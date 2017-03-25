import pandas as pd
from pandas import (Int64Index,
                    Float64Index,
                    MultiIndex,
                    CategoricalIndex,
                    DatetimeIndex,
                    RangeIndex,
                    PeriodIndex,
                    TimedeltaIndex)

from xl_link.xl_types import XLCell

index_proxies = {}

for indexer_type in (Int64Index, Float64Index, RangeIndex,
                     MultiIndex,
                     CategoricalIndex,
                     DatetimeIndex, PeriodIndex, TimedeltaIndex):

    class IndexerProxyFactory(indexer_type):

        _attributes = ["name", "xl"]
        _infer_as_myclass = True

        def __getitem__(self, item):
            obj = super().__getitem__(item)
            obj = self.__class__(obj)
            obj.xl = self.xl[item]
            return obj

    IndexerProxyFactory.__name__ = "{}Proxy".format(indexer_type.__name__)
    index_proxies[indexer_type.__name__] = IndexerProxyFactory


class IndexProxy(pd.Index):

    _attributes = ["name", "origin"]

    def __new__(cls, *args, **kwargs):
        obj = super().__new__(cls, *args, **kwargs)
        try:
            obj = index_proxies[obj.__class__.__name__](obj, *args, **kwargs)
        except KeyError:
            pass
        return obj


class FrameProxy(pd.DataFrame):

    @classmethod
    def create(cls, frame, startrow, startcol, sheetname):
        obj = cls(frame)
        for x in range(obj.index.size):
            for y in range(obj.columns.size):
                obj.ix[x, y] = XLCell(sheetname, x + startrow, y + startcol)

        i_start = obj.iat[0, 0].transform(0, -1)
        i_stop = obj.iat[-1, 0].transform(0, -1)
        cols_start = obj.iat[0, 0].transform(-1, 0)
        cols_stop = obj.iat[0, -1].transform(-1, 0)

        obj.columns = IndexProxy(obj.columns)
        obj.columns.xl = cols_start - cols_stop
        obj.columns.axis = 1

        obj.index = IndexProxy(obj.index)
        obj.index.xl = i_start - i_stop
        obj.index.axis = 0
        return obj

    @property
    def xl(self):
        return super().iat[0, 0] - super().iat[-1, -1]

    @property
    def _constructor(self):
        return FrameProxy

    @property
    def _constructor_sliced(self):
        return SeriesProxy


class SeriesProxy(pd.Series):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    @property
    def xl(self):
        return self.iat[0] - self.iat[-1]

    @property
    def _constructor(self):
        return SeriesProxy


class EmbededFrame(pd.DataFrame):

    def to_excel(self, excel_writer,
                 sheet_name="Sheet1",
                 columns=None,
                 header=True,
                 index=True,
                 index_label=None,
                 startrow=0,
                 startcol=0,
                 *args, **kwargs):
        super().to_excel(excel_writer,
                         sheet_name=sheet_name,
                         columns=columns,
                         header=header,
                         index=index,
                         index_label=index_label,
                         startrow=startrow,
                         startcol=startcol,
                         *args, **kwargs)
        f = self.copy()
        if isinstance(header, list):
            f = f[header]
        if isinstance(index, list):
            f = f.ix[index, :]
        if header:
            startrow += f.columns.nlevels
            if f.index.name and index_label:
                startrow += 1
        if index:
            startcol += f.index.nlevels
        frame_proxy = FrameProxy.create(f, startrow, startcol, sheet_name)
        return frame_proxy

    @property
    def _constructor(self):
        return EmbededFrame

