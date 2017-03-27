import pandas as pd
from pandas.lib import maybe_booleans_to_slice, maybe_indices_to_slice
from pandas.types.common import is_bool_dtype
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


class IndexerProxyMixin:
    """
    Provides universal settings for IndexerProxies
    """
    _attributes = ["name", "xl"] # Don't forget where I am
    _infer_as_myclass = True # Don't change me

    def __getitem__(self, item):
        obj = super().__getitem__(item)
        if isinstance(obj, pd.Index):
            obj = self.__class__(obj)
            obj.xl = self.xl[item]
        return obj


# Factory create IndexProxies as they're all the same
for indexer_type in (Int64Index, Float64Index, RangeIndex,
                     CategoricalIndex,
                     DatetimeIndex, PeriodIndex, TimedeltaIndex):

    class IndexerProxyFactory(IndexerProxyMixin, indexer_type):
        pass

    IndexerProxyFactory.__name__ = "{}Proxy".format(indexer_type.__name__)
    index_proxies[indexer_type.__name__] = IndexerProxyFactory


class MultiIndexProxy(MultiIndex):
    """
    Like a normal MultiIndex, but with the .xl attribute, that gives the location of the labels row
    """
    _attributes = ["name", "xl"] # Don't forget where I am
    _infer_as_myclass = True # Don't change me

    def __new__(cls, *args, **kwargs):
        """
        Overwritten to stop MultiIndex from sneakily change back to vanilla types
        """
        obj = super().__new__(cls, *args, **kwargs)
        if isinstance(obj, pd.Index):
            return IndexProxy(obj)
        elif isinstance(obj, pd.MultiIndex):
            return MultiIndexProxy.guarantee_proxy_new(obj)
        else:
            return obj

    @classmethod
    def guarantee_proxy_new(cls, multi_index):
        """
        From Pandas Source:  pandas/pandas/indexes/multi.py
        I don't really unerstand all of this bit! - DANGER!
        """
        result = object.__new__(MultiIndexProxy) # what is this ? 8O
        result._set_levels(multi_index.levels, copy=True, validate=False)
        result._set_labels(multi_index.labels, copy=True, validate=False)

        if multi_index.names is not None:
            # handles name validation
            result._set_names(multi_index.names)

        if multi_index.sortorder is not None:
            result.sortorder = int(multi_index.sortorder)
        else:
            result.sortorder = multi_index.sortorder
        result._reset_identity()
        return result
        """Stop"""

    def __getitem__(self, item):
        obj = super().__getitem__(item)
        if isinstance(obj, pd.MultiIndex):
            obj = MultiIndexProxy.guarantee_proxy_new(obj)
            obj.xl = self.xl[item]
        return obj


class IndexProxy(IndexerProxyMixin, pd.Index):
    """
    Needs to be treated specially as __new__ is often called, and allowed to determine what Index type to use.

    Overwriting this allows intercepting the vanilla Index types.
    """
    def __new__(cls, *args, **kwargs):
        if args and isinstance(args[0], pd.MultiIndex):
            index = args[0]
            # As pd.Index does not initialise to Multindexes, need to help out
            return MultiIndexProxy.guarantee_proxy_new(index)
        obj = super().__new__(cls, *args, **kwargs)
        try:
            obj = index_proxies[obj.__class__.__name__](obj, *args, **kwargs)
        except KeyError:
            pass
        return obj


class FrameProxy(pd.DataFrame):
    """
    DataFrame that contains all the information about it's parent DataFrame's positions on an Excel file
    """

    @classmethod
    def create(cls, frame, startrow, startcol, sheetname, index_label=False):
        """
        Initialise creating the 'Proxy' version of frame
        """
        obj = cls(frame)
        for x in range(obj.index.size):
            for y in range(obj.columns.size):
                obj.ix[x, y] = XLCell(sheetname, x + startrow, y + startcol)

        i_start = obj.iat[0, 0].translate(0, -1)
        i_stop = obj.iat[-1, 0].translate(0, -1)
        cols_start = obj.iat[0, 0].translate(-1, 0)
        cols_stop = obj.iat[0, -1].translate(-1, 0)
        if index_label: # Puts strange blank line between top of table and index!?
            cols_start = cols_start.translate(-1, 0)
            cols_stop = cols_stop.translate(-1, 0)

        obj.columns = IndexProxy(obj.columns)
        obj.columns.xl = cols_start - cols_stop
        obj.columns.axis = 1

        obj.index = IndexProxy(obj.index)
        obj.index.xl = i_start - i_stop
        obj.index.axis = 0

        return obj

    @property
    def xl(self):
        """
        Provides xl property that gives the XLRange corresponding to position
        """
        return super().iat[0, 0] - super().iat[-1, -1]

    @property
    def _constructor(self):
        return FrameProxy

    @property
    def _constructor_sliced(self):
        return SeriesProxy


class SeriesProxy(pd.Series):
    """
    Proxy for Pandas Series
    """

    @property
    def xl(self):
        """
        Provides xl property that gives the XLRange corresponding to position
        """
        return self.iat[0] - self.iat[-1]

    @property
    def _constructor(self):
        return SeriesProxy


class EmbededFrame(pd.DataFrame):
    """
    DataFrame with enhanced to_excel method
    """

    def to_excel(self, excel_writer,
                 sheet_name="Sheet1",
                 columns=None,
                 header=True,
                 index=True,
                 index_label=None,
                 startrow=0,
                 startcol=0,
                 *args, **kwargs):
        """
        See pandas docs for to_excel details. Parameters are the same.

        :return: A DataFrame proxy that provides pandas 'fancy indexing' for excel spreadsheets
        """
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
        if f.index.nlevels > 1 and index_label is not False:
            index_label = True
        if header:
            startrow += f.columns.nlevels
            if index_label:
                startrow += 1
        if index:
            startcol += f.index.nlevels
        frame_proxy = FrameProxy.create(f, startrow, startcol, sheet_name, index_label)
        return frame_proxy

    @property
    def _constructor(self):
        return EmbededFrame

