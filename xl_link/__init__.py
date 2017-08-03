import pandas as pd
import pandas.formats.format as fmt


from .xl_types import XLCell


class IndexProxyBase:

    _attributes = ["name", "xl"] # Don't forget where I am
    _infer_as_myclass = True # Don't change me

    def _constructor(self):
        return self.__class__

    def __getitem__(self, item):
        obj = super().__getitem__(item)
        if isinstance(item, slice):
            item = slice(item.start, item.stop, item.step)
        if not isinstance(obj, IndexProxyBase):
            obj = index_type_to_indexproxy(obj)
        obj.xl = self.xl[item]
        return obj

    def union(self, *args, **kwargs):
        raise NotImplementedError("Current Index Proxies do not support this mutation method")

    def intersection(self, *args, **kwargs):
        raise NotImplementedError("Current Index Proxies do not support this mutation method")

    def insert(self, *args, **kwargs):
        raise NotImplementedError("Current Index Proxies do not support this mutation method")

    def delete(self, *args, **kwargs):
        raise NotImplementedError("Current Index Proxies do not support this mutation method")

    def take(self, indices, axis=0, allow_fill=True,
             fill_value=None, **kwargs):
        obj = super().take(indices, axis, allow_fill, fill_value, **kwargs)
        print(indices)
        start = min(indices)
        stop = max(indices)
        if (stop - start) / len(indices) > 1:
            raise TypeError("XL_Link only supports plain slicing (step = 1)")
        xl = self.xl[start:stop]
        obj.xl = xl
        return obj


def index_type_to_indexproxy(index):

    if not isinstance(index, pd.MultiIndex):
        class IndexProxyFactory(IndexProxyBase, index.__class__):

                def __new__(cls, *args, **kwargs):
                    """
                    if args and isinstance(args[0], pd.MultiIndex):
                        index = args[0]
                        # As pd.Index does not initialise to Multindexes, need to help out
                        return MultiIndexProxy.guarantee_proxy_new(index)
                    """
                    obj = super().__new__(cls, *args, **kwargs)
                    if isinstance(obj, IndexProxyBase):
                        return obj
                    else:
                        return index_type_to_indexproxy(obj)
    else:
        class IndexProxyFactory(IndexProxyBase, index.__class__):

            def __new__(cls, *args, **kwargs):
                """
                    if args and isinstance(args[0], pd.MultiIndex):
                        index = args[0]
                        # As pd.Index does not initialise to Multindexes, need to help out
                        return MultiIndexProxy.guarantee_proxy_new(index)
                """
                obj = super().__new__(cls, *args, **kwargs)
                if isinstance(obj, IndexProxyBase):
                    return obj
                else:
                    return cls.guarantee_proxy_new(obj)

            @classmethod
            def guarantee_proxy_new(cls, multi_index):
                """
                From Pandas Source:  pandas/pandas/indexes/multi.py
                I don't really understand all of this bit! - DANGER!
                """
                result = object.__new__(cls) # what is this ? 8O
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
                """
                Stop
                """

    cls = IndexProxyFactory
    cls.__name__ = "{}Proxy".format(index.__class__.__name__)
    if isinstance(index, pd.MultiIndex):
        return cls.guarantee_proxy_new(index)
    else:
        return cls(index)


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

    @property
    def _constructor_expanddim(self):
        return FrameProxy


class FrameProxy(pd.DataFrame):
    """
    DataFrame that contains all the information about it's parent DataFrame's positions on an Excel file
    """

    @classmethod
    def create(cls, frame, col_range, index_range, data_range):
        """
        Initialise creating the 'Proxy' version of frame
        """
        obj = cls(frame)

        x_range = obj.index.size
        y_range = obj.columns.size
        for x in range(x_range):
            for y in range(y_range):
                obj.iloc[x, y] = data_range[x, y]

        obj.columns = index_type_to_indexproxy(obj.columns)
        obj.columns.xl = col_range

        obj.index = index_type_to_indexproxy(obj.index)
        obj.index.xl = index_range

        return obj

    @property
    def xl(self):
        """
        Provides xl property that gives the XLRange corresponding to position
        """
        return self.iat[0, 0] - self.iat[-1, -1]

    @property
    def _constructor(self):
        return FrameProxy

    @property
    def _constructor_sliced(self):
        return SeriesProxy

    @property
    def _constructor_expanddim(self):
        return pd.Panel


class EmbededFrame(pd.DataFrame):
    """
    DataFrame with enhanced to_excel method
    """

    def to_excel(self, excel_writer, sheet_name='Sheet1', na_rep='',
                 float_format=None, columns=None, header=True, index=True,
                 index_label=None, startrow=0, startcol=0, engine=None,
                 merge_cells=True, encoding=None, inf_rep='inf', verbose=True):
        """
        See pandas docs for to_excel details. Parameters are the same.

        :return: A DataFrame proxy that provides pandas 'fancy indexing' for excel spreadsheets
        """
        super().to_excel(excel_writer, sheet_name=sheet_name, na_rep=na_rep,
                         float_format=float_format, columns=columns, header=True, index=index,
                         index_label=index_label, startrow=startrow, startcol=startcol, engine=engine,
                         merge_cells=True, encoding=encoding, inf_rep=na_rep, verbose=verbose)
        f = self.copy()
        if isinstance(header, list):
            f = f[header]
        if isinstance(index, list):
            f = f.ix[index, :]

        formatter = fmt.ExcelFormatter(self, na_rep=na_rep, cols=columns,
                                       header=header,
                                       float_format=float_format, index=index,
                                       index_label=index_label,
                                       merge_cells=merge_cells,
                                       inf_rep=inf_rep)

        excel_header = list(formatter._format_header())
        col_start, col_stop = excel_header[0], excel_header[-1]

        col_start_cell = XLCell(sheet_name, col_stop.row + startrow, col_start.col + startcol)
        col_stop_cell = XLCell(sheet_name, col_stop.row + startrow, col_stop.col + startcol)

        if isinstance(self.index, pd.MultiIndex):
            col_start_cell = col_start_cell.translate(0, 1)

        col_range = col_start_cell - col_stop_cell

        body = list(formatter._format_body())

        if f.index.name or index_label:
            body.pop(0) # gets rid of index label cell that comes first!

        index_start_cell = XLCell(sheet_name, body[0].row + startrow, body[0].col + startcol + f.index.nlevels - 1)
        index_stop_cell = XLCell(sheet_name, body[-1].row + startrow, body[0].col + startcol + f.index.nlevels - 1)

        index_range = index_start_cell - index_stop_cell

        data_start_cell = XLCell(sheet_name, index_start_cell.row, col_start_cell.col)
        data_stop_cell = XLCell(sheet_name, index_stop_cell.row, col_stop_cell.col)

        data_range = data_start_cell - data_stop_cell

        frame_proxy = FrameProxy.create(f, col_range, index_range, data_range)

        return frame_proxy

    @property
    def _constructor(self):
        return EmbededFrame
