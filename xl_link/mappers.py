import pandas as pd
import pandas.formats.format as fmt

from .xl_types import XLCell


class IndexXLMapBase:
    _attributes = ["name", "xl"]  # Don't forget where I am
    _infer_as_myclass = True  # Don't change me

    def _constructor(self):
        return self.__class__

    def __getitem__(self, item):
        obj = super().__getitem__(item)
        if isinstance(item, slice):
            item = slice(item.start, item.stop, item.step)
        if not isinstance(obj, IndexXLMapBase):
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
        start = min(indices)
        stop = max(indices)
        if (stop - start) / len(indices) > 1:
            raise TypeError("XL_Link only supports plain slicing (step = 1)")
        xl = self.xl[start:stop]
        obj.xl = xl
        return obj


def index_type_to_indexproxy(index):
    if not isinstance(index, pd.MultiIndex):

        class IndexXLMapFactory(IndexXLMapBase, index.__class__):

            def __new__(cls, *args, **kwargs):
                obj = super().__new__(cls, *args, **kwargs)
                if isinstance(obj, IndexXLMapBase):
                    return obj
                else:
                    return index_type_to_indexproxy(obj)
    else:

        class IndexXLMapFactory(IndexXLMapBase, index.__class__):

            def __new__(cls, *args, **kwargs):
                obj = super().__new__(cls, *args, **kwargs)
                if isinstance(obj, IndexXLMapBase):
                    return obj
                else:
                    return cls.guarantee_proxy_new(obj)

            @classmethod
            def guarantee_proxy_new(cls, multi_index):
                """
                From Pandas Source:  pandas/pandas/indexes/multi.py
                I don't really understand all of this bit! - DANGER!
                """
                result = object.__new__(cls)  # what is this ? 8O
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

    cls = IndexXLMapFactory
    cls.__name__ = "{}Proxy".format(index.__class__.__name__)

    if isinstance(index, pd.MultiIndex):
        return cls.guarantee_proxy_new(index)
    else:
        return cls(index)


class SeriesXLMap(pd.Series):
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
        return SeriesXLMap

    @property
    def _constructor_expanddim(self):
        return FrameXLMap


class FrameXLMap(pd.DataFrame):
    """
    DataFrame that contains all the information about it's parent DataFrame's positions on an Excel file
    """

    @classmethod
    def create(cls, frame, col_range, index_range, data_range, excel_writer):
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

        obj.writer = excel_writer

        return obj

    @property
    def xl(self):
        """
        Provides xl property that gives the XLRange corresponding to position
        """
        return self.iat[0, 0] - self.iat[-1, -1]

    @property
    def _constructor(self):
        return FrameXLMap

    @property
    def _constructor_sliced(self):
        return SeriesXLMap

    @property
    def _constructor_expanddim(self):
        return pd.Panel

    def plot(self, x=None, y=None, kind='scatter', subkind=None,
             figsize=None, use_index=True,
             title=None, grid=False, legend=True, style=None,
             logx=False, logy=False, loglog=False, xticks=None, yticks=None, xlim=None, ylim=None, rot=None,
             fontsize=None, colormap=None,
             yerr=None, xerr=None,
             secondary_y=False,
             sort_columns=False):
        """
        Create xlsxwriter.chart.Chart object from dataframe to add to spreadsheet
        :param x: str/int, default None
                    label or position,
        :param y: str/list/tuple, default None
                    series of or label/ position
        :param kind: str, default 'scatter'
                    'line', 'bar', 'column', 'pie', 'doughnut', 'scatter', 'stock', 'radar',
        :param subkind: str, default straight_with_markers (scatter)
                    depends on kind i.e.
                        area: 'stacked', 'percent_stacked'
                        bar: 'stacked', 'percent_stacked'
                        column: 'stacked', 'percent_stacked'
                        scatter: 'straight_with_markers', 'straight', 'smooth_with_markers', 'smooth'
                        radar: 'with_markers', 'filled'
        :param figsize: tuple, default None
                    (width, height) of chart in pixels
        :param use_index: boolean, default True
                    use index as x axis series
        :param title: str, defaults None
                    title on top of plot
        :param grid: boolean, default False
                    include grid-lines
        :param legend: boolean, default True
                    include legend
        :param style: ???
        :param logx: boolean, default False
                    use log scale on x axis
        :param logy: boolean, default False
                    use log scale on y axis
        :param loglog: boolean, default False
                    use log scale on both axes
        :param xticks: tuple, default None
                    tick frequency for x axis in form (major, minor), None for auto
        :param yticks: tuple, default None
                    tick frequency for y axis in form (major, minor), None for auto
        :param xlim: tuple, default None
                    xaxis limits in form (lowlim, highlim), None for auto
        :param ylim: tuple, default None
                    yaxis limits in form (lowlim, highlim), None for auto
        :param rot: int, default None
                    rotation of axes labels in degrees ???
        :param fontsize: int, default None
                    font size, None for default
        :param colormap: ???
        :param yerr: dict, default None
                    dict mapping y columns to tuple containing plus and minus value column names e.g.
                        {'Column Name': ('plus val col', 'minus val col')}
        :param xerr: dict, default None
                    dict mapping x columns to tuple containing plus and minus value column names e.g.
                        {'Column Name': ('plus val col', 'minus val col')}
        :param secondary_y: str/list/tuple, default None
                    series of or label/ positions corresponding to columns to plot on secondary axis
        :return: xlsxwriter.chart.Chart object
        """
        pass


class EmbeddedFrame(pd.DataFrame):
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
            body.pop(0)  # gets rid of index label cell that comes first!

        index_start_cell = XLCell(sheet_name, body[0].row + startrow, body[0].col + startcol + f.index.nlevels - 1)
        index_stop_cell = XLCell(sheet_name, body[-1].row + startrow, body[0].col + startcol + f.index.nlevels - 1)

        index_range = index_start_cell - index_stop_cell

        data_start_cell = XLCell(sheet_name, index_start_cell.row, col_start_cell.col)
        data_stop_cell = XLCell(sheet_name, index_stop_cell.row, col_stop_cell.col)

        data_range = data_start_cell - data_stop_cell

        frame_proxy = FrameXLMap.create(f, col_range, index_range, data_range, excel_writer)

        return frame_proxy

    @property
    def _constructor(self):
        return EmbeddedFrame
