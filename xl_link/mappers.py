import pandas as pd

try:
    from pandas.io.formats.excel import ExcelFormatter
except ImportError:
    from pandas.formats.format import ExcelFormatter

from pandas.io.common import _stringify_path

from .xl_types import XLCell
from .chart_wrapper import create_chart, SINGLE_CATEGORY_CHARTS, CATEGORIES_REQUIRED_CHARTS


def get_xl_ranges(frame_index, frame_columns,
                  sheet_name='Sheet1',
                  columns=None,
                  header=True,
                  index=True,
                  index_label=None,
                  startrow=0,
                  startcol=0,
                  merge_cells=True):
    """
    Deduces location of data_range, index_range and col_range within excel spreadsheet, given the parameters provided.

    Does not require an actual DataFrame, which could be useful!

    Parameters
    ----------
    frame_index: Pandas Index or Array-like
        to determine location of index within spreadsheet.
    frame_columns: Pandas Index or Array-like
        used to determine location of column within spreadsheet.
    excel_writer : string or ExcelWriter

    sheet_name : str
        default ‘Sheet1’, Name of sheet which will contain DataFrame
    columns : sequence
        optional, Columns to write
    header : bool or list of strings,
        default True Write out the column names. If a list of strings is given it is assumed to be aliases for the column names
    index : bool
        default True. Write row names (index)
    index_label : str or sequence
        default None. Column label for index column(s) if desired. If None is given, and header and index are True, then the index names are used. A sequence should be given if the
        DataFrame uses MultiIndex.
    startrow : int
        upper left cell row to dump data frame
    startcol : int
        upper left cell column to dump data frame
    merge_cells : bool
        default True. Write MultiIndex and Hierarchical Rows as merged cells.

    Returns
    -------
    data_range, index_range, col_range : XLRange
        Each range represents where the data, index and columns can be found on the spreadsheet
    empty_f : DatFrame
        an empty DataFrame with matching Indices.
    """

    empty_f = pd.DataFrame(index=frame_index, columns=frame_columns)

    formatter = ExcelFormatter(empty_f,
                               cols=columns,
                               header=header,
                               index=index,
                               index_label=index_label,
                               merge_cells=merge_cells)

    excel_header = list(formatter._format_header())
    col_start, col_stop = excel_header[0], excel_header[-1]

    col_start_cell = XLCell(col_stop.row + startrow, col_start.col + startcol, sheet_name)
    col_stop_cell = XLCell(col_stop.row + startrow, col_stop.col + startcol, sheet_name)

    if isinstance(empty_f.index, pd.MultiIndex):
        col_start_cell = col_start_cell.translate(0, 1)

    col_range = col_start_cell - col_stop_cell

    body = list(formatter._format_body())

    if empty_f.index.name or index_label:
        body.pop(0)  # gets rid of index label cell that comes first!

    index_start_cell = XLCell(body[0].row + startrow, body[0].col + startcol + empty_f.index.nlevels - 1, sheet_name)
    index_stop_cell = XLCell(body[-1].row + startrow, body[0].col + startcol + empty_f.index.nlevels - 1, sheet_name)

    index_range = index_start_cell - index_stop_cell

    data_start_cell = XLCell(index_start_cell.row, col_start_cell.col, sheet_name)
    data_stop_cell = XLCell(index_stop_cell.row, col_stop_cell.col, sheet_name)

    data_range = data_start_cell - data_stop_cell

    return data_range, index_range, col_range, empty_f


def write_frame(f, excel_writer, to_excel_args=None):
    """
    Write a Pandas DataFrame to excel by calling to_excel, returning an XLMap, that can be used to determine
    the position of parts of f, using pandas indexing.

    Parameters
    ----------
    f : DataFrame
        Frame to write to excel
    excel_writer : str or ExcelWriter
        Path or existing Excel Writer to use to write frame
    to_excel_args : dict
        Additional arguments to pass to DataFrame.to_excel, see docs for DataFrame.to_excel

    Returns
    -------
    XLMap :
        Mapping that corresponds to the position in the spreadsheet that frame was written to.
    """
    xlf = XLDataFrame(f)

    return xlf.to_excel(excel_writer, **to_excel_args)


def _mapper_to_xl(value):
    """
    Convert mapper frame result to XLRange or XLCell
    """
    if isinstance(value, XLCell):
        return value

    if isinstance(value, pd.Series):
        return value.values[0] - value.values[-1]

    if isinstance(value, pd.DataFrame):
        return value.values[0, 0] - value.values[-1, -1]

    raise TypeError("Could not conver {} to XLRange or XLCell".format(value))


class _SelectorProxy:
    """
    Proxy object that intercepts calls to Pandas DataFrame indexers, and re-interprets result into excel locations.

    Parameters
    ----------
    mapper_frame: DataFrame
        with index the same as the DataFrame it is representing, however, each cell contains
        the location they sit within the spreadsheet.
    selector_name: str
        name of the indexer SelectorProxy is emulating, i.e. loc, iloc, ix, iat or at

    Notes
    -----
    Only implements __getitem__ behaviour of indexers.
    """

    def __init__(self, mapper_frame, selector_name):
        self.mapper_frame = mapper_frame
        self.selector_name = selector_name

    def __getitem__(self, key):
        val = getattr(self.mapper_frame, self.selector_name)[key]

        return _mapper_to_xl(val)


class XLMap:
    """
    An object that maps a Pandas DataFrame to it's positions on an excel spreadsheet.

    Provides access to basic pandas indexers - __getitem__, loc, iloc, ix, iat and at.
    These indexers are modified such that they return the cell/ range of the result.

    The idea is should make using the data in spreadsheet easy to access, by using Pandas indexing syntax.
    For example can be used to create charts more easily (see example below).

    Notes
    -----

    Recommended to not be created directly, instead via, XLDataFrame.to_excel.

    XLMap can only go 'one level deep' in terms of indexing, because each indexer always returns either an XLCell,
    or an XLRange. The only workaround is to reduce the size of your DataFrame BEFORE you call write_frame.
    This limitation drastically simplifies the implementation. Examples of what WON'T WORK:

    >>> xlmap.loc['Mon':'Tues', :].index
        AttributeError: 'XLRange' object has no attribute 'index'

    >>> xlmap.index['Mon':'Tues'] # Doesn't work because index is not a Pandas Index, but an XLRange.
        TypeError: unsupported operand type(s) for -: 'str' and 'int'

    Parameters
    ----------
    data_range, index_range, column_range : XLRange
     that represents the region the DataFrame's data sits in.
    f : DataFrame
     that has been written to excel.

    Attributes
    ----------
    index : XLRange
        range that the index column occupies
    columns : XLRange
        range that the frame columns occupy
    data : XLRange
        range that the frame data occupies
    writer : Pandas.ExcelWriter
        writer used to create spreadsheet
    sheet : object
        sheet object corresponding to sheet the frame was written to, handy if you want insert a chart into the same sheet

    Examples
    --------
    >>> calories_per_meal = XLDataFrame(columns=("Meal", "Mon", "Tues", "Weds", "Thur"),
                                         data={'Meal': ('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                                               'Mon': (15, 20, 12, 3),
                                               'Tues': (5, 16, 3, 0),
                                               'Weds': (3, 22, 2, 8),
                                               'Thur': (6, 7, 1, 9)})
    >>> calories_per_meal.set_index("Meal", drop=True, inplace=True)

    Write to excel

    >>> writer = pd.ExcelWriter("Example.xlsx", engine='xlsxwriter')
    >>> xlmap = calories_per_meal.to_excel(writer, sheet_name="XLLinked") # returns the XLMap

    Create chart with XLLink

    >>> workbook = writer.book
    >>> xl_linked_sheet = writer.sheets["XLLinked"]
    >>> xl_linked_chart = workbook.add_chart({'type': 'column'})
    >>> for time in calories_per_meal.index:
    >>>     xl_linked_chart.add_series({'name': time,
                                        'categories': proxy.columns.frange,
                                        'values': proxy.loc[time].frange})
    """

    def __init__(self, data_range, index_range, column_range, f, writer=None):
        self.index = index_range
        self.columns = column_range

        self.data = data_range
        self._f = f.copy()

        self.writer = writer
        self.book = writer.book
        self.sheet = writer.sheets[self.index.sheet]

        self._mapper_frame = f.copy().astype(XLCell)

        x_range = self._f.index.size
        y_range = self._f.columns.size

        for x in range(x_range):
            for y in range(y_range):
                self._mapper_frame.values[x, y] = data_range[x, y]

    @property
    def f(self):
        """
        for convenience provides read-only access to the DataFrame originally written to excel.
        """
        return self._f

    @property
    def df(self):
        """
        for convenience provides read-only access to the DataFrame originally written to excel.
        """
        return self._f

    def __repr__(self):
        return "<XLMap: index: {}, columns: {}, data: {}>".format(self.index, self.columns, self.data)

    def create_chart(self, type_='scatter',
                     values=None, categories=None, names=None,
                     subtype=None,
                     title=None, x_axis_name=None, y_axis_name=None):
        """
        Create excel chart object based off of data within the Frame.

        Parameters
        ----------
        type_ : str
            Type of chart to create.
        values : str or list or tuple
            label or list of labels to corresponding to column to use as values for each series in chart.
            Default all columns.
        categories : str or list or tuple
            label or list of labels to corresponding to column to use as categories for each series in chart.
            Default, use index for 'scatter' or None for everything else.
        names: str or list or tuple
            str or list of strs to corresponding to names for each series in chart.
            Default, column names corresponding to values.
        subtype : str
            subtype of type, only available for some chart types e.g. bar, see Excel writing package for details
        title : str
            chart title
        x_axis_name : str
            used as label on x_axis
        y_axis_name : str
            used as label on y_axis

        Returns
        -------

        Chart object corresponding to the engine selected

        Notes
        -----
        values, categories parameters can only correspond to columns.

        """

        if names is None and categories is None:
            names = tuple(name for name in self.f.columns.values)
        elif names is None and isinstance(categories, (str, int, list, tuple)):
            names = categories
        elif isinstance(names, (str, list, tuple)):
            names = names
        else:
            raise TypeError("Couldn't understand names input: " + names)

        if values is None:
            values = tuple(self[value] for value in self.f.columns)
        elif isinstance(values, list) or isinstance(values, tuple):
            values = tuple(self[value] for value in values)
        else:
            values = self[values]

        if categories is None and (type_ in SINGLE_CATEGORY_CHARTS and isinstance(values, tuple)) or \
                        type_ in CATEGORIES_REQUIRED_CHARTS:
            categories = self.index # Default, use x as index
        elif categories is None:
            pass
        elif isinstance(categories, (list, tuple)):
            categories = list(self[category] for category in categories)
        else:
            categories = self[categories]

        return create_chart(self.book, self.writer.engine, type_,
                            values, categories, names,
                            subtype, title,
                            x_axis_name, y_axis_name)

    def __getitem__(self, key):
        """
        Emulates DataFrame.__getitem__ (DataFrame[key] syntax), see Pandas DataFrame indexing for help on behaviour.

        Will return the location of the columns found, rather than the underlying data.

        Parameters
        ----------
        key : hashable or array-like
            hashables, corresponding to the names of the columns desired.

        Returns
        -------
        XLRange :
            corresponding to position of found colummn(s) within spreadsheet

        Example
        -------
        >>> xlmap['Col 1']
            <XLRange: B2:B10>
        """
        val = self._mapper_frame[key]

        return _mapper_to_xl(val)

    @property
    def loc(self):
        """
        Proxy for DataFrame.loc, see Pandas DataFrame loc help for behaviour.

        Will return location result rather than underlying data.

        Returns
        -------
        XLCell or XLRange
            corresponding to position of DataFrame, Series or Scalar found within spreadsheet.

        Example
        -------
        >>> xlmap.loc['Tues']
            <XLRange: A2:D2>
        """
        return _SelectorProxy(self._mapper_frame, 'loc')

    @property
    def iloc(self):
        """
        Proxy for DataFrame.iloc, see Pandas DataFrame iloc help for behaviour.

        Will return location result rather than underlying data.

        Returns
        -------
        XLCell or XLRange
            corresponding to position of DataFrame, Series or Scalar found within spreadsheet.

        Example
        -------
        >>> xlmap.iloc[3, :]
            <XLRange: A2:D2>
        """
        return _SelectorProxy(self._mapper_frame, 'iloc')

    @property
    def ix(self):
        """
        Proxy for DataFrame.ix, see Pandas DataFrame ix help for behaviour. (That said this is deprecated since 0.20!)

        Will return location result rather than underlying data.

        Returns
        -------
        XLCell or XLRange
            corresponding to position of DataFrame, Series or Scalar found within spreadsheet.


        Example
        -------
        >>> xlmap.ix[3, :]
            <XLRange A2:D2>
        """
        return _SelectorProxy(self._mapper_frame, 'ix')

    @property
    def iat(self):
        """
        Proxy for DataFrame.iat, see Pandas DataFrame iat help for behaviour.

        Will return location result rather than underlying data.

        Returns
        -------
        XLCell
            location corresponding to position value within spreadsheet.

        Example
        -------
        >>> xlmap.iat[3, 2]
            <XLCell C3>
        """
        return _SelectorProxy(self._mapper_frame, 'iat')

    @property
    def at(self):
        """
        Proxy for DataFrame.at, see Pandas DataFrame at help for behaviour.

        Will return location result rather than underlying data.

        Returns
        -------
        XLCell
            location corresponding to position value within spreadsheet.


        Example
        -------
        >>> xlmap.at["Mon", "Lunch"]
            <XLCell: C3>
        """
        return _SelectorProxy(self._mapper_frame, 'at')


class XLDataFrame(pd.DataFrame):
    """
    Monkeypatched DataFrame modified by xl_link!

    Changes:
    --------

    * to_excel modified to return an XLMap.

    * XLDataFrame._constructor set to XLDataFrame -> stops reverting to normal DataFrame

    Notes
    -----

    Conversions from this DataFrame to Series or Panels will return regular Panels and Series,
    which will convert back into regular DataFrame's upon expanding/ reducing dimensions.

    See Also
    --------
    Pandas.DataFrame
    """

    @property
    def _constructor(self):
        return XLDataFrame

    def to_excel(self, excel_writer, sheet_name='Sheet1', na_rep='',
                 float_format=None, columns=None, header=True, index=True,
                 index_label=None, startrow=0, startcol=0, engine=None,
                 merge_cells=True, encoding=None, inf_rep='inf', verbose=True,
                 **kwargs):
        """

        Monkeypatched DataFrame.to_excel by xl_link!

        Changes:
        --------

        Returns
        -------

        XLMap
            corresponding to position of frame as it appears in excel (see XLMap for details)

        See Also
        --------

        Pandas.DataFrame.to_excel for info on parameters

        Note
        ----
        When providing a path as excel_writer, default engine used is 'xlsxwriter', as xlsxwriter workbooks can only be
        saved once, xl_link suppresses calling `excel_writer.save()`, as a result, `xlmap.writer.save()` should be
        called once no further changes are to be made to the spreadsheet.
        """

        if isinstance(excel_writer, pd.ExcelWriter):
            need_save = False
        else:
            excel_writer = pd.ExcelWriter(_stringify_path(excel_writer), engine=engine)
            need_save = True if excel_writer.engine != 'xlsxwriter' else False # xlsxwriter can only save once!

        super().to_excel(excel_writer, sheet_name=sheet_name, na_rep=na_rep,
                 float_format=float_format, columns=columns, header=header, index=index,
                 index_label=index_label, startrow=startrow, startcol=startcol, engine=engine,
                 merge_cells=merge_cells, encoding=encoding, inf_rep=inf_rep, verbose=verbose,
                 **kwargs)

        if need_save:
            excel_writer.save()

        data_range, index_range, col_range, _ = get_xl_ranges(self.index, self.columns,
                                                              sheet_name=sheet_name,
                                                              columns=columns,
                                                              header=header,
                                                              index=index,
                                                              index_label=index_label,
                                                              startrow=startrow,
                                                              startcol=startcol,
                                                              merge_cells=merge_cells)
        f = self.copy()

        if isinstance(columns, list) or isinstance(columns, tuple):
            f = f[columns]

        return XLMap(data_range, index_range, col_range, f, writer=excel_writer)