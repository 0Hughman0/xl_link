import pandas as pd

try:
    from pandas.io.formats.excel import ExcelFormatter
except ImportError:
    from pandas.formats.format import ExcelFormatter

from .xl_types import XLCell

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
    frame_index: Pandas Index, or Array-like, used to determine location of index within spreadsheet.
    frame_columns: Pandas Index or Array-like, used to determine location of column within spreadsheet.
    excel_writer : string or ExcelWriter object
    sheet_name : string, default ‘Sheet1’, Name of sheet which will contain DataFrame
    columns : sequence, optional, Columns to write
    header : boolean or list of string, default Truem Write out the column names. If a list of strings is given
        it is assumed to be aliases for the column names
    index : boolean, default True. Write row names (index)
    index_label : string or sequence, default None. Column label for index column(s) if desired. If None is given,
        and header and index are True, then the index names are used. A sequence should be given if the
        DataFrame uses MultiIndex.
    startrow : upper left cell row to dump data frame
    startcol : upper left cell column to dump data frame
    merge_cells : boolean, default True. Write MultiIndex and Hierarchical Rows as merged cells.

    Returns
    -------
    tuple: (data_range, index_range, col_rangem, empty_f) of type (XLRange, XLRange, XLRange, DataFrame),
        where each range represents where the data, index and columns can be found on the spreadsheet, and empty_f is
        an empty DataFrame with matching Indexes.
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
    f : Pandas DataFrame to write to Excel
    excel_writer: Path or existing Excel Writer
    to_excel_args: Additional arguments to pass to DataFrame.to_excel, see docs for DataFrame.to_excel

    Returns
    -------
    XLMap: Mapping that corresponds to the position in the spreadsheet that frame was written to.
    """
    if to_excel_args is None:
        to_excel_args = {}

    default_args = {'sheet_name':'Sheet1',
                    'columns': None,
                    'header': True,
                    'index': True,
                    'index_label': None,
                    'startrow': 0,
                    'startcol': 0,
                    'merge_cells': True}
    default_args.update(to_excel_args)
    to_excel_args = default_args
    f.to_excel(excel_writer, **to_excel_args)

    data_range, index_range, col_range, _ = get_xl_ranges(f.index, f.columns,
                                                          sheet_name=to_excel_args['sheet_name'],
                                                          columns=to_excel_args['columns'],
                                                          header=to_excel_args['header'],
                                                          index=to_excel_args['index'],
                                                          index_label=to_excel_args['index_label'],
                                                          startrow=to_excel_args['startrow'],
                                                          startcol=to_excel_args['startcol'],
                                                          merge_cells=to_excel_args['merge_cells'])
    f = f.copy()

    columns = to_excel_args['columns']

    if isinstance(columns, list) or isinstance(columns, tuple):
        f = f[columns]

    return XLMap(data_range, index_range, col_range, f)


class SelectorProxy:
    """
    Proxy object that intercepts calls to Pandas DataFrame indexers, and re-interprets result into excel locations.

    Parameters
    ----------
    mapper_frame: Pandas DataFrame with index the same as the DataFrame it is representing, however, each cell contains
        the location they sit within the spreadsheet.
    selector_name: str, name of the indexer SelectorProxy is emulating, i.e. loc, iloc, ix, iat or at

    Only implements __getitem__ behaviour of indexers.
    """

    def __init__(self, mapper_frame, selector_name):
        self.mapper_frame = mapper_frame
        self.selector_name = selector_name

    def __getitem__(self, key):
        val = getattr(self.mapper_frame, self.selector_name)[key]

        if isinstance(val, pd.Series):
            return val.values[0] - val.values[-1]

        if isinstance(val, pd.DataFrame):

            return val.values[0, 0] - val.values[-1, -1]

        return val


class XLMap:
    """
    An object that maps a Pandas DataFrame to it's positions on an excel spreadsheet.

    Provides access to basic pandas indexers - __getitem__, loc, iloc, ix, iat and at.
    These indexers are modified such that they return the cell/ range of the result.

    The idea is should make using the data in spreadsheet easy to access, by using Pandas indexing syntax.
    For example can be used to create charts more easily (see example below).

    Beware!
    -------
    XLMap can only go 'one level deep' in terms of indexing, because each indexer always returns either an XLCell,
    or an XLRange. The only workaround is to reduce the size of your DataFrame BEFORE you call write_frame.
    This limitation drastically simplifies the implementation. Examples of what WON'T WORK:

    >>> map.loc['Mon':'Tues', :].index
        AttributeError: 'XLRange' object has no attribute 'index'

    >>> map.index['Mon':'Tues'] # Doesn't work because index is not a Pandas Index, but an XLRange.
        TypeError: unsupported operand type(s) for -: 'str' and 'int'

    Parameters
    ----------
    data_range : XLRange that represents the region the DataFrame's data sits in.
    index_range: XLRange that represents the region the DataFrame's index sits in.
    column_range: XLRange that represents the region the DataFrame's columns sit in.
    f: Pandas DataFrame that has been written to excel.

    Recommended to not be created directly, instead via, xl_link.write_frame(f, excel_writer, **kwargs)

    Examples
    --------
    calories_per_meal = pd.DataFrame(columns=("Meal", "Mon", "Tues", "Weds", "Thur"),
                                 data={'Meal': ('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                                       'Mon': (15, 20, 12, 3),
                                       'Tues': (5, 16, 3, 0),
                                       'Weds': (3, 22, 2, 8),
                                       'Thur': (6, 7, 1, 9)})
    calories_per_meal.set_index("Meal", drop=True, inplace=True)

    # Write to excel
    writer = pd.ExcelWriter("Example.xlsx", engine='xlsxwriter')
    map = write_frame(calories_per_meal, writer, sheet_name="XLLinked") # returns the 'ProxyFrame'

    # Create chart with XLLink
    workbook = writer.book
    xl_linked_sheet = writer.sheets["XLLinked"]
    xl_linked_chart = workbook.add_chart({'type': 'column'})

    for time in calories_per_meal.index:
        xl_linked_chart.add_series({'name': time,
                      'categories': proxy.columns.frange,
                      'values': proxy.loc[time].frange})
    """

    def __init__(self, data_range, index_range, column_range, f):
        self.index = index_range
        self.columns = column_range

        self.data = data_range
        self._f = f.copy()
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

    def __repr__(self):
        return "<XLMap: index: {}, columns: {}, data: {}>".format(self.index, self.columns, self.data)

    def __getitem__(self, key):
        """
        Emulates DataFrame.__getitem__ (DataFrame[key] syntax), see Pandas DataFrame indexing for help on behaviour.

        Will return the location of the columns found, rather than the underlying data.

        Parameters
        ----------
        key : hashable or array-like of hashables, corresponding to the names of the columns desired.

        Returns
        -------
        XLRange corresponding to position of found colummn(s) within spreadsheet

        Example
        -------
        >>> map['Col 1']
            <XLRange: B2:B10>
        """
        try:
            i = self.f.columns.get_loc(key)
            return self.columns[i]
        except KeyError:
            pass
        except TypeError:
            pass

        try:
            i = slice(*self.f.columns.slice_locs(key[0], key[-1]))
        except TypeError:
            raise TypeError("Cannot interpret key: {}".format(key))

        return self.columns[i]

    @property
    def loc(self):
        """
        Proxy for DataFrame.loc, see Pandas DataFrame loc help for behaviour.

        Will return location result rather than underlying data.

        Returns
        -------
        XLCell or XLRange corresponding to position of DataFrame, Series or Scalar found within spreadsheet.

        Example
        -------
        >>> map.loc['Tues']
            <XLRange: A2:D2>
        """
        return SelectorProxy(self._mapper_frame, 'loc')

    @property
    def iloc(self):
        """
        Proxy for DataFrame.iloc, see Pandas DataFrame iloc help for behaviour.

        Will return location result rather than underlying data.

        Returns
        -------
        XLCell or XLRange corresponding to position of DataFrame, Series or Scalar found within spreadsheet.

        Example
        -------
        >>> map.iloc[3, :]
            <XLRange: A2:D2>
        """
        return SelectorProxy(self._mapper_frame, 'iloc')

    @property
    def ix(self):
        """
        Proxy for DataFrame.ix, see Pandas DataFrame ix help for behaviour. (That said this is deprecated since 0.20!)

        Will return location result rather than underlying data.

        Returns
        -------
        XLCell or XLRange corresponding to position of DataFrame, Series or Scalar found within spreadsheet.


        Example
        -------
        >>> map.ix[3, :]
            <XLRange A2:D2>
        """
        return SelectorProxy(self._mapper_frame, 'ix')

    @property
    def iat(self):
        """
        Proxy for DataFrame.iat, see Pandas DataFrame iat help for behaviour.

        Will return location result rather than underlying data.

        Returns
        -------
        XLCell location corresponding to position value within spreadsheet.

        Example
        -------
        >>> map.iat[3, 2]
            <XLCell C3>
        """
        return SelectorProxy(self._mapper_frame, 'iat')

    @property
    def at(self):
        """
        Proxy for DataFrame.at, see Pandas DataFrame at help for behaviour.

        Will return location result rather than underlying data.

        Returns
        -------
        XLCell location corresponding to position value within spreadsheet.


        Example
        -------
        >>> map.at["Mon", "Lunch"]
            <XLCell: C3>
        """
        return SelectorProxy(self._mapper_frame, 'at')


class XLDataFrame(pd.DataFrame):
    """
    
    XLDataFrame - standard DataFrame monkeypatched to return xl_link.XLMap upon use of to_excel
    
    """

    __doc__ += pd.DataFrame.__doc__

    @property
    def _constructor(self):
        return XLDataFrame

    def to_excel(self, excel_writer, sheet_name='Sheet1', na_rep='',
                 float_format=None, columns=None, header=True, index=True,
                 index_label=None, startrow=0, startcol=0, engine=None,
                 merge_cells=True, encoding=None, inf_rep='inf', verbose=True,
                 freeze_panes=None):
        super().to_excel(excel_writer, sheet_name='Sheet1', na_rep='',
                 float_format=None, columns=None, header=True, index=True,
                 index_label=None, startrow=0, startcol=0, engine=None,
                 merge_cells=True, encoding=None, inf_rep='inf', verbose=True,
                 freeze_panes=None)

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

        return XLMap(data_range, index_range, col_range, f)