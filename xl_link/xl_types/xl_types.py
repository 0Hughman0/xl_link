from pandas.core.common import is_bool_indexer

from xl_link.xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol


def is_int_type(i):
    try:
        return int(i) == i
    except TypeError:
        return False
    except ValueError:
        return False


def fill_slice(holey_slice):
    return slice(holey_slice.start if holey_slice.start is not None else 0,
                 holey_slice.stop if holey_slice.stop is not None else -1,
                 holey_slice.step if holey_slice.step is not None else 1)

def to_series(cell_or_range):
    return '=' + cell_or_range


class XLCell:
    """
    Represents the location of a cell in an Excel spreadsheet

    Parameters
    ----------
    row : int
        corresponding to cell row number (0 indexed)
    col : int
        corresponding to cell col number (0 indexed)
    sheet: str
        corresponding to the sheet the cell is in, default='Sheet1'

    Attributes
    ---------
    row : int
        corresponding to cell row number (0 indexed)
    col : int
        corresponding to cell col number (0 indexed)
    sheet: str
        corresponding to the sheet the cell is in, default='Sheet1'

    Examples
    --------
    >>> XLCell(0, 5, 'Accounts')
        <XLCell: F1>

    >>> XLCell.from_cell('F1', 'Accounts')
        <XLCell: F1>

    >>> XLCell.from_fcell("'Accounts'!F1")
        <XLCell: F1>

    Notes
    -----
    sheet can be set to None if desired, keep in mind it defaults to 'Sheet1', mostly for convenience.

    row, col and sheet are attributes that can be accessed and set directly.

    All methods are designed to return new objects, preserving the state of self.

    See Also
    --------
    XLCell.from_cell
    XLCell.from_fcell
    """

    def __init__(self, row, col, sheet='Sheet1'):
        self.sheet = sheet
        self.row = row
        self.col = col

    @classmethod
    def from_cell(cls, cell, sheet="Sheet1"):
        """
        Alternative constructor

        Parameters
        ---------
        cell : str
            cell in excel notation (e.g. 'A1'), representing cell location.
        sheet : sheet
            sheet cell belongs to, default='Sheet1'

        Returns
        -------
        XLCell
            Initialised XLCell
        """
        return cls(*xl_cell_to_rowcol(cell), sheet)

    @classmethod
    def from_fcell(cls, fcell):
        """
        Alternative constructor

        Parameters
        ---------
        fcell: str
            cell in excel formula notation (e.g. "'Sheet1'!A1"), representing cell location.

        Returns
        -------
        XLCell
            Initialised XLCell
        """
        try:
            sheet, cell = fcell.split('!')
            sheet = sheet.replace("'", "")
        except ValueError:
            raise ValueError("""Could not parse fcell: {}, is your fcell in the form "'Sheet1'!A1:B2"?""".format(fcell))
        return cls(*xl_cell_to_rowcol(cell), sheet)

    @property
    def cell(self):
        """
        Gets the cell location str.

        Returns
        -------
        str
            cell location in excel notation

        Examples
        --------
        >>> cell = XLCell(0, 5, 'Accounts')
        >>> cell.cell
            'F1'
        """
        return xl_rowcol_to_cell(self.row, self.col)

    @property
    def fcell(self):
        """
        Gets the cell location for use in excel formulas.

        Returns
        -------
        str
            cell location in excel notation, in '{sheet}'!{cell} form

        Examples
        --------
        >>> cell = XLCell(0, 5, 'Accounts')
        >>> cell.fcell
            "'Accounts'!F1"
        """
        return "'{}'!{}".format(self.sheet, self.cell)

    @property
    def rowcol(self):
        """
        Coordinates of cell

        Returns
        -------
        row, col : tuple(int, int)
            (row, col) corresponding to location of cell
        """
        return self.row, self.col

    def range_between(self, other):
        """
        Get XLRange between this cell and other.

        Parameters
        ----------
        other : XLCell
            other XLCell that makes up stop of excel range

        Returns
        -------
        XLRange
            range between this cell and other

        Example
        -------
        >>> start = XLCell.from_cell('A1')
        >>> stop = XLCell.from_cell('B7')
        >>> start.range_between(stop)
            <XLRange: A1:B7>

        """
        return XLRange(self, other)

    def __sub__(self, other):
        """
        Alias for XLCell.range_between
        """
        return XLRange(self, other)

    def __eq__(self, other):
        """
        Equality check.

        Parameters
        ----------
        other : object
            to compare for equality.

        Notes
        -----
        other can be another XLCell or str, if a str is provided, XLCell.from_fcell will be called on other, and then
        the comparison made. Comparisons between XLCells first compares sheets, and then for equal rows and columns.
        """
        if isinstance(other, str):
            return self == XLCell.from_fcell(other)

        if isinstance(other, XLCell):
            if self.sheet != other.sheet:
                return False
            return self.row == other.row and self.col == other.col

        return False

    def __repr__(self):
        return "<XLCell: {}>".format(self.fcell)

    def copy(self):
        """
        Returns
        -------
        XLCell
            copy of self
        """
        return XLCell(self.row, self.col, self.sheet)

    def translate(self, row, col):
        """
        Returns new XLCell with translation applied

        Parameters
        ----------
        row : int
            corresponding to movement in row direction (+ve down the spreadsheet)
        col : int
            corresponding to movement in col direction (+ve right across the spreadsheet)

        Returns
        -------
        XLCell
         new cell with translation applied
        """
        cell = self.copy()
        cell.row += row or 0
        cell.col += col or 0
        return cell

    def trans(self, row, col):
        """
        Short for XLCell.translate

        See Also
        --------
        XLCell.translate
        """
        return self.translate(row, col)

    def __hash__(self):
        return hash((self.row, self.col, self.sheet))

class XLRange:
    """
    Represents a range in an Excel spreadsheet

    Parameters
    ----------
    start : XLCell
        corresponding to start cell
    stop : XLCell
        corresponding to stop cell

    Attributes
    ----------
    start : XLCell
        corresponding to start cell
    stop : XLCell
        corresponding to stop cell
    sheet : str
        name of sheet range exists in

    Examples
    --------
    Using __init__

    >>> start = XLCell.from_cell('A1')
    >>> stop = XLCell.from_cell('B7')
    >>> XLRange(start, stop)
        <XLRange: A1:B7>

    Notes
    -----
    Upon initialisation, if start and stop are not from the same sheet, an assertion error is raised.

    start and stop are attributes that can be accessed and modified directly.

    All methods are designed to return new objects, preserving the state of self.

    See Also
    --------
    XLRange.from_range
    XLRange.from_frange
    """
    def __init__(self, start, stop):
        assert start.sheet == stop.sheet, "start and stop must be in the same sheet"
        self._sheet = start.sheet
        self.start = start.copy()
        self.stop = stop.copy()

    @classmethod
    def from_range(cls, range, sheet='Sheet1'):
        """
        Alternative constructor

        Parameters
        ----------

        range : str
            str in excel notation (e.g. 'A1:B20'), representing excel range.
        sheet : str
            sheet range belongs to, default='Sheet1'

        Returns
        -------
        XLRange
            initialised XLRange
        """
        start, stop = range.split(':')
        return cls(XLCell.from_cell(start, sheet), XLCell.from_cell(stop, sheet))

    @classmethod
    def from_frange(cls, frange):
        """
        Alternative constructor

        Parameters
        ----------
        frange : str
            str in excel formula notation (e.g. "'Sheet1'!A1:B20"), representing excel range.

        Returns
        -------
        XLRange
            initialised XLRange
        """
        sheet, range = frange.split('!')
        sheet = sheet.replace("'", "")
        return cls.from_range(range, sheet)

    @property
    def sheet(self):
        return self._sheet

    @sheet.setter
    def sheet(self, value):
        self._sheet, self.start.sheet, self.stop.sheet = value, value, value

    @property
    def range(self):
        """
        Gets the Excel range this object represents

        Returns
        -------
        str
            representing this range in excel notation

        Examples
        --------
        >>> start = XLCell.from_cell('A1')
        >>> stop = XLCell.from_cell('B7')
        >>> range = XLRange(start, stop)
        >>> range.range
            'A1:B7'
        """
        return "{}:{}".format(self.start.cell, self.stop.cell)

    @property
    def frange(self):
        """
        Gets the Excel range this object represents, for use in excel formulas

        Returns
        -------
        str
            representing this range in excel notation for use in excel formulas (e.g. "'{sheet}'!{start}:{stop}"

        Examples
        --------
        >>> start = XLCell.from_cell('A1')
        >>> stop = XLCell.from_cell('B7')
        >>> range = XLRange(start, stop)
        >>> range.frange
            "'Sheet1'!A1:B7"
        """
        return "'{}'!{}".format(self._sheet, self.range)

    @property
    def rowcol_rowcol(self):
        """
        Coordinates of start and stop

        Returns
        -------
        start.row, stop.col : tuple
            start position in rows and column num (start.row, start.col)
        stop.row, stop.col : tuple
            stop position in rows and column num (stop.row, stop.col)
        """
        return self.start.rowcol, self.stop.rowcol

    @property
    def shape(self):
        """
        Returns
        -------
        tuple
            representing shape of XLRange in form (height, width)
        """
        return self.stop.row - self.start.row + 1, self.stop.col - self.start.col + 1

    @property
    def is_row(self):
        return self.shape[0] == 1

    @property
    def is_col(self):
        return self.shape[1] == 1

    @property
    def is_1D(self):
        return self.is_row or self.is_col

    @property
    def is_2D(self):
        return not self.is_1D

    def __repr__(self):
        return "<XLRange: {}>".format(self.frange)

    def __len__(self):
        """
        Length of self

        Notes
        -----
        Will raise TypeError unless self is 1 dimensional
        """
        if self.is_col:
            return self.shape[0]
        if self.is_row:
            return self.shape[1]
        raise TypeError("length is only defined for 1D ranges")

    def __iter__(self):
        """
        iterate over each XLCell with self, if self is a row or column

        Notes
        -----
        Will raise TypeError if self is not 1 dimensional
        """
        if self.is_1D:
            for x in range(len(self)):
                yield self[x]
        else:
            raise TypeError("Can only iterate over 1D ranges")

    def __getitem__(self, key):
        """
        Get subsection of self using given key

        Parameters
        ----------
        key : int or slice or boolean indexer or tuple of 2 ints/ slices!

        Returns
        -------
        XLCell or XLRange
            result after selection using key applied. Designed to emulate behaviour of Pandas Indexes.

        Notes
        -----
        int, single slice and boolean indexers can only be used on one dimensional XLRanges.

        Will always return an XLRange unless a scalar key is used.
        """
        if is_int_type(key):

            if not self.is_1D:
                raise TypeError("Can only do integer lookups on 1D ranges")

            elif key < 0:
                key += len(self)

            if self.is_row:
                return self.start.translate(0, key)
            else:
                return self.start.translate(key, 0)

        elif isinstance(key, slice):
            key = fill_slice(key)
            if key.step != 1:
                raise TypeError("Can only slice with step equal to 1")

            start, stop = key.start, (key.stop)

            return self[start] - self[stop]

        elif is_bool_indexer(key):

            if not self.is_1D:
                raise TypeError("Can only use Boolean indexers on 1D ranges")

            true_positions = list(i for i, bool_ in enumerate(key) if bool_ == True)

            start = min(true_positions)
            stop = max(true_positions)

            step_is_1 = ((stop - start) + 1) / len(true_positions) == 1

            if not step_is_1:
                raise TypeError("Bool indexers can't have any holes in (i.e. equivalent as slice must have step=1)")

            return self[start:stop]

        elif len(key) == 2:

            if self.is_1D:
                raise TypeError("2 Parameter indexing only available on 2D ranges")

            row_slice, col_slice = key

            if is_int_type(row_slice) and is_int_type(col_slice):

                if row_slice < 0:
                    row_slice += self.shape[0]
                if col_slice < 0:
                    col_slice += self.shape[1]

                return self.start.translate(row_slice, col_slice)

            elif isinstance(row_slice, slice) and isinstance(col_slice, slice):
                row_slice = fill_slice(row_slice)
                col_slice = fill_slice(col_slice)

                row_stop = row_slice.stop
                if row_slice.stop < 0:
                    row_stop += self.shape[0]

                col_stop = col_slice.stop
                if col_slice.stop < 0:
                    col_stop += self.shape[1]

                return XLRange(self.start.translate(row_slice.start, col_slice.start),
                               self.start.translate(row_stop, col_stop))

        raise TypeError("Expecting tuple of slices, boolean indexer, or an index or a slice if 1D, not {}".format(key))

    def iterrows(self):
        """
        Iterate over each row of self, yields each row as XLRange

        Yields
        -------
        XLRange
            corresponding to current row
        """
        if not self.is_2D:
            raise TypeError("Can only call iterrows on 2D ranges")

        for x in range(self.shape[0]):
            yield self[x, 0] - self[x, -1]

    def __eq__(self, other):
        return self.start == other.start and self.stop == other.stop

    def copy(self):
        return self.start.copy() - self.stop.copy()

    def __hash__(self):
        return hash((self.start.row, self.start.col, self.stop.row, self.stop.row, self._sheet))

    def translate(self, row, col):
        """
        Translates whole range by row, col.

        Parameters
        ----------
        row : int
            corresponding to movement in row direction (+ve down the spreadsheet)
        col : int
            corresponding to movement in col direction (+ve right across the spreadsheet)

        Returns
        -------
        XLRange
            new with translation applied

        Notes
        -----
        This moves the whole range, and cannot be used to change the shape of self.
        """
        new = self.copy()

        new.start = new.start.translate(row, col)
        new.stop = new.stop.translate(row, col)

        return new

    def trans(self, row, col):
        """
        Short for XLRange.translate

        See Also
        --------
        XLRange.translate
        """
        return self.translate(row, col)
