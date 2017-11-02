from xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol
from pandas.core.common import is_bool_indexer


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


class XLCell:
    """
    Represents a cell in an Excel spreadsheet
    """

    def __init__(self, row, col, sheet='Sheet1'):
        self.sheet = sheet
        self.row = row
        self.col = col

    @classmethod
    def from_cell(cls, cell, sheet="Sheet 1"):
        return XLCell(*xl_cell_to_rowcol(cell), sheet)

    @property
    def cell(self):
        """
        Gets the Excel cell this object represents
        """
        return xl_rowcol_to_cell(self.row, self.col)

    @property
    def fcell(self):
        """
        Gets the Excel cell this object represents for use in formulas
        """
        return "'{}'!{}".format(self.sheet, self.cell)

    @property
    def rowcol(self):
        return self.row, self.col

    def range_between(self, other):
        return XLRange(self, other)

    def __sub__(self, other):
        return XLRange(self, other)

    def __eq__(self, other):
        if isinstance(other, str):
            return self == XLCell.from_cell(other)
        return self.row == other.row and self.col == other.col

    def __repr__(self):
        return "<XLCell: {}>".format(self.cell)

    def copy(self):
        return XLCell(self.row, self.col, self.sheet)

    def translate(self, row, col):
        cell = self.copy()
        cell.row += row or 0
        cell.col += col or 0
        return cell

    def __hash__(self):
        return hash((self.row, self.col, self.sheet))

class XLRange:
    """
    Represents a range in an Excel spreadsheet
    """
    def __init__(self, start, stop):
        assert start.sheet == stop.sheet, "start and stop must be in the same sheet"
        self.sheet = start.sheet
        self.start = start.copy()
        self.stop = stop.copy()

    @property
    def range(self):
        """
        Gets the Excel range this object represents
        """
        return "{}:{}".format(self.start.cell, self.stop.cell)

    @property
    def frange(self):
        """
        Gets the Excel cell this object represents for use in formulas
        """
        return "'{}'!{}".format(self.sheet, self.range)

    @property
    def rowcol_rowcol(self):
        return self.start.rowcol, self.stop.rowcol

    @property
    def shape(self):
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

    def __repr__(self):
        return "<XLRange: {}>".format(self.range)

    def __len__(self):
        if self.is_col:
            return self.shape[0]
        if self.is_row:
            return self.shape[1]
        raise TypeError("length is only defined for 1D ranges")

    def __iter__(self):
        if self.is_1D:
            for x in range(len(self)):
                yield self[x]
        else:
            raise TypeError("Can only iterate over 1D ranges")

    def __getitem__(self, item):
        if is_int_type(item):

            if not self.is_1D:
                raise TypeError("Can only do integer lookups on 1D ranges")

            elif item < 0:
                item += len(self)

            if self.is_row:
                return self.start.translate(0, item)
            else:
                return self.start.translate(item, 0)

        elif isinstance(item, slice):
            item = fill_slice(item)
            if item.step != 1:
                raise TypeError("Can only slice with step equal to 1")

            start, stop = item.start, (item.stop) - 1

            return self[start] - self[stop]

        elif is_bool_indexer(item):

            if not self.is_1D:
                raise TypeError("Can only use Boolean indexers on 1D ranges")

            true_positions = list(i for i, bool_ in enumerate(item) if bool_ == True)

            start = min(true_positions)
            stop = max(true_positions)

            step_is_1 = (stop - start) / len(true_positions) == 1

            if not step_is_1:
                raise TypeError("Bool indexers can't have any holes in (i.e. equivalent as slice must have step=1)")

            return self[start:stop]

        elif len(item) == 2:

            row_slice, col_slice = item

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

        raise TypeError("Expecting tuple of slices, boolean indexer, or an index or a slice if 1D, not {}".format(item))

    def __eq__(self, other):
        return self.start == other.start and self.stop == other.stop

    def copy(self):
        return self.start.copy() - self.stop.copy()

    def __hash__(self):
        return hash((self.start.row, self.start.col, self.stop.row, self.stop.row, self.sheet))

    def translate(self, row, col):
        new = self.copy()

        new.start = new.start.translate(row, col)
        new.stop = new.stop.translate(row, col)

        return new

