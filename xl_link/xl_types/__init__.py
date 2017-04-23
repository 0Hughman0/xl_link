from xlsxwriter.utility import xl_rowcol_to_cell, xl_cell_to_rowcol
from pandas.core.common import is_bool_indexer


def is_int_type(i):
    try:
        return int(i) == i
    except TypeError:
        return False

class XLCell:
    """
    Represents a cell in an Excel spreadsheet
    """

    def __init__(self, sheet, row, col):
        self.sheet = sheet
        self.row = row
        self.col = col

    @classmethod
    def from_cell(cls, cell, sheet= "Sheet 1"):
        return XLCell(sheet, *xl_cell_to_rowcol(cell))

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
        return XLCell(self.sheet, self.row, self.col)

    def translate(self, row, col):
        cell = self.copy()
        cell.row += row or 0
        cell.col += col or 0
        return cell


class XLRange:
    """
    Represents a range in an Excel spreadsheet
    """
    def __init__(self, start, stop):
        assert start.sheet == stop.sheet, "start and stop must be in the same sheet"
        self.sheet = start.sheet
        self.start = start
        self.stop = stop

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
        return self.shape[1] == 1

    @property
    def is_col(self):
        return self.shape[0] == 1

    @property
    def is_1D(self):
        return self.is_row or self.is_col

    def __repr__(self):
        return "<XLRange: {}>".format(self.range)
        
    def __len__(self):
        if self.is_col:
            return self.shape[1]
        if self.is_row:
            return self.shape[0]
        raise TypeError("length is only defined for 1D ranges")

    def __iter__(self):
        if self.is_1D:
            for x in range(len(self)):
                yield self[x]
        else:
            raise TypeError("Can only iterate over 1D ranges")
        
    def __getitem__(self, item):
        if is_int_type(item):
            if item < 0:
                item += len(self)
            if self.is_row:
                return self.start.translate(item, 0)
            if self.is_col:
                return self.start.translate(0, item)
            else:
                raise TypeError("Can only do integer lookups on 1D ranges")

        elif isinstance(item, slice):
            start, stop = item.start, (item.stop or 0) - 1
            return self[start] - self[stop]

        elif is_bool_indexer(item):
            start = min((i for i, bool in enumerate(item) if bool == True))
            stop = max((i for i, bool in enumerate(item) if bool == True))
            return self[start, stop]

        elif len(item) == 2:
            row_slice, col_slice = item
            if is_int_type(row_slice) and is_int_type(col_slice):
                return self.start.translate(row_slice, col_slice)
            elif isinstance(row_slice, slice) and isinstance(col_slice, slice):
                return XLRange(self.start.translate(row_slice.start, col_slice.start),
                               self.stop.translate(-row_slice.stop, -col_slice.stop))
        else:
            raise TypeError("Excpecting tuple of slices, boolean indexer, or an index or a slice if 1D, not {}".format(item))

    def copy(self):
        return self.start.copy() - self.stop.copy()

