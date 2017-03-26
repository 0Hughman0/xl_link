from xlsxwriter.utility import xl_rowcol_to_cell


class XLCell:
    """
    Represents a cell in an Excel spreadsheet
    """

    def __init__(self, sheet, row, col):
        self.sheet = sheet
        self.row = row
        self.col = col

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
        return "{}!{}".format(self.sheet, self.cell)

    @property
    def rowcol(self):
        return self.row, self.col

    def range_between(self, other):
        return XLRange(self, other)

    def __sub__(self, other):
        return XLRange(self, other)

    def __eq__(self, other):
        return self.row == other.row and self.col == other.col

    def __repr__(self):
        return self.cell

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
        return "{}!{}".format(self.sheet, self.range)

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
        return self.is_row and self.is_col

    def __repr__(self):
        return self.range

    def __getitem__(self, item):
        if isinstance(item, slice):
            start, stop = item.start, item.stop - 1
            if self.is_row:
                return self.start.translate(start, 0) - self.start.translate(stop, 0)
            elif self.is_col:
                return self.start.translate(0, start) - self.start.translate(0, stop)
        if len(item) == 2:
            row_slice, col_slice = item
            if isinstance(row_slice, int) and isinstance(col_slice, int):
                return self.start.translate(row_slice, col_slice)
            elif isinstance(row_slice, slice) and isinstance(col_slice, slice):
                return XLRange(self.start.translate(row_slice.start, col_slice.start),
                               self.stop.translate(-row_slice.stop, -col_slice.stop))
        else:
            raise TypeError("Excpecting tuple of slices or indexes, or a slice if 1D")

    def copy(self):
        return self.start.copy() - self.stop.copy()

