"""
Tests for XLRange and XLCell

Note
----
The XLRange and XLCell classes build on code from xlsxwriter (see xlsx writer folder), functions from xlsxwriter
are not tested, but are used.

These tests build upon openpyxl, and work on the assumption that those modules are functional.
"""

import json
import unittest
from abc import abstractmethod

import numpy as np
from openpyxl import load_workbook
from openpyxl.chart.reference import Reference

from xl_link import xl_types
from xl_link.xlsxwriter.utility import xl_rowcol_to_cell

test_workbook = load_workbook(r"./tests/test_cases/XLTypesTestGrid.xlsx")


def process_cell(cell_text):
    return json.loads(cell_text)


class XLCellConstructorTestCase(unittest.TestCase):

    sheet = 'testSheet'
    row = 4
    col = 3
    xlcell = xl_rowcol_to_cell(row, col)

    def setUp(self):
        self.from_init = xl_types.XLCell(self.row, self.col, self.sheet)
        self.from_cell = xl_types.XLCell.from_cell(self.xlcell, self.sheet)

    def test_init(self):
        self.assertEqual(self.from_init.cell, self.xlcell)
        self.assertEqual(self.from_init.fcell, "'{}'!{}".format(self.sheet, self.xlcell))
        self.assertEqual(self.from_init, self.from_cell)

    def test_from_cell(self):
        self.assertEqual(self.from_cell.rowcol, (self.row, self.col))
        self.assertEqual(self.from_cell.fcell, "'{}'!{}".format(self.sheet, self.xlcell))
        self.assertEqual(self.from_cell, self.from_init)


def xl_cell_case_factory(row, col, sheetname):

    xlcell = xl_rowcol_to_cell(row, col)
    rowcol = (row, col)

    class XLCellCaseFactory(unittest.TestCase):

        def setUp(self):
            self.cell = xl_types.XLCell(row, col, sheetname)

        def test_cell_location(self):
            self.assertSequenceEqual(rowcol, [self.cell.row, self.cell.col])
            self.assertEqual(xlcell, self.cell.cell)
            self.assertSequenceEqual(rowcol, self.cell.rowcol)

        def test_fcell_form(self):
            try:
                Reference(range_string=self.cell.fcell)
            except ValueError as e:
                self.fail("Cannot use fcell in formula: {}".format(e))

        def test_translate(self):
            col_translate = (0, 3)
            row_translate = (5, 0)
            rowcol_tranlate = (4, 7)

            for translation in (col_translate, row_translate, rowcol_tranlate):
                row_t, col_t = translation
                new_position = self.cell.translate(row_t, col_t)
                new_position_cell = xl_rowcol_to_cell(row_t + row, col_t + col)
                self.assertEqual(new_position_cell, new_position.cell)
                # Check has not mutated original obj
                self.assertNotEqual(new_position_cell, self.cell.cell)

        def test_equal(self):
            other = xl_types.XLCell(row, col, sheetname)
            self.assertEqual(self.cell, other)
            self.assertEqual(self.cell.cell, xlcell)
            self.assertNotEqual(self.cell, self.cell.translate(0, 1))
            self.assertNotEqual(self.cell, self.cell.translate(1, 0))

        def test_to_xlrange(self):
            col_translate = (0, 6)
            row_translate = (8, 0)
            rowcol_tranlate = (1, 4)
            for translation in (col_translate, row_translate, rowcol_tranlate):
                row_t, col_t = translation
                other = self.cell.translate(row_t, col_t)
                for xlrange in (self.cell - other, self.cell.range_between(other)):
                    self.assertEqual(xlrange.start.cell, self.cell.cell)
                    self.assertEqual(xlrange.stop.cell, other.cell)

    obj = XLCellCaseFactory
    obj.__name__ = "XLCell{}by{}{}Case".format(row, col, sheetname)

    return obj


def xl_range_case_factory(start_rowcol, stop_rowcol, sheetname):

    start_cell = xl_types.XLCell(*start_rowcol, sheetname)
    stop_cell = xl_types.XLCell(*stop_rowcol, sheetname)

    is_2D = True
    is_col = False
    is_row = False

    if start_cell.row == stop_cell.row:
        is_2D = False
        is_row = True
    elif start_cell.col == stop_cell.col:
        is_2D = False
        is_col = True

    class XLRangeBaseFactoryCase(unittest.TestCase):

        def setUp(self):
            self.range = xl_types.XLRange(start_cell, stop_cell)

        def test_range_location(self):
            self.assertEqual(self.range.start, start_cell)
            self.assertEqual(self.range.stop, stop_cell)

            self.assertEqual(self.range.range, '{}:{}'.format(start_cell.cell, stop_cell.cell))
            self.assertEqual(self.range.rowcol_rowcol, (start_cell.rowcol, stop_cell.rowcol))

        def test_frange_form(self):
            try:
                Reference(range_string=self.range.frange)
            except ValueError as e:
                self.fail("Cannot use fcell in formula: {}".format(e))

        def test_shape(self):
            self.assertEqual(self.range.is_col, is_col)
            self.assertEqual(self.range.is_row, is_row)
            self.assertEqual(self.range.is_1D, is_col or is_row)

        @abstractmethod
        def test_len(self):
            pass

        @abstractmethod
        def test_iter(self):
            pass

        @abstractmethod
        def test_getitem_scalar(self):
            pass

        @abstractmethod
        def test_getitem_slice(self):
            pass

        @abstractmethod
        def test_getitem_bool_indexer(self):
            pass

        @abstractmethod
        def test_getitem_tuple(self):
            pass

    class XLRange2DFactory(XLRangeBaseFactoryCase):

        def test_len(self):
            self.assertRaises(TypeError, len, self.range)

        def test_iter(self):
            self.assertRaises(TypeError, lambda x: next(iter(x)), self.range)

        def test_getitem_scalar(self):
            self.assertRaises(TypeError, self.range.__getitem__, 0)
            self.assertRaises(TypeError, self.range.__getitem__, -1)
            self.assertRaises(TypeError, self.range.__getitem__, 4)

        def test_getitem_slice(self):
            self.assertRaises(TypeError, self.range.__getitem__, slice(0, 3))
            self.assertRaises(TypeError, self.range.__getitem__, slice(-1, -4))
            self.assertRaises(TypeError, self.range.__getitem__, slice(4, 5))

        def test_getitem_bool_indexer(self):
            try:
                self.range[np.array([False, True, True, True], dtype=bool)]
            except Exception as e:
                self.assertEqual(str(e), "Can only use Boolean indexers on 1D ranges")

        def test_getitem_tuple(self):
            top_left = self.range[0, 0]
            bottom_right = self.range[-1, -1]

            self.assertEqual(top_left, self.range.start)
            self.assertEqual(bottom_right, self.range.stop)

            halfway = (int(self.range.shape[0] / 2), int(self.range.shape[1] / 2))

            middle = self.range[halfway[0], halfway[1]]
            middle_cell = start_cell.translate(*halfway)
            self.assertEqual(middle_cell, middle)

            whole_slice = self.range[:, :]
            self.assertEqual(whole_slice, self.range)

            mid_slice = self.range[1:-2, 1:-2]
            one_in = self.range.start.translate(1, 1) - self.range.stop.translate(-1, -1)
            self.assertEqual(mid_slice, one_in)

            self.assertRaises(TypeError, self.range.__getitem__, "A1")

        def test_iterrows(self):
            start_cell
            top_right = start_cell.copy()
            top_right.col = stop_cell.col

            for row_range in self.range.iterrows():
                self.assertEqual(row_range, start_cell-top_right)
                start_cell.row += 1
                top_right.row += 1


    class XLRangeColFactory(XLRangeBaseFactoryCase):

        def test_len(self):
            self.assertEqual(len(self.range), (stop_cell.row - start_cell.row) + 1)

        def test_iter(self):
            start_row = start_cell.row
            start_col = start_cell.col
            for i, cell in enumerate(self.range):
                self.assertEqual(cell.cell, xl_rowcol_to_cell(start_row + i, start_col))

        def test_getitem_scalar(self):
            self.assertEqual(self.range[0], start_cell)
            self.assertEqual(self.range[-1], stop_cell)

            mid_cell = start_rowcol[0] + 4
            self.assertEqual(self.range[4].cell, xl_rowcol_to_cell(mid_cell, start_rowcol[1]))

        def test_getitem_slice(self):
            pass


    if is_2D:
        obj = XLRange2DFactory
    elif is_row:
        obj = XLRangeRowFactory
    elif is_col:
        obj = XLRangeColFactory

    obj.__name__ = 'XLRange{}to{}{}Case'.format(start_cell.cell, stop_cell.cell, sheetname)

    return obj


XLCell0by0Sheet1Case = xl_cell_case_factory(0, 0, 'Sheet1')
XLCell8by4Sheet1Case = xl_cell_case_factory(8, 4, 'Sheet1')
XLCell34by20Sheet1Case = xl_cell_case_factory(34, 20, 'Sheet1')
XLCell73by103DifferentSheetCase = xl_cell_case_factory(73, 103, 'DifferentSheet')

XLRange2DCase = xl_range_case_factory((3, 2), (6, 8), 'Sheet1')
XLRangeRowCase = xl_range_case_factory((5, 10), (10, 10), 'Sheet1')