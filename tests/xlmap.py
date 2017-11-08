"""
Tests for XLDataFrame and XLMap behaviour (indexing tests can be found in indexers.py).

Note
----
These tests build upon pandas and openpyxl, and work on the assumption that those modules are functional.
"""
import unittest

from xl_link import XLDataFrame

from .tools import XLMapBaseCase, path_for


test_frame = XLDataFrame(columns=("Meal", "Mon", "Tues", "Weds", "Thur"),
                          data={'Meal': ('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                                'Mon': ('Toast', 'Soup', 'Curry', 'Shmores'),
                                'Tues': ('Bagel', 'Something Different!', 'Stew', 'Cookies'),
                                'Weds': ('Cereal', 'Rice', 'Pasta', 'Biscuits'),
                                'Thur': ('Croissant', 'Hotpot', 'Gnocchi', 'Chocolate')})

class XLMapCase(XLMapBaseCase):

    test_frame = test_frame
    to_excel_args = {}
    indexer = ''

    def setUp(self):
        self.f = test_frame
        self.xlmap = self.f.to_excel(path_for(self.__name__), **self.to_excel_args)
        self.workbook = self.xlmap.writer.book
        self.f_indexer = getattr(self.f, self.indexer)
        self.map_indexer = getattr(self.xlmap, self.indexer)

    def check_cell(self, value, position, msg=None):
        found_value = self.workbook["Sheet1"][position.cell].value
        with self.subTest(msg=msg, value=value, position=position, found_value=found_value):
            self.assertEqual(value, found_value)

    def check_series(self, series, xlrange, msg=None):
        for val, cell in zip(series, xlrange):
            self.check_cell(val, cell, msg)

    def check_frame(self, frame, xlrange, msg=None):
        for row, xlrow in zip(frame.iterrows(), xlrange.iterrows()):
            self.check_series(row, xlrow, msg)

    def test_index(self):
        self.check_series(self.f.index, self.xlmap.index)

    def test_columns(self):
        self.check_series(self.f.columns, self.xlmap.columns)

    def test_data(self):
        self.check_frame(self.f, self.xlmap.data)

if __name__ == "__main__":
    unittest.main(verbosity=3)

