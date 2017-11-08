"""
Tests for XLMap Indexers

Note
----
These tests build upon pandas and openpyxl, and work on the assumption that those modules are functional.
"""

import unittest
from pathlib import Path

from xl_link import XLDataFrame

from .tools import XLMapBaseCase, path_for

TEST_CASE_FOLDER = Path("./tests/test_cases")


def path_for(filename):
    return (TEST_CASE_FOLDER / 'indexers' / (filename + '.xlsx')).absolute().as_posix()

test_frame = XLDataFrame(columns=("Meal", "Mon", "Tues", "Weds", "Thur"),
                          data={'Meal': ('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                                'Mon': ('Toast', 'Soup', 'Curry', 'Shmores'),
                                'Tues': ('Bagel', 'Something Different!', 'Stew', 'Cookies'),
                                'Weds': ('Cereal', 'Rice', 'Pasta', 'Biscuits'),
                                'Thur': ('Croissant', 'Hotpot', 'Gnocchi', 'Chocolate')})

class GetItemIndexerCase(XLMapBaseCase, unittest.TestCase):

    test_frame = test_frame

    def test_get_single_column(self):
        for col in self.f.columns:
            self.check_series(self.f[col], self.xlmap[col])

    def test_get_multiple_columns(self):
        column_names = list(self.f.columns[1:-1])
        cols = self.f[column_names]
        self.check_frame(cols, self.xlmap[column_names])


class BaseIndexerCase(XLMapBaseCase):

    indexer = None

    def setUp(self):
        super().setUp()
        self.f_indexer = getattr(self.f, self.indexer)
        self.map_indexer = getattr(self.xlmap, self.indexer)


class AtIndexerCase(BaseIndexerCase, unittest.TestCase):

    test_frame = test_frame
    indexer = 'at'

    def test_single_key(self):
        for index in self.f.index:
            for column in self.f.columns:
                self.check_cell(self.f_indexer[index, column], self.map_indexer[index, column])


class iAtIndexerCase(BaseIndexerCase, unittest.TestCase):

    test_frame = test_frame
    indexer = 'iat'

    def test_single_key(self):
        for index in range(self.f.index.size):
            for column in range(self.f.columns.size):
                self.check_cell(self.f_indexer[index, column], self.map_indexer[index, column])


class LocIndexerCase(AtIndexerCase):

    test_frame = test_frame
    indexer = 'loc'

    def test_slice_key(self):
        for index in self.f.index:
            self.check_series(self.f_indexer[index, :], self.map_indexer[index, :])

        for column in self.f.columns:
            self.check_series(self.f_indexer[:, column], self.map_indexer[:, column])


class iLocIndexerCase(iAtIndexerCase):

    test_frame = test_frame
    indexer = 'iloc'

    def test_slice_loc(self):
        for index in range(self.f.index.size):
            self.check_series(self.f_indexer[index, :], self.map_indexer[index, :])

        for column in range(self.f.columns.size):
            self.check_series(self.f_indexer[:, column], self.map_indexer[:, column])


if __name__ == "__main__":
    unittest.main(verbosity=3)
