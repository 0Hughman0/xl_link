import unittest
from collections import OrderedDict

import pandas as pd

from xl_link import EmbededFrame

from openpyxl import load_workbook

TEST_CASE_FOLDER = ".\\test_cases\\{}"


def path_for(filename):
    return TEST_CASE_FOLDER.format(filename)

idx = pd.IndexSlice

base_frame = pd.DataFrame(columns=("Meal", "Mon", "Tues", "Weds", "Thur"),
                          data={'Meal': ('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                                'Mon': ('Toast', 'Soup', 'Curry', 'Shmores'),
                                'Tues': ('Bagel', 'Something Different!', 'Stew', 'Cookies'),
                                'Weds': ('Cereal', 'Rice', 'Pasta', 'Biscuits'),
                                'Thur': ('Croissant', 'Hotpot', 'Gnocchi', 'Chocolate')})

print("\n\nTest Frame:\n\n{}\n\n".format(base_frame))


def case_factory(name, to_excel_args, to_excel_kwargs, f):

    class FactoryCase(unittest.TestCase):

        @classmethod
        def setUpClass(cls):
            cls.to_excel_args = to_excel_args
            cls.to_excel_kwargs = to_excel_kwargs
            cls.f = EmbededFrame(f)
            cls.__name__ = name

        def has_index_type(self, type):
            return isinstance(self.f.index, type)

        def setUp(self):
            file_name = path_for("{}.xlsx".format(self.__class__.__name__))
            self.frame_proxy = self.f.to_excel(file_name, *self.to_excel_args, engine="xlsxwriter", **self.to_excel_kwargs)
            self.workbook = load_workbook(file_name)

        def check_value(self, value, position, msg=None):
            found_value = self.workbook["Sheet1"][position].value
            with self.subTest(msg=msg, value=value, position=position, found_value=found_value):
                self.assertEqual(value, found_value)

        def check_data(self, frame, frame_proxy, msg=None):
            for f_row, p_row in zip(frame.iterrows(), frame_proxy.iterrows()):
                _, f_row = f_row
                _, p_row = p_row
                for value, position in zip(f_row, p_row):
                    self.check_value(value, position.cell, msg)

        def check_index(self, frame_index, proxy_index, msg=None):

            for value, position in zip(frame_index, proxy_index.xl):
                if self.has_index_type(pd.MultiIndex):
                    value = value[-1]
                self.check_value(value, position.cell, msg)

        def test_data(self):
            self.check_data(self.f, self.frame_proxy)

        def test_columns(self):
            self.check_index(self.f.columns, self.frame_proxy.columns)

        def test_index(self):
            self.check_index(self.f.index, self.frame_proxy.index)

        def test_series_slice(self):
            row_slice_row_indexer = "Breakfast"
            row_slice_col_indexer = slice("Mon", "Weds")

            col_slice_row_indexer = slice("Lunch", "Midnight Snack")
            col_slice_col_indexer = "Thur"

            if self.has_index_type(pd.RangeIndex):
                row_slice_row_indexer = 0
                col_slice_row_indexer = slice(1, 3)

            if self.has_index_type(pd.MultiIndex):
                row_slice_row_indexer = ("Pre-Noon", "Breakfast")
                col_slice_col_indexer = ("Late Week", "Thur")

            row_slice = self.f.loc[row_slice_row_indexer][row_slice_col_indexer]
            proxy_row_slice = self.frame_proxy.loc[row_slice_row_indexer][row_slice_col_indexer]

            self.check_index(row_slice.index, proxy_row_slice.index, msg="Series row slice check index")

            for value, position in zip(row_slice, proxy_row_slice):
                self.check_value(value, position.cell, msg="Series row slice check data")

            col_slice = self.f.loc[col_slice_row_indexer][col_slice_col_indexer]
            proxy_col_slice = self.frame_proxy.loc[col_slice_row_indexer][col_slice_col_indexer]

            self.check_index(col_slice.index, proxy_col_slice.index, msg="Series col slice check index")

            for value, position in zip(col_slice, proxy_col_slice):
                self.check_value(value, position.cell, msg="Series col slice check data")

        def test_frame_slice(self):
            slice_row_indexer = slice("Lunch", "Dinner")
            slice_col_indexer = slice("Tues", "Weds")

            if self.has_index_type(pd.RangeIndex):
                slice_row_indexer = slice(1, 2)

            frame_slice = self.f.loc[slice_row_indexer, slice_col_indexer]
            proxy_frame_slice = self.frame_proxy.loc[slice_row_indexer, slice_col_indexer]

            self.check_data(frame_slice, proxy_frame_slice, msg="Frame slice check data")
            self.check_index(frame_slice.index, proxy_frame_slice.index, msg="Frame slice check index")
            self.check_index(frame_slice.columns, proxy_frame_slice.columns, msg="Frame slice check columns")

    return FactoryCase


NoOffsetCase = case_factory("NoOffsetCase", [], {}, base_frame)
OffsetCase = case_factory("OffsetCase", [], {"startrow": 10, "startcol": 13}, base_frame)
TextIndexCase = case_factory("TextIndexCase", [], {}, base_frame.set_index("Meal", drop=True))

# MultiIndex Case setup
multi_cols = pd.MultiIndex.from_tuples(tuple(zip( ("Early Week", "Early Week", "Late Week", "Late Week"),
                                                  ("Mon"       , "Tues"      , "Weds"     , "Thur"     ))))
multi_index = pd.MultiIndex.from_tuples(tuple(zip(("Pre-Noon" , "Pre-Noon", "Post-Noon", "Post-Noon"     ),
                                                  ("Breakfast", "Lunch"   , "Dinner"   , "Midnight Snack"))))
multi_f = base_frame.copy()
multi_f.set_index(multi_index, inplace=True)
multi_f.drop("Meal", axis=1, inplace=True)
multi_f.columns = multi_cols
multi_f.sortlevel(inplace=True)
multi_f.sortlevel(axis=1, inplace=True)

print("\n\nMultiIndex Frame:\n\n{}\n\n".format(multi_f))

MultiIndexCase = case_factory("MultiIndexCase", [], {}, multi_f)

SlicedIndexCase = case_factory("SlicedIndexCase", [], {"index": ["Breakfast", "Lunch"]}, base_frame.set_index("Meal", drop=True))

if __name__ == "__main__":
    unittest.main(verbosity=1)


