import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from xl_link import write_frame


TEST_CASE_FOLDER = Path("./tests/test_cases")


def path_for(filename):
    return (TEST_CASE_FOLDER / filename).absolute().as_posix()

idx = pd.IndexSlice

base_frame = pd.DataFrame(columns=("Meal", "Mon", "Tues", "Weds", "Thur"),
                          data={'Meal': ('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                                'Mon': ('Toast', 'Soup', 'Curry', 'Shmores'),
                                'Tues': ('Bagel', 'Something Different!', 'Stew', 'Cookies'),
                                'Weds': ('Cereal', 'Rice', 'Pasta', 'Biscuits'),
                                'Thur': ('Croissant', 'Hotpot', 'Gnocchi', 'Chocolate')})


def case_factory(name, to_excel_kwargs, f):

    class FactoryCase(unittest.TestCase):

        @classmethod
        def setUpClass(cls):
            cls.to_excel_kwargs = to_excel_kwargs
            cls.f = f
            cls.__name__ = name

        def has_index_type(self, type, axis='index'):
            return isinstance(getattr(self.f, axis), type)

        def setUp(self):
            file_name = path_for("{}.xlsx".format(self.__class__.__name__))
            kwargs = {'engine':"openpyxl"}
            kwargs.update(self.to_excel_kwargs)
            self.frame_proxy = write_frame(self.f, file_name, kwargs)
            self.workbook = load_workbook(file_name)

        def check_value(self, value, position, msg=None):
            found_value = self.workbook["Sheet1"][position].value
            with self.subTest(msg=msg, value=value, position=position, found_value=found_value):
                self.assertEqual(value, found_value)

        def check_data(self, frame, xl_range, msg=None):
            for i, f_row in enumerate(frame.iterrows()):
                _, f_row = f_row
                for j, value in enumerate(f_row):
                    self.check_value(value, xl_range[i, j].cell, msg)

        def check_index(self, frame_index, proxy_index, msg=None):
            for value, position in zip(frame_index, proxy_index):
                if self.has_index_type(pd.MultiIndex):
                    value = value[-1]
                self.check_value(value, position.cell, msg)

        def test_data(self):
            self.check_data(self.f, self.frame_proxy.data)

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
                row_slice_row_indexer = idx["Pre-Noon", "Breakfast"]
                col_slice_row_indexer = idx[:, "Lunch":"Midnight Snack"]
            if self.has_index_type(pd.MultiIndex, 'columns'):
                row_slice_col_indexer = idx[:, "Mon":"Weds"]
                col_slice_col_indexer = idx["Late Week", "Thur"]

            row_slice = self.f.loc[row_slice_row_indexer, row_slice_col_indexer]
            proxy_row_slice = self.frame_proxy.loc[row_slice_row_indexer, row_slice_col_indexer]

            for value, position in zip(row_slice, proxy_row_slice):
                self.check_value(value, position.cell, msg="Series row slice check data")

            col_slice = self.f.loc[col_slice_row_indexer, col_slice_col_indexer]
            proxy_col_slice = self.frame_proxy.loc[col_slice_row_indexer, col_slice_col_indexer]

            for value, position in zip(col_slice, proxy_col_slice):
                self.check_value(value, position.cell, msg="Series col slice check data")

        def test_frame_slice(self):
            slice_row_indexer = slice("Lunch", "Dinner")
            slice_col_indexer = slice("Tues", "Weds")

            if self.has_index_type(pd.RangeIndex):
                slice_row_indexer = slice(1, 2)

            if self.has_index_type(pd.MultiIndex):
                slice_row_indexer = (idx[:, "Lunch"], idx[:, "Dinner"])
            if self.has_index_type(pd.MultiIndex, 'columns'):
                slice_col_indexer = (idx[:, "Tues"], idx[:, "Weds"])

            frame_slice = self.f.loc[slice_row_indexer, slice_col_indexer]
            proxy_frame_slice = self.frame_proxy.loc[slice_row_indexer, slice_col_indexer]

            self.check_data(frame_slice, proxy_frame_slice, msg="Frame slice check data")

    obj = FactoryCase
    obj.__name__ = name

    return obj

#  MultiIndex Case setup ###############################################################################################

multi_cols = pd.MultiIndex.from_tuples(tuple(zip( ("Early Week", "Early Week", "Late Week", "Late Week"),
                                                  ("Mon"       , "Tues"      , "Weds"     , "Thur"     ))))
multi_index = pd.MultiIndex.from_tuples(tuple(zip(("Pre-Noon" , "Pre-Noon", "Post-Noon", "Post-Noon"     ),
                                                  ("Breakfast", "Lunch"   , "Dinner"   , "Midnight Snack"))))
multi_f = base_frame.copy()
multi_f.set_index(multi_index, inplace=True)
multi_f.drop("Meal", axis=1, inplace=True)
multi_f.columns = multi_cols
multi_f = multi_f.sort_index().sort_index(axis=1)

########################################################################################################################


NoOffsetCase = case_factory("NoOffsetCase", {}, base_frame)
OffsetCase = case_factory("OffsetCase", {"startrow": 10, "startcol": 13}, base_frame)
TextIndexCase = case_factory("TextIndexCase", {}, base_frame.set_index("Meal", drop=True))
MultiIndexCase = case_factory("MultiIndexCase", {}, multi_f) # Fails one test due to strange index re-ordering
SlicedIndexCase = case_factory("SlicedIndexCase", {"columns": ["Mon", "Tues", "Weds", "Thur"]}, base_frame.set_index("Meal", drop=True))

if __name__ == "__main__":
    unittest.main(verbosity=3)


