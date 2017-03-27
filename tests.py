import unittest

import pandas as pd

from xl_link import EmbededFrame

idx = pd.IndexSlice

base_frame = pd.DataFrame({'Meal': ('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                            'Mon': ('Toast', 'Soup', 'Curry', 'Shmores'),
                            'Tues': ('Toast', 'Something Different!', 'Curry', 'Shmores'),
                            'Weds': ('Toast', 'Soup', 'Curry', 'Biscuits'),
                            'Thur': ('Toast', 'Hotpot', 'Curry', 'Chocolate')})


class NoOffsetCase(unittest.TestCase):

    def setUp(self):
        self.f = EmbededFrame(base_frame.copy())
        self.frame_proxy = self.f.to_excel("{}.xlsx".format(self.__class__.__name__), engine="xlsxwriter")

    def test_extreme_cells(self):
        self.assertEqual(self.frame_proxy.iat[0, 0].cell, "B2")
        self.assertEqual(self.frame_proxy.iat[-1, -1].cell, "F5")
        self.assertEqual(self.frame_proxy.xl.range, "{}:{}".format("B2", "F5"))

    def test_series_slices(self):
        slice = self.frame_proxy.loc[:, "Weds"]
        self.assertEqual(slice.xl.range, "F2:F5")
        self.assertEqual(slice.index.xl.range, "A2:A5")

    def test_frame_slices(self):
        slice = self.frame_proxy.loc[1:2, "Mon":"Tues"]
        self.assertEqual(slice.xl.range, "C3:E4")
        self.assertEqual(slice.index.xl.range, "A3:A4")
        self.assertEqual(slice.columns.xl.range, "C1:E1")


class OffsetCase(unittest.TestCase):

    def setUp(self):
        self.f = EmbededFrame(base_frame.copy())
        self.frame_proxy = self.f.to_excel("{}.xlsx".format(self.__class__.__name__), startrow=4, startcol=8, engine="xlsxwriter")

    def test_extreme_cells(self):
        self.assertEqual(self.frame_proxy.iat[0, 0].cell, "J6")
        self.assertEqual(self.frame_proxy.iat[-1, -1].cell, "N9")
        self.assertEqual(self.frame_proxy.xl.range, "{}:{}".format("J6", "N9"))

    def test_series_slices(self):
        # Col
        slice = self.frame_proxy.loc[1:3, "Meal"]
        self.assertEqual(slice.xl.range, "J7:J9")
        self.assertEqual(slice.index.xl.range, "I7:I9")
        # Row
        slice = self.frame_proxy.loc[2, "Mon":"Thur"]
        self.assertEqual(slice.xl.range, "K8:L8")
        self.assertEqual(slice.index.xl.range, "K5:L5")
        # Col by index
        slice = self.frame_proxy.iloc[1:4, 0]
        self.assertEqual(slice.xl.range, "J7:J9")
        self.assertEqual(slice.index.xl.range, "I7:I9")
        # Row by index
        slice = self.frame_proxy.iloc[2, 1:3]
        self.assertEqual(slice.xl.range, "K8:L8")
        self.assertEqual(slice.index.xl.range, "K5:L5")

    def test_frame_slices(self):
        # By label
        slice = self.frame_proxy.loc[1:2, "Thur":"Weds"]
        self.assertEqual(slice.xl.range, "L7:N8")
        self.assertEqual(slice.index.xl.range, "I7:I8")
        self.assertEqual(slice.columns.xl.range, "L5:N5")
        # By index
        slice = self.frame_proxy.iloc[1:3, 2:5]
        self.assertEqual(slice.xl.range, "L7:N8")
        self.assertEqual(slice.index.xl.range, "I7:I8")
        self.assertEqual(slice.columns.xl.range, "L5:N5")


class TextIndexCase(unittest.TestCase):

    def setUp(self):
        self.f = EmbededFrame(base_frame.copy())
        self.f.set_index("Meal", drop=True, inplace=True)
        self.frame_proxy = self.f.to_excel("{}.xlsx".format(self.__class__.__name__), startrow=4, startcol=8, engine="xlsxwriter")

    def test_extreme_cells(self):
        self.assertEqual(self.frame_proxy.iat[0, 0].cell, "J6")
        self.assertEqual(self.frame_proxy.iat[-1, -1].cell, "M9")
        self.assertEqual(self.frame_proxy.xl.range, "{}:{}".format("J6", "M9"))

    def test_series_slices(self):
        # Col
        slice = self.frame_proxy.loc["Lunch":"Midnight Snack", "Thur"]
        self.assertEqual(slice.xl.range, "K7:K9")
        self.assertEqual(slice.index.xl.range, "I7:I9")
        # Row
        slice = self.frame_proxy.loc["Lunch", "Mon":"Thur"]
        self.assertEqual(slice.xl.range, "J7:K7")
        # Col by index
        slice = self.frame_proxy.iloc[1:4, 0]
        self.assertEqual(slice.xl.range, "J7:J9")
        self.assertEqual(slice.index.xl.range, "I7:I9")
        # Row by index
        slice = self.frame_proxy.iloc[2, 1:3]
        self.assertEqual(slice.xl.range, "K8:L8")
        self.assertEqual(slice.index.xl.range, "K5:L5")

    def test_frame_slices(self):
        # By label
        slice = self.frame_proxy.loc["Lunch":"Dinner", "Thur":"Weds"]
        self.assertEqual(slice.xl.range, "K7:M8")
        self.assertEqual(slice.index.xl.range, "I7:I8")
        self.assertEqual(slice.columns.xl.range, "K5:M5")
        # By index
        slice = self.frame_proxy.iloc[1:3, 1:4]
        self.assertEqual(slice.xl.range, "K7:M8")
        self.assertEqual(slice.index.xl.range, "I7:I8")
        self.assertEqual(slice.columns.xl.range, "K5:M5")


class MultiIndexCase(unittest.TestCase):

    def setUp(self):
        self.f = EmbededFrame(base_frame.copy())

        multi_cols = pd.MultiIndex.from_tuples(tuple(zip(("Early Week", "Early Week", "Late Week", "Late Week"),
                                              ("Mon"       , "Thur"      , "Tues"     , "Weds"     ))))
        multi_index = pd.MultiIndex.from_tuples(tuple(zip(("Pre-Noon", "Pre-Noon", "Post-Noon", "Post-Noon"),
                                                    ("Breakfast"       , "Lunch"      , "Dinner"     , "Midnight Snack"     ))))
        self.f.set_index(multi_index, inplace=True)
        self.f.drop("Meal", axis=1, inplace=True)
        self.f.columns = multi_cols
        self.f.sortlevel(inplace=True)
        self.f.sortlevel(axis=1, inplace=True)
        self.frame_proxy = self.f.to_excel("{}.xlsx".format(self.__class__.__name__), startrow=4, startcol=8, engine="xlsxwriter")

    def test_extreme_cells(self):
        self.assertEqual(self.frame_proxy.iat[0, 0].cell, "K8")
        self.assertEqual(self.frame_proxy.iat[-1, -1].cell, "N11")
        self.assertEqual(self.frame_proxy.xl.range, "{}:{}".format("K8", "N11"))

    def test_series_slices(self):
        # Col
        slice = self.frame_proxy.loc[idx["Pre-Noon", :], idx["Early Week", "Thur"]]
        self.assertEqual(slice.xl.range, "L10:L11")
#        self.assertEqual(slice.index.xl.range, "J10:J11")
        # Row
        slice = self.frame_proxy.loc[idx["Pre-Noon", "Lunch"], :]
        self.assertEqual(slice.xl.range, "K9:N9")
        self.assertEqual(slice.index.xl.range, "K6:N6")
        # Col by index
        slice = self.frame_proxy.iloc[1:4, 0]
        self.assertEqual(slice.xl.range, "J7:J9")
        self.assertEqual(slice.index.xl.range, "I7:I9")
        # Row by index
        slice = self.frame_proxy.iloc[2, 1:3]
        self.assertEqual(slice.xl.range, "K8:L8")
        self.assertEqual(slice.index.xl.range, "K5:L5")

    def test_frame_slices(self):
        # By label
        slice = self.frame_proxy.loc[idx["Post-Noon", "Dinner"]:idx["Pre-Noon", "Breakfast"], idx["Early Week", :]]
        print(slice)
        self.assertEqual(slice.xl.range, "K8:L10")
        self.assertEqual(slice.index.xl.range, "J8:J10")
#        self.assertEqual(slice.columns.xl.range, "K5:L5")
        # By index
        slice = self.frame_proxy.iloc[1:3, 1:4]
        self.assertEqual(slice.xl.range, "K7:M8")
        self.assertEqual(slice.index.xl.range, "I7:I8")
        self.assertEqual(slice.columns.xl.range, "K5:M5")

if __name__ == "__main__":
    unittest.main(verbosity=2)

