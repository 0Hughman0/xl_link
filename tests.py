import unittest


from xl_link import *


class NoOffsetCase(unittest.TestCase):

    def setUp(self):
        self.f = EmbededFrame({"one": range(3, 14), "two": range(15, 26)})
        self.frame_proxy = self.f.to_excel("{}.xlsx".format(self.__class__.__name__), engine="xlsxwriter")

    def test_extreme_cells(self):
        self.assertEqual(self.frame_proxy.iat[0, 0].cell, "B2")
        self.assertEqual(self.frame_proxy.iat[-1, -1].cell, "C12")
        self.assertEqual(self.frame_proxy.xl.range, "{}:{}".format("B2", "C12"))

    def test_series_slices(self):
        slice = self.frame_proxy.loc[2:7, "two"]
        self.assertEqual(slice.xl.range, "C4:C9")
        self.assertEqual(slice.index.xl.range, "A4:A9")

    def test_frame_slices(self):
        slice = self.frame_proxy.loc[3:7, :]
        self.assertEqual(slice.index.xl.range, "A5:A9")
        self.assertEqual(slice.columns.xl.range, "B1:C1")


class OffsetCase(unittest.TestCase):

    def setUp(self):
        self.f = EmbededFrame({"one": range(3, 14), "two": range(15, 26)})
        self.frame_proxy = self.f.to_excel("{}.xlsx".format(self.__class__.__name__), startrow=4, startcol=8, engine="xlsxwriter")

    def test_extreme_cells(self):
        self.assertEqual(self.frame_proxy.iat[0, 0].cell, "J6")
        self.assertEqual(self.frame_proxy.iat[-1, -1].cell, "K16")
        self.assertEqual(self.frame_proxy.xl.range, "{}:{}".format("J6", "K16"))

    def test_series_slices(self):
        slice = self.frame_proxy.loc[2:7, "two"]
        self.assertEqual(slice.xl.range, "K8:K13")
        self.assertEqual(slice.index.xl.range, "I8:I13")

    def test_frame_slices(self):
        slice = self.frame_proxy.loc[3:7, :]
        self.assertEqual(slice.index.xl.range, "I9:I13")
        self.assertEqual(slice.columns.xl.range, "J5:K5")


if __name__ == "__main__":
    unittest.main()

