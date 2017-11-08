from pathlib import Path
from openpyxl import load_workbook


TEST_CASE_FOLDER = Path("./tests/test_cases")


def path_for(*args):
    *path, file = args
    file += '.xlsx'
    return Path(TEST_CASE_FOLDER , *path, file).absolute().as_posix()

class XLMapBaseCase:

    subdir = 'indexers'
    test_frame = None
    to_excel_args = {}

    def setUp(self):
        self.f = self.test_frame
        filename = path_for(self.subdir, self.__class__.__name__)
        self.xlmap = self.f.to_excel(filename, **self.to_excel_args)
        self.sheet = load_workbook(filename)[self.to_excel_args.get('sheet_name', 'Sheet1')]

    def check_cell(self, value, position, msg=None):
        found_value = self.sheet[position.cell].value
        with self.subTest(msg=msg, value=value, position=position, found_value=found_value):
            self.assertEqual(value, found_value)

    def check_series(self, series, xlrange, msg=None):
        for val, cell in zip(series.values, xlrange):
            self.check_cell(val, cell, msg)

    def check_frame(self, frame, xlrange, msg=None):
        for row, xlrow in zip(frame.iterrows(), xlrange.iterrows()):
            _, row = row
            self.check_series(row, xlrow, msg)