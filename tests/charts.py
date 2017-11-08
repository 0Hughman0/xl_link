"""
How the chuff do I automatically test these?

As a result of not answering that question, I have just tried to generate a load of chart combinations, and visually
inspected them. So please take these tests with a pinch of salt!
"""

import unittest
from .tools import path_for, TEST_CASE_FOLDER

import pandas as pd

from xl_link import XLDataFrame


test_frame = XLDataFrame(columns=("Mon", "Tues", "Weds", "Thur"),
                                 index=('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                                 data={'Mon': (15, 20, 12, 3),
                                       'Tues': (5, 16, 3, 0),
                                       'Weds': (3, 22, 2, 8),
                                       'Thur': (6, 7, 1, 9)})

class ChartBaseCase:

    test_frame = test_frame

    type_ = None
    subtype = None

    def setUp(self):
        super().setUp()

    @classmethod
    def setUpClass(cls):
        cls.f = test_frame
        cls.writer = pd.ExcelWriter(path_for('charts', cls.__name__), engine=cls.to_excel_args['engine'])
        cls.xlmap = cls.f.to_excel(cls.writer, **cls.to_excel_args)
        cls.workbook = cls.xlmap.writer.book
        return cls

    @classmethod
    def tearDownClass(cls):
        cls.writer.save()

    def insert_chart(self, chart, sheet):
        if self.xlmap.writer.engine == "xlsxwriter":
            sheet = self.workbook.add_worksheet(sheet)
            sheet.insert_chart('A1', chart)
        else:
            sheet = self.workbook.create_sheet(title=sheet)
            sheet.add_chart(chart, 'A1')

    def test_default_params(self):
        chart = self.xlmap.create_chart(type_=self.type_, subtype=self.subtype)
        self.insert_chart(chart, 'default')

    def test_1val(self):
        chart = self.xlmap.create_chart(type_=self.type_, subtype=self.subtype,
                                        values="Tues")
        self.insert_chart(chart, '1val')

    def test_1val1cat(self):
        chart = self.xlmap.create_chart(type_=self.type_, subtype=self.subtype,
                                        values="Tues", categories="Mon")
        self.insert_chart(chart, '1val1cat')

    def test_2val1cat(self):
        chart = self.xlmap.create_chart(type_=self.type_, subtype=self.subtype,
                                        values=("Mon", "Tues"), categories="Weds")
        self.insert_chart(chart, '2val1cat')

    def test_2val2cat(self):
        chart = self.xlmap.create_chart(type_=self.type_, subtype=self.subtype,
                                        values=("Mon", "Tues"), categories=("Weds", "Thur"))
        self.insert_chart(chart, '2val2cat')

suite = unittest.TestSuite()

CHART_TYPES = ['area', 'bar', 'column', 'line', 'pie', 'doughnut', 'scatter', 'stock', 'radar']

test_cases = []


for chart_type in CHART_TYPES:

    for engine in ("xlsxwriter", "openpyxl"):
        if chart_type == "column" and engine == "openpyxl":
            continue

        class FromFactory(ChartBaseCase, unittest.TestCase):

            type_ = chart_type

            to_excel_args = {'engine': engine}

        FromFactory.__name__ = engine + chart_type + "TestCase"

        testloader = unittest.TestLoader()
        testnames = testloader.getTestCaseNames(FromFactory)
        case_suite = unittest.TestSuite()
        for name in testnames:
            case_suite.addTest(FromFactory(name))

        suite.addTest(case_suite)
