import sys

sys.path.append('..')

import pandas as pd
from xl_link import XLDataFrame
import numpy as np


cat_data = XLDataFrame(columns=("Mon", "Tues", "Weds", "Thur"),
                       index=('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                       data={'Mon': (15, 20, 12, 3),
                                       'Tues': (5, 16, 3, 0),
                                       'Weds': (3, 22, 2, 8),
                                       'Thur': (6, 7, 1, 9)})

xy_data = XLDataFrame(data={'Y1': np.random.rand(10) * 10,
                            'Y2': np.random.rand(10) * 10})

#### Using openpyxl #####################################################################################################

writer = pd.ExcelWriter("OpenPyXLExamples.xlsx", engine="openpyxl")

### Bar Charts #########################################################################################################

xlmap = cat_data.to_excel(writer, sheet_name='Bar')
chart = xlmap.create_chart('bar', ['Mon', 'Tues'], title="Bar Example", x_axis_name="Meal", y_axis_name="Calories")

xlmap.sheet.add_chart(chart, xlmap.columns[-1].translate(0, 1).cell)

### Simple Scatter #####################################################################################################

xlmap = xy_data.to_excel(writer, sheet_name='Scatter')
scatter_chart = xlmap.create_chart('scatter', x_axis_name='x', y_axis_name='y', title='Scatter Example')
xlmap.sheet.add_chart(scatter_chart, xlmap.columns[-1].translate(0, 1).cell)

### Highlight values ####################################################################################################

from openpyxl.styles import Font
from openpyxl.styles.colors import RED

xlmap = xy_data.to_excel(writer, sheet_name="Highlighted")

for xlrow, frow in zip(xlmap.data.iterrows(), xlmap.f.iterrows()):
    i, row = frow
    for xlcell, val in zip(xlrow, row):
        if val > 5:
            xlmap.sheet[xlcell.cell].font = Font(color=RED, bold=True)

# insert 2 below bottom of table
xlmap.sheet[xlmap.index[-1].translate(2, 0).cell].value = "Values above 5 highlighted"

writer.save()

#### Using xlsxwriter ##################################################################################################

writer = pd.ExcelWriter("XlsxWriterExamples.xlsx", engine='xlsxwriter')

### Bar Charts #########################################################################################################

xlmap = cat_data.to_excel(writer, sheet_name="Bar") # returns the 'ProxyFrame'
xl_linked_chart = xlmap.create_chart('column', title="Bar Chart", x_axis_name="Meal", y_axis_name="Calories")

right_of_table = xlmap.columns[-1].translate(0, 1).cell
xlmap.sheet.insert_chart(right_of_table, xl_linked_chart)

### Simple Scatter #####################################################################################################

xlmap = xy_data.to_excel(writer, sheet_name='Scatter')
scatter_chart = xlmap.create_chart('scatter', x_axis_name='x', y_axis_name='y', title='Scatter Example')
xlmap.sheet.insert_chart(xlmap.columns[-1].translate(0, 1).cell, scatter_chart)

### Highlight values ####################################################################################################

xlmap = xy_data.to_excel(writer, sheet_name="Highlighted")

highlight = writer.book.add_format({'bold': True, 'font_color': 'red'})

for xlrow, frow in zip(xlmap.data.iterrows(), xlmap.f.iterrows()):
    i, row = frow
    for xlcell, val in zip(xlrow, row):
        if val > 5:
            xlmap.sheet.write(xlcell.cell, val, highlight)

# insert 2 below bottom of table
xlmap.sheet.write(xlmap.index[-1].translate(2, 0).cell, "Values above 5 highlighted")

writer.save()

#### Use with formulas #################################################################################################

xlmap = cat_data.to_excel("FormulaExamples.xlsx", engine='xlsxwriter')

xlmap.sheet.write(xlmap.index[-1].translate(1, 0).cell, "Sum")
xlmap.sheet.write(xlmap.index[-1].translate(2, 0).cell, "Std Dev")

for col in xlmap.f.columns:
    col_range = xlmap[col]
    sum_cell = col_range.stop.translate(1, 0)
    stddev_cell = col_range.stop.translate(2, 0)
    xlmap.sheet.write(sum_cell.cell, "=SUM({})".format(col_range.frange))
    xlmap.sheet.write(stddev_cell.cell, "=STDEV({})".format(col_range.frange))

xlmap.writer.save()