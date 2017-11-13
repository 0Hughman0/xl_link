import sys

sys.path.append('..')

import pandas as pd
from xl_link import XLDataFrame
import numpy as np

# Example data
cat_data = XLDataFrame(index=('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                       data={'Mon': (15, 20, 12, 3),
                                       'Tues': (5, 16, 3, 0),
                                       'Weds': (3, 22, 2, 8),
                                       'Thur': (6, 7, 1, 9)})

xy_data = XLDataFrame(data={'Y1': np.random.rand(10) * 10,
                            'Y2': np.random.rand(10) * 10})

#### Using openpyxl #####################################################################################################

writer = pd.ExcelWriter("OpenPyXLExamples.xlsx", engine="openpyxl") # Allows multiple sheets in same .xlsx file

### Bar Charts #########################################################################################################

xlmap = cat_data.to_excel(writer, sheet_name='Bar')
chart = xlmap.create_chart('bar',
                           ['Mon', 'Tues'], # Get values from 'Mon' and 'Tues' columns
                           title="Bar Example", x_axis_name="Meal", y_axis_name="Calories")

# Insert right of last column -> last column: columns[-1], one to right: .trans(0, 1)
xlmap.sheet.add_chart(chart, xlmap.columns[-1].trans(0, 1).cell)

### Simple Scatter #####################################################################################################

xlmap = xy_data.to_excel(writer, sheet_name='Scatter')
scatter_chart = xlmap.create_chart('scatter', x_axis_name='x', y_axis_name='y', title='Scatter Example')
xlmap.sheet.add_chart(scatter_chart, xlmap.columns[-1].trans(0, 1).cell)

### Highlight values ####################################################################################################

from openpyxl.styles import Font
from openpyxl.styles.colors import RED

xlmap = xy_data.to_excel(writer, sheet_name="Highlighted")

# Iterate through each row in xlsx sheet, and written frame f in parallel
for xlrow, frow in zip(xlmap.data.iterrows(), xlmap.f.iterrows()):
    i, row = frow
    for xlcell, val in zip(xlrow, row):
        if val > 5:
            xlmap.sheet[xlcell.cell].font = Font(color=RED, bold=True)

# insert 2 below bottom of table
xlmap.sheet[xlmap.index[-1].trans(2, 0).cell].value = "Values above 5 highlighted"

writer.save()

#### Using xlsxwriter ##################################################################################################

writer = pd.ExcelWriter("XlsxWriterExamples.xlsx", engine='xlsxwriter') # Allows multiple sheets in same .xlsx file

### Bar Charts #########################################################################################################

xlmap = cat_data.to_excel(writer, sheet_name="Bar")
xl_linked_chart = xlmap.create_chart('column', title="Bar Chart", x_axis_name="Meal", y_axis_name="Calories")

right_of_table = xlmap.columns[-1].trans(0, 1).cell
xlmap.sheet.insert_chart(right_of_table, xl_linked_chart)

### Simple Scatter #####################################################################################################

xlmap = xy_data.to_excel(writer, sheet_name='Scatter')
scatter_chart = xlmap.create_chart('scatter', x_axis_name='x', y_axis_name='y', title='Scatter Example')
xlmap.sheet.insert_chart(xlmap.columns[-1].trans(0, 1).cell, scatter_chart)

### Highlight values ####################################################################################################

xlmap = xy_data.to_excel(writer, sheet_name="Highlighted")

highlight = writer.book.add_format({'bold': True, 'font_color': 'red'})

for xlrow, frow in zip(xlmap.data.iterrows(), xlmap.f.iterrows()):
    i, row = frow
    for xlcell, val in zip(xlrow, row):
        if val > 5:
            xlmap.sheet.write(xlcell.cell, val, highlight)

# insert 2 below bottom of table
xlmap.sheet.write(xlmap.index[-1].trans(2, 0).cell, "Values above 5 highlighted")

writer.save()

#### Use with formulas #################################################################################################

xlmap = cat_data.to_excel("FormulaExamples.xlsx", engine='xlsxwriter')

xlmap.sheet.write(xlmap.index[-1].trans(1, 0).cell, "Sum") # Add to row below
xlmap.sheet.write(xlmap.index[-1].trans(2, 0).cell, "Std Dev") # Add to 2 rows below

for col in xlmap.f.columns:
    col_range = xlmap[col]
    sum_cell = col_range.stop.trans(1, 0)
    stddev_cell = col_range.stop.trans(2, 0)
    xlmap.sheet.write(sum_cell.cell, "=SUM({})".format(col_range.frange))
    xlmap.sheet.write(stddev_cell.cell, "=STDEV({})".format(col_range.frange))

xlmap.writer.save()

#### Use of chart_wrapper.create_chart #################################################################################

### Chart not using columns ############################################################################################

from xl_link.chart_wrapper import create_chart

writer = pd.ExcelWriter("CreateChartExample.xlsx", engine='xlsxwriter')

xlmap = cat_data.to_excel(writer, sheet_name="Non-column")

values = list(xlmap.data.iterrows())
categories = xlmap.columns
names = list(xlmap.f.index)

chart1 = create_chart(writer.book, writer.engine, 'column', values, categories, names)

xlmap.sheet.insert_chart(xlmap.columns.stop.trans(0, 1).cell, chart1)

### 2 DataFrames, one chart ############################################################################################

xy_data2 = xy_data.apply(lambda x: x * 2)

xlmap1 = xy_data.to_excel(writer, sheet_name="2 DataFrames")
xlmap2 = xy_data2.to_excel(writer, sheet_name="2 DataFrames", startcol=xlmap1.columns.stop.col + 2)

chart2 = create_chart(writer.book, writer.engine, 'scatter',
                     (xlmap1['Y1'], xlmap2['Y1']), # Values
                      xlmap1.index, # Categories
                     ('df1-Y1', 'df2-Y1')) # Names

xlmap1.sheet.insert_chart(xlmap2.columns[-1].trans(0, 1).cell, chart2)

writer.save()