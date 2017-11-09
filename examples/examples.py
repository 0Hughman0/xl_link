import pandas as pd
from xl_link import XLDataFrame
import random


calories_per_meal = XLDataFrame(columns=("Mon", "Tues", "Weds", "Thur"),
                                 index=('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                                 data={'Mon': (15, 20, 12, 3),
                                       'Tues': (5, 16, 3, 0),
                                       'Weds': (3, 22, 2, 8),
                                       'Thur': (6, 7, 1, 9)})

#### Using openpyxl #####################################################################################################

writer = pd.ExcelWriter("OpenPyXLExample.xlsx", engine="openpyxl")
xlmap = calories_per_meal.to_excel(writer, sheet_name='with openpyxl')

chart = xlmap.create_chart('bar', ['Mon', 'Tues'], title="Bar Chart", x_axis_name="Meal", y_axis_name="Calories")

xlmap.sheet.add_chart(chart, "A1")
xlmap.writer.save()

#### Using xlsxwriter ##################################################################################################

### Bar Charts #########################################################################################################

writer = pd.ExcelWriter("XlsxWriterExample.xlsx", engine='xlsxwriter')

f1 = XLDataFrame(columns=('X', 'Y1', 'Y2'),
                 data={'X': range(10),
                       'Y1': list(random.randrange(0, 10) for _ in range(10)),
                       'Y2': list(random.randrange(0, 10) for _ in range(10))})
f1.set_index('X', inplace=True)

xlmap2 = calories_per_meal.to_excel(writer, sheet_name="XLLinked") # returns the 'ProxyFrame'
xl_linked_chart = xlmap2.create_chart('column', title="Bar Chart", x_axis_name="Meal", y_axis_name="Calories")

right_of_table = xlmap2.columns[-1].translate(0, 1).cell
xlmap2.sheet.insert_chart(right_of_table, xl_linked_chart)

### Simple Scatter #####################################################################################################

xlmap1 = f1.to_excel(writer, sheet_name='scatter')
scatter_chart = xlmap1.create_chart('scatter', x_axis_name='x', y_axis_name='y', title='Scatter Example')
xlmap1.sheet.insert_chart(xlmap1.columns[-1].translate(0, 1).cell, scatter_chart)

writer.save()