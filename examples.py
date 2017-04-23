import pandas as pd
from xl_link import EmbededFrame

# Create EmbededFrame
calories_per_meal = EmbededFrame(pd.DataFrame(columns=("Meal", "Mon", "Tues", "Weds", "Thur"),
                                 data={'Meal': ('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                                       'Mon': (15, 20, 12, 3),
                                       'Tues': (5, 16, 3, 0),
                                       'Weds': (3, 22, 2, 8),
                                       'Thur': (6, 7, 1, 9)}))
calories_per_meal.set_index("Meal", drop=True, inplace=True)

# Write to excel
writer = pd.ExcelWriter("Example.xlsx", engine='xlsxwriter')
proxy = calories_per_meal.to_excel(writer, sheet_name="XLLinked") # returns the 'ProxyFrame'

# Create chart with XLLink ############################################################################################

workbook = writer.book
xl_linked_sheet = writer.sheets["XLLinked"]
xl_linked_chart = workbook.add_chart({'type': 'column'})

for time in calories_per_meal.index:
    xl_linked_chart.add_series({'name': time,
                      'categories': proxy.columns.xl.frange,
                      'values': proxy.loc[time].xl.frange})

right_of_table = proxy.columns.xl[-1].translate(0, 1)
xl_linked_sheet.insert_chart(right_of_table.cell, xl_linked_chart)

"""
Easy to read, and modify, intuitive
"""
######################################################################################################################

# Create chart without XLLink :( #####################################################################################

calories_per_meal.to_excel(writer, sheet_name="Without")

without_sheet = writer.sheets["Without"]
without_chart = workbook.add_chart({"type": "column"})

for col_num in range(1, len(calories_per_meal.index) + 1):
    without_chart.add_series({
        'name':       ["Without", col_num, 0],
        'categories': ["Without", 0, 1, 0, 4],
        'values':     ["Without", col_num, 1, col_num, 4]})

without_sheet.insert_chart(right_of_table.cell, without_chart)

"""
Overly complex, confusing, hard to change
"""

######################################################################################################################

writer.save()
