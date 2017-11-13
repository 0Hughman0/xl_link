# XLLink for Pandas

**Quick-start: https://0hughman0.github.io/xl_link/quickstart.html**

**Installation: `pip install xl_link`**

XLLink aims to provide a powerfull interface between Pandas DataFrames, and Excel spreadsheets.

XLLink tweaks the DataFrame.to_excel method to return its own XLMap object.

This XLMap stores all the information about the location of the written DataFrame within the spreadsheet.

By allowing you to use Pandas indexing methods, i.e. loc, iloc, at and iat. XLLink remains intuitive to use. But instead of returning a DataFrame, Series, or scalar, XLMap will instead return the XLRange, or XLCell corresponding to the location of the result within your spreadsheet.

Additionally XLMaps offer a wrapper around excel engines (currently supporting xlsxwriter and openpyxl) to make creating charts in excel far more intuitive. Providing a DataFrame.plot like interface for excel charts.

## Chart demo

Here's a teaser of what xl_link can do when combined with xlsx writer (for example):

    >>> from xl_link import XLDataFrame
	>>> import numpy as np
	>>> f = XLDataFrame(data={'Y1': np.random.randn(30),
                              'Y2': np.random.randn(30)})
    >>> xlmap = f.to_excel('Chart Demo.xlsx', sheet_name='scatter')
    >>> scatter_chart = xlmap.create_chart('scatter', x_axis_name='x', y_axis_name='y', title='Scatter Example')
    >>> xlmap.sheet.insert_chart('D1', scatter_chart)
    >>> xlmap.writer.save()

Which produces this chart:

![scatter chart](https://raw.githubusercontent.com/0Hughman0/xl_link/master/examples/ScatterExample.PNG)

Creating a complex chart like this:

![multi bar chart](https://raw.githubusercontent.com/0Hughman0/xl_link/master/examples/BarExample.png)

is as easy as:

    Setup

    >>> f = XLDataFrame(index=('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                                       data={'Mon': (15, 20, 12, 3),
                                             'Tues': (5, 16, 3, 0),
                                             'Weds': (3, 22, 2, 8),
                                             'Thur': (6, 7, 1, 9)})

    Create chart with xl_link

    >>> xlmap = f.to_excel('Compare.xlsx', sheet_name="XLLinked", engine='openpyxl')
    >>> xl_linked_chart = xlmap.create_chart('bar', title="With xl_link", x_axis_name="Meal", y_axis_name="Calories", subtype='col')
    >>> xlmap.sheet.add_chart(xl_linked_chart, 'F1')
    >>> xlmap.writer.save()

Check out the examples folder for more examples!

This package uses the utility functions from XlsxWriter under the BSD license found here: https://github.com/jmcnamara/XlsxWriter
