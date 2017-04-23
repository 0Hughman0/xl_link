# XLLink for Pandas

Love the fancy pandas indexing and slicing, but frustrated when writing to Excel, and loosing all that functionality?

Fear not! XLLink solves this by returning a 'proxy' frame object upon use of to_excel

This frame can be sliced and indexed just like your original frame, but at any point you can call:

    obj.xl -> XLRange or XLCell

This will return an XLRange or XLCell object that represents the range or cell that object takes up on your spreadsheet.

These can then be turning into excel notation with either:

    obj.xl.range -> <XLRange A1:C5> (for ranges)
    obj.xl.cell  -> <XLCell C5> (for cells)

For convenience add the f prefix for a formula compatible version

    obj.xl.frange -> "Sheet1"!A1:C5
    obj.xl.cell   -> C5

Example:

    In[3]: print(f)
    Out[3]:
                        Mon                  Tues      Weds       Thur
    Meal
    Breakfast         Toast                 Bagel    Cereal  Croissant
    Lunch              Soup  Something Different!      Rice     Hotpot
    Dinner            Curry                  Stew     Pasta    Gnocchi
    Midnight Snack  Shmores               Cookies  Biscuits  Chocolate

    In[7]: i = f.to_excel("t.xlsx")
    In[8]: i
    Out[8]:
                             Mon          Tues          Weds          Thur
    Meal
    Breakfast       <XLCell: B2>  <XLCell: C2>  <XLCell: D2>  <XLCell: E2>
    Lunch           <XLCell: B3>  <XLCell: C3>  <XLCell: D3>  <XLCell: E3>
    Dinner          <XLCell: B4>  <XLCell: C4>  <XLCell: D4>  <XLCell: E4>
    Midnight Snack  <XLCell: B5>  <XLCell: C5>  <XLCell: D5>  <XLCell: E5>

    In[10]: i.index.xl
    Out[10]: <XLRange: A2:A5>

    In[11]: i.columns.xl
    Out[11]: <XLRange: B1:E1>

    In[9]: i.loc["Lunch", :].xl
    Out[9]: <XLRange: B3:E3>


This really starts to look nice when using something like xlsxwriter.

For example, to create without XLWriter can look like:

    for col_num in range(1, len(calories_per_meal.index) + 1):
        without_chart.add_series({
            'name':       ["Without", col_num, 0],
            'categories': ["Without", 0, 1, 0, 4],
            'values':     ["Without", col_num, 1, col_num, 4]})

which to me looks pretty ugly and confusing. Instead with XLLink this becomes:


    for time in calories_per_meal.index:
        xl_linked_chart.add_series({
                            'name': time,
                            'categories': proxy.columns.xl.frange,
                            'values': proxy.loc[time].xl.frange})

NOTE: This is still in early stages and not fully tested. Please report any bugs for squishing, or help out! Check if your usage case is covered in test.py!
