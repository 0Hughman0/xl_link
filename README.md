# XLLink for Pandas

Love the fancy pandas indexing and slicing, but frustrated when writing to Excel, and loosing all that functionality?

Fear not! XLLink solves this by returning an XLMap object upon use of to_excel

This map supports all your favourite indexing methods, i.e. loc, iloc, at and iat (*ahem and ix... booooo!), but instead of returning a DataFrame, Series, or scalar, XLMap will instead return the XLRange, or XLCell corresponding to the location of the result within your spreadsheet.

XLCell and XLRange objects can return their location in excel notation via XLCell.cell and XLRange.range respectively:

    >>> first_temp_cell = XLCell(1, 1)
    >>> last_temp_cell = XLCell(1, 10)
    >>> temp_col = first_temp_cell - last_temp_cell (or could use XLRange(first_temp_cell, last_temp_cell))
    >>> first_temp_cell
        <XLCell: B2>
    >>> first_temp_cell.cell
        'B2'
    >>> temp_col
        <XLRange: B2:B11>
    >>> temp_col.range
        'B2:B11'

For convenience add the f prefix for a formula compatible version:

    >>> first_temp.fcell
        "'Sheet1'!B11"
    >>> temp_col.frange
        "'Sheet1'!B2:B11"

But rather than creating these XLCells and XLRanges from scratch, just let xl_link do it for you!:

    >>> print(f)
                            Mon                  Tues      Weds       Thur
        Meal
        Breakfast         Toast                 Bagel    Cereal  Croissant
        Lunch              Soup  Something Different!      Rice     Hotpot
        Dinner            Curry                  Stew     Pasta    Gnocchi
        Midnight Snack  Shmores               Cookies  Biscuits  Chocolate

    >>> map = xl_link.write_frame(f, "t.xlsx")
    >>> map
        <XLMap: index: <XLRange: A2:A5>, columns: <XLRange: B1:E1>, data: <XLRange: B2:E5>>

    >>> map.index
        <XLRange: A2:A5>

    >>> map.columns
        <XLRange: B1:E1>

    >>> map.loc["Lunch", :]
        <XLRange: B3:E3>

For convenience, you can access a copy of the frame f, as it was written to excel:

    >>> f.loc['Lunch'] = "Nom Nom Nom"
    >>> f
                                Mon         Tues         Weds         Thur
        Meal
        Breakfast             Toast        Bagel       Cereal    Croissant
        Lunch           Nom Nom Nom  Nom Nom Nom  Nom Nom Nom  Nom Nom Nom
        Dinner                Curry         Stew        Pasta      Gnocchi
        Midnight Snack      Shmores      Cookies     Biscuits    Chocolate

    >>> map.f # Preserved :)
                            Mon                  Tues      Weds       Thur
        Meal
        Breakfast         Toast                 Bagel    Cereal  Croissant
        Lunch              Soup  Something Different!      Rice     Hotpot
        Dinner            Curry                  Stew     Pasta    Gnocchi
        Midnight Snack  Shmores               Cookies  Biscuits  Chocolate


This all comes together to look nice when using something like xlsxwriter.
Allowing for something like:

    for time in map.f.index:
        xl_linked_chart.add_series({
                            'name': time,
                            'categories': proxy.columns.frange,
                            'values': proxy.loc[time].frange})

Compared to:

    for col_num in range(1, len(calories_per_meal.index) + 1):
        without_chart.add_series({
            'name':       ["Without", col_num, 0],
            'categories': ["Without", 0, 1, 0, 4],
            'values':     ["Without", col_num, 1, col_num, 4]})

Hopefully you agree that the former is far more appealing.

