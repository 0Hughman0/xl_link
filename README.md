# XLLink for Pandas

Love the fancy pandas indexing and slicing, but frustrated when writing to Excel, and loosing all that functionality?

Fear not! XLLink solves this by returning an XLMap object upon use of to_excel

This map supports all your favourite indexing methods, i.e. loc, iloc, at and iat (*ahem and ix... booooo!), but instead of returning a DataFrame, Series, or scalar, XLMap will instead return the XLRange, or XLCell corresponding to the location of the result within your spreadsheet.

Here's a teaser of what xl_link can do when combined with xlsx writer (for example):

    map = xl_link.write_frame(f, "t.xlsx")

    for time in map.f.index:
        xl_linked_chart.add_series({
                            'name': time,
                            'categories': map.columns.frange,
                            'values': map.loc[time].frange})

Compared to:

    for col_num in range(1, len(f.index) + 1):
        without_chart.add_series({
            'name':       ["Without", col_num, 0],
            'categories': ["Without", 0, 1, 0, 4],
            'values':     ["Without", col_num, 1, col_num, 4]})

Hopefully you agree that the former is far more appealing.

This mysterious map is an XLMap object, that represents the DataFrame f, frozen as it was written to excel, but crucially, it knows the location of every cell and index of f within the spreadsheet.

Let's look at XLMap with a more detailed example:

    >>> f = pd.DataFrame(columns=("Mon", "Tues", "Weds", "Thur"),
                         index=('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                         data={'Mon': (15, 20, 12, 3),
                               'Tues': (5, 16, 3, 0),
                               'Weds': (3, 22, 2, 8),
                               'Thur': (6, 7, 1, 9)})
    >>> f
                            Mon                  Tues      Weds       Thur
        Breakfast         Toast                 Bagel    Cereal  Croissant
        Lunch              Soup  Something Different!      Rice     Hotpot
        Dinner            Curry                  Stew     Pasta    Gnocchi
        Midnight Snack  Shmores               Cookies  Biscuits  Chocolate

    >>> map = xl_link.write_frame(f, "t.xlsx")
    >>> map
        <XLMap: index: <XLRange: 'Sheet1'!A2:A5>, columns: <XLRange: 'Sheet1'!B1:F1>, data: <XLRange: 'Sheet1'!B2:F5>>
    >>> map.index
        <XLRange: 'Sheet1'!A2:A5>
    >>> map.columns
        <XLRange: 'Sheet1'!B1:E1>

Or as an alternative to using xl_link.write_frame, you can use xl_link XLDataFrame instead, which redefines to_excel, to return an XLMap:

    >>> from xl_link import XLDataFrame
    >>> f = XLDataFrame(columns=("Mon", "Tues", "Weds", "Thur"),
                             index=('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
                             data={'Mon': (15, 20, 12, 3),
                                   'Tues': (5, 16, 3, 0),
                                   'Weds': (3, 22, 2, 8),
                                   'Thur': (6, 7, 1, 9)})
    >>> map = f.to_excel("t.xlsx")
    >>> map
        <XLMap: index: <XLRange: 'Sheet1'!A2:A5>, columns: <XLRange: 'Sheet1'!B1:E1>, data: <XLRange: 'Sheet1'!B2:E5>>


If you were to open t.xlsx you would find that the ranges described by map line up perfectly with where f was written. And, write_frame is smart, you can use all of the parameters you normally use with DataFrame.to_excel, just pass them as a dict to write_frame:

    >>> map = xl_link.write_frame(f, "t.xlsx", {'sheet_name': 'Demo Sheet', 'startrow': 7})
    >>> map
        <XLMap: index: <XLRange: 'Demo Sheet'!A9:A12>, columns: <XLRange: 'Demo Sheet'!B8:E8>, data: <XLRange: 'Demo Sheet'!B9:E12>>

Here are some more indexing examples:

    >>> # loc
    >>> map.loc['Lunch', 'Thur']
        <XLCell: 'Demo Sheet'!E10>
    >>> map.loc['Dinner', :]
        <XLRange: 'Demo Sheet'!B11:E11>
    >>> # iloc
    >>> map.iloc[3, 2]
        <XLCell: 'Demo Sheet'!D12>
    >>> map.iloc[:, 1]
        <XLRange: 'Demo Sheet'!C9:C12>
    >>> # at
    >>> map.at['Midnight Snack', 'Tues']
        <XLCell: 'Demo Sheet'!C12>
    >>> # iat
    >>> map.iat[0, 2]
        <XLCell: 'Demo Sheet'!D9>
    >>> # __getitem__
    >>> map['Mon']
        <XLCell: 'Demo Sheet'!B8>
    >>> map[['Mon', 'Tues', 'Weds']]
        <XLRange: 'Demo Sheet'!B8:E8>

For convenience, you can access a copy of the frame f, in it's state as it was written to excel:

    >>> f.loc['Lunch'] = "Nom Nom Nom"
    >>> f
                                Mon         Tues         Weds         Thur
        Breakfast             Toast        Bagel       Cereal    Croissant
        Lunch           Nom Nom Nom  Nom Nom Nom  Nom Nom Nom  Nom Nom Nom
        Dinner                Curry         Stew        Pasta      Gnocchi
        Midnight Snack      Shmores      Cookies     Biscuits    Chocolate

    >>> map.f # Preserved :)
                            Mon                  Tues      Weds       Thur
        Breakfast         Toast                 Bagel    Cereal  Croissant
        Lunch              Soup  Something Different!      Rice     Hotpot
        Dinner            Curry                  Stew     Pasta    Gnocchi
        Midnight Snack  Shmores               Cookies  Biscuits  Chocolate


Note that as a limitation, the map.index and map.columns are simply XLRange objects, so you cannot apply Pandas Index methods to get ranges within them.

That said, XLRanges support integer, slice and boolean indexing (For more details see below/ doc-strings), so there are workarounds:

>>> map.index["Lunch":"Dinner"]
    TypeError: Expecting tuple of slices, boolean indexer, or an index or a slice if 1D, not Lunch
>>> map.index[map.f.index.get_loc('Lunch'):map.f.index.get_loc('Dinner')] # probably more elegant workarounds possible!
    <XLRange: 'Demo Sheet'!A10:A11>


The XLCell and XLRange objects used to store the location of parts of f are powerful in themselves, if needs be, you can create them yourself:

    >>> from xl_link.xl_types import XLRange, XLCell
    >>> start = XLCell(1, 1) # using row, col
    >>> stop = XLCell(1, 8)
    >>> between = start - stop
        <XLRange: 'Sheet1'!B2:I2>

and you can get their location in excel notation via XLCell.cell and XLRange.range respectively:

    >>> start
        <XLCell: 'Sheet1'!B2>
    >>> start.cell
        'B2'
    >>> stop
        <XLCell: 'Sheet1'!I2>
    >>> between
        <XLRange: 'Sheet1'!B2:I2>
    >>> between.range
        'B2:I2'

For convenience add the f prefix for a formula compatible version:

    >>> start.fcell
        "'Sheet1'!B2"
    >>> between.frange
        "'Sheet1'!B2:I2"

And if you prefer to use this notation to initalise you XLRanges and XLCells, that's find too, using from_cell, from_fcell, from_range and from_frange:

    >>> XLCell.from_cell("A6")
        <XLCell: 'Sheet1'!A6>
    >>> XLRange.from_frange("'Another Sheet'!D2:R2")
        <XLRange: ''Another Sheet''!D2:R2>


Translate them, get items using a range of indexers, and even iterate over 1D XLRanges:

    >>> new_start = start.translate(0, 2)
    >>> new_stop = stop.translate(0, 2)
    >>> new_between = new_start - new_stop
    >>> new_between
        <XLRange: 'Sheet1'!D2:K2>
    >>> new_between[3:]
        <XLRange: 'Sheet1'!G2:K2>
    >>> for cell in new_between:
            print(cell.cell)
        D2
        E2
        F2
        G2
        H2
        I2
        J2
        K2
