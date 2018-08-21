==================
XLMap
==================


An XLMap object represents a DataFrame, frozen as it was written to excel, but crucially, it knows the location of every cell and index of f within the spreadsheet.

Let's look at XLMap with a more detailed example:

    >>> f = XLDataFrame(index=('Breakfast', 'Lunch', 'Dinner', 'Midnight Snack'),
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

    >>> xlmap = f.to_excel("t.xlsx")
    >>> xlmap
        <XLMap: index: <XLRange: 'Sheet1'!A2:A5>, columns: <XLRange: 'Sheet1'!B1:F1>, data: <XLRange: 'Sheet1'!B2:F5>>
    >>> xlmap.index
        <XLRange: 'Sheet1'!A2:A5>
    >>> xlmap.columns
        <XLRange: 'Sheet1'!B1:E1>

Here are some more indexing examples:

    >>> # loc
    >>> xlmap.loc['Lunch', 'Thur']
        <XLCell: 'Demo Sheet'!E10>
    >>> xlmap.loc['Dinner', :]
        <XLRange: 'Demo Sheet'!B11:E11>
    >>> # iloc
    >>> xlmap.iloc[3, 2]
        <XLCell: 'Demo Sheet'!D12>
    >>> xlmap.iloc[:, 1]
        <XLRange: 'Demo Sheet'!C9:C12>
    >>> # at
    >>> xlmap.at['Midnight Snack', 'Tues']
        <XLCell: 'Demo Sheet'!C12>
    >>> # iat
    >>> xlmap.iat[0, 2]
        <XLCell: 'Demo Sheet'!D9>
    >>> # __getitem__
    >>> xlmap['Mon']
        <XLCell: 'Demo Sheet'!B8>
    >>> xlmap[['Mon', 'Tues', 'Weds']]
        <XLRange: 'Demo Sheet'!B2:D5>

For convenience, you can access a copy of the frame f, in it's state as it was written to excel:

    >>> f.loc['Lunch'] = "Nom Nom Nom"
    >>> f
                                Mon         Tues         Weds         Thur
        Breakfast             Toast        Bagel       Cereal    Croissant
        Lunch           Nom Nom Nom  Nom Nom Nom  Nom Nom Nom  Nom Nom Nom
        Dinner                Curry         Stew        Pasta      Gnocchi
        Midnight Snack      Shmores      Cookies     Biscuits    Chocolate

    >>> xlmap.f # Preserved :)
                            Mon                  Tues      Weds       Thur
        Breakfast         Toast                 Bagel    Cereal  Croissant
        Lunch              Soup  Something Different!      Rice     Hotpot
        Dinner            Curry                  Stew     Pasta    Gnocchi
        Midnight Snack  Shmores               Cookies  Biscuits  Chocolate

