# XLLink for Pandas

Love the fancy pandas indexing and slicing, but frustrated when writing to Excel, and loosing all that functionality?

Fear not! XLLink solves this by returning a 'proxy' frame object upon use of to_excel

This frame can be sliced and indexed just like your original frame, but at any point you can call:

    obj.xl -> XLRange or XLCell

This will return an XLRange or XLCell object that represents the range or cell that object takes up on your spreadsheet.

These can then be turning into excel notation with either:

    obj.xl.range -> A1:C5 (for ranges)
    obj.xl.cell  -> C5 (for cells)

For convenience add the f prefix for a formula compatible version

    obj.xl.frange -> ="Sheet1"!A1:C5
    obj.xl.cell   -> ="Sheet1"!C5

Example:

    In[2]: f = EmbededFrame({"one": range(10), "two": range(10, 20), "three": range(20, 30)})

    In[3]: f
    Out[4]:
       one  three  two
    0    0     20   10
    1    1     21   11
    2    2     22   12
    3    3     23   13
    4    4     24   14
    5    5     25   15
    6    6     26   16
    7    7     27   17
    8    8     28   18
    9    9     29   19
    In[5]: i = f.to_excel("temp.xlsx", sheet_name="Cool Sheet", startrow=4, startcol=3)
    In[6]: i
    Out[6]:
       one three  two
    0   E6    F6   G6
    1   E7    F7   G7
    2   E8    F8   G8
    3   E9    F9   G9
    4  E10   F10  G10
    5  E11   F11  G11
    6  E12   F12  G12
    7  E13   F13  G13
    8  E14   F14  G14
    9  E15   F15  G15
    In[7]: i.loc[:, "two"]
    Out[7]:
    0     G6
    1     G7
    2     G8
    3     G9
    4    G10
    5    G11
    6    G12
    7    G13
    8    G14
    9    G15
    Name: two, dtype: object
    In[8]: i.loc[:, "two"].index.xl
    Out[8]: D6:D15
    In[9]: i.loc[:, "two"].xl
    Out[9]: G6:G15

    In[12]: i.ix[1:4, "two"].xl.frange
    Out[12]: 'Cool Sheet!G7:G10'

NOTE: This is not fully tested. Please report any bugs for squishing, or help out!
