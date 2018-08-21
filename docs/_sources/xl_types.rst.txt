==================
XLRange and XLCell
==================

These are the objects used within xl_link to represent ranges and cells within excel.

These objects have a ton of methods, making them powerful in themselves.

The primary way of getting ahold of these objects is from an ``XLMap``::


	>>> from xl_link import XLDataFrame # Rather than pandas DataFrame
	>>> f = XLDataFrame(data={'x': list(range(10)),
	                          'y': list(range(10, 20))})
	>>> xlmap = f.to_excel("book.xlsx")
	
	Get some XLRanges and XLCells
	
	>>> xlmap.index
	    <XLRange: 'Sheet1'!A2:A11>
	>>> xlmap['y']
	    <XLRange: 'Sheet1'!C2:C11>
	>>> xlmap.loc[3, 'x']
	    <XLCell: 'Sheet1'!B5>

if needs be, you can create them yourself:

	>>> from xl_link.xl_types import XLRange, XLCell
	>>> start = XLCell(1, 1) # using row, col
	>>> start
	    <XLCell: 'Sheet1'!B2>
	>>> stop = XLCell(1, 8)
	>>> between = start - stop
	    <XLRange: 'Sheet1'!B2:I2>
	or using XLCell.range_between
	>>> between = start.range_between(stop)
	    <XLRange: 'Sheet1'!B2:I2>

and you can get their location in excel notation via ``XLCell.cell`` and ``XLRange.range`` respectively:

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

And if you prefer to use this notation to initalise you ``XLRange`` s and ``XLCell`` s, that's find too, using ``from_cell``, ``from_fcell``, ``from_range`` and ``from_frange``:

	>>> XLCell.from_cell("A6")
	    <XLCell: 'Sheet1'!A6>
	>>> XLRange.from_frange("'Another Sheet'!D2:R2")
	    <XLRange: ''Another Sheet''!D2:R2>


Using the translate method finding relative positions is simple::

	>>> new_start = start.translate(0, 2)
	>>> new_start
	    <XLCell: 'Sheet1'!D2>
	>>> new_stop = stop.translate(0, 2)

XLRanges also support a range of indexers::

	>>> new_between = new_start - new_stop
	>>> new_between
	    <XLRange: 'Sheet1'!D2:K2>
	>>> new_between[-3] # integer
	    <XLCell: 'Sheet1'!I2>
	>>> new_between[3:] # slice
	    <XLRange: 'Sheet1'!G2:K2>
	>>> new_between[np.array([0, 1, 1, 1, 0, 0, 0], dtype=bool)] # boolean arrays (for use with Pandas!)
	    <XLRange: 'Sheet1'!E2:G2>

Iterate over 1D ranges::

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

Over 2D XLRanges you can use ``XLRange.iterrows()``:

	>>> square = XLCell(0, 0) - XLCell(3, 3)
	>>> for row_range in square.iterrows():
	        print(row_range)
	        for cell in row_range:
	        print(cell)
	    <XLRange: 'Sheet1'!A1:D1>
	    <XLCell: 'Sheet1'!A1>
	    <XLCell: 'Sheet1'!B1>
	    <XLCell: 'Sheet1'!C1>
	    <XLCell: 'Sheet1'!D1>
	    ...
	    <XLRange: 'Sheet1'!A4:D4>
	    <XLCell: 'Sheet1'!A4>
	    <XLCell: 'Sheet1'!B4>
	    <XLCell: 'Sheet1'!C4>
	    <XLCell: 'Sheet1'!D4>

