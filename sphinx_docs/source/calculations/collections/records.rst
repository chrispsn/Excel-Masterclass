Records
=======

Records are a kind of collection. Records are rows of data where each column of the data is a specific kind of information (a 'field'). Here's an example:

+--------+-------+------------+
|**Name**|**Age**|**Location**|
+--------+-------+------------+
|Alice   |25     |Canberra    |
+--------+-------+------------+
|Bob     |24     |Melbourne   |
+--------+-------+------------+
|Chris   |26     |Sydney      |
+--------+-------+------------+

Most of the time processing records is easy, because:

* there are built-in functions for most common data transformations
* most of these functions take columns of data as input, and
* Excel Tables' structured references makes referencing columns easy.

Here's an example::

	=AVERAGE(data[Age])

Excel has two families of function that work like this:

* Standard aggregation functions such as ``SUMIF``.
* Database functions such as ``DSUM``.

.. XXX elaborate on trade-offs between the two

.. XXX TEST ALL THIS CODE IN THE WORKBOOK


Filtering
---------

Filtering to get a specific record or field is not as easy using formulas.

``VLOOKUP`` is the classic 'get' function, but it only works if all of the following apply:

* you have one condition
* that condition is that a value be equal to another value
* the condition column is the first column in the data
* you want one value, and
* you want that value from the first record that meets the condition.

We can relax these constraints. Using the example records at the top, we'll:

* Get record row and field number(s) based on some conditions
* Get data out of the table using those numbers and the ``INDEX`` function.

By the way - there are many ways to do this. I find the ``INDEX`` approach is the best balance of functionality, calculation speed and brevity, but for completeness I've also briefly discussed alternatives at the bottom of the page.


Record Rows
~~~~~~~~~~~

If you want just one record but you have multiple conditions::

	<ARRAY> =MATCH(1, conditions, 0)

For example::

	<ARRAY> =MATCH(1, 1*(data[Name]="Bob"), 0)

If you want more than one record::

	<ARRAY> =SMALL(IF(conditions, ROW(data)-ROW(data[#Headers]), ""), n)

For example, getting the second record matching the criteria::

	<ARRAY> =SMALL(
		IF(
			(data[Age]>=25)*(0<((data[Location]="Sydney")+(data[Location]="Canberra"))), 
			ROW(data)-ROW(data[#Headers]), 
			""
		),
		2
	)

If you want all records matching the criteria:

* count the number of records matching the criteria (say ``X``)
* generate an array from 1 to ``X`` using ``ROW`` and ``OFFSET``.

For example::

	<ARRAY> =SMALL(
	 	IF(
	  		(data[Age]>=25)*(0<((data[Location]="Sydney")+(data[Location]="Canberra"))),
			ROW(data)-ROW(data[#Headers]),
	   		""
	  	),
	  	ROW(OFFSET($A$1, 0, 0, SUM(
	  		1*(
	  			0<(
	  				(data[Location]="Sydney")
	  				+(data[Location]="Canberra")
	  			)
	  		)
	  	)))
	 )

.. XXX is there a way to do it without OFFSET or other volatile functions?
.. XXX is there a way to do it that doesn't duplicate the conditional logic?

Field Columns
~~~~~~~~~~~~~

``MATCH`` is the obvious choice.

To get a specific field column::

	=MATCH(field_name, data[#Headers], 0)

To get multiple field columns::

	=MATCH({field_name1, field_name2 (...)}, data[#Headers], 0)

For example::

	=MATCH({"Age", "Location"}, data[#Headers], 0)

To get all field column numbers (only needed if not getting a specific record)::

	<ARRAY> =COLUMN(data[#Headers])-COLUMN(INDEX(data, 0, 1)) + 1


Combining with ``INDEX``
~~~~~~~~~~~~~~~~~~~~~~~~

Your final formula will depend on how many fields you need:

* an individual field: ``=INDEX(data, record_numbers, field_number)``
* multiple fields: ``=INDEX(data, record_numbers, field_numbers)``
* all fields for a specific single row: ``=INDEX(data, record_number, 0)`` [#index_column_argument_weird1]_
* all fields for an array of rows, one or more: ``=INDEX(data, record_numbers, COLUMN(data[#Headers])-COLUMN(INDEX(data, 0, 1)) + 1)`` [#index_column_argument_weird2]_

Examples::

	<ARRAY> =INDEX(data, {1; 2}, 2)       				[field 2]
	<ARRAY> =INDEX(data, {1; 2}, {2, 3})  				[fields 2 and 3]
	<ARRAY> =INDEX(data, 2, 0)           				[all fields for record 2]
	<ARRAY> =INDEX(data, {1; 2}, 
				COLUMN(data[#Headers])
				-COLUMN(INDEX(data, 0, 1)) + 1)         [all fields for records 1 and 2]


.. XXX apparently there's a VBA solution in Excel Gurus Gone Wild too


.. [#index_column_argument_weird1] ``INDEX(data, n)`` (leaving out the column argument) will fail if ``data`` is a range, but works just fine if it's an array.
.. [#index_column_argument_weird2] All this messing around with ``COLUMN`` just because we're using an array formula. Excel, man.

Other Methods
~~~~~~~~~~~~~

Other methods I know of can only get the first record.

One way is to transform the data before giving it to ``VLOOKUP``, by using ``CHOOSE`` to move columns around::

	<ARRAY> =VLOOKUP(value, CHOOSE({1, 2}, column1, column2), 2, FALSE)

For example::

	<ARRAY> =VLOOKUP(25, CHOOSE({1, 2}, data[Age], data[Name]), 2, FALSE)

Another uses ``DGET``.

.. XXX DGET still good because of speed?

