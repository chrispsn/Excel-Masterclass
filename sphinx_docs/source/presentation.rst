============
Presentation
============

General
-------

* Many features - list validation, conditional formatting, protection - rely on dialog boxes. Consider storing logic relating to these in a spreadsheet cell - not only to make model logic as transparent to other model maintainers, but also because it's hard to edit complex logic in Excel's tiny dialogue boxes. 


Conditional Formatting
----------------------

A spreadsheet user will often need to be alerted if certain conditions are met - for example, if not all the tests pass. Summarising test results can be simply computed by formulas such as `=COUNTIF(tests, FALSE)`, but their output may not be understood by the user.</p>

Suggestions:

* Store complex conditional formatting logic in a helper cell.
* Where you need to change the appearance of a value (particularly booleans) but also use it in a calculation, consider using `custom number formats`__. This is particularly useful for boolean values (just convert them to numbers first, eg using `*1`).

__ http://office.microsoft.com/en-au/excel-help/create-a-custom-number-format-HP010342372.aspx