# Excel to csv conversion tools
## Conversion from Excel (.xlsx) to (google) csv calendar

### Introduction
Tool enables conversion from Excel XLSX format to google calendar csv in PHP.

Tool expects certain format file named as excel.xlsx as an input. From that spreadsheet it takes
certain rows and generates csv format suitable for google calendar

Run:

php xlsx2csv.php

Output will be written to the file named as events.csv.

N.B PHPSpreadsheet is used.

### TODO
* Possibility to get input file name as an argument
* Automatic import to certain google calendar?
