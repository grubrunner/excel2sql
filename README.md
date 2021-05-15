# Excel 2 SQL

Converts Excels (`.xls` or `.xlsx`) into SQL create an insert statements and excutes those SQL statements into the specified database. If are SQL error occur this script will abort.

This code supports Excel cells that directly reference other cells. For example if a cell's content's are "=SheetName!C2", then that would be replaced by the contents of cell "C2" in sheet "SheetName"

This code does not currently support any other type of Excel formula.

## Prerequisites

### Step 1

Make sure you have python3 installed at `/usr/local/bin/python3`

Install the following python modules:

* pymysql
* openpyxl

### Step 2

Create an `.xls` or `.xlsx` document with one or more sheets. Each sheet should exactly match the name that you want the corresponding database table to have. If the database already contains a table of that name, the existing table with that name will be used.

Note: The first row in each sheet should match the exact name of the columns in the corresponding database table. The order of the columns in the sheet should match the order of the database table columns

## Example Usage

```bash
./scripts/excel2sql.py --debug false --user root --password password --host 127.0.0.1 --database dbname ./Sheet.xlsx
```

## Credits

This code was based off of: https://gist.github.com/antiproblemist/0c2694cc17d7e39e9d12