# Excel 2 SQL

Converts Excels (`.xls` or `.xlsx`) into SQL "create" & "insert into" statements and excutes those SQL statements into the specified database. If any SQL error occurs this script will abort.

This code supports Excel cells that directly reference other cells. For example if a cell's contents are "=SheetName!C2", then that would be replaced by the contents of cell "C2" in sheet "SheetName"

This code does not currently support any other type of Excel formula.

The `excel2sql.py` script will treat the first row as the column names of the table. So all column names should be included. It will treat additional rows as rows in the database's table. However it will ignore any rows that have a background color. This is so that you can include an "example" row in the spreadsheet that shows what a valid row looks like. This works nicely with the `--include_example_row true` flag in the `sql2empty-excel.py` script.

## Prerequisites

You can create the excel script structure automatically or manually.

Make sure you have python3 available in some folder in your $PATH

Install the following python modules:

* pymysql
* openpyxl

You can use this command to do that:

```bash
pip3 install pymysql
pip3 install openpyxl
```

### Automatic Setup

```bash
./sql2empty-excel.py \
  --debug false \
  --user root \
  --password password \
  --host 127.0.0.1 \
  --database dbname \
  --include_example_row true \
  --table_blacklist ignored_table1,ignored_table2 \
  ./EmptySheet.xlsx
```

### Manual Setup

Create an `.xls` or `.xlsx` document with one or more sheets. Each sheet should exactly match the name that you want the corresponding database table to have. If the database already contains a table of that name, the existing table with that name will be used.

Note: The first row in each sheet should match the exact name of the columns in the corresponding database table. You should include all columns from the database into the sheet. The order of the columns in the sheet should match the order of the database table columns

## Usage

```bash
./excel2sql.py \
  --debug false \
  --user root \
  --password password \
  --host 127.0.0.1 \
  --database dbname \
  ./Sheet.xlsx
```

## Why?

* If you want to import a lot of data into a SQL database, this may be for you, since Excel documents are easier to work with than SQL databases
* If your data entry team is more familiar with Excel than SQL, or you do not want to give them write access to your database, then this may be for you

## Credits

This code was based off of https://gist.github.com/antiproblemist/0c2694cc17d7e39e9d12
