# xl2db
Copy data from a Microsoft Excel file into a SQLite database

## Description

This is a command-line utility that is designed to read data from an Excel file
and to copy that data into a SQLite database. The Excel file can have multiple
sheets but it is a requirement that the data in the specific sheet to be
imported be in a tabular format with column headings starting from cell A1 and
data starting from A2.

The destination table in the database should be created manually before the
import. Destination tables should include conflict clauses and constraints
which can be used to prevent duplicate imports—a handy feature when data in the
Excel file grows over time and needs to be re-imported.

The mapping specifying which columns should go to which fields in the database
is described below.

## Dependencies

* Python 3 (built using Python 3.4)
* PyYaml (for specifying the mapping)
* Pandas (for connecting to Excel)

## Data import mapping

```yaml
# DB table name => Excel sheet name
relationships:
- [db_table_name, excel_sheet_name]
- [db_table_name, excel_sheet_name]
# ...

# DB table field name => Excel table column name
excel_sheet_name:
- [db_field_name, column_name]
- [db_field_name, column_name]
# ...
```

## Usage

* *To be completed*. In the meanwhile type `./xl2db.py -h` from the commandline.

## Todo list

* *Complete the README!*
