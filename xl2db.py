#!/usr/bin/env python
"""
Excel-to-Sqlite database data transfer tool
By Gustav Melck, 27 Oct 2015
"""

import argparse
import pandas as pd
#import q
import sqlite3
import yaml


def parse_args():
    """Parse arguments and return the arg object."""
    parser = argparse.ArgumentParser(
        description='Move data from an Excel file to a SQLite database',
        epilog='Created by Gustav Melck')
    parser.add_argument(
        'excel_file',
        help='the path to the Excel file containing your data',
        type=argparse.FileType('r'))
    parser.add_argument(
        'database_file',
        help='the path to the SQLite database',
        type=str)
    parser.add_argument(
        'mapping_file',
        help='Excel fields to database fields mapping',
        type=argparse.FileType('r'))
    return parser.parse_args()

def import_sheet(table_name, excel_data, db_conn, field_mappings):
    """Import data from a specific sheet into the database."""
    fields = []
    columns = []
    column_indices = []
    excel_columns = excel_data.columns.tolist()
    for mapping in field_mappings:
        if mapping[1] in excel_columns:
            fields.append(mapping[0])
            columns.append(mapping[1])
            column_indices.append(excel_columns.index(mapping[1]))
        else:
            print('Could not find ' + mapping[1] + ' in the Excel data')
    if len(fields) > 0:
        query = 'insert into ' + table_name + ' (' + str.join(', ', fields) + \
            ') values '
        query += '(' + str.join(', ', ['?' for i in range(len(fields))]) + ')'
        curs = db_conn.cursor()
        curs.executemany(
            query,
            excel_data.iloc[:, column_indices].values.tolist())
        db_conn.commit()
        curs.close()
    else:
        print('No fields identified for insertion.')

def run_import(excel_file, db_conn, mapping):
    """Run the import according to the mapping specification."""
    for (table_name, sheet_name) in mapping['relationships']:
        excel_data = pd.read_excel(excel_file, sheetname=sheet_name)
        import_sheet(
            table_name,
            excel_data,
            db_conn,
            mapping[sheet_name])

def db_connection(db_path):
    """Connect to the DB and turn on foreign key checking."""
    conn = sqlite3.connect(db_path)
    curs = conn.cursor()
    curs.execute('pragma foreign_keys = on')
    conn.commit()
    curs.close()
    return conn

def main():
    """Main commandline routine."""
    args = parse_args()
    print('Loading Excel file...')
    excel_file = pd.ExcelFile(args.excel_file.name)
    print('Done')
    print('Connecting to database...')
    db_conn = db_connection(args.database_file)
    print('Done')
    print('Loading mapping...')
    mapping = yaml.load(args.mapping_file)
    print('Done')
    run_import(excel_file, db_conn, mapping)
    db_conn.close()
    excel_file.close()

if __name__ == '__main__':
    main()
