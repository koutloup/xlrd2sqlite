#!/usr/bin/env python

import argparse
import json
import collections
import sqlite3
import xlrd
import xlrd.xldate

__author__ = 'Petros Koutloubasis'

parser = argparse.ArgumentParser(description='maps a xlrd file to sqlite')
parser.add_argument('excel_file', help='the excel file')
parser.add_argument('sqlite_file', help='the sqlite db')
parser.add_argument('json_map', help='the json mapping')
args = parser.parse_args()


def col_name_to_col_index(col_name):
    result = 0
    for i in xrange(len(col_name)):
        result = result * 26 + ord(col_name[i].lower()) - (ord("a") - 1)
    return result - 1


def extract_col_values_from_row(row_index, col_names):
    cell_values = []
    for col_name in col_names:
        if col_name[0] == '_':
            cell_values.append(col_name[1:])
            continue

        col_index = col_name_to_col_index(col_name)
        cell = sheet.cell(row_index, col_index)
        cell_value = cell.value

        # check for numbers (or dates) because there are no ints (or dates) in excel only floats
        if cell.ctype == xlrd.XL_CELL_DATE:
            xl_date = xlrd.xldate.xldate_as_datetime(cell_value, book.datemode)
            cell_value = xl_date.strftime('%m/%Y')
        elif cell.ctype == xlrd.XL_CELL_NUMBER and int(cell_value) == cell_value:
            cell_value = int(cell_value)
        elif cell.ctype == xlrd.XL_CELL_TEXT:
            cell_value = cell_value.strip()
        cell_values.append(cell_value)

    return cell_values


with open(args.json_map) as data_file:
    json_data = json.load(data_file, object_pairs_hook=collections.OrderedDict)


db_cols = json_data["db_cols"]
sheet_cols = json_data["sheet_cols"]


# open xlrd file
print 'opening', args.excel_file, '..'
book = xlrd.open_workbook(args.excel_file)
print '..done'

# open sheet
sheet_index = json_data.get('sheet_index', 0)
print 'opening sheet', sheet_index, '..'
sheet = book.sheet_by_index(sheet_index)
print '..done'

print 'reading data from sheet..'
rows_to_skip = json_data.get('rows_to_skip', 0)
insert_values = []
for row in xrange(rows_to_skip, sheet.nrows):
    if type(sheet_cols[0]) is list:
        for sheet_cols_array in sheet_cols:
            col_values = extract_col_values_from_row(row, sheet_cols_array)
            insert_values.append(col_values)
    else:
        col_values = extract_col_values_from_row(row, sheet_cols)
        insert_values.append(col_values)
print '..done'

print 'writing data to db..'
# connect to db
db_conn = sqlite3.connect(args.sqlite_file)
db_cursor = db_conn.cursor()

table_name = json_data["table_name"]
db_cursor.execute("DELETE FROM " + table_name)

insert_query = 'INSERT INTO %s (%s) VALUES(%s)' % \
               (table_name, ','.join(db_cols), ','.join(['?'] * len(db_cols)))

not_empty_fields = json_data.get('not_empty', [])
for vals in insert_values:
    empty = len(not_empty_fields) > 0
    for field in not_empty_fields:
        i = db_cols.index(field)
        if vals[i]:
            empty = False
            break

    if not empty:
        print vals
        db_cursor.execute(insert_query, vals)

db_conn.commit()
db_conn.close()
print '..done'
