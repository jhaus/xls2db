# -*- coding: utf-8 -*-
import os
import re
import sqlite3 as sqlite

from xls2db import xls2db

try:
    os.remove("example/stackhaus.db")
except:
    pass

print('--- running xls2db ---')

xls2db("example/stackhaus.xls", "example/stackhaus.db")

print('--- connecting to sqlite db ---')

stackhaus = sqlite.connect("example/stackhaus.db")

tests = {
    "locations": [
        "id string primary key",
        "short_descr string",
        "long_descr string",
        "special string"
    ],
    "links": [
        "src string",
        "dst string",
        "dir string"
    ],
    "items": [
        "id string primary key",
        "location string",
        "short_descr string",
        "long_descr string",
        "get_descr string",
        "get_pts integer",
        "use_desc string",
        "use_pts integer"
    ]
}

print('--- running tests ---');

for t in tests.items():
    table = t[0]
    headers = t[1]

    row = stackhaus.execute(
        "SELECT sql FROM sqlite_master WHERE tbl_name = ? AND type = 'table'", (table,)
    ).fetchone()

    for header in headers:
        msg = u'header ' + header + u' in ' + table
        try:
            assert re.search(header, row[0]), 'x ' + msg
        except AssertionError as err:
            print(err)
        else:
            print(u'✓ ' + msg)

print('--- done. ---');
