# -*- coding: utf-8 -*-
import os
import sys
import re
import sqlite3 as sqlite

try:
    if sys.version_info < (2, 3):
        raise ImportError
    import unittest2
    unittest = unittest2
except ImportError:
    import unittest
    unittest2 = None

import xls2db


def do_one(xls_filename, dbname):
    xls2db.xls2db(xls_filename, dbname)


class AllTests(unittest.TestCase):
    def test_stackhaus(self):
        xls_filename, dbname = 'stackhaus.xls', 'stackhaus.db'
        try:
            os.remove(dbname)
        except:
            pass

        do_one(xls_filename, dbname)

        stackhaus = sqlite.connect(dbname)

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

        for t in tests.items():
            table = t[0]
            headers = t[1]

            row = stackhaus.execute(
                "SELECT sql FROM sqlite_master WHERE tbl_name = ? AND type = 'table'", (table,)
            ).fetchone()

            for header in headers:
                msg = u'header ' + header + u' in ' + table
                self.assertTrue(re.search(header, row[0]), 'x ' + msg)


def main():
    # Some tests may use data files (without a full pathname)
    # set current working directory to test directory if
    # test is not being run from the same directory
    testpath = os.path.dirname(__file__)
    if testpath:
        testpath = os.path.join(testpath, 'example')
    else:
        # Just assume current directory
        testpath = 'example'
    try:
        os.chdir(testpath)
    except OSError:
        # this may be Jython 2.2 under OpenJDK...
        if sys.version_info <= (2, 3):
            pass
        else:
            raise
    unittest.main()

if __name__ == '__main__':
    main()
