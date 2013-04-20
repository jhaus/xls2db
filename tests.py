#!/usr/bin/env python
# -*- coding: ascii -*-
# vim:ts=4:sw=4:softtabstop=4:smarttab:expandtab
"""Test suite for xls2db
"""

import os
import sys
import sqlite3

try:
    if sys.version_info < (2, 3):
        raise ImportError
    import unittest2
    unittest = unittest2
except ImportError:
    import unittest
    unittest2 = None

import xls2db


def do_one(xls_filename, dbname, do_drop=False):
    if do_drop:
        xls2db.xls2db(xls_filename, dbname, do_drop=do_drop)
    else:
        xls2db.xls2db(xls_filename, dbname)


class AllTests(unittest.TestCase):
    def test_comma(self):
        xls_filename, dbname = 'comma_test.xls', ':memory:'
        # For now just see if it works without error
        do_one(xls_filename, dbname)

    def test_empty_worksheet(self):
        xls_filename, dbname = 'empty_worksheet_test.xls', ':memory:'
        db = sqlite3.connect(dbname)
        try:
            c = db.cursor()
            do_one(xls_filename, db)
            c.execute("SELECT tbl_name FROM sqlite_master WHERE type='table'")
            rows = c.fetchall()
            self.assertEqual(rows, [(u'simple_test',)])
        finally:
            db.close()

    def test_simple_test(self):
        xls_filename, dbname = 'simple_test.xls', ':memory:'
        db = sqlite3.connect(dbname)
        try:
            c = db.cursor()
            do_one(xls_filename, db)
            c.execute("SELECT tbl_name FROM sqlite_master WHERE type='table'")
            rows = c.fetchall()
            self.assertEqual(rows, [(u'simple_test',)])
            
            c.execute("SELECT * FROM simple_test ORDER BY 1")
            rows = c.fetchall()
            self.assertEqual(rows, [(1.0, u'one'), (2.0, u'two'), (3.0, u'three')])
        finally:
            db.close()

    def test_simple_test_twice(self):
        xls_filename, dbname = 'simple_test.xls', ':memory:'
        db = sqlite3.connect(dbname)
        c = db.cursor()
        
        def do_one_simple_conversion():
            do_one(xls_filename, db)
            c.execute("SELECT tbl_name FROM sqlite_master WHERE type='table'")
            rows = c.fetchall()
            self.assertEqual(rows, [(u'simple_test',)])
            
            c.execute("SELECT * FROM simple_test ORDER BY 1")
            rows = c.fetchall()
            self.assertEqual(rows, [(1.0, u'one'), (2.0, u'two'), (3.0, u'three')])
        
        try:
            do_one_simple_conversion()
            self.assertRaises(sqlite3.OperationalError, do_one_simple_conversion)
        finally:
            db.close()

    def test_simple_test_twice_with_drop(self):
        xls_filename, dbname = 'simple_test.xls', ':memory:'
        db = sqlite3.connect(dbname)
        c = db.cursor()
        
        def do_one_simple_conversion():
            do_one(xls_filename, db, do_drop=True)
            c.execute("SELECT tbl_name FROM sqlite_master WHERE type='table'")
            rows = c.fetchall()
            self.assertEqual(rows, [(u'simple_test',)])
            
            c.execute("SELECT * FROM simple_test ORDER BY 1")
            rows = c.fetchall()
            self.assertEqual(rows, [(1.0, u'one'), (2.0, u'two'), (3.0, u'three')])
        
        try:
            do_one_simple_conversion()
            do_one_simple_conversion()  # do again
        finally:
            db.close()


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
