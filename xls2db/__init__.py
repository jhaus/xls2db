import xlrd
import sqlite3 as sqlite
import itertools

def xls2db(infile, outfile):
    """
    Convert an xls file into an sqlite db!
    """
    #Now you can pass in a workbook!
    if type(infile) == str:
        wb = xlrd.open_workbook(infile)
    elif type(infile) == xlrd.Book:
        wb = infile
    else:
        raise TypeError

    #Now you can pass in a sqlite connection!
    if type(outfile) == str:
        db_conn = sqlite.connect(outfile)
        db_cursor = db_conn.cursor()
    elif type(outfile) == sqlite.Connection:
        db_conn = outfile
        db_cursor = db_conn.cursor()
    else:
        raise TypeError

    for s in wb.sheets():

        # Create the table.
        # Vulnerable to sql injection because ? is only able to handle inserts
        # I'm not sure what to do about that!
        if s.nrows > 0:
            db_cursor.execute("create table " + s.name + " ("
                + ','.join([s.cell(0,j).value for j in xrange(s.ncols)]) +");")

        if s.nrows > 1:
            tmp_sql = "insert into " + s.name + ' values (' +','.join(itertools.repeat('?', s.ncols)) +");"
            for rownum in xrange(1, s.nrows):
                db_cursor.execute(tmp_sql, s.row_values(rownum))

    db_conn.commit()
    #Only do this if we're not working on an externally-opened db
    if type(outfile) == str:
        db_cursor.close()
        db_conn.close()

def db2xls(infile, outfile):
    """
    Convert an sqlite db into an xls file. Not implemented!
    Some issues: one needs to be able to figure out what the table names are!
    """
    raise NotImplementedError
