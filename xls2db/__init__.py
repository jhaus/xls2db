#!/usr/bin/env python
# coding=utf-8

import sys, os, itertools, logging, datetime
import xlrd


if sys.version_info >= (3,):
    string_types = str,
else:
    string_types = basestring,
    range = xrange


logging.basicConfig() ## NO debug, no info. But logs warnings
log = logging.getLogger(__name__)
log.setLevel(logging.ERROR)
#log.setLevel(logging.DEBUG)


def xls2db(infile, outfile=None, column_name_start_row=0, data_start_row=1):
    """
    Convert an xls file into an Oracle db!
    """
    #Now you can pass in a workbook!
    if isinstance(infile, string_types):
        wb = xlrd.open_workbook(infile)
    elif isinstance(infile, xlrd.Book):
        wb = infile
        infile = "default"
    else:
        raise TypeError("infile must be a string or xlrd.Book")

    #Now you can pass in a Oracle connection!
    if outfile is None:
        raise NotImplementedError("outfile must be a string")

    if isinstance(outfile, string_types):
        outfile = outfile
    else:
        raise TypeError("outfile must be a string or cx_Oracle.Connection")

    of = open(outfile, 'w')

    # hack, avoid plac's annotations....
    column_name_start_row = int(column_name_start_row)
    data_start_row = int(data_start_row)
    for s in wb.sheets():
        # Create the table.
        # Vulnerable to sql injection because ? is only able to handle inserts
        # I'm not sure what to do about that!
        if s.nrows > column_name_start_row:
            column_names = []
            column_types = []
            for j in range(s.ncols):
                # FIXME TODO deal with embedded double quotes
                colname = s.cell(column_name_start_row, j).value
                if colname:
                    colname = '"c%d"' % (colname,) if isinstance(colname, float) else \
                              u'"%s"' % (colname,)
                    colname = colname.replace(" ","_")
                else:
                    colname = '"col%d"' % (j + 1,)
                # FIXME TODO deal with embedded spaces in names
                # (requires delimited identifiers) and missing column types
                coltype = s.cell(data_start_row, j).value
                if coltype:
                    if isinstance(coltype, float):
                        if colname[-5:] == 'Date"':
                            coltype = "TIMESTAMP"
                        elif colname[-5:] == 'Time"':
                            coltype = "TIMESTAMP"
                        else:
                            coltype = 'DOUBLE PRECISION'
                    else:
                        coltype = 'VARCHAR2(1000)'
                else:
                    coltype = 'VARCHAR2(1000)'
                column_names.append(colname)
                column_types.append(coltype)
                #print zip(column_names, column_types)
            column_names = ',\n'.join(map(lambda x: ' '.join(x), zip(column_names, column_types)))
            tmp_sql = u'create table "' + s.name + '" (' + column_names + ');\n'
            of.write(tmp_sql)

        if s.nrows > data_start_row:
#            tmp_sql = 'insert into "' + s.name + '" values (' +','.join(itertools.repeat('?', s.ncols)) +");"
            for rownum in range(data_start_row, s.nrows):
                of.write('insert into "' + s.name + '" values (')
                valcol = []
                for i, rv in enumerate(s.row_values(rownum)):
                    if isinstance(rv, float):
                        if(column_types[i] == 'TIMESTAMP'):
                            valcol.append("TO_TIMESTAMP('"+datetime.datetime(*xlrd.xldate_as_tuple(rv, wb.datemode)).isoformat(' ')+"', 'YYYY-MM-DD HH24:MI:SS')")
                        else:
                            valcol.append(str(rv))
                    else:
                        valcol.append("'"+rv.replace("'","''")+"'") # I know.
                of.write(", ".join(valcol))
                of.write(");\n")


def db2xls(infile, outfile):
    """
    Convert an sqlite db into an xls file. Not implemented!
    Some issues: one needs to be able to figure out what the table names are!
    """
    raise NotImplementedError


if __name__ == "__main__":
    if sys.version_info >= (3,):
        argv = sys.argv
    else:
        fse = sys.getfilesystemencoding()
        argv = [i.decode(fse) for i in sys.argv]

    #Apparently this thing's pretty magical.
    import plac
    plac.call(xls2db, argv[1:])
