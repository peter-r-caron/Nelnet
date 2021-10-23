'''
    Python filter program to execute SQL commands using sys, getopt, and pyodbc modules.
    Type "py sql.py" at command line to view syntax.

    Parameters:
        --colsep    Column separator on the output file.  Optional, default=",".  For tab delimited, use "\t"
        --database  Name of the ODBC datasource for database connection.  Required, no default.
        --encoding  Encoding of input/output files.  Optional, default="utf-8".
        --infile    Input SQL file name.  Optional, text may be piped in from console.
        --outfile   Output file name.  Optional, text will be written to the console if not used.
        --quote     Quote character around column values.  Optional, default='"'.  Eliminate quotes with --q=""

    Syntax: echo "select * from table_name" | py sql.py --database=DB_NAME > somefile.csv
    Alternate Syntax: py sql.py --colsep="|" --database="DB_NAME" --infile="sqlinput.sql" --outfile="output.csv" --quote='"'
    Short Syntax: py sql.py --c="," --d="MyDbName" --i="MySqlFile.sql" --o="MyOutputFile.csv" --q='"'

    Notes:
        On Windows, use the ODBC Data Sources app to name and configure the database connection.

        Double-Dash, (--) is required before command line parameter name.
        Equal sign, (=) is required between the command line parameter name and the parameter value.
        If unexpected results occur, check the command line for missing dashes and misspelling of parameter names.
        Put parameter values in quotes if they contain spaces or special characters.  Otherwise, they will be misinterpreted.

        If --infile is used, the SQL input file defaults to utf-8 encoding.  If your SQL input file is encoded with
        UTF-16 format, use --e="utf-16" in the command line parameters in order to read the infile properly.
        On Windows, you can change the file encoding of the SQL input file using notepad's "save as..." feature.

        For tab delimited output, use --colsep="\t" in the command line parameters.
'''

import sys
import getopt
import database
from database import Database

def _print_syntax():
    syntax = '''    Default syntax: echo "select * from table_name" | py sql.py --database=DB_NAME > somefile.csv
    Long Syntax:    py sql.py --colsep="|" --database="DB_NAME" --infile="sqlinput.sql" --outfile="output.csv" --quote='"'
    Short Syntax:   py sql.py --c="," --d="DbName" --i="SqlFile.sql" --o="OutputFile.csv" --q='"'
    Parameters:
        --colsep    Column separator on the output.  Optional, default=",".  For tab delimited, use "\\t".
        --database  Name of the ODBC datasource for database connection.  Required, no default.
        --encoding  Encoding of input/output file.  Optional, default="utf-8".
        --infile    Input SQL file name.  Optional, default=pipe SQL input from console.
        --outfile   Output file name.  Optional, default=print output to console.
        --quote     Quote character around column values.  Optional, default='"'.  Eliminate quotes with --q=\"\"
    '''
    print(syntax)
    sys.exit(1)

if len(sys.argv) < 2:
    _print_syntax()

# Set default values for command line parameters.
argv = sys.argv[1:]
COL_SEP = ','
DB_NAME = ''
ENCODING = 'utf-8'
IN_FILE = ''
OUT_FILE = ''
QUOTE='"'
SQL_STMT = ''

# Retrieve options from the command line.
opts, args = getopt.getopt(argv, "c:d:e:i:o:", ["colsep=","database=","encoding=","infile=","outfile=","quote="])
for opt, arg in opts:
    if opt in ("-c", "--colsep"):
        COL_SEP = arg
        if COL_SEP == "\\t":
            COL_SEP = "\N{TAB}"
    elif opt in ("-d", "--database"):
        DB_NAME = arg
    elif opt in ("-e", "--encoding"):
        ENCODING = arg
    elif opt in ("-i", "--infile"):
        IN_FILE = arg
    elif opt in ("-o", "--outfile"):
        OUT_FILE = arg
    elif opt in ("-q", "--quote"):
        QUOTE = arg

# Write the runtime parameters back to the console.
if OUT_FILE > '':
    print("Starting program " + sys.argv[0])
    print("-- Runtime Parameters --")
    print("  Column Seperator: " + COL_SEP)
    print("  Database Name:    " + DB_NAME)
    print("  Encoding:         " + ENCODING)
    print("  Input File:       " + IN_FILE)
    print("  Output File:      " + OUT_FILE)
    print("  Quote:            " + QUOTE)

# If a database name was provided, connect to the database.
if DB_NAME > '':
    pass
# Otherwise, fail with syntax in the error message.
else:
    print("--database parameter is a required value")
    _print_syntax()

# If an input file was passed from command line, open the file for reading
# and build the SQL statement.  Otherwise, build the SQL statement from console input.
if IN_FILE > '':
    try:
        with open(IN_FILE, "rt", encoding=ENCODING) as inf:
            for line in inf.readlines():
                SQL_STMT += line
    except FileNotFoundError as fail_error:
        print(fail_error)
        sys.exit(1)
else:
    for line in sys.stdin:
        SQL_STMT += line

# If an output file was passed from command line, open it for writing.
# Otherwise, any resultset will be output to stdout (console).
if OUT_FILE > '':
    try:
        OUT = open(OUT_FILE ,mode="wt",buffering=-1,encoding=ENCODING)
    except IOError as fail_error:
        print(fail_error)
        print("Unable to open " + OUT_FILE + " for writing.")
        sys.exit(1)
else:
    OUT = False

# Connect the database.
d = Database(DB_NAME)

# Execute the SQL statement.
rows = d.result_set(SQL_STMT)

# If rows were returned in the resultset, either write them to the output file, 
# or print them to the console, as directed by command line arguments.
if rows is None:
    pass
else:
    for row in rows:
        i = 0
        # Print each column value wrapped in quote characters 
        # and separated by the column separator.
        for x in row:
            if i > 0:
                if OUT:
                    OUT.write(COL_SEP)
                else:
                    print(",", end='')
            if OUT:
                OUT.write(QUOTE + str(x).strip() + QUOTE)
            else:
                print(QUOTE + str(x).strip() + QUOTE, end='')
            i += 1
        # Print a new line to end the tuple.
        if OUT:
             OUT.write("\n")
        else:
            print('')

# Clean up and exit.
if OUT:
    OUT.close()  # If an output file was created, close it.
d.close() # Close the cursor.

# Announce the end of the program.
if OUT_FILE > '':
    print("Ending program " + sys.argv[0] + " code 0")
sys.exit(0)
