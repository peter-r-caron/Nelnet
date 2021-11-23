'''
    Notes:
        On Windows, use the ODBC Data Sources app to name and configure the database connection.
'''

import sys
import pyodbc

class Database:
    """ Initialize a connection a SQL Server database. """
    def __init__(self, datasource):
        self.datasource = datasource
        self.connection = pyodbc.connect('DSN=' + self.datasource)
        self.cursor = self.connection.cursor()

    """ Close the database connection. """
    def close(self):
        self.connection.close()

    """ Return a list of table column names matching a wildcard pattern. """
    def columns(self, table, catalog, schema, column):
        return self.cursor.columns(table,catalog,schema,column)

    """ Commit changes to the database. """
    def commit(self):
        self.cursor.commit()

    """ Execute an SQL statement. """
    def execute(self, statement, commit):
        if self.cursor.execute(statement).rowcount > -1:
            if commit is True:
                self.commit()

    """ Use the pyodbc getinfo function to return the name of the database system. """
    def get_sql_dbms_name(self):
        return self.connection.getinfo(pyodbc.SQL_DBMS_NAME)

    """ Use the pyodbc getinfo function to return the version of the database system. """
    def get_sql_dbms_ver(self):
        return self.connection.getinfo(pyodbc.SQL_DBMS_VER)

    """ Use the pyodbc getinfo function to return the name of the database driver. """
    def get_sql_driver_name(self):
        return self.connection.getinfo(pyodbc.SQL_DRIVER_NAME)

    """ Execute the SQL statement and return all rows as a pyodbc resultset. """
    def result_set(self, statement):
        self.execute(statement, True)
        if self.cursor.rowcount < 0:
            return self.cursor.fetchall()
        else:
            return None

    """ Rollback changes to the database. """
    def rollback(self):
        self.cursor.rollback()

    """ Reinitialize the database connection from the given datasource. """
    def set(self, datasource):
        self.__init__(datasource)

    """ Return a list of database table names matching a wildcard pattern. """
    def tables(self, table_name, catalog, schema, type):
        return self.cursor.tables(table_name, catalog, schema, type)