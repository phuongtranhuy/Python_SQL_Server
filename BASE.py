# -*- coding: utf-8 -*-
"""
Created on Wed May 19 10:49:51 2021

@author: TNP2HC
"""

#*******************************IMPORT LIBRARY*********************************
import pyodbc 
import pandas as pd
import win32com.client
import win32ui

#*******************************SET UP SQL CONNECTION**************************

class SQL_server:
    def __init__(self):
        self.connection_str = {'RemoteDB':'Driver={ODBC Driver 17 for SQL Server};'
                                          'Server=SQL58.company,1433;'
                                          'Database=DB_C1_SQL;'
                                          'UID=WOM.C1-INT;'
                                          'Trusted_Connection=yes;',
                                          
                               'LocalDB':'Driver={SQL Server};'
                                         'Server=HC-UT40632N\SQLEXPRESS;'
                                         'Database=DB_PEP_OPERATION;'
                                         'Trusted_Connection=yes;'}
        
    def SQl_connection(self,DB_type):
        conn = pyodbc.connect(self.connection_str[DB_type])
        cursor = conn.cursor()        
        
        print('Set up connection successful')       
        return conn, cursor


class SQL_DML:
    def __init__(self,conn,cursor):
        self.conn = conn
        self.cursor = cursor
        
    def replace_null(self,sql_table):
        sql = "SELECT TOP (0) * FROM dbo.[" + str(sql_table) + "]"
        df = pd.read_sql(sql, self.conn)
        columns = df.columns
        for col in columns:
            sql = sql + "UPDATE " + str(sql_table) + " SET [" + str(col) + "] = ''" + "WHERE [" + str(col) + "] IS NULL;"
        self.cursor.execute(sql)
        self.conn.commit()
        return
      
    def delete_table(self,sql_table):
        sql = "TRUNCATE TABLE" + str(sql_table)
        self.cursor.execute(sql)
        self.conn.commit()
        return
    
    
    def Transformed_data_insert(self,lst_cols_name): #define how many cols need to input data

        col_name ='' 
        size_value = ''
        
        for col in lst_cols_name:
            if len(col)>2:
                col_name = col_name + '[' + str(col) +'],'
                size_value = size_value + "?,"
        size_value = '(' + size_value[0:len(size_value)-1]+ ")"
        col_name = '(' + col_name[0:len(col_name)-1]+ ")"
        
        return col_name,size_value 
    
    def write_to_sql(self,df,col_name,size_value,sql_table):
        sql = "INSERT INTO [" + str(self.sql_table) + "] " +  col_name  + ' VALUES ' + size_value 
        val = tuple(df.values)
        array_of_tuples = map(tuple, val)
        tuple_of_tuples = tuple(array_of_tuples)
        val = tuple_of_tuples
        self.cursor.executemany(sql,val)
        self.conn.commit()
        return    

    
#*******************************EXCEL MODULE***********************************
    
def Open_EXCEL_File(folder_path,sheet_name):
    
    xl= win32com.client.Dispatch("Excel.Application") # read active workbook (pep file)
    xl.Visible = True
    dlg = win32ui.CreateFileDialog (1) # 1 represents the Open File dialog box
    dlg.SetOFNInitialDir (folder_path) # Initial Set File Open dialog box displays the directory
    dlg.DoModal()
     
    excel_path = dlg.GetPathName () # get the selected file name
    wb = xl.Workbooks.Open(excel_path)
    wb.Worksheets(sheet_name).Visible = True
    ws = wb.Sheets(sheet_name)
    return ws



