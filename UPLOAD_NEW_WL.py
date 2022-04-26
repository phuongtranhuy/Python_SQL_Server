# -*- coding: utf-8 -*-
"""
Created on Wed May 19 17:40:30 2021

@author: TNP2HC
"""
import easygui
import BASE as base
import datetime
#*******************************SET UP SQL SERVER***********************************************************
sql = base.SQL_server()
conn, cursor = sql.SQl_connection('RemoteDB')
cursor.fast_executemany = True
sql_dml = base.SQL_DML(conn, cursor)
 
#********************************************************************************************************

link =easygui.fileopenbox(msg='Please Enter P94 Export Worklist Link', title='Link To excel',default=r'C:/Users/TNP2HC/Desktop/Division Data/', filetypes= '*.xlsx',multiple=False)
if link != None:
    
    print(link)
   
    df_upload_file = base.pd.read_excel(link,sheet_name='Sheet1',dtype=str)
    df_upload_file.fillna('',inplace=True)
    
    column_list1 = ['Created On','Changed On','Last Trigger Date','Upload time to SQL']
    column_list2 = ['Priority','Reliability Percentage','Number of similar rule based parts']
    
    for cols in column_list1:
        df_upload_file[cols]=df_upload_file[cols].replace('','1900-01-01 00:00:00.000')
        df_upload_file[cols] = df_upload_file[cols].astype('datetime64')
    
    for cols in column_list2:
        df_upload_file[cols] = df_upload_file[cols].replace('',0)
        df_upload_file[cols] = df_upload_file[cols].astype('float')              

    lst_column_name = df_upload_file.columns.tolist()
    print(lst_column_name)
    lst_column_name[24]= 'Prod#hierarchy'
    col_name,size_value = sql_dml.Transformed_data_insert(lst_column_name)
    
    sql = "SELECT [Product Number] FROM [dbo].[P94_worklist]" 

    df_all_wl = base.pd.read_sql(sql, conn)
    df_noexits_part = df_upload_file.merge(df_all_wl['Product Number'],how ='left',on ='Product Number',indicator = True)
    df_noexits_part = df_noexits_part[df_noexits_part['_merge']=='left_only']
           
    df_noexits_part = df_noexits_part.drop(['_merge'], axis=1)
    df_noexits_part['Upload time to SQL'] = datetime.now()
    
    
    numberofpart = df_noexits_part['Product Number'].count()
    print(numberofpart)    
    
    if numberofpart >0:
        sql_dml.write_to_sql(df_noexits_part,col_name,size_value,'P94_worklist')    
        sql_dml.replace_null('P94_worklist')
        easygui.msgbox(str(numberofpart) + '  part(s) updated', 'Upload new file')
    else:
        easygui.msgbox('These is no part to up', 'Upload new file')  
        
        
else:
    easygui.msgbox('You dont select file', 'Upload new file')
    
   

