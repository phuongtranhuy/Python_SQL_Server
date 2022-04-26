# -*- coding: utf-8 -*-
"""
Created on Tue Mar 30 15:35:54 2021

@author: TNP2HC
"""
option = input("Do you want to split MO data excel file?  y/n: ")

import easygui
import sys
from datetime import date
import BASE as base

#*******************************SET UP SQL SERVER***********************************************************
sql = base.SQL_server()
conn, cursor = sql.SQl_connection('RemoteDB')
cursor.fast_executemany = True
sql_dml = base.SQL_DML(conn, cursor) 
#********************************************************************************************************
    
if option == 'y': #_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
    
    link_MO_TABLE = easygui.fileopenbox(msg='Please Enter MO_TABLE EXCEL Link', title='Link To excel',default=r'C:\Users\TNP2HC\Desktop\Division Data\MO_TABLE_DATA', filetypes= '*.xlsx',multiple=False)
    
    if link_MO_TABLE == None  : 
        easygui.msgbox('Please select MO Data excel file') 
        sys.exit()
    elif link_MO_TABLE.find('MO') == -1:
        easygui.msgbox('Invalid file!!!') 
        sys.exit()    
    
    print(link_MO_TABLE)
    
    
    sql = "SELECT [Product Number] FROM [dbo].[MO_TABLE_DATA]"
    
    #-----------------------------SQL-----------------------------------------------------
    print('@Reading MO DATA excel file')
    MO_copy_columns = ['Logical System Group','Product Number','HS code','Hardness grade','Reference Number','Reference Group','Created on','REF']
    df_MO_TABLE  = base.pd.read_excel(link_MO_TABLE,'Sheet1', dtype= str)
    df_MO_TABLE = df_MO_TABLE.rename(columns = {"Number":"HS code","Created On":"Created on","Reference Number":"REF","Product Number.1":"Reference Number"})
    df_MO_TABLE = df_MO_TABLE[MO_copy_columns]
    
    #def Split_MO_TABLE_on_P94():
        
    link_not_duplicate_MO = 'C:/Users/TNP2HC/Desktop/Division Data/MO_TABLE_DATA/' + 'EXPORT_NEW_' + str(date.today()) + '.xlsx'
    link_duplicate_MO = 'C:/Users/TNP2HC/Desktop/Division Data/MO_TABLE_DATA/' + 'EXPORT_EXISTING_' + str(date.today())  + '.xlsx'
    
    df_all_MO = base.pd.read_sql(sql, conn)
    
    df_merge = df_MO_TABLE.merge(df_all_MO['Product Number'],how ='left',on ='Product Number',indicator = True)
    
    df_not_duplicate_MO = df_merge[df_merge['_merge']=='left_only']   
    df_duplicate_MO =  df_merge[df_merge['_merge']=='both']      
    
    
    df_not_duplicate_MO.to_excel(link_not_duplicate_MO, sheet_name='Sheet1' , index = False ,header=True)
    df_duplicate_MO.to_excel(link_duplicate_MO, sheet_name='Sheet1', index = False ,header=True)
    
    number_new = df_not_duplicate_MO['Product Number'].count()
    number_existing = df_duplicate_MO['Product Number'].count()
    
    print(str(number_new), ' parts for NEW MO DATA') 
    print(str(number_existing), ' parts for EXISTING MO DATA') 
    #Split_MO_TABLE_on_P94()
    
elif option == 'n':#_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
    
    option_ = input("INSERT new data or UPDATE existing data? new/existing: ")
    
    if option_ == 'new':#_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
        link_MO_TABLE_NEW = easygui.fileopenbox(msg='Please Enter MO_TABLE NEW Link', title='Link To excel',default=r'C:\Users\TNP2HC\Desktop\Division Data\MO_TABLE_DATA', filetypes= '*.xlsx',multiple=False)
        link_Backbone = easygui.fileopenbox(msg='Please Enter BACKBONE EXCEL Link', title='Link To excel',default=r'C:\Users\TNP2HC\Desktop\Division Data\MO_TABLE_DATA', filetypes= '*.xlsx',multiple=False)
        
        if link_MO_TABLE_NEW == None or link_Backbone == None: 
            easygui.msgbox('Please select at least MO & Backbone excel file') 
            sys.exit()
         
        if link_MO_TABLE_NEW.find('NEW') == -1: 
            easygui.msgbox('Please select MO NEW DATA only') 
            sys.exit()
        
        if link_Backbone.find('Spreadsheet') == -1: 
            easygui.msgbox('Please select BACKBONE Spreadsheet only') 
            sys.exit()
        
        excel_template = 'C:/temp/SQL table template.XLSX'
        termcode_hierachy_text = r'C:\Users\TNP2HC\Desktop\Division Data\EXCEL DATA\ProductHierarchy.xlsx'
        columns = ['Product Hierarchy','Definition']
        
        MO_copy_columns = ['Logical System Group','Product Number','HS code','Hardness grade','Reference Number','Reference Group','Created on']
        
        Backbone_copy_columns = ['Material Number','Material Description','Term Code','Product Hierarchy']
        
        print('@Reading from excel file')
        df_MO_TABLE_NEW  = base.pd.read_excel(link_MO_TABLE_NEW,'Sheet1', dtype= str)
        df_MO_TABLE_NEW = df_MO_TABLE_NEW[MO_copy_columns]
        
        df_BACKBONE  = base.pd.read_excel(link_Backbone,'Data', usecols = Backbone_copy_columns, dtype= str)
        df_BACKBONE = df_BACKBONE.rename(columns = {"Material Number":"Product Number","Product Hierarchy":"Hierarchy"})   
        df_BACKBONE.drop_duplicates(subset ="Product Number",keep = 'first', inplace = True)
                                                     
        
        print('@Reading excel file DONE!')
          
        #-----------------------------------Read data from excel ------------------------------
        print('@Reading from hierarchy excel file')
        df_template = base.pd.read_excel(excel_template,"SQL_MO_TABLE", dtype= str)
        
        df_term_text = base.pd.read_excel(termcode_hierachy_text,"Termcode", dtype= str)
        df_term_text = df_term_text.rename(columns = {"Term-Code-Num":"Term code"})
        df_term_text.drop_duplicates(subset ="Term code",keep = 'first', inplace = True)
        
        df_hie_PT = base.pd.read_excel(termcode_hierachy_text,"PT",usecols = columns, dtype= str)
        df_hie_BR = base.pd.read_excel(termcode_hierachy_text,"DC", usecols = columns,dtype= str)
        df_hie_TT = base.pd.read_excel(termcode_hierachy_text,"TT",usecols = columns, dtype= str)
        
        df_hie_ALL= base.pd.concat([df_hie_PT, df_hie_BR ,df_hie_TT])
        df_hie_ALL.drop_duplicates(subset ="Product Hierarchy",keep = 'first', inplace = True)
        
        print('@Reading hierarchy DONE!')
        #--------------------------------------------------------------------------------------
                  
        for cols in list(df_MO_TABLE_NEW.columns): df_template[cols] = df_MO_TABLE_NEW[cols]
             
        
        
        df_template['Product short text'] = base.pd.merge(df_template,df_BACKBONE , on = 'Product Number',how = 'left')['Material Description']
        df_template['Term code'] = base.pd.merge(df_template, df_BACKBONE, on = 'Product Number',how = 'left')['Term Code']
        df_template['Product Hierarchy'] = base.pd.merge(df_template, df_BACKBONE, on = 'Product Number',how = 'left')['Hierarchy']
        
        df_template['Term code text'] = base.pd.merge(df_template, df_term_text, on = 'Term code',how = 'left')['Term']
        df_template['Hierachy Definition'] = base.pd.merge(df_template, df_hie_ALL, on = 'Product Hierarchy',how = 'left')['Definition']
        
        df_template['Created on'] = [time[0:4] + '-' + time[4:6] + '-' + time[6:8] +  ' 00:00:00' for time in df_template['Created on']]
        
        lst_col1 = ['Upload time to P94','PIC upload time']        
        for cols in lst_col1: df_template[cols] = '1900-01-01 00:00:00' 
        
        df_template.fillna('',inplace=True)
        
        lst_col2 = ['Risk point','OLD RISK POINT','Priority_termcode']
        for cols in lst_col2:
            df_template[cols]=df_template[cols].replace('',0)
            df_template[cols]=df_template[cols].astype('int')
            
        #INSER TO SQL BELOW 
        lst_column_name = df_template.columns.tolist()
        col_name,size_value = sql_dml.Transformed_data_insert(lst_column_name)
        
        numberofpart = df_template['Product Number'].count()
        print(numberofpart)    
        
        if numberofpart >0:
            sql_dml.write_to_sql(df_template,col_name,size_value,'MO_TABLE_DATA')    
            sql_dml.replace_null('MO_TABLE_DATA')
            easygui.msgbox(str(numberofpart) + '  part(s) updated', 'Upload new file')           
        else:
            easygui.msgbox('These is no part to upload', 'Upload new file') 
        
        
    elif option_ == 'existing': #_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-  
        
        link_MO_TABLE_EXISTING = easygui.fileopenbox(msg='Please Enter MO_TABLE EXISTING Link', title='Link To excel',default=r'C:\Users\TNP2HC\Desktop\Division Data\MO_TABLE_DATA', filetypes= '*.xlsx',multiple=False)
        
        if link_MO_TABLE_EXISTING == None: 
            easygui.msgbox('Please select MO excel file') 
            sys.exit()
        
        if link_MO_TABLE_EXISTING.find('EXISTING') == -1: 
            easygui.msgbox('Please select MO EXISTING DATA only') 
            sys.exit()
            
        print('@Reading from excel file')
        df_MO_TABLE_EXISTING  = base.pd.read_excel(link_MO_TABLE_EXISTING,'Sheet1', dtype= str)
        df_MO_TABLE_EXISTING = df_MO_TABLE_EXISTING[df_MO_TABLE_EXISTING['REF'] != 'X']
        df_MO_TABLE_EXISTING = df_MO_TABLE_EXISTING[['Logical System Group','Product Number','HS code','Hardness grade','Reference Number','Reference Group']]
        
        sql_with_ref = "SELECT [Product Number],[Reference Number] FROM [dbo].[MO_TABLE_DATA]"
        df_all_with_REF_MO = base.pd.read_sql(sql_with_ref, conn)
       
        df_MO_TABLE_EXISTING = df_MO_TABLE_EXISTING.rename(columns = {"Reference Number":"Ref part"})
        
        df_MO_TABLE_EXISTING = df_MO_TABLE_EXISTING.merge(df_all_with_REF_MO[['Product Number', 'Reference Number']], 'left')
        
        df_ref_changed = df_MO_TABLE_EXISTING[df_MO_TABLE_EXISTING['Ref part'] != df_MO_TABLE_EXISTING['Reference Number']]
        
        df_ref_changed = df_ref_changed.dropna()
        
        numberofpart = df_ref_changed['Product Number'].count()
        
        if numberofpart >0: 
            a = 1
            for ind,row in df_ref_changed.iterrows():
  
                sql_update = " UPDATE [dbo].[MO_TABLE_DATA] SET " + \
                                        " [HS code] = " +  str(row['HS code']) + "," + \
                                        " [Reference Number] = '" + str(row['Ref part']) + "'," + \
                                        " [Reference Group] = '"  + str(row['Reference Group']) + "'," + \
                                        " [Hardness grade] = '"  + str(row['Hardness grade']) + "' " + \
                                        " WHERE [Product Number] = '" + str(row['Product Number']) + "'"
                                                                            
                cursor.execute(sql_update)
                conn.commit()
                
                sys.stdout.write('\r'+ str(a)  + ' Parts' + ': '+ str("upload done")) 
                a += 1

                                                            
        else:
            print('No Reference Changed have been found')
        
        
    else:
        easygui.msgbox('Invalid command. Abort function!') 
        sys.exit() 
else:
    easygui.msgbox('Invalid command. Abort function!') 
    sys.exit() 


























      


# =============================================================================
# excel_template = 'C:/temp/SQL table template.XLSX'
# termcode_hierachy_text = 'C:/Users/TNP2HC/Desktop/VBA training/ProductHierarchy.xlsx'
# columns = ['Product Hierarchy','Definition']
# MO_copy_columns = ['Logical System Group','Product Number','HS code','Hardness grade','Reference Number','Reference Group','Created on']
# Backbone_copy_columns = ['Material Number','Material Description','Term Code','Product Hierarchy']
# 
# print('@Reading from excel file')
# df_MO_TABLE  = base.pd.read_excel(link_MO_TABLE,'Sheet1', dtype= str)
# df_BACKBONE  = base.pd.read_excel(link_Backbone,'Data', usecols = Backbone_copy_columns, dtype= str)
# df_MO_TABLE = df_MO_TABLE.rename(columns = {"Number":"HS code","Created On":"Created on","Reference Number":"REF","Product Number.1":"Reference Number"})
# 
# df_MO_TABLE = df_MO_TABLE[MO_copy_columns]
# df_BACKBONE = df_BACKBONE.rename(columns = {"Material Number":"Product Number","Product Hierarchy":"Hierarchy"})   
# 
# print('@Reading excel file DONE!')
# 
# 
# 
# def INSERT_NEW_MO_TABLE_DATA():
#     
#     #-----------------------------------Read data from excel ------------------------------
#     print('@Reading from database excel file')
#     df_template = base.pd.read_excel(excel_template,"SQL_MO_TABLE", dtype= str)
#     
#     df_term_text = base.pd.read_excel(termcode_hierachy_text,"Termcode", dtype= str)
#     df_term_text = df_term_text.rename(columns = {"Term-Code-Num":"Term code"})
#     
#     
#     df_hie_PT = base.pd.read_excel(termcode_hierachy_text,"PT",usecols = columns, dtype= str)
#     df_hie_BR = base.pd.read_excel(termcode_hierachy_text,"DC", usecols = columns,dtype= str)
#     df_hie_TT = base.pd.read_excel(termcode_hierachy_text,"TT",usecols = columns, dtype= str)
#     
#     df_hie_ALL= base.pd.concat([df_hie_PT, df_hie_BR ,df_hie_TT])
#     print('@Reading database DONE!')
#     #--------------------------------------------------------------------------------------
#     
#     
#     for cols in list(df_MO_TABLE.columns): df_template[cols] = df_MO_TABLE[cols]
#          
#     
#     df_template['Product short text'] = base.pd.merge(df_template, df_BACKBONE, on = 'Product Number',how = 'left')['Material Description']
#     df_template['Term code'] = base.pd.merge(df_template, df_BACKBONE, on = 'Product Number',how = 'left')['Term Code']
#     df_template['Product Hierarchy'] = base.pd.merge(df_template, df_BACKBONE, on = 'Product Number',how = 'left')['Hierarchy']
#     
#     df_template['Term code text'] = base.pd.merge(df_template, df_term_text, on = 'Term code',how = 'left')['Term']
#     df_template['Hierachy Definition'] = base.pd.merge(df_template, df_hie_ALL, on = 'Product Hierarchy',how = 'left')['Definition']
#     
#     df_template['Created on'] = [time[0:4] + '-' + time[4:6] + '-' + time[6:8] +  ' 00:00:00' for time in df_template['Created on']]
#     
#     df_template['Upload time to P94'] = '1900-01-01 00:00:00'
#     df_template['PIC upload time'] = '1900-01-01 00:00:00'
# =============================================================================
            
        


    
#INSERT_NEW_MO_TABLE_DATA(df_not_duplicate_MO) 
   

    


   
    
    
        #df_template['Product short text'] = [df_BACKBONE[df_BACKBONE['Product Number'] == i]['Product short text'].values[0] for i in df_template['Product Number']]   
        #df_template['Term code'] = [df_BACKBONE[df_BACKBONE['Product Number'] == i]['Term code'].values[0] for i in df_template['Product Number']] 
        #df_template['Product Hierarchy'] = [df_BACKBONE[df_BACKBONE['Product Number'] == i]['Product Hierarchy'].values[0] for i in df_template['Product Number']]   
    
    
    
    
    