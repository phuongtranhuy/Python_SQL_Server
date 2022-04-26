# -*- coding: utf-8 -*-
"""
Created on Mon Mar  7 14:30:35 2022

@author: TNP2HC
"""

import numpy as np
import BASE as base
#*******************************SET UP SQL SERVER***********************************************************
sql = base.SQL_server()
conn, cursor = sql.SQl_connection('RemoteDB')

sql_dml = base.SQL_DML(conn, cursor) 
#********************************************************************************************************

sql_str = """select [Logical System Group] + '><' + [Product Number] + '<>' + replace([Change history],'#','') as [Change history]
				from P94_worklist where [Upload time to SQL] >= '2021-01-01' and [Change history] <> '' """
                                                                                                        
df_history_changed = base.pd.read_sql(sql_str, conn)

def add_log_prod_number(lst):
    #print(lst)
    new_list = [(lst[0][:lst[0].find('<>')] + '<>' if lst.index(x) > 0 and len(x) > 1 else '') + x for x in lst]        
    return new_list

df_history_changed['Change history'] = df_history_changed['Change history'].apply(lambda x: x.split(';'))
df_history_changed['Change history'] = df_history_changed['Change history'].apply(add_log_prod_number)

df_history_changed = df_history_changed.explode('Change history')           
df_history_changed['Change history'].replace('', np.nan, inplace=True)  
df_history_changed.dropna(subset=['Change history'], inplace=True)

df_history_changed['Change history'] = df_history_changed['Change history'].str.replace('<>','|')
df_history_changed['Change history'] = df_history_changed['Change history'].str.replace('><','|')
df_history_changed = df_history_changed['Change history'].str.split("|", n = 5, expand = True) 

df_history_changed.rename(columns={0:'Logical Group', 1:'Product Number', 2:'PIC classification', 3:'Product Group', 4:'Task', 5:'Upload time to SQL'},inplace=True)
df_history_changed['Upload time to SQL'] = df_history_changed['Upload time to SQL'].str[:10]

#df_history_changed['Upload time to SQL'] = pd.to_datetime(df_history_changed['Upload time to SQL'])


#df_history_changed.to_excel(r'C:\Users\TNP2HC\Desktop\POWER BI\PEP_KPI.xlsx', sheet_name='Sheet1', index = False ,header=True)
#df_test = df_history_changed.drop_duplicates(subset ="Upload time to SQL",keep = False, inplace = True)

#test_list = ['Y_UBK><0001030746<>Thien An|UNCLASSIFY|2nd Classification|04-05-2021 12 46 PM|', 'Duc Thanh|750 W < DC MOTOR <= 3.75 KW|2nd Classification|24-05-2021 04 23 PM|', '']
#new_list = test_list[0][:test_list[0].find('<>')] + str(s) for s in test_list if test_list.index(s) > 0]

#str_test = test_list[0].find('<>')
#get_str = test_list[0][:test_list[0].find('<>')]         
# =============================================================================
# input_list = ['VOD LN','HSBA LN', 'DOKA SS','SXNE', 'KERIN FH','YORK GY','SXNP']
#add_log_prod_number_ = [(test_list[0][:test_list[0].find('<>')] + '<>' if test_list.index(x) > 0 and len(x) > 1 else '') + x for x in test_list]
# result  
# =============================================================================
#new_list = add_log_prod_number(test_list)  
    