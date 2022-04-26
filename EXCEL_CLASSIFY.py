# -*- coding: utf-8 -*-
"""
Created on Wed May 19 17:40:30 2021

@author: TNP2HC
"""
option = input("Input Logical System Group to execute: ")

import BASE as base
from IPython.display import clear_output
#*******************************SET UP SQL & CALL EXCEL FILE**************************************
sql = base.SQL_server()
conn, cursor = sql.SQl_connection('RemoteDB')

ws = base.Open_EXCEL_File(r'U:\PEP FILE',"Sheet1")
df_hint = base.pd.read_excel(r'C:\Users\TNP2HC\Desktop\Division Data\EXCEL DATA\Logical_Data.xlsx','Sheet1',dtype= str)

#********************************************************************************************
def get_data_from_hsdatabase(Ref_number):
    sql_hsdatabase = "SELECT [PART NUMBER],[PRODUCT_GROUP],[HS CODE] FROM [DB_CTXFC1_SQL].[dbo].[hscodedata]"
    
    df_hscodedata = base.pd.read_sql(sql_hsdatabase, conn)
    df_hscodedata = df_hscodedata[(df_hscodedata['PART NUMBER'] != '') & (df_hscodedata['HS CODE'] != '')]
    
    df_ref_group = df_hscodedata[df_hscodedata['PART NUMBER'] == str(Ref_number)]
    
    ref_group = df_ref_group['PRODUCT_GROUP'].values[0] if df_ref_group.empty == False else ''
      
    return ref_group

def print_and_overwrite(text):
    '''Remember to add print() after the last print that you want to overwrite.'''
    clear_output(wait=True)
    print(text, end='\r')
    
#*******************************CLASSIFICATION BASED ON TEXT*****************************
def Classify_By_Text_Data_PT(termcode,drawing,compare_str,df_hint_logical): 
    
    comment_by_BOT,product_group = "",""
    key_matched, DRW_got = [],[]
    
    drw_matched, con = False,False
            
    drawing_split_ = drawing.strip().split(";")
    df_group_termcode = df_hint_logical[df_hint_logical["Term code"]== termcode]    

    if df_group_termcode.empty == False:       

        for ind, row in df_group_termcode.iterrows():
            
            if not str(row["Drawing Number"]) == '' or not str(row["Drawing Number"]) == None: 

                DRW_split = str(row["Drawing Number"]).strip().split(";")
                
                for drw in DRW_split:
                    if len(drw) > 1:
                        drw = drw.replace('\xa0',"")
                        if len(drawing_split_) == 1:
                            if drw.strip() == drawing.strip():
                                drw_matched, con = True, True
                                product_group = row["Product Group"]
                                DRW_got.append(drw.strip())
                                break
                    
                    
                        elif len(drawing_split_) > 1:
                            for each_drw in drawing_split_:
                                for drw in DRW_split:
                                    if len(drw) > 1 and len(each_drw) > 1:
                                        drw = drw.replace('\xa0',"")
                                        each_drw = each_drw.replace('\xa0',"")
                                                                                   
                                        if drw.strip() == each_drw.strip():
                                            drw_matched, con = True, True
                                            product_group = row["Product Group"]
                                            DRW_got.append(drw.strip())
                                            break
                    
                    if drw_matched == True: break    
                        
            
            if str(row["Keyword"]) != '' or str(row["Keyword"]) != None:
                keyword = str(row["Keyword"])
                KEY_split = keyword.strip().split(";")
                
                for element in KEY_split:
                    
                    if len(element) > 1:
                        if not (compare_str == None) or not (compare_str == ""):
                            
                            if compare_str.upper().find(element.upper().strip()) != -1:
                                con = True
                                product_group = row["Product Group"]
                                key_matched.append(element.strip())
                            
            if con == True:break        
            
  
    if len(key_matched) > 0 or drw_matched == True :
        
        comment_by_BOT =  "- Keyword matched: {0}\n- Drawing matched: {1}".format(key_matched,DRW_got)

    return product_group,comment_by_BOT
###############################################################################

def Classify_By_Text_Data_DC(config_material,motor_code,weight,df_hint_logical): 

    comment_by_BOT,product_group = "",""            
    # MATCH ZKMA NUMBER AND MOTOR CODE        
    if config_material != "":
    
        df_result = df_hint_logical[df_hint_logical['Configurable Material'].str.contains(config_material, na=False)]
        
        if len(df_result) == 1:                
            product_group = df_result.values[0][1]
            comment_by_BOT = 'Configuration Material Matched'               
        else:
            comment_by_BOT = '2nd Classification - BREX material'
    else:
        df_result = df_hint_logical[df_hint_logical['Keyword'].str.contains(motor_code,na=False)]
        
        if len(df_result) == 1:
            if df_result.values[0][1].find('74.6<P<=735W') > -1:
                convert_w = float(weight[0: weight.find('kg') - 1].replace(',','.'))
                if convert_w > 5:
                    product_group = 'DC-LT-MOTOR 3 PHASE (W>5KG, 74.6<P<=735W)'
                else:
                    product_group = 'DC-LT-MOTOR 3 PHASE (1<W<=5KG, 74.6<P<=735W)'
            else:
                product_group = df_result.values[0][1]
                
            comment_by_BOT = 'Motor Code Matched'    
            
    return product_group , comment_by_BOT    
############################################################################### 
       
def Classify_By_Text_Data_TT01(hierarchy,compare_str,df_hint_logical): 

    comment_by_BOT,product_group = "",""
        
    # MATCH KEYWORD GROUP BY PRODUCT HIERARCHY       
    df_group_hierarchy = df_hint_logical[df_hint_logical["Product Hierarchy"]== hierarchy]    
    if df_group_hierarchy.empty == False:
        for ind, row in df_group_hierarchy.iterrows():
            
            if str(row["Keyword"]) != '' or str(row["Keyword"]) != None:
                keyword = row["Keyword"]                       
                if not (compare_str == None) or not (compare_str == ""):
                    
                    if compare_str.upper().find(keyword.upper().strip()) > -1:
                        product_group = row["Product Group"]
                        comment_by_BOT = 'Keyword matched (*Auto): ' + str(keyword)
                
    return product_group , comment_by_BOT

#*******************************CLASSIFICATION BASED ON RULLING*****************************
def Classify_By_Statistics_Rulling(termcode,number_range,pro_hierarchy,df_rule,df_stats):
       
    #Logical : Y_PT, Y_BR______________________________________________________
    df_rule_termcode = df_rule[(df_rule['Term code'] != '___') & (df_rule['Number Range'] == '___') & (df_rule['Product Hierarchy'] == '___')]
    
    #Logical : Y_PT, Y_BR, Y_TT01, Y_ST______________________________________________________
    df_rule_hierarchy = df_rule[(df_rule['Term code'] == '___') & (df_rule['Number Range'] == '___') & (df_rule['Product Hierarchy'] != '___')]
    
    #Logical : Y_PT, Y_BR ______________________________________________________
    df_rule_termcode_hierarchy = df_rule[(df_rule['Term code'] != '___') & (df_rule['Number Range'] == '___') & (df_rule['Product Hierarchy'] != '___')]
    
    #Logical : Y_PT, Y_BR______________________________________________________
    df_rule_range_termcode = df_rule[(df_rule['Term code'] != '___') & (df_rule['Number Range'] != '___') & (df_rule['Product Hierarchy'] == '___')]
    
    #Logical :Y_PT, Y_TT01, Y_BR, Y_ST______________________________________________________
    df_rule_range_hierarchy = df_rule[(df_rule['Term code'] == '___') & (df_rule['Number Range'] != '___') & (df_rule['Product Hierarchy'] != '___')]
    
    
    df_filter_term = df_rule_termcode[df_rule_termcode['Term code'] == termcode] 
    df_filter_hierarchy = df_rule_hierarchy[df_rule_hierarchy['Product Hierarchy'] == pro_hierarchy] 
    df_filter_term_hie = df_rule_termcode_hierarchy[(df_rule_termcode_hierarchy['Term code'] == termcode) & (df_rule_termcode_hierarchy['Product Hierarchy'] == pro_hierarchy)] #combine Termcode & Hierachy
    df_filter_term_numRange = df_rule_range_termcode[(df_rule_range_termcode['Number Range'] == number_range) & (df_rule_range_termcode['Term code'] == termcode)] # combine number range & term code
    df_filter_hie_numRange = df_rule_range_hierarchy[(df_rule_range_hierarchy['Number Range'] == number_range) & (df_rule_range_hierarchy['Product Hierarchy'] == pro_hierarchy)] # combine number range & product hierachy

    percentage_follow_ref, amount_follower = 0,0
    product_group, comment_by_BOT, need_check = '','',''
    
    lst_num_follower = []
    lst_group = []
    lst_check = []
    lst_rule_name = [] 
    lst_ref = []

    if not df_filter_term.empty:
        lst_num_follower.append(df_filter_term.values[0][4])

        need_check = False if df_filter_term.values[0][4] >= 25 else True

        lst_group.append(df_filter_term.values[0][5])
        lst_check.append(need_check)
        lst_rule_name.append('Termcode')
        lst_ref.append(df_filter_term.values[0][3])
        
    if not df_filter_hierarchy.empty:
        lst_num_follower.append(df_filter_hierarchy.values[0][4])
        
        need_check = False if df_filter_hierarchy.values[0][4] >= 25 else True
        
        lst_group.append(df_filter_hierarchy.values[0][5])
        lst_check.append(need_check)
        lst_rule_name.append('Hierarchy')
        lst_ref.append(df_filter_hierarchy.values[0][3])
        
    if not df_filter_term_hie.empty:    
        lst_num_follower.append(df_filter_term_hie.values[0][4])
        
        need_check = False if df_filter_term_hie.values[0][4] >= 25 else True
        
        lst_group.append(df_filter_term_hie.values[0][5])
        lst_check.append(need_check)
        lst_rule_name.append('Termcode-Hierarchy')
        lst_ref.append(df_filter_term_hie.values[0][3])
        
    if not df_filter_term_numRange.empty: 
        lst_num_follower.append(df_filter_term_numRange.values[0][4])
        
        need_check = False if df_filter_term_numRange.values[0][4] >= 25 else True
        
        lst_group.append(df_filter_term_numRange.values[0][5])
        lst_check.append(need_check)
        lst_rule_name.append('Termcode-NumRange')
        lst_ref.append(df_filter_term_numRange.values[0][3])
        
    if not df_filter_hie_numRange.empty: 

        lst_num_follower.append(df_filter_hie_numRange.values[0][4])
        
        need_check = False if df_filter_hie_numRange.values[0][4] >= 25 else True
    
        lst_group.append(df_filter_hie_numRange.values[0][5])
        lst_check.append(need_check)
        lst_rule_name.append('Hierarchy-NumRange')
        lst_ref.append(df_filter_hie_numRange.values[0][3])
    
    if lst_group:
        #Check the highest num_follower from each rules to pick up ________________________
        ind_max_follower = lst_num_follower.index(max(lst_num_follower))
        
        product_group = lst_group[ind_max_follower]
        need_check = '' if lst_check[ind_max_follower] == False else 'Need to check'
       
        percentage_follow_ref = 100
        amount_follower = max(lst_num_follower)
        rule_name = lst_rule_name[ind_max_follower]
        comment_by_BOT = str(lst_ref[ind_max_follower]) + '>Total part similar: ' + str(amount_follower) + '>Rule Type: ' + rule_name
    
    else: # OUT OF RULLING - STATISTICS PHASE
        get_ref, percentage_follow_ref, amount_follower, rule_type = Statistics_For_AutoQC(termcode,number_range,pro_hierarchy,df_stats)
        if len(get_ref) >= 10:
            comment_by_BOT = str(get_ref) + '_' + str(round(percentage_follow_ref)) + '>Total part similar: ' + str(amount_follower) + '>Rule Type: ' + rule_type
            #Special case if percentage_follow_ref > 90% --> move to 1st check
            if float(percentage_follow_ref) >= 90:          
                product_group = get_data_from_hsdatabase(get_ref) 
                need_check = 'Need to check' if product_group != '' else ''
                #need_check = lambda x : ('','Need to check')[x != '']
        else:
            comment_by_BOT = 'No rules found'
            
    return product_group, need_check, comment_by_BOT, round(percentage_follow_ref), amount_follower


# =============================================================================
# =============================================================================
# termcode = '268313'
# number_range = 'R901'
# pro_hierarchy = '111131103'
# # 
# product_group, need_check, comment_by_BOT, percentage_follow_ref, amount_follower = Classify_By_Statistics_Rulling(termcode,number_range,pro_hierarchy,df_rule,df_stats)
# =============================================================================
# =============================================================================


def Statistics_For_AutoQC(termcode,number_range,pro_hierarchy,df_stats):
    percentage_follow_ref, amount_follower = 0,0
    get_ref, rule_type = '',''
    
    df_stats_termcode = df_stats[(df_stats['Term code'] != '___') & (df_stats['Number Range'] == '___') & (df_stats['Product Hierarchy'] == '___')]

    df_stats_hierarchy = df_stats[(df_stats['Term code'] == '___') & (df_stats['Number Range'] == '___') & (df_stats['Product Hierarchy'] != '___')]
    
    df_stats_termcode_hierarchy = df_stats[(df_stats['Term code'] != '___') & (df_stats['Number Range'] == '___') & (df_stats['Product Hierarchy'] != '___')]
    
    df_stats_range_termcode = df_stats[(df_stats['Term code'] != '___') & (df_stats['Number Range'] != '___') & (df_stats['Product Hierarchy'] == '___')]
    
    df_stats_range_hierarchy = df_stats[(df_stats['Term code'] == '___') & (df_stats['Number Range'] != '___') & (df_stats['Product Hierarchy'] != '___')]
    
    
    df_filter_term = df_stats_termcode[df_stats_termcode['Term code'] == termcode] 
    df_filter_hierarchy = df_stats_hierarchy[df_stats_hierarchy['Product Hierarchy'] == pro_hierarchy] 
    df_filter_term_hie = df_stats_termcode_hierarchy[(df_stats_termcode_hierarchy['Term code'] == termcode) & (df_stats_termcode_hierarchy['Product Hierarchy'] == pro_hierarchy)] #combine Termcode & Hierachy
    df_filter_term_numRange = df_stats_range_termcode[(df_stats_range_termcode['Number Range'] == number_range) & (df_stats_range_termcode['Term code'] == termcode)] # combine number range & term code
    df_filter_hie_numRange = df_stats_range_hierarchy[(df_stats_range_hierarchy['Number Range'] == number_range) & (df_stats_range_hierarchy['Product Hierarchy'] == pro_hierarchy)] # combine number range & product hierachy
    
    Ls_percent = []
    Ls_num_follow = []
    Ls_rule_name = [] 
    Ls_ref = []
    
    if df_filter_term.empty == False:
        ind_max = df_filter_term['Percentage'].idxmax()
               
        Ls_percent.append(df_filter_term['Percentage'].max())
        Ls_num_follow.append(df_filter_term.loc[ind_max].values[4])
        Ls_rule_name.append('Termcode')
        Ls_ref.append(df_filter_term.loc[ind_max].values[3])
        
        
    if df_filter_hierarchy.empty == False:
        ind_max = df_filter_hierarchy['Percentage'].idxmax()
               
        Ls_percent.append(df_filter_hierarchy['Percentage'].max())
        Ls_num_follow.append(df_filter_hierarchy.loc[ind_max].values[4])
        Ls_rule_name.append('Hierarchy')
        Ls_ref.append(df_filter_hierarchy.loc[ind_max].values[3])

        
    if df_filter_term_hie.empty == False:    
        ind_max = df_filter_term_hie['Percentage'].idxmax()
               
        Ls_percent.append(df_filter_term_hie['Percentage'].max())
        Ls_num_follow.append(df_filter_term_hie.loc[ind_max].values[4])
        Ls_rule_name.append('Termcode-Hierarchy')
        Ls_ref.append(df_filter_term_hie.loc[ind_max].values[3])

        
    if df_filter_term_numRange.empty == False: 
        ind_max = df_filter_term_numRange['Percentage'].idxmax()
               
        Ls_percent.append(df_filter_term_numRange['Percentage'].max())
        Ls_num_follow.append(df_filter_term_numRange.loc[ind_max].values[4])
        Ls_rule_name.append('Termcode-NumRange')
        Ls_ref.append(df_filter_term_numRange.loc[ind_max].values[3])

        
    if df_filter_hie_numRange.empty == False: 
        ind_max = df_filter_hie_numRange['Percentage'].idxmax()
               
        Ls_percent.append(df_filter_hie_numRange['Percentage'].max())
        Ls_num_follow.append(df_filter_hie_numRange.loc[ind_max].values[4])
        Ls_rule_name.append('Hierarchy-NumRange')
        Ls_ref.append(df_filter_hie_numRange.loc[ind_max].values[3])
    
    if Ls_ref: 
        ind_max_percent = Ls_percent.index(max(Ls_percent))
        
        get_ref = Ls_ref[ind_max_percent]  
        percentage_follow_ref = max(Ls_percent)
        amount_follower = Ls_num_follow[ind_max_percent]
        rule_type = Ls_rule_name[ind_max_percent]
    
    return get_ref, percentage_follow_ref, amount_follower, rule_type

#*********************************************************************************************************************************

if option.lower().find('y_pt') > -1:
       
    df_hint_PT = df_hint[df_hint['Logical Group'] == 'Y_PT']        
    
    sql_rules_based = " SELECT * FROM [DB_CTXFC1_SQL].[dbo].[RULE_PT_BR_TT] WHERE [Logical System Group] = 'Y_PT'"
    sql_rules_statistics = " SELECT * FROM [DB_CTXFC1_SQL].[dbo].[RULE_STATISTICS] WHERE [Logical System Group] = 'Y_PT'"
    
    df_rule = base.pd.read_sql(sql_rules_based, conn)
    df_stats = base.pd.read_sql(sql_rules_statistics, conn)
        
    lastrow = ws.Cells(ws.Cells.Rows.Count, "C").End(-4162).Row

    j = 2
    for j in range(j, lastrow + 1):
     
        termcode = str(ws.Cells(j,'I').Value).strip()
        if termcode.find('.0') > -1: termcode = termcode[:len(termcode) - 2]
        number_range = str(ws.Cells(j,'C').Value)[0:4]
        pro_hierarchy = str(ws.Cells(j,'K').Value)
        pro_hierarchy = pro_hierarchy.strip()
        
        compare_str = str(ws.Cells(j,'Q').Value) + ' ' +  str(ws.Cells(j,'D').Value)
        drawing = str(ws.Cells(j,'H').Value).replace(".", "")
        
        #Run auto-classification by looking up text data      
        product_group, comment_by_BOT = Classify_By_Text_Data_PT(termcode,drawing,compare_str,df_hint_PT)
        need_check = ''
        percentage_follow_ref,amount_of_part = 0,0
        #Classify by rulling if Classify by Configurable Material fails_____________________________________ 
        if  product_group == "": 
            product_group,\
            need_check,\
            comment_by_BOT,\
            percentage_follow_ref,\
            amount_of_part = Classify_By_Statistics_Rulling(termcode,number_range,pro_hierarchy,df_rule,df_stats)
                  
        ws.Cells(j,'T').Value = product_group
        ws.Cells(j,'U').Value = need_check
        ws.Cells(j,'N').Value = comment_by_BOT
        ws.Cells(j,'R').Value = percentage_follow_ref
        ws.Cells(j,'S').Value = amount_of_part
        
        
        str_ = "\r Line {0}: - {1}".format(str(j),str(product_group))
        print_and_overwrite(str_)

    
    
elif option.lower().find('y_br') > -1:   
    
    df_hint_DC = df_hint[df_hint['Logical Group'] == 'Y_BR']
    
    sql_rules_based = """ SELECT * FROM [DB_CTXFC1_SQL].[dbo].[RULE_PT_BR_TT] WHERE [Logical System Group] = 'Y_BR' """
    sql_rules_statistics = " SELECT * FROM [DB_CTXFC1_SQL].[dbo].[RULE_STATISTICS] WHERE [Logical System Group] = 'Y_BR'"
        
    df_rule = base.pd.read_sql(sql_rules_based, conn)
    df_stats = base.pd.read_sql(sql_rules_statistics, conn)
    
    lastrow = ws.Cells(ws.Cells.Rows.Count, "C").End(-4162).Row

    j = 2
    for j in range(j, lastrow + 1):
        
        termcode = str(ws.Cells(j,'I').Value).strip()
        number_range = str(ws.Cells(j,'C').Value)[0:4]
        pro_hierarchy = str(ws.Cells(j,'K').Value).strip()
        
        ZKMA = str(ws.Cells(j,'M').Value).strip()
        motor_code = str(ws.Cells(j,'F').Value)
        weight =  str(ws.Cells(j,'G').Value)
                
        #Classify by Configurable Material_____________________________________
        if len(ZKMA) > 5: #'None' string type has 4 digit  
            if motor_code.find("Without motor") > -1 or len(motor_code) < 5 :
                product_group, comment_by_BOT = Classify_By_Text_Data_DC(str(ZKMA).strip(),"",weight,df_hint_DC)            
            else:
                product_group, comment_by_BOT = Classify_By_Text_Data_DC("",motor_code,weight,df_hint_DC)
                
            need_check = ''
            percentage_follow_ref,amount_of_part = 0,0
        #Classify by rulling if Classify by Configurable Material fails_____________________________________           
        else:          
            product_group,\
            need_check,\
            comment_by_BOT,\
            percentage_follow_ref,\
            amount_of_part = Classify_By_Statistics_Rulling(termcode,number_range,pro_hierarchy,df_rule,df_stats)
                   
        ws.Cells(j,'T').Value = product_group
        ws.Cells(j,'U').Value = need_check
        ws.Cells(j,'N').Value = comment_by_BOT
        ws.Cells(j,'R').Value = percentage_follow_ref
        ws.Cells(j,'S').Value = amount_of_part
        
        str_ = "\r Line {0}: - {1}".format(str(j),str(product_group))
        print_and_overwrite(str_)
    
elif option.lower().find('y_tt01') > -1:   
    
    df_hint_TT = df_hint[df_hint['Logical Group'] == 'Y_TT01']
    
    sql_rules_based = """ SELECT * FROM [DB_CTXFC1_SQL].[dbo].[RULE_PT_BR_TT] WHERE [Logical System Group] = 'Y_TT01' """
    sql_rules_statistics = " SELECT * FROM [DB_CTXFC1_SQL].[dbo].[RULE_STATISTICS] WHERE [Logical System Group] = 'Y_TT01'"
    
    df_rule = base.pd.read_sql(sql_rules_based, conn)
    df_stats = base.pd.read_sql(sql_rules_statistics, conn)
    
    lastrow = ws.Cells(ws.Cells.Rows.Count, "C").End(-4162).Row

    j = 2
    for j in range(j, lastrow + 1):
        
        if  ws.Cells(j,'N').Value == None or ws.Cells(j,'N').Value == '' :
            
            number_range = str(ws.Cells(j,'C').Value)[0:4]
            pro_hierarchy = str(ws.Cells(j,'K').Value).strip()
            short_text = str(ws.Cells(j,'D').Value).strip()
            
            product_group, comment_by_BOT = Classify_By_Text_Data_TT01(pro_hierarchy,short_text,df_hint_TT)
            need_check = ''
            percentage_follow_ref,amount_of_part = 0,0
            #Classify by rulling if Classify by Configurable Material fails_____________________________________ 
            if product_group == "":    
                product_group,\
                need_check,\
                comment_by_BOT,\
                percentage_follow_ref,\
                amount_of_part = Classify_By_Statistics_Rulling("",number_range,pro_hierarchy,df_rule,df_stats)                              
    
            ws.Cells(j,'T').Value = product_group
            ws.Cells(j,'U').Value = need_check
            ws.Cells(j,'N').Value = comment_by_BOT
            ws.Cells(j,'R').Value = percentage_follow_ref
            ws.Cells(j,'S').Value = amount_of_part
        
        str_ = "\r Line {0}: - {1}".format(str(j),str(product_group))
        print_and_overwrite(str_)  
            
elif option.lower().find('y_st') > -1:   
        
    sql_rules_based = """ SELECT * FROM [DB_CTXFC1_SQL].[dbo].[RULE_PT_BR_TT] WHERE [Logical System Group] = 'Y_ST' """
    sql_rules_statistics = " SELECT * FROM [DB_CTXFC1_SQL].[dbo].[RULE_STATISTICS] WHERE [Logical System Group] = 'Y_ST'"
    
    df_rule = base.pd.read_sql(sql_rules_based, conn)
    df_stats = base.pd.read_sql(sql_rules_statistics, conn)
    
    lastrow = ws.Cells(ws.Cells.Rows.Count, "C").End(-4162).Row

    j = 2
    for j in range(j, lastrow + 1):
        
        if  ws.Cells(j,'N').Value == None or ws.Cells(j,'N').Value == '' :
            
            number_range = str(ws.Cells(j,'C').Value)[0:4]
            pro_hierarchy = str(ws.Cells(j,'K').Value).strip()
            short_text = str(ws.Cells(j,'D').Value).strip()
            
            product_group,need_check, comment_by_BOT = '','',''
            percentage_follow_ref,amount_of_part = 0,0
            #Classify by rulling if Classify by Configurable Material fails_____________________________________ 
   
            product_group,\
            need_check,\
            comment_by_BOT,\
            percentage_follow_ref,\
            amount_of_part = Classify_By_Statistics_Rulling("",number_range,pro_hierarchy,df_rule,df_stats)
   
            ws.Cells(j,'T').Value = product_group
            ws.Cells(j,'U').Value = need_check
            ws.Cells(j,'N').Value = comment_by_BOT
            ws.Cells(j,'R').Value = percentage_follow_ref
            ws.Cells(j,'S').Value = amount_of_part
        
        str_ = "\r Line {0}: - {1}".format(str(j),str(product_group))
        print_and_overwrite(str_)
     
else:
    print('Valid option!!!. Abort function')
    exit() 
    
#******************************************************************************   

    
    