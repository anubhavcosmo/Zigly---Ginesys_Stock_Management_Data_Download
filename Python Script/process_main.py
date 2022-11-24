# -*- coding: utf-8 -*-
"""
Created on Tue Nov 22 10:53:32 2022

@author: anubhav.anand
"""


import pandas as pd
from datetime import datetime
from datetime import timedelta
import warnings
import fuzzy
import json
from sqlalchemy import create_engine
import sys
import win32com.client as win32
import os
import matplotlib.pyplot as plt
import numpy as np

##################################2]:


soundex = fuzzy.Soundex(4)
warnings.filterwarnings(action="ignore")

#### READ CONFIG FILE
with open(
    "C:/Users/"+os.getlogin()+"/Documents/Common/zigly_analytics_database_config.json"
) as config_file:
    data = json.load(config_file)
#### GET DATA FROM CONFIG FILE
username = data["username"]
password = data["password"]
host = data["host"]
##################################
conversion_threshold_days = 3


def push_to_database(
    username, password, host, db_name, table_name, dataframe, operation_type
):
    temp_con = (
        "mysql+pymysql://"
        + username
        + ":"
        + password
        + "@"
        + host
        + "/"
        + db_name
    )
    temp_con = create_engine(temp_con)
    return dataframe.to_sql(
        table_name, if_exists=operation_type, index=False, con=temp_con
    )


def read_from_database(username, password, host, db_name, table_name):
    temp_con = (
        "mysql+pymysql://"
        + username
        + ":"
        + password
        + "@"
        + host
        + "/"
        + db_name
    )
    temp_con = create_engine(temp_con)
    return pd.read_sql("select * from " + str(table_name) + ";", con=temp_con)


def test_db_connection(username, password, host):
    temp_con = (
        "mysql+pymysql://"
        + username
        + ":"
        + password
        + "@"
        + host
        + "/"
        + "common"
    )
    temp_con = create_engine(temp_con)
    try:
        pd.read_sql(
            "select * from " + str("usecase_status LIMIT 1") + ";",
            con=temp_con,
        )
    except Exception as excp:
        raise Exception(str(excp) + "\t\t No Connection to database")


def find_stock_movement(stock_point, stock_flow_in, stock_flow_out):
    # stock_point = stock_data_df.copy()
    # stock_flow_in = stock_in_flow_df.copy()
    # stock_flow_out = stock_out_flow_df.copy()
    warehouse_stock_point_name_delete = ['Kailash Colony - EC',
                                         'WMS Stock point',
                                         'Fixed Asset',
                                         'Fixed Asset HO'
                                         'BT UAT',
                                         'BT UAT Return',
                                         'Expired Stock',
                                         'Damage/Rejected']
    store_stock_point_name_delete = ['Fixed Asset',
                                     'Fixed Assets',
                                     'Logistic']
    store_grp = stock_point.groupby(["NAME",'CODE'])
    for key, val in store_grp:
        temp_df = store_grp.get_group(key).sort_values(by = ['Stock_Point_Date']).reset_index(drop = True)
        store_name = key[0] #Get Store Name
        sku_name = key[1] # Get SKU (Item Code)
        
        # #Remove STOCK_POINT_NAMES from Warehouse
        # if("Cosmo First Limit".lower() in store_name.lower()):
        #     temp_df = temp_df[~temp_df['STOCKPOINT_NAME'].isin(warehouse_stock_point_name_delete)].reset_index(drop = True)
        # #Remove STOCK_POINT_NAMES from Store
        # if("Zigly Pet".lower() in store_name.lower()):
        #     temp_df = temp_df[~temp_df['STOCKPOINT_NAME'].isin(store_stock_point_name_delete)].reset_index(drop = True)
        # temp_stock_in = stock_flow_in[(stock_flow_in['SITE']==store_name) & (stock_flow_in['ICODE']==sku_name)].reset_index(drop = True)
        # temp_stock_out = stock_flow_out[(stock_flow_out['Source Location']==store_name) & (stock_flow_out['Item Code']==sku_name)].reset_index(drop = True)
        
        # if(temp_df.shape[0]>0):
        #     # Get Monthwise Data
        #     month_end_date_lst = sorted(temp_df['Stock_Point_Date'].unique().tolist())
            
        #     stock_point_sum_lst = []
        #     stock_in_sum_lst = []
        #     stock_out_sum_lst = []
        #     for i in range(0, len(month_end_date_lst)):
        #         month_start_date_str = datetime.strftime(month_end_date_lst[i],'%Y-%m-%d')
        #         month_start_date_str = month_start_date_str[:-2:]+"01"
        #         month_start_date = datetime.strptime(month_start_date_str,"%Y-%m-%d")
        #         stock_point_sum = temp_df[temp_df['Stock_Point_Date']==month_end_date_lst[i]]['CLOSING_STOCK_QTY'].astype('int').sum()
        #         stock_point_sum_lst.append(abs(stock_point_sum))
        #         stock_in_sum = temp_stock_in[(temp_stock_in['RECEIVE_DATE']>=pd.to_datetime(month_start_date,format = '%Y-%m-%d %H:%M:%S'))&
        #                                      (temp_stock_in['RECEIVE_DATE']<=pd.to_datetime(month_end_date_lst[i],format = "%Y-%m-%d"))]['Received Qty'].astype('int').sum()
        #         stock_in_sum_lst.append(abs(stock_in_sum))
        #         stock_out_sum = temp_stock_out[(temp_stock_out['SENT_DOCUMENT_DATE']>=pd.to_datetime(month_start_date,format = '%Y-%m-%d %H:%M:%S'))&
        #                                      (temp_stock_out['SENT_DOCUMENT_DATE']<=pd.to_datetime(month_end_date_lst[i],format = "%Y-%m-%d"))]['SEND_QTY'].astype('int').sum()
        #         stock_out_sum_lst.append(abs(stock_out_sum))
                
        #     # create data
        #     x = np.arange(len(month_end_date_lst))
        #     y1 = stock_in_sum_lst
        #     y2 = stock_out_sum_lst
        #     y3 = stock_point_sum_lst
        #     width = 0.2
              
        #     y_lim = max(max(stock_point_sum_lst),max(stock_in_sum_lst),max(stock_out_sum_lst))
        #     # plot data in grouped manner of bar type
        #     plt.bar(x-0.2, y1, width, color='cyan')
        #     plt.bar(x, y2, width, color='orange')
        #     plt.bar(x+0.2, y3, width, color='green')
        #     plt.xticks(x, month_end_date_lst)
        #     plt.yticks([i for i in range(0, y_lim,1)])
        #     plt.xlabel("Date")
        #     plt.ylabel("Quantity")
        #     plt.legend(["Stock In", "Current Out", "Stock Current"])
        #     plt.xticks(rotation=90)
        #     plt.show()
                
                
        
        if((temp_df.shape[0]>5) & (temp_stock_in.shape[0]>0) & ("Zigly Pet".lower() in store_name)):
            break


try:
    # check database connectivity
    test_db_connection(username, password, host)
    
    # Read current stock data
    stock_data_df = read_from_database(username, password, host, 'ginesys', 'stock_at_point_table')

    # Read Stock - In flow
    stock_in_flow_df = read_from_database(username, password, host, 'ginesys', 'grc_table')
    
    # Read Stock - Out Flow
    stock_out_flow_df = read_from_database(username, password, host, 'ginesys', 'product_movement_table')
    
    # 
except:
    error_cls, error, lineno = sys.exc_info()
    error_str = (
        "Class:"
        + str(error_cls)
        + "    Error:"
        + str(error)
        + "    Lineno:"
        + str(lineno.tb_lineno)
    )
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = "shashi.bhushan@cosmofirst.com;anubhav.anand@cosmofirst.com"
    mail.Subject = "Warehouse Forecast - Code Fail"
    mail.Body = (
        "Hi Everyone,\n\n Class:"
        + str(error_cls)
        + "    Error:"
        + str(error)
        + "    Lineno:"
        + str(lineno.tb_lineno)
    )

    # mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
    # To attach a file to the email (optional):
    # attachment  = os.getcwd()+"\\"+CONVERTED_DATA
    # mail.Attachments.Add(attachment)

    mail.Send()