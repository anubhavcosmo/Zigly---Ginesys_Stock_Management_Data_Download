# -*- coding: utf-8 -*-
"""
Created on Thu Nov 17 17:31:00 2022

@author: anubhav.anand
"""

import json
from sqlalchemy import create_engine
from datetime import datetime
import pandas as pd
import win32com.client as win32
import os

#### READ CONFIG FILE
with open(
    "D:/Users/"+os.getlogin()+"/Documents/Common/zigly_analytics_database_config.json"
) as config_file:
    data = json.load(config_file)
#### GET DATA FROM CONFIG FILE
username = data["username"]
password = data["password"]
host = data["host"]
toaddr = data["toaddr"]


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
        + 'common'
    )
    temp_con = create_engine(temp_con)
    try:
        pd.read_sql("select * from " + str('usecase_status LIMIT 1') + ";", con=temp_con)
    except Exception as excp:
        raise Exception(str(excp)+"\t\t No Connection to database")
test_db_connection(username, password, host)



temp_df = pd.DataFrame(
    [["Ginesys Warehouse Data Download Bot", datetime.today(), "Failed", "Bot Execution Failed"]],
    columns=["Usecase", "Exec_DateTime", "Status", "Error"],
)
push_to_database(
    username, password, host, "common", "usecase_status", temp_df, "append"
)
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = toaddr
mail.Subject = "Ginesys Warehouse Data Download - Bot Fail"
mail.Body = (
    "Hi Everyone,\n\n Bot Fail"
)

# mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
# To attach a file to the email (optional):
# attachment  = os.getcwd()+"\\"+CONVERTED_DATA
# mail.Attachments.Add(attachment)

mail.Send()