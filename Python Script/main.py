# -*- coding: utf-8 -*-
"""
Created on Thu Nov 17 15:49:17 2022

@author: anubhav.anand
"""

import os
import json
from datetime import datetime
import pandas as pd
from sqlalchemy import create_engine
import sys
import win32com.client as win32


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


try:
    #### CREATE A FREEZ BOT KEY FILE
    with open("temp_file_delete.txt", "w") as writer:
        writer.write("Hello World")
        writer.close()

    #### READ CONFIG FILE
    with open(
        "C:/Users/"+os.getlogin()+"/Documents/Common/zigly_analytics_database_config.json"
    ) as config_file:
        data = json.load(config_file)
    #### GET DATA FROM CONFIG FILE
    username = data["username"]
    password = data["password"]
    host = data["host"]
    toaddr = data["toaddr"]

    #### READ CONFIG FILE
    with open(
        "C:/Users/"+os.getlogin()+"/Documents/Common/common_path.json"
    ) as config_file:
        data = json.load(config_file)
    #### GET DATA FROM CONFIG FILE
    downloads_path = data["downloads_path"]

    test_db_connection(username, password, host)
    # =============================================================================
    # Read GRC Inward Detail data
    # =============================================================================
    file_df = pd.DataFrame(columns=["file", "time"])
    file_lst = []
    creation_list = []
    for file in os.listdir(downloads_path):
        if (".xlsx" in file) & ("GRC Inward Detail" in file):
            file_lst.append(file)
            creation_list.append(os.path.getctime(downloads_path + file))
    file_df["file"] = file_lst
    file_df["time"] = creation_list
    file_df = file_df.sort_values(by=["time"], ascending=False)
    COMMON_PATH_NEW = downloads_path + file_df["file"].iloc[0]
    print(COMMON_PATH_NEW)
    grc_data_df = pd.read_excel(COMMON_PATH_NEW)
    grc_data_df.columns = grc_data_df.loc[1:1, :].values.tolist()[0]
    grc_data_df = grc_data_df.iloc[2:-1:, :].reset_index(drop=True)
    grc_data_df["RECEIVE_DATE"] = pd.to_datetime(grc_data_df["RECEIVE_DATE"])
    for cols in grc_data_df.columns.tolist():
        if str(cols) == "nan":
            del grc_data_df[cols]

    # try:
    #     grc_existing_df = read_from_database(
    #         username, password, host, "ginesys", "grc_table"
    #     )
    # except:
    #     grc_existing_df = pd.DataFrame(columns=grc_data_df.columns.tolist())

    # grc_data_df["type"] = "new"
    # grc_existing_df["type"] = "old"
    # grc_no_dup = pd.concat([grc_existing_df, grc_data_df], axis=0).reset_index(
    #     drop=True
    # )
    # col_lst = grc_existing_df.columns.tolist()
    # col_lst.remove("type")
    # grc_no_dup = grc_no_dup.drop_duplicates(subset=col_lst).reset_index(
    #     drop=True
    # )
    # grc_no_dup = grc_no_dup[grc_no_dup["type"] == "new"].reset_index(drop=True)
    # del grc_no_dup["type"]
    if grc_data_df.shape[0] > 0:
        push_to_database(
            username,
            password,
            host,
            "ginesys",
            "grc_table",
            grc_data_df,
            "replace",
        )

    # =============================================================================
    # Read Product Inward Detail data
    # =============================================================================
    file_df = pd.DataFrame(columns=["file", "time"])
    file_lst = []
    creation_list = []
    for file in os.listdir(downloads_path):
        if (".xlsx" in file) & ("Product Movement" in file):
            file_lst.append(file)
            creation_list.append(os.path.getctime(downloads_path + file))
    file_df["file"] = file_lst
    file_df["time"] = creation_list
    file_df = file_df.sort_values(by=["time"], ascending=False)
    COMMON_PATH_NEW = downloads_path + file_df["file"].iloc[0]
    print(COMMON_PATH_NEW)

    product_data_df = pd.read_excel(COMMON_PATH_NEW)
    product_data_df.columns = product_data_df.loc[4:4, :].values.tolist()[0]
    product_data_df = product_data_df.iloc[5::, :].reset_index(drop=True)
    product_data_df["SENT_DOCUMENT_DATE"] = pd.to_datetime(
        product_data_df["SENT_DOCUMENT_DATE"]
    )
    # for cols in product_data_df.columns.tolist():
    #     if str(cols) == "nan":
    #         del product_data_df[cols]

    # try:
    #     product_existing_df = read_from_database(
    #         username, password, host, "ginesys", "product_movement_table"
    #     )
    # except:
    #     product_existing_df = pd.DataFrame(
    #         columns=product_data_df.columns.tolist()
    #     )

    # product_data_df["type"] = "new"
    # product_existing_df["type"] = "old"
    # product_no_dup = pd.concat(
    #     [product_existing_df, product_data_df], axis=0
    # ).reset_index(drop=True)
    # col_lst = product_existing_df.columns.tolist()
    # col_lst.remove("type")
    # product_no_dup = product_no_dup.drop_duplicates(
    #     subset=col_lst
    # ).reset_index(drop=True)
    # product_no_dup = product_no_dup[
    #     product_no_dup["type"] == "new"
    # ].reset_index(drop=True)
    # del product_no_dup["type"]
    if product_data_df.shape[0] > 0:
        push_to_database(
            username,
            password,
            host,
            "ginesys",
            "product_movement_table",
            product_data_df,
            "replace",
        )

    # =============================================================================
    # Stock at Point Data
    # =============================================================================
    file_df = pd.DataFrame(columns=["file", "time"])
    file_lst = []
    creation_list = []
    for file in os.listdir(downloads_path):
        if (".xlsx" in file) & ("Stock at Point" in file):
            file_lst.append(file)
            creation_list.append(os.path.getctime(downloads_path + file))
    file_df["file"] = file_lst
    file_df["time"] = creation_list
    file_df = file_df.sort_values(by=["time"], ascending=False)
    COMMON_PATH_NEW = downloads_path + file_df["file"].iloc[0]
    print(COMMON_PATH_NEW)

    stock_at_point_data_df = pd.read_excel(COMMON_PATH_NEW)
    stock_at_point_data_df.columns = stock_at_point_data_df.loc[
        1:1, :
    ].values.tolist()[0]
    stock_at_point_data_df = stock_at_point_data_df.iloc[2:-1:, :].reset_index(
        drop=True
    )
    
    for cols in stock_at_point_data_df.columns.tolist():
        if str(cols) == "nan":
            del stock_at_point_data_df[cols]

    stock_at_point_data_df["End_Point_Date"] = datetime.today().date()

    if stock_at_point_data_df.shape[0] > 0:
        push_to_database(
            username,
            password,
            host,
            "ginesys",
            "stock_at_point_movement_table",
            stock_at_point_data_df,
            "append",
        )

    #### REMOVE FREEZE BOT KEY FILE
    os.remove("temp_file_delete.txt")

    #### SEND MAIL COMMUNICATION
    temp_df = pd.DataFrame(
        [
            [
                "Ginesys Warehouse Data Download Bot",
                datetime.today(),
                "Success",
                "",
            ]
        ],
        columns=["Usecase", "Exec_DateTime", "Status", "Error"],
    )
    push_to_database(
        username, password, host, "common", "usecase_status", temp_df, "append"
    )
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = toaddr
    mail.Subject = "Ginesys Warehouse Data Download - Code Successful"
    mail.Body = "Hi Everyone,\n\n Execution successful."

    # mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional
    # To attach a file to the email (optional):
    # attachment  = os.getcwd()+"\\"+CONVERTED_DATA
    # mail.Attachments.Add(attachment)

    mail.Send()
except:
    try:
        os.remove("temp_file_delete.txt")
    except:
        pass
    error_cls, error, lineno = sys.exc_info()
    error_str = (
        "Class:"
        + str(error_cls)
        + "    Error:"
        + str(error)
        + "    Lineno:"
        + str(lineno.tb_lineno)
    )
    temp_df = pd.DataFrame(
        [
            [
                "Ginesys Warehouse Data Download Bot",
                datetime.today(),
                "Failed",
                error_str,
            ]
        ],
        columns=["Usecase", "Exec_DateTime", "Status", "Error"],
    )
    push_to_database(
        username, password, host, "common", "usecase_status", temp_df, "append"
    )
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = toaddr
    mail.Subject = "Ginesys Warehouse Data Download - Code Fail"
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
