# -*- coding: utf-8 -*-
"""
Created on Fri Apr 26 19:26:13 2019

@author: 19073
"""

import cx_Oracle
import xlwt
import os
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'


def index_application(module_name):
    conn = cx_Oracle.connect('c##LKX/0000@219.216.69.63:1521/orcl')  # 连接数据库
    cur = conn.cursor()
    cur2 = conn.cursor()
    a=0
    wbk = xlwt.Workbook()
    sheet = wbk.add_sheet('Sheet1',cell_overwrite_ok=True) 
    index_list = {}
    module_pre = module_name
     #获取cursor
    cur.execute("SELECT A.table_name,A.num_rows FROM ALL_TABLES A where A.table_name like '" + module_pre + "%' AND OWNER='C##SCYW' ")
    num_index_table=0#满足条件的表数
    for result in cur:
        num_index_table=num_index_table+1
        index_list[result[0]]={}
    print(index_list)
    cur2.execute("SELECT A.TABLE_NAME,A.index_name,A.NUM_ROWS FROM ALL_INDEXES A,ALL_IND_COLUMNS B,(SELECT C.table_name FROM ALL_TABLES C where C.table_name like '" + module_pre + "%' AND OWNER='C##SCYW' ) C where A.INDEX_NAME=B.INDEX_NAME AND A.TABLE_NAME=C.TABLE_NAME")
    num_index=0#满足条件的索引数
    for result in cur2:
        num_index=num_index+1
        index_list[result[0]][result[1]]=result[2]
    print(index_list)  
    print(num_index_table)
    if num_index_table != 0:
       print(num_index/num_index_table)#索引应用程度
    cur.close()  # 关闭cursor
    conn.close()  # 关闭连接  
    
    s= num_index/num_index_table
    w= num_index_table
    i= num_index
    sheet.write(0,0,'模块名')
    sheet.write(0,1,'索引数')
    sheet.write(0,2,'表数')
    sheet.write(0,3,'索引应用程度')
    a=a+1
    sheet.write(a,1,i)
    sheet.write(a,2,w)
    sheet.write(a,3,s)
    #sheet.write(a,1,s)
    wbk.save('ISC_USER0索引应用程度.xls')

index_application('ISC_USER')