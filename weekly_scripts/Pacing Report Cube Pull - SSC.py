
# coding: utf-8

# In[1]:


from io import BytesIO
import json, pycurl
import urllib3
import requests
import pandas as pd
http = urllib3.PoolManager()
from datetime import datetime, timedelta, date
import pandas.io.sql as psql
import pandasql
from sqlalchemy import create_engine
from sqlalchemy import MetaData
import win32com.client
connString = "PROVIDER=MSOLAP;Data Source={0};Database={1};Connect Timeout=3000".format('fcwPsqlanl03','Billing_2')
def ado(strsql, connString):
    rs = win32com.client.Dispatch(r'ADODB.Recordset')
    rs.Open(strsql, connString)
    rs.CommandTimeout=6000
    t = rs.GetRows()
    rs.Close()
    return t

#Define Connection to Postgres Database
hostname = 'mktstrategy.ciklurvi0auw.us-east-1.rds.amazonaws.com'
username = 'tronc'
password = 'tronc123123!'
database = 'Financial_Reporting'
port=5432

def connect(user, password, db, host, port=5432):
    url  = 'postgresql://{}:{}@{}:{}/{}'
    url = url.format(user, password, host, port, db)

    # The return value of create_engine() is our connection object
    con = create_engine(url, client_encoding='utf8')
    return con
c =connect(username, password, database, hostname)


# In[ ]:


#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS *  [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS * [Product].[GL Product].[GL Product].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201601], [Reporting Date].[Fiscal Period].&[201603], [Reporting Date].[Fiscal Period].&[201602], [Reporting Date].[Fiscal Period].&[201604], [Reporting Date].[Fiscal Period].&[201605], [Reporting Date].[Fiscal Period].&[201606] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company].&[SSC] } ) ON COLUMNS FROM [Revenue])))) WHERE ( [Company].[Company].&[SSC] ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
rs = win32com.client.Dispatch(r'ADODB.Recordset')
rs.Open(query, connString, CursorType=3)
a = rs.GetRows()
df = pd.DataFrame(data= list(a)).transpose()

df.rename(columns={0:'parent_name_number',2:'ledger_account',4:'sales_category',8:'sales_subcategory',10:'parent_product',
                   12:'product_type',14:'fiscal_quarter',16:'fiscal_period',18:'fiscal_week',20:'gl_product', 22:'net', 6:'product_code'},inplace=True)

df =df[['parent_name_number','ledger_account','sales_category','sales_subcategory','parent_product',
         'product_type','fiscal_quarter','fiscal_period','fiscal_week','net','gl_product', 'product_code']]


from sqlalchemy.sql import text as sa_text
c.execute(sa_text('''TRUNCATE TABLE "pacing_cube_pull_ssc"''').execution_options(autocommit=True))
df.to_sql('pacing_cube_pull_ssc', c, if_exists='append', index=False)




#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS * [Product].[GL Product].[GL Product].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201607], [Reporting Date].[Fiscal Period].&[201608], [Reporting Date].[Fiscal Period].&[201609], [Reporting Date].[Fiscal Period].&[201610], [Reporting Date].[Fiscal Period].&[201611], [Reporting Date].[Fiscal Period].&[201612] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company].&[SSC] } ) ON COLUMNS FROM [Revenue])))) WHERE ( [Company].[Company].&[SSC] ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
rs = win32com.client.Dispatch(r'ADODB.Recordset')
rs.Open(query, connString, CursorType=3)
a = rs.GetRows()
df = pd.DataFrame(data= list(a)).transpose()

df.rename(columns={0:'parent_name_number',2:'ledger_account',4:'sales_category',8:'sales_subcategory',10:'parent_product',
                   12:'product_type',14:'fiscal_quarter',16:'fiscal_period',18:'fiscal_week',20:'gl_product', 22:'net', 6:'product_code'},inplace=True)

df =df[['parent_name_number','ledger_account','sales_category','sales_subcategory','parent_product',
         'product_type','fiscal_quarter','fiscal_period','fiscal_week','net','gl_product', 'product_code']]


df.to_sql('pacing_cube_pull_ssc', c, if_exists='append', index=False)



#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS * [Product].[GL Product].[GL Product].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201701], [Reporting Date].[Fiscal Period].&[201702], [Reporting Date].[Fiscal Period].&[201703], [Reporting Date].[Fiscal Period].&[201704], [Reporting Date].[Fiscal Period].&[201705], [Reporting Date].[Fiscal Period].&[201706] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company].&[SSC] } ) ON COLUMNS FROM [Revenue])))) WHERE ( [Company].[Company].&[SSC] ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
rs = win32com.client.Dispatch(r'ADODB.Recordset')
rs.Open(query, connString, CursorType=3)
a = rs.GetRows()
df = pd.DataFrame(data= list(a)).transpose()

df.rename(columns={0:'parent_name_number',2:'ledger_account',4:'sales_category',8:'sales_subcategory',10:'parent_product',
                   12:'product_type',14:'fiscal_quarter',16:'fiscal_period',18:'fiscal_week',20:'gl_product', 22:'net', 6:'product_code'},inplace=True)

df =df[['parent_name_number','ledger_account','sales_category','sales_subcategory','parent_product',
         'product_type','fiscal_quarter','fiscal_period','fiscal_week','net','gl_product', 'product_code']]

df.to_sql('pacing_cube_pull_ssc', c, if_exists='append', index=False)



#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS * [Product].[GL Product].[GL Product].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201707], [Reporting Date].[Fiscal Period].&[201708], [Reporting Date].[Fiscal Period].&[201709], [Reporting Date].[Fiscal Period].&[201710], [Reporting Date].[Fiscal Period].&[201711], [Reporting Date].[Fiscal Period].&[201712] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company].&[SSC] } ) ON COLUMNS FROM [Revenue])))) WHERE ( [Company].[Company].&[SSC] ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
rs = win32com.client.Dispatch(r'ADODB.Recordset')
rs.Open(query, connString, CursorType=3)
a = rs.GetRows()
df = pd.DataFrame(data= list(a)).transpose()

df.rename(columns={0:'parent_name_number',2:'ledger_account',4:'sales_category',8:'sales_subcategory',10:'parent_product',
                   12:'product_type',14:'fiscal_quarter',16:'fiscal_period',18:'fiscal_week',20:'gl_product', 22:'net', 6:'product_code'},inplace=True)

df =df[['parent_name_number','ledger_account','sales_category','sales_subcategory','parent_product',
         'product_type','fiscal_quarter','fiscal_period','fiscal_week','net','gl_product', 'product_code']]


df.to_sql('pacing_cube_pull_ssc', c, if_exists='append', index=False)



#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS * [Product].[GL Product].[GL Product].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201801], [Reporting Date].[Fiscal Period].&[201802], [Reporting Date].[Fiscal Period].&[201803], [Reporting Date].[Fiscal Period].&[201804], [Reporting Date].[Fiscal Period].&[201805], [Reporting Date].[Fiscal Period].&[201806] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company].&[SSC] } ) ON COLUMNS FROM [Revenue])))) WHERE ( [Company].[Company].&[SSC] ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
rs = win32com.client.Dispatch(r'ADODB.Recordset')
rs.Open(query, connString, CursorType=3)
a = rs.GetRows()
df = pd.DataFrame(data= list(a)).transpose()

df.rename(columns={0:'parent_name_number',2:'ledger_account',4:'sales_category',8:'sales_subcategory',10:'parent_product',
                   12:'product_type',14:'fiscal_quarter',16:'fiscal_period',18:'fiscal_week',20:'gl_product', 22:'net', 6:'product_code'},inplace=True)

df =df[['parent_name_number','ledger_account','sales_category','sales_subcategory','parent_product',
         'product_type','fiscal_quarter','fiscal_period','fiscal_week','net','gl_product', 'product_code']]


df.to_sql('pacing_cube_pull_ssc', c, if_exists='append', index=False)




#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS * [Product].[GL Product].[GL Product].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201812], [Reporting Date].[Fiscal Period].&[201811], [Reporting Date].[Fiscal Period].&[201809], [Reporting Date].[Fiscal Period].&[201810], [Reporting Date].[Fiscal Period].&[201808], [Reporting Date].[Fiscal Period].&[201807] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company].&[SSC] } ) ON COLUMNS FROM [Revenue])))) WHERE ( [Company].[Company].&[SSC] ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
rs = win32com.client.Dispatch(r'ADODB.Recordset')
rs.Open(query, connString, CursorType=3)
a = rs.GetRows()
df = pd.DataFrame(data= list(a)).transpose()

df.rename(columns={0:'parent_name_number',2:'ledger_account',4:'sales_category',8:'sales_subcategory',10:'parent_product',
                   12:'product_type',14:'fiscal_quarter',16:'fiscal_period',18:'fiscal_week',20:'gl_product', 22:'net', 6:'product_code'},inplace=True)

df =df[['parent_name_number','ledger_account','sales_category','sales_subcategory','parent_product',
         'product_type','fiscal_quarter','fiscal_period','fiscal_week','net','gl_product', 'product_code']]


df.to_sql('pacing_cube_pull_ssc', c, if_exists='append', index=False)









#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS * [Product].[GL Product].[GL Product].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201901], [Reporting Date].[Fiscal Period].&[201902], [Reporting Date].[Fiscal Period].&[201903], [Reporting Date].[Fiscal Period].&[201904], [Reporting Date].[Fiscal Period].&[201905], [Reporting Date].[Fiscal Period].&[201906] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company].&[SSC] } ) ON COLUMNS FROM [Revenue])))) WHERE ( [Company].[Company].&[SSC] ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
rs = win32com.client.Dispatch(r'ADODB.Recordset')
rs.Open(query, connString, CursorType=3)
a = rs.GetRows()
df = pd.DataFrame(data= list(a)).transpose()

df.rename(columns={0:'parent_name_number',2:'ledger_account',4:'sales_category',8:'sales_subcategory',10:'parent_product',
                   12:'product_type',14:'fiscal_quarter',16:'fiscal_period',18:'fiscal_week',20:'gl_product', 22:'net', 6:'product_code'},inplace=True)

df =df[['parent_name_number','ledger_account','sales_category','sales_subcategory','parent_product',
         'product_type','fiscal_quarter','fiscal_period','fiscal_week','net','gl_product', 'product_code']]


df.to_sql('pacing_cube_pull_ssc', c, if_exists='append', index=False)




#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS * [Product].[GL Product].[GL Product].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201912], [Reporting Date].[Fiscal Period].&[201911], [Reporting Date].[Fiscal Period].&[201909], [Reporting Date].[Fiscal Period].&[201910], [Reporting Date].[Fiscal Period].&[201908], [Reporting Date].[Fiscal Period].&[201907] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company].&[SSC] } ) ON COLUMNS FROM [Revenue])))) WHERE ( [Company].[Company].&[SSC] ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
rs = win32com.client.Dispatch(r'ADODB.Recordset')
rs.Open(query, connString, CursorType=3)
a = rs.GetRows()
df = pd.DataFrame(data= list(a)).transpose()

df.rename(columns={0:'parent_name_number',2:'ledger_account',4:'sales_category',8:'sales_subcategory',10:'parent_product',
                   12:'product_type',14:'fiscal_quarter',16:'fiscal_period',18:'fiscal_week',20:'gl_product', 22:'net', 6:'product_code'},inplace=True)

df =df[['parent_name_number','ledger_account','sales_category','sales_subcategory','parent_product',
         'product_type','fiscal_quarter','fiscal_period','fiscal_week','net','gl_product', 'product_code']]


df.to_sql('pacing_cube_pull_ssc', c, if_exists='append', index=False)


