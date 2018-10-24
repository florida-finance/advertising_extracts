
# coding: utf-8

# In[1]:


from IPython.core.display import display, HTML
display(HTML("<style>.container { width:100% !important; }</style>"))


# In[2]:


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
from sqlalchemy.sql import text as sa_text
import win32com.client

#Define Connection to Postgres Database
def connect(user, password, db, host, port=5432):
    url  = 'postgresql://{}:{}@{}:{}/{}'
    url = url.format(user, password, host, port, db)

    # The return value of create_engine() is our connection object
    con = create_engine(url, client_encoding='utf8')
    return con


def migrate_records(query, connString, destination_table):
    rs = win32com.client.Dispatch(r'ADODB.Recordset')
    rs.Open(query, connString, CursorType=3)
    a = rs.GetRows()
    dataframe = pd.DataFrame(data= list(a)).transpose()

    dataframe.rename(columns={0:'company_code', 2:'gl_entity', 4:'gl_sub_entity', 6:'gl_product_code', 8:'ledger_account', 10:'order_ad_size', 12:'order_ad_type',
                       14:'parent_name_number', 16:'sales_category', 18:'sales_subcategory', 20:'product_code', 22:'parent_product', 24:'product_type', 
                       26:'fiscal_quarter',28:'fiscal_period',30:'fiscal_week',32:'net'},inplace=True)

    dataframe =dataframe[['company_code','gl_entity','gl_sub_entity','gl_product_code','ledger_account','order_ad_size','order_ad_type',
                       'parent_name_number', 'sales_category','sales_subcategory','product_code','parent_product','product_type', 
                       'fiscal_quarter','fiscal_period','fiscal_week','net']]

    dataframe.to_sql(destination_table, c, if_exists='append', index=False)

#Connection criteria for postgresdb    
hostname = 'mktstrategy.ciklurvi0auw.us-east-1.rds.amazonaws.com'
username = 'tronc'
password = 'tronc123123!'
database = 'Financial_Reporting'

#create connection object
c = connect(username, password, database, hostname)

#Define OLAP cube connection string and destination table to ingest
connString = "PROVIDER=MSOLAP;Data Source={0};Database={1}".format('fcwPsqlanl03','Billing_2')
destination_table = 'billing'

#truncate existing table
c.execute(sa_text('''TRUNCATE TABLE "billing"''').execution_options(autocommit=True))


# In[3]:


connString = "PROVIDER=MSOLAP;Data Source={0};Database={1}".format('fcwPsqlanl03','Billing_2')
destination_table = 'billing'

#2016 - P6
#Pacing Report Cube Pull
query ='SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201601], [Reporting Date].[Fiscal Period].&[201602], [Reporting Date].[Fiscal Period].&[201603], [Reporting Date].[Fiscal Period].&[201604], [Reporting Date].[Fiscal Period].&[201605], [Reporting Date].[Fiscal Period].&[201606] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Status].&[4] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
migrate_records(query, connString, destination_table)


#2016 - P12
#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201612], [Reporting Date].[Fiscal Period].&[201611], [Reporting Date].[Fiscal Period].&[201610], [Reporting Date].[Fiscal Period].&[201608], [Reporting Date].[Fiscal Period].&[201609], [Reporting Date].[Fiscal Period].&[201607] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Status].&[4] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
migrate_records(query, connString, destination_table)


#2017 - P6
#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201702], [Reporting Date].[Fiscal Period].&[201703], [Reporting Date].[Fiscal Period].&[201704], [Reporting Date].[Fiscal Period].&[201705], [Reporting Date].[Fiscal Period].&[201706] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Status].&[4] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
migrate_records(query, connString, destination_table)


#2017 - P12
#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201707], [Reporting Date].[Fiscal Period].&[201708], [Reporting Date].[Fiscal Period].&[201709], [Reporting Date].[Fiscal Period].&[201710], [Reporting Date].[Fiscal Period].&[201711], [Reporting Date].[Fiscal Period].&[201712] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Status].&[4] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
migrate_records(query, connString, destination_table)


#2018 - P6
#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201801], [Reporting Date].[Fiscal Period].&[201802], [Reporting Date].[Fiscal Period].&[201804], [Reporting Date].[Fiscal Period].&[201803], [Reporting Date].[Fiscal Period].&[201805], [Reporting Date].[Fiscal Period].&[201806] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Status].&[4] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
migrate_records(query, connString, destination_table)


#2018 - P12
#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201807], [Reporting Date].[Fiscal Period].&[201808], [Reporting Date].[Fiscal Period].&[201809], [Reporting Date].[Fiscal Period].&[201810], [Reporting Date].[Fiscal Period].&[201811], [Reporting Date].[Fiscal Period].&[201812] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Status].&[4] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
migrate_records(query, connString, destination_table)


#2019 - P6
#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201901], [Reporting Date].[Fiscal Period].&[201902], [Reporting Date].[Fiscal Period].&[201904], [Reporting Date].[Fiscal Period].&[201903], [Reporting Date].[Fiscal Period].&[201905], [Reporting Date].[Fiscal Period].&[201906] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Status].&[4] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
migrate_records(query, connString, destination_table)


#2019 - P12
#Pacing Report Cube Pull
query =' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Reporting Date].[Fiscal Period].&[201907], [Reporting Date].[Fiscal Period].&[201908], [Reporting Date].[Fiscal Period].&[201909], [Reporting Date].[Fiscal Period].&[201910], [Reporting Date].[Fiscal Period].&[201911], [Reporting Date].[Fiscal Period].&[201912] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Status].&[4] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
migrate_records(query, connString, destination_table)


# In[4]:


pd.set_option('display.max_columns', None)  
pd.set_option('display.expand_frame_repr', False)
pd.set_option('max_colwidth', -1)

