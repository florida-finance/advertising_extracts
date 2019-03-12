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
from datetime import datetime

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
    
    dataframe.loc[dataframe['gl_product_code']=='FHOMEZ', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='OS0000', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='PLAKE9', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='PORG99', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='PSEM99', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='PVOL99', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='TMCSSL', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='TMCUSX', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='TMCWRP', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='X21000', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='X22000', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='X30000', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='X45000', 'gl_product_code'] = '000000'  
    dataframe.loc[dataframe['gl_product_code']=='X10000', 'gl_product_code'] = 'DAO000'
    dataframe.loc[dataframe['gl_product_code']=='BRNDCT', 'gl_product_code'] = 'DBC000'
    dataframe.loc[dataframe['gl_product_code']=='X20000', 'gl_product_code'] = 'DCR000'
    dataframe.loc[dataframe['gl_product_code']=='LCLEDG', 'gl_product_code'] = 'DLP000'
    dataframe.loc[dataframe['gl_product_code']=='X55000', 'gl_product_code'] = 'DMS000'
    dataframe.loc[dataframe['gl_product_code']=='RCHEXT', 'gl_product_code'] = 'DRC000'
    dataframe.loc[dataframe['gl_product_code']=='ZEVENT', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='OSIGMG', 'gl_product_code'] = 'OSM000'
    dataframe.loc[dataframe['gl_product_code']=='POSCEO', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code'].isnull(), 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='RELSFR', 'gl_product_code'] = 'SEZ710'
    dataframe.loc[dataframe['gl_product_code']=='RHOMIM', 'gl_product_code'] = 'SHI000'
    dataframe.loc[dataframe['gl_product_code']=='Undefined', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='0', 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']==0, 'gl_product_code'] = '000000'
    dataframe.loc[dataframe['gl_product_code']=='EVT000', 'gl_product_code'] = '000000'
    dataframe.loc[(dataframe['gl_product_code']=='DA O000'),'gl_entity']= '50000'
    dataframe.loc[(dataframe['gl_product_code']=='DBC000'),'gl_entity']= '50000'
    dataframe.loc[(dataframe['gl_product_code']=='DCR000'),'gl_entity']= '50000'  
    dataframe.loc[(dataframe['gl_product_code']=='DLP000'),'gl_entity']= '50000'
    dataframe.loc[(dataframe['gl_product_code']=='DMS000'),'gl_entity']= '50000'
    dataframe.loc[(dataframe['gl_product_code']=='DRC000')&(dataframe['ledger_account']=='407000'),'gl_entity']= '50000'
    dataframe.loc[(dataframe['gl_product_code']=='DRC000')&(dataframe['ledger_account']=='407200'),'gl_entity']= '50000'
    dataframe.loc[(dataframe['gl_product_code']=='DRC000')&(dataframe['ledger_account']=='407420'),'gl_entity']= '50000'
    
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

now = datetime.now().strftime('%Y')
last_year = int(datetime.now().strftime('%Y'))-1
last_year_2 = int(datetime.now().strftime('%Y'))-2

print('Pulling First 6 Months - CY')
#first 6 months current year
query ='''SELECT NON EMPTY {{ [Measures].[Net] }} ON COLUMNS, NON EMPTY {{([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) }} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( {{ [Reporting Date].[Fiscal Period].&[{0}01], [Reporting Date].[Fiscal Period].&[{1}02], [Reporting Date].[Fiscal Period].&[{2}03], [Reporting Date].[Fiscal Period].&[{3}04], [Reporting Date].[Fiscal Period].&[{4}05], [Reporting Date].[Fiscal Period].&[{5}06] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Type].&[101] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Status].&[4] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Order Kind].&[Trade] }} ) ON COLUMNS FROM ( SELECT ( {{ [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] }} ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'''.format(now,now,now,now,now,now)
migrate_records(query, connString, destination_table)

print('Pulling Last 6 Months - CY')
#last 6 months current year
query ='''SELECT NON EMPTY {{ [Measures].[Net] }} ON COLUMNS, NON EMPTY {{([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) }} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( {{ [Reporting Date].[Fiscal Period].&[{0}07], [Reporting Date].[Fiscal Period].&[{1}08], [Reporting Date].[Fiscal Period].&[{2}09], [Reporting Date].[Fiscal Period].&[{3}10], [Reporting Date].[Fiscal Period].&[{4}11], [Reporting Date].[Fiscal Period].&[{5}12] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Type].&[101] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Status].&[4] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Order Kind].&[Trade] }} ) ON COLUMNS FROM ( SELECT ( {{ [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] }} ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'''.format(now,now,now,now,now,now)
migrate_records(query, connString, destination_table)

print('Pulling First 6 Months - PY')
#first 6 months last year
query ='''SELECT NON EMPTY {{ [Measures].[Net] }} ON COLUMNS, NON EMPTY {{([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) }} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( {{ [Reporting Date].[Fiscal Period].&[{0}01], [Reporting Date].[Fiscal Period].&[{1}02], [Reporting Date].[Fiscal Period].&[{2}03], [Reporting Date].[Fiscal Period].&[{3}04], [Reporting Date].[Fiscal Period].&[{4}05], [Reporting Date].[Fiscal Period].&[{5}06] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Type].&[101] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Status].&[4] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Order Kind].&[Trade] }} ) ON COLUMNS FROM ( SELECT ( {{ [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] }} ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'''.format(last_year,last_year,last_year,last_year,last_year,last_year)
migrate_records(query, connString, destination_table)

print('Pulling Last 6 Months - PY')
#last 6 months last year
query ='''SELECT NON EMPTY {{ [Measures].[Net] }} ON COLUMNS, NON EMPTY {{([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) }} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( {{ [Reporting Date].[Fiscal Period].&[{0}07], [Reporting Date].[Fiscal Period].&[{1}08], [Reporting Date].[Fiscal Period].&[{2}09], [Reporting Date].[Fiscal Period].&[{3}10], [Reporting Date].[Fiscal Period].&[{4}11], [Reporting Date].[Fiscal Period].&[{5}12] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Type].&[101] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Status].&[4] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Order Kind].&[Trade] }} ) ON COLUMNS FROM ( SELECT ( {{ [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] }} ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'''.format(last_year,last_year,last_year,last_year,last_year,last_year)
migrate_records(query, connString, destination_table)

print('Pulling P1-P3 - Prior 2 Years')
#P1-P3 prior 2 years
query ='''SELECT NON EMPTY {{ [Measures].[Net] }} ON COLUMNS, NON EMPTY {{([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) }} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( {{ [Reporting Date].[Fiscal Period].&[{0}01], [Reporting Date].[Fiscal Period].&[{1}02], [Reporting Date].[Fiscal Period].&[{2}03] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Type].&[101] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Status].&[4] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Order Kind].&[Trade] }} ) ON COLUMNS FROM ( SELECT ( {{ [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] }} ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'''.format(last_year_2,last_year_2,last_year_2)
migrate_records(query, connString, destination_table)

print('Pulling P4-P6 - Prior 2 Years')
#P4-P6 prior 2 years
query ='''SELECT NON EMPTY {{ [Measures].[Net] }} ON COLUMNS, NON EMPTY {{([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) }} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( {{ [Reporting Date].[Fiscal Period].&[{0}04], [Reporting Date].[Fiscal Period].&[{1}05], [Reporting Date].[Fiscal Period].&[{2}06] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Type].&[101] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Status].&[4] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Order Kind].&[Trade] }} ) ON COLUMNS FROM ( SELECT ( {{ [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] }} ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'''.format(last_year_2,last_year_2,last_year_2)
migrate_records(query, connString, destination_table)

print('Pulling P7-P9 - Prior 2 Years')
#P7-P9 prior 2 years
query ='''SELECT NON EMPTY {{ [Measures].[Net] }} ON COLUMNS, NON EMPTY {{([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) }} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( {{ [Reporting Date].[Fiscal Period].&[{0}07], [Reporting Date].[Fiscal Period].&[{1}08], [Reporting Date].[Fiscal Period].&[{2}09] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Type].&[101] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Status].&[4] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Order Kind].&[Trade] }} ) ON COLUMNS FROM ( SELECT ( {{ [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] }} ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'''.format(last_year_2,last_year_2,last_year_2)
migrate_records(query, connString, destination_table)

print('Pulling P10-P12 - Prior 2 Years')
#P10-P12 prior 2 years
query ='''SELECT NON EMPTY {{ [Measures].[Net] }} ON COLUMNS, NON EMPTY {{([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS ) }} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( {{ [Reporting Date].[Fiscal Period].&[{0}10], [Reporting Date].[Fiscal Period].&[{1}11], [Reporting Date].[Fiscal Period].&[{2}12] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Type].&[101] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Status].&[4] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Order Kind].&[Trade] }} ) ON COLUMNS FROM ( SELECT ( {{ [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] }} ) ON COLUMNS FROM [Revenue]))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'''.format(last_year_2,last_year_2,last_year_2)
migrate_records(query, connString, destination_table)
