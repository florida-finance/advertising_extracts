#!/usr/bin/env python
# coding: utf-8

# In[6]:


#Packages
import pandas as pd
from datetime import datetime, timedelta, date
import pandas.io.sql as psql
import pandasql
from sqlalchemy import create_engine
from sqlalchemy import MetaData
import time
import win32com.client

connString = "PROVIDER=MSOLAP;Data Source={0};Database={1}".format('fcwPsqlanl03','Crediting')
def ado(strsql, connString):
    rs = win32com.client.Dispatch(r'ADODB.Recordset')
    rs.Open(strsql, connString)
    t = rs.GetRows()
    rs.Close()
    return t


# ### Fetch Sun-Sentinel Local+Non Local Revenue 2018

# In[7]:


#Local Sans ADSS
query = 'SELECT NON EMPTY { [Measures].[Commission Net] } ON COLUMNS, { ([Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Auto], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Circulation], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Direct Mail], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_House/Other], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_Broward], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_Jewish Jrnl], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_Palm], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_SF Parenting], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_House/Other], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Inside_Sales], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_City], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_House/Othr], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_Palm], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_TMC], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Majors], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Marketing], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_4], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Niche_CityShore], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Niche_elSentinel], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Premium_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Support], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Transactional_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_WSFL], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Inactive], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Cars.com] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Year].&[2018] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Crediting]))))) WHERE ( [Company].[Company Code].&[SSC], [Reporting Date].[Fiscal Year].&[2018], [Credit Sales Assignment].[Sales Assignment Sub Team ID].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
a = ado(query, connString)
df = pd.DataFrame(data= list(a)).transpose()
df.rename(columns= {0:'Period',2:'Value'},inplace=True)
df['Period'] = df['Period'].apply(lambda x: x[-2:])
df['Period'] = df['Period'].apply(lambda x: int(x))
df = df[['Period','Value']]
df['Value_Type'] = 'Output_Net'
df.loc[df['Value'].isnull(),'Value']=0


#Local ADSS
query = 'SELECT NON EMPTY { [Measures].[Commission Net] } ON COLUMNS, { ([Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Self-Service] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Year].&[2018] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Crediting]))))) WHERE ( [Company].[Company Code].&[SSC], [Reporting Date].[Fiscal Year].&[2018], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Self-Service] ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
a = ado(query, connString)
df2 = pd.DataFrame(data= list(a)).transpose()
df2.rename(columns= {0:'Period',2:'Value'},inplace=True)
df2['Period'] = df2['Period'].apply(lambda x: x[-2:])
df2['Period'] = df2['Period'].apply(lambda x: int(x))
df2 = df2[['Period','Value']]
df2['Value_Type'] = 'Self-Service'
df2.loc[df2['Value'].isnull(),'Value']=0


#Local Digital Only
query = 'SELECT NON EMPTY { [Measures].[Commission Net] } ON COLUMNS, { ([Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Product].[Parent Product].&[FPG-INT-Forum Publishing], [Product].[Parent Product].&[SSC-INT-Hyperlocal], [Product].[Parent Product].&[SSC-INT-Hyperlocal Other], [Product].[Parent Product].&[SSC-INT-Replica], [Product].[Parent Product].&[SSC-INT-Sun-Sentinel], [Product].[Parent Product].&[GCP-INT-el Sentinel] } ) ON COLUMNS FROM ( SELECT ( { [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Auto], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Circulation], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Direct Mail], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_House/Other], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_Jewish Jrnl], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_Palm], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_SF Parenting], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_House/Other], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Inside_Sales], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_City], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_House/Othr], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_Palm], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_TMC], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Majors], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Marketing], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_4], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Niche_CityShore], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Niche_elSentinel], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Premium_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Support], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Transactional_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_WSFL], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Inactive], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_5] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Year].&[2018] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Crediting])))))) WHERE ( [Company].[Company Code].&[SSC], [Reporting Date].[Fiscal Year].&[2018], [Credit Sales Assignment].[Sales Assignment Sub Team ID].CurrentMember, [Product].[Parent Product].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
a = ado(query, connString)
df3 = pd.DataFrame(data= list(a)).transpose()
df3.rename(columns= {0:'Period',2:'Value'},inplace=True)
df3['Period'] = df3['Period'].apply(lambda x: x[-2:])
df3['Period'] = df3['Period'].apply(lambda x: int(x))
df3 = df3[['Period','Value']]
df3['Value_Type'] = 'Local_Digital'
df3.loc[df3['Value'].isnull(),'Value']=0


#Merge
df = df.append(df2)
df = df.append(df3)
df['Date_Run'] = datetime.today().strftime('%Y-%m-%d')
df = df[['Period','Value','Date_Run','Value_Type']]
df['fiscal_year'] = 2018
df1 = df.copy()

# In[6]:

#Local Sans ADSS
query = 'SELECT NON EMPTY { [Measures].[Commission Net] } ON COLUMNS, { ([Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Auto], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Circulation], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Direct Mail], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_House/Other], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_Broward], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_Jewish Jrnl], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_Palm], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_SF Parenting], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_House/Other], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Inside_Sales], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_City], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_House/Othr], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_Palm], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_TMC], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Majors], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Marketing], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_4], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Niche_CityShore], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Niche_elSentinel], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Premium_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Support], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Transactional_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_WSFL], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Inactive], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Cars.com] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Year].&[2019] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Crediting]))))) WHERE ( [Company].[Company Code].&[SSC], [Reporting Date].[Fiscal Year].&[2019], [Credit Sales Assignment].[Sales Assignment Sub Team ID].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
a = ado(query, connString)
df = pd.DataFrame(data= list(a)).transpose()
df.rename(columns= {0:'Period',2:'Value'},inplace=True)
df['Period'] = df['Period'].apply(lambda x: x[-2:])
df['Period'] = df['Period'].apply(lambda x: int(x))
df = df[['Period','Value']]
df['Value_Type'] = 'Output_Net'
df.loc[df['Value'].isnull(),'Value']=0


#Local ADSS
query = 'SELECT NON EMPTY { [Measures].[Commission Net] } ON COLUMNS, { ([Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Self-Service] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Year].&[2019] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Crediting]))))) WHERE ( [Company].[Company Code].&[SSC], [Reporting Date].[Fiscal Year].&[2019], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Self-Service] ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
a = ado(query, connString)
df2 = pd.DataFrame(data= list(a)).transpose()
df2.rename(columns= {0:'Period',2:'Value'},inplace=True)
df2['Period'] = df2['Period'].apply(lambda x: x[-2:])
df2['Period'] = df2['Period'].apply(lambda x: int(x))
df2 = df2[['Period','Value']]
df2['Value_Type'] = 'Self-Service'
df2.loc[df2['Value'].isnull(),'Value']=0


#Local Digital Only
query = 'SELECT NON EMPTY { [Measures].[Commission Net] } ON COLUMNS, { ([Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Product].[Parent Product].&[FPG-INT-Forum Publishing], [Product].[Parent Product].&[SSC-INT-Hyperlocal], [Product].[Parent Product].&[SSC-INT-Hyperlocal Other], [Product].[Parent Product].&[SSC-INT-Replica], [Product].[Parent Product].&[SSC-INT-Sun-Sentinel], [Product].[Parent Product].&[GCP-INT-el Sentinel] } ) ON COLUMNS FROM ( SELECT ( { [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Auto], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Circulation], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Direct Mail], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_House/Other], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_Jewish Jrnl], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_Palm], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_FPG_SF Parenting], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_House/Other], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Inside_Sales], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_City], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_House/Othr], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_Palm], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Local_TMC], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Majors], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Marketing], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_4], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Niche_CityShore], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Niche_elSentinel], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Premium_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Support], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Transactional_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_WSFL], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Inactive], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[SSC_Media_5] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Year].&[2019] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Crediting])))))) WHERE ( [Company].[Company Code].&[SSC], [Reporting Date].[Fiscal Year].&[2019], [Credit Sales Assignment].[Sales Assignment Sub Team ID].CurrentMember, [Product].[Parent Product].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
a = ado(query, connString)
df3 = pd.DataFrame(data= list(a)).transpose()
df3.rename(columns= {0:'Period',2:'Value'},inplace=True)
df3['Period'] = df3['Period'].apply(lambda x: x[-2:])
df3['Period'] = df3['Period'].apply(lambda x: int(x))
df3 = df3[['Period','Value']]
df3['Value_Type'] = 'Local_Digital'
df3.loc[df3['Value'].isnull(),'Value']=0


#Merge
df = df.append(df2)
df = df.append(df3)
df['Date_Run'] = datetime.today().strftime('%Y-%m-%d')
df = df[['Period','Value','Date_Run','Value_Type']]
df['fiscal_year'] = 2019
df2 = df.copy()

df = df1.append(df2)

#Append to SQL Table

#Define Connection to Postgres Database
hostname = 'mktstrategy.ciklurvi0auw.us-east-1.rds.amazonaws.com'
username = 'tronc'
password = 'tronc123123!'
database = 'OLAP_Exports'
port=5432

def connect(user, password, db, host, port=5432):
    url  = 'postgresql://{}:{}@{}:{}/{}'
    url = url.format(user, password, host, port, db)

    # The return value of create_engine() is our connection object
    con = create_engine(url, client_encoding='utf8')
    return con
c =connect(username, password, database, hostname)

df.to_sql('ssc_results', c, if_exists='append', index=False)
dataframe = psql.read_sql('SELECT * from "ssc_results"', c)
dataframe.sort_values(by=['Value'],ascending=False,inplace=True)
dataframe.drop_duplicates(subset=['Period','Date_Run','Value_Type'],inplace=True)
from sqlalchemy.sql import text as sa_text
c.execute(sa_text('''TRUNCATE TABLE "ssc_results"''').execution_options(autocommit=True))
dataframe.to_sql('ssc_results', c, if_exists='append', index=False)


# ### Fetch Orlando Sentinel Local+Non Local Revenue 2018


#Local Sans ADSS
query = ' SELECT NON EMPTY { [Measures].[Commission Net] } ON COLUMNS, {([Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( { [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Stats_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Support], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Transactional_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Retail_Ntnl_3],[Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Premium_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Retail_Local_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Retail_Local], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Premium_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Digital_Local], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Digital_Ntnl], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_DirectMail], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Digital_Ntnl_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Events], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_House/Other], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Marketing], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Inactive], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Media_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Media_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Media_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_NewBizDev], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Premium_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Auto], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Circulation]} ) ON COLUMNS FROM ( SELECT ( { [Company].[Company].&[OSC] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Year].&[2018] } ) ON COLUMNS FROM [Crediting]))))) WHERE ( [Reporting Date].[Fiscal Year].&[2018], [Company].[Company].CurrentMember, [Credit Sales Assignment].[Sales Assignment Sub Team ID].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
a = ado(query, connString)
df = pd.DataFrame(data= list(a)).transpose()
df.rename(columns= {0:'Period',2:'Value'},inplace=True)
df['Period'] = df['Period'].apply(lambda x: x[-2:])
df['Period'] = df['Period'].apply(lambda x: int(x))
df = df[['Period','Value']]
df['Value_Type'] = 'Output_Net'
df.loc[df['Value'].isnull(),'Value']=0

#Local ADSS
query='SELECT NON EMPTY { [Measures].[Commission Net] } ON COLUMNS, {([Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Self-Service] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Year].&[2018] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( {[Company].[Company].&[OSC] } ) ON COLUMNS FROM [Crediting]))))) WHERE ( [Company].[Company].CurrentMember, [Reporting Date].[Fiscal Year].&[2018], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Self-Service] ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
a = ado(query, connString)
df2 = pd.DataFrame(data= list(a)).transpose()
df2.rename(columns= {0:'Period',2:'Value'},inplace=True)
df2['Period'] = df2['Period'].apply(lambda x: x[-2:])
df2['Period'] = df2['Period'].apply(lambda x: int(x))
df2 = df2[['Period','Value']]
df2['Value_Type'] = 'Orlando_ADSS'
df2.loc[df2['Value'].isnull(),'Value']=0


#Local Digital Only
query=' SELECT NON EMPTY { [Measures].[Commission Net] } ON COLUMNS, NON EMPTY {([Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Product].[Parent Product].&[OSC-INT-Orlando Sentinel], [Product].[Parent Product].&[OSC-INT-Signature] } ) ON COLUMNS FROM ( SELECT ( { [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Auto], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Circulation], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Digital_Local], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Digital_Ntnl], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Digital_Ntnl_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_DirectMail], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Events], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_House/Other], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Inactive], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Marketing], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Media_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Media_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Media_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_NewBizDev], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Premium_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Premium_2],[Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Premium_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Premium_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Retail_Local], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Retail_Local_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Retail_Ntnl_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Self-Service], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Stats_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Support], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Transactional_1] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Year].&[2018] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[OSC] } ) ON COLUMNS FROM [Crediting])))))) WHERE ( [Company].[Company Code].&[OSC], [Reporting Date].[Fiscal Year].&[2018], [Credit Sales Assignment].[Sales Assignment Sub Team ID].CurrentMember, [Product].[Parent Product].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
a = ado(query, connString)
df3 = pd.DataFrame(data= list(a)).transpose()
df3.rename(columns= {0:'Period',2:'Value'},inplace=True)
df3['Period'] = df3['Period'].apply(lambda x: x[-2:])
df3['Period'] = df3['Period'].apply(lambda x: int(x))
df3 = df3[['Period','Value']]
df3['Value_Type'] = 'Local_Digital'
df3.loc[df3['Value'].isnull(),'Value']=0


#Merge
df = df.append(df2)
df = df.append(df3)
df['Date_Run'] = datetime.today().strftime('%Y-%m-%d')
df = df[['Period','Value','Date_Run','Value_Type']]
df['fiscal_year'] = 2018
df1 = df.copy()



# ### Fetch Orlando Sentinel Local+Non Local Revenue 2018


#Local Sans ADSS
query = ' SELECT NON EMPTY { [Measures].[Commission Net] } ON COLUMNS, {([Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( { [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Stats_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Support], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Transactional_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Retail_Ntnl_3],[Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Premium_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Retail_Local_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Retail_Local], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Premium_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Digital_Local], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Digital_Ntnl], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_DirectMail], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Digital_Ntnl_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Events], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_House/Other], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Marketing], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Inactive], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Media_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Media_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Media_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_NewBizDev], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Premium_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Auto], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Circulation]} ) ON COLUMNS FROM ( SELECT ( { [Company].[Company].&[OSC] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Year].&[2019] } ) ON COLUMNS FROM [Crediting]))))) WHERE ( [Reporting Date].[Fiscal Year].&[2019], [Company].[Company].CurrentMember, [Credit Sales Assignment].[Sales Assignment Sub Team ID].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
a = ado(query, connString)
df = pd.DataFrame(data= list(a)).transpose()
df.rename(columns= {0:'Period',2:'Value'},inplace=True)
df['Period'] = df['Period'].apply(lambda x: x[-2:])
df['Period'] = df['Period'].apply(lambda x: int(x))
df = df[['Period','Value']]
df['Value_Type'] = 'Output_Net'
df.loc[df['Value'].isnull(),'Value']=0

#Local ADSS
query='SELECT NON EMPTY { [Measures].[Commission Net] } ON COLUMNS, {([Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Self-Service] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Year].&[2019] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( {[Company].[Company].&[OSC] } ) ON COLUMNS FROM [Crediting]))))) WHERE ( [Company].[Company].CurrentMember, [Reporting Date].[Fiscal Year].&[2019], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Self-Service] ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
a = ado(query, connString)
df2 = pd.DataFrame(data= list(a)).transpose()
df2.rename(columns= {0:'Period',2:'Value'},inplace=True)
df2['Period'] = df2['Period'].apply(lambda x: x[-2:])
df2['Period'] = df2['Period'].apply(lambda x: int(x))
df2 = df2[['Period','Value']]
df2['Value_Type'] = 'Orlando_ADSS'
df2.loc[df2['Value'].isnull(),'Value']=0


#Local Digital Only
query=' SELECT NON EMPTY { [Measures].[Commission Net] } ON COLUMNS, NON EMPTY {([Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Product].[Parent Product].&[OSC-INT-Orlando Sentinel], [Product].[Parent Product].&[OSC-INT-Signature] } ) ON COLUMNS FROM ( SELECT ( { [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Auto], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Circulation], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Digital_Local], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Digital_Ntnl], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Digital_Ntnl_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_DirectMail], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Events], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_House/Other], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Inactive], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Marketing], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Media_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Media_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Media_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_NewBizDev], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Premium_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Premium_2],[Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Premium_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Premium_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Retail_Local], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Retail_Local_2], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Retail_Ntnl_3], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Self-Service], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Stats_1], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Support], [Credit Sales Assignment].[Sales Assignment Sub Team ID].&[OSC_Transactional_1] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( { [Reporting Date].[Fiscal Year].&[2019] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[OSC] } ) ON COLUMNS FROM [Crediting])))))) WHERE ( [Company].[Company Code].&[OSC], [Reporting Date].[Fiscal Year].&[2019], [Credit Sales Assignment].[Sales Assignment Sub Team ID].CurrentMember, [Product].[Parent Product].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
a = ado(query, connString)
df3 = pd.DataFrame(data= list(a)).transpose()
df3.rename(columns= {0:'Period',2:'Value'},inplace=True)
df3['Period'] = df3['Period'].apply(lambda x: x[-2:])
df3['Period'] = df3['Period'].apply(lambda x: int(x))
df3 = df3[['Period','Value']]
df3['Value_Type'] = 'Local_Digital'
df3.loc[df3['Value'].isnull(),'Value']=0


#Merge
df = df.append(df2)
df = df.append(df3)
df['Date_Run'] = datetime.today().strftime('%Y-%m-%d')
df = df[['Period','Value','Date_Run','Value_Type']]
df['fiscal_year'] = 2019
df2 = df.copy()

df = df1.append(df2)


df.to_sql('Results', c, if_exists='append', index=False)
dataframe = psql.read_sql('SELECT * from "Results"', c)
dataframe.sort_values(by=['Value'],ascending=False,inplace=True)
dataframe.drop_duplicates(subset=['Period','Date_Run','Value_Type'],inplace=True)
c.execute(sa_text('''TRUNCATE TABLE "Results"''').execution_options(autocommit=True))
dataframe.to_sql('Results', c, if_exists='append', index=False)
