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

connString = "PROVIDER=MSOLAP;Data Source={0};Database={1}".format('fcwPsqlanl03','Billing_2')
def ado(strsql, connString):
    rs = win32com.client.Dispatch(r'ADODB.Recordset')
    rs.Open(strsql, connString)
    t = rs.GetRows()
    rs.Close()
    return t


# In[2]:


now = datetime.now().strftime('%Y')
last_year = int(datetime.now().strftime('%Y'))-1
now = 'FY'+str(now)[-2:]
last_year = 'FY'+str(last_year)[-2:]


# In[3]:


Final = pd.DataFrame(columns=["Entity","SubEntity","Scenario","Version","Years", "Period", "Department", "Product","Retail Local Preprints (400300)","General Local Preprints (402300)","Retail T365 Preprints (400350)","General T365 Preprints"])
data = {"exportPlanningData" : True, "gridDefinition" : {"suppressMissingBlocks" : True, "pov" : {"dimensions" : ["Plan Element"],"members" : [ [ "Total Plan" ]]},"columns" : [ {"dimensions" : [ "Account" ],"members" : [ ['Retail Local Preprints (400300)','General Local Preprints (402300)','Retail T365 Preprints (400350)','General T365 Preprints (402350)'] ]}],"rows" : [ {"dimensions" : [ "Entity","SubEntity","Scenario","Version","Years", "Period","Department", "Product"],"members" : [ [ "Sun Sentinel (E_13000)", "South Florida Community News (E_13015)" ] ,["Default - SubEntity (S_00000)","El Sentinel (S_71013)", "City and Shore (S_13005)", "Tribune Direct (S_70000)", "South Florida Parenting (S_13016)", "Jewish Journal (S_13021)","Coral Springs Group (S_13018)"],["Actual","Plan","Forecast"],["Final","Working","Feb Fcst", "Mar Fcst", "Apr Fcst", "May Fcst", "Jun Fcst", "Jul Fcst", "Aug Fcst", "Sep Fcst", "Oct Fcst", "Nov Fcst", "Dec Fcst"], [last_year,now],["P1","P2","P3", "P4","P5","P6","P7","P8","P9","P10","P11","P12"],['All Department'],["AEBITDA Products"]]} ]}}
url="https://planning-a503899.pbcs.us6.oraclecloud.com//HyperionPlanning/rest/v3/applications/EPBCS/plantypes/OEP_FS/exportdataslice"

encoded_data = json.dumps(data).encode('utf-8')
r = http.request('POST', url, body=encoded_data, headers={'Content-Type': 'application/json','authorization': 'Basic YTUwMzg5OS5tcG9saXNza3k6UmltbWE5MzEyMw=='})
if r.status == 400:
    pass
else:
    response = json.loads(r.data.decode('utf-8'))
    rows = []
    for x in response['rows']:
        z = []
        for y in x['headers']:
            z.append(y)
        for m in x['data']:
            z.append(m)
        rows.append(z)

A = pd.DataFrame(columns=["Entity","SubEntity","Scenario","Version","Years", "Period", "Department", "Product","Retail Local Preprints (400300)","General Local Preprints (402300)","Retail T365 Preprints (400350)","General T365 Preprints"], data=rows)
Final = Final.append(A)
Final = Final.melt(id_vars = ["Entity","SubEntity","Scenario","Version","Years", "Period", "Department", "Product"], value_vars=["Retail Local Preprints (400300)","General Local Preprints (402300)","Retail T365 Preprints (400350)","General T365 Preprints"],var_name="Account")
Hyperion_Pull = Final[Final['value']!='']
Hyperion_Pull.loc[Hyperion_Pull.Period=='Jan','Period']=1
Hyperion_Pull.loc[Hyperion_Pull.Period=='Feb','Period']=2
Hyperion_Pull.loc[Hyperion_Pull.Period=='Mar','Period']=3
Hyperion_Pull.loc[Hyperion_Pull.Period=='Apr','Period']=4
Hyperion_Pull.loc[Hyperion_Pull.Period=='May','Period']=5
Hyperion_Pull.loc[Hyperion_Pull.Period=='Jun','Period']=6
Hyperion_Pull.loc[Hyperion_Pull.Period=='Jul','Period']=7
Hyperion_Pull.loc[Hyperion_Pull.Period=='Aug','Period']=8
Hyperion_Pull.loc[Hyperion_Pull.Period=='Sep','Period']=9
Hyperion_Pull.loc[Hyperion_Pull.Period=='Oct','Period']=10
Hyperion_Pull.loc[Hyperion_Pull.Period=='Nov','Period']=11
Hyperion_Pull.loc[Hyperion_Pull.Period=='Dec','Period']=12
Hyperion_Pull.loc[Hyperion_Pull.Years=='FY17','Years']=2017
Hyperion_Pull.loc[Hyperion_Pull.Years=='FY18','Years']=2018
Hyperion_Pull.loc[Hyperion_Pull.SubEntity=='S_13018','SubEntity'] = 'Coral Springs Group (S_13018)'
Hyperion_Pull.loc[Hyperion_Pull.SubEntity=='S_00000','SubEntity'] = 'Default - SubEntity (S_00000)'
Hyperion_Pull.loc[Hyperion_Pull.SubEntity=='S_13021','SubEntity'] = 'Jewish Journal (S_13021)'
Hyperion_Pull.loc[Hyperion_Pull.SubEntity=='S_71013','SubEntity'] = 'El Sentinel (S_71013)'
Hyperion_Pull.loc[Hyperion_Pull.Entity=='E_13015','Entity'] = 'South Florida Community News (E_13015)'
Hyperion_Pull.loc[Hyperion_Pull.Entity=='E_13000','Entity'] = 'Sun-Sentinel (E_13000)'
Hyperion_Pull['AsOf_Date'] = datetime.today().strftime('%Y-%m-%d')


# In[4]:


#Billing Cube Pull
query=' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Reporting Date].[Fiscal Year].[Fiscal Year].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Sold To].[Parent Category].[Parent Category].ALLMEMBERS * [Sold To].[Parent Sub Category].[Parent Sub Category].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Reporting Date].[Fiscal Year].&[2017], [Reporting Date].[Fiscal Year].&[2018], [Reporting Date].[Fiscal Year].&[2019] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Order Kind].&[Trade] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Status].&[4] } ) ON COLUMNS FROM ( SELECT ( -{ [Order].[Sales Type].&[101] } ) ON COLUMNS FROM ( SELECT ( { [Product].[Product Type].&[Alternative Print], [Product].[Product Type].&[Preprint] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company].&[SSC] } ) ON COLUMNS FROM [Revenue])))))) WHERE ( [Company].[Company].&[SSC], [Product].[Product Type].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
a = ado(query, connString)
df = pd.DataFrame(data= list(a)).transpose()

df.rename(columns={0:'fiscal_year',2:'fiscal_period',4:'fiscal_week',6:'sales_team',8:'customer_name_num',10:'customer_category',12:'customer_subcategory',14:'parent_product',16:'product_code',18:'ad_note', 20:'commission_net'},inplace=True)

df = df[['fiscal_year','fiscal_period','fiscal_week','sales_team','customer_name_num','customer_category','customer_subcategory','parent_product','product_code','ad_note','commission_net']]
df.fiscal_week = df.fiscal_week.apply(lambda x: int(x[-2:]))
df.fiscal_period = df.fiscal_period.apply(lambda x: int(x[-2:]))
df.head()
df['AsOf_Date'] = datetime.today().strftime('%Y-%m-%d')
df.loc[df['sales_team']=='400300','sales_team']= 'Retail Local Preprints (400300)'
df.loc[df['sales_team']=='402300','sales_team']= 'General Local Preprints (402300)'
df.loc[df['sales_team']=='400350','sales_team']= 'Retail T365 Preprints (400350)'
df.loc[df['sales_team']=='402350','sales_team']= 'General T365 Preprints (402350)'
df = df.loc[df['sales_team']!='103130',:]
df['sales team']='Local'
df.loc[df['sales_team']=='Retail Local Preprints (400300)','sales team']= 'Local'
df.loc[df['sales_team']=='General Local Preprints (402300)','sales team']= 'Local'
df.loc[df['sales_team']=='Retail T365 Preprints (400350)','sales team']= 'National'
df.loc[df['sales_team']=='General T365 Preprints (402350)','sales team']= 'National'
df.rename(columns={'sales_team':'gl_account'},inplace=True)
df.rename(columns={'sales team':'sales_team'},inplace=True)


# In[5]:


#Append to SQL Table

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

df.to_sql('preprint_by_sales_team_ssc', c, if_exists='replace', index=False)
Hyperion_Pull.to_sql('preprint_gl_summary_ssc', c, if_exists='replace', index=False)
