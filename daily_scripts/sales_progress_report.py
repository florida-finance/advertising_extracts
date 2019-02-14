import pandas as pd
from datetime import datetime
from sqlalchemy import create_engine
import win32com.client
from sqlalchemy.sql import text as sa_text

connString = "PROVIDER=MSOLAP;Data Source={0};Database={1}".format('fcwPsqlanl03','Crediting')
def ado(strsql, connString):
    rs = win32com.client.Dispatch(r'ADODB.Recordset')
    rs.Open(strsql, connString)
    t = rs.GetRows()
    rs.Close()
    return t

#Define Connection to Postgres Database
hostname = 'mktstrategy.ciklurvi0auw.us-east-1.rds.amazonaws.com'
username = 'tronc'
password = 'tronc123123!'
database = 'Crediting'
port=5432

def connect(user, password, db, host, port=5432):
    url  = 'postgresql://{}:{}@{}:{}/{}'
    url = url.format(user, password, host, port, db)

    # The return value of create_engine() is our connection object
    con = create_engine(url, client_encoding='utf8')
    return con
c =connect(username, password, database, hostname)

cy = str(datetime.now().year)
ly = str(int(datetime.now().year)-1)
ly2 = str(int(datetime.now().year)-2)

#Digital Revenue by SA ID from Crediting for cy+
query = 'SELECT NON EMPTY {{ [Measures].[Commission Net] }} ON COLUMNS, NON EMPTY {{ ([Credit Sales Assignment].[Sales Assignment Code].[Sales Assignment Code].ALLMEMBERS * [Credit Sales Assignment].[Sales Assignment Sub Team ID].[Sales Assignment Sub Team ID].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) }} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( -{{ [Credit Sales Assignment].[Sales Assignment Code].&[0] }} ) ON COLUMNS FROM ( SELECT ( {{[Reporting Date].[Fiscal Year].&[{0}] }} ) ON COLUMNS FROM ( SELECT ( {{ [Product].[Product Type].&[Alternative Digital], [Product].[Product Type].&[Online] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Status].&[4] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Type].&[101] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Order Kind].&[Trade] }} ) ON COLUMNS FROM ( SELECT ( {{ [Company].[Company Code].&[MOT], [Company].[Company Code].&[SSC], [Company].[Company Code].&[OSC] }} ) ON COLUMNS FROM [Crediting]))))))) WHERE ( [Company].[Company Code].CurrentMember, [Reporting Date].[Fiscal Year].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'.format(cy)
a = ado(query, connString)
df = pd.DataFrame(data= list(a)).transpose()
df.rename(columns = {0:'sa_id', 2:'sub_team', 4:'fiscal_period', 6:'commission_net_dig'},inplace=True)
df = df[['sa_id', 'sub_team', 'fiscal_period', 'commission_net_dig']]
df['fiscal_year'] = df.fiscal_period.apply(lambda x: x[0:4])
df['fiscal_period'] = df.fiscal_period.apply(lambda x: int(x[-2:]))
df = df[['sa_id', 'sub_team', 'fiscal_year', 'fiscal_period', 'commission_net_dig']]
df = df.groupby(by=['sa_id','fiscal_year', 'fiscal_period'])['commission_net_dig'].sum().reset_index()
digital = df.copy()

query = ' SELECT NON EMPTY {{ [Measures].[Commission Net] }} ON COLUMNS, NON EMPTY {{ ([Credit Sales Assignment].[Sales Assignment Code].[Sales Assignment Code].ALLMEMBERS * [Credit Sales Assignment].[Sales Assignment Sub Team ID].[Sales Assignment Sub Team ID].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) }} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( -{{ [Credit Sales Assignment].[Sales Assignment Code].&[0] }} ) ON COLUMNS FROM ( SELECT ( {{[Reporting Date].[Fiscal Year].&[{0}] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Product].[Product Type].&[Direct Mail] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Status].&[4] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Type].&[101] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Order Kind].&[Trade] }} ) ON COLUMNS FROM ( SELECT ( {{ [Company].[Company Code].&[MOT], [Company].[Company Code].&[SSC], [Company].[Company Code].&[OSC] }} ) ON COLUMNS FROM [Crediting]))))))) WHERE ( [Company].[Company Code].CurrentMember, [Reporting Date].[Fiscal Year].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'.format(cy)
a = ado(query, connString)
df = pd.DataFrame(data= list(a)).transpose()
df.rename(columns = {0:'sa_id', 2:'sub_team', 4:'fiscal_period', 6:'commission_net_all'},inplace=True)
df = df[['sa_id', 'sub_team', 'fiscal_period', 'commission_net_all']]
df['fiscal_year'] = df.fiscal_period.apply(lambda x: x[0:4])
df['fiscal_period'] = df.fiscal_period.apply(lambda x: int(x[-2:]))
df = df[['sa_id', 'sub_team', 'fiscal_year', 'fiscal_period', 'commission_net_all']]
df = df.groupby(by=['sa_id','fiscal_year', 'fiscal_period'])['commission_net_all'].sum().reset_index()
all_in = df.copy()


results = all_in.merge(digital, how='left', on=['sa_id', 'fiscal_year', 'fiscal_period'])
results['commission_net_print'] = results['commission_net_all'] - results['commission_net_dig']
results.fillna(0,inplace=True)


connString = "PROVIDER=MSOLAP;Data Source={0};Database={1}".format('fcwPsqlanl03','TMModelingCube')
query = 'SELECT NON EMPTY {{ [Measures].[Commission Net] }} ON COLUMNS, NON EMPTY {{ ([TM Sales Assignment].[SA ID].[SA ID].ALLMEMBERS *[TM Sales Assignment].[Sales Assignment Sub Team].[Sales Assignment Sub Team].ALLMEMBERS *[Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) }} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( {{ [Reporting Date].[Fiscal Year].&[{0}],[Reporting Date].[Fiscal Year].&[{1}] }} ) ON COLUMNS FROM ( SELECT ( -{{ [TM Sales Assignment].[SA ID].&[Missing] }} ) ON COLUMNS FROM ( SELECT ( {{ [Product].[Product Type].&[Alternative Digital], [Product].[Product Type].&[Online] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Type].&[101] }} ) ON COLUMNS FROM ( SELECT ( {{ [Order].[Sales Status].&[3] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Order Kind].&[Trade] }} ) ON COLUMNS FROM ( SELECT ( {{ [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC], [Company].[Company Code].&[MOT] }} ) ON COLUMNS FROM [Territory Management Modeling]))))))) WHERE ( [Company].[Company Code].CurrentMember, [Order].[Sales Status].&[3], [Product].[Product Type].CurrentMember, [Reporting Date].[Fiscal Year].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'.format(ly, ly2)
a = ado(query, connString)
df = pd.DataFrame(data= list(a)).transpose()
df.rename(columns = {0:'sa_id', 2:'sub_team', 4:'fiscal_period', 6:'commission_net_dig'},inplace=True)
df = df[['sa_id', 'sub_team', 'fiscal_period', 'commission_net_dig']]
df['fiscal_year'] = df.fiscal_period.apply(lambda x: x[0:4])
df['fiscal_period'] = df.fiscal_period.apply(lambda x: int(x[-2:]))
df = df[['sa_id', 'sub_team', 'fiscal_year', 'fiscal_period', 'commission_net_dig']]
df = df.groupby(by=['sa_id','fiscal_year', 'fiscal_period'])['commission_net_dig'].sum().reset_index()
digital = df.copy()


query = ' SELECT NON EMPTY {{ [Measures].[Commission Net] }} ON COLUMNS, NON EMPTY {{ ([TM Sales Assignment].[SA ID].[SA ID].ALLMEMBERS * [TM Sales Assignment].[Sales Assignment Sub Team].[Sales Assignment Sub Team].ALLMEMBERS *  [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS ) }} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( {{ [Reporting Date].[Fiscal Year].&[{0}],[Reporting Date].[Fiscal Year].&[{1}]  }} ) ON COLUMNS FROM ( SELECT ( -{{ [TM Sales Assignment].[SA ID].&[Missing] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Product].[Product Type].&[Direct Mail] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Type].&[101] }} ) ON COLUMNS FROM ( SELECT ( {{ [Order].[Sales Status].&[3] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Order Kind].&[Trade] }} ) ON COLUMNS FROM ( SELECT ( {{ [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC], [Company].[Company Code].&[MOT] }} ) ON COLUMNS FROM [Territory Management Modeling]))))))) WHERE ( [Company].[Company Code].CurrentMember, [Order].[Sales Status].&[3], [Reporting Date].[Fiscal Year].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'.format(ly2,ly)
a = ado(query, connString)
df = pd.DataFrame(data= list(a)).transpose()
df.rename(columns = {0:'sa_id', 2:'sub_team', 4:'fiscal_period', 6:'commission_net_all'},inplace=True)
df = df[['sa_id','sub_team', 'fiscal_period', 'commission_net_all']]
df['fiscal_year'] = df.fiscal_period.apply(lambda x: x[0:4])
df['fiscal_period'] = df.fiscal_period.apply(lambda x: int(x[-2:]))
df = df[['sa_id', 'sub_team','fiscal_year', 'fiscal_period', 'commission_net_all']]
df = df.groupby(by=['sa_id','fiscal_year', 'fiscal_period'])['commission_net_all'].sum().reset_index()
all_in = df.copy()


results2 = all_in.merge(digital, how='left', on=['sa_id', 'fiscal_year', 'fiscal_period'])
results2['commission_net_print'] = results2['commission_net_all'] - results2['commission_net_dig']
results2.fillna(0,inplace=True)


results['asof_date'] = datetime.today().strftime('%Y-%m-%d')
results2['asof_date'] = datetime.today().strftime('%Y-%m-%d')


results = results[['sa_id','fiscal_year', 'fiscal_period',
       'commission_net_all', 'commission_net_dig',
       'commission_net_print', 'asof_date']]

results2 = results2[['sa_id','fiscal_year', 'fiscal_period',
       'commission_net_all', 'commission_net_dig',
       'commission_net_print', 'asof_date']]

#delete today's records
today_date = datetime.today().strftime('%Y-%m-%d')
query = "DELETE FROM commission_sales WHERE asof_date = '{}".format(today_date)+"'"
c.execute(sa_text(query).execution_options(autocommit=True))

results.to_sql('commission_sales', c, if_exists='append', index=False)
results2.to_sql('commission_sales', c, if_exists='append', index=False)

