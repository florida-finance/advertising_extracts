import pandas as pd
from datetime import datetime
#from datetime import timedelta, date
#import pandas.io.sql as psql
#import pandasql
from sqlalchemy import create_engine
#from sqlalchemy.sql import text as sa_text
import psycopg2
import win32com.client
import tempfile
import boto3


#Function for chunking lists
def chunks(l, n):
    """Yield successive n-sized chunks from l."""
    for i in range(0, len(l), n):
        yield l[i:i + n]

#Define Connection to Postgres Database
def connect(user, password, db, host, port=5432):
    url  = 'postgresql://{}:{}@{}:{}/{}'
    url = url.format(user, password, host, port, db)

    # The return value of create_engine() is our connection objectpip
    con = create_engine(url, client_encoding='utf8')
    return con


#Connection criteria for postgresdb    
hostname = 'mktstrategy.ciklurvi0auw.us-east-1.rds.amazonaws.com'
username = 'tronc'
password = 'tronc123123!'
database = 'Crediting'

#create connection object
c = connect(username, password, database, hostname)

#Bring in all rep info
rep_mappings = pd.read_sql('SELECT DISTINCT company_name, sa_id, rep_name, manager_name, fiscal_year, sales_team FROM weekly_progress_report', c)
#sa_id_list = list(rep_mappings.sa_id.unique())
#sa_id_list = str(list(map(lambda x: '[TM Sales Assignment].[SA ID].&['+x+']',sa_id_list)))[1:-1].replace("'","")

#Get Customer Name Numbers
connString = "PROVIDER=MSOLAP;Data Source={0};Database={1}".format('fcwPsqlanl03','TMModelingCube')
query = 'SELECT NON EMPTY { [Measures].[Commission Net] } ON COLUMNS, NON EMPTY { ([Sold To].[Customer Name Number].[Customer Name Number].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Publication Date].[Fiscal Year].&[2018], [Publication Date].[Fiscal Year].&[2019] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company].&[OSC], [Company].[Company].&[SSC] } ) ON COLUMNS FROM [Territory Management Modeling])) WHERE ( [Company].[Company].CurrentMember, [Publication Date].[Fiscal Year].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
rs = win32com.client.Dispatch(r'ADODB.Recordset')
rs.Open(query, connString, CursorType=3)
customer_list = list(pd.DataFrame(data= list(rs.GetRows())).transpose()[0].unique())

#current_year = datetime.now().strftime('%Y')
#last_year = str(int(datetime.now().strftime('%Y'))-1)


query_result = pd.DataFrame()
#iterate through customer names and apply to tm mdx query
for i in chunks(customer_list,500):
    z = ",".join(list(map(lambda x: '[Sold To].[Customer Name Number].&['+x+']',i)))
    query = 'SELECT NON EMPTY {{ [Measures].[Commission Net], [Measures].[Net] }} ON COLUMNS, NON EMPTY {{ ([TM Sales Assignment].[SA ID].[SA ID].ALLMEMBERS * [TM Sales Assignment].[Sales Assignment Sub Team].[Sales Assignment Sub Team].ALLMEMBERS * [TM Sales Assignment].[Sales Assignment Employee].[Sales Assignment Employee].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Sold To].[Parent Name].[Parent Name].ALLMEMBERS * [Sold To].[Customer Name].[Customer Name].ALLMEMBERS * [Sold To].[Customer Name Number].[Customer Name Number].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Publication Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Publication Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Publication Date].[Date].[Date].ALLMEMBERS * [Bill To].[Customer Name Number].[Customer Name Number].ALLMEMBERS * [Bill To].[Customer Name].[Customer Name].ALLMEMBERS * [Bill To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Order].[Order Kind].[Order Kind].ALLMEMBERS * [Order].[Sales Type].[Sales Type].ALLMEMBERS * [Order].[Sales Status].[Sales Status].ALLMEMBERS  * [TM Sales Assignment].[Sales Assignment Percent].[Sales Assignment Percent].ALLMEMBERS) }} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( {{ [Publication Date].[Fiscal Year].&[2018], [Publication Date].[Fiscal Year].&[2019] }} ) ON COLUMNS FROM ( SELECT ( {{ [Company].[Company Code].&[SSC], [Company].[Company Code].&[OSC] }} ) ON COLUMNS FROM ( SELECT ( {{ {} }} ) ON COLUMNS FROM [Territory Management Modeling]))) WHERE ( [Company].[Company Code].CurrentMember, [Reporting Date].[Fiscal Year].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'.format(z)
    rs = win32com.client.Dispatch(r'ADODB.Recordset')
    rs.Open(query, connString, CursorType=3)
    query_result = query_result.append(pd.DataFrame(data= list(rs.GetRows())).transpose())


query_result.rename(columns = {0:'sa_id', 2:'sub_team', 4:'tm_employee_name', 6:'sold_to_parent_name_number',
                     8:'sold_to_parent_name', 10:'sold_to_customer_name', 12:'sold_to_customer_name_number', 14:'product_type',
                     16:'fiscal_period', 18:'fiscal_quarter', 20:'pub_date',  22:'bill_to_customer_name_number', 24:'bill_to_customer_name', 26:'bill_to_parent_name_number',
                     28:'product_code', 30:'order_sales_subcategory', 32:'order_kind',  34:'order_sales_type', 36:'sales_status',
                     38:'sales_assignment_percentage', 40:'commnet',41:'net'},inplace=True)

query_result = query_result[['sa_id', 'sub_team', 'tm_employee_name','sold_to_parent_name_number',
                     'sold_to_parent_name', 'sold_to_customer_name', 'sold_to_customer_name_number', 'product_type',
                     'fiscal_period', 'fiscal_quarter', 'pub_date',  'bill_to_customer_name_number', 'bill_to_customer_name', 'bill_to_parent_name_number',
                     'product_code', 'order_sales_subcategory', 'order_kind', 'order_sales_type', 'sales_status',
                     'sales_assignment_percentage','commnet', 'net']]

query_result = query_result.merge(rep_mappings, how='inner', left_on = ['sa_id'], right_on  = ['sa_id'])
query_result.drop(labels = ['fiscal_year'], axis =1, inplace=True)
query_result.rename(columns = {'rep_name':'sales_rep_name'}, inplace=True)

query_result = query_result[['company_name','sa_id', 'sub_team', 'sales_team','tm_employee_name',  'sales_rep_name', 'manager_name', 'sold_to_parent_name_number', 'sold_to_parent_name', 'sold_to_customer_name', 'sold_to_customer_name_number', 'product_type', 
                'fiscal_period', 'fiscal_quarter', 'pub_date', 'bill_to_customer_name_number', 'bill_to_customer_name', 'bill_to_parent_name_number', 'product_code', 'order_sales_subcategory', 
                'order_kind', 'order_sales_type', 'sales_status', 'sales_assignment_percentage', 'commnet', 'net']]

    
query_result['pub_date'] = query_result['pub_date'].apply(pd.to_datetime, format='%Y-%m-%d')



#upload file to S3
file_name = r'tm_modeling/tm_pull_'+str(datetime.now().strftime('%m_%d_%Y'))+'.csv'
aws_access_key_id = 'AKIAIG4JTEG4R2TJJREA'
aws_secret_access_key = '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'

#Create temporary file with results
with tempfile.NamedTemporaryFile(mode='wb', delete=False) as fp:
    query_result.to_csv(fp.name, index=False)    
    s3 = boto3.client('s3',aws_access_key_id=aws_access_key_id,aws_secret_access_key=aws_secret_access_key)
    s3.upload_file(fp.name,'bi-data-warehouse-00',file_name)


#Get results from S3 and load into redshift
from_file =  r"'s3://bi-data-warehouse-00/"+file_name+"'"
key = r"'"+aws_access_key_id+ r"'"
secret_key = r"'"+aws_secret_access_key+ r"'"

conn = psycopg2.connect("dbname = 'dev' user= 'mpolissky' host='redshift.troncdata.com' port='5439' password= '3C9oGfQddbh4E0Mq'")
cur = conn.cursor()
cur.execute("SET AUTOCOMMIT = ON;")

#load new file
query_template = """
     TRUNCATE TABLE financial_reporting.tm_modeling;
     copy financial_reporting.tm_modeling
     from {}
     access_key_id {}
     secret_access_key {}
     IGNOREHEADER AS 1
     delimiter ','
     csv;
     COMMIT;
 """.format(from_file, key, secret_key)

cur.execute(query_template)
