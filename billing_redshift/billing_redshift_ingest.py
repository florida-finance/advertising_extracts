
import pandas as pd
from datetime import datetime
import win32com.client
import psycopg2
import tempfile
import boto3


# import warnings
# warnings.filterwarnings("ignore")

# pd.set_option('display.max_columns', None)  
# pd.set_option('display.expand_frame_repr', False)
# pd.set_option('max_colwidth', -1)

#Function for chunking lists
def chunks(l, n):
    """Yield successive n-sized chunks from l."""
    for i in range(0, len(l), n):
        yield l[i:i + n]

def fix_records(dataframe):
    dataframe.rename(columns={0:'company_code', 2:'gl_entity', 4:'gl_sub_entity', 6:'gl_product_code', 8:'ledger_account', 10:'order_ad_size', 12:'order_ad_type',
                       14:'parent_name_number', 16:'sales_category', 18:'sales_subcategory', 20:'product_code', 22:'parent_product', 24:'product_type', 
                       26:'fiscal_quarter',28:'fiscal_period',30:'fiscal_week', 32:'sold_to_customer_name',34:'sold_to_customer_name_number',
                       36:'order_color',38:'product_section_code',40:'pub_date',42:'pub_fiscal_period', 44:'pub_fiscal_year', 46:'net'},inplace=True)
    
    #dataframe['pub_date'] = dataframe['pub_date'].apply(pd.to_datetime, format='%Y-%m-%d')
    
    return dataframe


#Connections string
connString = "PROVIDER=MSOLAP;Data Source={0};Database={1}".format('fcwPsqlanl03','Billing_2')

#Get Customer Name Numbers
query = ' SELECT NON EMPTY { [Measures].[Net] } ON COLUMNS, NON EMPTY { ([Sold To].[Customer Name Number].[Customer Name Number].ALLMEMBERS ) } DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( { [Publication Date].[Fiscal Year].&[2016], [Publication Date].[Fiscal Year].&[2017], [Publication Date].[Fiscal Year].&[2018], [Publication Date].[Fiscal Year].&[2019] } ) ON COLUMNS FROM ( SELECT ( { [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] } ) ON COLUMNS FROM [Revenue])) WHERE ( [Company].[Company Code].CurrentMember, [Publication Date].[Fiscal Year].CurrentMember ) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'
rs = win32com.client.Dispatch(r'ADODB.Recordset')
rs.Open(query, connString, CursorType=3)
customer_list = list(pd.DataFrame(data= list(rs.GetRows())).transpose()[0].unique())

#Get All Billing results and append to query results
query_result = pd.DataFrame()
for i in chunks(customer_list,300):
    z = ",".join(list(map(lambda x: '[Sold To].[Customer Name Number].&['+x+']',i)))
    query ='SELECT NON EMPTY {{ [Measures].[Net] }} ON COLUMNS, NON EMPTY {{ ([Company].[Company Code].[Company Code].ALLMEMBERS * [Product].[GL Entity Code].[GL Entity Code].ALLMEMBERS * [Product].[GL Sub Entity Code].[GL Sub Entity Code].ALLMEMBERS * [Product].[GL Product Code].[GL Product Code].ALLMEMBERS * [Order].[Ledger Account].[Ledger Account].ALLMEMBERS * [Order].[Ad Size].[Ad Size].ALLMEMBERS * [Order].[Ad Type].[Ad Type].ALLMEMBERS * [Sold To].[Parent Name Number].[Parent Name Number].ALLMEMBERS * [Order].[Sales Category].[Sales Category].ALLMEMBERS * [Order].[Sales Sub Category].[Sales Sub Category].ALLMEMBERS * [Product].[Product Code].[Product Code].ALLMEMBERS * [Product].[Parent Product].[Parent Product].ALLMEMBERS * [Product].[Product Type].[Product Type].ALLMEMBERS * [Reporting Date].[Fiscal Quarter].[Fiscal Quarter].ALLMEMBERS * [Reporting Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Reporting Date].[Fiscal Week].[Fiscal Week].ALLMEMBERS * [Sold To].[Customer Name].[Customer Name].ALLMEMBERS * [Sold To].[Customer Name Number].[Customer Name Number].ALLMEMBERS * [Order].[Color].[Color].ALLMEMBERS * [Product].[Section Code].[Section Code].ALLMEMBERS * [Publication Date].[Date].[Date].ALLMEMBERS * [Publication Date].[Fiscal Period].[Fiscal Period].ALLMEMBERS * [Publication Date].[Fiscal Year].[Fiscal Year].ALLMEMBERS  ) }} DIMENSION PROPERTIES MEMBER_CAPTION, MEMBER_UNIQUE_NAME ON ROWS FROM ( SELECT ( {{  [Publication Date].[Fiscal Year].&[2016], [Publication Date].[Fiscal Year].&[2017], [Publication Date].[Fiscal Year].&[2018], [Publication Date].[Fiscal Year].&[2019] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Type].&[101] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Sales Status].&[4] }} ) ON COLUMNS FROM ( SELECT ( -{{ [Order].[Order Kind].&[Trade] }} ) ON COLUMNS FROM ( SELECT ( {{ {} }} ) ON COLUMNS FROM ( SELECT ( {{ [Company].[Company Code].&[OSC], [Company].[Company Code].&[SSC] }} ) ON COLUMNS FROM [Revenue])))))) CELL PROPERTIES VALUE, BACK_COLOR, FORE_COLOR, FORMATTED_VALUE, FORMAT_STRING, FONT_NAME, FONT_SIZE, FONT_FLAGS'.format(z) 
    rs = win32com.client.Dispatch(r'ADODB.Recordset')
    rs.Open(query, connString, CursorType=3)
    query_result = query_result.append(pd.DataFrame(data= list(rs.GetRows())).transpose())
    #query_result = fix_records(query_result)
query_result  = fix_records(query_result)
query_result = query_result[['company_code','gl_entity','gl_sub_entity','gl_product_code','ledger_account','order_ad_size','order_ad_type', 'order_color',
                       'parent_name_number', 'sold_to_customer_name','sold_to_customer_name_number','sales_category','sales_subcategory',
                       'product_code','parent_product','product_type', 'product_section_code','pub_date','pub_fiscal_period','pub_fiscal_year',
                       'fiscal_quarter','fiscal_period','fiscal_week','net']]

#upload file to S3
file_name = r'billing/billing_pull_'+str(datetime.now().strftime('%m_%d_%Y'))+'.csv'
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
     TRUNCATE TABLE financial_reporting.billing;
     copy financial_reporting.billing
     from {}
     access_key_id {}
     secret_access_key {}
     IGNOREHEADER AS 1
     csv;
     COMMIT;
 """.format(from_file, key, secret_key)

cur.execute(query_template)

