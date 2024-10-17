import openpyxl
import pandas as pd
import credentials
from databricks import sql

#
FY = "2024" # Fiscal Year

#
connection = sql.connect(server_hostname = credentials.hostname ,http_path = credentials.path,access_token = credentials.token)
cursor = connection.cursor()

# load excel file
xlfile = "input.xlsx"
sheetname = "Sheet1"
wb_obj = openpyxl.load_workbook(xlfile,read_only=True)
sheet_obj = wb_obj[sheetname]
md_df = pd.DataFrame(sheet_obj.values)
wb_obj.close()
new_header0 = md_df.iloc[0] 
md_df = md_df[1:] 
md_df.columns = new_header0 
md_df.drop([1])
#print(md_df.to_markdown())



mm = md_df.copy(deep=True)

no_items = len(md_df)
for x in range(no_items):
    mnl = md_df.iloc[x]['Material No']
    bn = str(md_df.iloc[x]['Batch'])
    mn = str(mnl).zfill(len(str(mnl))+10)
    #sqlquery = "SELECT MATNR, CHARG, CINSM, LFMON FROM efdataonelh_prd.generaldiscovery_matmgt_r.all_mchbh_view where MATNR = '" + mn + "' AND CHARG = '" + bn + "' AND LFGJA = '2024' ORDER BY CHARG DESC LIMIT 1"
    sqlquery = "SELECT MATNR, CHARG, CINSM, LFMON FROM efdataonelh_prd.generaldiscovery_matmgt_r.all_mchbh_view where MATNR = '" + mn + "' AND CHARG = '" + bn + "'  AND LFGJA = '2024'"
    cursor.execute(sqlquery)
    mchbh = pd.DataFrame(cursor.fetchall(), columns=["Material Number","Batch Number", "Stocks", "Current Period"  ])
    print(mchbh.to_markdown())
    if len(mchbh) > 0:
        s = mchbh.iloc[0]['Stocks']
        md_df.iloc[x]['Remarks'] = s
    else:
        s = ''
    print("Material no: ", mnl," Stocks", s)
print(md_df.to_markdown())

