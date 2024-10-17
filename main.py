import openpyxl
import pandas as pd
import credentials


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
print(md_df.to_markdown())