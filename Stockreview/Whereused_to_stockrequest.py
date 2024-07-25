import pandas as pd

def whereused(file):
    whereuseddf = pd.read_excel(file,sheet_name="Sheet1",dtype={'CPN':'str'})
    whereuseddf = whereuseddf.groupby('CPN').agg(Model_List=('Model', lambda x: set(x)),Serial_Count=('Serial Number', 'count')).reset_index()
    notuseddf = pd.read_excel(file,sheet_name="Ignored Parts",dtype={'CPN':'str'})
    notuseddf["Model_List"] = "要調査"
    notuseddf["Serial_Count"] = "要調査"
    whereuseddf = pd.concat([whereuseddf,notuseddf]).reset_index(drop=True)
    return whereuseddf

def to_stockrequest(whereuseddf,stockrequestlist,output):
    stockrequestdf = pd.read_excel(stockrequestlist,sheet_name="MIN BAL解除予定",usecols=["Part Number","Description","Requested Qty","Cost Unit"],dtype={"Part Number":'str'})
    stockrequestdf = pd.merge(stockrequestdf,whereuseddf,left_on='Part Number',right_on='CPN',how="left")
    stockrequestdf = stockrequestdf.drop(columns=['CPN'])
    stockrequestdf = stockrequestdf.rename(columns={'Requested Qty':'現MIN Balance&在庫数','Cost Unit':'STD COST','Serial_Count':'稼働機数'})
    stockrequestdf.to_excel(output,index=False)

if __name__ == "__main__":
    filepath1 = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Stock Plan\whereusedreport\Where Used - Epiroc Japan - 7_23_2024 9_46_35 AM.xlsx"
    filepath2 = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Stock Plan\whereusedreport\Where Used - Epiroc Japan - 7_23_2024 9_52_04 AM.xlsx"
    filepath3 = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Stock Plan\whereusedreport\Where Used - Epiroc Japan - 7_23_2024 9_56_28 AM.xlsx"
    stockrequestlist = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Stock Plan\31303_PartNumbers_20240709.xlsx"
    outputpath = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Stock Plan\MIN BALANCE REVIEW.xlsx"
    filelist = [filepath1,filepath2,filepath3]
    df_list = []
    for file in filelist:
        df = whereused(file)
        df_list.append(df)

    resultdf = pd.concat(df_list,ignore_index=True).reset_index(drop=True)
    to_stockrequest(resultdf,stockrequestlist,outputpath)
