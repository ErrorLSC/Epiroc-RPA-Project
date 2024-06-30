import pandas as pd
#from std_lp_cal_EPL_CSV import calculate_margin 
from ODBC.BPCSquery import BPCSquery,read_sql_file

def calculate_margin(row):
    if row['ILIST'] != 0:
        margin = (row['ILIST'] - row['New STDCOST']) / row['ILIST']
        return round(margin,4)
    else:
        return None

def vendorquote(hqtdf,outputpath=None,fileoutput=True):

    hqtdf = hqtdf.sort_values('HQQDT')
    hqtdf = hqtdf.drop_duplicates(subset='HQPROD',keep='last')

    #print(hqtdf)
    if fileoutput is True:
        hqtdf.to_csv(outputpath,index=False)
    return hqtdf

def IIMrefresh_byvendorquote(hqt,IIM,output):
    IIMcols = ["IPROD","CXPPLC","IVEND","IXRATG","New_RG","ISCST","ILIST","STKOH01"]
    IIMdf = pd.read_csv(IIM,usecols=IIMcols,dtype={"IPROD":str})
    IIMdf = pd.merge(IIMdf,hqt,left_on="IPROD",right_on="HQPROD",how="left")
    IIMdf['New STDCOST'] = IIMdf.apply(lambda row: row['HQPR1'] if pd.notnull(row['HQPR1']) else row['ISCST'], axis=1)
    
    IIMdf['STKVAL'] = IIMdf["ISCST"] * IIMdf["STKOH01"]
    IIMdf['NEW STKVAL'] = IIMdf["New STDCOST"] * IIMdf["STKOH01"] 
    IIMdf['STDCOST Change Rate'] = round((IIMdf["New STDCOST"] - IIMdf['ISCST'])/IIMdf['ISCST'],4)
    IIMdf['CC Margin'] = IIMdf.apply(calculate_margin,axis=1)
    IIMdf.to_csv(output,index=False)

if __name__ == "__main__":
    vendor_quote_master_SQL = r"C:\Users\JPEQZ\OneDrive - Epiroc\Python\STDcost_LP_change\vendorquoteSQL.txt"
    IIM = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\EPL\filteredIIMdf_newRGNN.csv"
    pricelistpath = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\EPL\domestic_price.csv"
    IIMoutput = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\Items_PLC\NNitems_newprice.csv"

    query = read_sql_file(vendor_quote_master_SQL)
    vendor_quote_master_df = BPCSquery(query,"JPNPRDF",datecolumnlist=['HQQDT'])

    hqtdf = vendorquote(vendor_quote_master_df,pricelistpath)
    IIMrefresh_byvendorquote(hqtdf,IIM,IIMoutput)
    print("---------------FINISHED-----------------")
