import pandas as pd
from std_lp_cal_EPL import calculate_margin 

def vendorquote(filepath,outputpath=None,fileoutput=True):
    hqtdf = pd.read_excel(filepath,dtype={'HQPROD':str})
    hqtdf['HQPROD'] = hqtdf['HQPROD'].str.strip()
    hqtdf = hqtdf.sort_values('HQQDT')
    hqtdf = hqtdf.drop_duplicates(subset='HQPROD',keep='last')
    hqtdf = hqtdf.drop(columns=["ISCST"])
    #hqtdf = hqtdf.set_index("HQPROD")
    #print(hqtdf)
    if fileoutput is True:
        hqtdf.to_excel(outputpath,index=False)
    return hqtdf

def IIMrefresh_byvendorquote(hqt,IIM,output):
    IIMcols = ["IPROD","CXPPLC","IVEND","IXRATG","New_RG","ISCST","ILIST","STKOH01","CXLCFP"]
    IIMdf = pd.read_excel(IIM,sheet_name='NN',usecols=IIMcols,dtype={"IPROD":str})
    IIMdf = pd.merge(IIMdf,hqt,left_on="IPROD",right_on="HQPROD",how="left")
    IIMdf['New STDCOST'] = IIMdf.apply(lambda row: row['HQPR1'] if pd.notnull(row['HQPR1']) else row['ISCST'], axis=1)
    
    IIMdf['STKVAL'] = IIMdf["ISCST"] * IIMdf["STKOH01"]
    IIMdf['NEW STKVAL'] = IIMdf["New STDCOST"] * IIMdf["STKOH01"] 
    IIMdf['STDCOST Change Rate'] = round((IIMdf["New STDCOST"] - IIMdf['ISCST'])/IIMdf['ISCST'],4)
    IIMdf['CC Margin'] = IIMdf.apply(calculate_margin,axis=1)
    IIMdf.to_excel(output,index=False)
    



if __name__ == "__main__":
    vendor_quote_master = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\QRYs\Master File\Vendor Quote.xlsx"
    IIM = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\EPL\filteredIIMdf_newRG.xlsx"
    pricelistpath = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\EPL\domestic_price.xlsx"
    IIMoutput = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\Items_PLC\NNitems_newprice.xlsx"

    

    hqtdf = vendorquote(vendor_quote_master,pricelistpath)
    IIMrefresh_byvendorquote(hqtdf,IIM,IIMoutput)
    print("---------------FINISHED-----------------")
