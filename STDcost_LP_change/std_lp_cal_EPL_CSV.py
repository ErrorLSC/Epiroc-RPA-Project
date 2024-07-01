import pandas as pd
#import math

def calculate_stdcost(row,SEK_rate,LCF):
    if not pd.isnull(row['IRP']):
        if not pd.isnull(row['RF']):
            result=row['RF'] * SEK_rate * LCF * row['IRP']
            return round(result)
    else:
        return row['ISCST']
    
def calculate_margin(row):
    if row['ILIST'] != 0:
        margin = (row['ILIST'] - row['New STDCOST']) / row['ILIST']
        return round(margin,4)
    else:
        return None

def EPL_merge(EPL,IIM_PLC,RF,SEK_rate,LCF):
    IIMcols = ["IPROD","CXPPLC","IVEND",'IXRATG',"New_RG","ISCST","ILIST","STKOH01",'CXATLC','PGC']
    EPLcols = ["Item No","IRP"]
    RFcols = ["RG","RF"]
    IIMdf = pd.read_csv(IIM_PLC,usecols=IIMcols,dtype={"IPROD":str},index_col="IPROD")
    EPLdf = pd.read_excel(EPL,usecols=EPLcols,dtype={"Item No":str},index_col="Item No")
    RFdf = pd.read_excel(RF,usecols=RFcols)
    IIMdf = pd.merge(IIMdf,EPLdf,left_index=True,right_index=True,how="left")
    IIMdf = IIMdf.reset_index()
    IIMdf = pd.merge(IIMdf,RFdf,left_on="New_RG",right_on='RG', how="left")
    IIMdf['STKVAL'] = IIMdf["ISCST"] * IIMdf["STKOH01"]
    #IIMdf['RG'] = IIMdf["RG"].fillna("NN")
    IIMdf['New STDCOST'] = IIMdf.apply(calculate_stdcost,args=(SEK_rate,LCF,),axis=1)
    IIMdf['NEW STKVAL'] = IIMdf["New STDCOST"] * IIMdf["STKOH01"] 
    IIMdf['STDCOST Change Rate'] = round((IIMdf["New STDCOST"] - IIMdf['ISCST'])/IIMdf['ISCST'],4)
    IIMdf['CC Margin'] = IIMdf.apply(calculate_margin,axis=1)
    path = IIM_PLC[:-4] + "_withNEWSTDCOST.csv"
    IIMdf.to_csv(path,index=False)
    #print(IIMdf)

if __name__ == "__main__":
    EPL_HAT = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\EPL\HAT_EPL_20240701_SEK_total.xlsx"
    EPL_RDD = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\EPL\RDD_EPL_20240701.xlsx"
    EPL_PSD = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\EPL\PSD_EPL_20240701.xlsx"

    RF_RDD = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\RF\JPE_RF_Jul2024(RDD).xlsx"
    RF_PSD = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\RF\ZXJPE_LF&RF_Jul2024(MRS).xlsx"
    RF_HAT = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\RF\RF 07 2024 Japan JPE_JPY(HAT).xlsx"
    
    IIM_HAT = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\Items_PLC\Items_newRG_inEPLHAT.csv"
    IIM_PSD = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\Items_PLC\Items_newRG_inEPLPSD.csv"
    IIM_RDD = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\Items_PLC\Items_newRG_inEPLRDD.csv"

    SEK_rate = 13.6989492
    LCF = 1.098

    PLCinfo = [(EPL_HAT, IIM_HAT,RF_HAT), (EPL_RDD, IIM_RDD,RF_RDD), (EPL_PSD, IIM_PSD,RF_PSD)]
    for EPL,IIM,RF in PLCinfo:
        EPL_merge(EPL,IIM,RF,SEK_rate,LCF)
    #EPL_merge(EPL_PSD, IIM_PSD,RF_PSD,SEK_rate,LCF)
    print("-----------Finished!-------------")