import pandas as pd
from ODBC.BPCSquery import BPCSquery,read_sql_file

def IIMcleaning(IIMdf,outputpath,fileoutput=True):
 
    IIMdf = IIMdf[IIMdf["IPROD"].str.len().isin([8,10]) & IIMdf["IPROD"].str.isdigit()]
    if fileoutput is True:
        IIMdf.to_excel(outputpath,index=False)
        return
    else:
        return IIMdf

def rategroup_update(filteredIIMdf,EPLALL,output,fileoutput=True):
    #filteredIIMdf = pd.read_excel(filteredIIM,dtype={"IPROD":str})
    EPLalldf = pd.read_excel(EPLALL,usecols=['Item No','RG','PGC','GAC'],dtype={"IPROD":str})
    filteredIIMdf_newRG = pd.merge(filteredIIMdf,EPLalldf,left_on='IPROD',right_on='Item No',how='left')
    filteredIIMdf_newRG['New_RG'] = filteredIIMdf_newRG["RG"].fillna('NN')
    filteredIIMdf_newRG.loc[filteredIIMdf_newRG['IVEND'].astype(str).str.len() == 5, 'New_RG'] = 'NN'
    #filteredIIMdf_newRG['New_RG'] = filteredIIMdf_newRG.apply(lambda row: row['RG'] if pd.notnull(row['RG']) else row['IXRATG'], axis=1)
    filteredIIMdf_newRG_NN = filteredIIMdf_newRG[filteredIIMdf_newRG['New_RG'] == 'NN']
    filteredIIMdf_newRG_notNN = filteredIIMdf_newRG[filteredIIMdf_newRG['New_RG'] != 'NN']
    
    #print(filteredIIMdf_newRG)
    if fileoutput is True:
        NotNN_output = output + "notNN.csv"
        NN_output = output + "NN.csv"
        filteredIIMdf_newRG_notNN.to_csv(NotNN_output, index=False)
        filteredIIMdf_newRG_NN.to_csv(NN_output, index=False)
        
    return filteredIIMdf_newRG_notNN

def PLC_filtering(PLCpath,filteredIIM_newRG_notNN,PLCname,outputpath):
    columns = list(pd.read_excel(PLCpath, nrows=0).columns)
    
    # 检查列名并选择合适的列
    if "RG" in columns:
        usecols = ["RG"]
    elif "Rate group" in columns:
        usecols = ["Rate group"]
    else:
        raise ValueError("Neither 'RG' nor 'Rate group' columns found in the Excel file.")

    # 读取所需的列
    PLC_RG_df = pd.read_excel(PLCpath, usecols=usecols)
    if PLC_RG_df.shape[1] != 1:
        raise ValueError("PLC_RGdf should have exactly one column.")
    PLC_RG_columnname = PLC_RG_df.columns[0]
    #print(PLC_RG)
    IIM_PLC_df = filteredIIM_newRG_notNN[filteredIIM_newRG_notNN['New_RG'].isin(PLC_RG_df[PLC_RG_columnname])]
    final_outputpath = outputpath + "\\Items_newRG_inEPL" + PLCname + ".csv"
    IIM_PLC_df.to_csv(final_outputpath,index=False)

if __name__ == "__main__":
    IIMSQLpath = r"C:\Users\JPEQZ\OneDrive - Epiroc\Python\STDcost_LP_change\itemmasterSQL.sql"
    RDD = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\RF\JPE_RF_Jul2024(RDD).xlsx"
    PSD = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\RF\ZXJPE_LF&RF_Jul2024(MRS).xlsx"
    HAT = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\RF\RF 07 2024 Japan JPE_JPY(HAT).xlsx"
    filteredIIMpath = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\QRYs\Master File\Item_Master_filtered.xlsx"
    EPLALL = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\EPL\totalEPL20240701.xlsx"
    filteredIIMdf_newRG_output = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\EPL\filteredIIMdf_newRG"
    IIM_PLC_output_path = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\Items_PLC"
    PLCinfo = [(RDD, "RDD"), (PSD, "PSD"), (HAT, "HAT")]

    IIMSQL = read_sql_file(IIMSQLpath)

    IIMdf = BPCSquery(IIMSQL,"JPNPRDF")
    filteredIIMdf = IIMcleaning(IIMdf,filteredIIMpath,fileoutput=True)
    filteredIIMdf_newRG_notNN = rategroup_update(filteredIIMdf,EPLALL,filteredIIMdf_newRG_output,fileoutput=True)

    for PLC,PLCname in PLCinfo:
        PLC_filtering(PLC,filteredIIMdf_newRG_notNN,PLCname,IIM_PLC_output_path)
    print("-----------Finished!-------------")