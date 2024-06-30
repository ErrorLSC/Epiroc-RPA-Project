import pandas as pd

def pricelist_merge(domestic,oversea):
    cols = ["IPROD","New STDCOST"]
    domesticdf = pd.read_csv(domestic,usecols=cols,dtype={'IPROD':str})
    overseadf = pd.read_csv(oversea,usecols=cols,dtype={'IPROD':str})
    pricelist = pd.concat([domesticdf,overseadf],axis=0,ignore_index=True)
    #print(pricelist)
    return pricelist

def kitprice(pricelist,kitassy,IIM,output):
    IIMdf = pd.read_csv(IIM,usecols=["IPROD","ISCST"],dtype={'IPROD':str})
    kitassydf = pd.read_excel(kitassy,usecols=["SBSWKT","SBMPNO","LPROD","LQORD","WLEAD"],dtype={"SBMPNO":str,"LPROD":str})
    kitassydf["LPROD"] = kitassydf["LPROD"].str.strip()
    kitassydf["SBMPNO"] = kitassydf["SBMPNO"].str.strip()
    kitassydf = pd.merge(kitassydf,pricelist,left_on="LPROD",right_on="IPROD",how='left')
    
    #print(kitassydf)
    kitassydf["LINECOST"] = kitassydf["LQORD"] * kitassydf["New STDCOST"]
    kitassydf = kitassydf.groupby("SBMPNO").agg({
        'LINECOST':'sum',
        'WLEAD':'max',
        'SBSWKT':"first"
    }).reset_index()
    kitassydf = pd.merge(kitassydf,IIMdf,left_on="SBMPNO",right_on="IPROD",how="left")
    #print(kitassydf)
    kitassydf = kitassydf.drop(columns=["IPROD"])
    new_column_names={"SBMPNO":"IPROD","LINECOST":"REAL COST","WLEAD":"LEADTIME","SBSWKT":"TYPE","ISCST":"ISCST"}
    kitassydf = kitassydf.rename(columns=new_column_names)
    desired_order = ["TYPE","IPROD","ISCST","REAL COST","LEADTIME"]
    kitassydf = kitassydf[desired_order]
    kitassydf["Preferred COST"] = round(kitassydf["REAL COST"] * 1.4)
    kitassydf["Preferred LEADTIME"] = kitassydf["LEADTIME"] + 7
    kitassydf.to_excel(output,index=False)


if __name__ == "__main__":
    domestic = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\Items_PLC\NNitems_newprice.csv"
    oversea = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\DATA brush up\LISTPRICE&STD COST\Items_PLC\Items_newRG_inEPLPSD_withNEWSTDCOST.csv"
    kitassy = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\kit&assy\KITQUOTE.xlsx"
    kitassyoutput = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\kit&assy\KITQUOTE_newprice.xlsx"

    pricelist = pricelist_merge(domestic,oversea)
    kitprice(pricelist,kitassy,domestic,kitassyoutput)
    print("------------FINISHED!---------------")

