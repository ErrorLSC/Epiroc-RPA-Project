import pandas as pd
from datetime import datetime
import os

def newest_file_in_directory(directory):
    # 获取目录中所有文件的列表
    files = os.listdir(directory)
    # 过滤掉非文件的项目
    files = [f for f in files if os.path.isfile(os.path.join(directory, f))]
    # 如果文件列表为空，返回None
    if not files:
        return None
    # 获取文件列表中第一个文件的路径和最后修改时间
    newest_file = files[0]
    newest_mtime = os.path.getmtime(os.path.join(directory, newest_file))
    # 遍历文件列表，查找最新修改的文件
    for file in files:
        file_path = os.path.join(directory, file)
        mtime = os.path.getmtime(file_path)
        if mtime > newest_mtime:
            newest_file = file
            newest_mtime = mtime
    # 返回最新修改的文件的路径
    return os.path.join(directory, newest_file)

def real_stock(stkoh,ELAL01):
    stkdf = pd.read_excel(stkoh,skiprows=4,usecols=["Whs","Loc","DIVISION","Item Number","ITEM DESCRIPTION","STOCK ON HAND"],dtype={"Item Number":str})
    stkdf = stkdf.rename(columns = {"Whs":"LWHS","Loc":"LLOC","Item Number":"LPROD","ITEM DESCRIPTION":"IDESC","STOCK ON HAND":"QTYOH"})
    stkdf["LLOC"] = stkdf["LLOC"].str.strip()
    stkdf = stkdf[stkdf["LWHS"] == "5"]
    stkdf = stkdf[stkdf["LLOC"] != "X"]
    stkdf = stkdf[stkdf["LLOC"] != "CC MIN"]
    stkdf = stkdf[stkdf["LLOC"] != "MM CC"]
    stkdf = stkdf[stkdf["LLOC"] != "YM MIN"]
    stkdf["LPROD"] = stkdf["LPROD"].str.strip()
    stkdf = stkdf.groupby("LPROD").agg({"QTYOH":'sum',"LLOC":'first',"IDESC":'first',"DIVISION":'first'})
    ELALdf = pd.read_excel(ELAL01,dtype={"Item Number":str,"Whse":str},skiprows=3)
    ELALdf = ELALdf[ELALdf["Whse"]== "5"]
    ELALdf["Item Number"] = ELALdf["Item Number"].str.strip()
    ELALdf = ELALdf.groupby("Item Number")['Quant Alloc'].sum()
    realstockse = stkdf["QTYOH"].sub(ELALdf, fill_value=0)
    #realstockse.to_excel("temp.xlsx")
    stkdf["QTYOH"] = realstockse
    #stkdf.to_excel("temp.xlsx")
    #print(ELALdf)
    return stkdf

def difflist(BPCS,LCAT,ResultSave):
    LCATdf = pd.read_excel(LCAT,usecols=["商品コード","数量","ロケーション"],dtype={"商品コード":str})
    LCATdf = LCATdf.groupby("商品コード").agg({"数量":'sum',"ロケーション":'first'})
    LCATdf.index.name = "LPROD"
    LCATdf = LCATdf.rename(columns={"数量":"QTYOH_LCAT","ロケーション":'LOC_LCAT'})
    LCATdf = LCATdf[LCATdf["LOC_LCAT"] != "X"]

    resultdf = pd.concat([BPCS,LCATdf], axis = 1)
    resultdf['QTYOH'] = resultdf['QTYOH'].fillna(0)
    resultdf['QTYOH_LCAT'] = resultdf['QTYOH_LCAT'].fillna(0)
    resultdf['DIFF'] = resultdf["QTYOH"] - resultdf['QTYOH_LCAT']
    resultdf = resultdf[resultdf['DIFF'] !=0]
    zerostockdf = resultdf[resultdf["QTYOH_LCAT"] == 0]
    with pd.ExcelWriter(ResultSave) as writer:
        zerostockdf.to_excel(writer, sheet_name='LCAT在庫なし')
        resultdf.to_excel(writer,sheet_name='総合差異表')
    

if __name__ == "__main__":

    BPCSdic = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\QRYs\STKOH"
    LCATdic = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\QRYs\LCAT_STOCK"
    Resultdic = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Stock take\202404"
    ELALdic = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\QRYs\OUTBOUND\ELAL01"

    stkoh = newest_file_in_directory(BPCSdic)
    LCAT = newest_file_in_directory(LCATdic)
    ELAL01 = newest_file_in_directory(ELALdic)
    today_date = datetime.today().date().strftime("%Y-%m-%d")
    ResultSave = Resultdic + "\DiffList_" + today_date + ".xlsx"

    BPCS = real_stock(stkoh,ELAL01)
    difflist(BPCS,LCAT,ResultSave)
    print("Finished!")
