import pandas as pd
import numpy as np
from ODBC.BPCSquery import BPCSquery
from CurrentMonth.CurrentMonth import dateperiod

def outboundhist(PICKHIST,KITQUOTEMASTER):
    pickdf = PICKHIST
    kitquotedf = pd.read_excel(KITQUOTEMASTER,sheet_name="KITLINENUMS",dtype={"SBMPNO":str})
    kitquotedf["SBMPNO"] = kitquotedf["SBMPNO"].str.strip()
    pickdf = pd.merge(pickdf, kitquotedf, how='left', left_on='S#PROD', right_on='SBMPNO')
    pickdf['PICKCOUNT'] = pickdf.apply(lambda row: 0 if row['SBSWKT'] in ['KIT', 'ASX'] else row['Count of LLINE'],axis=1)
    pickdf['PICKCOUNT'] = pickdf['PICKCOUNT'].fillna(1)    
    pickdf.to_csv('pickdf.csv')
    oversea_count = pickdf[pickdf['HCUST'] == 9125]['PICKCOUNT'].sum()
    pickdf = pickdf[pickdf['HCUST'] != 9125] 
    
    pickdf = pickdf.groupby('CXPPLC')['PICKCOUNT'].count()
    pickdf.loc['LDC+'] = oversea_count
    print(pickdf)
    return pickdf

def inboundhist(INBOUNDHIST):
    inbounddf = INBOUNDHIST
    inbounddf = inbounddf.groupby(['CXPPLC','TTYPE'])['TPROD'].count().reset_index()
    inbounddf = inbounddf.rename(columns={'TPROD': 'Total'})
    LDCindex = inbounddf.loc[(inbounddf['CXPPLC'] == 'MRS') & (inbounddf['TTYPE'] == 'H')].index
    inbounddf.loc[LDCindex,'CXPPLC'] = 'LDC+'
    inbounddf = inbounddf.groupby('CXPPLC')['Total'].sum()
    return inbounddf

def totalcount(pickdf,inbounddf):
    totaldf = pickdf.add(inbounddf, fill_value=0)
    ratio = totaldf.div(totaldf.sum())
    return ratio

def HIST_bill(YAMATOBILL,totalratio):
    billdf = pd.read_excel(YAMATOBILL,sheet_name='pagetable2')
    billdf['本体金額'] = billdf['本体金額'].str.replace(' ','').str.replace(',','')
    billdf['本体金額'] = billdf['本体金額'].astype(float)
    billdf = billdf[billdf['項目'] != '宅急便運賃']
    division_fee = billdf.iloc[:4]['本体金額'].sum()
    DC_fee = billdf.iloc[4:]['本体金額'].sum()
    fee_detail = totalratio * division_fee
    fee_detail.loc['7718'] = DC_fee
    return fee_detail

def YamatoShipment(YAMATODETAIL,PICKHIST):
    detaildf = pd.read_excel(YAMATODETAIL,skiprows=1,parse_dates=['日付'])
    #print(detaildf)
    detaildf = detaildf.dropna(subset=['伝票番号'])
    df_nonorders = detaildf[pd.to_numeric(detaildf['ORD#'], errors='coerce').isna()]
    df_orders = detaildf[~pd.to_numeric(detaildf['ORD#'], errors='coerce').isna()]
    pickdf = PICKHIST
    df_orders = df_orders[['日付','ORD#','単価']]
    
    pickdf = pickdf.drop_duplicates(subset=['S#ORD','S#CDTE'])
    pickdf['S#ORD'] = pickdf['S#ORD'].astype(str)
    pickdf['S#CDTE'] = pickdf['S#CDTE'].astype(str)
    df_orders['ORD#'] = df_orders['ORD#'].astype(str)
    df_orders['日付'] = df_orders['日付'].astype(str)
    #pickdf = pickdf.set_index(['S#ORD', 'S#CDTE'])
    #pickdf.index.names = ['ORD#', '日付']
    #df_orders = df_orders.set_index(['ORD#', '日付']).join(pickdf['CXPPLC'],how='left')
    df_orders = pd.merge(df_orders, pickdf, left_on=['ORD#', '日付'], right_on=['S#ORD', 'S#CDTE'], how='left')
    df_blankorders = df_orders[df_orders['CXPPLC'].isnull()]
    #print(df_orders)
    df_orders = df_orders.groupby('CXPPLC')['単価'].sum()
    with pd.ExcelWriter('YAMATO_BILL/OUTPUT/Yamato/YamatoSHIPMENT.xlsx') as writer:
        df_nonorders.to_excel(writer,sheet_name='YamatoNONORDER')
        df_orders.to_excel(writer,sheet_name='YamatoORDER')
        df_blankorders.to_excel(writer,sheet_name='blankORDER')
    return df_orders

def FinalFee(fee_detail):
    FinalFee = fee_detail
    FinalFee = np.ceil(FinalFee)
    FinalFee.to_excel('YAMATO_BILL/OUTPUT/Yamato/人件費.xlsx')

if __name__ == "__main__":
    YAMATOBILL =  r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Bills\Yamato\202407\Yamatotable.xlsx"
    YAMATODETAIL = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Bills\Yamato\202407\【7月度】運賃明細データ.xlsx"
    KITQUOTEMASTER = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\kit&assy\KITQUOTE.xlsx"
    start_date = dateperiod("int")[0]
    end_date = dateperiod("int")[1]

    #start_date = 20240416
    #end_date = 20240515

    PICKHIST_SQL = f"""SELECT P.S#ORD,P.S#PROD,O.HCUST,X.CXPPLC,P.S#CDTE,S.SBSWKT
    FROM JPNLOCF.ESRW P
    LEFT JOIN JPNPRDF.ECHL02 O
    ON P.S#ORD=O.HORD
    LEFT JOIN JPNPRDF.IIML01 I
    ON P.S#PROD=I.IPROD
    LEFT JOIN JPNPRDF.ICXL01 X
    ON I.ICLAS=X.CXCLAS
    LEFT JOIN JPNPRDF.JSBL02 S
    ON P.S#ORD=S.SBHORD
    WHERE P.S#WHSE='5'
    AND P.S#CDTE BETWEEN {start_date} AND {end_date}"""

    INBOUND_SQL = f"""SELECT T01.TREF,T01.TPROD,T02.CXPPLC,T01.TTYPE
    FROM ITH T01
    INNER JOIN ICX T02 
    ON T01.TCLAS=T02.CXCLAS
    WHERE T01.TWHS IN ('5', 'A1')
    AND T01.TTDTE BETWEEN {start_date} AND {end_date}
    AND T01.TTYPE IN ('H', 'GA')
    AND TLOCT NOT IN ('BARAKI', '')"""

    PICKHIST = BPCSquery(PICKHIST_SQL,"JPNALL",datecolumnlist=['S#CDTE'])

    INBOUNDHIST = BPCSquery(INBOUND_SQL,"JPNPRDF")
    pickdf = outboundhist(PICKHIST,KITQUOTEMASTER)
    
    inbounddf = inboundhist(INBOUNDHIST)
    totalratio = totalcount(pickdf,inbounddf)
    feedetail = HIST_bill(YAMATOBILL,totalratio)
    df_orders = YamatoShipment(YAMATODETAIL,PICKHIST)
    FinalFee(feedetail)
    print("------------FINISHED!------------")



