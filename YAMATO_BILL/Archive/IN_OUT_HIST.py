import pandas as pd
import numpy as np

def outboundhist(PICKHIST,start_date,end_date):
    pickdf = pd.read_excel(PICKHIST)
    pickdf['S#CDTE'] = pd.to_datetime(pickdf['S#CDTE'], format='%Y%m%d')
    pickdf = pickdf.set_index('S#CDTE')
    pickdf = pickdf.loc[start_date:end_date]
    oversea_count = (pickdf['HCUST']==9125).sum()
    pickdf = pickdf[pickdf['HCUST'] != 9125] 
    pickdf = pickdf.groupby('CXPPLC')['S#PROD'].count()
    pickdf.loc['LDC+'] = oversea_count
    return pickdf

def inboundhist(INBOUNDHIST,start_date,end_date):
    inbounddf = pd.read_excel(INBOUNDHIST)
    inbounddf['TTDTE'] = pd.to_datetime(inbounddf['TTDTE'], format='%Y%m%d')
    inbounddf = inbounddf.set_index('TTDTE')
    inbounddf = inbounddf.loc[start_date:end_date]
    inbounddf = inbounddf.groupby(['CXPPLC','TTYPE'])['TPROD'].count().reset_index()
    inbounddf = inbounddf.rename(columns={'TPROD': 'Total'})
    LDCindex = inbounddf.loc[(inbounddf['CXPPLC'] == 'MRS') & (inbounddf['TTYPE'] == 'H ')].index
    inbounddf.loc[LDCindex,'CXPPLC'] = 'LDC+'
    inbounddf = inbounddf.groupby('CXPPLC')['Total'].sum()
    return inbounddf

def totalcount(pickdf,inbounddf):
    totaldf = pickdf.add(inbounddf, fill_value=0)
    ratio = totaldf.div(totaldf.sum())
    return ratio

def HIST_bill(YAMATOBILL,totalratio):
    billdf = pd.read_excel(YAMATOBILL,sheet_name='page1table2')
    billdf['本体金額'] = billdf['本体金額'].str.replace(' ','').str.replace(',','')
    billdf['本体金額'] = billdf['本体金額'].astype(float)
    billdf = billdf[billdf['項目'] != '宅急便運賃']
    division_fee = billdf.iloc[:4]['本体金額'].sum()
    DC_fee = billdf.iloc[4:]['本体金額'].sum()
    fee_detail = totalratio * division_fee
    fee_detail.loc['7718'] = DC_fee
    return fee_detail

def YamatoShipment(YAMATODETAIL,PICKHIST):
    detaildf = pd.read_excel(YAMATODETAIL,skiprows=1)
    detaildf = detaildf.dropna(subset=['伝票番号'])
    df_nonorders = detaildf[pd.to_numeric(detaildf['ORD#'], errors='coerce').isna()]
    df_orders = detaildf[~pd.to_numeric(detaildf['ORD#'], errors='coerce').isna()]
    pickdf = pd.read_excel(PICKHIST)
    df_orders = df_orders[['ORD#','単価']]
    pickdf = pickdf.drop_duplicates(subset=['S#ORD'])
    pickdf['S#ORD'] = pickdf['S#ORD'].astype(str)
    df_orders['ORD#'] = df_orders['ORD#'].astype(str)
    df_orders = df_orders.set_index('ORD#').join(pickdf.set_index('S#ORD')[['CXPPLC']],how='left')
    df_orders = df_orders.groupby('CXPPLC')['単価'].sum()
    with pd.ExcelWriter('YAMATO_BILL\OUTPUT\Yamato\YamatoSHIPMENT.xlsx') as writer:
        df_nonorders.to_excel(writer,sheet_name='YamatoNONORDER')
        df_orders.to_excel(writer,sheet_name='YamatoORDER')
    return df_orders

def FinalFee(fee_detail,df_orders):
    FinalFee = fee_detail
    FinalFee = np.ceil(FinalFee)
    FinalFee.to_excel('YAMATO_BILL\OUTPUT\Yamato\人件費.xlsx')

if __name__ == "__main__":
    PICKHIST = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\QRYs\OUTBOUND\PICKHIST.xlsx"
    INBOUNDHIST = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\QRYs\INBOUND\ARRIVALYLC.xlsx"
    YAMATOBILL =  r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Bills\Yamato\202404\Yamatotable.xlsx"
    YAMATODETAIL = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Bills\Yamato\202404\【4月度】運賃明細データ.xlsx"
    start_date = '2024-03-16'
    end_date = '2024-04-15'
    pickdf = outboundhist(PICKHIST,start_date,end_date)
    inbounddf = inboundhist(INBOUNDHIST,start_date,end_date)
    totalratio = totalcount(pickdf,inbounddf)
    feedetail = HIST_bill(YAMATOBILL,totalratio)
    df_orders = YamatoShipment(YAMATODETAIL,PICKHIST)
    FinalFee(feedetail,df_orders)


