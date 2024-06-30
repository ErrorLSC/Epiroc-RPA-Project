import pandas as pd

def seinou(pickhist,seinoubill,seinoudetail):
    pickdf = pd.read_excel(pickhist,usecols=["S#ORD","CXPPLC"],dtype={"S#ORD":str})
    pickdf = pickdf.drop_duplicates(subset=['S#ORD'])
    seinoubilldf = pd.read_csv(seinoubill,dtype={'原票No.':str},usecols=['原票No.','合計(運賃)'])
    seinoudetaildf = pd.read_excel(seinoudetail,sheet_name='西濃運輸',skiprows=1,usecols=['お問合せ番号','ORD#'],dtype={'お問合せ番号':str,'ORD#':str})
    seinoudetaildf.columns = ['原票No.','ORD#']
    seinoubilldf = seinoubilldf.set_index('原票No.').join(seinoudetaildf.set_index('原票No.')[['ORD#']],how='left')
    seinoubilldf = seinoubilldf.join(pickdf.set_index('S#ORD'), on='ORD#')
    seinoubillnonorderdf = seinoubilldf[seinoubilldf['CXPPLC'].isna()]
    seinoubillorderdf = seinoubilldf[~seinoubilldf['CXPPLC'].isna()]
    #print(seinoubillorderdf)
    seinoubillorderdf = seinoubillorderdf.groupby('CXPPLC')['合計(運賃)'].sum()
    with pd.ExcelWriter('YAMATO_BILL\OUTPUT\SEINOU\Seinou.xlsx') as writer:
        seinoubillnonorderdf.to_excel(writer, sheet_name='NONORDER')
        seinoubillorderdf.to_excel(writer,sheet_name='ORDER')
    #print(seinoubillorderdf)


if __name__ == "__main__":
    pickhist = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\QRYs\OUTBOUND\PICKHIST.xlsx"
    seinoubill = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Bills\Yamato\202404\seinou\0296496731_20240415_01359.csv"
    seinoudetail = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Bills\Yamato\202404\【4月度】運賃明細データ.xlsx"

    seinou(pickhist,seinoubill,seinoudetail)