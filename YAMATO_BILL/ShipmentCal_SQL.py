import pandas as pd
from ODBC.BPCSquery import BPCSquery
from CurrentMonth.CurrentMonth import dateperiod

def seinou(pickhist,seinoubill,seinoudetail):
    pickdf = pickhist
    pickdf['S#ORD'] = pickdf['S#ORD'].astype(str)
    pickdf['S#CDTE'] = pickdf['S#CDTE'].astype(str)
    #print(pickdf.dtypes)
    pickdf = pickdf.drop_duplicates(subset=['S#ORD','S#CDTE'])
    pickdf = pickdf.set_index(['S#ORD', 'S#CDTE'])
    seinoubilldf = pd.read_csv(seinoubill,dtype={'原票No.':str},usecols=['原票No.','合計(運賃)'])
    seinoudetaildf = pd.read_excel(seinoudetail,sheet_name='西濃運輸',skiprows=1,usecols=['出荷予定日','お問合せ番号','ORD#'],dtype={'出荷予定日':str,'お問合せ番号':str,'ORD#':str})
    seinoudetaildf.columns = ['出荷予定日','原票No.','ORD#']
    seinoubilldf = seinoubilldf.set_index('原票No.').join(seinoudetaildf.set_index('原票No.')[['ORD#','出荷予定日']],how='left')
    #print(seinoubilldf)
    seinoubilldf = seinoubilldf.join(pickdf['CXPPLC'], on=['ORD#',"出荷予定日"])
    seinoubillnonorderdf = seinoubilldf[seinoubilldf['CXPPLC'].isna()]
    seinoubillorderdf = seinoubilldf[~seinoubilldf['CXPPLC'].isna()]
    print(seinoubilldf)
    seinoubillorderdf = seinoubillorderdf.groupby('CXPPLC')['合計(運賃)'].sum()
    with pd.ExcelWriter('YAMATO_BILL/OUTPUT/SEINOU/Seinou.xlsx') as writer:
        seinoubillnonorderdf.to_excel(writer, sheet_name='NONORDER')
        seinoubillorderdf.to_excel(writer,sheet_name='ORDER')
    #print(seinoubillorderdf)


if __name__ == "__main__":
    #pickhist = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\QRYs\OUTBOUND\PICKHIST.xlsx"
    seinoubill = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Bills\Yamato\202406\seinou\0296496731_20240615_01358.csv"
    seinoudetail = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Bills\Yamato\202406\【6月度】運賃明細データ.xlsx"

    start_date = dateperiod("int")[0]
    end_date = dateperiod("int")[1]

    #start_date = 20240416
    #end_date = 20240515

    PICKHIST_SQL = f"""SELECT P.S#ORD,P.S#PROD,O.HCUST,X.CXPPLC,P.S#CDTE
    FROM JPNLOCF.ESRW P
    LEFT JOIN JPNPRDF.ECHL02 O
    ON P.S#ORD=O.HORD
    LEFT JOIN JPNPRDF.IIML01 I
    ON P.S#PROD=I.IPROD
    LEFT JOIN JPNPRDF.ICXL01 X
    ON I.ICLAS=X.CXCLAS
    WHERE P.S#WHSE='5'
    AND P.S#CDTE BETWEEN {start_date} AND {end_date}
    AND O.HCUST != 9125
    """

    pickhist = BPCSquery(PICKHIST_SQL,"JPNALL")
    seinou(pickhist,seinoubill,seinoudetail)
    print("------------FINISHED!------------")