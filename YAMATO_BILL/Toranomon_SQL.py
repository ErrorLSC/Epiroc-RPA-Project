import pandas as pd
from CurrentMonth.CurrentMonth import dateperiod
from ODBC.BPCSquery import BPCSquery

def samepacklength(seinoudetail):
    seinoudetaildf = pd.read_excel(seinoudetail,sheet_name='西濃運輸',dtype=str,skiprows=1)
    ORD_num = seinoudetaildf.shape[1] - 7
    columnlist = []
    for i in range(ORD_num):
        ORDcolumn = f"ORD#.{i}" if i != 0 else "ORD#"
        columnlist.append(ORDcolumn)
    return columnlist

def seinou(seinoubill,seinoudetail,samepack_column_list):
    seinoubilldf = pd.read_csv(seinoubill,dtype={'原票No.':str,'合計(運賃)':float},usecols=['原票No.','合計(運賃)'])
    seinoudetaildf = pd.read_excel(seinoudetail,sheet_name='西濃運輸',dtype=str,skiprows=1)
    columnlist1 = ["出荷予定日","お問合せ番号"]
    columnlist3 = ["お届け先名称１","お届け先住所１"]
    columnlist_1 = columnlist1 + samepack_column_list + columnlist3
    seinoudetaildf = seinoudetaildf[columnlist_1]
    seinoudetaildf = seinoudetaildf.rename(columns={"お問合せ番号":"原票No.","お届け先名称１":"届け先会社名","お届け先住所１":"届け先住所"})
    seinoubilldf = seinoubilldf.set_index('原票No.').join(seinoudetaildf.set_index('原票No.'),how='left').reset_index()
    seinoubilldf = seinoubilldf.rename(columns={"出荷予定日":"日付",'原票No.':'伝票番号','合計(運賃)':'単価'})

    columnlist4 = ["日付","伝票番号"]
    columnlist5 = ["届け先会社名","届け先住所","単価"]
    columnlist_2 = columnlist4 + samepack_column_list + columnlist5
    seinoubilldf = seinoubilldf.reindex(columns=columnlist_2)
    seinoubilldf["日付"]  = pd.to_datetime(seinoubilldf["日付"], format='%Y%m%d')
    return seinoubilldf
    
def yamato(YAMATODETAIL):
    detaildf = pd.read_excel(YAMATODETAIL,skiprows=1,usecols=["日付",'伝票番号','ORD#','単価',"届け先住所","届け先会社名"],dtype={'伝票番号':str,'ORD#':str,"日付":str})
    detaildf = detaildf.dropna(subset=['伝票番号'])
    detaildf = detaildf.reindex(columns=["日付","伝票番号","ORD#","届け先会社名","届け先住所","単価"])
    return detaildf

def total(yamato,seinou,pickhist,samepack_column_list):
    totaldf = pd.concat([yamato,seinou]).reset_index()
    pickdf = pickhist
    pickdf['S#ORD'] = pickdf['S#ORD'].astype(str)
    pickdf["HCPO"] = pickdf["HCPO"].str.strip()  
    pickdf = pickdf.drop(columns=['HCUST','CXPPLC'])
    pickdf = pickdf.rename(columns={"S#ORD":'ORD#','S#CDTE':'日付'}) 
    pickdf = pickdf.groupby(['ORD#','日付']).agg({'HCPO':'first'}).reset_index()
    is_tora_dict = {}
    for column_name in samepack_column_list:
        is_tora_dict[column_name] = totaldf[column_name].isin(pickdf["ORD#"])
    is_tora_df = pd.DataFrame(is_tora_dict)
    is_tora_df['is_tora'] = is_tora_df.sum(axis=1)
    totaldf["is_tora"] = is_tora_df['is_tora']
    totaldf = totaldf[totaldf["is_tora"] > 0]
    totaldf["日付"]=pd.to_datetime(totaldf["日付"]).dt.normalize()

    totaldf['ORD#'] = totaldf[samepack_column_list].apply(lambda x: list(x.values), axis=1) #将每行的订单号总和为一个列表
    totaldf = totaldf.explode("ORD#") #将列表爆炸成单行
    totaldf = totaldf.dropna(subset=['ORD#'],how='any')
    columns_to_drop = ['is_tora'] + samepack_column_list[1:]
    totaldf = totaldf.drop(columns=columns_to_drop)
    totaldf = pd.merge(totaldf,pickdf[['ORD#','HCPO','日付']],on=["ORD#","日付"],how='left')
    totaldf = totaldf.sort_values(by='日付')
    mask = totaldf.duplicated(subset=['伝票番号', '単価'], keep='first') #找出爆炸出的重复项
    totaldf.loc[mask, '単価'] = 0 #将重复之后的第二项以后的价格设置为0
    totaldf = totaldf.reindex(columns=["日付","伝票番号","ORD#",'HCPO',"届け先会社名","届け先住所","単価"])
    totaldf = totaldf.drop(columns=['伝票番号'])
    totaldf['日付'] = totaldf['日付'].astype(str)
    price_total = totaldf['単価'].sum(skipna=True)
    new_row = pd.DataFrame({'単価': [price_total], '日付': ['合計']})
    totaldf = pd.concat([totaldf,new_row],axis=0, ignore_index=True)
    totaldf["届け先会社名"] = totaldf["届け先会社名"].str.replace("虎ノ門","虎乃門")
    totaldf["届け先住所"] = totaldf["届け先住所"].str.replace("２－３－２","2-3-2")
    date = dateperiod("str")
    datestr = "期間："+ date[0] + "～" + date[1] + "間"

    with pd.ExcelWriter('YAMATO_BILL/OUTPUT/TORANOMON/Toranomon_Shipment.xlsx', engine='xlsxwriter') as writer:
        totaldf.to_excel(writer,index=False,startrow = 4)
        
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        currencyformat = workbook.add_format({'num_format': '¥#,##0'})
        worksheet.set_column('A:A',12)
        worksheet.set_column('D:D',40)
        worksheet.set_column('E:E',65)
        worksheet.set_column('F:F',None, currencyformat)

        worksheet.write('A1',"虎乃門建設機械株式会社　様",workbook.add_format({'bold': True}))
        worksheet.write('A4',datestr)

if __name__ == "__main__":
    seinoubill = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Bills\Yamato\202407\seinou\0296496731_20240715_01337.csv"
    detail = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Bills\Yamato\202407\【7月度】運賃明細データ.xlsx"

    start_date = dateperiod("int")[0]
    end_date = dateperiod("int")[1]

    #start_date = 20240416
    #end_date = 20240515

    PICKHIST_SQL = f"""SELECT P.S#ORD,P.S#PROD,O.HCUST,X.CXPPLC,O.HCPO,P.S#CDTE
    FROM JPNLOCF.ESRW P
    LEFT JOIN JPNPRDF.ECHL02 O
    ON P.S#ORD=O.HORD
    LEFT JOIN JPNPRDF.IIML01 I
    ON P.S#PROD=I.IPROD
    LEFT JOIN JPNPRDF.ICXL01 X
    ON I.ICLAS=X.CXCLAS
    WHERE P.S#WHSE='5'
    AND P.S#CDTE BETWEEN {start_date} AND {end_date}
    AND O.HCUST = 901100
    AND X.CXPPLC = 'MRS'
    """

    pickhist = BPCSquery(PICKHIST_SQL,"JPNALL",datecolumnlist=['S#CDTE'])

    samepack_column_list = samepacklength(detail)
    seinoubillorderdf = seinou(seinoubill,detail,samepack_column_list)
    yamatodf = yamato(detail)
    total(yamatodf,seinoubillorderdf,pickhist,samepack_column_list)
    print('------Finished!------')