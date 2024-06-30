import pandas as pd
from datetime import datetime, timedelta
import xlsxwriter

def seinou(seinoubill,seinoudetail):
    seinoubilldf = pd.read_csv(seinoubill,dtype={'原票No.':str,'合計(運賃)':float},usecols=['原票No.','合計(運賃)'])
    seinoudetaildf = pd.read_excel(seinoudetail,sheet_name='西濃運輸',dtype=str,skiprows=1)
    seinoudetaildf = seinoudetaildf[["出荷予定日","お問合せ番号","ORD#","ORD#.1","ORD#.2","ORD#.3","ORD#.4","お届け先名称１","お届け先住所１"]]
    seinoudetaildf = seinoudetaildf.rename(columns={"お問合せ番号":"原票No.","お届け先名称１":"届け先会社名","お届け先住所１":"届け先住所"})
    #seinoudetaildf[['ORD#.1', 'ORD#.2', 'ORD#.3']] = seinoudetaildf[['ORD#.1', 'ORD#.2', 'ORD#.3']].fillna('')
    #seinoudetaildf['ORD#'] = seinoudetaildf['ORD#'].astype(str) + " " +seinoudetaildf['ORD#.1'] + " "+ seinoudetaildf['ORD#.2'] + " "+seinoudetaildf['ORD#.3']
    #seinoudetaildf = seinoudetaildf.drop(columns=['ORD#.1','ORD#.2','ORD#.3'])
    seinoubilldf = seinoubilldf.set_index('原票No.').join(seinoudetaildf.set_index('原票No.'),how='left').reset_index()
    #seinoubillorderdf = seinoubilldf[~seinoubilldf['ORD#'].isna()].reset_index()
    seinoubilldf = seinoubilldf.rename(columns={"出荷予定日":"日付",'原票No.':'伝票番号','合計(運賃)':'単価'})
    seinoubilldf = seinoubilldf.reindex(columns=["日付","伝票番号","ORD#",'ORD#.1', 'ORD#.2', 'ORD#.3','ORD#.4',"届け先会社名","届け先住所","単価"])
    seinoubilldf["日付"]  = pd.to_datetime(seinoubilldf["日付"], format='%Y-%m-%d %H:%M:%S')
    #print(seinoubilldf)
    return seinoubilldf
    
def yamato(YAMATODETAIL):
    detaildf = pd.read_excel(YAMATODETAIL,skiprows=1,usecols=["日付",'伝票番号','ORD#','単価',"届け先住所","届け先会社名"],dtype={'伝票番号':str,'ORD#':str,"日付":str})
    detaildf = detaildf.dropna(subset=['伝票番号'])
    detaildf = detaildf.reindex(columns=["日付","伝票番号","ORD#","届け先会社名","届け先住所","単価"])
    #print(detaildf)
    return detaildf

def total(yamato,seinou,pickhist):
    totaldf = pd.concat([yamato,seinou]).reset_index()
    #print(totaldf)
    pickdf = pd.read_excel(pickhist,usecols=["S#ORD","CXPPLC","HCUST","S#CDTE","HCPO"],dtype={"S#ORD":str})
    pickdf["HCPO"] = pickdf["HCPO"].str.strip()  
    pickdf = pickdf[pickdf['CXPPLC']=="MRS"]
    pickdf = pickdf[pickdf['HCUST']==901100]
    pickdf = pickdf.drop(columns=['HCUST','CXPPLC'])
    pickdf = pickdf.rename(columns={"S#ORD":'ORD#'}) 
    pickdf = pickdf.groupby('ORD#').agg({'HCPO':'first'}).reset_index()
    is_tora1 = totaldf["ORD#"].isin(pickdf["ORD#"])
    is_tora2 = totaldf["ORD#.1"].isin(pickdf["ORD#"])
    is_tora3 = totaldf["ORD#.2"].isin(pickdf["ORD#"])
    is_tora4 = totaldf["ORD#.3"].isin(pickdf["ORD#"])
    is_tora5 = totaldf["ORD#.4"].isin(pickdf["ORD#"])
    is_tora = sum([is_tora1,is_tora2,is_tora3,is_tora4,is_tora5])
    
    totaldf["is_tora"] = is_tora
    totaldf = totaldf[totaldf["is_tora"] > 0]
    totaldf["日付"]=pd.to_datetime(totaldf["日付"]).dt.normalize()
    #totaldf['ORD#.1'] =  totaldf['ORD#.1'].fillna("")
    #totaldf['ORD#.2'] =  totaldf['ORD#.2'].fillna("")
    #totaldf['ORD#.3'] =  totaldf['ORD#.3'].fillna("")
    totaldf['ORD#'] = totaldf[['ORD#', 'ORD#.1', 'ORD#.2', 'ORD#.3','ORD#.4']].apply(lambda x: [x[0], x[1], x[2], x[3],x[4]], axis=1) #将每行的订单号总和为一个列表
    totaldf = totaldf.explode("ORD#") #将列表爆炸成单行
    totaldf = totaldf.dropna(subset=['ORD#'],how='any')
    totaldf = totaldf.drop(columns=['is_tora','ORD#.1', 'ORD#.2', 'ORD#.3','ORD#.4'])
    totaldf = pd.merge(totaldf,pickdf[['ORD#','HCPO']],on="ORD#",how='left')
    totaldf = totaldf.sort_values(by='日付')
    mask = totaldf.duplicated(subset=['伝票番号', '単価'], keep='first') #找出爆炸出的重复项
    totaldf.loc[mask, '単価'] = 0 #将重复之后的第二项以后的价格设置为0
    totaldf = totaldf.reindex(columns=["日付","伝票番号","ORD#",'HCPO',"届け先会社名","届け先住所","単価"])
    totaldf = totaldf.drop(columns=['伝票番号'])
    totaldf['日付'] = totaldf['日付'].astype(str)
    #totaldf['日付'] = totaldf['日付'].str[:-]
    price_total = totaldf['単価'].sum(skipna=True)
    #totaldf["日付"]=pd.to_datetime(totaldf["日付"]).dt.date()
    totaldf = totaldf.append({'単価': price_total,'日付':'合計'}, ignore_index=True)
    totaldf["届け先会社名"] = totaldf["届け先会社名"].str.replace("虎ノ門","虎乃門")
    totaldf["届け先住所"] = totaldf["届け先住所"].str.replace("２－３－２","2-3-2")
    #totaldf['単価'] = totaldf['単価'].replace(0,"")
    #print(mask)
    #print(totaldf["日付"])
    #finaldf = pd.merge(pickdf,totaldf, on='ORD#')
    #totaldf = totaldf.groupby(['日付','伝票番号']).agg({'単価':'first','ORD#':list,'HCPO':list,'届け先会社名':'first','届け先住所':'first'}).reset_index()
    #finaldf = finaldf.drop_duplicates(subset=['伝票番号','ORD#'])
    #sum = finaldf['単価'].sum()
    #print(totaldf)
    date = dateperiod()
    datestr = "期間："+ date[0] + "～" + date[1] + "間"

    with pd.ExcelWriter('YAMATO_BILL\OUTPUT\TORANOMON\Toranomon_Shipment.xlsx', engine='xlsxwriter') as writer:
        totaldf.to_excel(writer,index=False,startrow = 4)
        
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        currencyformat = workbook.add_format({'num_format': '¥#,##0'})
        worksheet.set_column('A:A', 12)
        worksheet.set_column('D:D',40)
        worksheet.set_column('E:E',65)
        #worksheet.autofit_column(3)
        #worksheet.autofit_column(4)
        worksheet.set_column('F:F', None, currencyformat)

        worksheet.write('A1',"虎乃門建設機械株式会社　様",workbook.add_format({'bold': True}))
        worksheet.write('A4',datestr)



def dateperiod():
    current_date = datetime.now()
    # 计算上个月的年份和月份
    last_month_year = current_date.year if current_date.month != 1 else current_date.year - 1
    last_month_month = current_date.month - 1 if current_date.month != 1 else 12

    # 计算上个月的日期
    last_month_date = datetime(last_month_year, last_month_month, 16)

    this_month_date = datetime(current_date.year, current_date.month, 15)

    last_month_date_str = last_month_date.strftime('%Y年%m月%d日')
    this_month_date_str = this_month_date.strftime('%Y年%m月%d日')

    #print("上个月16号的日期:", last_month_date_str)
    #print("这个月15号的日期:", this_month_date_str)

    return (last_month_date_str,this_month_date_str)

if __name__ == "__main__":
    pickhist = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\QRYs\OUTBOUND\PICKHIST.xlsx"
    seinoubill = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Bills\Yamato\202404\seinou\0296496731_20240415_01359.csv"
    detail = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\Bills\Yamato\202404\【4月度】運賃明細データ.xlsx"

    seinoubillorderdf = seinou(seinoubill,detail)
    yamatodf = yamato(detail)
    total(yamatodf,seinoubillorderdf,pickhist)
    dateperiod()
    print('Finished!')