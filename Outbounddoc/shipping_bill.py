import pandas as pd
import re
import chardet
import openpyxl as ox
import unicodedata

class ShipmentDirection:
    def __init__(self):
        self.shipments = {}
        self.consolidated_shipment_orders = {}

    def add_shipment(self, picknum, total_orders, tracking_needed_orders=None, special_instructions=None):
        self.shipments[picknum] = {
            'total_orders': total_orders,
            'tracking_needed_orders': tracking_needed_orders,
            'special_instructions': special_instructions
        }
    def update_tracking_needed_orders(self,picknum,tracking_needed_orders):
        self.shipments[picknum]['tracking_needed_orders'] = tracking_needed_orders

    def update_special_instructions(self,picknum,special_instructions):
        self.shipments[picknum]['special_instructions'] = special_instructions

    # consolidated_shipname_ordernum 为一个字典
    def add_consolidated_shipment_orders(self, consolidated_shipname_ordernum):
        self.consolidated_shipment_orders.update(consolidated_shipname_ordernum)

    def get_shipment_by_picktime(self, picknum):
        return self.shipments.get(picknum, None)

    def get_all_shipments(self):
        return self.shipments

    def get_consolidated_shipment_orders(self):
        return self.consolidated_shipment_orders

def detect_encoding(file):
    with open(file, 'rb') as f:
        result = chardet.detect(f.read())
    return result['encoding']

def pickingcsv_loading(csv_list,header_list):
    dataframes = []

    for i, file in enumerate(csv_list):
        encoding = detect_encoding(file)
        df = pd.read_csv(file, encoding=encoding,names=header_list,dtype="str")
        df['SourceFile'] = i + 1  # 增加一列并赋值，从1开始
        dataframes.append(df)

    merged_df = pd.concat(dataframes, ignore_index=True)

    return merged_df

def order_num_count(pickdf):
    order_countdf = pickdf.groupby('SourceFile')['OSONO'].nunique().reset_index()
    
    order_count = order_countdf.set_index('SourceFile')['OSONO'].to_dict()
    #print(order_count)
    return order_count

# Set OMEMO3,4 and 3rd line of SHIPTO as special note
def special_note(pickdf):
    special_note_df = pickdf[['OSONO','OSHAD3','OMEMO1','OMEMO2','OMEMO3','OMEMO4','SourceFile']].drop_duplicates()

    special_note_df['PROFILE'] = special_note_df['OSHAD3'].fillna('').apply(lambda x: ' '.join(re.findall(r'\*(.*?)\*', x)))

    special_note_df['SpecialNote'] = special_note_df['OMEMO1'].fillna('').apply(lambda x: unicodedata.normalize('NFKC',x)) + special_note_df['OMEMO2'].fillna('').apply(lambda x: unicodedata.normalize('NFKC',x)) + special_note_df['OMEMO3'].fillna('').apply(lambda x: unicodedata.normalize('NFKC',x)) + special_note_df['OMEMO4'].fillna('').apply(lambda x: unicodedata.normalize('NFKC',x))

    special_note_df['SpecialNote'] = special_note_df['SpecialNote'].apply(lambda x: ' '.join(re.findall(r'\*(.*?)\*', x)))
    special_note_df['SpecialNote'] = special_note_df['PROFILE'].fillna('') + special_note_df['SpecialNote'].fillna('')

    special_note_df['SpecialNote'] = special_note_df['SpecialNote'].str.replace("送り状要","")
    special_note_df['SpecialNote'] = special_note_df['SpecialNote'].str.replace("同梱不可","")
    special_note_df['SpecialNote'] = special_note_df['SpecialNote'].str.strip()
    special_note_df = special_note_df[special_note_df['SpecialNote'] != '']

    special_note_df['SpecialNote'] = special_note_df['OSONO'] + ":" + special_note_df['SpecialNote']
    
    special_note_df = special_note_df.groupby('SourceFile')['SpecialNote'].apply(list)
    special_note_dict = special_note_df.to_dict()
    #print(special_note_dict)
    return special_note_dict

# Set OMEMO3 as waybill request
def waybill_request(pickdf):
    waybill_request_df = pickdf[['OSONO','OSHAD3','OMEMO1','OMEMO2','OMEMO3','OMEMO4','SourceFile']].drop_duplicates()
    #waybill_request_df = waybill_request_df.dropna(subset = ['OMEMO3'])
    waybill_request_df['WB'] = waybill_request_df['OSHAD3'].fillna('') + waybill_request_df['OMEMO1'].fillna('')+ waybill_request_df['OMEMO2'].fillna('') + waybill_request_df['OMEMO3'].fillna('') + waybill_request_df['OMEMO4'].fillna('')
    waybill_request_df = waybill_request_df[waybill_request_df['WB'].str.contains("送り状要")]
    waybill_request_df = waybill_request_df.groupby('SourceFile')['OSONO']
    waybill_request_dict = waybill_request_df.apply(list).to_dict()
    return waybill_request_dict

def consolidate_shipment(pickdf,name_block_list=None,address_block_list=None,fixed_consolidate_dict=None):
    name_block_pattern = '|'.join(name_block_list)
    address_block_pattern = '|'.join(address_block_list)
    consolidate_df = pickdf[~pickdf['OSHNA1'].str.contains(name_block_pattern)]
    consolidate_df = consolidate_df[~consolidate_df['OSHAD1'].str.contains(address_block_pattern)]
    consolidate_df = consolidate_df[['OSONO','OTELNO','OSHNA1','OSHAD1','OSHAD2','OSHAD3','OMEMO1','OMEMO2','OMEMO3','OMEMO4','SourceFile']].drop_duplicates()

    specific_note = "同梱不可"

    contains_note = consolidate_df.applymap(lambda x: specific_note in x if isinstance(x, str) else False)
    rows_to_drop = contains_note.any(axis=1)
    consolidate_df = consolidate_df.drop(index=consolidate_df[rows_to_drop].index)

    consolidate_df['OTELNO'] = consolidate_df['OTELNO'].str.strip()
    consolidate_df['OTELNO'] = consolidate_df['OTELNO'].str.replace("-","")
    fixed_consolidate_df = consolidate_df[['OSONO','OTELNO']]
    consolidate_order_dict ={}
    for location, phone in fixed_consolidate_dict.items():
        consolidate_order_dict[location] = fixed_consolidate_df[fixed_consolidate_df['OTELNO'] == phone]['OSONO'].to_list()
    
    other_df = consolidate_df[~consolidate_df['OTELNO'].isin(list(fixed_consolidate_dict.values()))]
    other_df.loc[:, 'SHIPADD'] = other_df['OSHAD1'] + other_df['OSHAD2'].fillna('')
    other_df.loc[:, 'SHIPADD'] = other_df.loc[:, 'SHIPADD'].apply(lambda x: unicodedata.normalize('NFKC',x))
    other_df['SHIPADD'] = other_df['SHIPADD'].str.replace(" ","")
    other_df['OSHNA1'] = other_df['OSHNA1'].apply(lambda x: unicodedata.normalize('NFKC',x))
    other_df = other_df[other_df.duplicated(subset=['SHIPADD'],keep=False)]
    other_dict = other_df.groupby('SHIPADD').apply(lambda x: {'OSHNA1': x['OSHNA1'].iloc[0], 'OSONO': x['OSONO'].tolist()}).to_dict()
    other_dict = {v['OSHNA1']: v['OSONO'] for v in other_dict.values()}
    consolidate_order_dict.update(other_dict)
    #print(other_df)
    return consolidate_order_dict

def shipment_fulfillment():
    pass


def template_fulfillment(excel_template,shipment_diretion):
    pass

if __name__ == '__main__':
    csv0 = r"\\ssisjpfs0004\JPN\MRBA\Logistics\CMT Logistics\FromBPCS\DOWNLOADS\Yamato\送信済みデータ\lypl20240723 1100.csv"
    csv1 = r"\\ssisjpfs0004\JPN\MRBA\Logistics\CMT Logistics\FromBPCS\DOWNLOADS\Yamato\送信済みデータ\lypl20240723 1145.csv"
    csv2 = r"\\ssisjpfs0004\JPN\MRBA\Logistics\CMT Logistics\FromBPCS\DOWNLOADS\Yamato\送信済みデータ\lypl20240723 1430.csv"
    csv3 = r"\\ssisjpfs0004\JPN\MRBA\Logistics\CMT Logistics\FromBPCS\DOWNLOADS\Yamato\送信済みデータ\lypl20240723 1530.csv"

    tempcsv = "Outbounddoc\pick0723.csv"
    excel_template = r"C:\Users\jpeqz\OneDrive - Epiroc\Tempfiles\送り状鑑（更新版）.xls"
    header_list = ["OSONO","OSHIP","OTYPE","OCUSPO","OCUSNO","OCUSNA","OCUSA1","OCUSA2","ODATE","OSHNA1","OSHNA2","OSHZIP","OSHAD1","OSHAD2","OSHAD3","OSHATN","OTELNO","OSOLNE","OITMN","OSERN","OLOCN","OIDESC","OQTY","ODDATE","ODTIME","OTRNSP","OMEMO1","OMEMO2","OMEMO3","OMEMO4","OSHIPR","OPGC","OPLC"]
    name_block_list = ["戸髙","鳥形","峩朗"]
    address_block_list = ["高知県吾川郡仁淀川町","峩朗"]
    fixed_consolidate_dict = {"福冈営業所":"0925580621", "大阪営業所":"0727754511","仙台営業所": "0223473755", "DM兵庫":"0795360461","虎乃門千葉": "0436222141"}
    csv_list = [csv0,csv1,csv2,csv3]
    
    shipment_direction = ShipmentDirection()
    #pickdf = pickingcsv_loading(csv_list,header_list)
    #pickdf.to_csv(tempcsv,encoding='UTF_8_sig')
    pickdf = pd.read_csv(tempcsv,encoding='UTF_8_sig',dtype="str")
    ordercount = order_num_count(pickdf)
    for picktime in ordercount:
        shipment_direction.add_shipment(picktime,ordercount[picktime])
    
    waybill_request_dict = waybill_request(pickdf)
    for picktime in waybill_request_dict:
        shipment_direction.update_tracking_needed_orders(picktime,waybill_request_dict[picktime])
    
    special_note_dict = special_note(pickdf)
    for picktime in special_note_dict:
        shipment_direction.update_special_instructions(picktime,special_note_dict[picktime])

    #print(shipment_direction.get_shipment_by_picktime('1'))
    consolidate_shipment_dict = consolidate_shipment(pickdf,name_block_list,address_block_list,fixed_consolidate_dict)
    shipment_direction.add_consolidated_shipment_orders(consolidate_shipment_dict)
    print(shipment_direction.get_all_shipments())
    print(shipment_direction.get_consolidated_shipment_orders())
