import os
import pandas as pd

def stock_by_PLC(stockfile,snapdf):
    stockdf = pd.read_excel(stockfile,header=4)
    stockdf_CC = stockdf[stockdf['Whs'] != 'N1']
    stockdf_CC = stockdf_CC.groupby('CODE')['STOCK VALUE'].sum().reset_index()
    row_name = stockfile[-14:-4]
    for index, row in stockdf_CC.iterrows():
        snapdf.loc[row_name, row['CODE']] = row['STOCK VALUE']
    #snapdf = snapdf.rename_axis('Date')
    return snapdf

if __name__ == "__main__":
    folder_path = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\QRYs\STKOH"
    snapdf = pd.DataFrame(columns=['CTS','HAT','MRS','RDT','RGU','SED','SMT'])
    snapdf = snapdf.rename_axis('Date')
    for root,dirs,files in os.walk(folder_path):
        if root == folder_path:
            for file in files:
                stockfile = os.path.join(root,file)
                snapdf = stock_by_PLC(stockfile,snapdf)
    snapdf.to_csv('stockonhandsnap.csv')
    print("Finished!")
