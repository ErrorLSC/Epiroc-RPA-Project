import os
import pandas as pd
from StockSnapShot import stock_by_PLC 

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



folder_path = r"C:\Users\jpeqz\OneDrive - Epiroc\SCX\QRYs\STKOH"
snapfile = r'C:\Users\jpeqz\OneDrive - Epiroc\Python\STOCKSNAP\stockonhandsnap.csv'
stocksnapdf = pd.read_csv(snapfile,index_col='Date')
newestfile = newest_file_in_directory(folder_path)
    
newestdf = stock_by_PLC(newestfile,stocksnapdf)
newestdf.to_csv('STOCKSNAP\stockonhandsnap.csv')
print('Finished!')

    
