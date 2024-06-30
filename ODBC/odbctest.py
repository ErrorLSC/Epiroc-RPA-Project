import pyodbc
import pandas as pd

conn_str = (
    "DRIVER={IBM i Access ODBC Driver};"
    "DSN = JPNPRDF;"
    "system = EPISBE20;"
    "Trusted_Connection=yes;"
)

conn = pyodbc.connect(conn_str)

# 创建一个游标对象
#cursor = conn.cursor()
query =  """
SELECT HQPROD,HQPR1,HQQDT
FROM HQTL01
"""

#query = "SELECT WPROD,WMIN,SUM(WOPB+WRCT-WISS+WADJ) AS STKOH FROM IWIL01 GROUP BY WPROD,WMIN"

df = pd.read_sql(query,conn,parse_dates={"HQQDT": {"format": "%Y%m%d"}})
#df = pd.read_sql(query,conn)
def strip_and_convert(x):
    if isinstance(x, str):
        return x.strip()
    elif isinstance(x, float) and x.is_integer():
        return int(x)
    else:
        return x
    
df = df.map(strip_and_convert)
print(df)

#df.to_csv("ITHhistory.csv",index=False)
#print(df)

# 执行一个 SQL 查询
#cursor.execute("SELECT * FROM ")

#print(cursor.description)

#columns = [column[0] for column in cursor.description]

# 获取并打印所有结果
#rows = cursor.fetchall()
#for row in rows:
#    print(row)

#print("Column Names:", columns)

#num_columns = len(cursor.description)

# 关闭游标和连接
#cursor.close()
conn.close()

# 将结果和表头转换为 Pandas DataFrame
#df = pd.DataFrame(rows, columns=columns)


#print(num_columns)

#df = pd.DataFrame(rows,columns=['col1'])

#df_explode = df.explode('col1')
#print(df)