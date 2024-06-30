import pandas as pd
from ODBC.BPCSquery import BPCSquery
query = """
SELECT 
    E.LORD,
    E.LPROD,
    SUM(E.LQORD - E.LQSHP) AS LORDLEFT,
    E.LNET,
    (SUM(E.LQORD - E.LQSHP) * E.LNET) AS LVALUE,
    COALESCE(STKOH.SUM_STKOH, 0) AS STKOH,
    COALESCE(IML.IONOD, 0) AS IONOD,
    (COALESCE(STKOH.SUM_STKOH, 0) + COALESCE(IML.IONOD, 0)) AS EFF_STOCK
FROM 
    ECL E
LEFT JOIN 
    (
        SELECT 
            WPROD, 
            SUM(WOPB + WRCT - WISS + WADJ) AS SUM_STKOH
        FROM 
            IWI
        WHERE 
            WPROD NOT LIKE 'FA%' 
        GROUP BY 
            WPROD
    ) STKOH 
ON 
    E.LPROD = STKOH.WPROD
LEFT JOIN 
    (
        SELECT 
            IPROD, 
            IONOD
        FROM 
            IIM
    ) IML 
ON 
    E.LPROD = IML.IPROD
WHERE 
    E.LQORD > E.LQSHP
    AND E.LRDTE > 20231030
    AND E.LCUST > 900
GROUP BY 
    E.LORD, E.LPROD, E.LNET, E.LQORD, IML.IONOD, STKOH.SUM_STKOH
    """
df = BPCSquery(query,"JPNPRDF")
print(df)
df.to_csv("output.csv",index=False)