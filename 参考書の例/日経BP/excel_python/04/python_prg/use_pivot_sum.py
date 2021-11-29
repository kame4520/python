import pandas as pd

df = pd.read_csv("..\data\ordersList.csv",encoding="cp932",header = 0)
print(df.pivot_table(index="品名",columns="得意先名", values="金額", \
    fill_value=0, margins=True, aggfunc="sum"))