import pandas as pd

df = pd.read_excel("TTT10nhap.xlsx")

df["NGÀY SINH"] = df["NGÀY SINH"].astype(str) + '/' + df["Unnamed: 7"].astype(str) + '/'  + df["Unnamed: 8"].astype(str)
df = df.drop(["Unnamed: 7", "Unnamed: 8"], axis=1)

for row in df.itertuples():
    print(row._13)

print(df)
