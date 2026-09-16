import pandas as pd
# print(pd.__version__)
file_path = "/Users/changting-chia/Desktop/1.2.1 樣表.xlsx"
df = pd.read_excel(file_path)
# print(df.head())

names = ["Alice", "Bob", "Charlie"]
UperrNames= list(map(lambda x: x.upper(), names))
print(UperrNames)

Names_ContainE = list(filter(lambda x: "e" in x, names))
print(Names_ContainE)
