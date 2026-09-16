"""
mini_pandas_demo.py
------------------------------------------------------
教學用途：用一個簡化版的 MyDataFrame class，
模擬 pandas 的 pd.DataFrame / pd.read_csv / pd.concat / df.loc / df.head()

目的：讓你親手看到「類別（class）」跟「實體（instance）」
      是怎麼被呼叫、怎麼互相搭配的。
------------------------------------------------------
"""

import csv


class MyDataFrame:
    """模擬 pandas.DataFrame 的迷你教學版（這是「模具」本身）"""

    def __init__(self, data, columns=None):
        # __init__ 就是「造出一片餅乾」的過程
        self.columns = columns if columns else [f"col{i}" for i in range(len(data[0]))]
        self.data = data  # 真正存放資料的地方

    # ---------- 模擬 pd.read_csv(...) ----------
    @classmethod
    def read_csv(cls, filepath):
        """
        classmethod：不需要先有實體就能呼叫，專門負責「造出新實體」，
        很像 Java 裡常見的 static factory method，例如：
            DataFrame df = DataFrame.readCsv("data.csv");
        """
        with open(filepath, newline='', encoding='utf-8') as f:
            reader = csv.reader(f)
            rows = list(reader)
        header, body = rows[0], rows[1:]
        return cls(body, columns=header)  # 呼叫 __init__，造出新的 MyDataFrame

    # ---------- 模擬 pd.concat([...]) ----------
    @staticmethod
    def concat(df_list):
        """
        staticmethod：跟 class 沒有直接關聯（不用 cls，也不用 self），
        只是邏輯上放在這個 class 裡比較方便理解、比較好找。
        """
        if not df_list:
            raise ValueError("concat 至少需要一個 DataFrame")
        columns = df_list[0].columns
        merged_data = []
        for df in df_list:          # 你熟悉的 for 迴圈上場！
            merged_data.extend(df.data)
        return MyDataFrame(merged_data, columns=columns)

    # ---------- 模擬 df.head() ----------
    def head(self, n=5):
        """一般的 instance method，一定要透過「已經做好的餅乾」(df) 來呼叫"""
        return MyDataFrame(self.data[:n], columns=self.columns)

    # ---------- 模擬 df.loc[...] ----------
    @property
    def loc(self):
        """
        property：讓 df.loc 看起來像「屬性」，但其實背後偷偷造了一個
        _Locator 物件出來，負責處理 [] 裡面的篩選邏輯。
        這是「組合（composition）」：一個物件裡面藏著另一個物件在做事。
        """
        return _Locator(self)

    def __repr__(self):
        lines = ["\t".join(self.columns)]
        for row in self.data:
            lines.append("\t".join(str(x) for x in row))
        return "\n".join(lines)


class _Locator:
    """專門負責處理 df.loc[條件] 篩選邏輯的小幫手類別"""

    def __init__(self, parent_df):
        self.parent_df = parent_df

    def __getitem__(self, condition):
        # condition 是一個跟 data 等長的 True/False list
        filtered = [row for row, keep in zip(self.parent_df.data, condition) if keep]
        return MyDataFrame(filtered, columns=self.parent_df.columns)


# ================= 使用範例（對應你原本的程式碼）=================
if __name__ == "__main__":
    columnName_list = ["name", "age"]
    data = [
        ["alan", 12],
        ["bill", 13],
        ["catht", 14],
    ]

    # 1. 造出實體：呼叫「類別」pd.DataFrame(...) 的概念
    df = MyDataFrame(data, columns=columnName_list)
    print("=== df ===")
    print(df)

    # 2. 模擬 df["age"] > 12，用你熟悉的 list comprehension（濃縮版 for 迴圈）
    condition = [row[1] > 12 for row in df.data]

    # 3. 呼叫 df.loc[...]：先觸發 property -> 建立 _Locator -> 執行 __getitem__
    df2 = df.loc[condition]
    print("\n=== df2 (age > 12) ===")
    print(df2)

    # 4. 呼叫 df.head()：對「已存在的實體」做操作
    print("\n=== df.head(2) ===")
    print(df.head(2))

    # 5. 呼叫 MyDataFrame.concat([...])：呼叫「類別方法」，把多個實體合併成新實體
    df3 = MyDataFrame.concat([df, df2])
    print("\n=== concat(df, df2) ===")
    print(df3)