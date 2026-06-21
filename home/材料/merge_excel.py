import pandas as pd

# 读取两个Excel文件
file1 = r'F:\code\数据清洗\1月材料出库单明细表20210207-西航港.xlsx'
file2 = r'F:\code\数据清洗\1月材料出库单明细表20210207.xlsx'

df1 = pd.read_excel(file1)
df2 = pd.read_excel(file2)

# 只保留两个表都存在的相同列
common_cols = df1.columns.intersection(df2.columns)
df1_common = df1[common_cols]
df2_common = df2[common_cols]

# 合并两个数据框，标题行引用第一个表
merged_df = pd.concat([df1_common, df2_common], ignore_index=True)

# 将合并后的数据框写入新的Excel文件，保存在同一文件夹下
output_file = r'F:\code\数据清洗\合并后的出库单明细表.xlsx'
merged_df.to_excel(output_file, index=False)

print(f"合并完成，结果已保存到 {output_file}")