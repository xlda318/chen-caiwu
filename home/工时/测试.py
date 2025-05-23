import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
df = pd.read_excel(r'D:\code\家里电脑\工时\1月导管厂登记表.xlsx')
print(df.shape)
print(df.columns)
print(df.head())

