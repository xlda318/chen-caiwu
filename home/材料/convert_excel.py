import pandas as pd
import os

# 定义要处理的文件列表
files = [
    r"f:\code\数据清洗\1月材料出库单明细表20210207-西航港.xls",
    r"f:\code\数据清洗\1月材料出库单明细表20210207.xls"
]

for file in files:
    try:
        # 读取Excel文件
        df = pd.read_excel(file)
        
        # 生成新的xlsx文件名
        new_file = os.path.splitext(file)[0] + ".xlsx"
        
        # 保存为xlsx格式
        df.to_excel(new_file, index=False)
        
        # 删除原文件
        os.remove(file)
        
        print(f"成功转换并删除: {file} -> {new_file}")
    except Exception as e:
        print(f"处理文件 {file} 时出错: {str(e)}")