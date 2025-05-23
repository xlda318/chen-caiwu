import pandas as pd

def process_excel(file_path):
    # 读取Excel文件，跳过前6行，将第7行作为标题
    df = pd.read_excel(file_path, header=6)
    
    # 删除第一列
    df.drop(df.columns[0], axis=1, inplace=True)
    
    # 查找"合计："所在行
    for i, row in df.iterrows():
        if "合计：" in str(row.values):
            end_row = i
            break
    
    # 保留标题行至"合计："行
    df = df.iloc[:end_row+1]
    
    # 保存处理后的文件
    df.to_excel(file_path, index=False, engine='openpyxl')
    print(f"文件处理完成，保存路径: {file_path}")

# 使用示例
if __name__ == "__main__":
    process_excel(r"f:\code\数据清洗\1月材料出库单明细表20210207.xlsx")