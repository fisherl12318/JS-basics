import pandas as pd

def find_items(column_names_input):
    # 第一行是数据，所以不把它当列名读取
    df = pd.read_excel('data.xlsx', sheet_name='Tabelle2', header=None)

    # 获取第1行数据
    first_row = df.iloc[0]

    # 将输入转换为列名列表（支持空格分隔的字符串或列表）
    if isinstance(column_names_input, str):
        column_names = column_names_input.split()
    else:
        column_names = column_names_input

    # 找到所有列名对应的列索引
    col_indices = []
    for col_name in column_names:
        if col_name not in first_row.values:
            print(f"没有找到列名: {col_name}")
            continue
        col_index = first_row[first_row == col_name].index[0]
        col_indices.append(col_index)

    if not col_indices:
        return []

    # 从第二行开始，要求所有指定列都为"Y"
    mask = pd.Series([False] * len(df.iloc[1:]), index=df.iloc[1:].index)
    for col_idx in col_indices:
        mask = mask | (df.iloc[1:, col_idx] == "Y")

    # 获取满足条件的行
    rows = df.iloc[1:][mask]

    # 返回这些行的第1列值
    items = rows.iloc[:, 1].tolist()
    return items

# 测试
if __name__ == "__main__":
    col = input("请输入要查询的列名: ")
    result = find_items(col)
    print("符合条件的第一列数据:")
    print(result)