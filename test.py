import os
import pandas as pd

# ====== 路径配置 ======
txt_dir = r"./2015 2"  # 存放 txt 文件的文件夹路径
excel_path = r"./常用控制变量2000_2024_Ver3.1.xlsx"

# 从文件夹名提取年份（取前4位数字）
folder_name = os.path.basename(txt_dir)
year = int(folder_name[:4])

# ====== 1. 读取 Excel ======
df = pd.read_excel(excel_path)

# 确保 stkcode 是 6 位字符串，year 是整数
df['stkcode'] = df['stkcode'].astype(str).str.zfill(6)
df['year'] = df['year'].astype(int)

# ====== 2. 筛选该年份的制造业 ======
manufacturing_codes = set(
    df.loc[(df['year'] == year) & (df['ind1'] == 'C'), 'stkcode']
)
print(f"Excel 中 {year} 年制造业股票代码数量：{len(manufacturing_codes)}")

# ====== 3. 读取 txt 文件名 ======
txt_codes = set()

for filename in os.listdir(txt_dir):
    if filename.endswith('.txt'):
        code = filename[:6]  # 文件名前 6 位
        txt_codes.add(code)

# ====== 4. 匹配制造业 ======
manufacturing_txt_codes = txt_codes & manufacturing_codes

# ====== 5. 输出结果 ======
print("制造业公司数量：", len(manufacturing_txt_codes))
print("制造业公司代码示例：", list(manufacturing_txt_codes)[:10])
