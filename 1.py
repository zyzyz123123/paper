import pandas as pd

# 读取数据
df = pd.read_excel('绿色化转型分析结果_final_2014-2023.xlsx')

print('=== 重复数据检查 (以 stkcd + 年份 为键) ===')
print()
print('列名:', df.columns.tolist())
print('数据行数:', len(df))
print()

# 检查重复
duplicates = df[df.duplicated(subset=['stkcd', '年份'], keep=False)]
dup_count = df.duplicated(subset=['stkcd', '年份'], keep=False).sum()

print(f'总数据行数: {len(df)}')
# 推荐写法：外层双引号，内层单引号，无转义，无冲突
print(f"唯一 (stkcd, 年份) 组合数: {df.drop_duplicates(subset=['stkcd', '年份']).shape[0]}")
print(f'重复记录数: {dup_count}')
print()

if dup_count > 0:
    print('=== 重复的记录详情 ===')
    # 按 stkcd 和年份排序显示重复记录
    duplicates_sorted = duplicates.sort_values(['stkcd', '年份'])
    # 检查列名
    cols_to_show = ['stkcd', '年份']
    for col in ['证券名称', '公司简称', '绿色化转型词数', '年报词数']:
        if col in df.columns:
            cols_to_show.append(col)
    print(duplicates_sorted[cols_to_show].to_string())
    print()
    
    # 统计每个重复组合出现的次数
    print('=== 重复组合统计 ===')
    dup_groups = df.groupby(['stkcd', '年份']).size().reset_index(name='count')
    dup_groups = dup_groups[dup_groups['count'] > 1]
    print(f'有 {len(dup_groups)} 个 (stkcd, 年份) 组合出现了重复')
    print(dup_groups.to_string())
else:
    print('没有发现重复数据！')