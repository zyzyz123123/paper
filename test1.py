import os
import re
import shutil
import pandas as pd

# ========= 路径配置 =========
base_dir = "."
excel_path = os.path.join(base_dir, "常用控制变量2000_2024_Ver3.1.xlsx")

# 年份范围：2014～2023 的文件夹
YEAR_START, YEAR_END = 2011, 2024
YEAR_RANGE = range(YEAR_START, YEAR_END + 1)

# 改为匹配 "年年度报告" 或 "年度报告"（但排除"半年度报告"）
REPORT_YEAR_PATTERN = re.compile(r"(\d{4})年(?:年度|度)报告")

# 重复文件夹中要直接删除的关键词（文件名包含即删除）
EXCLUDE_KEYWORDS = ("已取消", "英文版", "英文")

# 文件名末尾日期格式：_YYYY-MM-DD.txt
DATE_PATTERN = re.compile(r"_(\d{4}-\d{2}-\d{2})\.txt$")


def parse_date_from_filename(filename):
    """从文件名解析日期，返回 YYYY-MM-DD 字符串，无法解析则返回 None。"""
    m = DATE_PATTERN.search(filename)
    return m.group(1) if m else None


# ========= 0. 按「年年度报告」前的数字重分类到对应年份文件夹 =========
print("========= 步骤 0：按报告年份重分类 =========")
for year in YEAR_RANGE:
    src_dir = os.path.join(base_dir, str(year))
    if not os.path.isdir(src_dir):
        continue
    # 遍历列表副本，避免目录变化导致异常
    for filename in list(os.listdir(src_dir)):
        if not filename.endswith(".txt"):
            continue
        m = REPORT_YEAR_PATTERN.search(filename)
        if not m:
            continue
        report_year = int(m.group(1))
        if report_year == year:
            continue
        if report_year < YEAR_START or report_year > YEAR_END:
            continue
        target_dir = os.path.join(base_dir, str(report_year))
        os.makedirs(target_dir, exist_ok=True)
        src_path = os.path.join(src_dir, filename)
        dst_path = os.path.join(target_dir, filename)
        # 目标已存在则跳过，避免覆盖
        if os.path.exists(dst_path):
            continue
        shutil.move(src_path, dst_path)
        print(f"  移动: {year}/{filename} -> {report_year}/")

# ========= 1. 读取 Excel，按年份提取制造业 stkcode =========
print("\n========= 步骤 1：读取制造业股票代码（按年份区分）=========")
df = pd.read_excel(excel_path)
# 股票代码转为6位字符串（补零），保证匹配一致性
df["stkcode"] = df["stkcode"].astype(str).str.zfill(6)
# 确保 year 列为整数
df["year"] = df["year"].astype(int)

# 按年份构建制造业股票代码集合：{year: set(stkcode1, stkcode2, ...)}
manufacturing_by_year = {}
for year in YEAR_RANGE:
    year_manu_codes = set(
        df.loc[(df["year"] == year) & (df["ind1"] == "C"), "stkcode"]
    )
    manufacturing_by_year[year] = year_manu_codes
    print(f"  {year} 年制造业股票代码数量：{len(year_manu_codes)}")

# ========= 2. 按年遍历：仅保留制造业+去重+手动文件彻底移出 =========
for year in YEAR_RANGE:
    src_dir = os.path.join(base_dir, str(year))
    if not os.path.isdir(src_dir):
        print(f"\n{year} 年文件夹不存在，跳过...")
        continue

    # 手动判断文件夹（仅保留这一个辅助文件夹）
    manual_dir = os.path.join(base_dir, f"{year}_重复_手动")
    os.makedirs(manual_dir, exist_ok=True)

    print(f"\n========= {year} 年 处理（仅保留制造业+无重复）=========")

    # 初始化统计变量
    del_non_manufacturing = 0  # 删除非制造业文件数
    del_english_count = 0      # 删除含排除关键词的制造业文件数

    # 预处理：清空手动文件夹（避免上次运行遗留干扰，重新按逻辑填充）
    cleared_manual_count = 0
    for filename in list(os.listdir(manual_dir)):
        if filename.endswith(".txt"):
            os.remove(os.path.join(manual_dir, filename))
            cleared_manual_count += 1
    if cleared_manual_count > 0:
        print(f"  清空手动文件夹遗留文件：{cleared_manual_count} 个")
    del_older_count = 0        # 自动删除旧版本制造业文件数
    move_manual_count = 0      # 彻底移出到手动文件夹的制造业文件数

    # 获取该年份的制造业股票代码集合
    manufacturing_codes = manufacturing_by_year.get(year, set())

    # 步骤1：先删除所有非制造业TXT文件，仅保留制造业文件
    all_files = list(os.listdir(src_dir))
    for filename in all_files:
        if not filename.endswith(".txt"):
            continue
        stk_code = filename[:6]
        if stk_code not in manufacturing_codes:
            os.remove(os.path.join(src_dir, filename))
            del_non_manufacturing += 1

    # 步骤2：将剩余制造业TXT文件按股票代码分组
    manu_txt_files = [f for f in os.listdir(src_dir) if f.endswith(".txt")]
    stkcode_file_map = {}
    for filename in manu_txt_files:
        stk_code = filename[:6]
        stkcode_file_map.setdefault(stk_code, []).append(filename)

    # 步骤3：逐股票代码处理制造业重复文件
    for stk_code, file_list in stkcode_file_map.items():
        # 无重复，直接保留，无需处理
        if len(file_list) <= 1:
            continue

        # 子步骤1：删除含「英文版/英文」的制造业文件
        remain_files = []
        for file in file_list:
            if any(kw in file for kw in EXCLUDE_KEYWORDS):
                os.remove(os.path.join(src_dir, file))
                del_english_count += 1
            else:
                remain_files.append(file)
        # 删完英文后无重复，结束该股票代码处理
        if len(remain_files) <= 1:
            continue

        # 子步骤2：解析文件日期，区分可解析/不可解析日期
        file_with_date = [(f, parse_date_from_filename(f)) for f in remain_files]
        valid_date_files = [(f, d) for f, d in file_with_date if d is not None]
        invalid_date_files = [f for f, d in file_with_date if d is None]

        # 子步骤3：无任何可解析日期 → 全部彻底移出到手动文件夹
        if not valid_date_files:
            for file in remain_files:
                src_path = os.path.join(src_dir, file)
                dst_path = os.path.join(manual_dir, file)
                if not os.path.exists(dst_path):
                    shutil.move(src_path, dst_path)
                    move_manual_count += 1
            continue

        # 子步骤4：有可解析日期 → 先移出不可解析日期文件到手动文件夹
        for file in invalid_date_files:
            src_path = os.path.join(src_dir, file)
            dst_path = os.path.join(manual_dir, file)
            if not os.path.exists(dst_path):
                shutil.move(src_path, dst_path)
                move_manual_count += 1

        # 子步骤5：可解析日期文件 → 保留最新，删除旧版本
        latest_date = max(d for _, d in valid_date_files)
        latest_date_files = [f for f, d in valid_date_files if d == latest_date]

        # 仅1个最新日期文件 → 保留该文件，删除其他
        if len(latest_date_files) == 1:
            keep_file = latest_date_files[0]
            for file, _ in valid_date_files:
                if file != keep_file:
                    os.remove(os.path.join(src_dir, file))
                    del_older_count += 1
        # 多个最新日期文件 → 全部移出到手动文件夹
        else:
            for file, _ in valid_date_files:
                src_path = os.path.join(src_dir, file)
                dst_path = os.path.join(manual_dir, file)
                if not os.path.exists(dst_path):
                    shutil.move(src_path, dst_path)
                    move_manual_count += 1

    # ========= 该年份处理结果精准统计 =========
    # 原年份文件夹最终剩余制造业文件数（无重复）
    final_manu_files = [f for f in os.listdir(src_dir) if f.endswith(".txt")]
    # 原年份文件夹最终制造业股票代码数（1:1匹配文件数）
    final_manu_stkcodes = {f[:6] for f in final_manu_files}
    # 手动文件夹待处理制造业文件数
    final_manual_files = [f for f in os.listdir(manual_dir) if f.endswith(".txt")]

    print(f"  📊 {year} 年处理统计：")
    print(f"     ├─ 删除非制造业文件：{del_non_manufacturing} 个")
    print(f"     ├─ 删除含排除关键词文件（已取消/英文）：{del_english_count} 个")
    print(f"     ├─ 制造业文件自动删旧保新：{del_older_count} 个")
    print(f"     ├─ 制造业文件移出到手动文件夹：{move_manual_count} 个")
    print(f"     ├─ 原文件夹最终制造业文件数：{len(final_manu_files)}（无重复）")
    print(f"     ├─ 原文件夹制造业股票代码数：{len(final_manu_stkcodes)}（1:1匹配）")
    print(f"     └─ 手动文件夹待处理制造业文件：{len(final_manual_files)} 个（原文件夹无残留）")

# ========= 最终清理：删除所有冗余中间文件夹 =========
print("\n========= 步骤 3：清理冗余中间文件夹 =========")
del_folder_count = 0
for year in YEAR_RANGE:
    redundant_folders = [
        os.path.join(base_dir, f"{year}_C"),
        os.path.join(base_dir, f"{year}_重复")
    ]
    for folder in redundant_folders:
        if os.path.isdir(folder):
            shutil.rmtree(folder)
            del_folder_count += 1
            print(f"  删除冗余文件夹：{folder}")
if del_folder_count == 0:
    print(f"  无冗余文件夹，无需清理")

# ========= 最终操作提示 =========
print("\n" + "="*65)
print("✅ 所有年份处理完成！完全匹配诉求：仅保留制造业+无重复+手动文件彻底分离")
print("  1. 原年份文件夹（2014~2023）：仅含制造业文件，每个股票代码1个，无任何重复")
print("  2. 非制造业文件：已全部删除，无任何残留")
print("  3. 手动文件夹：仅含制造业待判断文件，原文件夹无任何手动文件残留")
print("\n📌 后续手动操作步骤（仅处理制造业文件）：")
print("  1. 打开「XX_重复_手动」文件夹，筛选需要保留的制造业文件；")
print("  2. 将需要保留的文件移动到对应「XX」年份文件夹（无重复风险）；")
print("  3. 无需保留的制造业文件，直接在手动文件夹删除即可；")
print("  4. 手动处理完成后，手动文件夹可按需保留/删除。")
print("="*65)