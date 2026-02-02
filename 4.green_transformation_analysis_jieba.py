#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''
@Project ：Annualreport_tools
@File    ：4.green_transformation_analysis_jieba.py
@IDE     ：PyCharm
@Author  ：lingxiaotian
@Date    ：2026/02/01
@Description: 绿色化转型关键词分析脚本（jieba分词版本）
    - 读取113个绿色化转型关键词
    - 使用jieba分词统计年报中关键词词数和年报总词数
    - 计算绿色化转型词频 = 绿色化转型词数 / 年报词数
    - 计算 log后的绿色化转型词频 = ln(词频 + 1)
    - 与公司元数据合并，输出完整表格
    
输出列：
    证券代码, 证券名称, stkcd, 年份,
    绿色化转型词数, 年报词数, 绿色化转型词频, log后的绿色化转型词频,
    行业代码, 行业名称, 省份代码, 省份名称, 城市代码, 城市名称
'''

from __future__ import annotations

import logging
import math
import os
import re
from collections import Counter
from dataclasses import dataclass
from multiprocessing import Pool, cpu_count
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import jieba
import pandas as pd

# 日志配置
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)


@dataclass
class GreenAnalysisConfig:
    """绿色化转型分析配置类。"""
    txt_folder: str  # TXT年报文件夹路径
    keywords_file: str  # 关键词Excel文件路径
    company_info_file: str  # 公司元数据Excel文件路径
    output_file: str  # 输出文件路径
    start_year: int = 2012  # 起始年份
    end_year: int = 2023  # 结束年份
    processes: Optional[int] = None  # 进程数，None表示自动
    chunk_size: int = 100  # 增量保存间隔


class KeywordLoader:
    """关键词加载器。"""
    
    @staticmethod
    def load_from_excel(file_path: str, column_name: str = None) -> List[str]:
        """从Excel文件加载关键词列表。
        
        Args:
            file_path: Excel文件路径
            column_name: 关键词所在列名，如果为None则使用第一列
            
        Returns:
            关键词列表
        """
        try:
            # 先尝试读取，判断第一行是否为标题
            df = pd.read_excel(file_path, header=None)
            first_cell = str(df.iloc[0, 0]).strip() if not df.empty else ""
            
            # 如果第一行看起来像标题（包含"关键词"、"keyword"等），则跳过第一行
            header_keywords = ['关键词', '关键字', 'keyword', '词汇', '词语']
            if any(h in first_cell.lower() for h in header_keywords):
                df = df.iloc[1:]  # 跳过标题行
            
            if column_name:
                # 如果指定了列名，重新读取带header的版本
                df = pd.read_excel(file_path)
                keywords = df[column_name].dropna().astype(str).tolist()
            else:
                # 使用第一列
                keywords = df.iloc[:, 0].dropna().astype(str).tolist()
            
            # 清理关键词（去除空白）
            keywords = [kw.strip() for kw in keywords if kw.strip()]
            
            logging.info(f"成功加载 {len(keywords)} 个关键词")
            return keywords
            
        except FileNotFoundError:
            logging.error(f"关键词文件不存在: {file_path}")
            raise
        except Exception as e:
            logging.error(f"加载关键词失败: {e}")
            raise


class CompanyInfoLoader:
    """公司元数据加载器。"""
    
    # 列名映射（支持多种常见命名）
    COLUMN_MAPPING = {
        'stock_code': ['证券代码', '股票代码', 'stkcd', 'code', 'Stkcd', 'STKCD', 'Symbol'],
        'stock_name': ['证券简称', '公司简称', '股票简称', 'name', 'Stknm', 'ShortName'],
        'industry_code': ['行业代码', 'IndustryCode', 'industry_code', 'IndCd'],
        'industry_name': ['行业名称', 'IndustryName', 'industry_name', 'IndNm', 'Industry'],
        'province': ['所属省份', '省份', 'Province', 'province'],
        'province_code': ['所属省份代码', '省份代码', 'ProvinceCode', 'province_code'],
        'city': ['所属城市', '城市', 'City', 'city'],
        'city_code': ['所属城市代码', '城市代码', 'CityCode', 'city_code'],
    }
    
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.df = None
        self._load_data()
    
    def _find_column(self, df: pd.DataFrame, possible_names: List[str]) -> Optional[str]:
        """在DataFrame中查找匹配的列名。"""
        for name in possible_names:
            if name in df.columns:
                return name
        return None
    
    def _load_data(self) -> None:
        """加载公司元数据。"""
        try:
            self.df = pd.read_excel(self.file_path)
            logging.info(f"成功加载公司元数据: {len(self.df)} 条记录")
            logging.info(f"可用列: {list(self.df.columns)}")
            
            # 标准化股票代码列
            code_col = self._find_column(self.df, self.COLUMN_MAPPING['stock_code'])
            if code_col:
                # 确保股票代码为6位字符串
                self.df['_stock_code'] = self.df[code_col].astype(str).str.zfill(6)
            else:
                logging.warning("未找到股票代码列，请检查公司元数据文件")
                
        except FileNotFoundError:
            logging.error(f"公司元数据文件不存在: {self.file_path}")
            raise
        except Exception as e:
            logging.error(f"加载公司元数据失败: {e}")
            raise
    
    def get_company_info(self, stock_code: str) -> Dict[str, str]:
        """获取指定股票代码的公司信息。
        
        Args:
            stock_code: 股票代码（6位）
            
        Returns:
            包含公司信息的字典
        """
        result = {
            'industry_code': '',
            'industry_name': '',
            'province': '',
            'province_code': '',
            'city': '',
            'city_code': '',
        }
        
        if self.df is None or '_stock_code' not in self.df.columns:
            return result
        
        # 标准化输入的股票代码
        stock_code = str(stock_code).zfill(6)
        
        # 查找匹配记录
        matches = self.df[self.df['_stock_code'] == stock_code]
        
        if matches.empty:
            return result
        
        row = matches.iloc[0]
        
        # 提取各字段
        for field, possible_names in self.COLUMN_MAPPING.items():
            if field in ['stock_code', 'stock_name']:
                continue
            col = self._find_column(self.df, possible_names)
            if col and pd.notna(row[col]):
                result[field] = str(row[col])
        
        return result


# 全局变量，用于多进程共享关键词集合
_global_keywords_set = None


def init_worker(keywords: List[str]):
    """初始化工作进程，加载关键词到jieba词典。"""
    global _global_keywords_set
    _global_keywords_set = set(keywords)
    
    # 将关键词添加到jieba词典，提高分词准确率
    for word in keywords:
        jieba.add_word(word)


def analyze_single_file_jieba(file_path: str) -> Optional[Dict]:
    """使用jieba分词分析单个TXT文件的关键词。
    
    Args:
        file_path: 文件路径
        
    Returns:
        分析结果字典
    """
    global _global_keywords_set
    
    try:
        # 解析文件名：000001_平安银行_2023.txt
        filename = os.path.basename(file_path)
        match = re.match(r'^(\d{6})_(.+?)_(\d{4})\.txt$', filename)
        
        if not match:
            logging.warning(f"文件名格式不符: {filename}")
            return None
        
        stock_code = match.group(1)
        stock_name = match.group(2)
        year = match.group(3)
        
        # 读取文件内容
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            with open(file_path, 'r', encoding='gbk', errors='ignore') as f:
                content = f.read()
        
        if not content.strip():
            logging.warning(f"文件内容为空: {filename}")
            return None
        
        # 使用jieba分词
        words = list(jieba.cut(content))
        
        # 过滤掉空白词和标点符号，只保留有意义的词
        # 中文词：至少包含一个中文字符
        chinese_pattern = re.compile(r'[\u4e00-\u9fff]')
        meaningful_words = [w for w in words if chinese_pattern.search(w)]
        
        # 统计年报总词数（使用中文字符数，与字符串匹配版保持一致）
        total_word_count = len(chinese_pattern.findall(content))
        
        # 使用Counter统计词频
        word_counts = Counter(meaningful_words)
        
        # 统计所有关键词的词频并求和
        green_word_count = sum(word_counts.get(kw, 0) for kw in _global_keywords_set)
        
        # 计算绿色化转型词频 = 绿色化转型词数 / 年报词数
        if total_word_count > 0:
            green_freq = green_word_count / total_word_count
        else:
            green_freq = 0.0
        
        # 计算 log 后的绿色化转型词频 = ln(词频 + 1)
        green_freq_log = math.log(green_freq + 1)
        
        return {
            'stock_code': stock_code,
            'stock_name': stock_name,
            'year': year,
            'green_word_count': green_word_count,
            'total_word_count': total_word_count,
            'green_freq': green_freq,
            'green_freq_log': green_freq_log,
        }
        
    except Exception as e:
        logging.error(f"分析文件失败 {file_path}: {e}")
        return None


class GreenTransformationAnalyzerJieba:
    """绿色化转型分析器主类（jieba版本）。"""
    
    def __init__(self, config: GreenAnalysisConfig):
        self.config = config
        self.keywords = []
        self.company_info_loader = None
    
    def _load_keywords(self) -> None:
        """加载关键词。"""
        self.keywords = KeywordLoader.load_from_excel(self.config.keywords_file)
        logging.info(f"关键词列表前10个: {self.keywords[:10]}")
    
    def _load_company_info(self) -> None:
        """加载公司元数据。"""
        if os.path.exists(self.config.company_info_file):
            self.company_info_loader = CompanyInfoLoader(self.config.company_info_file)
        else:
            logging.warning(f"公司元数据文件不存在: {self.config.company_info_file}")
            logging.warning("将只输出基础词频数据，不包含行业/地区信息")
    
    def _collect_txt_files(self) -> List[str]:
        """收集所有符合条件的TXT文件。"""
        txt_files = []
        
        for root, dirs, files in os.walk(self.config.txt_folder):
            for filename in files:
                if not filename.endswith('.txt'):
                    continue
                
                # 检查年份范围
                match = re.search(r'_(\d{4})\.txt$', filename)
                if match:
                    year = int(match.group(1))
                    if year < self.config.start_year or year > self.config.end_year:
                        continue
                
                txt_files.append(os.path.join(root, filename))
        
        logging.info(f"找到 {len(txt_files)} 个符合条件的TXT文件")
        return txt_files
    
    def _build_output_dataframe(self, results: List[Dict]) -> pd.DataFrame:
        """构建输出DataFrame。"""
        rows = []
        
        for result in results:
            if result is None:
                continue
            
            stock_code = result['stock_code']
            
            # 获取公司元数据
            if self.company_info_loader:
                info = self.company_info_loader.get_company_info(stock_code)
            else:
                info = {
                    'industry_code': '',
                    'industry_name': '',
                    'province': '',
                    'province_code': '',
                    'city': '',
                    'city_code': '',
                }
            
            row = {
                '证券代码': stock_code,
                '证券名称': result['stock_name'],
                'stkcd': int(stock_code),
                '年份': int(result['year']),
                '绿色化转型词数': result['green_word_count'],
                '年报词数': result['total_word_count'],
                '绿色化转型词频': round(result['green_freq'], 10),
                'log后的绿色化转型词频': round(result['green_freq_log'], 10),
                '行业代码': info['industry_code'],
                '行业名称': info['industry_name'],
                '省份代码': info['province_code'],
                '省份名称': info['province'],
                '城市代码': info['city_code'],
                '城市名称': info['city'],
            }
            rows.append(row)
        
        df = pd.DataFrame(rows)
        
        # 按股票代码和年份排序
        if not df.empty:
            df = df.sort_values(['stkcd', '年份']).reset_index(drop=True)
        
        return df
    
    def run(self) -> None:
        """执行分析流程。"""
        logging.info("=" * 60)
        logging.info("绿色化转型关键词分析启动（jieba分词版本）")
        logging.info(f"年份范围: {self.config.start_year} - {self.config.end_year}")
        logging.info(f"TXT文件夹: {self.config.txt_folder}")
        logging.info(f"关键词文件: {self.config.keywords_file}")
        logging.info(f"公司元数据: {self.config.company_info_file}")
        logging.info("=" * 60)
        
        # 加载关键词
        self._load_keywords()
        
        # 加载公司元数据
        self._load_company_info()
        
        # 收集TXT文件
        txt_files = self._collect_txt_files()
        
        if not txt_files:
            logging.error("未找到符合条件的TXT文件")
            return
        
        # 多进程分析
        worker_count = self.config.processes or min(cpu_count(), len(txt_files))
        logging.info(f"使用 {worker_count} 个进程处理 {len(txt_files)} 个文件")
        
        results = []
        with Pool(
            processes=worker_count,
            initializer=init_worker,
            initargs=(self.keywords,)
        ) as pool:
            for i, result in enumerate(pool.imap_unordered(analyze_single_file_jieba, txt_files)):
                results.append(result)
                
                # 显示进度
                progress = (i + 1) / len(txt_files) * 100
                print(f"\r分析进度: {i + 1}/{len(txt_files)} ({progress:.1f}%)", end='', flush=True)
        
        print()  # 换行
        
        # 过滤空结果
        valid_results = [r for r in results if r is not None]
        logging.info(f"成功分析 {len(valid_results)} 个文件")
        
        # 构建输出DataFrame
        df = self._build_output_dataframe(valid_results)
        
        # 保存结果
        output_path = Path(self.config.output_file)
        
        if output_path.suffix == '.xlsx':
            df.to_excel(output_path, index=False, engine='openpyxl')
        else:
            df.to_excel(output_path.with_suffix('.xlsx'), index=False, engine='openpyxl')
        
        logging.info("=" * 60)
        logging.info(f"分析完成！")
        logging.info(f"总记录数: {len(df)}")
        logging.info(f"输出文件: {output_path}")
        logging.info("=" * 60)
        
        # 显示数据预览
        if not df.empty:
            print("\n数据预览（前5行）:")
            print(df.head().to_string())


if __name__ == '__main__':
    # ==================== 配置区域 ====================
    
    # TXT年报文件夹路径（包含按年份组织的子文件夹）
    # 文件命名格式: 000001_平安银行_2023.txt
    TXT_FOLDER = "年报文件"
    
    # 113个绿色化转型关键词Excel文件路径
    KEYWORDS_FILE = "113个绿色化转型关键词.xlsx"
    
    # 公司元数据Excel文件路径
    # 需要包含列：证券代码、行业代码、行业名称、所属省份、所属省份代码、所属城市、所属城市代码
    COMPANY_INFO_FILE = "公司基础信息表.xlsx"
    
    # 输出文件路径（jieba版本）
    OUTPUT_FILE = "绿色化转型分析结果_jieba.xlsx"
    
    # 年份范围
    START_YEAR = 2024
    END_YEAR = 2024
    
    # 进程数（None表示自动根据CPU核心数调整）
    PROCESSES = None
    
    # ==================== 执行逻辑 ====================
    
    config = GreenAnalysisConfig(
        txt_folder=TXT_FOLDER,
        keywords_file=KEYWORDS_FILE,
        company_info_file=COMPANY_INFO_FILE,
        output_file=OUTPUT_FILE,
        start_year=START_YEAR,
        end_year=END_YEAR,
        processes=PROCESSES,
    )
    
    analyzer = GreenTransformationAnalyzerJieba(config)
    analyzer.run()
    
    print("\n提示：")
    print("1. 本脚本使用jieba分词，年报词数为分词后的有效中文词数量")
    print("2. 关键词已添加到jieba词典，提高分词准确率")
    print("3. 如需对比字符串匹配版本，请运行 4.green_transformation_analysis.py")
