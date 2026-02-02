#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''
@Project ：Annualreport_tools
@File    ：4.green_transformation_analysis_final.py
@IDE     ：PyCharm
@Author  ：lingxiaotian
@Date    ：2026/02/01
@Description: 绿色化转型关键词分析脚本（最终版）

统计方法：
    - 关键词匹配：字符串匹配（不会漏计专业术语）
    - 年报总词数：jieba分词后的有效中文词数量
    - 绿色化转型词频 = 关键词出现次数 / 年报总词数
    - 绿色化转型指数 = ln(关键词出现次数 + 1)
    
输出列：
    证券代码, 证券名称, stkcd, 年份,
    绿色化转型词数, 年报词数, 绿色化转型词频, 绿色化转型指数,
    行业代码, 行业名称, 省份代码, 省份名称, 城市代码, 城市名称
'''

from __future__ import annotations

import logging
import math
import os
import re
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


class KeywordLoader:
    """关键词加载器。"""
    
    @staticmethod
    def load_from_excel(file_path: str, column_name: str = None) -> List[str]:
        """从Excel文件加载关键词列表。"""
        try:
            df = pd.read_excel(file_path, header=None)
            first_cell = str(df.iloc[0, 0]).strip() if not df.empty else ""
            
            header_keywords = ['关键词', '关键字', 'keyword', '词汇', '词语']
            if any(h in first_cell.lower() for h in header_keywords):
                df = df.iloc[1:]
            
            if column_name:
                df = pd.read_excel(file_path)
                keywords = df[column_name].dropna().astype(str).tolist()
            else:
                keywords = df.iloc[:, 0].dropna().astype(str).tolist()
            
            keywords = [kw.strip() for kw in keywords if kw.strip()]
            logging.info(f"成功加载 {len(keywords)} 个关键词")
            return keywords
            
        except Exception as e:
            logging.error(f"加载关键词失败: {e}")
            raise


class CompanyInfoLoader:
    """公司元数据加载器。"""
    
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
        for name in possible_names:
            if name in df.columns:
                return name
        return None
    
    def _load_data(self) -> None:
        try:
            self.df = pd.read_excel(self.file_path)
            logging.info(f"成功加载公司元数据: {len(self.df)} 条记录")
            
            code_col = self._find_column(self.df, self.COLUMN_MAPPING['stock_code'])
            if code_col:
                self.df['_stock_code'] = self.df[code_col].astype(str).str.zfill(6)
            else:
                logging.warning("未找到股票代码列")
                
        except Exception as e:
            logging.error(f"加载公司元数据失败: {e}")
            raise
    
    def get_company_info(self, stock_code: str) -> Dict[str, str]:
        result = {
            'industry_code': '', 'industry_name': '',
            'province': '', 'province_code': '',
            'city': '', 'city_code': '',
        }
        
        if self.df is None or '_stock_code' not in self.df.columns:
            return result
        
        stock_code = str(stock_code).zfill(6)
        matches = self.df[self.df['_stock_code'] == stock_code]
        
        if matches.empty:
            return result
        
        row = matches.iloc[0]
        for field, possible_names in self.COLUMN_MAPPING.items():
            if field in ['stock_code', 'stock_name']:
                continue
            col = self._find_column(self.df, possible_names)
            if col and pd.notna(row[col]):
                result[field] = str(row[col])
        
        return result


# 全局变量
_global_keywords = None


def init_worker(keywords: List[str]):
    """初始化工作进程。"""
    global _global_keywords
    _global_keywords = keywords


def analyze_single_file(file_path: str) -> Optional[Dict]:
    """分析单个TXT文件。
    
    统计方法：
    - 关键词词数：字符串匹配（不漏计）
    - 年报总词数：jieba分词后的有效中文词数
    """
    global _global_keywords
    
    try:
        filename = os.path.basename(file_path)
        match = re.match(r'^(\d{6})_(.+?)_(\d{4})\.txt$', filename)
        
        if not match:
            return None
        
        stock_code = match.group(1)
        stock_name = match.group(2)
        year = match.group(3)
        
        # 读取文件
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            with open(file_path, 'r', encoding='gbk', errors='ignore') as f:
                content = f.read()
        
        if not content.strip():
            return None
        
        # ========== 关键词词数：字符串匹配 ==========
        # 对于专业术语，字符串匹配不会漏计
        green_word_count = sum(content.count(kw) for kw in _global_keywords)
        
        # ========== 年报总词数：jieba分词 ==========
        # 分词后过滤，只保留包含中文的词
        words = jieba.cut(content)
        chinese_pattern = re.compile(r'[\u4e00-\u9fff]')
        total_word_count = sum(1 for w in words if chinese_pattern.search(w))
        
        # ========== 计算指标 ==========
        # 词频 = 关键词出现次数 / 年报总词数
        if total_word_count > 0:
            green_freq = green_word_count / total_word_count
        else:
            green_freq = 0.0
        
        # 指数 = ln(关键词出现次数 + 1)
        green_index = math.log(green_word_count + 1)
        
        return {
            'stock_code': stock_code,
            'stock_name': stock_name,
            'year': year,
            'green_word_count': green_word_count,
            'total_word_count': total_word_count,
            'green_freq': green_freq,
            'green_index': green_index,
        }
        
    except Exception as e:
        logging.error(f"分析文件失败 {file_path}: {e}")
        return None


class GreenTransformationAnalyzer:
    """绿色化转型分析器。"""
    
    def __init__(self, config: GreenAnalysisConfig):
        self.config = config
        self.keywords = []
        self.company_info_loader = None
    
    def _load_keywords(self) -> None:
        self.keywords = KeywordLoader.load_from_excel(self.config.keywords_file)
    
    def _load_company_info(self) -> None:
        if os.path.exists(self.config.company_info_file):
            self.company_info_loader = CompanyInfoLoader(self.config.company_info_file)
    
    def _collect_txt_files(self) -> List[str]:
        txt_files = []
        for root, dirs, files in os.walk(self.config.txt_folder):
            for filename in files:
                if not filename.endswith('.txt'):
                    continue
                match = re.search(r'_(\d{4})\.txt$', filename)
                if match:
                    year = int(match.group(1))
                    if self.config.start_year <= year <= self.config.end_year:
                        txt_files.append(os.path.join(root, filename))
        
        logging.info(f"找到 {len(txt_files)} 个符合条件的TXT文件")
        return txt_files
    
    def _build_output_dataframe(self, results: List[Dict]) -> pd.DataFrame:
        rows = []
        for result in results:
            if result is None:
                continue
            
            stock_code = result['stock_code']
            info = self.company_info_loader.get_company_info(stock_code) if self.company_info_loader else {
                'industry_code': '', 'industry_name': '',
                'province': '', 'province_code': '',
                'city': '', 'city_code': '',
            }
            
            rows.append({
                '证券代码': stock_code,
                '证券名称': result['stock_name'],
                'stkcd': int(stock_code),
                '年份': int(result['year']),
                '绿色化转型词数': result['green_word_count'],
                '年报词数': result['total_word_count'],
                '绿色化转型词频': round(result['green_freq'], 10),
                '绿色化转型指数': round(result['green_index'], 6),
                '行业代码': info['industry_code'],
                '行业名称': info['industry_name'],
                '省份代码': info['province_code'],
                '省份名称': info['province'],
                '城市代码': info['city_code'],
                '城市名称': info['city'],
            })
        
        df = pd.DataFrame(rows)
        if not df.empty:
            df = df.sort_values(['stkcd', '年份']).reset_index(drop=True)
        return df
    
    def run(self) -> None:
        logging.info("=" * 60)
        logging.info("绿色化转型关键词分析（最终版）")
        logging.info(f"统计方法：关键词=字符串匹配，总词数=jieba分词")
        logging.info(f"年份范围: {self.config.start_year} - {self.config.end_year}")
        logging.info("=" * 60)
        
        self._load_keywords()
        self._load_company_info()
        txt_files = self._collect_txt_files()
        
        if not txt_files:
            logging.error("未找到符合条件的TXT文件")
            return
        
        worker_count = self.config.processes or min(cpu_count(), len(txt_files))
        logging.info(f"使用 {worker_count} 个进程处理 {len(txt_files)} 个文件")
        
        results = []
        with Pool(processes=worker_count, initializer=init_worker, initargs=(self.keywords,)) as pool:
            for i, result in enumerate(pool.imap_unordered(analyze_single_file, txt_files)):
                results.append(result)
                print(f"\r进度: {i + 1}/{len(txt_files)} ({(i + 1) / len(txt_files) * 100:.1f}%)", end='', flush=True)
        
        print()
        
        valid_results = [r for r in results if r is not None]
        logging.info(f"成功分析 {len(valid_results)} 个文件")
        
        df = self._build_output_dataframe(valid_results)
        
        output_path = Path(self.config.output_file)
        df.to_excel(output_path.with_suffix('.xlsx'), index=False, engine='openpyxl')
        
        logging.info(f"输出文件: {output_path}")
        
        if not df.empty:
            print("\n数据预览:")
            print(df.head().to_string())


if __name__ == '__main__':
    config = GreenAnalysisConfig(
        txt_folder="年报文件",
        keywords_file="113个绿色化转型关键词.xlsx",
        company_info_file="公司基础信息表.xlsx",
        output_file="绿色化转型分析结果_final.xlsx",
        start_year=2024,
        end_year=2024,
    )
    
    analyzer = GreenTransformationAnalyzer(config)
    analyzer.run()
