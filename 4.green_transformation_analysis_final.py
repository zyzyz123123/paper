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
    
输出列：
    证券代码, 证券名称, stkcd, 年份, 绿色化转型词数, 年报词数
'''

from __future__ import annotations

import logging
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
    
    支持三种文件名格式：
    - 旧格式：600582_天地科技_2014.txt
    - 2014文件夹格式：600582-天地科技-2014年年度报告（更新版）_2015-03-28.txt
    - 无“年份年”的：600367-红星发展-年报（修订版）_2015-04-03.txt（年份从所在目录名取）
    """
    global _global_keywords
    
    try:
        filename = os.path.basename(file_path)
        parent_dir = os.path.basename(os.path.dirname(file_path))
        folder_year = parent_dir if parent_dir.isdigit() and len(parent_dir) == 4 else None

        # 格式1：6位代码_公司名_年份.txt
        match = re.match(r'^(\d{6})_(.+?)_(\d{4})\.txt$', filename)
        if match:
            stock_code = match.group(1)
            stock_name = match.group(2)
            year = match.group(3)
        else:
            # 格式2：6位代码-公司名-...2014年....txt（如 2014 文件夹）
            match2 = re.match(r'^(\d{6})-([^-]+)-.*?(\d{4})年.*\.txt$', filename)
            if match2:
                stock_code = match2.group(1)
                stock_name = match2.group(2)
                year = match2.group(3)
            else:
                # 格式3：6位代码-公司名-年报（修订版）等，无“年份年”，年份用目录名
                match3 = re.match(r'^(\d{6})-([^-]+)-.+\.txt$', filename)
                if match3 and folder_year:
                    stock_code = match3.group(1)
                    stock_name = match3.group(2)
                    year = folder_year
                else:
                    return None
        
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
        green_word_count = sum(content.count(kw) for kw in _global_keywords)
        
        # ========== 年报总词数：jieba分词 ==========
        words = jieba.cut(content)
        chinese_pattern = re.compile(r'[\u4e00-\u9fff]')
        total_word_count = sum(1 for w in words if chinese_pattern.search(w))
        
        return {
            'stock_code': stock_code,
            'stock_name': stock_name,
            'year': year,
            'green_word_count': green_word_count,
            'total_word_count': total_word_count,
        }
        
    except Exception as e:
        logging.error(f"分析文件失败 {file_path}: {e}")
        return None


class GreenTransformationAnalyzer:
    """绿色化转型分析器。"""
    
    def __init__(self, config: GreenAnalysisConfig):
        self.config = config
        self.keywords = []
    
    def _load_keywords(self) -> None:
        self.keywords = KeywordLoader.load_from_excel(self.config.keywords_file)
    
    def _collect_txt_files(self) -> List[str]:
        # 只从“纯年份”文件夹收集（2014、2015、…、2023），排除 2014_重复_手动 等
        allowed_year_folders = {str(y) for y in range(self.config.start_year, self.config.end_year + 1)}
        txt_files = []
        for root, dirs, files in os.walk(self.config.txt_folder):
            root_basename = os.path.basename(root.rstrip(os.sep))
            folder_year = int(root_basename) if root_basename.isdigit() and len(root_basename) == 4 else None
            # 只处理目录名恰好为 2014、2015、…、2023 的文件夹
            if root_basename not in allowed_year_folders:
                continue

            for filename in files:
                if not filename.endswith('.txt'):
                    continue
                # 格式1：末尾 _年份.txt
                match = re.search(r'_(\d{4})\.txt$', filename)
                if match:
                    year = int(match.group(1))
                    if self.config.start_year <= year <= self.config.end_year:
                        txt_files.append(os.path.join(root, filename))
                    continue
                # 格式2：文件名中含 "年份年"（如 2014年年度报告）
                match2 = re.search(r'(\d{4})年', filename)
                if match2:
                    year = int(match2.group(1))
                    if self.config.start_year <= year <= self.config.end_year:
                        txt_files.append(os.path.join(root, filename))
                    continue
                # 格式3：在“年份”目录下且为 6位代码-公司名-...txt（如 年报（修订版）、年报（更新））
                if folder_year is not None and self.config.start_year <= folder_year <= self.config.end_year:
                    if re.match(r'^\d{6}-[^-]+-.+\.txt$', filename):
                        txt_files.append(os.path.join(root, filename))
        
        logging.info(f"找到 {len(txt_files)} 个符合条件的TXT文件")
        return txt_files
    
    def _build_output_dataframe(self, results: List[Dict]) -> pd.DataFrame:
        rows = []
        for result in results:
            if result is None:
                continue
            
            stock_code = result['stock_code']
            rows.append({
                '证券代码': stock_code,
                '证券名称': result['stock_name'],
                'stkcd': int(stock_code),
                '年份': int(result['year']),
                '绿色化转型词数': result['green_word_count'],
                '年报词数': result['total_word_count'],
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
    START_YEAR = 2011
    END_YEAR = 2024
    
    # 生成带年份的输出文件名
    if START_YEAR == END_YEAR:
        output_filename = f"绿色化转型分析结果_final_{START_YEAR}.xlsx"
    else:
        output_filename = f"绿色化转型分析结果_final_{START_YEAR}-{END_YEAR}.xlsx"
    
    # 从当前目录遍历，会扫描 2014、2015、…、2023 等年份子文件夹内的 txt
    config = GreenAnalysisConfig(
        txt_folder=".",
        keywords_file="副本数字化转型关键词(1).xlsx",
        output_file=output_filename,
        start_year=START_YEAR,
        end_year=END_YEAR,
    )
    
    analyzer = GreenTransformationAnalyzer(config)
    analyzer.run()
