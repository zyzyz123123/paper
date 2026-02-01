#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''
@Project ：Annualreport_tools
@File    ：4.green_transformation_analysis.py
@IDE     ：PyCharm
@Author  ：lingxiaotian
@Date    ：2026/02/01
@Description: 绿色化转型关键词分析脚本
    - 读取113个绿色化转型关键词
    - 统计年报中关键词词频并求和
    - 计算绿色化转型指数 ln(词频+1)
    - 与公司元数据合并，输出完整表格
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
            df = pd.read_excel(file_path)
            
            if column_name:
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


def analyze_single_file(args: Tuple) -> Optional[Dict]:
    """分析单个TXT文件的关键词。
    
    Args:
        args: (file_path, keywords) 元组
        
    Returns:
        分析结果字典
    """
    file_path, keywords = args
    
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
        
        # 统计所有关键词的词频并求和
        total_keyword_freq = 0
        for keyword in keywords:
            count = words.count(keyword)
            total_keyword_freq += count
        
        # 计算绿色化转型指数 = ln(词频 + 1)
        green_index = math.log(total_keyword_freq + 1)
        
        return {
            'stock_code': stock_code,
            'stock_name': stock_name,
            'year': year,
            'green_freq': total_keyword_freq,
            'green_index': green_index,
        }
        
    except Exception as e:
        logging.error(f"分析文件失败 {file_path}: {e}")
        return None


class GreenTransformationAnalyzer:
    """绿色化转型分析器主类。"""
    
    def __init__(self, config: GreenAnalysisConfig):
        self.config = config
        self.keywords = []
        self.company_info_loader = None
    
    def _load_keywords(self) -> None:
        """加载关键词并添加到jieba词典。"""
        self.keywords = KeywordLoader.load_from_excel(self.config.keywords_file)
        
        # 将关键词添加到jieba词典
        for word in self.keywords:
            jieba.add_word(word)
        
        logging.info(f"已将 {len(self.keywords)} 个关键词添加到jieba词典")
    
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
                '证券简称': result['stock_name'],
                'stkcd': int(stock_code),
                'year': int(result['year']),
                '绿色化转型词频': result['green_freq'],
                '绿色化转型指数': round(result['green_index'], 6),
                '行业代码': info['industry_code'],
                '行业名称': info['industry_name'],
                '所属省份': info['province'],
                '所属省份代码': info['province_code'],
                '所属城市': info['city'],
                '所属城市代码': info['city_code'],
            }
            rows.append(row)
        
        df = pd.DataFrame(rows)
        
        # 按股票代码和年份排序
        if not df.empty:
            df = df.sort_values(['stkcd', 'year']).reset_index(drop=True)
        
        return df
    
    def run(self) -> None:
        """执行分析流程。"""
        logging.info("=" * 60)
        logging.info("绿色化转型关键词分析启动")
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
        
        # 准备任务
        tasks = [(f, self.keywords) for f in txt_files]
        
        # 多进程分析
        worker_count = self.config.processes or min(cpu_count(), len(tasks))
        logging.info(f"使用 {worker_count} 个进程处理 {len(tasks)} 个文件")
        
        results = []
        with Pool(processes=worker_count) as pool:
            for i, result in enumerate(pool.imap_unordered(analyze_single_file, tasks)):
                results.append(result)
                
                # 显示进度
                progress = (i + 1) / len(tasks) * 100
                print(f"\r分析进度: {i + 1}/{len(tasks)} ({progress:.1f}%)", end='', flush=True)
        
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
    
    # 输出文件路径
    OUTPUT_FILE = "绿色化转型分析结果.xlsx"
    
    # 年份范围
    START_YEAR = 2012
    END_YEAR = 2012  # 先用2012年测试
    
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
    
    analyzer = GreenTransformationAnalyzer(config)
    analyzer.run()
    
    print("\n提示：")
    print("1. 如果行业/地区信息为空，请检查公司元数据文件是否存在且格式正确")
    print("2. TXT文件命名格式需为: 股票代码_公司简称_年份.txt (如: 000001_平安银行_2023.txt)")
    print("3. 公司元数据文件需包含: 证券代码、行业代码、行业名称、所属省份、所属省份代码、所属城市、所属城市代码")
