#!/usr/bin/env python
# -*- coding: UTF-8 -*-
'''
@Project ：Annualreport_tools
@File    ：5.company_info_crawler.py
@Description: 上市公司基础信息爬虫
    - 从东方财富获取公司详细信息
    - 包括：行业代码、行业名称、省份、城市等
    - 输出公司基础信息表
'''

from __future__ import annotations

import logging
import os
import re
import time
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import requests

# 日志配置
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# 省份代码映射表
PROVINCE_CODE_MAP = {
    '北京': '11', '天津': '12', '河北': '13', '山西': '14', '内蒙古': '15',
    '辽宁': '21', '吉林': '22', '黑龙江': '23',
    '上海': '31', '江苏': '32', '浙江': '33', '安徽': '34', '福建': '35', '江西': '36', '山东': '37',
    '河南': '41', '湖北': '42', '湖南': '43', '广东': '44', '广西': '45', '海南': '46',
    '重庆': '50', '四川': '51', '贵州': '52', '云南': '53', '西藏': '54',
    '陕西': '61', '甘肃': '62', '青海': '63', '宁夏': '64', '新疆': '65',
}

# 城市代码映射表将在运行时从GitHub加载
CITY_CODE_MAP = {}


def load_city_codes() -> Dict[str, str]:
    """从GitHub加载城市代码映射表。"""
    global CITY_CODE_MAP
    
    if CITY_CODE_MAP:
        return CITY_CODE_MAP
    
    url = "https://raw.githubusercontent.com/uiwjs/province-city-china/gh-pages/city.json"
    
    try:
        logging.info("正在从GitHub加载城市代码数据...")
        response = requests.get(url, timeout=15)
        if response.status_code == 200:
            city_list = response.json()
            # 构建映射表：城市名 -> 城市代码
            for city in city_list:
                name = city.get("name", "")
                code = city.get("code", "")
                if name and code:
                    CITY_CODE_MAP[name] = code
                    # 同时添加不带"市"的版本方便匹配
                    if name.endswith("市"):
                        CITY_CODE_MAP[name[:-1]] = code
            logging.info(f"成功加载 {len(city_list)} 个城市代码")
        else:
            logging.warning(f"加载城市代码失败: HTTP {response.status_code}")
    except Exception as e:
        logging.warning(f"加载城市代码失败: {e}")
    
    # 如果加载失败，使用备用的直辖市数据
    if not CITY_CODE_MAP:
        CITY_CODE_MAP.update({
            '北京市': '110100', '北京': '110100',
            '天津市': '120100', '天津': '120100',
            '上海市': '310100', '上海': '310100',
            '重庆市': '500100', '重庆': '500100',
        })
    
    return CITY_CODE_MAP


class EastMoneyClient:
    """东方财富API客户端。"""
    
    # 公司资料API
    COMPANY_INFO_URL = "https://emweb.securities.eastmoney.com/PC_HSF10/CompanySurvey/CompanySurveyAjax"
    
    HEADERS = {
        "Accept": "*/*",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "Referer": "https://emweb.securities.eastmoney.com/",
    }
    
    def __init__(self, timeout: int = 15, retry_delay: float = 1.0) -> None:
        self.timeout = timeout
        self.retry_delay = retry_delay
        self.session = requests.Session()
        self.session.headers.update(self.HEADERS)
    
    def _get_market_code(self, stock_code: str) -> str:
        """根据股票代码判断市场。"""
        code = str(stock_code).zfill(6)
        # 6开头是沪市，0/3开头是深市
        if code.startswith('6'):
            return f"SH{code}"
        else:
            return f"SZ{code}"
    
    def get_company_info(self, stock_code: str, max_retries: int = 3) -> Optional[Dict]:
        """获取公司基本信息。
        
        Args:
            stock_code: 股票代码
            max_retries: 最大重试次数
            
        Returns:
            公司信息字典
        """
        code = str(stock_code).zfill(6)
        market_code = self._get_market_code(code)
        
        params = {"code": market_code}
        
        for attempt in range(1, max_retries + 1):
            try:
                response = self.session.get(
                    self.COMPANY_INFO_URL,
                    params=params,
                    timeout=self.timeout
                )
                
                if response.status_code == 200:
                    data = response.json()
                    return data
                else:
                    logging.warning(f"请求失败 {code}: 状态码 {response.status_code}")
                    
            except requests.exceptions.Timeout:
                logging.warning(f"请求超时 ({attempt}/{max_retries}): {code}")
            except requests.exceptions.RequestException as e:
                logging.warning(f"请求错误 ({attempt}/{max_retries}): {e}")
            except Exception as e:
                logging.warning(f"解析错误 ({attempt}/{max_retries}): {e}")
            
            if attempt < max_retries:
                time.sleep(self.retry_delay)
        
        return None


class CompanyInfoCrawler:
    """公司信息爬虫主类。"""
    
    def __init__(
        self,
        source_excel: str,
        output_file: str,
        request_delay: float = 0.3,
        max_retries: int = 3
    ) -> None:
        self.source_excel = source_excel
        self.output_file = output_file
        self.request_delay = request_delay
        self.max_retries = max_retries
        self.client = EastMoneyClient()
    
    def _load_company_codes(self) -> List[str]:
        """从Excel加载公司代码列表。"""
        try:
            df = pd.read_excel(self.source_excel)
            
            # 尝试多种可能的列名
            code_column = None
            for col_name in ['公司代码', '证券代码', '股票代码', 'code', 'stkcd']:
                if col_name in df.columns:
                    code_column = col_name
                    break
            
            if code_column is None:
                logging.error(f"Excel中未找到公司代码列，可用列: {list(df.columns)}")
                return []
            
            # 提取唯一的公司代码
            codes = df[code_column].astype(str).str.zfill(6).unique().tolist()
            logging.info(f"从Excel加载了 {len(codes)} 个唯一公司代码")
            return codes
            
        except FileNotFoundError:
            logging.error(f"Excel文件不存在: {self.source_excel}")
            return []
        except Exception as e:
            logging.error(f"加载Excel失败: {e}")
            return []
    
    def _extract_province(self, address: str) -> tuple:
        """从地址提取省份信息。"""
        if not address:
            return "", ""
        
        # 直辖市特殊处理
        for city in ['北京', '天津', '上海', '重庆']:
            if city in address:
                return city, PROVINCE_CODE_MAP.get(city, "")
        
        # 匹配省份
        for prov_name, prov_code in PROVINCE_CODE_MAP.items():
            if prov_name in address:
                return prov_name, prov_code
        
        return "", ""
    
    def _extract_city(self, address: str) -> tuple:
        """从地址提取城市信息。"""
        if not address:
            return "", ""
        
        # 确保城市代码已加载
        city_map = load_city_codes()
        
        # 直辖市特殊处理
        for city in ['北京', '天津', '上海', '重庆']:
            if city in address:
                city_name = city + "市"
                city_code = city_map.get(city_name, "")
                return city_name, city_code
        
        # 先清理地址，去掉省份部分
        # 匹配并移除 "XX省" 或 "XX自治区" 部分
        clean_address = re.sub(r'^[\u4e00-\u9fa5]{2,6}(省|自治区)', '', address)
        
        # 用正则提取"XX市"格式的城市名（不贪婪，取最短匹配）
        # 排除"省"字结尾的匹配
        city_match = re.search(r'(?<!省)([\u4e00-\u9fa5]{2,4}市)', clean_address)
        if city_match:
            city_name = city_match.group(1)
            # 确保不是"XX省市"这种错误匹配
            if not city_name.startswith(('省', '区')):
                city_code = city_map.get(city_name, "")
                if not city_code:
                    # 尝试不带"市"匹配
                    city_code = city_map.get(city_name[:-1], "")
                return city_name, city_code
        
        # 如果没有匹配到"XX市"，尝试从映射表反向匹配
        for city_name, city_code in city_map.items():
            if city_name.endswith("市"):
                city_short = city_name[:-1]
                if len(city_short) >= 2 and city_short in clean_address:
                    return city_name, city_code
        
        return "", ""
    
    def _parse_company_info(self, stock_code: str, data: Dict) -> Dict:
        """解析公司信息。"""
        result = {
            "证券代码": stock_code,
            "证券简称": "",
            "行业代码": "",
            "行业名称": "",
            "所属省份": "",
            "所属省份代码": "",
            "所属城市": "",
            "所属城市代码": "",
        }
        
        if not data:
            return result
        
        try:
            # 获取基本信息
            jbzl = data.get("jbzl", {})
            if jbzl:
                result["证券简称"] = jbzl.get("gsmc", "") or jbzl.get("agjc", "")
                
                # 行业信息
                industry = jbzl.get("sshy", "")
                if industry:
                    # 尝试提取行业代码（如果有）
                    code_match = re.match(r'^([A-Z]\d{0,2})\s*[-—]?\s*(.*)$', industry)
                    if code_match:
                        result["行业代码"] = code_match.group(1)
                        result["行业名称"] = code_match.group(2) or industry
                    else:
                        result["行业名称"] = industry
                
                # 地址信息
                address = jbzl.get("zcdz", "") or jbzl.get("bgdz", "")
                if address:
                    province, province_code = self._extract_province(address)
                    city, city_code = self._extract_city(address)
                    result["所属省份"] = province
                    result["所属省份代码"] = province_code
                    result["所属城市"] = city
                    result["所属城市代码"] = city_code
            
            # 尝试从其他字段获取
            if not result["证券简称"]:
                fxxg = data.get("fxxg", {})
                if fxxg:
                    result["证券简称"] = fxxg.get("agdm", "")
                    
        except Exception as e:
            logging.warning(f"解析公司 {stock_code} 信息失败: {e}")
        
        return result
    
    def run(self) -> None:
        """执行爬取任务。"""
        logging.info("=" * 60)
        logging.info("公司基础信息爬虫启动（东方财富数据源）")
        logging.info(f"来源文件: {self.source_excel}")
        logging.info(f"输出文件: {self.output_file}")
        logging.info("=" * 60)
        
        # 预加载城市代码
        load_city_codes()
        
        # 加载公司代码
        company_codes = self._load_company_codes()
        if not company_codes:
            logging.error("未获取到公司代码")
            return
        
        # 爬取公司信息
        results = []
        total = len(company_codes)
        success_count = 0
        
        for idx, code in enumerate(company_codes, 1):
            # 获取公司信息
            data = self.client.get_company_info(code, self.max_retries)
            
            # 解析信息
            result = self._parse_company_info(code, data)
            results.append(result)
            
            if result["证券简称"]:
                success_count += 1
                city_info = result['所属城市'] if result['所属城市'] else result['所属省份']
                city_code_info = f"({result['所属城市代码']})" if result['所属城市代码'] else ""
                logging.info(f"[{idx}/{total}] {code} - {result['证券简称']} - {result['行业名称']} - {city_info}{city_code_info}")
            else:
                logging.warning(f"[{idx}/{total}] {code} - 获取失败")
            
            # 请求间隔
            if idx < total:
                time.sleep(self.request_delay)
            
            # 进度显示
            if idx % 10 == 0:
                progress = (idx / total) * 100
                print(f"\r进度: {idx}/{total} ({progress:.1f}%)", end='', flush=True)
        
        print()  # 换行
        
        # 保存结果
        df = pd.DataFrame(results)
        
        output_path = Path(self.output_file)
        df.to_excel(output_path, index=False, engine='openpyxl')
        
        logging.info("=" * 60)
        logging.info(f"爬取完成!")
        logging.info(f"总公司数: {len(results)}")
        logging.info(f"成功获取: {success_count}")
        logging.info(f"获取失败: {len(results) - success_count}")
        logging.info(f"输出文件: {output_path}")
        logging.info("=" * 60)


if __name__ == '__main__':
    # ==================== 配置区域 ====================
    
    # 来源Excel文件路径（包含公司代码的年报链接表）
    SOURCE_EXCEL = "年报链接_2012.xlsx"
    
    # 输出文件路径
    OUTPUT_FILE = "公司基础信息表.xlsx"
    
    # 爬虫配置
    REQUEST_DELAY = 0.3  # 请求间隔（秒）
    MAX_RETRIES = 3  # 最大重试次数
    
    # ==================== 执行逻辑 ====================
    
    if not os.path.exists(SOURCE_EXCEL):
        logging.error(f"来源文件不存在: {SOURCE_EXCEL}")
    else:
        crawler = CompanyInfoCrawler(
            source_excel=SOURCE_EXCEL,
            output_file=OUTPUT_FILE,
            request_delay=REQUEST_DELAY,
            max_retries=MAX_RETRIES,
        )
        crawler.run()
