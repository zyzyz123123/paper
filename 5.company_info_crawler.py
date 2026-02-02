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

# 证监会行业分类代码映射表（2012版）
# 格式：行业名称 -> 行业代码
CSRC_INDUSTRY_CODE_MAP = {
    # A 农、林、牧、渔业
    "农业": "A01", "林业": "A02", "畜牧业": "A03", "渔业": "A04", 
    "农、林、牧、渔专业及辅助性活动": "A05", "农、林、牧、渔服务业": "A05",
    "农、林、牧、渔业": "A",
    
    # B 采矿业
    "煤炭开采和洗选业": "B06", "石油和天然气开采业": "B07", 
    "黑色金属矿采选业": "B08", "有色金属矿采选业": "B09",
    "非金属矿采选业": "B10", "开采专业及辅助性活动": "B11", "开采辅助活动": "B11",
    "其他采矿业": "B12", "采矿业": "B",
    
    # C 制造业
    "农副食品加工业": "C13", "食品制造业": "C14", 
    "酒、饮料和精制茶制造业": "C15", "烟草制品业": "C16",
    "纺织业": "C17", "纺织服装、服饰业": "C18", 
    "皮革、毛皮、羽毛及其制品和制鞋业": "C19",
    "木材加工和木、竹、藤、棕、草制品业": "C20", "木材加工及木、竹、藤、棕、草制品业": "C20",
    "家具制造业": "C21", "造纸和纸制品业": "C22", "造纸及纸制品业": "C22",
    "印刷和记录媒介复制业": "C23", "印刷业和记录媒介的复制": "C23",
    "文教、工美、体育和娱乐用品制造业": "C24", "文教体育用品制造业": "C24",
    "石油、煤炭及其他燃料加工业": "C25", "石油加工、炼焦及核燃料加工业": "C25",
    "化学原料和化学制品制造业": "C26", "化学原料及化学制品制造业": "C26",
    "医药制造业": "C27", "化学纤维制造业": "C28",
    "橡胶和塑料制品业": "C29", "非金属矿物制品业": "C30",
    "黑色金属冶炼和压延加工业": "C31", "黑色金属冶炼及压延加工业": "C31",
    "有色金属冶炼和压延加工业": "C32", "有色金属冶炼及压延加工业": "C32",
    "金属制品业": "C33", "通用设备制造业": "C34", "专用设备制造业": "C35",
    "汽车制造业": "C36", "铁路、船舶、航空航天和其他运输设备制造业": "C37",
    "铁路、船舶、航空航天和其它运输设备制造业": "C37",
    "电气机械和器材制造业": "C38", "电气机械及器材制造业": "C38",
    "计算机、通信和其他电子设备制造业": "C39", "计算机、通信和其它电子设备制造业": "C39",
    "仪器仪表制造业": "C40", "其他制造业": "C41", "废弃资源综合利用业": "C42",
    "金属制品、机械和设备修理业": "C43", "制造业": "C",
    
    # D 电力、热力、燃气及水生产和供应业
    "电力、热力生产和供应业": "D44", "电力、热力的生产和供应业": "D44",
    "燃气生产和供应业": "D45", "燃气生产及供应业": "D45",
    "水的生产和供应业": "D46", "电力、热力、燃气及水生产和供应业": "D",
    
    # E 建筑业
    "房屋建筑业": "E47", "土木工程建筑业": "E48", 
    "建筑安装业": "E49", "建筑装饰、装修和其他建筑业": "E50",
    "建筑装饰和其他建筑业": "E50", "建筑业": "E",
    
    # F 批发和零售业
    "批发业": "F51", "零售业": "F52", "批发和零售业": "F",
    
    # G 交通运输、仓储和邮政业
    "铁路运输业": "G53", "道路运输业": "G54", "水上运输业": "G55",
    "航空运输业": "G56", "管道运输业": "G57", 
    "多式联运和运输代理业": "G58", "装卸搬运和仓储业": "G59",
    "装卸搬运和运输代理业": "G58", "仓储业": "G59",
    "邮政业": "G60", "交通运输、仓储和邮政业": "G",
    
    # H 住宿和餐饮业
    "住宿业": "H61", "餐饮业": "H62", "住宿和餐饮业": "H",
    
    # I 信息传输、软件和信息技术服务业
    "电信、广播电视和卫星传输服务": "I63", "电信、广播电视和卫星传输服务业": "I63",
    "互联网和相关服务": "I64", "互联网和相关服务业": "I64",
    "软件和信息技术服务业": "I65", "信息传输、软件和信息技术服务业": "I",
    
    # J 金融业
    "货币金融服务": "J66", "货币金融服务业": "J66",
    "资本市场服务": "J67", "资本市场服务业": "J67",
    "保险业": "J68", "其他金融业": "J69", "金融业": "J",
    
    # K 房地产业
    "房地产业": "K70", "房地产": "K70",
    
    # L 租赁和商务服务业
    "租赁业": "L71", "商务服务业": "L72", "租赁和商务服务业": "L",
    
    # M 科学研究和技术服务业
    "研究和试验发展": "M73", "专业技术服务业": "M74",
    "科技推广和应用服务业": "M75", "科学研究和技术服务业": "M",
    
    # N 水利、环境和公共设施管理业
    "水利管理业": "N76", "生态保护和环境治理业": "N77",
    "公共设施管理业": "N78", "土地管理业": "N79",
    "水利、环境和公共设施管理业": "N",
    
    # O 居民服务、修理和其他服务业
    "居民服务业": "O80", "机动车、电子产品和日用产品修理业": "O81",
    "其他服务业": "O82", "居民服务、修理和其他服务业": "O",
    
    # P 教育
    "教育": "P83", "教育业": "P83",
    
    # Q 卫生和社会工作
    "卫生": "Q84", "卫生业": "Q84", "社会工作": "Q85", "卫生和社会工作": "Q",
    
    # R 文化、体育和娱乐业
    "新闻和出版业": "R85", "广播、电视、电影和录音制作业": "R86",
    "广播、电视、电影和影视录音制作业": "R86",
    "文化艺术业": "R87", "体育": "R88", "娱乐业": "R89",
    "文化、体育和娱乐业": "R",
    
    # S 综合
    "综合": "S90",
}


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
    
    # 始终添加直辖市的城市代码（确保直辖市被正确处理）
    # 直辖市在数据源中可能作为省份而非城市，所以需要手动添加
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
        # 6开头是沪市
        if code.startswith('6'):
            return f"SH{code}"
        # 43/83/87/92开头是北交所
        elif code.startswith(('43', '83', '87', '92')):
            return f"BJ{code}"
        # 0/3开头是深市
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
    
    def _get_industry_code(self, industry_name: str) -> str:
        """根据行业名称获取证监会行业代码。
        
        Args:
            industry_name: 行业名称，如"制造业-酒、饮料和精制茶制造业"
            
        Returns:
            行业代码，如"C15"
        """
        if not industry_name:
            return ""
        
        # 先尝试精确匹配
        if industry_name in CSRC_INDUSTRY_CODE_MAP:
            return CSRC_INDUSTRY_CODE_MAP[industry_name]
        
        # 如果是"大类-子类"格式，尝试匹配子类
        if "-" in industry_name:
            parts = industry_name.split("-")
            sub_industry = parts[-1].strip()  # 取最后一个部分（子类）
            if sub_industry in CSRC_INDUSTRY_CODE_MAP:
                return CSRC_INDUSTRY_CODE_MAP[sub_industry]
            
            # 尝试匹配大类
            main_industry = parts[0].strip()
            if main_industry in CSRC_INDUSTRY_CODE_MAP:
                return CSRC_INDUSTRY_CODE_MAP[main_industry]
        
        # 模糊匹配：遍历映射表，查找包含关系
        for name, code in CSRC_INDUSTRY_CODE_MAP.items():
            if name in industry_name or industry_name in name:
                return code
        
        return ""
    
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
                
                # 行业信息 - 优先使用证监会行业分类(sszjhhy)
                csrc_industry = jbzl.get("sszjhhy", "")  # 证监会行业分类，如"制造业-酒、饮料和精制茶制造业"
                em_industry = jbzl.get("sshy", "")  # 东方财富行业分类，如"酿酒行业"
                
                if csrc_industry:
                    # 解析证监会行业分类
                    result["行业代码"] = self._get_industry_code(csrc_industry)
                    # 取子类名称作为行业名称
                    if "-" in csrc_industry:
                        result["行业名称"] = csrc_industry.split("-")[-1].strip()
                    else:
                        result["行业名称"] = csrc_industry
                elif em_industry:
                    # 退回使用东方财富行业分类
                    result["行业名称"] = em_industry
                    result["行业代码"] = self._get_industry_code(em_industry)
                
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
                industry_info = f"{result['行业代码']}-{result['行业名称']}" if result['行业代码'] else result['行业名称']
                logging.info(f"[{idx}/{total}] {code} - {result['证券简称']} - {industry_info} - {city_info}{city_code_info}")
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


def collect_unique_codes_from_multiple_years(
    start_year: int,
    end_year: int,
    file_pattern: str = "年报链接_{year}.xlsx"
) -> List[str]:
    """从多年份的年报链接文件中收集所有唯一公司代码。
    
    Args:
        start_year: 起始年份
        end_year: 结束年份
        file_pattern: 文件名模式，{year}会被替换为年份
        
    Returns:
        去重后的公司代码列表
    """
    all_codes = set()
    
    for year in range(start_year, end_year + 1):
        file_path = file_pattern.format(year=year)
        
        if not os.path.exists(file_path):
            logging.warning(f"年报链接文件不存在，跳过: {file_path}")
            continue
        
        try:
            df = pd.read_excel(file_path)
            
            # 尝试多种可能的列名
            code_column = None
            for col_name in ['公司代码', '证券代码', '股票代码', 'code', 'stkcd']:
                if col_name in df.columns:
                    code_column = col_name
                    break
            
            if code_column is None:
                logging.warning(f"文件 {file_path} 中未找到公司代码列，可用列: {list(df.columns)}")
                continue
            
            # 提取公司代码
            codes = df[code_column].astype(str).str.zfill(6).unique()
            year_count = len(codes)
            all_codes.update(codes)
            logging.info(f"从 {file_path} 加载了 {year_count} 个公司代码")
            
        except Exception as e:
            logging.error(f"加载文件 {file_path} 失败: {e}")
    
    unique_codes = sorted(list(all_codes))
    logging.info(f"共收集到 {len(unique_codes)} 个唯一公司代码")
    return unique_codes


class MultiYearCompanyInfoCrawler:
    """支持多年份的公司信息爬虫。"""
    
    def __init__(
        self,
        start_year: int,
        end_year: int,
        output_file: str,
        file_pattern: str = "年报链接_{year}.xlsx",
        request_delay: float = 0.3,
        max_retries: int = 3
    ) -> None:
        self.start_year = start_year
        self.end_year = end_year
        self.output_file = output_file
        self.file_pattern = file_pattern
        self.request_delay = request_delay
        self.max_retries = max_retries
        self.client = EastMoneyClient()
    
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
        clean_address = re.sub(r'^[\u4e00-\u9fa5]{2,6}(省|自治区)', '', address)
        
        # 用正则提取"XX市"格式的城市名
        city_match = re.search(r'(?<!省)([\u4e00-\u9fa5]{2,4}市)', clean_address)
        if city_match:
            city_name = city_match.group(1)
            if not city_name.startswith(('省', '区')):
                city_code = city_map.get(city_name, "")
                if not city_code:
                    city_code = city_map.get(city_name[:-1], "")
                return city_name, city_code
        
        # 如果没有匹配到"XX市"，尝试从映射表反向匹配
        for city_name, city_code in city_map.items():
            if city_name.endswith("市"):
                city_short = city_name[:-1]
                if len(city_short) >= 2 and city_short in clean_address:
                    return city_name, city_code
        
        return "", ""
    
    def _get_industry_code(self, industry_name: str) -> str:
        """根据行业名称获取证监会行业代码。"""
        if not industry_name:
            return ""
        
        # 先尝试精确匹配
        if industry_name in CSRC_INDUSTRY_CODE_MAP:
            return CSRC_INDUSTRY_CODE_MAP[industry_name]
        
        # 如果是"大类-子类"格式，尝试匹配子类
        if "-" in industry_name:
            parts = industry_name.split("-")
            sub_industry = parts[-1].strip()
            if sub_industry in CSRC_INDUSTRY_CODE_MAP:
                return CSRC_INDUSTRY_CODE_MAP[sub_industry]
            
            main_industry = parts[0].strip()
            if main_industry in CSRC_INDUSTRY_CODE_MAP:
                return CSRC_INDUSTRY_CODE_MAP[main_industry]
        
        # 模糊匹配
        for name, code in CSRC_INDUSTRY_CODE_MAP.items():
            if name in industry_name or industry_name in name:
                return code
        
        return ""
    
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
            jbzl = data.get("jbzl", {})
            if jbzl:
                result["证券简称"] = jbzl.get("gsmc", "") or jbzl.get("agjc", "")
                
                csrc_industry = jbzl.get("sszjhhy", "")
                em_industry = jbzl.get("sshy", "")
                
                if csrc_industry:
                    result["行业代码"] = self._get_industry_code(csrc_industry)
                    if "-" in csrc_industry:
                        result["行业名称"] = csrc_industry.split("-")[-1].strip()
                    else:
                        result["行业名称"] = csrc_industry
                elif em_industry:
                    result["行业名称"] = em_industry
                    result["行业代码"] = self._get_industry_code(em_industry)
                
                address = jbzl.get("zcdz", "") or jbzl.get("bgdz", "")
                if address:
                    province, province_code = self._extract_province(address)
                    city, city_code = self._extract_city(address)
                    result["所属省份"] = province
                    result["所属省份代码"] = province_code
                    result["所属城市"] = city
                    result["所属城市代码"] = city_code
            
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
        logging.info("公司基础信息爬虫启动（多年份合并模式）")
        logging.info(f"年份范围: {self.start_year} - {self.end_year}")
        logging.info(f"文件模式: {self.file_pattern}")
        logging.info(f"输出文件: {self.output_file}")
        logging.info("=" * 60)
        
        # 预加载城市代码
        load_city_codes()
        
        # 从多年份文件收集唯一公司代码
        company_codes = collect_unique_codes_from_multiple_years(
            self.start_year,
            self.end_year,
            self.file_pattern
        )
        
        if not company_codes:
            logging.error("未获取到公司代码")
            return
        
        # 爬取公司信息
        results = []
        total = len(company_codes)
        success_count = 0
        
        for idx, code in enumerate(company_codes, 1):
            data = self.client.get_company_info(code, self.max_retries)
            result = self._parse_company_info(code, data)
            results.append(result)
            
            if result["证券简称"]:
                success_count += 1
                city_info = result['所属城市'] if result['所属城市'] else result['所属省份']
                city_code_info = f"({result['所属城市代码']})" if result['所属城市代码'] else ""
                industry_info = f"{result['行业代码']}-{result['行业名称']}" if result['行业代码'] else result['行业名称']
                logging.info(f"[{idx}/{total}] {code} - {result['证券简称']} - {industry_info} - {city_info}{city_code_info}")
            else:
                logging.warning(f"[{idx}/{total}] {code} - 获取失败")
            
            if idx < total:
                time.sleep(self.request_delay)
            
            if idx % 10 == 0:
                progress = (idx / total) * 100
                print(f"\r进度: {idx}/{total} ({progress:.1f}%)", end='', flush=True)
        
        print()
        
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


def get_stock_basic_info(stock_code: str) -> Optional[Dict]:
    """通过行情API获取股票基本信息（用于北交所等特殊股票）。
    
    Args:
        stock_code: 股票代码
        
    Returns:
        包含股票名称的字典，或None
    """
    code = str(stock_code).zfill(6)
    
    # 根据股票代码判断secid格式
    # 北交所: 0.代码, 深市: 0.代码, 沪市: 1.代码
    if code.startswith('6'):
        secid = f"1.{code}"
    else:
        secid = f"0.{code}"
    
    url = f"https://push2.eastmoney.com/api/qt/stock/get?secid={secid}&fields=f57,f58"
    
    try:
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            data = response.json()
            if data.get("data"):
                name = data["data"].get("f58", "")
                return {"name": name, "code": code}
    except Exception as e:
        logging.debug(f"获取行情信息失败 {code}: {e}")
    
    return None


def retry_failed_records(
    source_file: str = "公司基础信息表.xlsx",
    output_file: str = "公司基础信息表.xlsx",
    request_delay: float = 0.3,
    max_retries: int = 3
) -> None:
    """重试获取失败的记录。
    
    读取已有的输出文件，找出获取失败的记录（证券简称为空），
    重新爬取这些记录并更新到文件中。
    
    对于北交所等特殊股票，如果无法获取详细信息，会尝试获取基本信息并标记状态。
    
    Args:
        source_file: 已有的输出文件
        output_file: 更新后的输出文件（可以与source_file相同）
        request_delay: 请求间隔
        max_retries: 最大重试次数
    """
    logging.info("=" * 60)
    logging.info("重试失败记录模式")
    logging.info(f"来源文件: {source_file}")
    logging.info(f"输出文件: {output_file}")
    logging.info("=" * 60)
    
    # 读取已有文件
    if not os.path.exists(source_file):
        logging.error(f"文件不存在: {source_file}")
        return
    
    df = pd.read_excel(source_file)
    logging.info(f"读取到 {len(df)} 条记录")
    
    # 确保有备注列
    if '备注' not in df.columns:
        df['备注'] = ''
    
    # 找出失败的记录（证券简称为空）
    failed_mask = df['证券简称'].isna() | (df['证券简称'] == '')
    failed_codes = df.loc[failed_mask, '证券代码'].astype(str).str.zfill(6).tolist()
    
    if not failed_codes:
        logging.info("没有需要重试的记录")
        return
    
    logging.info(f"发现 {len(failed_codes)} 条失败记录需要重试")
    
    # 预加载城市代码
    load_city_codes()
    
    # 创建爬虫实例
    client = EastMoneyClient()
    
    # 辅助函数
    def extract_province(address: str) -> tuple:
        if not address:
            return "", ""
        for city in ['北京', '天津', '上海', '重庆']:
            if city in address:
                return city, PROVINCE_CODE_MAP.get(city, "")
        for prov_name, prov_code in PROVINCE_CODE_MAP.items():
            if prov_name in address:
                return prov_name, prov_code
        return "", ""
    
    def extract_city(address: str) -> tuple:
        if not address:
            return "", ""
        city_map = load_city_codes()
        for city in ['北京', '天津', '上海', '重庆']:
            if city in address:
                city_name = city + "市"
                city_code = city_map.get(city_name, "")
                return city_name, city_code
        clean_address = re.sub(r'^[\u4e00-\u9fa5]{2,6}(省|自治区)', '', address)
        city_match = re.search(r'(?<!省)([\u4e00-\u9fa5]{2,4}市)', clean_address)
        if city_match:
            city_name = city_match.group(1)
            if not city_name.startswith(('省', '区')):
                city_code = city_map.get(city_name, "")
                if not city_code:
                    city_code = city_map.get(city_name[:-1], "")
                return city_name, city_code
        for city_name, city_code in city_map.items():
            if city_name.endswith("市"):
                city_short = city_name[:-1]
                if len(city_short) >= 2 and city_short in clean_address:
                    return city_name, city_code
        return "", ""
    
    def get_industry_code(industry_name: str) -> str:
        if not industry_name:
            return ""
        if industry_name in CSRC_INDUSTRY_CODE_MAP:
            return CSRC_INDUSTRY_CODE_MAP[industry_name]
        if "-" in industry_name:
            parts = industry_name.split("-")
            sub_industry = parts[-1].strip()
            if sub_industry in CSRC_INDUSTRY_CODE_MAP:
                return CSRC_INDUSTRY_CODE_MAP[sub_industry]
            main_industry = parts[0].strip()
            if main_industry in CSRC_INDUSTRY_CODE_MAP:
                return CSRC_INDUSTRY_CODE_MAP[main_industry]
        for name, code in CSRC_INDUSTRY_CODE_MAP.items():
            if name in industry_name or industry_name in name:
                return code
        return ""
    
    # 重试爬取
    success_count = 0
    still_failed = 0
    total = len(failed_codes)
    
    for idx, code in enumerate(failed_codes, 1):
        data = client.get_company_info(code, max_retries)
        
        # 解析数据
        result = {
            "证券简称": "",
            "行业代码": "",
            "行业名称": "",
            "所属省份": "",
            "所属省份代码": "",
            "所属城市": "",
            "所属城市代码": "",
        }
        
        if data:
            try:
                jbzl = data.get("jbzl", {})
                if jbzl:
                    result["证券简称"] = jbzl.get("gsmc", "") or jbzl.get("agjc", "")
                    
                    csrc_industry = jbzl.get("sszjhhy", "")
                    em_industry = jbzl.get("sshy", "")
                    
                    if csrc_industry:
                        result["行业代码"] = get_industry_code(csrc_industry)
                        if "-" in csrc_industry:
                            result["行业名称"] = csrc_industry.split("-")[-1].strip()
                        else:
                            result["行业名称"] = csrc_industry
                    elif em_industry:
                        result["行业名称"] = em_industry
                        result["行业代码"] = get_industry_code(em_industry)
                    
                    address = jbzl.get("zcdz", "") or jbzl.get("bgdz", "")
                    if address:
                        province, province_code = extract_province(address)
                        city, city_code = extract_city(address)
                        result["所属省份"] = province
                        result["所属省份代码"] = province_code
                        result["所属城市"] = city
                        result["所属城市代码"] = city_code
                
                if not result["证券简称"]:
                    fxxg = data.get("fxxg", {})
                    if fxxg:
                        result["证券简称"] = fxxg.get("agdm", "")
            except Exception as e:
                logging.warning(f"解析公司 {code} 信息失败: {e}")
        
        # 更新DataFrame
        mask = df['证券代码'].astype(str).str.zfill(6) == code
        
        if result["证券简称"]:
            success_count += 1
            for key, value in result.items():
                df.loc[mask, key] = value
            
            city_info = result['所属城市'] if result['所属城市'] else result['所属省份']
            industry_info = f"{result['行业代码']}-{result['行业名称']}" if result['行业代码'] else result['行业名称']
            logging.info(f"[{idx}/{total}] {code} - {result['证券简称']} - {industry_info} - {city_info}")
        else:
            # 尝试通过行情API获取基本信息
            basic_info = get_stock_basic_info(code)
            if basic_info and basic_info.get("name"):
                name = basic_info["name"]
                df.loc[mask, '证券简称'] = name
                
                # 标记特殊状态
                if "已切换" in name or "已退市" in name:
                    df.loc[mask, '备注'] = "已转板或退市"
                    logging.warning(f"[{idx}/{total}] {code} - {name} (已转板或退市)")
                else:
                    df.loc[mask, '备注'] = "仅获取到名称，详细信息缺失"
                    logging.warning(f"[{idx}/{total}] {code} - {name} (详细信息缺失)")
                still_failed += 1
            else:
                still_failed += 1
                df.loc[mask, '备注'] = "无法获取信息"
                logging.warning(f"[{idx}/{total}] {code} - 仍然获取失败")
        
        if idx < total:
            time.sleep(request_delay)
        
        if idx % 10 == 0:
            progress = (idx / total) * 100
            print(f"\r进度: {idx}/{total} ({progress:.1f}%)", end='', flush=True)
    
    print()
    
    # 保存更新后的结果
    df.to_excel(output_file, index=False, engine='openpyxl')
    
    logging.info("=" * 60)
    logging.info(f"重试完成!")
    logging.info(f"重试记录数: {total}")
    logging.info(f"成功修复: {success_count}")
    logging.info(f"仍然失败: {still_failed}")
    logging.info(f"输出文件: {output_file}")
    logging.info("=" * 60)


if __name__ == '__main__':
    # ==================== 配置区域 ====================
    
    # 运行模式: "full" 完整爬取, "retry" 重试失败记录
    MODE = "retry"
    
    # 年份范围（会自动读取 年报链接_2012.xlsx 到 年报链接_2024.xlsx）
    START_YEAR = 2012
    END_YEAR = 2024
    
    # 年报链接文件命名模式（{year}会被替换为具体年份）
    FILE_PATTERN = "年报链接_{year}.xlsx"
    
    # 输出文件路径（所有年份合并后的唯一公司信息）
    OUTPUT_FILE = "公司基础信息表.xlsx"
    
    # 爬虫配置
    REQUEST_DELAY = 0.3  # 请求间隔（秒）
    MAX_RETRIES = 3  # 最大重试次数
    
    # ==================== 执行逻辑 ====================
    
    if MODE == "retry":
        # 重试失败记录模式
        retry_failed_records(
            source_file=OUTPUT_FILE,
            output_file=OUTPUT_FILE,
            request_delay=REQUEST_DELAY,
            max_retries=MAX_RETRIES,
        )
    else:
        # 完整爬取模式
        crawler = MultiYearCompanyInfoCrawler(
            start_year=START_YEAR,
            end_year=END_YEAR,
            output_file=OUTPUT_FILE,
            file_pattern=FILE_PATTERN,
            request_delay=REQUEST_DELAY,
            max_retries=MAX_RETRIES,
        )
        crawler.run()
