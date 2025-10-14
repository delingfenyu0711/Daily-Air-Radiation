import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import random
import datetime 
from fake_useragent import UserAgent
import warnings
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import re 

# 忽略HTTPS证书警告
warnings.filterwarnings("ignore", category=requests.packages.urllib3.exceptions.InsecureRequestWarning)

def get_radiation_data(url):
    """获取辐射监测数据，优先使用requests，失败则使用Selenium"""
    # 尝试使用requests获取（高效）
    try:
        ua = UserAgent()
        headers = {
            "User-Agent": ua.random,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
            "Connection": "keep-alive"
        }
        
        time.sleep(random.uniform(1, 3))  # 随机延迟防反爬
        response = requests.get(url, headers=headers, timeout=15, verify=False)
        response.raise_for_status()
        response.encoding = response.apparent_encoding
        print("通过requests成功获取网页内容")
        return response.text
    except Exception as e:
        print(f"requests获取失败，尝试使用Selenium: {e}")
    
    # 失败后使用Selenium模拟浏览器（处理动态页面）
    try:
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')  # 无头模式
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options
        )
        driver.get(url)
        time.sleep(2)  # 等待页面加载
        html_content = driver.page_source
        driver.quit()
        print("通过Selenium成功获取网页内容")
        return html_content
    except Exception as e:
        print(f"Selenium获取失败: {e}")
        return None

def parse_by_html_tags(html_content):
    """通过HTML标签解析数据（适用于列表结构）"""
    if not html_content:
        return []
    
    try:
        soup = BeautifulSoup(html_content, 'lxml')
        data = []
        # 查找监测点列表项（尝试多种可能的标签和类名）
        monitor_items = soup.find_all(['li', 'div'], class_=['datali', 'data-item', 'monitor-item'])
        
        if not monitor_items:
            print("未找到监测点列表项，尝试通过其他方式解析")
            return []
        
        for item in monitor_items:
            try:
                # 提取监测点名称
                name_div = item.find('div', class_=['divname', 'station-name'])
                station = name_div.text.strip() if name_div else "名称缺失"
                
                # 提取辐射值
                val_div = item.find('div', class_=['divval', 'radiation-value-div'])
                radiation_span = val_div.find('span', class_=['label', 'radiation-value']) if val_div else None
                radiation = radiation_span.text.strip() if radiation_span else "数值缺失"
                
                # 提取更新时间
                time_span = val_div.find('span', class_=['showtime', 'update-time']) if val_div else None
                time_str = time_span.text.strip() if time_span else "时间缺失"
                
                # 提取省份
                province = station.split(' (')[0] if ' (' in station else "省份未知"
                
                data.append({
                    "省份": province,
                    "监测点": station,
                    "辐射值": radiation,
                    "更新时间": time_str
                })
            except Exception as e:
                print(f"解析单个监测点（标签方式）出错: {e}")
        
        return data
    except Exception as e:
        print(f"标签解析整体出错: {e}")
        return []

def parse_by_text_lines(html_content):
    """通过文本行解析数据（适用于纯文本结构）"""
    if not html_content:
        return []
    
    try:
        soup = BeautifulSoup(html_content, 'lxml')
        data = []
        # 提取所有文本行并过滤空行
        lines = soup.get_text().split('\n')
        valid_lines = [line.strip() for line in lines if line.strip()]
        
        # 查找数据起始位置
        start_index = -1
        for i, line in enumerate(valid_lines):
            if "省会城市空气吸收剂量率 监测值" in line:
                start_index = i + 2
                break
        
        if start_index == -1:
            print("未找到数据起始标记，尝试从开头解析")
            start_index = 0
        
        i = start_index
        while i < len(valid_lines):
            # 跳过空行和非数据行
            if len(valid_lines[i]) < 3 or "nGy/h" not in valid_lines[i+1:i+2]:
                i += 1
                continue
            
            # 提取监测点
            station = valid_lines[i].strip()
            i += 1
            
            # 提取辐射值（确保包含nGy/h）
            radiation = ""
            while i < len(valid_lines) and "nGy/h" not in valid_lines[i]:
                i += 1
            if i < len(valid_lines):
                radiation = valid_lines[i].strip()
                i += 1
            else:
                break
            
            # 提取时间（格式为YYYY-MM-DD）
            time_str = ""
            while i < len(valid_lines):
                if len(valid_lines[i]) == 10 and '-' in valid_lines[i]:
                    time_str = valid_lines[i].strip()
                    break
                i += 1
            if not time_str:
                time_str = "时间缺失"
            i += 1  # 移动到下一组开始
            
            # 提取省份
            province = station.split(' (')[0] if ' (' in station else "省份未知"
            
            data.append({
                "省份": province,
                "监测点": station,
                "辐射值": radiation,
                "更新时间": time_str
            })
        
        return data
    except Exception as e:
        print(f"文本行解析出错: {e}")
        return []

def parse_html(html_content):
    """智能选择解析方式，先尝试标签解析，再尝试文本行解析"""
    # 先尝试标签解析（更准确）
    data = parse_by_html_tags(html_content)
    if data:
        print(f"通过标签解析获取{len(data)}条数据")
        return data
    
    # 再尝试文本行解析
    data = parse_by_text_lines(html_content)
    if data:
        print(f"通过文本行解析获取{len(data)}条数据")
        return data
    
    print("两种解析方式均未获取到有效数据")
    return []

def save_to_excel(data, filename_prefix="辐射监测数据"):
    """保存数据到Excel，正确设置sheet名称"""
    if not data:
        print("没有数据可保存")
        return False
    
    try:
        # 获取当前时间和日期
        current_time = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        sheet_date = datetime.datetime.now().strftime("%Y%m%d")
        
        # 生成文件名和规范sheet名称
        filename = f"{filename_prefix}_{current_time}.xlsx"
        sheet_name = re.sub(r'[\\/:*?"<>|]', '', sheet_date)
        sheet_name = sheet_name if sheet_name else "辐射数据"
        
        # 保存数据
        df = pd.DataFrame(data)
        df.to_excel(filename, index=False, sheet_name=sheet_name)
        print(f"数据已保存至 {filename}，sheet名称: {sheet_name}")
        return True
    except Exception as e:
        print(f"保存文件出错: {e}")
        return False


def display_data(data):
    """以表格形式显示数据"""
    if not data:
        print("没有可显示的数据")
        return
    
    print("\n=== 省会城市空气吸收剂量率监测数据 ===")
    print(f"{'省份':<10}{'监测点':<30}{'辐射值':<15}{'更新时间':<15}")
    print("-" * 70)
    
    for item in data:
        print(f"{item['省份']:<10}{item['监测点']:<30}{item['辐射值']:<15}{item['更新时间']:<15}")

def main():
    """主函数"""
    url = "https://data.rmtc.org.cn/gis/listtype0M.html"
    print(f"正在获取数据: {url}")
    
    html_content = get_radiation_data(url)
    if not html_content:
        print("获取数据失败，程序退出")
        return
    
    data = parse_html(html_content)
    if not data:
        print("解析数据失败，未找到有效数据")
        return
    
    display_data(data)
    save_to_excel(data)

if __name__ == "__main__":
    main()