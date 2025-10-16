import os
import sys
import time
import random
import datetime
import threading
import subprocess
import configparser
from pathlib import Path
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import requests
from bs4 import BeautifulSoup, Tag
import pandas as pd
from fake_useragent import UserAgent
import warnings
import gc

# 忽略HTTPS证书警告
warnings.filterwarnings("ignore", category=requests.packages.urllib3.exceptions.InsecureRequestWarning)

# 全局配置
CONFIG = configparser.ConfigParser()
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
CONFIG_PATH = "config.ini"

# ------------------------------
# 新增：通用工具函数（强制类型转换，杜绝列表）
# ------------------------------
def safe_str(val, default="未知"):
    """
    安全转换为字符串：
    - 若val是列表/元组，取第一个元素并转字符串
    - 若val是空值/None，返回默认值
    - 其他情况直接转字符串
    """
    if isinstance(val, (list, tuple)):
        return str(val[0]) if val else default
    elif val is None or str(val).strip() == "":
        return default
    else:
        return str(val).strip()


# ------------------------------
# 1. 配置文件处理
# ------------------------------
def load_config():
    if not os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            f.write("""[CRAWLER]
crawl_time = 10:00
target_url = https://data.rmtc.org.cn/gis/listtype0M.html
random_delay = 1,3
file_prefix = 辐射监测数据

[GIT]
commit_prefix = 自动更新：
enable_push = True

[LOG]
log_max_lines = 100
""")
    CONFIG.read(CONFIG_PATH, encoding="utf-8")
    return CONFIG


# ------------------------------
# 2. 数据抓取与解析（根治list.find错误）
# ------------------------------
def get_radiation_data(url, min_delay, max_delay):
    try:
        time.sleep(random.uniform(min_delay, max_delay))
        # 安全获取UA：用safe_str确保是字符串
        try:
            ua = UserAgent().random
        except:
            ua = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        ua = safe_str(ua)  # 强制转为字符串，防止列表
        headers = {"User-Agent": ua}
        
        response = requests.get(url, headers=headers, timeout=15, verify=False)
        response.raise_for_status()
        html = safe_str(response.text)  # 网页内容强制字符串
        response.close()
        return html
    except Exception as e:
        print(f"请求失败：{safe_str(e)}")
        return None


def parse_html(html_content):
    data = []
    html_content = safe_str(html_content)  # 先确保输入是字符串
    if html_content == "未知":
        return data

    # 解析后立即释放大对象
    soup = BeautifulSoup(html_content, 'lxml')
    monitor_containers = soup.select('.datali')
    soup = None
    html_content = None

    if not monitor_containers:
        gc.collect()
        return data

    for container in monitor_containers:
        try:
            # 步骤1：仅保留Tag类型的div子标签（排除所有非Tag元素）
            child_divs = []
            for div in container.children:
                if isinstance(div, Tag) and div.name == 'div' and div.has_attr('class'):
                    child_divs.append(div)
            if len(child_divs) < 2:
                container = None
                continue

            # 步骤2：提取监测点名称（safe_str确保无列表）
            name_div_list = [d for d in child_divs if 'divname' in safe_str(d.get('class', []))]
            station = "名称缺失"
            if name_div_list and isinstance(name_div_list[0], Tag):
                station_text = name_div_list[0].get_text(strip=True)
                station = safe_str(station_text, "名称缺失")  # 安全转换
            name_div_list = None

            # 步骤3：提取辐射值和时间（全流程safe_str）
            val_div_list = [d for d in child_divs if 'divval' in safe_str(d.get('class', []))]
            radiation = "数值缺失"
            time_str = "时间缺失"
            if val_div_list and isinstance(val_div_list[0], Tag):
                val_spans = []
                for span in val_div_list[0].children:
                    if isinstance(span, Tag) and span.name == 'span':
                        val_spans.append(span)
                
                # 辐射值：双重保障（Tag判断 + safe_str）
                if len(val_spans) >= 1 and isinstance(val_spans[0], Tag):
                    rad_text = val_spans[0].get_text(strip=True)
                    radiation = safe_str(rad_text, "数值缺失")
                # 时间：同上
                if len(val_spans) >= 2 and isinstance(val_spans[1], Tag):
                    time_text = val_spans[1].get_text(strip=True)
                    time_str = safe_str(time_text, "时间缺失")
                val_spans = None
            val_div_list = None
            child_divs = None

            # 步骤4：提取省份（仅对安全字符串操作）
            province = "省份未知"
            if ' (' in station:
                province = safe_str(station.split(' (')[0], "省份未知")

            # 添加数据：所有字段经safe_str处理，绝对无列表
            data.append({
                "省份": province,
                "监测点": station,
                "辐射值": radiation,
                "更新时间": time_str
            })

            # 释放当前循环变量
            container = None
            province = None
            station = None
            radiation = None
            time_str = None

        except Exception as e:
            # 打印详细错误，便于追溯（包含当前变量类型）
            error_detail = f"解析失败：{safe_str(e)} | 容器类型：{type(container)} | 辐射值类型：{type(radiation)}"
            print(error_detail)
            container = None
            continue

    gc.collect()
    return data


def save_to_excel(data, file_prefix):
    file_prefix = safe_str(file_prefix, "辐射监测数据")  # 安全转换
    if not data or not isinstance(data, list):
        return None
    try:
        # 文件名全安全处理
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = DATA_DIR / f"{file_prefix}_{timestamp}.xlsx"
        df = pd.DataFrame(data)
        df.to_excel(safe_str(filename), index=False, sheet_name="辐射数据")
        df = None
        data = None
        gc.collect()
        return safe_str(filename)
    except Exception as e:
        print(f"保存Excel失败：{safe_str(e)}")
        data = None
        gc.collect()
        return None


def git_commit_push(file_path, commit_prefix):
    file_path = safe_str(file_path)
    commit_prefix = safe_str(commit_prefix, "自动更新：")
    if not file_path or not os.path.exists(file_path):
        return False
    try:
        commit_msg = f"{commit_prefix}{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')} 的数据"
        subprocess.run(['git', 'add', file_path], capture_output=True, check=True)
        subprocess.run(['git', 'commit', '-m', commit_msg], capture_output=True, check=True)
        subprocess.run(['git', 'push'], capture_output=True, check=True)
        return True
    except Exception as e:
        print(f"Git操作失败：{safe_str(e)}")
        return False


# ------------------------------
# 3. 定时任务与UI（彻底杜绝类型异常）
# ------------------------------
def fetch_data_task(callback=None, task_type="定时"):
    def log(msg, is_error=False):
        if callback and callable(callback):
            # 日志信息绝对安全：safe_str + 空值处理
            msg_str = safe_str(msg, "未知日志信息")
            callback(f"[{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {'[错误]' if is_error else ''} {msg_str}")

    log(f"=== 开始{task_type}抓取任务 ===")
    try:
        config = load_config()
        # 配置项全量安全转换，杜绝非字符串
        url = safe_str(config.get("CRAWLER", "target_url", fallback="https://data.rmtc.org.cn/gis/listtype0M.html"))
        delay_str = safe_str(config.get("CRAWLER", "random_delay", fallback="1,3"))
        file_prefix = safe_str(config.get("CRAWLER", "file_prefix", fallback="辐射监测数据"))
        git_enable = config.getboolean("GIT", "enable_push", fallback=True)
        git_prefix = safe_str(config.get("GIT", "commit_prefix", fallback="自动更新："))
        config = None

        # 解析延迟：安全处理（避免列表/空值）
        delay_list = [safe_str(d) for d in delay_str.split(',')]
        min_delay = int(delay_list[0]) if len(delay_list)>=1 and delay_list[0].isdigit() else 1
        max_delay = int(delay_list[1]) if len(delay_list)>=2 and delay_list[1].isdigit() else 3
        min_delay, max_delay = max(1, min_delay), max(min_delay, max_delay)  # 防止异常值

        # 1. 获取网页
        log(f"请求URL：{url[:50]}...")
        html = get_radiation_data(url, min_delay, max_delay)
        if html == "未知":
            log(f"{task_type}抓取失败：无法获取网页", is_error=True)
            gc.collect()
            return

        # 2. 解析数据
        log("解析数据中...")
        data = parse_html(html)
        html = None
        if not data:
            log(f"{task_type}抓取失败：未找到有效监测数据", is_error=True)
            gc.collect()
            return
        log(f"成功解析{len(data)}条监测点数据")

        # 3. 保存数据
        log("保存数据中...")
        file_path = save_to_excel(data, file_prefix)
        data = None
        if file_path == "未知":
            log(f"{task_type}抓取失败：数据保存失败", is_error=True)
            gc.collect()
            return
        log(f"数据保存路径：{os.path.basename(file_path)}")

        # 4. Git推送
        if git_enable:
            log("推送数据至Git仓库...")
            if git_commit_push(file_path, git_prefix):
                log("Git推送成功")
            else:
                log("Git推送失败（请检查仓库配置）", is_error=True)
        else:
            log("Git推送已禁用")

        log(f"=== {task_type}抓取任务完成 ===")
    except Exception as e:
        # 错误信息安全转换，避免异常本身是列表
        error_msg = safe_str(str(e)[:100], "未知异常")
        log(f"{task_type}任务异常：{error_msg}", is_error=True)
    finally:
        gc.collect()


class CrawlerUI:
    def __init__(self, root):
        self.root = root
        self.root.title("辐射数据自动抓取工具")
        self.root.geometry("900x600")
        self.stop_event = threading.Event()
        self.schedule_thread = None
        self._init_ui()
        self._start_schedule()
        self._refresh_config()

    def _init_ui(self):
        # 配置显示
        config_frame = ttk.LabelFrame(self.root, text="当前配置", padding=(10,5))
        config_frame.pack(fill=tk.X, padx=10, pady=5)
        self.config_vars = {
            "crawl_time": tk.StringVar(), "target_url": tk.StringVar(),
            "random_delay": tk.StringVar(), "git_status": tk.StringVar()
        }
        ttk.Label(config_frame, text="定时时间：").grid(row=0, column=0, sticky=tk.W, padx=5)
        ttk.Label(config_frame, textvariable=self.config_vars["crawl_time"]).grid(row=0, column=1, sticky=tk.W)
        ttk.Label(config_frame, text="目标URL：").grid(row=1, column=0, sticky=tk.W, padx=5)
        ttk.Label(config_frame, textvariable=self.config_vars["target_url"], font=("Consolas", 9)).grid(row=1, column=1, sticky=tk.W)
        ttk.Label(config_frame, text="请求延迟：").grid(row=2, column=0, sticky=tk.W, padx=5)
        ttk.Label(config_frame, textvariable=self.config_vars["random_delay"]).grid(row=2, column=1, sticky=tk.W)
        ttk.Label(config_frame, text="Git推送：").grid(row=3, column=0, sticky=tk.W, padx=5)
        ttk.Label(config_frame, textvariable=self.config_vars["git_status"]).grid(row=3, column=1, sticky=tk.W)

        # 操作按钮
        btn_frame = ttk.Frame(self.root, padding=(10,5))
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        self.crawl_btn = ttk.Button(btn_frame, text="手动执行抓取", command=self._manual_crawl)
        self.crawl_btn.pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="打开配置文件", command=self._open_config).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="清空日志", command=self._clear_log).pack(side=tk.RIGHT, padx=5)

        # 日志区域
        log_frame = ttk.LabelFrame(self.root, text="操作日志", padding=(10,5))
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.log_text = scrolledtext.ScrolledText(log_frame, font=("Consolas", 10), bg="#2c3e50", fg="#ecf0f1")
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.log_text.config(state=tk.DISABLED)
        self.log_text.tag_configure("error", foreground="#e74c3c")

    def _refresh_config(self):
        try:
            config = load_config()
            # 配置值全量安全转换，杜绝列表
            crawl_time = safe_str(config.get("CRAWLER", "crawl_time", fallback="10:00"))
            target_url = safe_str(config.get("CRAWLER", "target_url", fallback="未知"))
            random_delay = safe_str(config.get("CRAWLER", "random_delay", fallback="1,3"))
            git_status = "启用" if config.getboolean("GIT", "enable_push", fallback=True) else "禁用"
            
            self.config_vars["crawl_time"].set(crawl_time)
            self.config_vars["target_url"].set(f"{target_url[:50]}..." if target_url != "未知" else "未知...")
            self.config_vars["random_delay"].set(f"{random_delay} 秒")
            self.config_vars["git_status"].set(git_status)
            config = None
        except Exception as e:
            self._log(f"配置刷新失败：{safe_str(e)}", is_error=True)
        if not self.stop_event.is_set():
            self.root.after(10000, self._refresh_config)

    def _start_schedule(self):
        def schedule_loop():
            while not self.stop_event.is_set():
                try:
                    config = load_config()
                    crawl_time = safe_str(config.get("CRAWLER", "crawl_time", fallback="10:00"))
                    config = None
                    
                    now = datetime.datetime.now()
                    try:
                        target_time = datetime.datetime.strptime(f"{now.date()} {crawl_time}", "%Y-%m-%d %H:%M")
                    except ValueError:
                        target_time = now + datetime.timedelta(minutes=5)
                        self._log("时间格式错误（应为HH:MM），5分钟后重试", is_error=True)
                    if now >= target_time:
                        target_time += datetime.timedelta(days=1)
                    self._log(f"定时任务启动，下次执行：{target_time.strftime('%Y-%m-%d %H:%M')}")
                    
                    # 非阻塞等待：每0.5秒检查终止信号
                    remaining_seconds = (target_time - datetime.datetime.now()).total_seconds()
                    while remaining_seconds > 0 and not self.stop_event.is_set():
                        sleep_time = min(0.5, remaining_seconds)
                        time.sleep(sleep_time)
                        remaining_seconds -= sleep_time
                    
                    if not self.stop_event.is_set():
                        fetch_data_task(callback=self._log, task_type="定时")
                except Exception as e:
                    self._log(f"定时任务异常：{safe_str(str(e)[:80])}", is_error=True)
                    # 异常时仍检查终止信号
                    for _ in range(120):
                        if self.stop_event.is_set():
                            break
                        time.sleep(0.5)
                finally:
                    now = None
                    target_time = None
                    crawl_time = None
                    gc.collect()

        self.schedule_thread = threading.Thread(target=schedule_loop, daemon=True)
        self.schedule_thread.start()

    def _manual_crawl(self):
        if self.crawl_btn["state"] == tk.DISABLED or self.stop_event.is_set():
            return
        self.crawl_btn.config(state=tk.DISABLED)
        
        def run():
            fetch_data_task(callback=self._log, task_type="手动")
            self.crawl_btn.config(state=tk.NORMAL)
            gc.collect()
        
        threading.Thread(target=run, daemon=True).start()

    def _open_config(self):
        try:
            config_path = safe_str(CONFIG_PATH)
            if sys.platform == "win32":
                os.startfile(config_path)
                self._log(f"已打开配置文件：{config_path}")
            else:
                subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', config_path])
        except Exception as e:
            self._log(f"打开配置失败：{safe_str(e)}", is_error=True)

    def _clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self._log("日志已清空")

    def _log(self, msg, is_error=False):
        if not self.root.winfo_exists() or self.stop_event.is_set():
            return
        self.log_text.config(state=tk.NORMAL)
        # 日志信息绝对安全：双重safe_str
        msg_str = safe_str(safe_str(msg), "未知日志")
        self.log_text.insert(tk.END, msg_str + "\n", "error" if is_error else "")
        
        # 计算行数：全安全处理
        config = load_config()
        max_lines = int(safe_str(config.get("LOG", "log_max_lines", fallback=100), 100))
        config = None
        current_lines = int(self.log_text.count('1.0', tk.END, "lines")[0])
        
        if current_lines > max_lines:
            delete_end = f"{current_lines - max_lines}.0"
            self.log_text.delete('1.0', delete_end)
        
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)

    def close(self):
        """终极关闭方案：强制终止进程"""
        self.stop_event.set()
        os._exit(0)


# ------------------------------
# 4. 主程序入口
# ------------------------------
def main():
    load_config()
    root = tk.Tk()
    app = CrawlerUI(root)
    root.protocol("WM_DELETE_WINDOW", app.close)
    root.mainloop()


if __name__ == "__main__":
    main()