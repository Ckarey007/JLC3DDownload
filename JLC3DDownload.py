"""
嘉立创 3D 模型下载器 (JLC 3D Model Downloader)

功能概览：
  - 单个下载：输入元器件编号（如 C8734），一键下载对应 STEP 3D 模型
  - 批量下载：导入 Excel (.xlsx) 或 CSV 文件，自动识别所有元器件编号并逐个下载
    - 下载间隔随机 10~30 秒，避免对服务器造成压力
    - 支持随时取消，安全中断
    - 自动去重，跳过重复编号
  - 下载路径持久化记忆
  - 一键定位已下载文件

版本：1.3
作者：Jupiter
贡献者：Ckare（新增批量下载功能）
项目地址：https://github.com/zhutongxueya/JLC3DDownload

新增功能（v1.3 - 批量下载）：
  1. 支持从 Excel (.xlsx/.xls) 和 CSV 文件中导入元器件编号
  2. 自动扫描所有单元格，识别 "C + 数字" 格式的编号（如 C8734）
  3. 每次下载后随机等待 10~30 秒，降低服务器负载
  4. 日志实时显示进度 [当前/总数]，下载完成后汇总统计
  5. 支持中途取消，当前器件下载完成后安全停止

依赖：
  pip install requests openpyxl
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import requests
import json
import os
import threading
import configparser
import csv          # 批量下载：读取 CSV 文件
import random       # 批量下载：随机等待间隔
import time         # 批量下载：控制下载间隔
from datetime import datetime
import webbrowser
import subprocess
import sys
import openpyxl     # 批量下载：读取 Excel 文件

# =====================================================
# 配置与逻辑处理
# =====================================================
# 应用配置目录，用于持久化保存用户的下载路径
APP_DIR = os.path.join(os.path.expanduser("~"), ".jlc3d")
CONFIG_FILE = os.path.join(APP_DIR, "config.ini")


def get_runtime_dir():
    """Return the packaged exe directory, or the script directory."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def ensure_app_dir():
    """确保应用配置目录存在"""
    if not os.path.exists(APP_DIR):
        os.makedirs(APP_DIR)


def save_download_path(path):
    """将用户选择的下载路径保存到配置文件"""
    ensure_app_dir()
    cfg = configparser.ConfigParser()
    cfg["PATH"] = {"DownloadPath": path}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        cfg.write(f)


def load_download_path():
    """从配置文件中读取上次保存的下载路径"""
    if not os.path.exists(CONFIG_FILE):
        return None
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_FILE, encoding="utf-8")
    return cfg.get("PATH", "DownloadPath", fallback=None)


def default_download_dir():
    """Use the current runtime directory as the default download location."""
    return get_runtime_dir()


# =====================================================
# 嘉立创 API 逻辑
# =====================================================
BASE_API = "https://pro.lceda.cn/api"
HEADERS = {"User-Agent": "Mozilla/5.0"}


def search_product(code):
    """根据元器件编号搜索产品，返回设备 UUID"""
    r = requests.post(
        f"{BASE_API}/eda/product/search",
        data={"keyword": code, "needAggs": "true", "currPage": "1", "pageSize": "10"},
        headers=HEADERS, timeout=10
    )
    r.raise_for_status()
    products = r.json()["result"]["productList"]
    if not products:
        raise ValueError("未找到该器件")
    return products[0]["hasDevice"]


def get_model_uuid(device_uuid):
    """根据设备 UUID 获取 3D 模型的 UUID"""
    r = requests.post(
        f"{BASE_API}/devices/searchByIds",
        data={"uuids[]": device_uuid},
        headers=HEADERS, timeout=10
    )
    r.raise_for_status()
    attrs = r.json()["result"][0]["attributes"]
    if "3D Model" not in attrs:
        raise ValueError("该器件没有 3D 模型")
    return attrs["3D Model"]


def get_model_file(model_uuid):
    """根据模型 UUID 获取 STEP 文件的下载路径标识"""
    r = requests.post(
        f"{BASE_API}/components/searchByIds?forceOnline=1",
        data={"uuids[]": model_uuid, "dataStr": "yes"},
        headers=HEADERS, timeout=10
    )
    r.raise_for_status()
    data = json.loads(r.json()["result"][0]["dataStr"])
    return data["model"]


def download_step_file(model_file):
    """从嘉立创 CDN 下载 STEP 模型文件的二进制内容"""
    url = f"https://modules.lceda.cn/qAxj6KHrDKw4blvCG8QJPs7Y/{model_file}"
    r = requests.get(url, headers=HEADERS, timeout=15)
    r.raise_for_status()
    return r.content


# =====================================================
# 主程序 UI
# =====================================================
class JLC3DApp(tk.Tk):
    """嘉立创 3D 模型下载器主窗口"""

    def __init__(self):
        super().__init__()
        self.title("嘉立创 3D 模型下载器")

        # 窗口大小及居中（v1.3 加宽以容纳批量下载按钮）
        self.app_width = 650
        self.app_height = 480
        self.center_window(self, self.app_width, self.app_height)

        self.configure(bg="#f8f9fa")
        self.download_path = default_download_dir()
        self.last_download_file = None

        # --- 批量下载状态 ---
        self.batch_running = False   # 批量下载是否正在运行
        self.batch_cancel = False    # 用户是否请求取消批量下载

        self._build_menu()
        self._build_ui()

    def center_window(self, target, width, height):
        """将指定窗口居中显示"""
        screen_width = target.winfo_screenwidth()
        screen_height = target.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        target.geometry(f"{width}x{height}+{x}+{y}")

    def _build_menu(self):
        """构建原生菜单栏"""
        menu = tk.Menu(self)
        self.config(menu=menu)

        file_menu = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="查看下载路径", command=self.choose_path)
        file_menu.add_command(label="批量导入 (Excel/CSV)", command=self.import_batch_file)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.quit)

        help_menu = tk.Menu(menu, tearoff=0)
        menu.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="关于", command=self.show_about)

    def _build_ui(self):
        """构建主界面布局"""
        main_container = tk.Frame(self, bg="#f8f9fa", padx=20, pady=10)
        main_container.pack(fill="both", expand=True)

        # 1. 操作行：编号输入 + 下载按钮 + 批量按钮 + 取消按钮
        center_row = tk.Frame(main_container, bg="#f8f9fa")
        center_row.pack(pady=(40, 40))

        tk.Label(
            center_row, text="元器件编号:", bg="#f8f9fa",
            font=("Microsoft YaHei", 12, "bold")
        ).pack(side="left", padx=(0, 10))

        self.entry = ttk.Entry(center_row, font=("Arial", 14), width=14)
        self.entry.insert(0, "C8734")
        self.entry.pack(side="left", padx=5)

        # 单个下载按钮
        self.btn_download = tk.Button(
            center_row, text="立即下载", bg="#28a745", fg="white",
            font=("Microsoft YaHei", 11, "bold"), relief="flat",
            width=10, height=1, cursor="hand2", command=self.start_download
        )
        self.btn_download.pack(side="left", padx=10)

        # [新增] 批量下载按钮 - 打开文件选择对话框导入 Excel/CSV
        self.btn_batch = tk.Button(
            center_row, text="批量下载", bg="#007bff", fg="white",
            font=("Microsoft YaHei", 11, "bold"), relief="flat",
            width=10, height=1, cursor="hand2", command=self.import_batch_file
        )
        self.btn_batch.pack(side="left", padx=5)

        # [新增] 取消批量下载按钮 - 仅在批量下载进行时可用
        self.btn_cancel = tk.Button(
            center_row, text="取消批量", bg="#dc3545", fg="white",
            font=("Microsoft YaHei", 11, "bold"), relief="flat",
            width=10, height=1, cursor="hand2", command=self.cancel_batch,
            state="disabled"
        )
        self.btn_cancel.pack(side="left", padx=5)

        # 2. 日志区
        self.log = scrolledtext.ScrolledText(
            main_container, height=8, font=("Consolas", 10),
            bg="white", relief="solid", borderwidth=1
        )
        self.log.pack(fill="both", expand=True)

        # 3. 底部栏：显示当前保存路径 + 定位文件按钮
        bottom_frame = tk.Frame(main_container, bg="#f8f9fa", pady=15)
        bottom_frame.pack(fill="x")

        self.path_label = tk.Label(
            bottom_frame, text=f"保存至: {self.download_path}",
            bg="#f8f9fa", fg="#495057", font=("Microsoft YaHei", 10), anchor="w"
        )
        self.path_label.pack(side="left", fill="x", expand=True)

        btn_style = ttk.Style()
        btn_style.configure("Small.TButton", font=("Microsoft YaHei", 10))
        ttk.Button(
            bottom_frame, text="定位文件", style="Small.TButton",
            width=10, command=self.locate_file
        ).pack(side="right")

    # =====================================================
    # 通用方法
    # =====================================================

    def log_msg(self, msg):
        """向日志区追加一条带时间戳的消息"""
        now = datetime.now().strftime("%H:%M:%S")
        self.log.insert(tk.END, f"[{now}] {msg}\n")
        self.log.see(tk.END)

    def choose_path(self):
        """弹出目录选择对话框，修改下载保存路径"""
        self.download_path = default_download_dir()
        self.path_label.config(text=f"保存至: {self.download_path}")
        messagebox.showinfo("提示", f"下载目录固定为程序当前目录：\n{self.download_path}")

    def locate_file(self):
        """在系统文件管理器中定位最近下载的文件"""
        if not self.last_download_file or not os.path.exists(self.last_download_file):
            messagebox.showinfo("提示", "还没有可定位的下载文件")
            return
        path = self.last_download_file
        if sys.platform.startswith("win"):
            subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
        elif sys.platform.startswith("darwin"):
            subprocess.call(["open", "-R", path])
        else:
            subprocess.call(["xdg-open", os.path.dirname(path)])

    def show_about(self):
        """显示"关于"对话框"""
        about = tk.Toplevel(self)
        about.title("关于")
        self.center_window(about, 400, 250)
        about.resizable(False, False)
        text = tk.Text(about, wrap="word", padx=15, pady=15, font=("Microsoft YaHei", 11))
        text.pack(fill="both", expand=True)
        content = (
            "嘉立创 3D 模型下载器\n"
            "版本：1.3\n"
            "作者：Jupiter\n\n"
            "项目地址：\n"
            "https://github.com/zhutongxueya/JLC3DDownload\n\n"
            "感谢 kulya97 的原始思路"
        )
        text.insert("1.0", content)
        text.config(state="disabled")
        text.tag_add("link", "5.0", "5.end")
        text.tag_config("link", foreground="blue", underline=True)
        text.tag_bind("link", "<Button-1>",
                       lambda e: webbrowser.open("https://github.com/zhutongxueya/JLC3DDownload"))

    # =====================================================
    # 单个下载
    # =====================================================

    def start_download(self):
        """启动单个下载（在后台线程中执行，避免阻塞 UI）"""
        self.btn_download.config(state="disabled", bg="#6c757d", text="下载中...")
        threading.Thread(target=self.download_task, daemon=True).start()

    def download_task(self):
        """单个下载的后台任务：搜索器件 → 获取模型 UUID → 下载 STEP 文件"""
        code = self.entry.get().strip()
        if not code:
            self.after(0, lambda: messagebox.showwarning("提示", "请输入元器件编号"))
            self.after(0, lambda: self.btn_download.config(state="normal", bg="#28a745", text="立即下载"))
            return
        try:
            self.after(0, lambda: self.log_msg(f"搜索器件【{code}】…"))
            device = search_product(code)
            self.after(0, lambda: self.log_msg("解析模型 ID…"))
            model_uuid = get_model_uuid(device)
            self.after(0, lambda: self.log_msg("下载 STEP 文件…"))
            model_file = get_model_file(model_uuid)
            data = download_step_file(model_file)
            os.makedirs(self.download_path, exist_ok=True)
            filepath = os.path.join(self.download_path, f"{code}.step")
            with open(filepath, "wb") as f:
                f.write(data)
            self.last_download_file = filepath
            self.after(0, lambda: self.log_msg(f"下载完成 ✔\n保存至：{filepath}"))
        except Exception as e:
            self.after(0, lambda: self.log_msg(f"错误：{e}"))
        finally:
            self.after(0, lambda: self.btn_download.config(state="normal", bg="#28a745", text="立即下载"))

    # =====================================================
    # [新增] 批量下载功能
    #
    # 使用方式：
    #   1. 准备一个 Excel (.xlsx) 或 CSV 文件，其中包含元器件编号
    #   2. 点击"批量下载"按钮或菜单"文件 → 批量导入"
    #   3. 程序会自动扫描文件中所有单元格，提取 C+数字 格式的编号
    #   4. 逐个下载，每次间隔随机 10~30 秒，避免服务器压力
    #   5. 下载过程中可随时点击"取消批量"安全中断
    # =====================================================

    def parse_codes_from_file(self, filepath):
        """
        从 Excel 或 CSV 文件中提取元器件编号列表。

        识别规则：扫描所有单元格，匹配 "C" + 纯数字 格式（如 C8734, C14663）。
        自动去重并保持原始顺序。

        Args:
            filepath: Excel (.xlsx/.xls) 或 CSV (.csv) 文件路径

        Returns:
            去重后的元器件编号列表，如 ["C8734", "C7171", "C14663"]
        """
        codes = []
        ext = os.path.splitext(filepath)[1].lower()

        if ext in (".xlsx", ".xls"):
            # 使用 openpyxl 读取 Excel，只读模式更节省内存
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    if cell is not None:
                        val = str(cell).strip()
                        if val.upper().startswith("C") and val[1:].isdigit():
                            codes.append(val)
            wb.close()
        elif ext == ".csv":
            # 使用 utf-8-sig 编码兼容带 BOM 的 CSV（Excel 导出常见）
            with open(filepath, "r", encoding="utf-8-sig") as f:
                reader = csv.reader(f)
                for row in reader:
                    for cell in row:
                        val = cell.strip()
                        if val.upper().startswith("C") and val[1:].isdigit():
                            codes.append(val)
        else:
            raise ValueError("不支持的文件格式，请使用 .xlsx 或 .csv")

        # 去重并保持顺序
        seen = set()
        unique = []
        for c in codes:
            if c not in seen:
                seen.add(c)
                unique.append(c)
        return unique

    def import_batch_file(self):
        """打开文件选择对话框，导入 Excel/CSV 并启动批量下载"""
        if self.batch_running:
            messagebox.showinfo("提示", "批量下载正在进行中，请等待完成或取消后再试")
            return

        filepath = filedialog.askopenfilename(
            title="选择包含元器件编号的文件",
            filetypes=[("Excel/CSV 文件", "*.xlsx *.xls *.csv"), ("所有文件", "*.*")]
        )
        if not filepath:
            return

        try:
            codes = self.parse_codes_from_file(filepath)
        except Exception as e:
            messagebox.showerror("导入失败", f"读取文件出错：{e}")
            return

        if not codes:
            messagebox.showwarning("提示", "文件中未找到有效的元器件编号（如 C8734）")
            return

        # 弹窗确认，显示识别到的编号数量和预计耗时
        confirm = messagebox.askyesno(
            "确认批量下载",
            f"共识别到 {len(codes)} 个元器件编号，下载间隔 10~30 秒。\n"
            f"预计耗时 {len(codes) * 20 // 60} ~ {len(codes) * 30 // 60 + 1} 分钟。\n\n"
            f"前 10 个：{', '.join(codes[:10])}{'...' if len(codes) > 10 else ''}\n\n"
            f"是否开始下载？"
        )
        if not confirm:
            return
        self.start_batch_download(codes)

    def start_batch_download(self, codes):
        """初始化批量下载状态并启动后台线程"""
        self.batch_running = True
        self.batch_cancel = False
        # 批量下载期间禁用其他下载按钮，防止冲突
        self.btn_batch.config(state="disabled", bg="#6c757d")
        self.btn_download.config(state="disabled", bg="#6c757d")
        self.btn_cancel.config(state="normal")
        threading.Thread(target=self.batch_download_task, args=(codes,), daemon=True).start()

    def cancel_batch(self):
        """用户请求取消批量下载，设置标志位，当前器件完成后安全停止"""
        self.batch_cancel = True
        self.after(0, lambda: self.log_msg("⏹ 用户已请求取消，将在当前下载完成后停止…"))

    def batch_download_task(self, codes):
        """
        批量下载的后台任务。

        逐个下载列表中的元器件 3D 模型，每次下载后随机等待 10~30 秒。
        等待期间每秒检查取消标志，确保能及时响应用户的取消请求。
        """
        total = len(codes)
        success = 0
        fail = 0
        self.after(0, lambda: self.log_msg(f"═══ 批量下载开始，共 {total} 个器件 ═══"))

        for i, code in enumerate(codes, 1):
            # 检查取消标志
            if self.batch_cancel:
                self.after(0, lambda s=success, f=fail, skip=total - i + 1:
                           self.log_msg(f"═══ 批量下载已取消 ({s} 成功 / {f} 失败 / {skip} 跳过) ═══"))
                break

            self.after(0, lambda c=code, idx=i: self.log_msg(f"[{idx}/{total}] 开始下载 {c}"))
            try:
                device = search_product(code)
                model_uuid = get_model_uuid(device)
                model_file = get_model_file(model_uuid)
                data = download_step_file(model_file)
                os.makedirs(self.download_path, exist_ok=True)
                filepath = os.path.join(self.download_path, f"{code}.step")
                with open(filepath, "wb") as f:
                    f.write(data)
                self.last_download_file = filepath
                success += 1
                self.after(0, lambda c=code, idx=i: self.log_msg(f"[{idx}/{total}] {c} 下载完成 ✔"))
            except Exception as e:
                fail += 1
                self.after(0, lambda c=code, idx=i, err=e:
                           self.log_msg(f"[{idx}/{total}] {c} 失败：{err}"))

            # 随机等待 10~30 秒，最后一个器件下载完不等待
            if i < total and not self.batch_cancel:
                delay = random.uniform(10, 30)
                self.after(0, lambda d=delay: self.log_msg(f"    等待 {d:.1f} 秒后继续…"))
                # 分段 sleep（每秒检查一次），以便及时响应取消操作
                elapsed = 0.0
                while elapsed < delay and not self.batch_cancel:
                    time.sleep(min(1.0, delay - elapsed))
                    elapsed += 1.0

        # 正常完成（非取消）时显示汇总
        if not self.batch_cancel:
            self.after(0, lambda s=success, f=fail:
                       self.log_msg(f"═══ 批量下载完成：{s} 成功 / {f} 失败 ═══"))

        # 恢复所有按钮状态
        self.batch_running = False
        self.batch_cancel = False
        self.after(0, lambda: self.btn_batch.config(state="normal", bg="#007bff"))
        self.after(0, lambda: self.btn_download.config(state="normal", bg="#28a745", text="立即下载"))
        self.after(0, lambda: self.btn_cancel.config(state="disabled"))


if __name__ == "__main__":
    JLC3DApp().mainloop()
