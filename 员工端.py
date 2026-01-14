import time, ctypes, threading, requests, socket, sys, os  # pyright: ignore[reportMissingModuleSource]
from ctypes import wintypes
from datetime import datetime, timedelta, timezone
import winreg

# PyQt5导入
from PyQt5.QtWidgets import (  # pyright: ignore[reportMissingImports]
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QLabel, QPushButton, QSystemTrayIcon, QMenu, QAction, QMessageBox,
    QDialog, QTextEdit  # 消息弹窗所需的类
)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QEvent  # pyright: ignore[reportMissingImports]
from PyQt5.QtGui import QIcon, QFont, QPixmap, QPainter, QColor  # pyright: ignore[reportMissingImports]

SCAN_INTERVAL = 0.1     # 扫描间隔：0.1秒
REPORT_INTERVAL = 0.5   # 上报间隔：0.5秒
VALID_REPLY_THRESHOLD = 0.5  # ≥0.5秒才算有效回复
SERVER_URL = "http://101.42.32.73:9999"
EMPLOYEE_NAME = socket.gethostname()

# 修复5：全局变量，用于判断是否成功连接到服务器
last_report_success = False

# 北京时区（UTC+8）
BEIJING_TZ = timezone(timedelta(hours=8))

def get_beijing_now():
    """获取北京时间的当前时间"""
    return datetime.now(BEIJING_TZ)

def get_beijing_date_str():
    """获取北京时间的日期字符串（格式：YYYY-MM-DD）"""
    return get_beijing_now().strftime("%Y-%m-%d")

def create_cow_icon():
    """Q字图标"""
    p = QPixmap(32, 32)
    p.fill(Qt.transparent)
    painter = QPainter(p)
    painter.setFont(QFont("Arial", 24, QFont.Bold))
    painter.setPen(QColor(0, 120, 215))
    painter.drawText(0, 0, 32, 32, Qt.AlignCenter, "Q")
    painter.end()
    return QIcon(p)

user32 = ctypes.windll.user32
user32.EnumWindows.argtypes = [ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM), wintypes.LPARAM]
user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetClassNameW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
user32.GetWindowRect.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.RECT)]
user32.IsWindowVisible.argtypes = [wintypes.HWND]

reception_windows = {}
popup_info = {}
daily_stats = {"last_reset": get_beijing_date_str(), "today_consult": 0, "today_replied": 0, "today_reply_time": 0.0}
next_unknown_id = 1

def load_stats_from_server():
    """从服务器加载当天的统计数据"""
    global daily_stats
    today = get_beijing_date_str()
    
    try:
        resp = requests.get(f"{SERVER_URL}/get_stats", params={"employee_name": EMPLOYEE_NAME}, timeout=3)
        if resp.status_code == 200:
            data = resp.json()
            
            # 验证服务器返回的数据日期
            data_date = data.get("data_date", "")
            
            if data_date != today:
                # 数据日期和今天不一致，可能是旧数据或跨天
                print(f"[启动] 检测到日期不一致：服务器数据日期={data_date}, 今天={today}")
                print(f"[启动] 清空数据，从0开始")
                daily_stats.update({
                    "last_reset": today,
                    "today_consult": 0,
                    "today_replied": 0,
                    "today_reply_time": 0.0
                })
            else:
                # 日期一致，正常加载数据（当天重启场景）
                daily_stats.update({
                    "last_reset": today,
                    "today_consult": data.get("today_consult", 0),
                    "today_replied": data.get("replied_count", 0),
                    "today_reply_time": data.get("total_reply_time", 0.0)
                })
                print(f"[启动] 已从服务器加载统计数据: 咨询={daily_stats['today_consult']}, 回复={daily_stats['today_replied']}")
    except Exception as e:
        print(f"[启动] 加载统计数据失败: {e}，从0开始")
        # 如果加载失败，初始化为今天的日期，从0开始
        daily_stats.update({
            "last_reset": today,
            "today_consult": 0,
            "today_replied": 0,
            "today_reply_time": 0.0
        })

# sync_stats_to_server() 函数已删除
# 原因：/report 接口已经每0.5秒实时更新数据库，不需要额外的同步
# 数据通过 /report 接口统一管理，避免重复上报和数据不一致

def reset_daily():
    """检查是否需要重置（每天0:00，使用北京时间）"""
    global next_unknown_id, daily_stats
    today = get_beijing_date_str()
    if daily_stats["last_reset"] != today:
        print(f"[RESET] 检测到新的一天（北京时间）: {daily_stats['last_reset']} -> {today}")
        
        # 🔧 修复：跨天前先上报最后的数据（防止丢失最后0.5秒内的增量）
        if daily_stats["today_consult"] > 0 or daily_stats["today_replied"] > 0:
            print(f"[RESET] 跨天前最后上报昨天数据：咨询={daily_stats['today_consult']}, 回复={daily_stats['today_replied']}")
            try:
                report_to_server()  # 立即上报昨天的最后数据
                print(f"[RESET] 昨天最后数据上报成功")
            except Exception as e:
                print(f"[RESET] 昨天最后数据上报失败（但不影响清零）: {e}")
        
        # 重置本地统计数据为0（新的一天从0开始）
        # 注意：昨天的数据已经通过上面的最后上报保存到数据库了
        daily_stats.update({"last_reset": today, "today_consult": 0, "today_replied": 0, "today_reply_time": 0.0})
        next_unknown_id = 1
        
        print(f"[RESET] 新的一天，已重置本地统计数据为0，从头开始计算")

def get_window_text(hwnd):
    length = user32.GetWindowTextLengthW(hwnd)
    if length == 0: return ""
    title = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, title, length + 1)
    return title.value

def get_class_name(hwnd):
    class_name = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, class_name, 256)
    return class_name.value

def get_window_rect(hwnd):
    rect = wintypes.RECT()
    if user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        return (rect.right - rect.left, rect.bottom - rect.top)
    return None

def get_customer_count_from_height(h):
    if h < 120: return 0
    count = (h // 60) - 1
    return max(0, min(8, count))

def get_virtual_shop_name(hwnd):
    global next_unknown_id
    if hwnd not in popup_info:
        return f"未知店铺{next_unknown_id}"
    if "virtual_id" not in popup_info[hwnd]:
        popup_info[hwnd]["virtual_id"] = next_unknown_id
        next_unknown_id += 1
    vid = popup_info[hwnd]["virtual_id"]
    return f"未知店铺{vid}"

def match_windows():
    for info in popup_info.values():
        if not info.get("permanently_bound", False):
            info["owner_shop"] = None
            info["matched"] = False
    if not reception_windows or not popup_info:
        return
    if len(reception_windows) == 1 and len(popup_info) == 1:
        p_hwnd = next(iter(popup_info))
        r_info = next(iter(reception_windows.values()))
        popup_info[p_hwnd]["owner_shop"] = r_info["shop"]
        popup_info[p_hwnd]["matched"] = True
        return
    used_receptions = set()
    for p_hwnd, p_info in popup_info.items():
        if p_info.get("permanently_bound", False):
            continue
        best_shop = None
        best_r_hwnd = None
        min_diff = float('inf')
        for r_hwnd, r_info in reception_windows.items():
            if r_hwnd in used_receptions:
                continue
            diff = abs(p_info["create_time"] - r_info["first_seen"])
            if diff < 0.3 and diff < min_diff:
                min_diff = diff
                best_shop = r_info["shop"]
                best_r_hwnd = r_hwnd
        if best_shop:
            p_info["owner_shop"] = best_shop
            p_info["matched"] = True
            used_receptions.add(best_r_hwnd)

def handle_customer_close(customer, now):
    duration = now - customer["enter_time"]
    # 仅当停留时间 >= 0.5 秒才计入有效回复
    if duration >= VALID_REPLY_THRESHOLD:
        daily_stats["today_replied"] += 1
        daily_stats["today_reply_time"] += duration

def scan_and_update():
    """高频扫描（0.1秒），更新内部状态"""
    global reception_windows, popup_info
    now = time.time()
    reset_daily()
    current_receptions = {}
    current_popups = {}

    def enum_cb(hwnd, _):
        if not user32.IsWindowVisible(hwnd):
            return True
        cls = get_class_name(hwnd)
        title = get_window_text(hwnd)
        size = get_window_rect(hwnd)
        if not size:
            return True
        w, h = size
        if cls == "Qt5152QWindowIcon" and "-接待中心" in title:
            shop = title.split("-接待中心")[0].strip()
            if shop:
                if hwnd not in reception_windows:
                    reception_windows[hwnd] = {"shop": shop, "first_seen": now}
                current_receptions[hwnd] = reception_windows[hwnd]
        elif cls == "Qt5152QWindowToolSaveBits" and title == "消息提醒" and 380 <= w <= 420 and 120 <= h <= 540:
            current_popups[hwnd] = h
        return True

    enum_proc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(enum_cb)
    user32.EnumWindows(enum_proc, 0)

    # 更新接待窗口
    for hwnd in list(reception_windows.keys()):
        if hwnd not in current_receptions:
            del reception_windows[hwnd]

    # 处理弹窗变化
    for hwnd, h in current_popups.items():
        new_count = get_customer_count_from_height(h)
        if hwnd not in popup_info:
            popup_info[hwnd] = {
                "create_time": now,
                "customers": [],
                "owner_shop": None,
                "matched": False,
                "permanently_bound": False
            }
        old_count = len(popup_info[hwnd]["customers"])
        # 新增客户
        if new_count > old_count:
            for _ in range(new_count - old_count):
                popup_info[hwnd]["customers"].append({"enter_time": now})
                daily_stats["today_consult"] += 1
        # 客户减少 - 从等待时间最长的开始移除
        elif new_count < old_count:
            # 按等待时间降序排序（最久的在前），使用sorted创建新列表
            customers = popup_info[hwnd]["customers"]
            customers_sorted = sorted(customers, key=lambda c: now - c["enter_time"], reverse=True)
            # 移除等待时间最长的顾客
            removed = customers_sorted[:old_count - new_count]
            popup_info[hwnd]["customers"] = customers_sorted[old_count - new_count:]
            for cust in removed:
                cust["owner_shop"] = popup_info[hwnd].get("owner_shop") or get_virtual_shop_name(hwnd)
                handle_customer_close(cust, now)

    # 弹窗消失
    disappeared = [h for h in popup_info if h not in current_popups]
    for hwnd in disappeared:
        # 先获取虚拟店铺名称（在pop之前）
        virtual_shop = get_virtual_shop_name(hwnd)
        info = popup_info.pop(hwnd, {})
        for cust in info.get("customers", []):
            cust["owner_shop"] = info.get("owner_shop") or virtual_shop
            handle_customer_close(cust, now)

    match_windows()

    # 单空闲窗口兜底绑定
    bound_shop_names = set()
    for info in popup_info.values():
        if (info.get("permanently_bound") or info["matched"]) and info.get("owner_shop"):
            bound_shop_names.add(info["owner_shop"])
    idle_shops = [info["shop"] for info in current_receptions.values() if info["shop"] not in bound_shop_names]
    if len(idle_shops) == 1:
        target_shop = idle_shops[0]
        for p_hwnd, info in popup_info.items():
            if not info["matched"] and not info.get("permanently_bound"):
                info["owner_shop"] = target_shop
                info["matched"] = True
                info["permanently_bound"] = True
                break

def build_display_lines():
    now = time.time()
    display_lines = []
    bound_shops = {}
    for info in popup_info.values():
        if info["matched"] or info.get("permanently_bound"):
            shop = info["owner_shop"]
            if shop not in bound_shops:
                bound_shops[shop] = []
            bound_shops[shop].extend(info["customers"])
    for shop in sorted(bound_shops.keys()):
        waits = [int(now - c["enter_time"]) for c in bound_shops[shop]]
        waits.sort(reverse=True)  # 降序排序，等待时间最久的排前面
        lines = [f"{shop}-{waits[0]}秒"] if waits else [shop]
        for w in waits[1:8]:
            lines.append(" " * len(shop) + f"-{w}秒")
        display_lines.append("\n".join(lines))
    # 未绑定弹窗
    unmatched = [(hwnd, info) for hwnd, info in popup_info.items() if not info["matched"] and not info.get("permanently_bound")]
    for hwnd, info in unmatched:
        virtual_shop = get_virtual_shop_name(hwnd)
        waits = [int(now - c["enter_time"]) for c in info["customers"]]
        waits.sort(reverse=True)  # 降序排序，等待时间最久的排前面
        if waits:
            lines = [f"{virtual_shop}-{waits[0]}秒"]
            for w in waits[1:8]:
                lines.append(" " * len(virtual_shop) + f"-{w}秒")
            display_lines.append("\n".join(lines))
    # 无弹窗但有接待窗口
    bound_names = {info["owner_shop"] for info in popup_info.values() if info.get("owner_shop")}
    for info in reception_windows.values():
        if info["shop"] not in bound_names:
            display_lines.append(info["shop"])
    return display_lines

def report_to_server():
    """低频上报（0.5秒）"""
    total_customers = sum(len(info["customers"]) for info in popup_info.values())
    avg_reply = 0
    if daily_stats["today_replied"] > 0:
        avg_reply = round(daily_stats["today_reply_time"] / daily_stats["today_replied"])
    
    # 添加日期标记和时间戳，让服务器知道这是哪一天的数据（使用北京时间）
    today = get_beijing_date_str()
    
    report = {
        "employee_name": EMPLOYEE_NAME,
        "report_date": today,  # 数据日期标记
        "report_timestamp": time.time(),  # 数据上报时间戳
        "total_customers": total_customers,
        "total_shops": len(reception_windows),
        "shops_list": build_display_lines(),
        "today_consult": daily_stats["today_consult"],
        "today_replied": daily_stats["today_replied"],  # 今日回复数
        "total_reply_time": daily_stats["today_reply_time"],  # 总回复时长
        "avg_reply": avg_reply,  # 整数！
        "online": True
    }
    global last_report_success
    try:
        resp = requests.post(f"{SERVER_URL}/report", json=report, timeout=2)
        if resp.status_code == 200:
            # 修复5：成功上报后设置连接状态
            last_report_success = True
            print(f"[员工端] 上报成功: {EMPLOYEE_NAME}")
        else:
            print(f"[员工端] 上报失败，状态码: {resp.status_code}")
    except Exception as e:
        print(f"[员工端] 上报异常: {e}")

def check_startup_status():
    """检查是否已设置开机自启"""
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ)
        try:
            value, _ = winreg.QueryValueEx(key, "QN_Monitor")
            winreg.CloseKey(key)
            return True
        except FileNotFoundError:
            winreg.CloseKey(key)
            return False
    except:
        return False

def set_startup(enable):
    """设置或取消开机自启"""
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    app_path = os.path.abspath(sys.argv[0])
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE)
        if enable:
            winreg.SetValueEx(key, "QN_Monitor", 0, winreg.REG_SZ, f'"{app_path}"')
        else:
            try:
                winreg.DeleteValue(key, "QN_Monitor")
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"设置开机自启失败: {e}")
        return False

# 全局变量，用于GUI更新
gui_instance = None

class EmployeeMonitorGUI(QMainWindow):
    """员工端GUI主窗口"""
    
    # 修复6：定义信号，用于在其他线程中安全地显示消息弹窗
    show_message_signal = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        global gui_instance
        gui_instance = self
        
        # 极致压缩：窗口标题只显示"千牛识别"
        self.setWindowTitle("千牛识别")
        
        # 极致压缩：最小化窗口尺寸
        self.setFixedSize(280, 180)
        
        # 居中显示
        screen = QApplication.desktop().screenGeometry()
        self.move((screen.width() - 280) // 2, (screen.height() - 180) // 2)
        
        # 窗口置顶，保留最小化和关闭按钮
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.Window)
        
        # 极致压缩：主widget，零边距，零间距
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(8, 5, 8, 5)  # 极小边距
        main_layout.setSpacing(3)  # 极小间距
        central_widget.setLayout(main_layout)
        
        # 极致压缩：连接状态和时间合并为一行
        self.status_label = QLabel("连接中...")
        self.status_label.setFont(QFont("微软雅黑", 9))
        main_layout.addWidget(self.status_label)
        
        # 极致压缩：昨日数据
        self.yesterday_label = QLabel("昨日平均回复时长: -- 秒")
        self.yesterday_label.setFont(QFont("微软雅黑", 9))
        main_layout.addWidget(self.yesterday_label)
        
        # 极致压缩：今日数据
        self.today_label = QLabel("今日平均回复时长: -- 秒")
        self.today_label.setFont(QFont("微软雅黑", 9))
        main_layout.addWidget(self.today_label)
        
        # 极致压缩：分隔线（极细）
        line = QLabel()
        line.setFrameStyle(QLabel.HLine | QLabel.Sunken)
        line.setMaximumHeight(1)
        main_layout.addWidget(line)
        
        # 极致压缩：开机自启按钮（更小）
        self.startup_btn = QPushButton()
        self.startup_btn.setFont(QFont("微软雅黑", 9))
        self.startup_btn.setFixedHeight(28)
        self.startup_btn.clicked.connect(self.toggle_startup)
        main_layout.addWidget(self.startup_btn)
        
        # 更新开机自启状态显示
        self.update_startup_status()
        
        # 创建系统托盘图标
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setToolTip("千牛识别")
        
        # 设置牛🐮图标
        cow_icon = create_cow_icon()
        self.tray_icon.setIcon(cow_icon)
        self.setWindowIcon(cow_icon)
        
        # 单击无反应，仅双击打开窗口
        self.tray_icon.activated.connect(self.on_tray_activated)
        
        # 显示托盘图标
        self.tray_icon.show()
        
        # 修复6：连接消息信号到槽函数，确保在主线程中显示弹窗
        self.show_message_signal.connect(self.show_message_in_main_thread)
        
        # 定时器：更新UI数据
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_ui_data)
        self.update_timer.start(1000)  # 每秒更新一次
        
        # 加载昨日数据
        self.yesterday_avg = 0.0
        self.load_yesterday_stats()
        
        # 修复5：记录启动时间，用于判断连接状态
        self.start_time = time.time()
    
    def update_startup_status(self):
        """修复1：更新开机自启状态显示 - 按钮背景色"""
        is_enabled = check_startup_status()
        if is_enabled:
            self.startup_btn.setText("开启自启 已开启 ✓")
            self.startup_btn.setStyleSheet("""
                QPushButton {
                    background-color: #4CAF50;
                    color: white;
                    border: none;
                    border-radius: 5px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #45a049;
                }
            """)
        else:
            self.startup_btn.setText("开启自启 未开启 ✗")
            self.startup_btn.setStyleSheet("""
                QPushButton {
                    background-color: #f44336;
                    color: white;
                    border: none;
                    border-radius: 5px;
                    font-weight: bold;
                }
                QPushButton:hover {
                    background-color: #da190b;
                }
            """)
    
    def toggle_startup(self):
        """切换开机自启状态"""
        is_enabled = check_startup_status()
        if set_startup(not is_enabled):
            self.update_startup_status()
            status_text = "已开启" if not is_enabled else "已关闭"
            QMessageBox.information(self, "开机自启", f"开机自启{status_text}")
        else:
            QMessageBox.critical(self, "错误", "设置失败，请以管理员权限运行")
    
    def load_yesterday_stats(self):
        """从服务器加载昨日数据"""
        try:
            resp = requests.get(
                f"{SERVER_URL}/history",
                params={"employee_id": EMPLOYEE_NAME, "period": "yesterday"},
                timeout=3
            )
            if resp.status_code == 200:
                data = resp.json()
                if data and len(data) > 0:
                    self.yesterday_avg = float(data[0].get("avg_reply", 0))
                    print(f"[GUI] 加载昨日数据成功: {self.yesterday_avg}秒")
        except Exception as e:
            print(f"[GUI] 加载昨日数据失败: {e}")
    
    def update_ui_data(self):
        """极致压缩：更新UI显示数据"""
        global last_report_success
        try:
            # 极致压缩：今日平均回复时长（显示一位小数）
            if daily_stats["today_replied"] > 0:
                today_avg = daily_stats["today_reply_time"] / daily_stats["today_replied"]
                self.today_label.setText(f"今日平均回复时长: {today_avg:.1f} 秒")
            else:
                self.today_label.setText("今日平均回复时长: -- 秒")
            
            # 极致压缩：昨日平均回复时长（显示一位小数）
            if self.yesterday_avg > 0:
                self.yesterday_label.setText(f"昨日平均回复时长: {self.yesterday_avg:.1f} 秒")
            else:
                self.yesterday_label.setText("昨日平均回复时长: -- 秒")
            
            # 极致压缩：连接状态和时间合并显示
            now = datetime.now()
            if last_report_success:
                self.status_label.setText(f"✓ 已连接服务器: {now.strftime('%H:%M:%S')}")
                self.status_label.setStyleSheet("color: green; font-weight: bold;")
            else:
                self.status_label.setText(f"连接中... {now.strftime('%H:%M:%S')}")
                self.status_label.setStyleSheet("color: orange;")
        except Exception as e:
            print(f"[GUI] 更新UI错误: {e}")
    
    def show_window(self):
        """显示窗口"""
        # 修复2：确保窗口正确显示并置顶
        self.setWindowState(self.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
        self.show()
        self.raise_()
        self.activateWindow()
        print("[GUI] 窗口已显示")
    
    def on_tray_activated(self, reason):
        """托盘图标激活事件：单击无反应，双击打开窗口"""
        # 仅处理双击事件
        if reason == QSystemTrayIcon.DoubleClick:
            self.show_window()
        # 单击（Trigger）不做任何处理
    
    def show_window(self):
        """显示窗口"""
        self.setWindowState(self.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
        self.show()
        self.raise_()
        self.activateWindow()
        print("[GUI] 窗口已显示")
    
    def show_message_in_main_thread(self, message):
        """修复6：在主线程中显示消息弹窗（避免崩溃）"""
        print(f"[弹窗] 准备显示消息: {message[:50]}...")
        try:
            dialog = QDialog(self)
            dialog.setWindowTitle("来自管理端的消息")
            dialog.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.Dialog)
            dialog.resize(500, 300)
            
            # 居中显示
            screen = QApplication.desktop().screenGeometry()
            dialog.move((screen.width() - 500) // 2, (screen.height() - 300) // 2)
            
            layout = QVBoxLayout()
            
            # 文本框（可复制）
            text_widget = QTextEdit()
            text_widget.setPlainText(message)
            text_widget.setReadOnly(False)  # 允许复制
            text_widget.setFont(QFont("微软雅黑", 12))
            layout.addWidget(text_widget)
            
            # 确认按钮
            confirm_btn = QPushButton("确认")
            confirm_btn.setFont(QFont("微软雅黑", 12))
            confirm_btn.clicked.connect(dialog.accept)
            layout.addWidget(confirm_btn)
            
            dialog.setLayout(layout)
            print(f"[弹窗] 显示对话框")
            dialog.exec_()
            print(f"[弹窗] 对话框已关闭")
        except Exception as e:
            print(f"[弹窗] 显示错误: {e}")
            import traceback
            traceback.print_exc()
    
    def changeEvent(self, event):
        """最小化时隐藏到系统托盘"""
        if event.type() == QEvent.WindowStateChange:
            if self.isMinimized():
                QTimer.singleShot(0, self.hide)
                # 显示最小化提示
                QTimer.singleShot(100, self.show_minimize_message)
        super().changeEvent(event)
    
    def show_minimize_message(self):
        """显示最小化提示弹窗"""
        msg_box = QMessageBox()
        msg_box.setWindowTitle("千牛识别")
        msg_box.setText("已最小化到系统托盘，双击即可运行")
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.setWindowFlags(Qt.WindowStaysOnTopHint)
        msg_box.exec_()
    
    def closeEvent(self, event):
        """点击X关闭按钮时确认退出"""
        reply = QMessageBox.question(
            self, 
            "确认退出",
            "确定要退出千牛识别程序吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.tray_icon.hide()
            QApplication.quit()
            os._exit(0)
        else:
            event.ignore()
    

def poll_messages_loop():
    """长轮询线程，定期查询消息"""
    global gui_instance
    print(f"[消息轮询] 已启动，员工ID: {EMPLOYEE_NAME}")
    
    # 等待GUI实例初始化
    while gui_instance is None:
        time.sleep(0.5)
    
    while True:
        try:
            resp = requests.get(
                f"{SERVER_URL}/poll_messages/{EMPLOYEE_NAME}",
                timeout=35  # 比服务器超时时间稍长
            )
            if resp.status_code == 200:
                data = resp.json()
                messages = data.get("messages", [])
                if messages:
                    print(f"[消息轮询] 收到 {len(messages)} 条消息")
                    for msg_data in messages:
                        message = msg_data.get("message", "")
                        if message:
                            print(f"[消息] {message[:50]}...")
                            # 修复6：使用信号在主线程中显示弹窗，避免崩溃
                            if gui_instance:
                                gui_instance.show_message_signal.emit(message)
        except requests.exceptions.Timeout:
            # 超时是正常的，继续轮询
            pass
        except Exception as e:
            print(f"[消息轮询] 错误: {e}")
            time.sleep(5)  # 出错后等待5秒再重试

def main_loop():
    last_report = 0
    while True:
        scan_and_update()  # 0.1秒高频扫描
        now = time.time()
        if now - last_report >= REPORT_INTERVAL:
            report_to_server()
            last_report = now
        time.sleep(SCAN_INTERVAL)

if __name__ == "__main__":
    # 单实例检查（使用命名互斥量）
    mutex = ctypes.windll.kernel32.CreateMutexW(None, True, "Global\\QN_Monitor_Mutex_V2")
    if ctypes.windll.kernel32.GetLastError() == 183:
        # 已有实例运行，尝试通过文件通信让现有实例显示窗口
        signal_file = os.path.join(os.environ.get('TEMP', '.'), 'qn_monitor_show_window.signal')
        try:
            with open(signal_file, 'w') as f:
                f.write('show')
        except:
            pass
        print("程序已在运行，正在激活主窗口...")
        time.sleep(1)
        sys.exit(0)
    
    # 隐藏控制台窗口
    if sys.platform == "win32":
        try:
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd:
                ctypes.windll.user32.ShowWindow(hwnd, 0)  # SW_HIDE
        except:
            pass
    
    # 启动时从服务器加载统计数据
    load_stats_from_server()
    
    # 启动后台监控线程
    threading.Thread(target=main_loop, daemon=True).start()
    
    # 启动消息长轮询线程
    threading.Thread(target=poll_messages_loop, daemon=True).start()
    
    # 程序退出时最后上报一次数据
    import atexit
    atexit.register(report_to_server)
    
    # 启动PyQt5应用
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)  # 关闭窗口不退出程序
    
    # 创建主窗口
    window = EmployeeMonitorGUI()
    window.show()
    
    # 检查显示窗口信号文件的定时器
    def check_show_signal():
        signal_file = os.path.join(os.environ.get('TEMP', '.'), 'qn_monitor_show_window.signal')
        if os.path.exists(signal_file):
            try:
                os.remove(signal_file)
                if window:
                    window.show_window()
            except:
                pass
    
    signal_timer = QTimer()
    signal_timer.timeout.connect(check_show_signal)
    signal_timer.start(500)  # 每0.5秒检查一次
    
    sys.exit(app.exec_())
