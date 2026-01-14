import sys
import requests  # pyright: ignore[reportMissingModuleSource]
import time
import json
import os
from datetime import datetime, timedelta

# 重要：必须在导入QApplication之前导入QWebEngineWidgets
from PyQt5.QtWebEngineWidgets import QWebEngineView  # pyright: ignore[reportMissingImports]

from PyQt5.QtWidgets import (  # pyright: ignore[reportMissingImports]
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, 
    QPushButton, QHBoxLayout, QVBoxLayout, QWidget, QLabel,
    QHeaderView, QMessageBox, QDialog, QLineEdit, QRadioButton,
    QButtonGroup, QDateEdit, QGridLayout, QAbstractItemView,
    QInputDialog, QComboBox, QGroupBox, QColorDialog, QScrollArea, QSpinBox,
    QStackedWidget, QTextEdit, QDialogButtonBox, QListWidget, QListWidgetItem
)
from PyQt5.QtCore import Qt, QTimer, QDate, QUrl, pyqtSlot, QObject, QThread, pyqtSignal, QSettings  # pyright: ignore[reportMissingImports]
from PyQt5.QtGui import QColor, QFont  # pyright: ignore[reportMissingImports]

# 尝试导入 QWebChannel，如果失败则禁用双击功能
try:
    from PyQt5.QtWebChannel import QWebChannel  # pyright: ignore[reportMissingImports]
    WEBCHANNEL_AVAILABLE = True
except ImportError:
    WEBCHANNEL_AVAILABLE = False
    print("警告: QWebChannel 不可用，双击查看员工详情功能将被禁用")

SERVER_URL = "http://101.42.32.73:9999"

# 网络请求超时时间（秒）
REQUEST_TIMEOUT = 3  # 修复：减少超时时间，提高响应速度

def handle_request_error(parent, error, operation="操作"):
    """统一处理网络请求错误"""
    if isinstance(error, requests.exceptions.Timeout):
        QMessageBox.critical(parent, "错误", f"{operation}超时，请检查网络连接")
    elif isinstance(error, requests.exceptions.ConnectionError):
        QMessageBox.critical(parent, "错误", "无法连接到服务器，请检查服务器地址和网络连接")
    else:
        QMessageBox.critical(parent, "错误", f"{operation}失败：{str(error)}")

class DeleteEmployeeDialog(QDialog):
    """删除员工对话框"""
    def __init__(self, parent=None, employees_list=None):
        super().__init__(parent)
        self.setWindowTitle("删除员工记录")
        self.resize(500, 400)
        self.employees_list = employees_list or []
        self.delete_all = True
        
        layout = QVBoxLayout()
        
        info_label = QLabel("选择要删除的员工和删除方式：")
        info_label.setFont(QFont("Arial", 10))
        layout.addWidget(info_label)
        
        employee_group = QGroupBox("选择员工")
        employee_layout = QVBoxLayout()
        self.employee_combo = QComboBox()
        if self.employees_list:
            for emp in self.employees_list:
                display_name = emp.get("display_name", emp.get("employee_name", ""))
                employee_id = emp.get("employee_name", "")
                self.employee_combo.addItem(f"{display_name} ({employee_id})", employee_id)
        else:
            self.employee_combo.addItem("暂无员工", "")
        employee_layout.addWidget(self.employee_combo)
        employee_group.setLayout(employee_layout)
        layout.addWidget(employee_group)
        
        delete_type_group = QGroupBox("删除方式")
        delete_type_layout = QVBoxLayout()
        self.delete_type_group = QButtonGroup(self)
        
        self.delete_today_radio = QRadioButton("只删除今日记录（保留历史数据）")
        self.delete_today_radio.setChecked(False)
        self.delete_all_radio = QRadioButton("完全删除（删除所有记录）")
        self.delete_all_radio.setChecked(True)
        
        self.delete_type_group.addButton(self.delete_today_radio, 0)
        self.delete_type_group.addButton(self.delete_all_radio, 1)
        
        delete_type_layout.addWidget(self.delete_today_radio)
        delete_type_layout.addWidget(self.delete_all_radio)
        delete_type_group.setLayout(delete_type_layout)
        layout.addWidget(delete_type_group)
        
        warning_label = QLabel("⚠️ 警告：删除操作不可逆，请确认后再操作！")
        warning_label.setStyleSheet("color: red; font-weight: bold;")
        layout.addWidget(warning_label)
        
        button_layout = QHBoxLayout()
        self.delete_btn = QPushButton("确认删除")
        self.delete_btn.setStyleSheet("background-color: #ff4444; color: white; font-weight: bold;")
        self.cancel_btn = QPushButton("取消")
        button_layout.addStretch()
        button_layout.addWidget(self.delete_btn)
        button_layout.addWidget(self.cancel_btn)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        self.delete_btn.clicked.connect(self.confirm_delete)
        self.cancel_btn.clicked.connect(self.reject)
    
    def confirm_delete(self):
        """确认删除"""
        employee_id = self.employee_combo.currentData()
        if not employee_id:
            QMessageBox.warning(self, "警告", "请选择要删除的员工")
            return
        
        delete_all = self.delete_all_radio.isChecked()
        delete_type_text = "完全删除（删除所有记录）" if delete_all else "只删除今日记录"
        
        # 二次确认
        reply = QMessageBox.question(
            self, 
            "确认删除", 
            f"确定要{delete_type_text}员工 '{self.employee_combo.currentText()}' 吗？\n\n此操作不可逆！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            # 执行删除
            self.delete_employee(employee_id, delete_all)
    
    def delete_employee(self, employee_id, delete_all):
        """执行删除操作"""
        try:
            # 通过API删除员工
            resp = requests.post(
                f"{SERVER_URL}/delete_employee",
                json={
                    "employee_id": employee_id,
                    "delete_all": delete_all
                },
                timeout=REQUEST_TIMEOUT
            )
            
            if resp.status_code == 200:
                result = resp.json()
                if result.get("success"):
                    QMessageBox.information(self, "删除成功", f"员工 '{employee_id}' 的记录已删除\n\n管理端会在下次刷新时自动更新")
                    self.accept()
                else:
                    QMessageBox.warning(self, "删除失败", result.get("message", "删除操作失败"))
            else:
                error_msg = resp.json().get("detail", "删除操作失败") if resp.text else f"服务器返回错误：{resp.status_code}"
                QMessageBox.warning(self, "删除失败", error_msg)
        except Exception as e:
            handle_request_error(self, e, "删除员工")

class EmployeeManagementDialog(QDialog):
    """员工管理对话框（颜色+排序+隐藏）"""
    def __init__(self, parent=None, employees_list=None):
        super().__init__(parent)
        self.setWindowTitle("员工管理")
        self.resize(800, 600)
        self.employees_list = employees_list or []
        self.color_configs = {}  # {employee_id: {'bar': '#xxx'}}
        self.sort_orders = {}  # {employee_id: order}
        self.visibility_configs = {}  # {employee_id: {'hidden': bool, 'is_manual': bool}}
        
        layout = QVBoxLayout()
        
        info_label = QLabel("为每个员工设置柱状图颜色，并设置全局线形图颜色：")
        info_label.setFont(QFont("Arial", 10))
        layout.addWidget(info_label)
        
        # 全局线形图颜色设置
        global_group = QGroupBox("全局设置（所有员工）")
        global_layout = QHBoxLayout()
        global_label = QLabel("线形图统一颜色:")
        self.global_line_color_btn = QPushButton()
        global_line_color = self.color_configs.get('__global__', {}).get('line', '#FFA500')
        self.global_line_color_btn.setStyleSheet(f"background-color: {global_line_color}; min-width: 100px; min-height: 30px;")
        self.global_line_color_btn.clicked.connect(lambda: self.choose_color('__global__', 'line', self.global_line_color_btn))
        global_layout.addWidget(global_label)
        global_layout.addWidget(self.global_line_color_btn)
        global_layout.addStretch()
        global_group.setLayout(global_layout)
        layout.addWidget(global_group)
        
        # 分隔线
        separator = QLabel("—" * 60)
        separator.setAlignment(Qt.AlignCenter)
        layout.addWidget(separator)
        
        # 员工颜色设置区域（使用滚动区域）
        scroll = QScrollArea()
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        # 加载现有颜色配置
        self.load_color_configs()
        
        for emp in self.employees_list:
            employee_id = emp.get("employee_name", "")
            display_name = emp.get("display_name", employee_id)
            
            # 每个员工的设置区域
            emp_group = QGroupBox(f"{display_name} ({employee_id})")
            emp_layout = QHBoxLayout()
            
            # 1. 柱状图颜色
            bar_label = QLabel("柱状图颜色:")
            bar_color_btn = QPushButton()
            bar_color = self.color_configs.get(employee_id, {}).get('bar', '#4CAF50')
            bar_color_btn.setStyleSheet(f"background-color: {bar_color}; min-width: 80px; min-height: 30px;")
            bar_color_btn.clicked.connect(lambda checked, eid=employee_id, btn=bar_color_btn: self.choose_color(eid, 'bar', btn))
            emp_layout.addWidget(bar_label)
            emp_layout.addWidget(bar_color_btn)
            
            # 2. 排序
            sort_label = QLabel("  排序:")
            sort_spinbox = QSpinBox()
            sort_spinbox.setRange(1, 9999)
            sort_spinbox.setValue(self.sort_orders.get(employee_id, 9999))
            sort_spinbox.setMinimumWidth(80)
            emp_layout.addWidget(sort_label)
            emp_layout.addWidget(sort_spinbox)
            
            # 3. 当日数据显示/隐藏
            visibility_label = QLabel("  当日数据:")
            visibility_btn = QPushButton()
            vis_config = self.visibility_configs.get(employee_id, {'hidden': False, 'is_manual': True})
            is_hidden = vis_config.get('hidden', False)
            is_manual = vis_config.get('is_manual', True)
            
            # 设置按钮文本和颜色
            if is_hidden:
                if is_manual:
                    visibility_btn.setText("隐藏数据")
                else:
                    visibility_btn.setText("隐藏数据(自动)")
                visibility_btn.setStyleSheet("background-color: #F44336; color: white; font-weight: bold; min-width: 120px; min-height: 30px;")
            else:
                visibility_btn.setText("展示数据")
                visibility_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; min-width: 120px; min-height: 30px;")
            
            # 点击切换状态
            visibility_btn.clicked.connect(lambda checked, eid=employee_id, btn=visibility_btn: self.toggle_visibility(eid, btn))
            emp_layout.addWidget(visibility_label)
            emp_layout.addWidget(visibility_btn)
            
            emp_layout.addStretch()
            emp_group.setLayout(emp_layout)
            scroll_layout.addWidget(emp_group)
            
            # 保存控件引用
            if employee_id not in self.color_configs:
                self.color_configs[employee_id] = {}
            self.color_configs[employee_id]['bar_btn'] = bar_color_btn
            self.color_configs[employee_id]['sort_spinbox'] = sort_spinbox
            self.color_configs[employee_id]['visibility_btn'] = visibility_btn
        
        scroll_layout.addStretch()
        scroll.setWidget(scroll_widget)
        scroll.setWidgetResizable(True)
        layout.addWidget(scroll)
        
        # 按钮
        button_layout = QHBoxLayout()
        save_btn = QPushButton("保存")
        save_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        cancel_btn = QPushButton("取消")
        button_layout.addStretch()
        button_layout.addWidget(save_btn)
        button_layout.addWidget(cancel_btn)
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
        save_btn.clicked.connect(self.save_all_configs)
        cancel_btn.clicked.connect(self.reject)
    
    def load_color_configs(self):
        """加载现有配置（颜色+排序+隐藏）"""
        # 1. 加载颜色配置
        try:
            employee_ids = [emp.get("employee_name", "") for emp in self.employees_list]
            employee_ids_str = ",".join(employee_ids)
            
            resp = requests.get(
                f"{SERVER_URL}/color_configs",
                params={"employee_ids": employee_ids_str},
                timeout=REQUEST_TIMEOUT
            )
            
            if resp.status_code == 200:
                data = resp.json()
                # 设置全局线形图颜色
                global_line_color = data.get("global_line_color", "#FFA500")
                self.color_configs['__global__'] = {
                    'line': global_line_color
                }
                # 更新全局线形图颜色按钮
                if hasattr(self, 'global_line_color_btn'):
                    self.global_line_color_btn.setStyleSheet(f"background-color: {global_line_color}; min-width: 100px; min-height: 30px;")
                
                # 设置员工柱状图颜色
                employee_colors = data.get("employee_colors", {})
                for employee_id, color_data in employee_colors.items():
                    if employee_id not in self.color_configs:
                        self.color_configs[employee_id] = {}
                    self.color_configs[employee_id]['bar'] = color_data.get("bar_color", "#4CAF50")
        except Exception as e:
            print(f"加载颜色配置错误: {e}")
        
        # 2. 加载排序配置
        try:
            resp = requests.get(f"{SERVER_URL}/employee_order", timeout=REQUEST_TIMEOUT)
            if resp.status_code == 200:
                order_data = resp.json()
                for item in order_data:
                    self.sort_orders[item['employee_id']] = item['order']
        except Exception as e:
            print(f"加载排序配置错误: {e}")
        
        # 3. 加载隐藏配置
        try:
            resp = requests.get(f"{SERVER_URL}/employee_visibility", timeout=REQUEST_TIMEOUT)
            if resp.status_code == 200:
                visibility_data = resp.json()
                for item in visibility_data:
                    self.visibility_configs[item['employee_id']] = {
                        'hidden': item['hidden'],
                        'is_manual': item['is_manual']
                    }
        except Exception as e:
            print(f"加载隐藏配置错误: {e}")
    
    async def _load_colors_async(self):
        """异步加载颜色配置（通过API）"""
        try:
            # 获取所有员工ID
            employee_ids = [emp.get("employee_name", "") for emp in self.employees_list]
            employee_ids_str = ",".join(employee_ids)
            
            # 通过API获取颜色配置
            resp = requests.get(
                f"{SERVER_URL}/color_configs",
                params={"employee_ids": employee_ids_str},
                timeout=REQUEST_TIMEOUT
            )
            
            if resp.status_code == 200:
                data = resp.json()
                # 设置全局线形图颜色
                global_line_color = data.get("global_line_color", "#FFA500")
                self.color_configs['__global__'] = {
                    'line': global_line_color
                }
                # 更新全局线形图颜色按钮
                if hasattr(self, 'global_line_color_btn'):
                    self.global_line_color_btn.setStyleSheet(f"background-color: {global_line_color}; min-width: 100px; min-height: 30px;")
                
                # 设置员工柱状图颜色
                employee_colors = data.get("employee_colors", {})
                for employee_id, color_data in employee_colors.items():
                    if employee_id not in self.color_configs:
                        self.color_configs[employee_id] = {}
                    self.color_configs[employee_id]['bar'] = color_data.get("bar_color", "#4CAF50")
        except Exception as e:
            print(f"加载颜色配置错误: {e}")
    
    def choose_color(self, employee_id, color_type, button):
        """选择颜色（预设颜色）"""
        # 只保留四个颜色：绿色、蓝色、红色、橙色
        preset_colors = {
            '绿': '#4CAF50',
            '蓝': '#2196F3',
            '红': '#F44336',
            '橙': '#FFA500'
        }
        
        current_color = self.color_configs.get(employee_id, {}).get(color_type, '#4CAF50' if color_type == 'bar' else '#FFA500')
        current_name = [name for name, hex_val in preset_colors.items() if hex_val == current_color]
        current_name = current_name[0] if current_name else ('绿' if color_type == 'bar' else '橙')
        
        # 创建颜色选择对话框
        color_dialog = QDialog(self)
        color_dialog.setWindowTitle(f"选择{color_type}颜色")
        color_dialog.resize(400, 200)
        dialog_layout = QVBoxLayout()
        
        # 颜色按钮网格（只显示四个颜色）
        color_grid = QGridLayout()
        row, col = 0, 0
        for color_name, color_hex in preset_colors.items():
            color_btn = QPushButton(color_name)
            color_btn.setStyleSheet(f"background-color: {color_hex}; color: white; font-weight: bold; min-width: 80px; min-height: 50px; font-size: 14px;")
            color_btn.clicked.connect(lambda checked, name=color_name, hex_val=color_hex: self._set_color(employee_id, color_type, button, hex_val, color_dialog))
            color_grid.addWidget(color_btn, row, col)
            col += 1
            if col >= 2:
                col = 0
                row += 1
        
        dialog_layout.addLayout(color_grid)
        cancel_btn = QPushButton("取消")
        cancel_btn.clicked.connect(color_dialog.reject)
        dialog_layout.addWidget(cancel_btn)
        color_dialog.setLayout(dialog_layout)
        color_dialog.exec_()
    
    def _set_color(self, employee_id, color_type, button, color_hex, dialog):
        """设置颜色"""
        if employee_id not in self.color_configs:
            self.color_configs[employee_id] = {}
        self.color_configs[employee_id][color_type] = color_hex
        button.setStyleSheet(f"background-color: {color_hex}; min-width: 80px; min-height: 30px; border: 2px solid #333;")
        dialog.accept()
    
    def toggle_visibility(self, employee_id, button):
        """切换显示/隐藏状态"""
        vis_config = self.visibility_configs.get(employee_id, {'hidden': False, 'is_manual': True})
        is_hidden = vis_config.get('hidden', False)
        
        # 切换状态
        new_hidden = not is_hidden
        self.visibility_configs[employee_id] = {
            'hidden': new_hidden,
            'is_manual': True  # 手动设置
        }
        
        # 更新按钮
        if new_hidden:
            button.setText("隐藏数据")
            button.setStyleSheet("background-color: #F44336; color: white; font-weight: bold; min-width: 120px; min-height: 30px;")
        else:
            button.setText("展示数据")
            button.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; min-width: 120px; min-height: 30px;")
    
    def save_all_configs(self):
        """保存所有配置（颜色+排序+隐藏）"""
        try:
            # 1. 保存颜色配置
            colors = []
            for employee_id, config in self.color_configs.items():
                if employee_id == '__global__':
                    continue
                bar_color = config.get('bar', '#4CAF50')
                colors.append({
                    "employee_id": employee_id,
                    "bar_color": bar_color
                })
            
            global_line_color = self.color_configs.get('__global__', {}).get('line', '#FFA500')
            
            resp = requests.post(
                f"{SERVER_URL}/save_color_configs",
                json={
                    "colors": colors,
                    "global_line_color": global_line_color
                },
                timeout=REQUEST_TIMEOUT
            )
            
            if resp.status_code != 200:
                QMessageBox.warning(self, "警告", "保存颜色配置失败")
                return
            
            # 2. 保存排序配置
            orders = []
            for employee_id, config in self.color_configs.items():
                if employee_id == '__global__':
                    continue
                spinbox = config.get('sort_spinbox')
                if spinbox:
                    order = spinbox.value()
                    orders.append({
                        "employee_id": employee_id,
                        "order": order
                    })
            
            resp = requests.post(
                f"{SERVER_URL}/save_employee_order",
                json={"orders": orders},
                timeout=REQUEST_TIMEOUT
            )
            
            if resp.status_code != 200:
                QMessageBox.warning(self, "警告", "保存排序配置失败")
                return
            
            # 3. 保存隐藏配置
            visibility = []
            for employee_id, vis_config in self.visibility_configs.items():
                visibility.append({
                    "employee_id": employee_id,
                    "hidden": vis_config.get('hidden', False),
                    "is_manual": vis_config.get('is_manual', True)
                })
            
            resp = requests.post(
                f"{SERVER_URL}/employee_visibility",
                json={"visibility": visibility},
                timeout=REQUEST_TIMEOUT
            )
            
            if resp.status_code != 200:
                QMessageBox.warning(self, "警告", "保存隐藏配置失败")
                return
            
            QMessageBox.information(self, "成功", "所有配置已保存")
            self.accept()
        except Exception as e:
            handle_request_error(self, e, "保存配置")
    
    async def _save_colors_async(self):
        """异步保存颜色配置（通过API）"""
        try:
            # 准备颜色配置数据
            colors = []
            for employee_id, config in self.color_configs.items():
                if employee_id == '__global__':
                    continue  # 跳过全局设置
                bar_color = config.get('bar', '#4CAF50')
                colors.append({
                    "employee_id": employee_id,
                    "bar_color": bar_color
                })
            
            global_line_color = self.color_configs.get('__global__', {}).get('line', '#FFA500')
            
            # 通过API保存颜色配置
            resp = requests.post(
                f"{SERVER_URL}/save_color_configs",
                json={
                    "colors": colors,
                    "global_line_color": global_line_color
                },
                timeout=REQUEST_TIMEOUT
            )
            
            if resp.status_code == 200:
                return True
            else:
                print(f"保存颜色配置失败，状态码: {resp.status_code}")
                return False
        except Exception as e:
            print(f"保存颜色配置错误: {e}")
            return False

if WEBCHANNEL_AVAILABLE:
    class ChartBridge(QObject):
        """JavaScript 和 Python 之间的桥接类"""
        def __init__(self, dialog):
            super().__init__()
            self.dialog = dialog
        
        @pyqtSlot(str, str)
        def onEmployeeDoubleClick(self, employee_id, employee_name):
            """处理员工柱状图双击事件（总览→单员工）"""
            print(f"双击员工: {employee_name} (ID: {employee_id})")
            self.dialog.show_single_employee(employee_id, employee_name)
        
        @pyqtSlot()
        def onBackToOverview(self):
            """处理双击返回总览事件（单员工→总览）"""
            print("双击返回总览")
            self.dialog.back_to_overview()

class HistoryDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("员工效率统计图表")
        
        # 创建QSettings（保存到Windows注册表，不生成文件）
        self.settings = QSettings("QianNiuMonitor", "HistoryDialog")
        
        # 恢复窗口大小和位置
        self.restore_window_geometry()
        
        # 设置为非模态对话框，允许同时查看主窗口
        self.setWindowFlags(Qt.Window)
        # 允许最小化和调整大小
        self.setWindowModality(Qt.NonModal)
        
        # 防止窗口自动置顶
        self.setAttribute(Qt.WA_ShowWithoutActivating, False)  # 允许正常激活
        self._prevent_raise = False  # 标志：是否阻止自动置顶
        
        # 单员工视图状态
        self.single_employee_mode = False
        self.current_employee_id = None
        self.current_employee_name = None
        
        layout = QVBoxLayout()
        
        # 时间维度控制区域
        control_layout = QHBoxLayout()
        control_layout.addWidget(QLabel("时间维度:"))
        
        self.period_group = QButtonGroup(self)
        # 修复：添加"昨日"选项，排序为：今日/昨日/本周/本月/自定义
        self.day_radio = QRadioButton("今日")
        self.yesterday_radio = QRadioButton("昨日")  # 新增：昨日选项
        self.week_radio = QRadioButton("本周")
        self.month_radio = QRadioButton("本月")
        self.custom_radio = QRadioButton("自定义")
        self.day_radio.setChecked(True)
        
        # 修复：调整ID分配，添加"昨日"ID=1，后续依次递增
        self.period_group.addButton(self.day_radio, 0)       # 今日
        self.period_group.addButton(self.yesterday_radio, 1)  # 昨日
        self.period_group.addButton(self.week_radio, 2)       # 本周
        self.period_group.addButton(self.month_radio, 3)      # 本月
        self.period_group.addButton(self.custom_radio, 4)     # 自定义
        
        # 修复：去掉单选按钮的小球，改为选中后变绿色的样式
        radio_style = """
            QRadioButton {
                spacing: 5px;
                padding: 5px 10px;
                color: #333;
                background-color: transparent;
                border: 1px solid #ccc;
                border-radius: 3px;
            }
            QRadioButton::indicator {
                width: 0px;
                height: 0px;
            }
            QRadioButton:checked {
                color: white;
                background-color: #4CAF50;
                border: 1px solid #4CAF50;
            }
            QRadioButton:hover {
                border: 1px solid #4CAF50;
            }
        """
        self.day_radio.setStyleSheet(radio_style)
        self.yesterday_radio.setStyleSheet(radio_style)
        self.week_radio.setStyleSheet(radio_style)
        self.month_radio.setStyleSheet(radio_style)
        self.custom_radio.setStyleSheet(radio_style)
        
        control_layout.addWidget(self.day_radio)
        control_layout.addWidget(self.yesterday_radio)  # 添加昨日选项
        control_layout.addWidget(self.week_radio)
        control_layout.addWidget(self.month_radio)
        control_layout.addWidget(self.custom_radio)
        control_layout.addStretch()
        
        self.start_date = QDateEdit()
        self.end_date = QDateEdit()
        self.start_date.setDate(QDate.currentDate().addDays(-7))
        self.end_date.setDate(QDate.currentDate())
        self.start_date.setEnabled(False)
        self.end_date.setEnabled(False)
        self.start_date.setCalendarPopup(True)
        self.end_date.setCalendarPopup(True)
        control_layout.addWidget(QLabel("开始:"))
        control_layout.addWidget(self.start_date)
        control_layout.addWidget(QLabel("结束:"))
        control_layout.addWidget(self.end_date)
        
        query_btn = QPushButton("查询")
        query_btn.clicked.connect(self.query_stats)
        control_layout.addWidget(query_btn)
        
        # 添加员工管理按钮
        manage_btn = QPushButton("员工管理")
        manage_btn.clicked.connect(self.open_employee_management_dialog)
        control_layout.addWidget(manage_btn)
        
        # 添加展示/隐藏员工按钮
        self.toggle_hidden_btn = QPushButton("展示隐藏员工")
        self.toggle_hidden_btn.clicked.connect(self.toggle_hidden_employees)
        control_layout.addWidget(self.toggle_hidden_btn)
        
        layout.addLayout(control_layout)
        
        self.period_group.buttonClicked.connect(self.on_period_changed)
        self.custom_radio.toggled.connect(self.on_custom_toggled)
        
        # 使用QStackedWidget管理2个WebView（总览 + 单员工）
        self.stacked_widget = QStackedWidget()
        
        # 页面0：总览WebView
        self.web_view_overview = QWebEngineView()
        self.enable_hardware_acceleration(self.web_view_overview)
        self.stacked_widget.addWidget(self.web_view_overview)  # index=0
        
        # 页面1：单员工视图（含WebView+员工列表）
        single_employee_widget = QWidget()
        single_employee_layout = QHBoxLayout()
        single_employee_layout.setContentsMargins(0, 0, 0, 0)
        
        # 单员工WebView（左侧）
        self.web_view_single = QWebEngineView()
        self.enable_hardware_acceleration(self.web_view_single)
        single_employee_layout.addWidget(self.web_view_single)
        
        # 修复：添加员工列表选择器（右侧）
        self.employee_list = QListWidget()
        self.employee_list.setMaximumWidth(200)
        self.employee_list.setStyleSheet("""
            QListWidget::item {
                padding: 8px;
                border-bottom: 1px solid #ddd;
            }
            QListWidget::item:hover {
                background-color: #e8f5e9;
            }
            QListWidget::item:selected {
                background-color: #4CAF50;
                color: white;
            }
        """)
        self.employee_list.itemClicked.connect(self.on_employee_list_clicked)
        single_employee_layout.addWidget(self.employee_list)
        
        single_employee_widget.setLayout(single_employee_layout)
        self.stacked_widget.addWidget(single_employee_widget)  # index=1
        
        # 兼容性：保留 self.web_view 指向当前显示的WebView
        self.web_view = self.web_view_overview
        
        # 单员工图表初始化状态
        self.single_chart_initialized = False  # 单员工图表是否已初始化
        
        # 员工列表数据
        self.all_employees = []  # 所有员工数据
        
        layout.addWidget(self.stacked_widget)
        
        # 设置 WebChannel 以便 JavaScript 和 Python 通信（如果可用）
        if WEBCHANNEL_AVAILABLE:
            # 总览WebView的WebChannel
            self.channel_overview = QWebChannel()
            self.bridge = ChartBridge(self)
            self.channel_overview.registerObject('bridge', self.bridge)
            self.web_view_overview.page().setWebChannel(self.channel_overview)
            
            # 单员工WebView的WebChannel
            self.channel_single = QWebChannel()
            self.channel_single.registerObject('bridge', self.bridge)
            self.web_view_single.page().setWebChannel(self.channel_single)
        else:
            self.channel_overview = None
            self.channel_single = None
            self.bridge = None
        
        # 方案3：图表实例缓存复用 + 分模式复用
        self.chart_initialized = False  # 图表是否已初始化
        self.last_mode = None  # 上一次的模式：'overview' 或 'single'
        self.last_data_json = None  # 上一次的数据JSON（用于避免重复更新）
        
        # 月度数据缓存（预加载整月数据，实现快速切换）
        self.monthly_data_cache = {}  # {employee_id: [{"date": "2026-01-01", "total_consult": 10, "avg_reply": 5}, ...]}
        self.employee_name_cache = {}  # {employee_id: display_name} 员工显示名称缓存
        self.monthly_data_loaded = False  # 是否已加载月数据
        self.cached_month = None  # 缓存的月份 (year, month)
        self.cached_date = None  # 缓存加载的日期（用于检测跨天）
        
        # 颜色配置缓存
        self.color_cache = {}  # {employee_id: {"bar": "#xxx", "line": "#xxx"}}
        self.global_line_color_cache = "#FFA500"  # 全局线形图颜色缓存
        self.color_cache_loaded = False
        
        self.setLayout(layout)
        self.init_chart()
        
        # 预加载当月数据（优先级最高，提前加载）
        QTimer.singleShot(50, self.preload_monthly_data)
        
        # 预加载颜色配置
        QTimer.singleShot(100, self.preload_color_configs)
        
        # 修复2：进入图表后立即查询一次总览，加载今日数据
        QTimer.singleShot(100, self.safe_query_stats)
        
        # 初始化展示/隐藏按钮文本
        QTimer.singleShot(200, self.update_toggle_button_text)
    
    def restore_window_geometry(self):
        """恢复窗口大小和位置（从Windows注册表读取）"""
        # 恢复窗口大小
        width = self.settings.value("window/width", 1400, type=int)
        height = self.settings.value("window/height", 800, type=int)
        self.resize(width, height)
        
        # 恢复窗口位置
        x = self.settings.value("window/x", None, type=int)
        y = self.settings.value("window/y", None, type=int)
        
        if x is not None and y is not None:
            # 检查位置是否在屏幕范围内
            screen = QApplication.desktop().screenGeometry()
            if 0 <= x < screen.width() - 100 and 0 <= y < screen.height() - 100:
                self.move(x, y)
        
        print(f"[窗口配置] 恢复窗口大小: {width}x{height}, 位置: ({x}, {y})")
    
    def save_window_geometry(self):
        """保存窗口大小和位置（保存到Windows注册表）"""
        # 保存窗口大小
        self.settings.setValue("window/width", self.width())
        self.settings.setValue("window/height", self.height())
        
        # 保存窗口位置
        self.settings.setValue("window/x", self.x())
        self.settings.setValue("window/y", self.y())
        
        print(f"[窗口配置] 保存窗口大小: {self.width()}x{self.height()}, 位置: ({self.x()}, {self.y()})")
    
    def closeEvent(self, event):
        """窗口关闭时保存配置"""
        self.save_window_geometry()
        event.accept()

    def enable_hardware_acceleration(self, web_view):
        """启用WebView的硬件加速"""
        try:
            from PyQt5.QtWebEngineWidgets import QWebEngineSettings  # pyright: ignore[reportMissingImports]
            settings = web_view.settings()
            settings.setAttribute(QWebEngineSettings.Accelerated2dCanvasEnabled, True)
            settings.setAttribute(QWebEngineSettings.WebGLEnabled, True)
        except Exception as e:
            print(f"硬件加速启用失败（不影响功能）: {e}")

    def on_period_changed(self, button):
        """时间维度切换（自动查询，除了自定义模式需要手动点查询）"""
        # 修复：调整ID判断，自定义现在是4
        is_custom = (self.period_group.id(button) == 4)
        self.start_date.setEnabled(is_custom)
        self.end_date.setEnabled(is_custom)
        
        # 检查是否需要刷新月度缓存（跨月切换）
        self.check_and_refresh_monthly_cache()
        
        # 修复：恢复自动查询，切换时间维度时自动查询（除了自定义模式）
        if not is_custom:
            self.safe_query_stats()
    
    def on_custom_toggled(self, checked):
        self.start_date.setEnabled(checked)
        self.end_date.setEnabled(checked)
    
    def keyPressEvent(self, event):
        """处理键盘事件：按回车键触发查询"""
        from PyQt5.QtCore import Qt  # pyright: ignore[reportMissingImports]
        if event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
            print("[回车键] 触发查询")
            self.query_stats()
        else:
            super().keyPressEvent(event)
    
    def check_and_refresh_monthly_cache(self):
        """检查并刷新月度缓存（当切换到不同月份或跨天时）"""
        now = datetime.now()
        current_month = (now.year, now.month)
        current_date = now.date()  # 当前日期
        
        # 如果缓存的月份不是当前月份，重新加载
        if self.cached_month != current_month:
            print(f"[缓存刷新] 检测到月份变化: {self.cached_month} -> {current_month}")
            self.preload_monthly_data()
        # 修复：如果日期变化了（跨天），也刷新缓存，确保显示今天的最新数据
        elif self.cached_date is not None and self.cached_date != current_date:
            print(f"[缓存刷新] 检测到日期变化（跨天）: {self.cached_date} -> {current_date}，刷新缓存")
            self.preload_monthly_data()
    
    def preload_monthly_data(self):
        """预加载当月所有员工的每日数据（提升响应速度）"""
        try:
            now = datetime.now()
            year = now.year
            month = now.month
            
            print(f"[预加载] 开始加载 {year}年{month}月 的所有员工每日数据...")
            
            resp = requests.get(
                f"{SERVER_URL}/monthly_daily_stats",
                params={"year": year, "month": month},
                timeout=REQUEST_TIMEOUT
            )
            
            if resp.status_code == 200:
                data = resp.json()
                employees = data.get("employees", [])
                
                # 缓存每个员工的每日数据和显示名称
                self.monthly_data_cache = {}
                self.employee_name_cache = {}  # 缓存员工显示名称
                
                for emp in employees:
                    employee_id = emp.get("employee_id")
                    employee_name = emp.get("employee_name", employee_id)
                    daily_data = emp.get("daily_data", [])
                    
                    self.monthly_data_cache[employee_id] = daily_data
                    self.employee_name_cache[employee_id] = employee_name
                
                self.monthly_data_loaded = True
                self.cached_month = (year, month)
                self.cached_date = now.date()  # 记录缓存加载的日期，用于检测跨天
                
                total_records = sum(len(v) for v in self.monthly_data_cache.values())
                print(f"[预加载] ✓ 成功加载 {len(employees)} 个员工的月度数据，共 {total_records} 条每日记录")
            else:
                print(f"[预加载] ✗ 服务器返回错误: {resp.status_code}")
                self.monthly_data_loaded = False
        except Exception as e:
            print(f"[预加载] ✗ 加载月度数据失败: {e}")
            self.monthly_data_loaded = False
    
    def preload_color_configs(self):
        """预加载颜色配置（提升响应速度）"""
        try:
            print("[预加载] 开始加载颜色配置...")
            
            resp = requests.get(
                f"{SERVER_URL}/color_configs",
                timeout=REQUEST_TIMEOUT
            )
            
            if resp.status_code == 200:
                data = resp.json()
                self.global_line_color_cache = data.get("global_line_color", "#FFA500")
                employee_colors = data.get("employee_colors", {})
                
                # 缓存所有员工的颜色配置
                self.color_cache = {}
                for employee_id, color_data in employee_colors.items():
                    self.color_cache[employee_id] = {
                        "bar": color_data.get("bar_color", "#4CAF50"),
                        "line": self.global_line_color_cache
                    }
                
                self.color_cache_loaded = True
                print(f"[预加载] ✓ 成功加载 {len(self.color_cache)} 个员工的颜色配置")
            else:
                print(f"[预加载] ✗ 颜色配置加载失败: {resp.status_code}")
                self.color_cache_loaded = False
        except Exception as e:
            print(f"[预加载] ✗ 颜色配置加载失败: {e}")
            self.color_cache_loaded = False
    
    def extract_data_from_monthly_cache(self, period, start_date=None, end_date=None):
        """从月度缓存中提取指定时间段的数据（快速响应，无需请求服务器）
        
        Args:
            period: 时间周期 'day'/'week'/'month'/'custom'
            start_date: 开始日期字符串 "YYYY-MM-DD"（仅custom需要）
            end_date: 结束日期字符串 "YYYY-MM-DD"（仅custom需要）
        
        Returns:
            list: 员工统计数据列表，格式同 /stats_by_employee 接口
        """
        if not self.monthly_data_loaded or not self.monthly_data_cache:
            print("[缓存提取] 月度数据未加载，无法从缓存提取")
            return None
        
        # 确定需要筛选的日期范围
        now = datetime.now()
        if period == 'day':
            # 今日
            target_dates = [now.strftime("%Y-%m-%d")]
        elif period == 'yesterday':
            # 昨日（新增）
            yesterday = now - timedelta(days=1)
            target_dates = [yesterday.strftime("%Y-%m-%d")]
        elif period == 'week':
            # 本周（从周一到今天）
            today = now
            week_start = today - timedelta(days=today.weekday())
            target_dates = [(week_start + timedelta(days=i)).strftime("%Y-%m-%d") 
                          for i in range((today - week_start).days + 1)]
        elif period == 'month':
            # 本月（从1号到今天）
            month_start = now.replace(day=1)
            target_dates = [(month_start + timedelta(days=i)).strftime("%Y-%m-%d")
                          for i in range((now - month_start).days + 1)]
        elif period == 'custom' and start_date and end_date:
            # 自定义时间段
            start = datetime.strptime(start_date, "%Y-%m-%d")
            end = datetime.strptime(end_date, "%Y-%m-%d")
            target_dates = [(start + timedelta(days=i)).strftime("%Y-%m-%d")
                          for i in range((end - start).days + 1)]
        else:
            print("[缓存提取] 未知的时间周期或缺少参数")
            return None
        
        print(f"[缓存提取] 从缓存中提取 {period} 数据，日期范围: {target_dates[0]} 至 {target_dates[-1]}")
        
        # 从缓存中提取并聚合数据
        result = []
        for employee_id, daily_data in self.monthly_data_cache.items():
            # 筛选出目标日期范围内的数据
            filtered_data = [d for d in daily_data if d["date"] in target_dates]
            
            if not filtered_data:
                continue
            
            # 聚合统计
            total_consult = sum(d["total_consult"] for d in filtered_data)
            avg_reply_sum = sum(d["avg_reply"] for d in filtered_data)
            avg_reply_count = len([d for d in filtered_data if d["avg_reply"] > 0])
            avg_reply = int(avg_reply_sum / avg_reply_count) if avg_reply_count > 0 else 0
            
            # 计算效率
            if avg_reply > 0 and total_consult > 0:
                efficiency = total_consult / avg_reply
            else:
                efficiency = 0.0
            
            # 从缓存中获取员工显示名称
            employee_name = self.employee_name_cache.get(employee_id, employee_id)
            
            result.append({
                "employee_id": employee_id,
                "employee_name": employee_name,
                "total_consult": total_consult,
                "avg_reply": avg_reply,
                "efficiency": round(efficiency, 2)
            })
        
        # 按咨询量降序排序
        result.sort(key=lambda x: x["total_consult"], reverse=True)
        
        print(f"[缓存提取] ✓ 成功提取 {len(result)} 个员工的数据")
        return result
    
    def init_chart(self):
        """初始化ECharts图表HTML"""
        html_content = self.get_chart_html([], "", "")
        self.web_view_overview.setHtml(html_content)
    
    def get_echarts_script(self):
        """获取ECharts脚本（优先本地，fallback到CDN）"""
        # 尝试从本地加载ECharts
        local_echarts_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'echarts.min.js')
        
        # 打包后的路径
        if getattr(sys, 'frozen', False):
            # 运行在打包后的exe中
            bundle_dir = sys._MEIPASS
            local_echarts_path = os.path.join(bundle_dir, 'echarts.min.js')
        
        if os.path.exists(local_echarts_path):
            try:
                with open(local_echarts_path, 'r', encoding='utf-8') as f:
                    echarts_code = f.read()
                print("[本地ECharts] 使用本地ECharts库（秒加载）")
                return f'<script>{echarts_code}</script>'
            except Exception as e:
                print(f"[本地ECharts] 读取失败: {e}，fallback到CDN")
        
        print("[CDN ECharts] 使用CDN加载ECharts")
        return '<script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"></script>'
    
    def get_chart_html(self, data, period_name, date_range, single_mode=False):
        """生成包含ECharts图表的HTML"""
        # 获取ECharts脚本（本地或CDN）
        echarts_script = self.get_echarts_script()
        
        if not data:
            return '''
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>员工效率统计</title>
    ''' + echarts_script + '''
    <style>
        body {
            margin: 0;
            padding: 10px;
            font-family: "Microsoft YaHei", Arial, sans-serif;
        }
        #chart {
            width: 100%;
            height: 700px;
        }
        #loading {
            text-align: center;
            padding-top: 300px;
            font-size: 16px;
            color: #666;
        }
    </style>
</head>
<body>
    <div id="loading">正在加载图表库...</div>
    <div id="chart" style="width:100%;height:700px;display:none;"></div>
    <script>
        var loadStartTime = Date.now();
        var maxLoadTime = 15000; // 15秒超时
        
        function initChart() {
            // 检查超时
            if (Date.now() - loadStartTime > maxLoadTime) {
                var loading = document.getElementById('loading');
                if (loading) {
                    loading.innerHTML = '图表库加载超时，请检查网络连接';
                    loading.style.color = 'red';
                }
                return;
            }
            
            if (typeof echarts === 'undefined') {
                setTimeout(initChart, 100);
                return;
            }
            document.getElementById('loading').style.display = 'none';
            document.getElementById('chart').style.display = 'block';
            var chartDom = document.getElementById('chart');
            var myChart = echarts.init(chartDom);
            myChart.setOption({
                title: {
                    text: '暂无数据',
                    left: 'center',
                    top: 'middle',
                    textStyle: { fontSize: 24, color: '#999' }
                }
            });
        }
        
        // 等待ECharts加载完成后初始化图表
        function tryInitChart() {
            if (typeof echarts !== 'undefined') {
                initChart();
            } else {
                // 检查超时
                if (Date.now() - loadStartTime > maxLoadTime) {
                    var loading = document.getElementById('loading');
                    if (loading) {
                        loading.innerHTML = '图表库加载超时，请检查网络连接';
                        loading.style.color = 'red';
                    }
                    return;
                }
                setTimeout(tryInitChart, 100);
            }
        }
        
        // 如果ECharts已经加载，直接初始化
        if (typeof echarts !== 'undefined') {
            initChart();
        } else {
            // 等待加载
            setTimeout(tryInitChart, 500);
        }
    </script>
</body>
</html>
'''
        
        # 准备数据
        employee_names = [item["employee_name"] for item in data]
        employee_ids = [item.get("employee_id", item["employee_name"]) for item in data]
        total_consults = [item["total_consult"] for item in data]
        avg_replies = [item["avg_reply"] for item in data]
        efficiencies = [item["efficiency"] for item in data]
        
        # 单员工模式标题调整
        if single_mode and data:
            title_suffix = f" - {employee_names[0]}"
        else:
            title_suffix = ""
        
        # 计算Y轴最大值（由于使用stack会叠加数据，需要手动设置最大值）
        max_consult = max(total_consults) if total_consults and len(total_consults) > 0 else 100
        y_axis_max = int(max_consult * 1.1) if max_consult > 0 else 110  # 留10%的边距
        
        # 获取员工自定义颜色配置
        bar_colors, line_colors = self.get_employee_colors(employee_ids)
        line_color = line_colors[0] if line_colors else '#FFA500'
        
        # 如果没有配置颜色，使用默认颜色循环
        if not bar_colors or len(bar_colors) != len(employee_names):
            bar_color_palette = ['#4CAF50', '#2196F3', '#F44336']
            bar_colors = [bar_color_palette[i % 3] for i in range(len(employee_names))]
        
        # 构建柱状图数据（包含值和标签信息）
        bar_data = []
        for i, (consult, avg_reply) in enumerate(zip(total_consults, avg_replies)):
            bar_data.append({
                "value": consult,
                "label": f"{consult}\\n({avg_reply}秒)"
            })
        
        # 使用字符串格式化避免f-string中的嵌套问题
        # 如果 WebChannel 可用，加载相关脚本
        webchannel_script = '<script src="qrc:///qtwebchannel/qwebchannel.js"></script>' if WEBCHANNEL_AVAILABLE else ''
        
        html_template = '''
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>员工效率统计</title>
    __WEBCHANNEL_SCRIPT__
    ''' + echarts_script + '''
    <style>
        body {
            margin: 0;
            padding: 10px;
            font-family: "Microsoft YaHei", Arial, sans-serif;
        }
        #chart {
            width: 100%;
            height: 700px;
        }
        #loading {
            text-align: center;
            padding-top: 300px;
            font-size: 16px;
            color: #666;
        }
    </style>
</head>
<body>
    <div id="loading">正在加载图表库...</div>
    <div id="chart" style="width:100%;height:700px;display:none;"></div>
    <script>
        var employeeNames = __EMPLOYEE_NAMES__;
        var employeeIds = __EMPLOYEE_IDS__;
        var totalConsults = __TOTAL_CONSULTS__;
        var avgReplies = __AVG_REPLIES__;
        var efficiencies = __EFFICIENCIES__;
        var barColors = __BAR_COLORS__;
        var lineColor = __LINE_COLOR__;
        var loadStartTime = Date.now();
        var maxLoadTime = 15000; // 15秒超时
        
        // 初始化 WebChannel（如果可用）
        var bridge = null;
        var webChannelReady = false;
        if (typeof QWebChannel !== 'undefined' && typeof qt !== 'undefined') {
            try {
                new QWebChannel(qt.webChannelTransport, function(channel) {
                    bridge = channel.objects.bridge;
                    window.bridge = bridge;
                    webChannelReady = true;
                });
            } catch (e) {
            }
        }
        
        function initChart() {
            // 检查超时
            if (Date.now() - loadStartTime > maxLoadTime) {
                var loading = document.getElementById('loading');
                if (loading) {
                    loading.innerHTML = '图表库加载超时，请检查网络连接';
                    loading.style.color = 'red';
                }
                return;
            }
            
            // 检查echarts是否已加载
            if (typeof echarts === 'undefined') {
                setTimeout(initChart, 100);
                return;
            }
            
            // 隐藏加载提示，显示图表
            document.getElementById('loading').style.display = 'none';
            document.getElementById('chart').style.display = 'block';
            
            var chartDom = document.getElementById('chart');
            // 使用暗色主题或默认主题，加快渲染
            var myChart = echarts.init(chartDom, null, {
                renderer: 'canvas',
                useDirtyRect: false  // 禁用脏矩形优化，加快首次渲染
            });
            
            var option = {
                title: {
                    text: '员工工作效率统计 - __PERIOD_NAME__',
                    subtext: '__DATE_RANGE__',
                    left: 'center',
                    top: 15,
                    textStyle: { 
                        fontSize: 18, 
                        fontWeight: 'bold',
                        color: '#333'
                    },
                    subtextStyle: {
                        fontSize: 13,
                        color: '#666',
                        padding: [0, 0, 10, 0]
                    },
                    itemGap: 8
                },
                tooltip: {
                    trigger: 'axis',
                    axisPointer: {
                        type: 'cross',
                        crossStyle: {
                            color: '#999'
                        }
                    },
                    backgroundColor: 'rgba(50, 50, 50, 0.9)',
                    borderColor: '#333',
                    borderWidth: 1,
                    textStyle: {
                        color: '#fff',
                        fontSize: 13
                    },
                    formatter: function(params) {
                        var result = '<div style="font-weight: bold; margin-bottom: 5px; font-size: 14px;">' + params[0].axisValue + '</div>';
                        
                        // 只处理"咨询数量"和"效率指标"，跳过"标签层"
                        params.forEach(function(item) {
                            if (item.seriesName === '咨询数量') {
                                var index = employeeNames.indexOf(item.axisValue);
                                var avgReply = avgReplies[index];
                                // 咨询数量
                                result += '<div style="margin: 3px 0;">' + item.marker + '<span style="margin-left: 5px; color: #fff; font-size: 13px;">咨询数量: <strong>' + item.value + '</strong> 次</span></div>';
                                // 平均回复时长（样式与咨询数量一致）
                                result += '<div style="margin: 3px 0; margin-left: 20px;"><span style="color: #fff; font-size: 13px;">平均回复时长: <strong>' + avgReply + '</strong> 秒</span></div>';
                            } else if (item.seriesName === '效率指标') {
                                // 效率指标（根据时间周期显示不同文字）
                                result += '<div style="margin: 3px 0;">' + item.marker + '<span style="margin-left: 5px; color: #fff; font-size: 13px;">__PERIOD_TOOLTIP__效率指标: <strong>' + item.value.toFixed(2) + '</strong></span></div>';
                            }
                            // 跳过"标签层"
                        });
                        return result;
                    }
                },
                legend: {
                    data: ['咨询数量', '效率指标'],
                    top: 65,
                    itemGap: 25,
                    textStyle: {
                        fontSize: 13,
                        fontWeight: 'bold'
                    },
                    icon: 'rect'
                },
                grid: {
                    left: '8%',
                    right: '8%',
                    bottom: '15%',
                    top: '22%',
                    containLabel: true
                },
                xAxis: {
                    type: 'category',
                    data: employeeNames,
                    boundaryGap: true,  // 确保类别之间有间距，对齐柱状图
                    axisLabel: {
                        rotate: 0,
                        interval: 0,
                        formatter: function(value) {
                            // 如果名字太长，使用省略号，避免换行
                            if (value.length > 10) {
                                return value.substring(0, 10) + '...';
                            }
                            return value;
                        },
                        overflow: 'truncate',
                        width: 100,
                        ellipsis: '...'
                    },
                    axisTick: {
                        alignWithLabel: true,
                        interval: 0  // 确保每个类别都有刻度线
                    }
                },
                yAxis: [
                    {
                        type: 'value',
                        name: '咨询数量',
                        position: 'left',
                        axisLabel: { 
                            color: '#4CAF50',
                            fontSize: 12
                        },
                        nameTextStyle: { 
                            color: '#4CAF50',
                            fontSize: 13,
                            fontWeight: 'bold'
                        },
                        splitLine: { 
                            show: true, 
                            lineStyle: { 
                                color: '#e6e6e6',
                                type: 'dashed',
                                width: 1
                            } 
                        },
                        scale: false,  // 不从0开始，更清晰显示差异
                        // 手动设置最大值，避免stack导致的数据叠加影响Y轴范围
                        max: __Y_AXIS_MAX__
                    },
                    {
                        type: 'value',
                        name: '效率指标',
                        position: 'right',
                        axisLabel: { 
                            color: '#FFA500',
                            fontSize: 12
                        },
                        nameTextStyle: { 
                            color: '#FFA500',
                            fontSize: 13,
                            fontWeight: 'bold'
                        },
                        splitLine: { show: false }  // 右Y轴不显示分割线，避免混乱
                    }
                ],
                series: [
                    {
                        name: '咨询数量',
                        type: 'bar',
                        yAxisIndex: 0,
                        z: 1,  // 柱状图在底层
                        stack: 'barStack',  // 使用stack让两个bar系列完全重叠
                        barWidth: '60%',
                        data: totalConsults.map(function(value, index) {
                            return {
                                value: value,
                                itemStyle: { color: barColors[index] }
                            };
                        }),
                        label: {
                            show: false  // 原始柱状图不显示标签
                        },
                        emphasis: {
                            itemStyle: {
                                shadowBlur: 10,
                                shadowOffsetX: 0,
                                shadowColor: 'rgba(0, 0, 0, 0.5)'
                            }
                        }
                    },
                    {
                        name: '效率指标',
                        type: 'line',
                        yAxisIndex: 1,
                        data: efficiencies,
                        z: 2,  // 折线图在中间层（显示在柱状图前面）
                        lineStyle: {
                            width: 3,
                            color: lineColor
                        },
                        itemStyle: {
                            color: lineColor,
                            borderWidth: 2,
                            borderColor: '#fff'
                        },
                        symbol: 'circle',
                        symbolSize: 8,
                        label: {
                            show: false  // 不显示折线图的数字标签
                        },
                        emphasis: {
                            lineStyle: {
                                width: 4
                            },
                            itemStyle: {
                                borderWidth: 3,
                                borderColor: '#fff',
                                shadowBlur: 10,
                                shadowColor: 'rgba(255, 165, 0, 0.6)'
                            }
                        }
                    },
                    {
                        // 单独的透明柱状图系列，专门用于显示标签（最顶层）
                        // 不使用stack，独立显示，确保标签位置准确在原始柱状图顶部
                        name: '标签层',
                        type: 'bar',
                        yAxisIndex: 0,
                        z: 3,  // 标签在最顶层
                        barGap: '-100%',  // 完全重叠在原始柱状图上
                        barWidth: '60%',  // 与原始柱状图完全一致
                        // 使用相同的数据值，确保标签位置与柱状图顶部完全对齐
                        data: totalConsults.map(function(value, index) {
                            return {
                                value: value,  // 使用原始柱状图的值
                                itemStyle: { 
                                    color: 'transparent',  // 柱体完全透明
                                    borderColor: 'transparent',
                                    borderWidth: 0
                                }
                            };
                        }),
                        label: {
                            show: true,
                            position: 'top',  // 标签在柱状图顶部上方
                            formatter: function(params) {
                                var index = params.dataIndex;
                                return '人数：' + totalConsults[index] + '\\n时长：' + avgReplies[index] + '秒';
                            },
                            fontSize: 11,
                            fontWeight: 'bold',
                            color: '#333',
                            offset: [0, 0],  // 紧贴柱状图顶部
                            lineHeight: 16,
                            // 透明背景，只显示文字
                            backgroundColor: 'transparent',
                            borderColor: 'transparent',
                            borderWidth: 0,
                            padding: [2, 5, 0, 5],
                            // 文字描边，让文字在折线上也能看清楚
                            textBorderColor: '#fff',
                            textBorderWidth: 2
                        },
                        // 不显示在图例中，不响应鼠标事件
                        silent: true,
                        legendHoverLink: false
                    }
                ]
            };
            
            myChart.setOption(option);
            
            // 禁用图表右键菜单
            chartDom.oncontextmenu = function(e) {
                e.preventDefault();
                return false;
            };
            
            // 添加双击事件监听（支持双向切换 + 防抖）
            var singleMode = __SINGLE_MODE__;
            var doubleClickLocked = false;  // 防抖锁
            
            myChart.on('dblclick', function(params) {
                if (doubleClickLocked) {
                    return;
                }
                
                // 只处理柱状图的双击事件
                if (params.componentType === 'series' && 
                    (params.seriesName === '咨询数量' || params.seriesName === '标签层')) {
                    
                    // 锁定双击事件（1.5秒内不响应新的双击）
                    doubleClickLocked = true;
                    setTimeout(function() {
                        doubleClickLocked = false;
                    }, 1500);
                    
                    if (singleMode) {
                        function tryBackToOverview(retryCount) {
                            if (window.bridge && typeof window.bridge.onBackToOverview === 'function') {
                                window.bridge.onBackToOverview();
                            } else if (retryCount > 0) {
                                setTimeout(function() { tryBackToOverview(retryCount - 1); }, 200);
                            }
                        }
                        tryBackToOverview(5);
                    } else {
                        // 总览模式 → 双击查看单员工详情
                        var employeeIndex = params.dataIndex;
                        var employeeName = employeeNames[employeeIndex];
                        var employeeId = employeeIds[employeeIndex];
                        
                        function tryShowEmployee(retryCount) {
                            if (window.bridge && typeof window.bridge.onEmployeeDoubleClick === 'function') {
                                window.bridge.onEmployeeDoubleClick(employeeId, employeeName);
                            } else if (retryCount > 0) {
                                setTimeout(function() { tryShowEmployee(retryCount - 1); }, 200);
                            }
                        }
                        tryShowEmployee(5);
                    }
                }
            });
            
            // 响应式调整
            window.addEventListener('resize', function() {
                myChart.resize();
            });
            
            // 将图表实例暴露为全局变量（方案3：图表实例复用）
            window.chartInstance = myChart;
        }
        
        // 方案3：只更新数据的函数（不重新创建图表）
        window.updateChartData = function(newData) {
            if (!window.chartInstance) return;
            
            // 更新全局变量
            employeeNames = newData.employeeNames;
            employeeIds = newData.employeeIds;
            totalConsults = newData.totalConsults;
            avgReplies = newData.avgReplies;
            efficiencies = newData.efficiencies;
            barColors = newData.barColors;
            lineColor = newData.lineColor;
            
            // 只更新数据，不重新创建图表
            window.chartInstance.setOption({
                title: {
                    text: '员工工作效率统计 - ' + newData.periodName,
                    subtext: newData.dateRange
                },
                xAxis: {
                    data: employeeNames
                },
                yAxis: [
                    {
                        max: newData.yAxisMax
                    },
                    {}
                ],
                series: [
                    {
                        data: totalConsults.map(function(value, index) {
                            return {
                                value: value,
                                itemStyle: { color: barColors[index] }
                            };
                        })
                    },
                    {
                        data: efficiencies,
                        lineStyle: {
                            color: lineColor
                        },
                        itemStyle: {
                            color: lineColor
                        }
                    },
                    {
                        data: totalConsults.map(function(value, index) {
                            return {
                                value: value,
                                itemStyle: { 
                                    color: 'transparent',
                                    borderColor: 'transparent',
                                    borderWidth: 0
                                }
                            };
                        })
                    }
                ]
            });
        };
        
        // 等待ECharts加载完成后初始化图表
        function tryInitChart() {
            if (typeof echarts !== 'undefined') {
                initChart();
            } else {
                // 检查超时
                if (Date.now() - loadStartTime > maxLoadTime) {
                    var loading = document.getElementById('loading');
                    if (loading) {
                        loading.innerHTML = '图表库加载超时，请检查网络连接';
                        loading.style.color = 'red';
                    }
                    return;
                }
                setTimeout(tryInitChart, 100);
            }
        }
        
        // 如果ECharts已经加载，直接初始化
        if (typeof echarts !== 'undefined') {
            initChart();
        } else {
            // 等待加载
            setTimeout(tryInitChart, 500);
        }
    </script>
</body>
</html>
'''
        # 使用format方法替换变量
        # 使用字符串替换避免format方法的占位符问题
        html_content = html_template.replace('__WEBCHANNEL_SCRIPT__', webchannel_script)
        html_content = html_content.replace('__EMPLOYEE_NAMES__', json.dumps(employee_names, ensure_ascii=False))
        html_content = html_content.replace('__EMPLOYEE_IDS__', json.dumps(employee_ids, ensure_ascii=False))
        html_content = html_content.replace('__TOTAL_CONSULTS__', json.dumps(total_consults))
        html_content = html_content.replace('__AVG_REPLIES__', json.dumps(avg_replies))
        html_content = html_content.replace('__EFFICIENCIES__', json.dumps(efficiencies))
        html_content = html_content.replace('__BAR_COLORS__', json.dumps(bar_colors))
        html_content = html_content.replace('__LINE_COLOR__', json.dumps(line_color))
        html_content = html_content.replace('__Y_AXIS_MAX__', str(y_axis_max))
        html_content = html_content.replace('__SINGLE_MODE__', 'true' if single_mode else 'false')
        # 转义特殊字符，避免JavaScript语法错误
        period_name_safe = (period_name + title_suffix).replace("'", "\\'").replace('"', '\\"')
        date_range_safe = date_range.replace("'", "\\'").replace('"', '\\"')
        
        # 生成tooltip专用的时间周期文字
        period_tooltip = period_name
        if period_name == "今日":
            period_tooltip = "当日"
        elif period_name == "自定义":
            period_tooltip = "自定义时间"
        # "本周"和"本月"保持不变
        
        html_content = html_content.replace('__PERIOD_NAME__', period_name_safe)
        html_content = html_content.replace('__PERIOD_TOOLTIP__', period_tooltip)
        html_content = html_content.replace('__DATE_RANGE__', date_range_safe)
        return html_content
    
    def back_to_overview(self):
        """返回到总览视图（通过双击触发） - 瞬间切换"""
        print("[⚡返回总览] 切换到总览")
        self.single_employee_mode = False
        self.current_employee_id = None
        self.current_employee_name = None
        self.setWindowTitle("员工效率统计图表")
        
        # 切换到总览WebView（瞬间）
        self.stacked_widget.setCurrentIndex(0)
        self.web_view = self.web_view_overview
        
        # 刷新总览数据
        QTimer.singleShot(50, self.query_stats)
    
    def show_single_employee(self, employee_id, employee_name):
        """显示单个员工的详细数据（通过双击触发） - 瞬间切换，只更新数据"""
        print(f"[⚡切换员工] {employee_name} (ID: {employee_id})")
        
        # 更新状态
        self.single_employee_mode = True
        self.current_employee_id = employee_id
        self.current_employee_name = employee_name
        self.setWindowTitle(f"员工详细数据 - {employee_name}")
        
        # 切换到单员工WebView（瞬间）
        self.stacked_widget.setCurrentIndex(1)
        self.web_view = self.web_view_single
        
        # 修复1：加载员工列表（如果还未加载）
        if self.employee_list.count() == 0:
            self.load_employee_list()
        
        # 修复1：选中当前员工（背景变绿）
        self.select_employee_in_list(employee_id)
        
        # 获取当前时间参数
        params, period_name, date_range = self.get_current_time_params()
        
        # 查询并更新数据（从缓存提取，快速响应）
        QTimer.singleShot(10, lambda: self.query_single_employee_stats(params, period_name, date_range))
    
    def load_employee_list(self):
        """修复1：加载所有员工列表"""
        try:
            self.employee_list.clear()
            # 从缓存或服务器获取员工列表
            if hasattr(self, 'parent') and self.parent() and hasattr(self.parent(), 'current_data'):
                employees = self.parent().current_data
            else:
                # 从服务器获取
                resp = requests.get(f"{SERVER_URL}/employees", timeout=REQUEST_TIMEOUT)
                employees = resp.json() if resp.status_code == 200 else []
            
            self.all_employees = employees
            
            # 添加到列表
            for emp in employees:
                item_text = emp.get("display_name", emp.get("employee_name", "未知"))
                item = QListWidgetItem(item_text)
                item.setData(Qt.UserRole, emp.get("employee_name", ""))  # 存储employee_id
                self.employee_list.addItem(item)
                
            print(f"[员工列表] 加载了 {len(employees)} 个员工")
        except Exception as e:
            print(f"[员工列表] 加载失败: {e}")
    
    def select_employee_in_list(self, employee_id):
        """修复1：选中指定员工（背景变绿）"""
        for i in range(self.employee_list.count()):
            item = self.employee_list.item(i)
            if item.data(Qt.UserRole) == employee_id:
                self.employee_list.setCurrentItem(item)
                break
    
    def on_employee_list_clicked(self, item):
        """修复1：点击员工列表项，切换到该员工"""
        employee_id = item.data(Qt.UserRole)
        employee_name = item.text()
        
        # 如果已经是当前员工，不重复切换
        if employee_id == self.current_employee_id:
            return
        
        print(f"[员工列表] 切换到: {employee_name}")
        
        # 更新状态
        self.current_employee_id = employee_id
        self.current_employee_name = employee_name
        self.setWindowTitle(f"员工详细数据 - {employee_name}")
        
        # 获取当前时间参数
        params, period_name, date_range = self.get_current_time_params()
        
        # 查询并更新数据
        QTimer.singleShot(10, lambda: self.query_single_employee_stats(params, period_name, date_range))
    
    def safe_query_stats(self):
        """安全地查询统计数据（带异常处理）"""
        try:
            self.query_stats()
        except Exception as e:
            print(f"[查询失败] {e}")
            import traceback
            traceback.print_exc()
            # 不显示错误框，只是打印日志
    
    def get_current_time_params(self):
        """获取当前选择的时间参数（统一方法，确保总览和单员工视图时间同步）"""
        period = self.period_group.checkedId()
        params = {}
        period_name = "今日"
        date_range = ""
        
        if period == 0:
            # 今日
            params["period"] = "day"
            period_name = "今日"
            date_range = datetime.now().strftime("%Y-%m-%d")
        elif period == 1:
            # 昨日
            params["period"] = "yesterday"  # 新增：昨日
            period_name = "昨日"
            yesterday = datetime.now() - timedelta(days=1)
            date_range = yesterday.strftime("%Y-%m-%d")
        elif period == 2:
            # 本周
            params["period"] = "week"
            period_name = "本周"
            today = datetime.now()
            week_start = (today - timedelta(days=today.weekday())).strftime("%Y-%m-%d")
            date_range = f"{week_start} 至 {today.strftime('%Y-%m-%d')}"
        elif period == 3:
            # 本月
            params["period"] = "month"
            period_name = "本月"
            today = datetime.now()
            month_start = today.replace(day=1).strftime("%Y-%m-%d")
            date_range = f"{month_start} 至 {today.strftime('%Y-%m-%d')}"
        elif period == 4:
            # 自定义
            params["period"] = "custom"
            start_str = self.start_date.date().toString("yyyy-MM-dd")
            end_str = self.end_date.date().toString("yyyy-MM-dd")
            params["start"] = start_str
            params["end"] = end_str
            period_name = "自定义"
            date_range = f"{start_str} 至 {end_str}"
        
        return params, period_name, date_range
    
    def query_stats(self):
        """查询统计数据（根据当前模式：总览 or 单员工）"""
        # 修复：在查询前检查日期变化，确保跨天后刷新缓存
        self.check_and_refresh_monthly_cache()
        
        params, period_name, date_range = self.get_current_time_params()
        
        # 判断当前模式：单员工 or 总览
        if self.single_employee_mode and self.current_employee_id:
            # 单员工模式：更新单员工数据
            self.query_single_employee_stats(params, period_name, date_range)
        else:
            # 总览模式：更新所有员工数据
            self.query_overview_stats(params, period_name, date_range)
    
    def query_overview_stats(self, params, period_name, date_range):
        """查询总览数据（所有员工）- 优先使用缓存，实现快速响应"""
        try:
            import time
            start_time = time.time()
            
            # 获取当前选择的时间周期
            period = params.get("period", "day")
            start_date = params.get("start")
            end_date = params.get("end")
            
            # 尝试从缓存中提取数据（快速模式）
            data = None
            use_cache = False
            
            # 判断是否可以使用缓存（检查月份和日期，确保跨天后刷新）
            now = datetime.now()
            current_month = (now.year, now.month)
            current_date = now.date()
            cache_valid = (self.monthly_data_loaded and 
                          self.cached_month == current_month and
                          self.cached_date == current_date)  # 修复：也检查日期，确保跨天后刷新
            
            # 修复：查询"今日"时，总是从服务器获取实时数据，不使用缓存
            # 因为今天的数据会持续更新，需要实时显示
            # "昨日"是历史数据，可以使用缓存
            if cache_valid and period != 'day':
                # 对于昨日/周/月，可以从缓存提取（历史数据不会变化）
                if period in ['yesterday', 'week', 'month']:
                    data = self.extract_data_from_monthly_cache(period)
                    if data is not None:
                        use_cache = True
                        elapsed = (time.time() - start_time) * 1000
                        print(f"[快速模式] ✓ 从缓存中提取 {period_name} 数据，耗时 {elapsed:.1f}ms")
                
                # 对于自定义时间，检查是否在当月范围内
                elif period == 'custom' and start_date and end_date:
                    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
                    end_dt = datetime.strptime(end_date, "%Y-%m-%d")
                    month_start = now.replace(day=1)
                    
                    # 如果自定义时间完全在当月内，使用缓存
                    if start_dt >= month_start and end_dt <= now:
                        data = self.extract_data_from_monthly_cache(period, start_date, end_date)
                        if data is not None:
                            use_cache = True
                            elapsed = (time.time() - start_time) * 1000
                            print(f"[快速模式] ✓ 从缓存中提取自定义时间数据，耗时 {elapsed:.1f}ms")
            
            # 如果缓存不可用或提取失败，从服务器请求（兼容模式）
            if data is None:
                print(f"[服务器模式] 从服务器请求 {period_name} 数据...")
                resp = requests.get(f"{SERVER_URL}/stats_by_employee", params=params, timeout=REQUEST_TIMEOUT)
                data = resp.json() if resp.status_code == 200 else []
                elapsed = (time.time() - start_time) * 1000
                print(f"[服务器模式] 请求完成，耗时 {elapsed:.1f}ms")
            
            if not data:
                # 没有数据时，显示空数据HTML，但重置chart_initialized，下次有数据时重新加载
                html_content = self.get_chart_html([], period_name, date_range)
                self.web_view_overview.setHtml(html_content)
                self.chart_initialized = False  # 重置标志，下次有数据时重新加载完整HTML
                self.last_mode = 'overview'
                print("[提示] 当前时间段没有数据")
                return
            
            # 应用排序
            data = self.apply_employee_order(data)
            
            # 应用隐藏规则（只在"今日"模式）
            if self.period_group.checkedId() == 0:  # 今日
                data = self.apply_visibility_filter(data)
            
            # 分模式复用逻辑
            if not self.chart_initialized or self.last_mode != 'overview':
                # 首次加载 或 模式切换 或 上次数据为空 → 重新加载完整HTML
                mode_text = "缓存" if use_cache else "服务器"
                print(f"[总览-{mode_text}] 重新加载完整HTML")
                html_content = self.get_chart_html(data, period_name, date_range)
                self.web_view_overview.setHtml(html_content)
                self.chart_initialized = True
                self.last_mode = 'overview'
            else:
                # 同一模式内切换时间周期 → 只更新数据（0.1秒）
                mode_text = "缓存" if use_cache else "服务器"
                print(f"[总览-{mode_text}] 快速更新数据")
                self.update_chart_data_only(data, period_name, date_range, single_mode=False)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"查询总览数据失败: {str(e)}")
    
    def query_single_employee_stats(self, params, period_name, date_range):
        """查询单员工数据（刷新当前员工视图）- 优先使用缓存，快速更新"""
        if not self.current_employee_id:
            print("[错误] 当前没有选中的员工")
            return
        
        try:
            import time
            start_time = time.time()
            
            print(f"[单员工] 查询数据: {self.current_employee_name} ({period_name})")
            
            # 获取当前选择的时间周期
            period = params.get("period", "day")
            start_date = params.get("start")
            end_date = params.get("end")
            
            # 尝试从缓存中提取数据（快速模式）
            all_data = None
            use_cache = False
            
            # 判断是否可以使用缓存（检查月份和日期，确保跨天后刷新）
            now = datetime.now()
            current_month = (now.year, now.month)
            current_date = now.date()
            cache_valid = (self.monthly_data_loaded and 
                          self.cached_month == current_month and
                          self.cached_date == current_date)  # 修复：也检查日期，确保跨天后刷新
            
            # 修复：查询"今日"时，总是从服务器获取实时数据，不使用缓存
            # 因为今天的数据会持续更新，需要实时显示
            # "昨日"是历史数据，可以使用缓存
            if cache_valid and period != 'day':
                if period in ['yesterday', 'week', 'month']:
                    all_data = self.extract_data_from_monthly_cache(period)
                    if all_data is not None:
                        use_cache = True
                        elapsed = (time.time() - start_time) * 1000
                        print(f"[单员工-快速模式] ✓ 从缓存提取数据，耗时 {elapsed:.1f}ms")
                elif period == 'custom' and start_date and end_date:
                    start_dt = datetime.strptime(start_date, "%Y-%m-%d")
                    end_dt = datetime.strptime(end_date, "%Y-%m-%d")
                    month_start = now.replace(day=1)
                    if start_dt >= month_start and end_dt <= now:
                        all_data = self.extract_data_from_monthly_cache(period, start_date, end_date)
                        if all_data is not None:
                            use_cache = True
                            elapsed = (time.time() - start_time) * 1000
                            print(f"[单员工-快速模式] ✓ 从缓存提取自定义数据，耗时 {elapsed:.1f}ms")
            
            # 如果缓存不可用，从服务器请求
            if all_data is None:
                print(f"[单员工-服务器模式] 从服务器请求数据...")
                resp = requests.get(f"{SERVER_URL}/stats_by_employee", params=params, timeout=REQUEST_TIMEOUT)
                all_data = resp.json() if resp.status_code == 200 else []
                elapsed = (time.time() - start_time) * 1000
                print(f"[单员工-服务器模式] 请求完成，耗时 {elapsed:.1f}ms")
            
            # 筛选出当前员工的数据
            data = [item for item in all_data if item.get("employee_id") == self.current_employee_id]
            
            if not data:
                print("[单员工] 当前时间段没有数据")
                # 显示空数据
                if not self.single_chart_initialized:
                    html_content = self.get_chart_html([], period_name, date_range, single_mode=True)
                    self.web_view_single.setHtml(html_content)
                    self.single_chart_initialized = True
                return
            
            # 首次加载或模式切换：加载完整HTML
            if not self.single_chart_initialized:
                mode_text = "缓存" if use_cache else "服务器"
                print(f"[单员工-{mode_text}] 首次加载，生成完整HTML")
                html_content = self.get_chart_html(data, period_name, date_range, single_mode=True)
                self.web_view_single.setHtml(html_content)
                self.single_chart_initialized = True
            else:
                # 已初始化：只更新数据（快速模式）
                mode_text = "缓存" if use_cache else "服务器"
                print(f"[单员工-{mode_text}] 快速更新数据")
                self.update_chart_data_only(data, period_name, date_range, single_mode=True)
        
        except Exception as e:
            print(f"[错误] 查询单员工数据失败: {e}")
            import traceback
            traceback.print_exc()
    
    def update_chart_data_only(self, data, period_name, date_range, single_mode=False):
        """只更新图表数据，不重新加载HTML（方案3：图表实例复用）"""
        if not data:
            return
        
        # 准备数据
        employee_names = [item["employee_name"] for item in data]
        employee_ids = [item.get("employee_id", item["employee_name"]) for item in data]
        total_consults = [item["total_consult"] for item in data]
        avg_replies = [item["avg_reply"] for item in data]
        efficiencies = [item["efficiency"] for item in data]
        
        # 计算Y轴最大值
        max_consult = max(total_consults) if total_consults else 100
        y_axis_max = int(max_consult * 1.1) if max_consult > 0 else 110
        
        # 获取颜色配置
        bar_colors, line_colors = self.get_employee_colors(employee_ids)
        line_color = line_colors[0] if line_colors else '#FFA500'
        
        if not bar_colors or len(bar_colors) != len(employee_names):
            bar_color_palette = ['#4CAF50', '#2196F3', '#F44336']
            bar_colors = [bar_color_palette[i % 3] for i in range(len(employee_names))]
        
        # 构建JavaScript数据对象
        import json
        new_data = {
            'employeeNames': employee_names,
            'employeeIds': employee_ids,
            'totalConsults': total_consults,
            'avgReplies': avg_replies,
            'efficiencies': efficiencies,
            'barColors': bar_colors,
            'lineColor': line_color,
            'periodName': period_name,
            'dateRange': date_range,
            'yAxisMax': y_axis_max
        }
        
        # 使用当前的WebView
        target_webview = self.web_view
        
        # 调用JavaScript函数更新数据
        js_code = f"if (window.updateChartData) {{ window.updateChartData({json.dumps(new_data)}); }}"
        target_webview.page().runJavaScript(js_code)
    
    def get_employee_colors(self, employee_ids):
        """获取员工颜色配置（优先使用缓存）"""
        try:
            # 如果缓存已加载，直接从缓存获取（快速模式）
            if self.color_cache_loaded and self.color_cache:
                bar_colors = []
                for emp_id in employee_ids:
                    if emp_id in self.color_cache:
                        bar_colors.append(self.color_cache[emp_id].get("bar", "#4CAF50"))
                    else:
                        bar_colors.append("#4CAF50")  # 默认绿色
                
                # 所有员工使用统一的线形图颜色
                line_colors = [self.global_line_color_cache] * len(employee_ids)
                
                return bar_colors, line_colors
            
            # 缓存未加载，从服务器获取（兼容模式）
            print("[颜色配置] 缓存未加载，从服务器获取...")
            employee_ids_str = ",".join(employee_ids)
            resp = requests.get(
                f"{SERVER_URL}/color_configs",
                params={"employee_ids": employee_ids_str},
                timeout=REQUEST_TIMEOUT
            )
            
            if resp.status_code == 200:
                data = resp.json()
                global_line_color = data.get("global_line_color", "#FFA500")
                employee_colors = data.get("employee_colors", {})
                
                # 构建柱状图颜色列表
                bar_colors = []
                for emp_id in employee_ids:
                    if emp_id in employee_colors:
                        bar_colors.append(employee_colors[emp_id].get("bar_color", "#4CAF50"))
                    else:
                        bar_colors.append("#4CAF50")  # 默认绿色
                
                # 所有员工使用统一的线形图颜色
                line_colors = [global_line_color] * len(employee_ids)
                
                return bar_colors, line_colors
            else:
                # API调用失败，使用默认颜色
                default_bar = '#4CAF50'
                default_line = '#FFA500'
                return [default_bar] * len(employee_ids), [default_line] * len(employee_ids)
        except Exception as e:
            print(f"获取颜色配置错误: {e}")
            # 如果获取失败，使用默认颜色
            default_bar = '#4CAF50'
            default_line = '#FFA500'
            return [default_bar] * len(employee_ids), [default_line] * len(employee_ids)
    
    async def _get_colors_async(self, employee_ids):
        """异步获取员工颜色（通过API）"""
        try:
            # 通过API获取颜色配置
            employee_ids_str = ",".join(employee_ids)
            resp = requests.get(
                f"{SERVER_URL}/color_configs",
                params={"employee_ids": employee_ids_str},
                timeout=REQUEST_TIMEOUT
            )
            
            if resp.status_code == 200:
                data = resp.json()
                global_line_color = data.get("global_line_color", "#FFA500")
                employee_colors = data.get("employee_colors", {})
                
                # 构建柱状图颜色列表
                bar_colors = []
                for emp_id in employee_ids:
                    if emp_id in employee_colors:
                        bar_colors.append(employee_colors[emp_id].get("bar_color", "#4CAF50"))
                    else:
                        bar_colors.append("#4CAF50")  # 默认绿色
                
                # 所有员工使用统一的线形图颜色
                line_colors = [global_line_color] * len(employee_ids)
                
                return bar_colors, line_colors
            else:
                # API调用失败，返回默认颜色
                default_bar = '#4CAF50'
                default_line = '#FFA500'
                return [default_bar] * len(employee_ids), [default_line] * len(employee_ids)
        except Exception as e:
            print(f"获取颜色配置错误: {e}")
            # 返回默认颜色
            default_bar = '#4CAF50'
            default_line = '#FFA500'
            return [default_bar] * len(employee_ids), [default_line] * len(employee_ids)
    
    def open_employee_management_dialog(self):
        """打开员工管理对话框（颜色+排序+隐藏）"""
        try:
            # 获取当前员工列表
            resp = requests.get(f"{SERVER_URL}/employees", timeout=REQUEST_TIMEOUT)
            employees = resp.json() if resp.status_code == 200 else []
            
            if not employees:
                QMessageBox.information(self, "提示", "当前没有员工记录")
                return
            
            dialog = EmployeeManagementDialog(self, employees)
            if dialog.exec_() == QDialog.Accepted:
                # 刷新缓存（颜色配置可能已更改）
                self.preload_color_configs()
                self.preload_monthly_data()  # 也刷新月度数据（可能有名称变更）
                
                # 刷新图表和按钮状态
                self.update_toggle_button_text()
                QTimer.singleShot(200, self.query_stats)  # 延迟查询，确保缓存已刷新
        except Exception as e:
            QMessageBox.critical(self, "错误", f"获取员工列表失败：\n{str(e)}")
    
    def toggle_hidden_employees(self):
        """切换展示/隐藏员工状态"""
        try:
            # 获取当前全局显示模式
            resp = requests.get(f"{SERVER_URL}/global_visibility_mode", timeout=REQUEST_TIMEOUT)
            current_mode = resp.json() if resp.status_code == 200 else {"show_all": False}
            current_show_all = current_mode.get("show_all", False)
            
            # 切换状态
            new_show_all = not current_show_all
            
            # 保存到服务器
            resp = requests.post(
                f"{SERVER_URL}/global_visibility_mode",
                json={"show_all": new_show_all},
                timeout=REQUEST_TIMEOUT
            )
            
            if resp.status_code == 200:
                # 更新按钮文本
                self.update_toggle_button_text()
                # 刷新图表（只在今日模式）
                if self.period_group.checkedId() == 0:  # 今日
                    self.query_stats()
            else:
                QMessageBox.warning(self, "警告", "更新显示模式失败")
        except Exception as e:
            handle_request_error(self, e, "切换显示模式")
    
    def update_toggle_button_text(self):
        """更新展示/隐藏按钮的文本"""
        try:
            resp = requests.get(f"{SERVER_URL}/global_visibility_mode", timeout=REQUEST_TIMEOUT)
            mode = resp.json() if resp.status_code == 200 else {"show_all": False}
            show_all = mode.get("show_all", False)
            
            if show_all:
                self.toggle_hidden_btn.setText("隐藏员工")
            else:
                self.toggle_hidden_btn.setText("展示隐藏员工")
        except Exception as e:
            print(f"更新按钮文本失败: {e}")
    
    def apply_employee_order(self, data):
        """应用员工排序"""
        try:
            # 获取排序配置
            resp = requests.get(f"{SERVER_URL}/employee_order", timeout=REQUEST_TIMEOUT)
            order_data = resp.json() if resp.status_code == 200 else []
            
            # 构建排序字典
            order_dict = {item['employee_id']: item['order'] for item in order_data}
            
            # 对数据排序
            def get_sort_key(item):
                employee_id = item.get("employee_id", item.get("employee_name", ""))
                return order_dict.get(employee_id, 9999)
            
            sorted_data = sorted(data, key=get_sort_key)
            return sorted_data
        except Exception as e:
            print(f"应用排序失败: {e}")
            return data  # 如果排序失败，返回原始数据
    
    def apply_visibility_filter(self, data):
        """应用员工可见性过滤（只在今日模式）- 新版自动隐藏逻辑"""
        try:
            # 获取全局显示模式
            resp = requests.get(f"{SERVER_URL}/global_visibility_mode", timeout=REQUEST_TIMEOUT)
            mode = resp.json() if resp.status_code == 200 else {"show_all": False}
            show_all = mode.get("show_all", False)
            
            if show_all:
                # 显示所有员工，不过滤
                print("[可见性] 显示所有员工")
                return data
            
            # 获取实时在线状态（从/employees接口）
            try:
                resp_realtime = requests.get(f"{SERVER_URL}/employees", timeout=REQUEST_TIMEOUT)
                realtime_data = resp_realtime.json() if resp_realtime.status_code == 200 else []
                online_dict = {emp.get("employee_name", ""): emp.get("online", False) for emp in realtime_data}
                print(f"[可见性] 获取到 {len(online_dict)} 个员工的实时在线状态")
            except Exception as e:
                print(f"[可见性] 获取实时在线状态失败: {e}，默认所有员工在线")
                online_dict = {}
            
            # 获取员工隐藏配置
            resp = requests.get(f"{SERVER_URL}/employee_visibility", timeout=REQUEST_TIMEOUT)
            visibility_data = resp.json() if resp.status_code == 200 else []
            hidden_dict = {item['employee_id']: item['hidden'] for item in visibility_data}
            
            # 收集需要更新状态的员工（批量处理）
            visibility_updates = []
            filtered_data = []
            
            for item in data:
                employee_id = item.get("employee_id", item.get("employee_name", ""))
                # 修复：使用实时在线状态，而不是历史数据中的online
                online = online_dict.get(employee_id, True)  # 默认在线（如果获取不到实时状态）
                total_consult = item.get("total_consult", 0)
                avg_reply = item.get("avg_reply", 0)
                
                # 当前隐藏状态
                is_hidden = hidden_dict.get(employee_id, False)
                
                # 【自动隐藏条件】
                should_auto_hide = False
                hide_reason = ""
                
                # 条件1：离线员工自动隐藏
                if not online:
                    should_auto_hide = True
                    hide_reason = "离线"
                
                # 条件2：在线但无数据（咨询=0且回复=0）
                elif online and total_consult == 0 and avg_reply == 0:
                    should_auto_hide = True
                    hide_reason = "无数据"
                
                # 【自动显示条件】
                # 在线且有数据（咨询>0或回复>0）
                should_auto_show = online and (total_consult > 0 or avg_reply > 0)
                
                # 更新隐藏状态
                if should_auto_hide and not is_hidden:
                    # 需要自动隐藏
                    is_hidden = True
                    visibility_updates.append({
                        "employee_id": employee_id,
                        "hidden": True,
                        "is_manual": False  # 自动隐藏
                    })
                    print(f"[自动隐藏] {employee_id} - 原因: {hide_reason}")
                
                elif should_auto_show and is_hidden:
                    # 需要自动显示（取消隐藏）
                    is_hidden = False
                    visibility_updates.append({
                        "employee_id": employee_id,
                        "hidden": False,
                        "is_manual": False  # 自动显示
                    })
                    print(f"[自动显示] {employee_id} - 原因: 在线且有数据")
                
                # 过滤：只添加未隐藏的员工
                if not is_hidden:
                    filtered_data.append(item)
            
            # 批量保存状态更新到服务器
            if visibility_updates:
                try:
                    requests.post(
                        f"{SERVER_URL}/employee_visibility",
                        json={"visibility": visibility_updates},
                        timeout=REQUEST_TIMEOUT
                    )
                    print(f"[可见性] 批量更新 {len(visibility_updates)} 个员工状态")
                except Exception as e:
                    print(f"[可见性] 批量保存失败: {e}")
            
            print(f"[可见性] 过滤后显示 {len(filtered_data)}/{len(data)} 个员工")
            return filtered_data
        except requests.exceptions.Timeout:
            print(f"[可见性] 请求超时，返回原始数据")
            return data
        except Exception as e:
            print(f"应用可见性过滤失败: {e}")
            return data  # 如果过滤失败，返回原始数据

class NetworkWorker(QThread):
    """网络请求工作线程，避免阻塞UI"""
    data_ready = pyqtSignal(list)  # 数据准备好的信号
    error_occurred = pyqtSignal(str)  # 错误信号
    
    def __init__(self, url, timeout=REQUEST_TIMEOUT):
        super().__init__()
        self.url = url
        self.timeout = timeout
        self._is_cancelled = False
    
    def cancel(self):
        """取消当前请求"""
        self._is_cancelled = True
    
    def run(self):
        """在后台线程中执行网络请求"""
        if self._is_cancelled:
            return
        
        try:
            resp = requests.get(self.url, timeout=self.timeout)
            
            if self._is_cancelled:
                return
            
            if resp.status_code == 200:
                employees = resp.json()
                self.data_ready.emit(employees)
            else:
                self.error_occurred.emit(f"HTTP {resp.status_code}")
        except requests.exceptions.Timeout:
            if not self._is_cancelled:
                self.error_occurred.emit("请求超时")
        except requests.exceptions.ConnectionError:
            if not self._is_cancelled:
                self.error_occurred.emit("连接失败")
        except Exception as e:
            if not self._is_cancelled:
                self.error_occurred.emit(str(e))

class CustomTextEdit(QTextEdit):
    """修复3：自定义文本框，支持回车发送，CTRL+回车换行"""
    enter_pressed = pyqtSignal()  # 发送回车信号
    
    def keyPressEvent(self, event):
        # 修复3：按回车立即发送（不换行）
        if event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
            # CTRL+回车：换行
            if event.modifiers() & Qt.ControlModifier:
                super().keyPressEvent(event)
            # 单独回车：发送
            else:
                self.enter_pressed.emit()
        else:
            super().keyPressEvent(event)

class MessageInputDialog(QDialog):
    """自定义的消息输入对话框，提供大的文本输入框"""
    def __init__(self, employee_name, parent=None):
        super().__init__(parent)
        self.setWindowTitle("发送消息")
        self.resize(500, 300)
        
        layout = QVBoxLayout()
        
        # 提示标签
        label = QLabel(f"给员工 {employee_name} 发送消息:")
        label.setFont(QFont("微软雅黑", 10))
        layout.addWidget(label)
        
        # 修复3：使用自定义文本框，更新提示文字
        self.text_edit = CustomTextEdit()
        self.text_edit.setFont(QFont("微软雅黑", 11))
        self.text_edit.setPlaceholderText("按回车立即发送，按CTRL+回车换行，请输入要发送的消息")
        self.text_edit.enter_pressed.connect(self.accept)  # 回车时自动接受
        layout.addWidget(self.text_edit)
        
        # 按钮
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        self.setLayout(layout)
        
        # 焦点设置到文本框
        self.text_edit.setFocus()
    
    def get_message(self):
        """获取输入的消息内容"""
        return self.text_edit.toPlainText().strip()

class ManagerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("客服监控管理端")
        
        # 创建QSettings（保存主窗口配置到Windows注册表）
        self.settings = QSettings("QianNiuMonitor", "ManagerApp")
        
        # 恢复窗口大小和位置
        self.restore_window_geometry()
        
        font = QFont("Arial", 11)
        self.setFont(font)
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
        status_layout = QHBoxLayout()
        self.last_update_label = QLabel("最新收到数据时间：--.--.--")
        self.connection_status = QLabel("❌️ 未连接")
        self.reconnect_btn = QPushButton("重新连接")
        status_layout.addWidget(self.last_update_label)
        status_layout.addStretch()
        status_layout.addWidget(self.connection_status)
        status_layout.addWidget(self.reconnect_btn)
        main_layout.addLayout(status_layout)
        
        title_label = QLabel("实时状态")
        title_label.setFont(QFont("Arial", 11, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "员工名", "顾客数", "接待窗口", "店铺列表", 
            "今日咨询", "平均回复(秒)", "状态"
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setColumnWidth(0, 150)
        self.table.setColumnWidth(1, 80)
        self.table.setColumnWidth(2, 90)
        self.table.setColumnWidth(3, 320)
        self.table.setColumnWidth(4, 90)
        self.table.setColumnWidth(5, 120)
        self.table.setColumnWidth(6, 80)
        self.table.setWordWrap(True)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.table.setSelectionBehavior(QTableWidget.SelectItems)
        self.table.setSelectionMode(QTableWidget.NoSelection)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        main_layout.addWidget(self.table)
        
        button_layout = QHBoxLayout()
        self.rename_btn = QPushButton("启用改名")
        self.history_btn = QPushButton("历史统计查询")
        self.offline_toggle_btn = QPushButton("隐藏离线员工")
        self.delete_btn = QPushButton("删除员工")
        self.delete_btn.setStyleSheet("background-color: #ff6666; color: white;")
        self.sort_btn = QPushButton("员工排序")
        self.sort_btn.setStyleSheet("background-color: #4CAF50; color: white;")
        button_layout.addWidget(self.rename_btn)
        button_layout.addWidget(self.history_btn)
        button_layout.addWidget(self.offline_toggle_btn)
        button_layout.addWidget(self.delete_btn)
        button_layout.addWidget(self.sort_btn)
        button_layout.addStretch()
        main_layout.addLayout(button_layout)
        
        self.rename_btn.clicked.connect(self.toggle_rename)
        self.history_btn.clicked.connect(self.open_history)
        self.offline_toggle_btn.clicked.connect(self.toggle_offline_display)
        self.delete_btn.clicked.connect(self.open_delete_dialog)
        self.sort_btn.clicked.connect(self.open_sort_dialog)
        self.reconnect_btn.clicked.connect(self.reconnect_server)
        
        self.rename_mode = False
        self.show_offline = True
        self.last_update_time = 0
        self.last_request_time = 0  # 记录最后一次请求的时间戳，用于判断是否处理响应
        self.connected = False
        
        # 网络工作线程
        self.network_worker = None
        
        # 连接双击事件（始终保持连接，通过rename_mode判断行为）
        self.table.cellDoubleClicked.connect(self.on_double_click)
        
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_realtime)
        self.timer.start(500)  # 修复：改为500ms（0.5秒）刷新，提高实时性
        
        # 缓存上次数据，只在变化时更新
        self.last_employees_data = None
        
        self.update_realtime()
    
    def reconnect_server(self):
        self.update_realtime(force=True)
    
    def toggle_rename(self):
        self.rename_mode = not self.rename_mode
        if self.rename_mode:
            self.rename_btn.setText("关闭改名")
            self.rename_btn.setStyleSheet("background-color: lightgreen;")
        else:
            self.rename_btn.setText("启用改名")
            self.rename_btn.setStyleSheet("")
    
    def on_double_click(self, row, column):
        # 双击员工名（第0列）
        if column != 0:
            return
        
        # 从表格项的data中获取原始ID
        name_item = self.table.item(row, 0)
        if not name_item:
            return
        original_id = name_item.data(Qt.UserRole)  # 获取存储的原始ID
        current_display_name = name_item.text()
        
        # 如果处于改名模式，执行改名操作
        if self.rename_mode:
            new_name, ok = QInputDialog.getText(
                self, "修改员工名", "请输入新名称:", text=current_display_name
            )
            if ok and new_name.strip() and new_name.strip() != current_display_name:
                try:
                    resp = requests.post(
                        f"{SERVER_URL}/rename_employee",
                        json={"original_id": original_id, "new_name": new_name.strip()},
                        timeout=REQUEST_TIMEOUT
                    )
                    if resp.status_code == 200:
                        self.update_realtime()
                    else:
                        QMessageBox.critical(self, "错误", "改名失败，请重试")
                except Exception as ex:
                    QMessageBox.critical(self, "网络错误", f"无法连接服务器:\n{str(ex)}")
        else:
            # 非改名模式，发送消息给员工
            dialog = MessageInputDialog(current_display_name, self)
            if dialog.exec_() == QDialog.Accepted:
                message = dialog.get_message()
                if message:
                    try:
                        resp = requests.post(
                            f"{SERVER_URL}/send_message",
                            json={"employee_id": original_id, "message": message},
                            timeout=REQUEST_TIMEOUT
                        )
                        if resp.status_code == 200:
                            QMessageBox.information(self, "成功", f"消息已发送给 {current_display_name}")
                        else:
                            QMessageBox.critical(self, "错误", "消息发送失败，请重试")
                    except Exception as ex:
                        QMessageBox.critical(self, "网络错误", f"无法连接服务器:\n{str(ex)}")
    
    def open_history(self):
        dialog = HistoryDialog(self)
        dialog.show()  # 使用show()而不是exec_()，允许同时查看主窗口
    
    def toggle_offline_display(self):
        self.show_offline = not self.show_offline
        self.offline_toggle_btn.setText("隐藏离线员工" if self.show_offline else "显示离线员工")
        self.update_realtime()
    
    def open_delete_dialog(self):
        """打开删除员工对话框"""
        try:
            # 获取当前员工列表
            resp = requests.get(f"{SERVER_URL}/employees", timeout=REQUEST_TIMEOUT)
            employees = resp.json() if resp.status_code == 200 else []
            
            if not employees:
                QMessageBox.information(self, "提示", "当前没有员工记录")
                return
            
            # 打开删除对话框
            dialog = DeleteEmployeeDialog(self, employees)
            if dialog.exec_() == QDialog.Accepted:
                # 删除成功后刷新列表
                self.update_realtime()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"获取员工列表失败：\n{str(e)}")
    
    def open_sort_dialog(self):
        """打开员工排序对话框"""
        try:
            # 获取当前员工列表
            resp = requests.get(f"{SERVER_URL}/employees", timeout=REQUEST_TIMEOUT)
            employees = resp.json() if resp.status_code == 200 else []
            
            if not employees:
                QMessageBox.information(self, "提示", "当前没有员工记录")
                return
            
            # 创建排序对话框
            dialog = QDialog(self)
            dialog.setWindowTitle("员工排序设置")
            dialog.resize(600, 500)
            
            layout = QVBoxLayout()
            
            # 说明
            info_label = QLabel("为每个员工设置排序序号（数字越小，排序越靠前）\n离线员工自动排在在线员工下面")
            info_label.setFont(QFont("Arial", 10))
            layout.addWidget(info_label)
            
            # 获取现有排序配置
            try:
                resp_order = requests.get(f"{SERVER_URL}/employee_order", timeout=REQUEST_TIMEOUT)
                order_data = resp_order.json() if resp_order.status_code == 200 else []
                order_dict = {item['employee_id']: item['order'] for item in order_data}
            except:
                order_dict = {}
            
            # 创建表格
            table = QTableWidget()
            table.setColumnCount(4)
            table.setHorizontalHeaderLabels(["员工名称", "状态", "排序序号", "员工ID"])
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            table.setRowCount(len(employees))
            
            for row, emp in enumerate(employees):
                display_name = emp.get("display_name", emp.get("employee_name", ""))
                employee_id = emp.get("employee_name", "")
                online = emp.get("online", False)
                status = "🟢在线" if online else "🔴离线"
                current_order = order_dict.get(employee_id, 9999)
                
                # 员工名称
                name_item = QTableWidgetItem(display_name)
                name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
                table.setItem(row, 0, name_item)
                
                # 状态
                status_item = QTableWidgetItem(status)
                status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
                status_item.setTextAlignment(Qt.AlignCenter)
                table.setItem(row, 1, status_item)
                
                # 排序序号（可编辑）
                order_spinbox = QSpinBox()
                order_spinbox.setRange(1, 9999)
                order_spinbox.setValue(current_order)
                order_spinbox.setAlignment(Qt.AlignCenter)
                table.setCellWidget(row, 2, order_spinbox)
                
                # 员工ID
                id_item = QTableWidgetItem(employee_id)
                id_item.setFlags(id_item.flags() & ~Qt.ItemIsEditable)
                table.setItem(row, 3, id_item)
            
            layout.addWidget(table)
            
            # 按钮
            button_layout = QHBoxLayout()
            save_btn = QPushButton("保存")
            save_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
            cancel_btn = QPushButton("取消")
            button_layout.addStretch()
            button_layout.addWidget(save_btn)
            button_layout.addWidget(cancel_btn)
            layout.addLayout(button_layout)
            
            dialog.setLayout(layout)
            
            def save_order():
                try:
                    orders = []
                    for row in range(table.rowCount()):
                        employee_id = table.item(row, 3).text()
                        spinbox = table.cellWidget(row, 2)
                        order = spinbox.value()
                        orders.append({
                            "employee_id": employee_id,
                            "order": order
                        })
                    
                    # 保存到服务器
                    resp = requests.post(
                        f"{SERVER_URL}/save_employee_order",
                        json={"orders": orders},
                        timeout=REQUEST_TIMEOUT
                    )
                    
                    if resp.status_code == 200:
                        QMessageBox.information(dialog, "成功", "排序配置已保存")
                        dialog.accept()
                        # 刷新表格
                        self.update_realtime(force=True)
                    else:
                        QMessageBox.warning(dialog, "失败", "保存排序配置失败")
                except Exception as e:
                    QMessageBox.critical(dialog, "错误", f"保存时发生错误：\n{str(e)}")
            
            save_btn.clicked.connect(save_order)
            cancel_btn.clicked.connect(dialog.reject)
            
            dialog.exec_()
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"打开排序对话框失败：\n{str(e)}")
    
    def apply_employee_sort(self, employees):
        """应用员工排序（离线员工排在在线员工下面）"""
        try:
            # 获取排序配置
            resp = requests.get(f"{SERVER_URL}/employee_order", timeout=REQUEST_TIMEOUT)
            order_data = resp.json() if resp.status_code == 200 else []
            order_dict = {item['employee_id']: item['order'] for item in order_data}
            
            # 分离在线和离线员工
            online_employees = [emp for emp in employees if emp.get("online", False)]
            offline_employees = [emp for emp in employees if not emp.get("online", False)]
            
            # 分别排序
            def get_sort_key(emp):
                employee_id = emp.get("employee_name", "")
                return order_dict.get(employee_id, 9999)
            
            online_sorted = sorted(online_employees, key=get_sort_key)
            offline_sorted = sorted(offline_employees, key=get_sort_key)
            
            # 在线员工在前，离线员工在后
            return online_sorted + offline_sorted
        except Exception as e:
            print(f"应用员工排序失败: {e}")
            return employees  # 失败时返回原始列表
    
    def update_realtime(self, force=False):
        """非阻塞的实时数据更新"""
        # 修复：如果有正在运行的任务，取消旧的请求，启动新的请求
        # 这样可以确保总是获取最新的数据，避免延迟累积
        if self.network_worker and self.network_worker.isRunning():
            print("[实时更新] 检测到前一个请求还在运行，取消旧请求，启动新请求")
            self.network_worker.cancel()  # 取消旧的请求
            self.network_worker.quit()  # 退出线程
            self.network_worker.wait(50)  # 等待线程退出（最多等待50ms，不阻塞太久）
            # 如果50ms内没有退出，强制终止（但这样可能会导致资源泄漏，所以避免）
            if self.network_worker.isRunning():
                print("[实时更新] 警告：前一个请求仍在运行，将直接启动新请求（旧请求会被忽略）")
        
        # 创建新的工作线程（使用请求时间戳来标识最新的请求）
        self.network_worker = NetworkWorker(f"{SERVER_URL}/employees")
        # 使用lambda捕获当前时间戳，确保只处理最新的响应
        # 注意：使用默认参数来避免闭包问题
        request_time = time.time()
        # 先更新请求时间（确保后续响应能正确判断）
        self.last_request_time = request_time
        self.network_worker.data_ready.connect(
            lambda employees, rt=request_time, f=force: self.on_data_received(employees, f, rt)
        )
        self.network_worker.error_occurred.connect(self.on_network_error)
        self.network_worker.start()
    
    def on_data_received(self, employees, force=False, request_time=None):
        """数据接收完成（在主线程中执行UI更新）"""
        try:
            # 修复：如果这不是最新的请求响应，忽略它（避免处理过期的数据）
            if request_time is not None:
                if request_time < self.last_request_time:
                    print(f"[实时更新] 忽略过期的响应（请求时间：{request_time:.3f}，最新请求时间：{self.last_request_time:.3f}）")
                    return
                # 这是最新的响应，更新请求时间
                self.last_request_time = request_time
            
            # 数据对比，只在变化时更新UI
            employees_str = json.dumps(employees, sort_keys=True)
            if not force and employees_str == self.last_employees_data:
                return  # 数据未变化，不更新UI
            
            self.last_employees_data = employees_str
            self.connected = True
            self.connection_status.setText("✅️ 已连接")
            self.last_update_time = time.time()
            # 更新最新收到数据时间显示
            current_time = datetime.now()
            time_str = current_time.strftime("%H.%M.%S")
            self.last_update_label.setText(f"最新收到数据时间：{time_str}")
            
            filtered_employees = [
                emp for emp in employees
                if self.show_offline or emp["online"]
            ]
            
            # 应用员工排序（离线员工排在在线员工下面）
            filtered_employees = self.apply_employee_sort(filtered_employees)
            
            self.table.setRowCount(len(filtered_employees))
            
            for row, emp in enumerate(filtered_employees):
                name = emp["display_name"]
                original_id = emp["employee_name"]  # 获取原始ID
                customers = str(emp["total_customers"]) if emp["online"] else ""
                shops = str(emp["total_shops"]) if emp["online"] else ""
                
                # 排序店铺列表：未知店铺排在已知店铺后面
                raw_shops = emp.get("shops_list", [])
                if raw_shops:
                    known_shops = [s for s in raw_shops if not s.startswith("未知店铺")]
                    unknown_shops = [s for s in raw_shops if s.startswith("未知店铺")]
                    sorted_shops = known_shops + unknown_shops
                    shop_list = "\n".join(sorted_shops)
                else:
                    shop_list = "无店铺" if emp["online"] else ""
                
                consult = str(emp["today_consult"])
                avg_reply = str(emp["avg_reply"])  # 已是整数
                status = "🟢在线" if emp["online"] else "🔴离线"
                
                # 设置员工名称，并存储原始ID
                name_item = QTableWidgetItem(name)
                name_item.setTextAlignment(Qt.AlignCenter)
                name_item.setData(Qt.UserRole, original_id)  # 存储原始ID到UserRole
                self.table.setItem(row, 0, name_item)
                self._set_table_item(row, 1, customers, align=Qt.AlignCenter)
                self._set_table_item(row, 2, shops, align=Qt.AlignCenter)
                self._set_table_item(row, 3, shop_list, align=Qt.AlignLeft | Qt.AlignTop)
                self._set_table_item(row, 4, consult, align=Qt.AlignCenter)
                self._set_table_item(row, 5, avg_reply, align=Qt.AlignCenter)
                self._set_table_item(row, 6, status, align=Qt.AlignCenter)
                self.table.resizeRowToContents(row)
                
                if emp["online"]:
                    cust_val = emp["total_customers"]
                    color = QColor("black") if cust_val <= 3 else QColor("blue") if cust_val <= 5 else QColor("red")
                    self.table.item(row, 1).setForeground(color)
                    shop_val = emp["total_shops"]
                    color = QColor("black") if shop_val <= 3 else QColor("blue") if shop_val <= 5 else QColor("red")
                    self.table.item(row, 2).setForeground(color)
                else:
                    self.table.item(row, 1).setForeground(QColor("black"))
                    self.table.item(row, 2).setForeground(QColor("black"))
        except Exception as e:
            print(f"[UI更新错误] {e}")
            self.connected = False
            self.connection_status.setText("❌️ 数据处理失败")
    
    def on_network_error(self, error_msg):
        """网络错误处理（在主线程中执行）"""
        self.connected = False
        self.connection_status.setText(f"❌️ {error_msg}")

    def _set_table_item(self, row, col, text, align=Qt.AlignCenter):
        item = QTableWidgetItem(str(text) if text is not None else "")
        item.setTextAlignment(align)
        self.table.setItem(row, col, item)
    
    def restore_window_geometry(self):
        """恢复主窗口大小和位置（从Windows注册表读取）"""
        # 恢复窗口大小
        width = self.settings.value("window/width", 1100, type=int)
        height = self.settings.value("window/height", 700, type=int)
        self.resize(width, height)
        
        # 恢复窗口位置
        x = self.settings.value("window/x", None, type=int)
        y = self.settings.value("window/y", None, type=int)
        
        if x is not None and y is not None:
            # 检查位置是否在屏幕范围内
            screen = QApplication.desktop().screenGeometry()
            if 0 <= x < screen.width() - 100 and 0 <= y < screen.height() - 100:
                self.move(x, y)
        
        print(f"[主窗口配置] 恢复窗口大小: {width}x{height}, 位置: ({x}, {y})")
    
    def save_window_geometry(self):
        """保存主窗口大小和位置（保存到Windows注册表）"""
        # 保存窗口大小
        self.settings.setValue("window/width", self.width())
        self.settings.setValue("window/height", self.height())
        
        # 保存窗口位置
        self.settings.setValue("window/x", self.x())
        self.settings.setValue("window/y", self.y())
        
        print(f"[主窗口配置] 保存窗口大小: {self.width()}x{self.height()}, 位置: ({self.x()}, {self.y()})")
    
    def closeEvent(self, event):
        """主窗口关闭时保存配置"""
        self.save_window_geometry()
        event.accept()

if __name__ == "__main__":
    # 不隐藏控制台窗口，方便查看错误信息
    try:
        # 设置共享OpenGL上下文（WebEngine需要）
        QApplication.setAttribute(Qt.AA_ShareOpenGLContexts)
        
        app = QApplication(sys.argv)
        window = ManagerApp()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        # 如果出错，显示错误信息
        import traceback
        print("\n" + "=" * 60)
        print("程序启动失败！")
        print("=" * 60)
        print(f"\n错误类型: {type(e).__name__}")
        print(f"错误信息: {str(e)}")
        print("\n详细错误堆栈:")
        traceback.print_exc()
        input("\n按 Enter 键退出...")
