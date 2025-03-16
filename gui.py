"""
图形界面实现文件 - 包含所有UI组件和布局
"""

import os
import sys
import time
import threading
import logging
import base64  # 添加base64模块导入
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QLineEdit, QPushButton, 
                            QTextEdit, QFileDialog, QProgressBar, QGroupBox,
                            QComboBox, QMessageBox, QFrame, QSplitter, QStatusBar,
                            QScrollArea)  # 添加QScrollArea
from PyQt5.QtCore import Qt, pyqtSignal, QObject, QSize
from PyQt5.QtGui import QIcon, QPixmap, QFont

# 导入样式
from style import (MAIN_STYLE, LOG_STYLE, STATUS_INDICATOR_STYLE, 
                  SUCCESS_STYLE, FAILURE_STYLE, PROCESSING_STYLE)

# 导入资源
from resources import LOGO_BASE64  # 导入logo数据

# 导入导出功能
import TG_DataExporter as exporter

# 自定义信号类，用于线程间通信
class ExporterSignals(QObject):
    update_log = pyqtSignal(str)
    update_progress = pyqtSignal(int, int)  # 当前进度, 总数
    update_status = pyqtSignal(str, str)  # 状态, 样式
    export_finished = pyqtSignal(list)  # 失败的客户端列表
    client_processed = pyqtSignal(str, bool)  # 客户端名称, 是否成功

class TelegramExporterGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.signals = ExporterSignals()
        self.init_ui()
        self.setup_logging()
        self.connect_signals()
        
    def init_ui(self):
        """初始化UI组件"""
        self.setWindowTitle("Telegram数据自动导出工具")
        self.setMinimumSize(900, 1200)  # 增加默认最小尺寸
        
        # 设置应用程序图标
        if LOGO_BASE64.strip():
            try:
                # 解码base64数据
                logo_data = base64.b64decode(LOGO_BASE64)
                # 创建QPixmap
                pixmap = QPixmap()
                pixmap.loadFromData(logo_data)
                # 设置窗口图标
                self.setWindowIcon(QIcon(pixmap))
            except Exception as e:
                print(f"加载logo时出错: {str(e)}")
        
        # 设置应用程序默认字体
        app = QApplication.instance()
        font = app.font()
        font.setPointSize(11)  # 增加默认字体大小
        app.setFont(font)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # 标题
        title_label = QLabel("Telegram数据自动导出工具")
        title_font = QFont()
        title_font.setPointSize(16)  # 设置标题字体大小
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setObjectName("titleLabel")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 副标题
        subtitle_label = QLabel("通过图像识别实现自动化导出Telegram聊天记录")
        subtitle_font = QFont()
        subtitle_font.setPointSize(12)  # 设置副标题字体大小
        subtitle_label.setFont(subtitle_font)
        subtitle_label.setObjectName("subtitleLabel")
        subtitle_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(subtitle_label)
        
        # 分隔线
        line = QFrame()
        line.setObjectName("line")
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line)
        
        # 配置区域
        config_group = QGroupBox("配置")
        config_layout = QVBoxLayout(config_group)
        
        # 创建滚动区域用于客户端路径
        client_paths_scroll = QScrollArea()
        client_paths_scroll.setWidgetResizable(True)
        client_paths_scroll.setMinimumHeight(150)  # 设置最小高度
        client_paths_widget = QWidget()
        client_paths_scroll.setWidget(client_paths_widget)
        
        # 客户端目录选择 - 改为动态添加的容器
        self.client_paths_container = QVBoxLayout(client_paths_widget)
        self.client_paths_container.setContentsMargins(5, 5, 5, 5)
        self.client_paths_container.setSpacing(10)
        
        # 添加第一个客户端目录选择行
        self.add_client_path_row()
        
        # 添加"添加路径"按钮
        add_path_layout = QHBoxLayout()
        add_path_layout.addStretch(1)
        add_path_btn = QPushButton("添加客户端路径")
        add_path_btn.clicked.connect(self.add_client_path_row)
        add_path_layout.addWidget(add_path_btn)
        
        # 将滚动区域和添加按钮添加到配置布局
        config_layout.addWidget(client_paths_scroll)
        config_layout.addLayout(add_path_layout)
        
        # 导出目录选择
        export_layout = QHBoxLayout()
        export_label = QLabel("导出目录:")
        self.export_path_edit = QLineEdit()
        self.export_path_edit.setPlaceholderText("选择导出数据的保存目录")
        export_browse_btn = QPushButton("浏览...")
        export_browse_btn.clicked.connect(self.browse_export_dir)
        export_layout.addWidget(export_label)
        export_layout.addWidget(self.export_path_edit, 1)
        export_layout.addWidget(export_browse_btn)
        config_layout.addLayout(export_layout)
        
        # 移除语言选择部分
        # language_layout = QHBoxLayout()
        # language_label = QLabel("界面语言:")
        # self.language_combo = QComboBox()
        # for lang in exporter.SUPPORTED_LANGUAGES:
        #     self.language_combo.addItem(f"{lang} - {'英语' if lang == 'en' else '俄语'}", lang)
        # language_layout.addWidget(language_label)
        # language_layout.addWidget(self.language_combo)
        # language_layout.addStretch(1)
        # config_layout.addLayout(language_layout)
        
        main_layout.addWidget(config_group)
        
        # 操作区域
        action_layout = QHBoxLayout()
        self.start_btn = QPushButton("开始导出")
        self.start_btn.clicked.connect(self.start_export)
        self.stop_btn = QPushButton("停止")
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_export)
        action_layout.addStretch(1)
        action_layout.addWidget(self.start_btn)
        action_layout.addWidget(self.stop_btn)
        action_layout.addStretch(1)
        main_layout.addLayout(action_layout)
        
        # 进度条
        progress_layout = QHBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%v/%m (%p%)")
        progress_layout.addWidget(self.progress_bar)
        main_layout.addLayout(progress_layout)
        
        # 状态指示器
        status_layout = QHBoxLayout()
        self.status_indicator = QLabel("就绪")
        self.status_indicator.setObjectName("statusIndicator")
        self.status_indicator.setAlignment(Qt.AlignCenter)
        self.status_indicator.setStyleSheet(STATUS_INDICATOR_STYLE)
        status_layout.addWidget(self.status_indicator)
        main_layout.addLayout(status_layout)
        
        # 日志区域
        log_group = QGroupBox("运行日志")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setObjectName("logTextEdit")
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        main_layout.addWidget(log_group)
        
        # 设置状态栏
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("准备就绪")
        
        # 应用样式
        self.setStyleSheet(MAIN_STYLE)
        self.log_text.setStyleSheet(LOG_STYLE)
        
        # 初始化变量
        self.export_thread = None
        self.stop_requested = False
        
    def setup_logging(self):
        """设置日志处理"""
        # 创建自定义处理器，将日志输出到GUI
        class QTextEditHandler(logging.Handler):
            def __init__(self, signals):
                super().__init__()
                self.signals = signals
                
            def emit(self, record):
                msg = self.format(record)
                self.signals.update_log.emit(msg)
        
        # 配置根日志记录器
        root_logger = logging.getLogger()
        
        # 清除现有处理器
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)
        
        # 添加文件处理器
        file_handler = logging.FileHandler('telegram_export.log')
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s: %(message)s'))
        root_logger.addHandler(file_handler)
        
        # 添加GUI处理器
        gui_handler = QTextEditHandler(self.signals)
        gui_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s: %(message)s'))
        root_logger.addHandler(gui_handler)
        
        # 设置日志级别
        root_logger.setLevel(logging.INFO)
    
    def connect_signals(self):
        """连接信号和槽"""
        self.signals.update_log.connect(self.update_log)
        self.signals.update_progress.connect(self.update_progress)
        self.signals.update_status.connect(self.update_status)
        self.signals.export_finished.connect(self.export_finished)
        self.signals.client_processed.connect(self.client_processed)
    
    def add_client_path_row(self):
        """添加一个新的客户端路径输入行"""
        row_layout = QHBoxLayout()
        
        # 如果是第一行，添加标签
        if self.client_paths_container.count() == 0:
            client_label = QLabel("客户端根目录:")
            row_layout.addWidget(client_label)
        else:
            # 如果不是第一行，添加空白标签保持对齐
            empty_label = QLabel("")
            row_layout.addWidget(empty_label)
        
        # 路径输入框
        path_edit = QLineEdit()
        path_edit.setPlaceholderText(f"选择包含Telegram客户端的根目录 #{self.client_paths_container.count()+1}")
        
        # 浏览按钮
        browse_btn = QPushButton("浏览...")
        browse_btn.clicked.connect(lambda: self.browse_client_dir_for_row(path_edit))
        
        # 删除按钮 (第一行不显示删除按钮)
        if self.client_paths_container.count() > 0:
            delete_btn = QPushButton("删除")
            delete_btn.clicked.connect(lambda: self.delete_client_path_row(row_layout))
            row_layout.addWidget(path_edit, 1)
            row_layout.addWidget(browse_btn)
            row_layout.addWidget(delete_btn)
        else:
            row_layout.addWidget(path_edit, 1)
            row_layout.addWidget(browse_btn)
        
        # 将行添加到容器
        self.client_paths_container.addLayout(row_layout)
    
    def browse_client_dir_for_row(self, line_edit):
        """为特定行浏览选择客户端目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择客户端根目录")
        if dir_path:
            line_edit.setText(dir_path)
    
    def delete_client_path_row(self, row_layout):
        """删除一个客户端路径输入行"""
        # 移除行中的所有控件
        while row_layout.count():
            item = row_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()
        
        # 从容器中移除行布局
        self.client_paths_container.removeItem(row_layout)
    
    def browse_client_dir(self):
        """浏览选择客户端目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择客户端根目录")
        if dir_path:
            self.client_path_edit.setText(dir_path)
    
    def browse_export_dir(self):
        """浏览选择导出目录"""
        dir_path = QFileDialog.getExistingDirectory(self, "选择导出目录")
        if dir_path:
            self.export_path_edit.setText(dir_path)
    
    def start_export(self):
        """开始导出过程"""
        # 收集所有客户端目录
        client_dirs = []
        for i in range(self.client_paths_container.count()):
            layout = self.client_paths_container.itemAt(i)
            for j in range(layout.count()):
                item = layout.itemAt(j)
                widget = item.widget()
                if isinstance(widget, QLineEdit):
                    path = widget.text().strip()
                    if path:
                        client_dirs.append(path)
        
        export_dir = self.export_path_edit.text().strip()
        
        if not client_dirs or not export_dir:
            QMessageBox.warning(self, "输入错误", "请填写至少一个客户端根目录和导出目录")
            return
        
        # 检查截图目录
        try:
            # 检查所有支持的语言截图目录
            for language in exporter.SUPPORTED_LANGUAGES:
                exporter.check_screenshot_dir(language)
        except Exception as e:
            QMessageBox.critical(self, "截图文件错误", f"截图文件检查失败: {str(e)}")
            return
        
        # 创建导出目录
        os.makedirs(export_dir, exist_ok=True)
        
        # 禁用开始按钮，启用停止按钮
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        
        # 更新状态
        self.update_status("正在导出...", PROCESSING_STYLE)
        
        # 重置停止标志
        self.stop_requested = False
        
        # 启动导出线程
        self.export_thread = threading.Thread(
            target=self.run_export,
            args=(client_dirs, export_dir)  # 传递目录列表
        )
        self.export_thread.daemon = True
        self.export_thread.start()
    
    def stop_export(self):
        """停止导出过程"""
        if self.export_thread and self.export_thread.is_alive():
            self.stop_requested = True
            self.update_status("正在停止...", PROCESSING_STYLE)
            self.stop_btn.setEnabled(False)
    
    def progress_callback(self, message):
        """处理进度回调，同时更新日志和进度条"""
        self.signals.update_log.emit(message)
        
        # 检查消息是否包含进度信息
        if "处理进度：" in message:
            try:
                # 从消息中提取进度信息
                progress_info = message.split("处理进度：")[1].split("/")[0]
                total_info = message.split("/")[1].split("\n")[0]
                current = int(progress_info.strip())
                total = int(total_info.strip())
                
                # 更新进度条
                self.signals.update_progress.emit(current, total)
            except Exception as e:
                logging.debug(f"解析进度信息失败: {str(e)}")

    def run_export(self, client_dirs, export_dir):
        """在后台线程中运行导出过程"""
        try:
            # 使用exporter的run_export函数处理多个目录
            result = exporter.run_export(
                client_dirs, 
                export_dir,
                callback=self.progress_callback  # 使用专门的进度回调函数
            )
            
            if "error" in result:
                self.signals.update_log.emit(f"导出过程发生错误: {result['error']}")
                self.signals.update_status.emit("导出失败", FAILURE_STYLE)
                self.signals.export_finished.emit([])
                return
            
            # 处理完成
            self.signals.export_finished.emit(result.get("failed_list", []))
            
        except Exception as e:
            self.signals.update_log.emit(f"导出过程发生错误: {str(e)}")
            self.signals.update_status.emit("导出失败", FAILURE_STYLE)
            self.signals.export_finished.emit([])
    
    def update_log(self, message):
        """更新日志显示"""
        self.log_text.append(message)
        # 自动滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def update_progress(self, current, total):
        """更新进度条"""
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
    
    def update_status(self, status, style):
        """更新状态指示器"""
        self.status_indicator.setText(status)
        self.status_indicator.setStyleSheet(style)
        self.statusBar.showMessage(status)
    
    def client_processed(self, client_name, success):
        """处理单个客户端完成的回调"""
        if success:
            self.statusBar.showMessage(f"成功导出: {client_name}")
        else:
            self.statusBar.showMessage(f"导出失败: {client_name}")
    
    def export_finished(self, failed_clients):
        """导出完成的回调"""
        # 启用开始按钮，禁用停止按钮
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        
        # 更新状态
        if self.stop_requested:
            self.update_status("已停止", FAILURE_STYLE)
            self.signals.update_log.emit("导出过程已被用户中断")
        elif not failed_clients:
            self.update_status("导出完成", SUCCESS_STYLE)
            self.signals.update_log.emit("\n所有客户端导出成功！")
        else:
            self.update_status("部分导出成功", PROCESSING_STYLE)
            self.signals.update_log.emit(f"\n有 {len(failed_clients)} 个客户端导出失败")
            
            # 显示失败客户端列表
            self.signals.update_log.emit("\n失败的客户端:")
            for client_path in failed_clients:
                self.signals.update_log.emit(f"- {client_path}")
            
            # 保存失败记录到文件
            try:
                current_dir = os.getcwd()
                failed_log_path = os.path.join(current_dir, "failed_exports.txt")
                with open(failed_log_path, "w", encoding="utf-8") as f:
                    f.write(f"导出失败的客户端列表 (总计 {len(failed_clients)} 个):\n")
                    f.write("="*50 + "\n")
                    for client_path in failed_clients:
                        f.write(f"{client_path}\n")
                self.signals.update_log.emit(f"\n失败记录已保存至: {failed_log_path}")
            except Exception as e:
                self.signals.update_log.emit(f"保存失败记录文件时出错: {str(e)}")
        
        # 重置停止标志
        self.stop_requested = False


def main():
    """主函数"""
    app = QApplication(sys.argv)
    window = TelegramExporterGUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()