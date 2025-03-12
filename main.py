"""
主程序入口文件 - 启动Telegram数据导出工具
"""

import sys
from PyQt5.QtWidgets import QApplication
from gui import TelegramExporterGUI
import TG_DataExporter

def main():
    """程序主入口函数"""
    # 创建QApplication实例
    app = QApplication(sys.argv)
    
    # 创建主窗口
    window = TelegramExporterGUI()
    
    # 显示窗口
    window.show()
    
    # 进入应用程序主循环
    sys.exit(app.exec_())

if __name__ == "__main__":
    # 移除COM初始化和清理代码
    try:
        # 启动GUI界面
        main()
    finally:
        pass