"""
样式表定义文件 - 包含应用程序的所有样式
"""

# 主应用样式
MAIN_STYLE = """
QWidget {
    font-family: 'Microsoft YaHei', 'SimHei', sans-serif;
    font-size: 11pt;  /* 增加默认字体大小 */
}

QGroupBox {
    font-size: 12pt;  /* 增加分组框标题字体大小 */
    font-weight: bold;
    border: 1px solid #CCCCCC;
    border-radius: 5px;
    margin-top: 10px;
    padding-top: 15px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top center;
    padding: 0 5px;
}

QPushButton {
    font-size: 11pt;  /* 增加按钮字体大小 */
    padding: 8px 15px;  /* 增加按钮内边距 */
    background-color: #4A86E8;
    color: white;
    border: none;
    border-radius: 4px;
}

QPushButton:hover {
    background-color: #3D7DD6;
}

QPushButton:pressed {
    background-color: #2D6BC4;
}

QPushButton:disabled {
    background-color: #CCCCCC;
    color: #888888;
}

QLineEdit {
    font-size: 11pt;  /* 增加输入框字体大小 */
    padding: 8px;  /* 增加输入框内边距 */
    border: 1px solid #CCCCCC;
    border-radius: 4px;
}

QComboBox {
    font-size: 11pt;  /* 增加下拉框字体大小 */
    padding: 8px;  /* 增加下拉框内边距 */
    border: 1px solid #CCCCCC;
    border-radius: 4px;
}

QLabel {
    font-size: 11pt;  /* 增加标签字体大小 */
}

#titleLabel {
    font-size: 16pt;  /* 增加标题字体大小 */
    font-weight: bold;
    color: #2D6BC4;
    margin: 10px 0;
}

#subtitleLabel {
    font-size: 12pt;  /* 增加副标题字体大小 */
    color: #555555;
    margin-bottom: 15px;
}

#statusIndicator {
    font-size: 12pt;  /* 增加状态指示器字体大小 */
    font-weight: bold;
    padding: 8px;
    border-radius: 4px;
}
"""

# 日志样式
LOG_STYLE = """
QTextEdit {
    font-family: 'Consolas', 'Courier New', monospace;
    font-size: 11pt;  /* 增加日志字体大小 */
    background-color: #F8F8F8;
    border: 1px solid #CCCCCC;
    border-radius: 4px;
    padding: 5px;
}
"""

# 状态指示器样式
STATUS_INDICATOR_STYLE = """
QLabel#statusIndicator {
    font-weight: bold;
    padding: 5px;
    border-radius: 4px;
}
"""

# 成功状态样式
SUCCESS_STYLE = """
QLabel#statusIndicator {
    background-color: #2ecc71;
    color: white;
}
"""

# 失败状态样式
FAILURE_STYLE = """
QLabel#statusIndicator {
    background-color: #e74c3c;
    color: white;
}
"""

# 进行中状态样式
PROCESSING_STYLE = """
QLabel#statusIndicator {
    background-color: #f39c12;
    color: white;
}
"""