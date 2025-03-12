# TG-DataExporter (Telegram 数据自动导出工具)

## 简介

TG-DataExporter 是一个自动化工具，用于批量导出 Telegram 客户端的聊天数据。该工具通过图像识别和模拟点击操作，实现了无需人工干预的自动化导出流程，特别适合需要批量处理多个 Telegram 账号数据的场景。

---

## 功能特点

- **多客户端支持**：批量处理多个 Telegram 客户端。
- **多语言适配**：自动检测客户端界面语言（支持英语和俄语）。
- **数据导出**：自动导出聊天记录、媒体文件和其他数据。
- **文件整理**：自动将导出文件分类存储到指定目录。
- **日志与报告**：详细的日志记录、错误处理，以及失败客户端的记录和报告。

---

## 安装要求

1. **Python 环境**：需 Python 3.7 或更高版本。
2. **依赖安装**：

   pip install -r requirements.txt

---

## 使用方法

### 1. 准备截图文件
- 为每种支持的语言（英语、俄语）准备界面元素截图。
- 截图保存在 screenshots/{语言代码}/ 目录下。
- **必要截图文件**：
  - hamburger_menu.png（左上角三横线菜单按钮）
  - settings_menu_item.png（设置选项）
  - advanced_tab.png（高级选项卡）
  - export_button.png（导出数据按钮）
  - 其他详见 [截图准备指南](#截图准备指南)。

### 2. 运行程序

python TG-DataExporter.py

### 3. 输入参数
- **客户端根目录**：包含多个 Telegram 客户端的目录。
- **导出目标目录**：数据导出的目标路径。

### 4. 自动化流程
程序将依次执行以下操作：
1. 扫描并识别所有客户端。
2. 逐个启动客户端并执行导出操作。
3. 整理导出的数据至目标目录。
4. 生成处理报告（report.txt 和 failed_exports.txt）。

---

## 截图准备指南

程序需要以下截图文件（每种语言均需独立配置）：

| 截图文件名                | 描述                          |
|---------------------------|-------------------------------|
| hamburger_menu.png      | 左上角三横线菜单按钮（白底样式）|
| hamburger_menu_dark.png | 左上角三横线菜单按钮（黑底样式，可选）|
| settings_menu_item.png  | 菜单中的设置选项              |
| advanced_tab.png        | 设置页面的高级选项卡          |
| export_button.png       | 高级设置中的导出数据按钮      |
| save_button.png         | 导出设置页面的保存按钮        |
| export_settings_title.png | 导出设置窗口的标题          |
| show_my_data_button.png | 导出完成后出现的按钮          |
| close_button.png        | 导出完成窗口的关闭按钮        |
| start_messaging_button.png | 未登录状态下的按钮         |

**导出选项文字截图**：
以下是需要勾选的导出选项的文字截图，每个选项都需要单独截图：

| 截图文件名                | 描述                          |
|---------------------------|-------------------------------|
| option_only_my_messages.png | "只导出我的消息"选项文字    |
| option_videos.png        | "视频"选项文字                |
| option_voice_messages.png | "语音消息"选项文字           |
| option_video_messages.png | "视频消息"选项文字           |
| option_stickers.png      | "贴纸"选项文字                |
| option_gifs.png          | "GIF"选项文字                 |
| option_files.png         | "文件"选项文字                |
| option_both.png          | "HTML和JSON格式"选项文字      |

**截图要求**：
- 仅包含目标元素，避免冗余内容。
- 使用与运行环境相同的屏幕分辨率。
- 按目录结构保存（如 screenshots/en/、screenshots/ru/）。
- 截图时确保只包含需要识别的元素，不要包含太多周围内容。
- 由于不同电脑分辨率不同，请在您自己的电脑上截取上述元素的图片。

---

## 注意事项

- 🚫 程序运行时请勿操作鼠标和键盘。
- ⏳ 导出耗时较长，请耐心等待。
- 🔍 失败记录可通过 failed_exports.txt 查看。

---

## 故障排除

若遇到识别问题，请尝试：
1. 重新截取更清晰的界面元素图片。
2. 调整代码中的 confidence 参数（默认值 0.6-0.8）。
3. 查看日志文件 telegram_export.log 获取详细错误信息。

---

**提示**：建议在首次运行前，先手动完成一次导出流程以熟悉界面布局。