import os
import time
import logging
import pyautogui
import shutil
from pywinauto import Application

# 配置参数
SCREENSHOT_DIR = "screenshots"  # 界面元素截图目录
SCROLL_ATTEMPTS = 5             # 最大滚动尝试次数
SCROLL_DISTANCE = -1200         # 加大滚动距离
EXPORT_OPTIONS = [              # 需要勾选的导出选项
    "option_only_my_messages",  # 添加回来，确保所有需要的选项都在这里
    "option_videos",
    "option_voice_messages",
    "option_video_messages",    # 添加这个选项
    "option_stickers",
    "option_gifs",
    "option_files",
    "option_both"
]

# 添加语言支持
SUPPORTED_LANGUAGES = ["en", "ru"]  # 支持的语言列表：英语和俄语
DEFAULT_LANGUAGE = "en"             # 默认语言：英语

# 初始化日志
logging.basicConfig(
    filename='telegram_export.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s: %(message)s'
)

def check_screenshot_dir(language="en"):
    """检查截图目录结构，支持多语言"""
    lang_dir = os.path.join(SCREENSHOT_DIR, language)
    
    if not os.path.exists(lang_dir):
        raise FileNotFoundError(f"找不到语言目录：{lang_dir}")
    
    required_files = [
        'hamburger_menu.png',
        'settings_menu_item.png',
        'advanced_tab.png',
        'export_button.png',
        'save_button.png',
        # 移除未使用的复选框图像
        # 'checkbox_selected.png',
        # 'checkbox_unselected.png',
        'show_my_data_button.png',
        'close_button.png',
        'start_messaging_button.png',
        'export_settings_title.png'  # 添加这个，因为它在select_export_options中被使用
    ] + [f"{opt}.png" for opt in EXPORT_OPTIONS]
    
    missing = []
    for f in required_files:
        if not os.path.exists(os.path.join(lang_dir, f)):
            missing.append(f)
    
    if missing:
        raise FileNotFoundError(f"语言 {language} 缺少必要截图文件：{', '.join(missing)}")

def find_and_click(image_path, timeout=15, confidence=0.6, language="en"):
    """通过图像识别定位并点击元素，支持多语言"""
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            location = pyautogui.locateOnScreen(
                os.path.join(SCREENSHOT_DIR, language, image_path),
                confidence=confidence
            )
            if location:
                center = pyautogui.center(location)
                pyautogui.click(center)
                return True
        except pyautogui.ImageNotFoundException:
            time.sleep(1)
    logging.warning(f"未找到元素：{image_path} (语言: {language})")
    return False

def select_export_options(language="en"):
    """选择导出选项（直接点击所有指定选项），支持多语言"""
    # 通过识别导出设置窗口标题来获取焦点
    try:
        # 尝试查找"Chat export settings"标题
        title_loc = pyautogui.locateOnScreen(
            os.path.join(SCREENSHOT_DIR, language, "export_settings_title.png"),
            confidence=0.7
        )
        if title_loc:
            center = pyautogui.center(title_loc)
            # 直接点击标题区域以获取焦点
            pyautogui.click(center)
            logging.info("已通过识别标题获取导出窗口焦点")
        else:
            # 如果找不到标题，退回到中心点击方法
            logging.warning("未找到导出设置标题，使用屏幕中心点击获取焦点")
            pyautogui.click(pyautogui.size()[0] // 2, pyautogui.size()[1] // 2)
    except Exception as e:
        logging.warning(f"获取窗口焦点异常: {str(e)}，使用屏幕中心点击")
        pyautogui.click(pyautogui.size()[0] // 2, pyautogui.size()[1] // 2)
    
    time.sleep(0.5)
    
    # 直接使用全局定义的EXPORT_OPTIONS，不再重复定义all_options
    
    # 初始强力滚动到顶部
    for _ in range(3):
        pyautogui.scroll(800)  # 向上滚动
        time.sleep(1.5)
    
    # 再向下滚动一点，确保从选项开始的位置
    pyautogui.scroll(-400)
    time.sleep(1.5)
    
    # 动态滚动查找
    options_found = set()
    
    for attempt in range(10):
        logging.info(f"选项查找尝试 #{attempt+1}")
        
        for option in EXPORT_OPTIONS:  # 使用EXPORT_OPTIONS替代all_options
            if option in options_found:
                continue
                
            try:
                # 定位选项文字，使用对应语言的截图
                text_loc = pyautogui.locateOnScreen(
                    os.path.join(SCREENSHOT_DIR, language, f"{option}.png"),
                    confidence=0.7
                )
                if text_loc:
                    # 直接点击选项文字
                    center = pyautogui.center(text_loc)
                    pyautogui.click(center)  # 直接点击文字中央
                    #pyautogui.click(center.x - 50, center.y)  # 点击文字左侧约50像素处的复选框
                    logging.info(f"已点击选项：{option}")
                    time.sleep(0.2)
                    
                    options_found.add(option)
            except Exception as e:
                logging.debug(f"选项处理异常：{option} - {str(e)}")
                continue
        
        # 如果所有选项都已找到，则退出循环
        if len(options_found) == len(EXPORT_OPTIONS):  # 使用EXPORT_OPTIONS替代all_options
            logging.info("所有选项已找到并处理")
            break
            
        # 强力向下滚动，使用更大的滚动距离
        pyautogui.scroll(-500)
        time.sleep(1)
        
    # 记录未找到的选项
    if len(options_found) < len(EXPORT_OPTIONS):  # 使用EXPORT_OPTIONS替代all_options
        missing = set(EXPORT_OPTIONS) - options_found  # 使用EXPORT_OPTIONS替代all_options
        logging.warning(f"未能找到以下选项: {missing}")

def detect_language(client_path):
    """检测客户端界面语言，使用设置菜单项来判断，并直接点击进入设置"""
    # 先点击汉堡菜单，以便能看到设置菜单项
    try:
        # 汉堡菜单在所有语言中都是相同的
        location = pyautogui.locateOnScreen(
            os.path.join(SCREENSHOT_DIR, DEFAULT_LANGUAGE, "hamburger_menu.png"),
            confidence=0.8
        )
        if location:
            center = pyautogui.center(location)
            pyautogui.click(center)
            time.sleep(1.5)  # 等待菜单展开
        else:
            logging.warning("无法找到汉堡菜单按钮，语言检测可能不准确")
    except Exception as e:
        logging.debug(f"点击汉堡菜单异常: {str(e)}")
    
    # 检测各语言的设置菜单项
    for language in SUPPORTED_LANGUAGES:
        try:
            # 尝试查找该语言的设置菜单项
            location = pyautogui.locateOnScreen(
                os.path.join(SCREENSHOT_DIR, language, "settings_menu_item.png"),
                confidence=0.75
            )
            if location:
                logging.info(f"检测到界面语言: {language}")
                # 直接点击设置菜单项，进入设置
                center = pyautogui.center(location)
                pyautogui.click(center)
                time.sleep(1)  # 等待设置页面加载
                return language
        except Exception as e:
            logging.debug(f"语言检测异常 ({language}): {str(e)}")
    
    # 点击ESC关闭可能打开的菜单
    pyautogui.press('escape')
    time.sleep(0.5)
    
    # 如果无法检测到语言，返回默认语言
    logging.warning(f"无法检测界面语言，使用默认语言: {DEFAULT_LANGUAGE}")
    return DEFAULT_LANGUAGE

def export_telegram_data(client_path, export_base_dir):
    try:
        # 启动客户端
        app = Application().start(f'"{client_path}"')
        time.sleep(8)
        
        # 检测界面语言并已经点击了设置菜单
        language = detect_language(client_path)
        logging.info(f"使用语言: {language}")
        
        # 获取客户端目录名，用于文件命名
        client_dir = os.path.basename(os.path.dirname(client_path))
        
        # 在点击设置菜单后进行截图
        settings_screenshot = os.path.join(export_base_dir, f"{client_dir}_settings.png")
        pyautogui.screenshot(settings_screenshot)
        logging.info(f"已保存设置页面截图: {settings_screenshot}")
        time.sleep(1)

        # 检查是否有"Start Messaging"按钮（未登录状态）
        try:
            start_messaging = pyautogui.locateOnScreen(
                os.path.join(SCREENSHOT_DIR, language, "start_messaging_button.png"),
                confidence=0.7
            )
            if start_messaging:
                logging.warning(f"客户端未登录，跳过处理：{client_path}")
                return None  # 返回None表示未登录
        except Exception:
            pass  # 忽略查找异常
        
        # 检查是否能找到汉堡菜单按钮（已登录状态）
        hamburger_found = False
        for attempt in range(3):  # 尝试3次
            try:
                hamburger = pyautogui.locateOnScreen(
                    os.path.join(SCREENSHOT_DIR, language, "hamburger_menu.png"),
                    confidence=0.8
                )
                if hamburger:
                    hamburger_found = True
                    break
            except Exception:
                pass
            time.sleep(1)
        
        if not hamburger_found:
            logging.warning(f"无法确认客户端登录状态，跳过处理：{client_path}")
            return None  # 返回None表示状态异常
        
        # 由于detect_language已经点击了设置菜单，无需再点击汉堡菜单和设置
        # 直接进入高级设置
        if not find_and_click("advanced_tab.png", language=language):
            raise Exception("找不到高级选项")
        time.sleep(1)

        # 滚动查找导出按钮
        if not scroll_and_find_export(language):
            raise Exception("找不到导出按钮")
        time.sleep(1)
        
        # 执行选项勾选
        select_export_options(language)
        
        # 点击最终保存按钮
        if not find_and_click("save_button.png", timeout=20, language=language):
            raise Exception("找不到保存按钮")
        
        # 处理保存路径
        client_dir = os.path.basename(os.path.dirname(client_path))
        export_path = os.path.join(export_base_dir, client_dir)
        os.makedirs(export_path, exist_ok=True)
        time.sleep(2)
        
        # 获取导出前下载文件夹中的文件列表
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads", "Telegram Desktop")
        before_export_folders = []
        if os.path.exists(downloads_path):
            before_export_folders = [f for f in os.listdir(downloads_path) if os.path.isdir(os.path.join(downloads_path, f))]
            logging.info(f"导出前下载文件夹中的文件夹数量: {len(before_export_folders)}")
        
        # 输入路径并确认 (这里使用默认路径，不再手动指定)
        pyautogui.press('enter')
        time.sleep(1)
        
        # 添加一个返回值标志，表示是否成功导出
        export_success = False
        
        try:
            # 等待"Show My Data"按钮出现，表示导出已完成
            logging.info("等待导出完成，寻找'Show My Data'按钮...")
            show_data_found = False
            max_wait_time = 300  # 最长等待5分钟
            start_time = time.time()
            
            while time.time() - start_time < max_wait_time:
                try:
                    # 尝试定位"Show My Data"按钮
                    location = pyautogui.locateOnScreen(
                        os.path.join(SCREENSHOT_DIR, language, "show_my_data_button.png"),
                        confidence=0.7
                    )
                    if location:
                        logging.info("导出完成，已找到'Show My Data'按钮")
                        show_data_found = True
                        export_success = True  # 设置成功标志
                        
                        # 先关闭导出窗口
                        if find_and_click("close_button.png", timeout=10, language=language):
                            logging.info("已关闭导出窗口")
                        else:
                            logging.warning("未能找到关闭按钮，尝试继续执行")
                        
                        time.sleep(2)  # 等待窗口关闭
                        break
                except Exception as e:
                    logging.debug(f"查找按钮异常: {str(e)}")
                    pass
                
                time.sleep(2)
            
            if not show_data_found:
                logging.warning(f"等待超时，未找到'Show My Data'按钮，可能导出未完成")
                return False  # 返回失败标志
            
            # 查找导出后新生成的文件夹
            time.sleep(3)  # 等待文件系统更新
            if os.path.exists(downloads_path):
                after_export_folders = [f for f in os.listdir(downloads_path) if os.path.isdir(os.path.join(downloads_path, f))]
                
                # 找出新增的文件夹
                new_folders = [f for f in after_export_folders if f not in before_export_folders]
                
                if new_folders:
                    # 如果有多个新文件夹，选择最新创建的一个
                    newest_folder = None
                    newest_time = 0
                    
                    for folder in new_folders:
                        folder_path = os.path.join(downloads_path, folder)
                        folder_time = os.path.getctime(folder_path)
                        if folder_time > newest_time:
                            newest_time = folder_time
                            newest_folder = folder
                    
                    if newest_folder:
                        source_path = os.path.join(downloads_path, newest_folder)
                        
                        # 如果目标路径已存在，先删除
                        if os.path.exists(export_path) and os.path.isdir(export_path):
                            try:
                                shutil.rmtree(export_path)
                                logging.info(f"已删除已存在的目标文件夹: {export_path}")
                            except Exception as e:
                                logging.error(f"删除目标文件夹失败: {str(e)}")
                        
                        # 直接复制文件夹
                        try:
                            shutil.copytree(source_path, export_path)
                            logging.info(f"已将导出文件夹 {newest_folder} 复制到 {export_path}")
                        except Exception as e:
                            logging.error(f"复制文件夹失败: {str(e)}")
        
                # 复制成功后尝试删除源文件夹
                try:
                    shutil.rmtree(source_path)
                    logging.info(f"已删除源文件夹: {source_path}")
                except:
                    logging.warning(f"无法删除源文件夹: {source_path}")
            else:
                logging.warning(f"下载路径不存在: {downloads_path}")
            
            # 在成功移动导出文件夹后，移动截图到导出文件夹
            if os.path.exists(settings_screenshot):
                try:
                    final_screenshot_path = os.path.join(export_path, f"{client_dir}_settings.png")
                    shutil.move(settings_screenshot, final_screenshot_path)
                    logging.info(f"已将设置页面截图移动到导出文件夹: {final_screenshot_path}")
                except Exception as e:
                    logging.error(f"移动截图失败: {str(e)}")
            
            logging.info(f"导出完成：{client_dir}")
            return export_success  # 返回成功标志
            
        except Exception as e:
            # 如果处理失败，清理临时截图
            if 'settings_screenshot' in locals() and os.path.exists(settings_screenshot):
                try:
                    os.remove(settings_screenshot)
                    logging.info("已清理临时截图文件")
                except:
                    pass
            logging.error(f"处理失败：{client_path} - {str(e)}")
            return False  # 异常情况返回失败
    finally:
        try:
            # 尝试多种方式关闭应用
            close_success = False
            
            # 方法1: 使用app.kill()
            try:
                app.kill()
                logging.info("已通过app.kill()关闭客户端")
                # 不要立即标记为成功，因为可能只关闭了当前实例
            except Exception as e:
                logging.debug(f"app.kill()关闭失败: {str(e)}")
            
            # 方法2: 无论方法1是否成功，都尝试使用taskkill关闭所有Telegram实例
            try:
                # 使用taskkill强制关闭所有Telegram.exe进程
                os.system("taskkill /f /im Telegram.exe")
                logging.info("已通过taskkill强制关闭所有Telegram客户端")
                close_success = True
            except Exception as e:
                logging.debug(f"taskkill关闭失败: {str(e)}")
            
            # 无论关闭成功与否，都显示桌面
            try:
                # 使用Windows快捷键显示桌面
                pyautogui.hotkey('win', 'd')
                logging.info("已显示桌面，确保所有窗口不会干扰后续操作")
            except Exception as e:
                logging.debug(f"显示桌面失败: {str(e)}")
            
            # 无论关闭成功与否，都等待一段时间确保资源释放
            time.sleep(3)
        except Exception as e:
            # 捕获所有可能的异常，确保不影响后续程序运行
            logging.debug(f"关闭客户端过程中出现异常: {str(e)}")
            # 最后再尝试一次显示桌面
            try:
                pyautogui.hotkey('win', 'd')
            except:
                pass

# 添加支持多语言的滚动查找函数
def scroll_and_find_export(language="en"):
    """滚动屏幕并查找导出按钮，支持多语言"""
    for _ in range(SCROLL_ATTEMPTS):
        pyautogui.scroll(SCROLL_DISTANCE)
        time.sleep(0.8)
        if find_and_click("export_button.png", timeout=2, language=language):
            return True
    return False

# 在main函数中修改相应的判断逻辑
def main():
    try:
        # 检查所有支持语言的截图目录
        for language in SUPPORTED_LANGUAGES:
            try:
                check_screenshot_dir(language)
                logging.info(f"语言 {language} 的截图文件检查通过")
            except Exception as e:
                logging.error(f"语言 {language} 的截图文件检查失败: {str(e)}")
                print(f"语言 {language} 的截图文件检查失败: {str(e)}")
                return
    except Exception as e:
        print(f"配置错误：{str(e)}")
        return
    
    # 使用分隔线和表情符号增强交互体验
    print("\n")
    print("-" * 50 + "\n")

    # 客户端根目录输入
    source_dir = input("📁 请输入客户端根目录 :").strip()
    print(f"✅ 已确认客户端根目录：{source_dir}")

    # 导出目录输入
    export_dir = input("\n📁 请输入导出目录 :").strip()
    print(f"✅ 已确认导出目录：{export_dir}\n")
    os.makedirs(export_dir, exist_ok=True)
    
    # 查找所有客户端
    clients = []
    for root, dirs, files in os.walk(source_dir):
        if "最新.exe" in files:
            clients.append(os.path.join(root, "最新.exe"))
    
    print(f"找到 {len(clients)} 个客户端")
    
    # 记录所有导出的文件夹名称
    exported_folders = []
    
    # 记录处理失败的客户端
    failed_clients = []
    
    # 记录成功导出的客户端
    success_clients = []
    
    # 批量处理
    for idx, exe_path in enumerate(clients, 1):
        print(f"\n处理进度：{idx}/{len(clients)}")
        print("正在处理：" + os.path.dirname(exe_path))
        
        # 获取导出前的文件夹列表
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads", "Telegram Desktop")
        before_folders = set()
        if os.path.exists(downloads_path):
            before_folders = set(f for f in os.listdir(downloads_path) if os.path.isdir(os.path.join(downloads_path, f)))
        
        # 执行导出并获取结果
        client_dir = os.path.basename(os.path.dirname(exe_path))
        
        try:
            # 执行导出并获取是否成功的返回值
            export_success = export_telegram_data(exe_path, export_dir)
            
            # 根据export_telegram_data的返回值判断是否成功
            if export_success is False:  # 明确检查是否为False，因为未登录的情况下返回None
                print(f"警告：客户端 {client_dir} 未成功导出数据 (未找到'Show My Data'按钮)")
                failed_clients.append(os.path.dirname(exe_path))  # 保存完整路径
                logging.warning(f"客户端 {client_dir} 未成功导出数据 (未找到'Show My Data'按钮)")
            elif export_success is None:  # 处理未登录或其他提前返回的情况
                print(f"警告：客户端 {client_dir} 未处理 (可能未登录或状态异常)")
                failed_clients.append(os.path.dirname(exe_path))  # 保存完整路径
                logging.warning(f"客户端 {client_dir} 未处理 (可能未登录或状态异常)")
            else:
                print(f"客户端 {client_dir} 数据导出成功")
                success_clients.append(client_dir)  # 添加到成功列表
        except Exception as e:
            print(f"错误：客户端 {client_dir} 导出失败 - {str(e)}")
            failed_clients.append(os.path.dirname(exe_path))  # 保存完整路径
            logging.error(f"客户端 {client_dir} 导出失败 - {str(e)}")
        
        time.sleep(5)
        
        # 获取导出后的文件夹列表，找出新增的文件夹
        if os.path.exists(downloads_path):
            after_folders = set(f for f in os.listdir(downloads_path) if os.path.isdir(os.path.join(downloads_path, f)))
            new_folders = after_folders - before_folders
            exported_folders.extend(new_folders)
    
    # 所有处理完成后，只删除导出过程中生成的文件夹
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads", "Telegram Desktop")
    if os.path.exists(downloads_path):
        for folder in exported_folders:
            folder_path = os.path.join(downloads_path, folder)
            if os.path.exists(folder_path):
                try:
                    shutil.rmtree(folder_path)
                    logging.info(f"已删除导出文件夹: {folder_path}")
                    print(f"已删除导出文件夹: {folder_path}")
                except Exception as e:
                    logging.error(f"删除文件夹失败 {folder}: {str(e)}")
                    print(f"删除文件夹失败 {folder}: {str(e)}")
    
    # 输出处理结果摘要 - 移到循环外部
    print("\n========== 处理结果摘要 ==========")
    print(f"总客户端数量: {len(clients)}")
    print(f"成功导出数量: {len(success_clients)}")
    print(f"失败客户端数量: {len(failed_clients)}")
    
    # 将失败的客户端列表写入文件
    if failed_clients:
        print("\n以下客户端导出失败:")
        for client_path in failed_clients:
            print(f"- {client_path}")
        
        # 写入失败记录文件到程序所在目录
        failed_log_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "failed_exports.txt")
        try:
            with open(failed_log_path, "w", encoding="utf-8") as f:
                f.write(f"导出失败的客户端列表 (总计 {len(failed_clients)} 个):\n")
                f.write("="*50 + "\n")
                for client_path in failed_clients:
                    f.write(f"{client_path}\n")
            print(f"\n失败记录已保存至: {failed_log_path}")
        except Exception as e:
            print(f"保存失败记录文件时出错: {str(e)}")


if __name__ == "__main__":
    # 打印工具名称，使用格式化字符串和分隔线增强视觉效果，添加表情符号
    print("=" * 50)
    print(f"{'🎈Telegram数据自动导出工具🎈':^50}")
    print("=" * 50)

    # 输出支持的语言信息
    print("本程序通过关键按钮图像识别实现模拟点击的自动化导出，需要对相关按钮进行截图保存")
    print(f"支持的语言: {', '.join(SUPPORTED_LANGUAGES)}")
    print("-" * 50)

    # 提示用户准备截图文件，使用多行字符串和循环使代码更清晰，添加表情符号
    screenshot_prompts = [
        "📸 请确保已准备好以下截图文件（每种语言都需要）：",
        "👉 1. 菜单按钮 (hamburger_menu.png) - 左上角的三横线菜单按钮",
        "👉 2. 设置菜单项 (settings_menu_item.png) - 点击菜单后出现的设置选项",
        "👉 3. 高级选项卡 (advanced_tab.png) - 设置页面中的高级选项卡",
        "👉 4. 导出按钮 (export_button.png) - 高级设置中的导出数据按钮",
        "👉 5. 保存按钮 (save_button.png) - 导出设置页面底部的保存按钮",
        # 移除未使用的复选框提示
        # "👉 6. 复选框状态图 (checkbox_selected.png / checkbox_unselected.png) - 导出选项的勾选框状态",
        "👉 6. 导出设置标题 (export_settings_title.png) - 导出设置窗口的标题",
        "👉 7. 导出选项文字截图 - 所有需要勾选的选项文字："
    ]
    
    # 添加所有导出选项的提示
    for option in EXPORT_OPTIONS:
        screenshot_prompts.append(f"   • {option}.png - 对应选项的文字")
    
    screenshot_prompts.extend([
        "👉 8. 显示数据按钮 (show_my_data_button.png) - 导出完成后出现的按钮",
        "👉 9. 关闭按钮 (close_button.png) - 导出完成窗口的关闭按钮",
        "👉 10. 开始消息按钮 (start_messaging_button.png) - 未登录状态下的按钮"
    ])
    
    # 添加关于分辨率的提示
    resolution_prompts = [
        "⚠️ 重要提示：",
        "由于不同电脑分辨率不同，请在您自己的电脑上截取上述元素的图片",
        "截图时请确保只包含需要识别的元素，不要包含太多周围内容",
        "所有截图请保存在以下目录结构中：",
        f"screenshots/{'{语言代码}'}/图片名称.png",
        f"例如英文界面的菜单按钮：screenshots/en/hamburger_menu.png",
        f"当前支持的语言代码：{', '.join(SUPPORTED_LANGUAGES)}"
    ]
    
    # 输出所有提示
    for prompt in screenshot_prompts:
        print(prompt)
    
    print("\n" + "-" * 50 + "\n")
    
    for prompt in resolution_prompts:
        print(prompt)
    
    print("\n" + "-" * 50)

    # 调用主函数，添加表情符号
    input("\n🚀 准备好后，按回车键开始执行...")
    main()
