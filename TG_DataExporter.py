import os
import time
import logging
import pyautogui
import shutil
import subprocess  # 确保在文件顶部导入subprocess模块
import win32api
import win32con

# 有条件导入pythoncom，如果不可用则跳过
try:
    import pythoncom
    PYTHONCOM_AVAILABLE = True
except ImportError:
    PYTHONCOM_AVAILABLE = False
    logging.warning("pythoncom模块不可用，某些功能可能受限")

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
        time.sleep(0.5)
    
    # 再向下滚动一点，确保从选项开始的位置
    pyautogui.scroll(-400)
    time.sleep(0.5)
    
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
        # 添加调试信息，记录尝试查找白底汉堡菜单
        logging.info("尝试查找白底汉堡菜单...")
        white_menu_path = os.path.join(SCREENSHOT_DIR, DEFAULT_LANGUAGE, "hamburger_menu.png")
        logging.debug(f"白底菜单图片路径: {white_menu_path}")
        
        # 检查文件是否存在
        if not os.path.exists(white_menu_path):
            logging.warning(f"白底菜单图片不存在: {white_menu_path}")
        
        # 尝试查找白底汉堡菜单
        location = None
        try:
            location = pyautogui.locateOnScreen(
                white_menu_path,
                confidence=0.8
            )
            
            if location:
                logging.info(f"找到白底汉堡菜单，位置: {location}")
        except Exception as e:
            logging.error(f"查找白底汉堡菜单时出错: {str(e)}")
            # 不返回，继续尝试黑底菜单
        
        # 无论白底菜单是否找到或出错，都尝试查找黑底菜单
        if not location:
            logging.info("未找到白底汉堡菜单，尝试查找黑底汉堡菜单...")
            
            # 添加调试信息，记录尝试查找黑底汉堡菜单
            dark_menu_path = os.path.join(SCREENSHOT_DIR, DEFAULT_LANGUAGE, "hamburger_menu_dark.png")
            logging.debug(f"黑底菜单图片路径: {dark_menu_path}")
            
            # 检查文件是否存在
            if not os.path.exists(dark_menu_path):
                logging.warning(f"黑底菜单图片不存在: {dark_menu_path}")
            else:
                # 尝试降低confidence值查找黑底菜单
                for conf in [0.8, 0.7, 0.6]:
                    try:
                        logging.info(f"尝试使用confidence={conf}查找黑底菜单")
                        location = pyautogui.locateOnScreen(
                            dark_menu_path,
                            confidence=conf
                        )
                        if location:
                            logging.info(f"找到黑底汉堡菜单，位置: {location}，confidence: {conf}")
                            break
                    except Exception as e:
                        logging.error(f"使用confidence={conf}查找黑底菜单时出错: {str(e)}")
                        # 继续尝试下一个confidence值
            
        if location:
            center = pyautogui.center(location)
            logging.info(f"点击汉堡菜单，位置: {center}")
            pyautogui.click(center)
            time.sleep(2)  # 等待菜单展开
        else:
            # 如果使用默认语言找不到，尝试使用其他支持的语言
            logging.info("默认语言未找到汉堡菜单，尝试其他语言...")
            for language in SUPPORTED_LANGUAGES:
                if language == DEFAULT_LANGUAGE:
                    continue  # 跳过默认语言，因为已经尝试过了
                
                logging.info(f"尝试使用语言 {language} 查找汉堡菜单...")
                
                # 尝试查找白底汉堡菜单
                white_path = os.path.join(SCREENSHOT_DIR, language, "hamburger_menu.png")
                if os.path.exists(white_path):
                    logging.debug(f"尝试白底菜单: {white_path}")
                    location = pyautogui.locateOnScreen(
                        white_path,
                        confidence=0.8
                    )
                else:
                    logging.warning(f"白底菜单图片不存在: {white_path}")
                
                # 如果找不到白底菜单，尝试查找黑底菜单
                if not location:
                    dark_path = os.path.join(SCREENSHOT_DIR, language, "hamburger_menu_dark.png")
                    if os.path.exists(dark_path):
                        logging.debug(f"尝试黑底菜单: {dark_path}")
                        # 尝试降低confidence值查找黑底菜单
                        for conf in [0.8, 0.7, 0.6]:
                            logging.info(f"尝试使用confidence={conf}查找黑底菜单")
                            location = pyautogui.locateOnScreen(
                                dark_path,
                                confidence=conf
                            )
                            if location:
                                logging.info(f"找到黑底汉堡菜单，位置: {location}，confidence: {conf}")
                                break
                    else:
                        logging.warning(f"黑底菜单图片不存在: {dark_path}")
                
                if location:
                    center = pyautogui.center(location)
                    logging.info(f"点击汉堡菜单，位置: {center}，语言: {language}")
                    pyautogui.click(center)
                    time.sleep(2)  # 等待菜单展开
                    break
            
            if not location:
                # 添加截图以便调试
                debug_screenshot = os.path.join(os.getcwd(), "debug_screenshot.png")
                pyautogui.screenshot(debug_screenshot)
                logging.warning(f"无法找到汉堡菜单按钮，已保存调试截图: {debug_screenshot}")
                logging.warning("无法找到汉堡菜单按钮，语言检测可能不准确")
    except Exception as e:
        logging.debug(f"点击汉堡菜单异常: {str(e)}")
        # 保存异常时的截图
        try:
            error_screenshot = os.path.join(os.getcwd(), "error_screenshot.png")
            pyautogui.screenshot(error_screenshot)
            logging.error(f"发生异常，已保存错误截图: {error_screenshot}")
        except:
            pass
    
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
    process = None  # 在函数开始处定义process变量
    try:
        # 首先验证是否为Telegram客户端
        is_telegram, exe_name = is_telegram_exe(client_path)
        if not is_telegram:
            logging.warning(f"不是有效的Telegram客户端: {client_path}")
            return None
        
        # 使用subprocess启动客户端
        logging.info(f"启动客户端: {client_path} ({exe_name})")
        process = subprocess.Popen([client_path])
        time.sleep(5)  # 等待客户端启动
        
        # 检测界面语言并已经点击了设置菜单
        language = detect_language(client_path)
        logging.info(f"使用语言: {language}")
        
        # 获取客户端目录名，用于文件命名
        client_dir = os.path.basename(os.path.dirname(client_path))
        
        # 确保导出基础目录存在
        os.makedirs(export_base_dir, exist_ok=True)
        
        # 在点击设置菜单后进行截图
        settings_screenshot = os.path.join(export_base_dir, f"{client_dir}_settings.png")
        try:
            pyautogui.screenshot(settings_screenshot)
            logging.info(f"已保存设置页面截图: {settings_screenshot}")
        except Exception as e:
            logging.error(f"保存设置页面截图失败: {str(e)}")
            # 即使截图失败也继续执行
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
        except Exception as e:
            logging.debug(f"检查登录状态异常: {str(e)}")
            # 忽略查找异常，继续执行
        
        # 修改登录状态检查逻辑
        # 由于我们已经成功进入了设置页面，可以认为客户端已登录
        # 不再尝试查找汉堡菜单按钮，而是直接继续执行
        
        # 直接进入高级设置
        if not find_and_click("advanced_tab.png", language=language):
            logging.warning(f"找不到高级选项，可能客户端状态异常: {client_path}")
            # 保存当前屏幕截图以便调试
            debug_screenshot = os.path.join(export_base_dir, f"{client_dir}_debug.png")
            pyautogui.screenshot(debug_screenshot)
            logging.info(f"已保存调试截图: {debug_screenshot}")
            return None  # 返回None表示状态异常
        
        time.sleep(1)

        # 滚动查找导出按钮
        if not scroll_and_find_export(language):
            logging.warning(f"找不到导出按钮，可能客户端状态异常: {client_path}")
            return None  # 返回None表示状态异常
        
        time.sleep(1)
        
        # 执行选项勾选
        select_export_options(language)
        
        # 点击最终保存按钮
        if not find_and_click("save_button.png", timeout=20, language=language):
            logging.warning(f"找不到保存按钮，可能客户端状态异常: {client_path}")
            return None  # 返回None表示状态异常
        
        # 处理保存路径
        client_dir = os.path.basename(os.path.dirname(client_path))
        export_path = os.path.join(export_base_dir, client_dir)
        # 不再提前创建文件夹，而是在确认需要复制时再创建
        # os.makedirs(export_path, exist_ok=True)  # 移除这行
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
            max_wait_time = 1800  
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
                # 如果导出未完成，清理临时截图
                if settings_screenshot and os.path.exists(settings_screenshot):
                    try:
                        os.remove(settings_screenshot)
                        logging.info("已清理临时截图文件")
                    except Exception as e:
                        logging.debug(f"清理临时截图失败: {str(e)}")
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
                        
                        # 在确认有源文件夹后，再创建或清理目标文件夹
                        if os.path.exists(export_path):
                            try:
                                shutil.rmtree(export_path)
                                logging.info(f"已删除已存在的目标文件夹: {export_path}")
                            except Exception as e:
                                logging.error(f"删除目标文件夹失败: {str(e)}")
                        
                        # 创建父目录（如果不存在）
                        os.makedirs(os.path.dirname(export_path), exist_ok=True)
                        
                        # 直接复制文件夹
                        try:
                            shutil.copytree(source_path, export_path)
                            logging.info(f"已将导出文件夹 {newest_folder} 复制到 {export_path}")
                            
                            # 只有在成功复制导出文件夹后，才移动截图到导出文件夹
                            if settings_screenshot and os.path.exists(settings_screenshot):
                                try:
                                    final_screenshot_path = os.path.join(export_path, f"{client_dir}_settings.png")
                                    shutil.move(settings_screenshot, final_screenshot_path)
                                    logging.info(f"已将设置页面截图移动到导出文件夹: {final_screenshot_path}")
                                except Exception as e:
                                    logging.error(f"移动截图失败: {str(e)}")
                                    # 如果移动失败，不要删除原始截图
                        except Exception as e:
                            logging.error(f"复制文件夹失败: {str(e)}")
                            # 如果复制失败，清理临时截图
                            if settings_screenshot and os.path.exists(settings_screenshot):
                                try:
                                    os.remove(settings_screenshot)
                                    logging.info("已清理临时截图文件")
                                except Exception as e2:
                                    logging.debug(f"清理临时截图失败: {str(e2)}")
        
                # 复制成功后尝试删除源文件夹
                try:
                    shutil.rmtree(source_path)
                    logging.info(f"已删除源文件夹: {source_path}")
                    # 不在控制台输出删除源文件夹的信息
                except Exception as e:
                    logging.warning(f"无法删除源文件夹: {source_path}")
                    # 不在控制台输出无法删除的信息
            else:
                logging.warning(f"下载路径不存在: {downloads_path}")
                # 如果下载路径不存在，清理临时截图
                if settings_screenshot and os.path.exists(settings_screenshot):
                    try:
                        os.remove(settings_screenshot)
                        logging.info("已清理临时截图文件")
                    except Exception as e:
                        logging.debug(f"清理临时截图失败: {str(e)}")
            
            logging.info(f"导出完成：{client_dir}")
            return export_success  # 返回成功标志
            
        except Exception as e:
            # 如果处理失败，清理临时截图
            if settings_screenshot and os.path.exists(settings_screenshot):
                try:
                    os.remove(settings_screenshot)
                    logging.info("已清理临时截图文件")
                except Exception as e2:
                    logging.debug(f"清理临时截图失败: {str(e2)}")
            logging.error(f"处理失败：{client_path} - {str(e)}")
            return False  # 异常情况返回失败
    finally:
        try:
            # 尝试多种方式关闭应用
            close_success = False
            
            # 首先尝试使用process.terminate()关闭
            if process is not None:
                try:
                    process.terminate()
                    time.sleep(2)
                    logging.info("已尝试通过process.terminate()关闭客户端")
                except Exception as e:
                    logging.debug(f"process.terminate()关闭失败: {str(e)}")
            
            # 使用新的进程关闭函数
            telegram_processes = find_telegram_processes()
            if telegram_processes:
                for exe_name in telegram_processes:
                    try:
                        # 首先尝试正常关闭
                        subprocess.run(f"taskkill /IM {exe_name} /T", 
                                    shell=True, 
                                    stdout=subprocess.PIPE, 
                                    stderr=subprocess.PIPE)
                        logging.info(f"已尝试关闭进程: {exe_name}")
                        
                        # 如果正常关闭失败，使用强制关闭
                        subprocess.run(f"taskkill /F /IM {exe_name} /T", 
                                    shell=True, 
                                    stdout=subprocess.PIPE, 
                                    stderr=subprocess.PIPE)
                        logging.info(f"已强制关闭进程: {exe_name}")
                    except Exception as e:
                        logging.error(f"关闭进程失败 {exe_name}: {str(e)}")
            
            # 最后验证是否所有进程都已关闭
            remaining_processes = find_telegram_processes()
            if remaining_processes:
                logging.warning(f"以下进程仍在运行: {', '.join(remaining_processes)}")
                # 最后的强制关闭尝试
                for exe_name in remaining_processes:
                    try:
                        subprocess.run(f"taskkill /F /IM {exe_name}", 
                                    shell=True, 
                                    stdout=subprocess.PIPE, 
                                    stderr=subprocess.PIPE)
                    except Exception as e:
                        logging.error(f"最终强制关闭失败 {exe_name}: {str(e)}")
            
            # 无论关闭成功与否，都显示桌面并按Alt+Tab切换窗口焦点
            try:
                pyautogui.hotkey('win', 'd')
                time.sleep(1)
                pyautogui.hotkey('alt', 'tab')
                time.sleep(0.5)
                pyautogui.hotkey('win', 'd')
                logging.info("已显示桌面并切换窗口焦点")
            except Exception as e:
                logging.debug(f"显示桌面失败: {str(e)}")
            
            # 等待一段时间确保资源释放
            time.sleep(2)
            
        except Exception as e:
            logging.debug(f"关闭客户端过程中出现异常: {str(e)}")
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

def is_telegram_exe(exe_path):
    """
    检查可执行文件是否为Telegram客户端
    返回: (bool, str) - (是否为Telegram客户端, 可执行文件名)
    """
    try:
        # 获取文件版本信息
        info = win32api.GetFileVersionInfo(exe_path, "\\")
        ms = info['FileVersionMS']
        ls = info['FileVersionLS']
        version = f"{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}.{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}"
        
        # 获取文件属性信息
        lang, codepage = win32api.GetFileVersionInfo(exe_path, '\\VarFileInfo\\Translation')[0]
        str_info_path = f'\\StringFileInfo\\{lang:04x}{codepage:04x}\\'
        
        # 获取文件说明和产品名称
        file_description = win32api.GetFileVersionInfo(exe_path, str_info_path + 'FileDescription')
        product_name = win32api.GetFileVersionInfo(exe_path, str_info_path + 'ProductName')
        
        # 检查是否为Telegram Desktop
        is_telegram = (
            'Telegram' in file_description and 'Desktop' in file_description
        ) or (
            'Telegram' in product_name and 'Desktop' in product_name
        )
        
        if is_telegram:
            exe_name = os.path.basename(exe_path)
            logging.info(f"找到Telegram客户端: {exe_name} (版本: {version})")
            return True, exe_name
        return False, None
        
    except Exception as e:
        logging.debug(f"读取文件属性失败: {exe_path} - {str(e)}")
        return False, None

def find_telegram_processes():
    """
    查找所有正在运行的Telegram客户端进程
    返回: list of str - 进程名称列表
    """
    telegram_processes = set()
    try:
        # 使用wmic命令获取所有进程的详细信息
        cmd = 'wmic process get ExecutablePath,ProcessId /format:csv'
        result = subprocess.run(cmd, 
                              shell=True, 
                              stdout=subprocess.PIPE, 
                              stderr=subprocess.PIPE,
                              text=True)
        
        if result.returncode == 0:
            lines = result.stdout.strip().split('\n')
            for line in lines[1:]:  # 跳过标题行
                if line.strip():
                    parts = line.split(',')
                    if len(parts) >= 2 and parts[1].strip():  # 确保有可执行文件路径
                        exe_path = parts[1].strip()
                        if exe_path and os.path.exists(exe_path):
                            is_telegram, exe_name = is_telegram_exe(exe_path)
                            if is_telegram and exe_name:
                                telegram_processes.add(exe_name)
        
        return list(telegram_processes)
    except Exception as e:
        logging.error(f"查找Telegram进程失败: {str(e)}")
        return []

# 添加一个新函数，用于GUI程序调用
def run_export(source_dirs, export_dir, callback=None):
    """
    执行Telegram数据导出的主要功能，适用于GUI程序调用
    
    参数:
        source_dirs: 客户端根目录列表，支持多个目录
        export_dir: 导出目录
        callback: 可选的回调函数，用于更新GUI进度
    
    返回:
        dict: 包含导出结果的字典，包括成功列表、失败列表等
    """
    try:
        # 检查所有支持语言的截图目录
        for language in SUPPORTED_LANGUAGES:
            try:
                check_screenshot_dir(language)
                logging.info(f"语言 {language} 的截图文件检查通过")
                if callback:
                    callback(f"语言 {language} 的截图文件检查通过")
            except Exception as e:
                logging.error(f"语言 {language} 的截图文件检查失败: {str(e)}")
                if callback:
                    callback(f"语言 {language} 的截图文件检查失败: {str(e)}")
                return {"error": str(e)}
    except Exception as e:
        logging.error(f"配置错误：{str(e)}")
        if callback:
            callback(f"配置错误：{str(e)}")
        return {"error": str(e)}
    
    # 创建导出目录
    os.makedirs(export_dir, exist_ok=True)
    
    # 将单个目录转换为列表以统一处理
    if isinstance(source_dirs, str):
        source_dirs = [source_dirs]
    
    # 查找所有客户端
    clients = []
    for source_dir in source_dirs:
        source_dir_name = os.path.basename(source_dir)
        for root, dirs, files in os.walk(source_dir):
            for file in files:
                if file.lower().endswith('.exe'):
                    exe_path = os.path.join(root, file)
                    is_telegram, _ = is_telegram_exe(exe_path)
                    if is_telegram:
                        clients.append({
                            "path": exe_path,
                            "root_dir_name": source_dir_name
                        })
    
    if callback:
        callback(f"找到 {len(clients)} 个客户端")
    logging.info(f"找到 {len(clients)} 个客户端")
    
    # 记录所有导出的文件夹名称
    exported_folders = []
    
    # 记录处理失败的客户端
    failed_clients = []
    
    # 记录成功导出的客户端
    success_clients = []
    
    # 批量处理
    for idx, client_info in enumerate(clients, 1):
        exe_path = client_info["path"]
        root_dir_name = client_info["root_dir_name"]
        
        # 更新进度信息 - 这里需要发送进度信号
        if hasattr(callback, '__self__') and hasattr(callback.__self__, 'signals'):
            # 如果callback是GUI对象的方法，尝试发送进度信号
            try:
                callback.__self__.signals.update_progress.emit(idx, len(clients))
            except Exception as e:
                logging.debug(f"发送进度信号失败: {str(e)}")
        
        if callback:
            callback(f"处理进度：{idx}/{len(clients)}\n正在处理：{os.path.dirname(exe_path)}")
        logging.info(f"处理进度：{idx}/{len(clients)}")
        logging.info(f"正在处理：{os.path.dirname(exe_path)}")
        
        # 获取导出前的文件夹列表
        downloads_path = os.path.join(os.path.expanduser("~"), "Downloads", "Telegram Desktop")
        before_folders = set()
        if os.path.exists(downloads_path):
            before_folders = set(f for f in os.listdir(downloads_path) if os.path.isdir(os.path.join(downloads_path, f)))
        
        # 执行导出并获取结果
        client_dir = os.path.basename(os.path.dirname(exe_path))
        
        # 创建按照新结构的导出目录：导出目录-客户端根目录文件夹名-客户端文件夹名
        client_export_dir = os.path.join(export_dir, root_dir_name)
        
        try:
            # 执行导出并获取是否成功的返回值
            export_success = export_telegram_data(exe_path, client_export_dir)
            
            # 根据export_telegram_data的返回值判断是否成功
            if export_success is False:  # 明确检查是否为False，因为未登录的情况下返回None
                message = f"警告：客户端 {client_dir} (根目录: {root_dir_name}) 未成功导出数据 (未找到'Show My Data'按钮)"
                if callback:
                    callback(message)
                logging.warning(message)
                failed_clients.append(os.path.dirname(exe_path))  # 保存完整路径
            elif export_success is None:  # 处理未登录或其他提前返回的情况
                message = f"警告：客户端 {client_dir} (根目录: {root_dir_name}) 未处理 (可能未登录或状态异常)"
                if callback:
                    callback(message)
                logging.warning(message)
                failed_clients.append(os.path.dirname(exe_path))  # 保存完整路径
            else:
                message = f"客户端 {client_dir} (根目录: {root_dir_name}) 数据导出成功"
                if callback:
                    callback(message)
                logging.info(message)
                success_clients.append(f"{root_dir_name}/{client_dir}")  # 添加到成功列表，包含根目录信息
        except Exception as e:
            message = f"错误：客户端 {client_dir} (根目录: {root_dir_name}) 导出失败 - {str(e)}"
            if callback:
                callback(message)
            logging.error(message)
            failed_clients.append(os.path.dirname(exe_path))  # 保存完整路径
        
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
                except Exception as e:
                    logging.error(f"删除文件夹失败 {folder}: {str(e)}")
    
    # 处理结果摘要
    summary = {
        "total": len(clients),
        "success": len(success_clients),
        "failed": len(failed_clients),
        "success_list": success_clients,
        "failed_list": failed_clients
    }
    
    # 将失败的客户端列表写入文件
    if failed_clients:
        current_dir = os.getcwd()
        failed_log_path = os.path.join(current_dir, "failed_exports.txt")
        try:
            with open(failed_log_path, "w", encoding="utf-8") as f:
                f.write(f"导出失败的客户端列表 (总计 {len(failed_clients)} 个):\n")
                f.write("="*50 + "\n")
                for client_path in failed_clients:
                    f.write(f"{client_path}\n")
            if callback:
                callback(f"失败记录已保存至: {failed_log_path}")
            logging.info(f"失败记录已保存至: {failed_log_path}")
            summary["failed_log"] = failed_log_path
        except Exception as e:
            if callback:
                callback(f"保存失败记录文件时出错: {str(e)}")
            logging.error(f"保存失败记录文件时出错: {str(e)}")
    
    return summary

def close_telegram_processes():
    """关闭所有Telegram客户端进程"""
    try:
        # 获取所有Telegram进程
        telegram_processes = find_telegram_processes()
        
        if not telegram_processes:
            logging.info("未发现正在运行的Telegram客户端进程")
            return True
        
        for exe_name in telegram_processes:
            try:
                # 首先尝试正常关闭
                subprocess.run(f"taskkill /IM {exe_name} /T", 
                              shell=True, 
                              stdout=subprocess.PIPE, 
                              stderr=subprocess.PIPE)
                logging.info(f"已尝试关闭进程: {exe_name}")
                
                # 如果正常关闭失败，使用强制关闭
                subprocess.run(f"taskkill /F /IM {exe_name} /T", 
                              shell=True, 
                              stdout=subprocess.PIPE, 
                              stderr=subprocess.PIPE)
                logging.info(f"已强制关闭进程: {exe_name}")
            except Exception as e:
                logging.error(f"关闭进程失败 {exe_name}: {str(e)}")
        
        # 最后验证是否所有进程都已关闭
        remaining_processes = find_telegram_processes()
        if remaining_processes:
            logging.warning(f"以下进程仍在运行: {', '.join(remaining_processes)}")
            return False
        return True
    except Exception as e:
        logging.error(f"关闭进程时出错: {str(e)}")
        return False


# 保留原始main函数，但重命名为console_main，用于命令行模式
def console_main():
    # 原来的main函数代码
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

    # 客户端根目录输入，支持多个目录
    source_dirs_input = input("📁 请输入客户端根目录(多个目录用英文分号;分隔) :").strip()
    source_dirs = [dir.strip() for dir in source_dirs_input.split(';') if dir.strip()]
    
    if not source_dirs:
        print("❌ 未输入有效的客户端根目录，程序退出")
        return
    
    print(f"✅ 已确认客户端根目录：{', '.join(source_dirs)}")

    # 导出目录输入
    export_dir = input("\n📁 请输入导出目录 :").strip()
    print(f"✅ 已确认导出目录：{export_dir}\n")
    
    # 调用新的run_export函数，传入目录列表
    result = run_export(source_dirs, export_dir, callback=print)
    
    # 输出处理结果摘要
    print("\n========== 处理结果摘要 ==========")
    print(f"总客户端数量: {result['total']}")
    print(f"成功导出数量: {result['success']}")
    print(f"失败客户端数量: {result['failed']}")
    
    if result['failed'] > 0:
        print("\n以下客户端导出失败:")
        for client_path in result['failed_list']:
            print(f"- {client_path}")
        
        if 'failed_log' in result:
            print(f"\n失败记录已保存至: {result['failed_log']}")
    
    # 添加等待用户按键确认后退出
    print("\n" + "=" * 50)
    print("程序执行完毕！按任意键退出...")
    input()  # 等待用户按任意键

# 修改主入口点
if __name__ == "__main__":
    # 有条件初始化COM
    if PYTHONCOM_AVAILABLE:
        pythoncom.CoInitialize()
    
    try:
        # 打印工具名称，使用格式化字符串和分隔线增强视觉效果，添加表情符号
        print("=" * 50)
        print(f"{'🎈Telegram数据自动导出工具🎈':^50}")
        print("=" * 50)

        # 输出支持的语言信息
        print("本程序通过关键按钮图像识别实现模拟点击的自动化导出，需要对相关按钮进行截图保存")
        print(f"支持的语言: {', '.join(SUPPORTED_LANGUAGES)}")
        print("-" * 50)


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
        # 修改为调用console_main而不是main
        console_main()  # 将main()改为console_main()
    finally:
        # 有条件反初始化COM
        if PYTHONCOM_AVAILABLE:
            pythoncom.CoUninitialize()