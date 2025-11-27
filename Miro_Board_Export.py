import time
import json
import os
import csv
import datetime
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
# from webdriver_manager.microsoft import EdgeChromiumDriverManager # 移除依赖，避免网络问题

# ==================== 用户配置区域 (必填) ====================

# 1. Edge User Data 路径 (请确保 Edge 已完全关闭)
USER_DATA_DIR = r"C:\Users\112560\AppData\Local\Microsoft\Edge\User Data"

# 2. 配置文件目录，通常是 "Default"
PROFILE_DIR = "Default"

# 3. 链接保存的文件名
LINK_FILE = "miro_board_links.json"

# 4. 报告文件名
REPORT_FILE = "miro_export_report.csv"

# ============================================================

def setup_driver():
    """初始化浏览器驱动 (Edge)"""
    print(">>> [初始化] 正在启动 Edge 浏览器...")
    options = Options()
    options.add_argument(f"user-data-dir={USER_DATA_DIR}")
    options.add_argument(f"profile-directory={PROFILE_DIR}")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("detach", True)
    options.add_experimental_option('excludeSwitches', ['enable-logging']) # 屏蔽浏览器内部日志
    options.add_argument("--start-maximized")
    options.add_argument("--log-level=3")
    
    # 增加稳定性参数
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--enable-unsafe-swiftshader") # 允许使用软件渲染 WebGL
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    
    # 忽略证书错误等干扰
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')

    try:
        # 使用 Selenium 内置的驱动管理 (Selenium Manager)
        # 如果报错找不到驱动，请手动下载 msedgedriver.exe 放到同级目录
        driver = webdriver.Edge(service=Service(), options=options)
        print(">>> [初始化] 浏览器启动成功！")
        return driver
    except Exception as e:
        print("\n" + "="*60)
        print("!!! 启动浏览器失败 !!!")
        print("原因可能是：")
        print("1. Edge 浏览器正在运行。请先关闭所有 Edge 窗口！")
        print("2. 缺少 msedgedriver.exe 驱动。")
        print(f"详细错误: {e}")
        print("="*60 + "\n")
        raise e

def scrape_dashboard_links(driver, existing_links=None):
    """功能一：从 Dashboard 抓取所有白板链接 (支持增量更新) - 包含白板名称"""
    if existing_links is None:
        existing_links = []
        
    # 将现有链接转换为 URL 集合，方便去重
    seen_urls = set()
    for item in existing_links:
        if isinstance(item, dict) and 'url' in item:
            seen_urls.add(item['url'])
        elif isinstance(item, str): # 兼容旧格式
            seen_urls.add(item)
        
    print(">>> [阶段一] 正在打开 Dashboard 抓取链接...")
    driver.get("https://miro.com/app/dashboard/")
    
    # 等待 dashboard 核心元素加载
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "[data-testid='grid-view'], [data-testid='board-card']"))
        )
    except:
        print("等待 Dashboard 加载超时，尝试继续滚动...")

    time.sleep(5) # 给足时间让初始内容渲染
    print("正在执行自动滚动以加载所有白板...")
    
    # 使用字典存储本次会话抓取到的所有链接 (URL -> Item)
    scraped_items_map = {}
    
    scroll_unchanged_count = 0
    max_scroll_unchanged = 5 # 连续几次滚动位置没变判定为到底
    
    while True:
        # 1. 抓取当前视图可见的白板链接
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, "a[href*='/app/board/']")
            for el in elements:
                try:
                    href = el.get_attribute('href')
                    if not href or len(href) < 10: continue
                    
                    if href in scraped_items_map:
                        continue

                    name = el.text.strip()
                    if not name:
                        name = el.get_attribute("aria-label")
                    if not name:
                        try:
                            name_el = el.find_element(By.XPATH, ".//div[contains(@class, 'title')] | .//span[contains(@class, 'title')]")
                            name = name_el.text.strip()
                        except:
                            pass
                    
                    if name and '\n' in name:
                        name = name.split('\n')[0].strip()
                    
                    if not name:
                        name = "Untitled Board"

                    item = {"name": name, "url": href}
                    scraped_items_map[href] = item
                except:
                    continue
        except Exception as e:
            print(f"  [!] 抓取元素时发生轻微错误: {e}")

        print(f"  -> 当前已累计发现 {len(scraped_items_map)} 个白板...")

        # 2. 尝试定位滚动容器
        scrollable_container = None
        try:
            scrollable_container = driver.find_element(By.CSS_SELECTOR, "[data-testid='grid-view']")
        except:
            pass
            
        if not scrollable_container:
            try:
                divs = driver.find_elements(By.TAG_NAME, "div")
                max_scroll = 0
                for div in divs:
                    try:
                        sh = int(div.get_attribute("scrollHeight"))
                        ch = int(div.get_attribute("clientHeight"))
                        if sh > ch and sh > max_scroll:
                            max_scroll = sh
                            scrollable_container = div
                    except:
                        continue
            except:
                pass

        # 3. 执行滚动
        current_scroll_top = 0
        scrolled_via_script = False
        
        if scrollable_container:
            try:
                current_scroll_top = driver.execute_script("return arguments[0].scrollTop", scrollable_container)
                driver.execute_script("arguments[0].scrollBy(0, 400)", scrollable_container)
                scrolled_via_script = True
            except:
                pass
        
        if not scrolled_via_script:
            current_scroll_top = driver.execute_script("return window.pageYOffset || document.documentElement.scrollTop")
            driver.execute_script("window.scrollBy(0, 400);")

        time.sleep(3)
        
        # 4. 检查滚动位置
        new_scroll_top = 0
        if scrollable_container:
            try:
                new_scroll_top = driver.execute_script("return arguments[0].scrollTop", scrollable_container)
            except:
                pass
        else:
            new_scroll_top = driver.execute_script("return window.pageYOffset || document.documentElement.scrollTop")
            
        if abs(new_scroll_top - current_scroll_top) < 5:
            scroll_unchanged_count += 1
            print(f"     [.] 滚动位置未变化 ({scroll_unchanged_count}/{max_scroll_unchanged})")
            
            try:
                ActionChains(driver).send_keys(Keys.PAGE_DOWN).perform()
                time.sleep(1)
            except:
                pass
                
            if scroll_unchanged_count >= max_scroll_unchanged:
                print("  -> 判定已到达底部。")
                break
        else:
            scroll_unchanged_count = 0

    print("正在整合链接...")
    
    new_items = []
    for url, item in scraped_items_map.items():
        if url not in seen_urls:
            new_items.append(item)
            seen_urls.add(url)
            print(f"     [+] 新增: {item['name']} ({url[-20:]}...)")
            
    final_links = []
    for item in existing_links:
        if isinstance(item, str):
            final_links.append({"name": "Unknown (Old)", "url": item})
        else:
            final_links.append(item)
            
    final_links.extend(new_items)
    
    print(f"抓取完成。总计: {len(final_links)} (本次新增: {len(new_items)})")
    
    with open(LINK_FILE, 'w', encoding='utf-8') as f:
        json.dump(final_links, f, indent=4, ensure_ascii=False)
    print(f"链接已保存到: {LINK_FILE}\n")
    return final_links

def download_vector_pdf(driver, links_data):
    """功能二：批量导出 Vector PDF (智能兼容版) - 支持增量 CSV 报告"""
    wait_normal = WebDriverWait(driver, 20)      # 常规 UI 等待
    wait_long = WebDriverWait(driver, 600)       # 矢量图生成专用长等待 (10分钟)
    
    # 初始化 CSV 报告 (如果不存在则写入表头)
    if not os.path.exists(REPORT_FILE):
        try:
            with open(REPORT_FILE, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(["Timestamp", "Board Name", "URL", "Status", "Error Message"])
            print(f">>> [报告] 已创建新的报告文件: {REPORT_FILE}")
        except Exception as e:
            print(f"[警告] 无法创建报告文件: {e}")

    results = [] # 记录执行结果
    print(f">>> [阶段二] 开始执行批量下载，共 {len(links_data)} 个任务...")

    for index, item in enumerate(links_data):
        # 兼容处理：如果 item 是字符串 (旧格式)，转为 dict
        if isinstance(item, str):
            url = item
            name = "Unknown"
        else:
            url = item.get('url')
            name = item.get('name', 'Unknown')
            
        result = {"name": name, "url": url, "status": "Pending", "error": ""}
        
        try:
            print(f"[{index+1}/{len(links_data)}] 正在处理: {name}")
            print(f"     URL: {url}")
            driver.get(url)
            
            # 1. 强制等待页面基础渲染
            time.sleep(8) 

            # 消除可能的干扰弹窗 (Esc)
            try:
                ActionChains(driver).send_keys("\ue00c").perform() # ESC 键
            except:
                pass

            # 2. 寻找导出入口 (Main Menu -> Board -> Export)
            print("  -> 寻找 Export 菜单 (3级菜单模式)...")
            export_menu_opened = False
            
            try:
                # 第一级：点击 Main Menu (汉堡菜单)
                menu_btn = None
                try:
                    # 优先尝试 aria-label
                    menu_btn = driver.find_element(By.XPATH, "//button[@aria-label='Main menu'] | //div[@aria-label='Main menu']")
                except:
                    # 备用 selector
                    menu_btn = driver.find_element(By.CSS_SELECTOR, "[data-testid='board-header__main-menu-button']")
                
                if menu_btn:
                    menu_btn.click()
                    print("     [UI] 点击 Main Menu 成功")
                    time.sleep(1.5) # 稍作等待，确保菜单动画完成
                    
                    # 第二级 & 第三级：尝试悬停 Board 并寻找 Export (增加重试机制)
                    max_retries = 3
                    for attempt in range(max_retries):
                        print(f"     [UI] 尝试悬停 Board (第 {attempt+1} 次)...")
                        
                        # 2.1 寻找并悬停 Board
                        board_found = False
                        # 使用精确匹配，避免匹配到侧边栏的 "Board" 字样
                        candidates = driver.find_elements(By.XPATH, "//*[normalize-space(text())='Board']")
                        for cand in candidates:
                            try:
                                if cand.is_displayed():
                                    # 先移回菜单按钮重置一下位置，防止鼠标死锁
                                    ActionChains(driver).move_to_element(menu_btn).perform()
                                    time.sleep(0.2)
                                    
                                    ActionChains(driver).move_to_element(cand).perform()
                                    print("     [UI] 悬停 Board 动作已执行")
                                    board_found = True
                                    break
                            except:
                                continue
                        
                        if not board_found:
                            print("     [X] 未找到可见的 Board 选项，重试中...")
                            time.sleep(1)
                            continue

                        time.sleep(1.5) # 等待子菜单展开
                        
                        # 2.2 寻找并悬停 Export
                        print("     [UI] 正在寻找 Export 选项...")
                        export_found = False
                        # 使用精确匹配
                        candidates = driver.find_elements(By.XPATH, "//*[normalize-space(text())='Export']")
                        for cand in candidates:
                            try:
                                if cand.is_displayed():
                                    ActionChains(driver).move_to_element(cand).perform()
                                    print("     [UI] 悬停 Export 成功")
                                    export_found = True
                                    export_menu_opened = True
                                    break
                            except:
                                continue
                        
                        if export_found:
                            break # 成功找到 Export，跳出重试循环
                        else:
                            print("     [!] 未找到 Export 选项，可能是 Board 悬停失效，准备重试...")
                            time.sleep(1)

            except Exception as e:
                print(f"  [X] 无法打开菜单: {e}")
                result["status"] = "Failed"
                result["error"] = f"Menu Error: {str(e)}"
                write_csv_report(result)
                results.append(result)
                continue
            
            if not export_menu_opened:
                print(f"  [X] 导出菜单未打开，跳过: {url}")
                result["status"] = "Failed"
                result["error"] = "Export menu not opened"
                write_csv_report(result)
                results.append(result)
                continue

            time.sleep(1)

            # 3. 点击 "Save as PDF"
            try:
                save_pdf = wait_normal.until(EC.element_to_be_clickable(
                    (By.XPATH, "//span[contains(text(), 'Save as PDF')] | //div[contains(text(), 'Save as PDF')]")
                ))
                save_pdf.click()
            except Exception as e:
                print(f"  [X] 点击 Save as PDF 失败: {e}")
                result["status"] = "Failed"
                result["error"] = f"Save PDF Error: {str(e)}"
                write_csv_report(result)
                results.append(result)
                continue

            # 4. 选择 "Vector" 选项 (根据用户截图恢复)
            try:
                print("     [UI] 等待导出弹窗并选择 Vector...")
                # 尝试定位 Vector 选项 (通常是 Radio Button 或 Label)
                vector_option = wait_normal.until(EC.element_to_be_clickable(
                    (By.XPATH, "//label[contains(., 'Vector')] | //div[contains(text(), 'Vector')]")
                ))
                vector_option.click()
                print("     [UI] 已选择 Vector")
            except Exception as e:
                print(f"  [X] 选择 Vector 选项失败: {e}")
                result["status"] = "Failed"
                result["error"] = f"Vector Option Error: {str(e)}"
                write_csv_report(result)
                results.append(result)
                continue

            # 5. 点击 Download (Export) 按钮
            try:
                print("     [UI] 正在点击 Export 按钮...")
                # 按钮文本通常是 "Export"
                export_btn = wait_normal.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(., 'Export')] | //button[contains(@class, 'button') and contains(., 'Export')]")
                ))
                export_btn.click()
                print(f"     [SUCCESS] 导出指令已发送: {url}")
                
                # 给予一定的缓冲时间让下载请求发出
                time.sleep(2)
                
            except Exception as e:
                print(f"  [X] 点击 Export 按钮失败: {e}")
                result["status"] = "Failed"
                result["error"] = f"Export Button Error: {str(e)}"
                write_csv_report(result)
                results.append(result)
                continue

            # 6. 等待并点击 "Download file" 按钮 (生成完成后)
            try:
                print("     [UI] 等待 PDF 生成及下载按钮出现 (可能需要较长时间)...")
                # 按钮文本是 "Download file"
                # 使用 wait_long (10分钟) 因为生成可能很慢
                download_file_btn = wait_long.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(., 'Download file')] | //div[contains(text(), 'Download file')]")
                ))
                download_file_btn.click()
                print(f"     [SUCCESS] 已点击 Download file 按钮，开始下载")
                result["status"] = "Success"
                
                # 缓冲时间确保下载开始
                time.sleep(5)
            except Exception as e:
                print(f"  [X] 等待/点击 Download file 按钮失败 (或生成超时): {e}")
                result["status"] = "Failed"
                result["error"] = f"Download File Button Error: {str(e)}"
                write_csv_report(result)
                results.append(result)
                continue

        except Exception as e:
            print(f"  [ERROR] 处理白板出错: {e}")
            result["status"] = "Failed"
            result["error"] = f"General Error: {str(e)}"
            write_csv_report(result)
            results.append(result)
            continue
        
        # 成功的情况也要写入
        write_csv_report(result)
        results.append(result)
    
    return results

def write_csv_report(result):
    """辅助函数：将单条结果写入 CSV"""
    try:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(REPORT_FILE, 'a', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow([timestamp, result["name"], result["url"], result["status"], result["error"]])
        print(f"     [报告] 已写入 CSV: {result['status']}")
    except Exception as e:
        print(f"     [警告] 写入 CSV 失败: {e}")

def main():
    print("="*50)
    print("       Miro 批量导出脚本启动 (Edge版)")
    print("="*50)
    driver = setup_driver()
    try:
        # 1. 链接管理 (读取 + 更新)
        links = []
        if os.path.exists(LINK_FILE):
            print(f"\n[提示] 发现本地文件 {LINK_FILE}，读取现有链接...")
            try:
                with open(LINK_FILE, 'r', encoding='utf-8') as f:
                    links = json.load(f)
            except json.JSONDecodeError:
                print(f"[警告] {LINK_FILE} 文件格式有误，将重新抓取...")
                links = []
        else:
            print(f"\n[提示] 本地未发现 {LINK_FILE}，将执行全新抓取...")

        # 始终尝试更新/补充链接
        print(f"\n[提示] 正在检查 Dashboard 以获取新白板...")
        links = scrape_dashboard_links(driver, existing_links=links)
            
        if not links:
            print("[警告] 未找到任何白板链接，程序即将退出。")
            return

        print(f"\n>>> 准备处理 {len(links)} 个白板...")
        
        # 2. 执行下载并获取结果
        results = download_vector_pdf(driver, links)
        
        # 3. 生成报告 (控制台简报)
        print("\n" + "="*50)
        print("       执行报告 Summary")
        print("="*50)
        
        success_count = sum(1 for r in results if r['status'] == 'Success')
        fail_count = len(results) - success_count
        
        print(f"总任务数: {len(results)}")
        print(f"成功: {success_count}")
        print(f"失败: {fail_count}")
        print(f"详细报告已保存至: {REPORT_FILE}")
        
        if fail_count > 0:
            print("\n[失败详情]")
            for r in results:
                if r['status'] == 'Failed':
                    print(f" - {r['name']} ({r['url']}): {r['error']}")
        
        print("="*50)

    except Exception as e:
        print(f"\n[!!!] 主程序发生异常: {e}")
    finally:
        print("\n" + "="*50)
        print("       任务结束")
        print("="*50)
        # input("按 Enter 关闭浏览器...") # 调试时可开启此行
        driver.quit()

if __name__ == "__main__":
    main()