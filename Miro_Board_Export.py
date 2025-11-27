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
# from webdriver_manager.microsoft import EdgeChromiumDriverManager # Removed dependency to avoid network issues / 移除依赖，避免网络问题

# ==================== User Configuration (Required) / 用户配置区域 (必填) ====================

# 1. Edge User Data Path (Ensure Edge is completely closed) / Edge User Data 路径 (请确保 Edge 已完全关闭)
USER_DATA_DIR = r"C:\Users\112560\AppData\Local\Microsoft\Edge\User Data"

# 2. Profile Directory, usually "Default" / 配置文件目录，通常是 "Default"
PROFILE_DIR = "Default"

# 3. Filename to save links / 链接保存的文件名
LINK_FILE = "miro_board_links.json"

# 4. Report filename / 报告文件名
REPORT_FILE = "miro_export_report.csv"

# ============================================================

def setup_driver():
    """Initialize Browser Driver (Edge) / 初始化浏览器驱动 (Edge)"""
    print(">>> [Init] Starting Edge browser... / [初始化] 正在启动 Edge 浏览器...")
    options = Options()
    options.add_argument(f"user-data-dir={USER_DATA_DIR}")
    options.add_argument(f"profile-directory={PROFILE_DIR}")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("detach", True)
    options.add_experimental_option('excludeSwitches', ['enable-logging']) # Suppress internal browser logs / 屏蔽浏览器内部日志
    options.add_argument("--start-maximized")
    options.add_argument("--log-level=3")
    
    # Stability options / 增加稳定性参数
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-gpu")
    options.add_argument("--enable-unsafe-swiftshader") # Allow software rendering for WebGL / 允许使用软件渲染 WebGL
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    
    # Ignore certificate errors / 忽略证书错误等干扰
    options.add_argument('--ignore-certificate-errors')
    options.add_argument('--ignore-ssl-errors')

    try:
        # Use Selenium Manager / 使用 Selenium 内置的驱动管理
        # If driver not found, manually download msedgedriver.exe to the same directory / 如果报错找不到驱动，请手动下载 msedgedriver.exe 放到同级目录
        driver = webdriver.Edge(service=Service(), options=options)
        print(">>> [Init] Browser started successfully! / [初始化] 浏览器启动成功！")
        return driver
    except Exception as e:
        print("\n" + "="*60)
        print("!!! Failed to start browser / 启动浏览器失败 !!!")
        print("Possible reasons / 原因可能是：")
        print("1. Edge is running. Please close all Edge windows! / Edge 浏览器正在运行。请先关闭所有 Edge 窗口！")
        print("2. Missing msedgedriver.exe driver. / 缺少 msedgedriver.exe 驱动。")
        print(f"Detailed error / 详细错误: {e}")
        print("="*60 + "\n")
        raise e

def scrape_dashboard_links(driver, existing_links=None):
    """Feature 1: Scrape all board links from Dashboard (Incremental Update) / 功能一：从 Dashboard 抓取所有白板链接 (支持增量更新)"""
    if existing_links is None:
        existing_links = []
        
    # Convert existing links to a set of URLs for deduplication / 将现有链接转换为 URL 集合，方便去重
    seen_urls = set()
    for item in existing_links:
        if isinstance(item, dict) and 'url' in item:
            seen_urls.add(item['url'])
        elif isinstance(item, str): # Compatibility with old format / 兼容旧格式
            seen_urls.add(item)
        
    print(">>> [Phase 1] Opening Dashboard to scrape links... / [阶段一] 正在打开 Dashboard 抓取链接...")
    driver.get("https://miro.com/app/dashboard/")
    
    # Wait for dashboard core elements / 等待 dashboard 核心元素加载
    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "[data-testid='grid-view'], [data-testid='board-card']"))
        )
    except:
        print("Timeout waiting for Dashboard, attempting to scroll anyway... / 等待 Dashboard 加载超时，尝试继续滚动...")

    time.sleep(5) # Allow time for initial content rendering / 给足时间让初始内容渲染
    print("Executing auto-scroll to load all boards... / 正在执行自动滚动以加载所有白板...")
    
    # Use a dictionary to store all links scraped in this session (URL -> Item)
    # This handles virtual scrolling (items being unloaded) / 使用字典存储本次会话抓取到的所有链接 (URL -> Item)
    scraped_items_map = {}
    
    scroll_unchanged_count = 0
    max_scroll_unchanged = 5 # Stop if scroll position doesn't change for N checks / 连续几次滚动位置没变判定为到底
    
    while True:
        # 1. Scrape currently visible board links / 抓取当前视图可见的白板链接
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
            print(f"  [!] Minor error scraping elements: {e} / 抓取元素时发生轻微错误: {e}")

        print(f"  -> Currently found {len(scraped_items_map)} boards... / 当前已累计发现 {len(scraped_items_map)} 个白板...")

        # 2. Try to locate scrollable container / 尝试定位滚动容器
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

        # 3. Execute Scroll / 执行滚动
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
        
        # 4. Check scroll position / 检查滚动位置
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
            print(f"     [.] Scroll position unchanged ({scroll_unchanged_count}/{max_scroll_unchanged}) / 滚动位置未变化")
            
            try:
                ActionChains(driver).send_keys(Keys.PAGE_DOWN).perform()
                time.sleep(1)
            except:
                pass
                
            if scroll_unchanged_count >= max_scroll_unchanged:
                print("  -> Reached bottom. / 判定已到达底部。")
                break
        else:
            scroll_unchanged_count = 0

    print("Consolidating links... / 正在整合链接...")
    
    new_items = []
    for url, item in scraped_items_map.items():
        if url not in seen_urls:
            new_items.append(item)
            seen_urls.add(url)
            print(f"     [+] New: {item['name']} ({url[-20:]}...) / 新增: {item['name']}")
            
    final_links = []
    for item in existing_links:
        if isinstance(item, str):
            final_links.append({"name": "Unknown (Old)", "url": item})
        else:
            final_links.append(item)
            
    final_links.extend(new_items)
    
    print(f"Scraping completed. Total: {len(final_links)} (New: {len(new_items)}) / 抓取完成。总计: {len(final_links)} (本次新增: {len(new_items)})")
    
    with open(LINK_FILE, 'w', encoding='utf-8') as f:
        json.dump(final_links, f, indent=4, ensure_ascii=False)
    print(f"Links saved to: {LINK_FILE} / 链接已保存到: {LINK_FILE}\n")
    return final_links

def download_vector_pdf(driver, links_data):
    """Feature 2: Batch Export Vector PDF (Smart Compatible) / 功能二：批量导出 Vector PDF (智能兼容版)"""
    wait_normal = WebDriverWait(driver, 20)      # Normal UI wait / 常规 UI 等待
    wait_long = WebDriverWait(driver, 600)       # Long wait for Vector generation (10 mins) / 矢量图生成专用长等待 (10分钟)
    
    # Initialize CSV report / 初始化 CSV 报告
    successful_urls = set()
    
    # Check and normalize CSV format / 检查并标准化 CSV 格式
    if os.path.exists(REPORT_FILE):
        try:
            rows_to_keep = []
            header_needs_update = False
            
            with open(REPORT_FILE, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                try:
                    header = next(reader)
                except StopIteration:
                    header = []
                
                # Check if it's the old 4-column header / 检查是否为旧的 4 列头部
                if header and "Board Name" not in header and "URL" in header:
                    print(">>> [Init] Detected old report format, migrating... / [初始化] 检测到旧版报告格式，正在迁移...")
                    header_needs_update = True
                
                for row in reader:
                    if not row: continue
                    
                    # Normalize row to 5 columns / 将行标准化为 5 列
                    # Old format: Timestamp, URL, Status, Error Message (4 cols)
                    # New format: Timestamp, Board Name, URL, Status, Error Message (5 cols)
                    
                    new_row = []
                    if len(row) == 4:
                        # Migration: Insert "Unknown" as Board Name / 迁移：插入 "Unknown" 作为白板名称
                        new_row = [row[0], "Unknown", row[1], row[2], row[3]]
                    elif len(row) >= 5:
                        new_row = row[:5] # Keep first 5 columns / 保留前 5 列
                    else:
                        # Invalid row, skip or try best effort / 无效行，跳过或尽力尝试
                        continue
                        
                    rows_to_keep.append(new_row)
                    
                    # Collect successful URLs / 收集成功的 URL
                    # Index 2 is URL, Index 3 is Status in the NEW format
                    if new_row[3] == "Success" and new_row[2]:
                        successful_urls.add(new_row[2])

            # Rewrite file if needed (header update or mixed content) / 如果需要（头部更新或混合内容），重写文件
            # We always rewrite to ensure consistency / 我们始终重写以确保持致性
            with open(REPORT_FILE, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(["Timestamp", "Board Name", "URL", "Status", "Error Message"])
                writer.writerows(rows_to_keep)
                
            print(f">>> [Init] Report normalized. Found {len(successful_urls)} exported boards. / [初始化] 报告已标准化。发现 {len(successful_urls)} 个已导出的白板。")

        except Exception as e:
            print(f"[Warning] Failed to process report file: {e} / [警告] 处理报告文件失败: {e}")
            # Fallback: try to continue without history if critical error / 降级：如果发生严重错误，尝试不带历史记录继续
    else:
        try:
            with open(REPORT_FILE, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(["Timestamp", "Board Name", "URL", "Status", "Error Message"])
            print(f">>> [Report] Created new report file: {REPORT_FILE} / [报告] 已创建新的报告文件")
        except Exception as e:
            print(f"[Warning] Failed to create report file: {e} / [警告] 无法创建报告文件: {e}")

    results = [] # Record results / 记录执行结果
    print(f">>> [Phase 2] Starting batch download, {len(links_data)} tasks... / [阶段二] 开始执行批量下载，共 {len(links_data)} 个任务...")

    for index, item in enumerate(links_data):
        # Compatibility / 兼容处理
        if isinstance(item, str):
            url = item
            name = "Unknown"
        else:
            url = item.get('url')
            name = item.get('name', 'Unknown')
            
        # Skip if already successful / 如果已成功则跳过
        if url in successful_urls:
            print(f"[{index+1}/{len(links_data)}] Skipping (Already Exported): {name} / 跳过 (已导出)")
            continue

        result = {"name": name, "url": url, "status": "Pending", "error": ""}
        
        try:
            print(f"[{index+1}/{len(links_data)}] Processing: {name} / 正在处理: {name}")
            print(f"     URL: {url}")
            driver.get(url)
            
            # 1. Force wait for page render / 强制等待页面基础渲染
            time.sleep(8) 

            # Dismiss potential popups (Esc) / 消除可能的干扰弹窗 (Esc)
            try:
                ActionChains(driver).send_keys("\ue00c").perform() # ESC Key
            except:
                pass

            # 2. Find Export Menu (Main Menu -> Board -> Export) / 寻找导出入口
            print("  -> Finding Export menu... / 寻找 Export 菜单...")
            export_menu_opened = False
            
            try:
                # Level 1: Click Main Menu / 第一级：点击 Main Menu
                menu_btn = None
                try:
                    # Try aria-label first / 优先尝试 aria-label
                    menu_btn = driver.find_element(By.XPATH, "//button[@aria-label='Main menu'] | //div[@aria-label='Main menu']")
                except:
                    # Backup selector / 备用 selector
                    menu_btn = driver.find_element(By.CSS_SELECTOR, "[data-testid='board-header__main-menu-button']")
                
                if menu_btn:
                    menu_btn.click()
                    print("     [UI] Clicked Main Menu / 点击 Main Menu 成功")
                    time.sleep(1.5) # Wait for animation / 稍作等待
                    
                    # Level 2 & 3: Hover Board and find Export / 第二级 & 第三级：尝试悬停 Board 并寻找 Export
                    max_retries = 3
                    for attempt in range(max_retries):
                        print(f"     [UI] Attempting to hover Board (Try {attempt+1})... / 尝试悬停 Board (第 {attempt+1} 次)...")
                        
                        # 2.1 Find and Hover Board / 寻找并悬停 Board
                        board_found = False
                        # Exact match / 使用精确匹配
                        candidates = driver.find_elements(By.XPATH, "//*[normalize-space(text())='Board']")
                        for cand in candidates:
                            try:
                                if cand.is_displayed():
                                    # Reset mouse position / 先移回菜单按钮重置一下位置
                                    ActionChains(driver).move_to_element(menu_btn).perform()
                                    time.sleep(0.2)
                                    
                                    ActionChains(driver).move_to_element(cand).perform()
                                    print("     [UI] Hovered Board / 悬停 Board 动作已执行")
                                    board_found = True
                                    break
                            except:
                                continue
                        
                        if not board_found:
                            print("     [X] Board option not visible, retrying... / 未找到可见的 Board 选项，重试中...")
                            time.sleep(1)
                            continue

                        time.sleep(1.5) # Wait for submenu / 等待子菜单展开
                        
                        # 2.2 Find and Hover Export / 寻找并悬停 Export
                        print("     [UI] Finding Export option... / 正在寻找 Export 选项...")
                        export_found = False
                        # Exact match / 使用精确匹配
                        candidates = driver.find_elements(By.XPATH, "//*[normalize-space(text())='Export']")
                        for cand in candidates:
                            try:
                                if cand.is_displayed():
                                    ActionChains(driver).move_to_element(cand).perform()
                                    print("     [UI] Hovered Export / 悬停 Export 成功")
                                    export_found = True
                                    export_menu_opened = True
                                    break
                            except:
                                continue
                        
                        if export_found:
                            break # Success / 成功
                        else:
                            print("     [!] Export option not found, retrying... / 未找到 Export 选项，准备重试...")
                            time.sleep(1)

            except Exception:
                # Simplified error / 简化错误信息
                print(f"  [X] Failed to open menu / 无法打开菜单")
                result["status"] = "Failed"
                result["error"] = "Main Menu or Board option not found"
                write_csv_report(result)
                results.append(result)
                continue
            
            if not export_menu_opened:
                print(f"  [X] Export menu not opened, skipping: {url} / 导出菜单未打开，跳过")
                result["status"] = "Failed"
                result["error"] = "Export menu not opened"
                write_csv_report(result)
                results.append(result)
                continue

            time.sleep(1)

            # 3. Click "Save as PDF" / 点击 "Save as PDF"
            try:
                save_pdf = wait_normal.until(EC.element_to_be_clickable(
                    (By.XPATH, "//span[contains(text(), 'Save as PDF')] | //div[contains(text(), 'Save as PDF')]")
                ))
                save_pdf.click()
            except Exception:
                print(f"  [X] 'Save as PDF' not found / 未找到 'Save as PDF'")
                result["status"] = "Failed"
                result["error"] = "'Save as PDF' option not found"
                write_csv_report(result)
                results.append(result)
                continue

            # 4. Select "Vector" option / 选择 "Vector" 选项
            try:
                print("     [UI] Waiting for popup and selecting Vector... / 等待导出弹窗并选择 Vector...")
                vector_option = wait_normal.until(EC.element_to_be_clickable(
                    (By.XPATH, "//label[contains(., 'Vector')] | //div[contains(text(), 'Vector')]")
                ))
                vector_option.click()
                print("     [UI] Selected Vector / 已选择 Vector")
            except Exception:
                print(f"  [X] 'Vector' option not found / 未找到 'Vector' 选项")
                result["status"] = "Failed"
                result["error"] = "'Vector' option not found (Maybe no frames?)"
                write_csv_report(result)
                results.append(result)
                continue

            # 5. Click Download (Export) button / 点击 Download (Export) 按钮
            try:
                print("     [UI] Clicking Export button... / 正在点击 Export 按钮...")
                export_btn = wait_normal.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(., 'Export')] | //button[contains(@class, 'button') and contains(., 'Export')]")
                ))
                export_btn.click()
                print(f"     [SUCCESS] Export command sent: {url} / 导出指令已发送")
                
                time.sleep(2)
                
            except Exception:
                print(f"  [X] Export button not found / 未找到 Export 按钮")
                result["status"] = "Failed"
                result["error"] = "Export button not found"
                write_csv_report(result)
                results.append(result)
                continue

            # 6. Wait and Click "Download file" / 等待并点击 "Download file"
            try:
                print("     [UI] Waiting for PDF generation (may take long)... / 等待 PDF 生成及下载按钮出现 (可能需要较长时间)...")
                # Use wait_long (10 mins) / 使用 wait_long (10分钟)
                download_file_btn = wait_long.until(EC.element_to_be_clickable(
                    (By.XPATH, "//button[contains(., 'Download file')] | //div[contains(text(), 'Download file')]")
                ))
                download_file_btn.click()
                print(f"     [SUCCESS] Clicked Download file button / 已点击 Download file 按钮，开始下载")
                result["status"] = "Success"
                
                time.sleep(5)
            except Exception:
                print(f"  [X] Download file button missing (Timeout) / 下载按钮未出现 (超时)")
                result["status"] = "Failed"
                result["error"] = "Download file button missing (Generation Timeout)"
                write_csv_report(result)
                results.append(result)
                continue

        except Exception as e:
            print(f"  [ERROR] Unexpected error: {e} / 发生意外错误: {e}")
            result["status"] = "Failed"
            result["error"] = f"Unexpected Error: {str(e)}"
            write_csv_report(result)
            results.append(result)
            continue
        
        # Write success / 成功的情况也要写入
        write_csv_report(result)
        results.append(result)
    
    return results

def write_csv_report(result):
    """Helper: Write single result to CSV / 辅助函数：将单条结果写入 CSV"""
    try:
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(REPORT_FILE, 'a', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow([timestamp, result["name"], result["url"], result["status"], result["error"]])
        print(f"     [Report] Written to CSV: {result['status']} / [报告] 已写入 CSV: {result['status']}")
    except Exception as e:
        print(f"     [Warning] Failed to write CSV: {e} / [警告] 写入 CSV 失败: {e}")

def main():
    print("="*50)
    print("       Miro Batch Export Script (Edge) / Miro 批量导出脚本 (Edge版)")
    print("="*50)
    driver = setup_driver()
    try:
        # 1. Link Management / 链接管理
        links = []
        if os.path.exists(LINK_FILE):
            print(f"\n[Info] Found local file {LINK_FILE}, reading links... / [提示] 发现本地文件 {LINK_FILE}，读取现有链接...")
            try:
                with open(LINK_FILE, 'r', encoding='utf-8') as f:
                    links = json.load(f)
            except json.JSONDecodeError:
                print(f"[Warning] {LINK_FILE} format error, will re-scrape... / [警告] {LINK_FILE} 文件格式有误，将重新抓取...")
                links = []
        else:
            print(f"\n[Info] {LINK_FILE} not found, starting fresh scrape... / [提示] 本地未发现 {LINK_FILE}，将执行全新抓取...")

        # Always try to update/supplement links / 始终尝试更新/补充链接
        print(f"\n[Info] Checking Dashboard for new boards... / [提示] 正在检查 Dashboard 以获取新白板...")
        links = scrape_dashboard_links(driver, existing_links=links)
            
        if not links:
            print("[Warning] No board links found, exiting. / [警告] 未找到任何白板链接，程序即将退出。")
            return

        print(f"\n>>> Ready to process {len(links)} boards... / 准备处理 {len(links)} 个白板...")
        
        # 2. Execute Download / 执行下载
        results = download_vector_pdf(driver, links)
        
        # 3. Generate Report / 生成报告
        print("\n" + "="*50)
        print("       Execution Summary / 执行报告 Summary")
        print("="*50)
        
        success_count = sum(1 for r in results if r['status'] == 'Success')
        fail_count = len(results) - success_count
        
        print(f"Total Tasks / 总任务数: {len(results)}")
        print(f"Success / 成功: {success_count}")
        print(f"Failed / 失败: {fail_count}")
        print(f"Detailed report saved to: {REPORT_FILE} / 详细报告已保存至: {REPORT_FILE}")
        
        if fail_count > 0:
            print("\n[Failure Details / 失败详情]")
            for r in results:
                if r['status'] == 'Failed':
                    print(f" - {r['name']} ({r['url']}): {r['error']}")
        
        print("="*50)

    except Exception as e:
        print(f"\n[!!!] Main Program Exception: {e} / 主程序发生异常: {e}")
    finally:
        print("\n" + "="*50)
        print("       Task Finished / 任务结束")
        print("="*50)
        # input("Press Enter to close browser... / 按 Enter 关闭浏览器...") 
        driver.quit()

if __name__ == "__main__":
    main()