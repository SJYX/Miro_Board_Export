import time
import json
import os
import csv
import datetime
import logging
from dataclasses import dataclass
from typing import List, Dict, Optional, Set, Any

from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ==================== Configuration ====================

@dataclass
class MiroConfig:
    """Configuration settings for Miro Export."""
    user_data_dir: str = r"C:\Users\112560\AppData\Local\Microsoft\Edge\User Data"
    profile_dir: str = "Default"
    link_file: str = "miro_board_links.json"
    report_file: str = "miro_export_report.csv"
    headless: bool = False
    log_level: int = logging.INFO

# ==================== Logging Setup ====================

def setup_logger(level: int = logging.INFO):
    """Configure structured logging."""
    logging.basicConfig(
        level=level,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%H:%M:%S',
        handlers=[
            logging.StreamHandler()
        ]
    )
    return logging.getLogger("MiroExport")

logger = setup_logger()

# ==================== CSV Report Handler ====================

class CsvReport:
    """Handles CSV report operations: Initialization, Reading, and Upserting."""

    HEADER = ["Timestamp", "Board Name", "URL", "Owner", "Status", "Error Message"]

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.initialize()

    def initialize(self):
        """Ensure CSV exists and has the correct header. Normalize if needed."""
        if not os.path.exists(self.filepath):
            self._create_new()
        else:
            self._normalize_header()

    def _create_new(self):
        try:
            with open(self.filepath, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                writer.writerow(self.HEADER)
            logger.info(f"Created new report file: {self.filepath}")
        except Exception as e:
            logger.error(f"Failed to create report file: {e}")

    def _normalize_header(self):
        """Check and migrate old CSV formats to the new 6-column format."""
        try:
            rows_to_keep = []
            needs_rewrite = False
            
            with open(self.filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                try:
                    header = next(reader)
                except StopIteration:
                    header = []

                if header != self.HEADER:
                    logger.info("Detected old or mismatching report format, normalizing...")
                    needs_rewrite = True

                for row in reader:
                    if not row: continue
                    
                    # Normalize to 6 columns
                    new_row = []
                    if len(row) == 4: # Old format (Timestamp, Name, URL, Status)
                        new_row = [row[0], "Unknown", row[1], "Unknown", row[2], row[3]]
                    elif len(row) == 5: # Previous format (Timestamp, Name, URL, Status, Error)
                        new_row = [row[0], row[1], row[2], "Unknown", row[3], row[4]]
                    elif len(row) >= 6:
                        new_row = row[:6]
                    else:
                        continue # Skip invalid
                    
                    rows_to_keep.append(new_row)

            if needs_rewrite:
                self._write_all(rows_to_keep)
                logger.info("Report normalized.")

        except Exception as e:
            logger.warning(f"Failed to process/normalize report file: {e}")

    def _write_all(self, rows: List[List[str]]):
        """Rewrite the entire CSV file."""
        with open(self.filepath, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(self.HEADER)
            writer.writerows(rows)
            f.flush()
            os.fsync(f.fileno())

    def get_successful_urls(self) -> Set[str]:
        """Return a set of URLs that have been successfully exported."""
        successful = set()
        if not os.path.exists(self.filepath):
            return successful

        try:
            with open(self.filepath, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                next(reader, None) # Skip header
                for row in reader:
                    # Check index 4 for Status (0:Time, 1:Name, 2:URL, 3:Owner, 4:Status, 5:Error)
                    if len(row) >= 6 and row[4] == "Success":
                        successful.add(row[2])
        except Exception:
            pass
        return successful

    def upsert_result(self, result: Dict[str, str]):
        """Update existing row or append new row based on URL."""
        try:
            rows = []
            if os.path.exists(self.filepath):
                with open(self.filepath, 'r', encoding='utf-8-sig') as f:
                    reader = csv.reader(f)
                    try:
                        next(reader) # Skip header
                        rows = list(reader)
                    except StopIteration:
                        pass

            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            updated = False
            # 0:Time, 1:Name, 2:URL, 3:Owner, 4:Status, 5:Error
            new_entry = [timestamp, result["name"], result["url"], result.get("owner", "Unknown"), result["status"], result["error"]]

            for i, row in enumerate(rows):
                if len(row) >= 3 and row[2] == result["url"]:
                    rows[i] = new_entry
                    updated = True
                    break
            
            if not updated:
                rows.append(new_entry)

            self._write_all(rows)
            action = "Updated" if updated else "Added"
            logger.info(f"[Report] {action} CSV: {result['status']}")

        except Exception as e:
            logger.error(f"Failed to write CSV: {e}")

# ==================== Miro Automator ====================

class MiroAutomator:
    """Main class for Miro automation logic."""

    def __init__(self, config: MiroConfig):
        self.config = config
        self.driver = None
        self.wait_normal = None
        self.wait_long = None

    def start_driver(self):
        """Initialize Edge driver with options."""
        logger.info("Starting Edge browser...")
        options = Options()
        options.add_argument(f"user-data-dir={self.config.user_data_dir}")
        options.add_argument(f"profile-directory={self.config.profile_dir}")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("detach", True)
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        options.add_argument("--start-maximized")
        options.add_argument("--log-level=3")
        
        # Stability
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-gpu")
        options.add_argument("--enable-unsafe-swiftshader")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument('--ignore-certificate-errors')

        try:
            self.driver = webdriver.Edge(service=Service(), options=options)
            self.wait_normal = WebDriverWait(self.driver, 20)
            self.wait_long = WebDriverWait(self.driver, 600) # 10 mins for export
            logger.info("Browser started successfully.")
        except Exception as e:
            logger.critical(f"Failed to start browser: {e}")
            raise

    def stop_driver(self):
        if self.driver:
            self.driver.quit()
            logger.info("Browser closed.")

    def _smart_wait(self, by: str, value: str, timeout: int = 5) -> bool:
        """Helper for optional element waiting."""
        try:
            WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((by, value)))
            return True
        except TimeoutException:
            return False

    def scrape_dashboard(self, existing_links: List[Dict] = None) -> List[Dict]:
        """Scrape board links from dashboard incrementally."""
        if existing_links is None:
            existing_links = []
        
        seen_urls = {item['url'] for item in existing_links if isinstance(item, dict)}
        # Handle old format strings
        for item in existing_links:
            if isinstance(item, str):
                seen_urls.add(item)

        logger.info("Opening Dashboard to scrape links...")
        self.driver.get("https://miro.com/app/dashboard/")

        # Wait for dashboard
        if not self._smart_wait(By.CSS_SELECTOR, "[data-testid='grid-view'], [data-testid='board-card'], a[href*='/app/board/']", 15):
            logger.warning("Timeout waiting for Dashboard, attempting to scroll anyway...")

        time.sleep(2) # Short buffer
        logger.info("Executing auto-scroll...")

        scraped_items_map = {}
        scrollable_container = self._find_scrollable_container()
        
        scroll_unchanged_count = 0
        max_scroll_unchanged = 5

        while True:
            # 1. JS Scrape
            self._js_scrape_visible_boards(scraped_items_map)
            logger.info(f"  -> Found {len(scraped_items_map)} boards so far...")

            # 2. Scroll
            if self._scroll_step(scrollable_container):
                scroll_unchanged_count = 0
            else:
                scroll_unchanged_count += 1
                logger.debug(f"Scroll unchanged {scroll_unchanged_count}/{max_scroll_unchanged}")
                
                # Try Page Down
                ActionChains(self.driver).send_keys(Keys.PAGE_DOWN).perform()
                time.sleep(0.5)

                if scroll_unchanged_count >= max_scroll_unchanged:
                    logger.info("Reached bottom.")
                    break
            
            time.sleep(0.5)

        # Consolidate
        new_items = []
        for url, item in scraped_items_map.items():
            if url not in seen_urls:
                new_items.append(item)
                seen_urls.add(url)
                logger.info(f"     [+] New: {item['name']} (Owner: {item.get('owner', 'Unknown')})")

        final_links = []
        # Keep old links
        for item in existing_links:
            if isinstance(item, str):
                final_links.append({"name": "Unknown (Old)", "url": item, "owner": "Unknown"})
            else:
                final_links.append(item)
        
        final_links.extend(new_items)
        
        logger.info(f"Scraping completed. Total: {len(final_links)} (New: {len(new_items)})")
        return final_links

    def _find_scrollable_container(self):
        try:
            return self.driver.find_element(By.CSS_SELECTOR, "[data-testid='grid-view']")
        except:
            # Fallback JS
            try:
                return self.driver.execute_script("""
                    let maxScroll = 0;
                    let maxDiv = null;
                    document.querySelectorAll('div').forEach(div => {
                        if (div.scrollHeight > div.clientHeight && div.scrollHeight > maxScroll) {
                            maxScroll = div.scrollHeight;
                            maxDiv = div;
                        }
                    });
                    return maxDiv;
                """)
            except:
                return None

    def _js_scrape_visible_boards(self, scraped_map):
        try:
            js_script = """
            return Array.from(document.querySelectorAll("a[href*='/app/board/']")).map(el => {
                let name = el.innerText.trim();
                if (!name) name = el.getAttribute("aria-label");
                if (!name) {
                    let titleEl = el.querySelector(".title, [class*='title']");
                    if (titleEl) name = titleEl.innerText.trim();
                }
                if (name && name.includes('\\n')) name = name.split('\\n')[0].trim();
                
                // Try to find Owner
                let owner = "Unknown";
                try {
                    // Strategy 1: List View (Row)
                    let row = el.closest('[role="row"]');
                    if (row) {
                        // In list view, Owner is typically in a specific column.
                        // We can look for text that looks like a name, or specific class.
                        // Assuming standard grid cells:
                        let cells = row.querySelectorAll('[role="gridcell"]');
                        if (cells.length >= 5) {
                             // Try the last few cells for owner name
                             // Usually: Name, Users, Project, ..., Last Opened, Owner, Actions
                             // Let's try to get text from the cell before the actions menu
                             // Or just grab all text and guess.
                             
                             // Better: Look for an element that is NOT the date and NOT the name.
                             // Let's assume it's the 6th column as per user image (index 5)
                             if (cells[5]) {
                                owner = cells[5].innerText.trim();
                             } else if (cells.length > 2) {
                                // Fallback: try the last text cell
                                owner = cells[cells.length - 2].innerText.trim();
                             }
                        }
                    }
                    
                    // Strategy 2: Grid View (Card)
                    if (owner === "Unknown") {
                        let card = el.closest('[data-testid="board-card"]');
                        if (card) {
                            // In card view, owner might be in footer
                            let footer = card.querySelector('[class*="footer"], [class*="bottom"]');
                            if (footer) owner = footer.innerText.trim();
                        }
                    }
                } catch (e) {
                    // Ignore owner extraction errors
                }

                return {
                    url: el.href,
                    name: name || "Untitled Board",
                    owner: owner || "Unknown"
                };
            });
            """
            items = self.driver.execute_script(js_script)
            for item in items:
                href = item.get('url')
                if href and len(href) > 10 and href not in scraped_map:
                    scraped_map[href] = item
        except Exception as e:
            logger.warning(f"Minor scraping error: {e}")

    def _scroll_step(self, container) -> bool:
        """Returns True if scroll position changed."""
        try:
            if container:
                current = self.driver.execute_script("return arguments[0].scrollTop", container)
                self.driver.execute_script("arguments[0].scrollBy(0, 400)", container)
                new_pos = self.driver.execute_script("return arguments[0].scrollTop", container)
                return abs(new_pos - current) > 5
            else:
                current = self.driver.execute_script("return window.pageYOffset || document.documentElement.scrollTop")
                self.driver.execute_script("window.scrollBy(0, 400);")
                new_pos = self.driver.execute_script("return window.pageYOffset || document.documentElement.scrollTop")
                return abs(new_pos - current) > 5
        except:
            # Fallback
            self.driver.execute_script("window.scrollBy(0, 400);")
            return True

    def batch_export(self, links: List[Dict], report: CsvReport):
        """Process all links for export."""
        successful_urls = report.get_successful_urls()
        
        logger.info(f"Starting batch export for {len(links)} boards...")
        
        for index, item in enumerate(links):
            # Normalize item
            if isinstance(item, str):
                url = item
                name = "Unknown"
                owner = "Unknown"
            else:
                url = item.get('url')
                name = item.get('name', 'Unknown')
                owner = item.get('owner', 'Unknown')

            if url in successful_urls:
                logger.info(f"[{index+1}/{len(links)}] Skipping (Already Exported): {name}")
                continue

            logger.info(f"[{index+1}/{len(links)}] Processing: {name} (Owner: {owner})")
            
            result = self._export_single_board(url, name, owner)
            report.upsert_result(result)

    def _export_single_board(self, url: str, name: str, owner: str) -> Dict[str, str]:
        result = {"name": name, "url": url, "owner": owner, "status": "Pending", "error": ""}
        
        try:
            self.driver.get(url)
            
            # 1. Optimized Board Load: Directly check for UI elements
            # User requested to skip waiting for full page/canvas load
            # We will rely on the "Share" button check or Menu button check to act as our "wait"
            
            self._dismiss_popups()

            # 2. Permission Check (Acts as the primary wait for board interactivity)
            # If "Share" button appears, the UI is ready enough for us to proceed.
            if not self._check_permissions():
                result["status"] = "Failed"
                result["error"] = "Insufficient permissions (Share button missing)"
                return result

            # 3. Open Export Menu
            if not self._open_export_menu():
                result["status"] = "Failed"
                result["error"] = "Could not open Export menu"
                return result

            # 4. Click Save as PDF
            if not self._click_save_as_pdf():
                result["status"] = "Failed"
                result["error"] = "'Save as PDF' not found"
                return result

            # 5. Check "Need 1 frame" popup
            if self._check_no_frame_popup():
                result["status"] = "Failed"
                result["error"] = "No frames to export"
                return result

            # 6. Select Vector
            if not self._select_vector_option():
                result["status"] = "Failed"
                result["error"] = "'Vector' option not found"
                return result

            # 7. Click Export
            if not self._click_export_button():
                result["status"] = "Failed"
                result["error"] = "Export button not found"
                return result

            # 8. Wait for Download
            if self._wait_for_download():
                result["status"] = "Success"
            else:
                result["status"] = "Failed"
                result["error"] = "Download timeout"

        except Exception as e:
            logger.error(f"Unexpected error processing {url}: {e}")
            result["status"] = "Failed"
            result["error"] = f"Unexpected: {str(e)}"

        return result

    def _dismiss_popups(self):
        try:
            ActionChains(self.driver).send_keys("\ue00c").perform() # ESC
        except:
            pass

    def _check_permissions(self) -> bool:
        try:
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Share')] | //div[contains(text(), 'Share')]"))
            )
            return True
        except:
            return False

    def _open_export_menu(self) -> bool:
        try:
            # Find Main Menu
            menu_btn = None
            try:
                menu_btn = self.driver.find_element(By.XPATH, "//button[@aria-label='Main menu'] | //div[@aria-label='Main menu']")
            except:
                menu_btn = self.driver.find_element(By.CSS_SELECTOR, "[data-testid='board-header__main-menu-button']")
            
            menu_btn.click()
            
            # Wait for Board option
            try:
                self.wait_normal.until(EC.visibility_of_element_located((By.XPATH, "//*[normalize-space(text())='Board']")))
            except:
                time.sleep(1)

            # Hover Board -> Find Export
            for _ in range(3):
                board_opts = self.driver.find_elements(By.XPATH, "//*[normalize-space(text())='Board']")
                for opt in board_opts:
                    if opt.is_displayed():
                        ActionChains(self.driver).move_to_element(menu_btn).pause(0.2).move_to_element(opt).perform()
                        
                        # Check for Export
                        try:
                            self.wait_normal.until(EC.visibility_of_element_located((By.XPATH, "//*[normalize-space(text())='Export']")))
                        except:
                            time.sleep(0.5)

                        export_opts = self.driver.find_elements(By.XPATH, "//*[normalize-space(text())='Export']")
                        for ex_opt in export_opts:
                            if ex_opt.is_displayed():
                                ActionChains(self.driver).move_to_element(ex_opt).perform()
                                return True
                time.sleep(1)
            return False
        except Exception as e:
            logger.debug(f"Menu navigation failed: {e}")
            return False

    def _click_save_as_pdf(self) -> bool:
        try:
            time.sleep(1)
            btn = self.wait_normal.until(EC.element_to_be_clickable(
                (By.XPATH, "//span[contains(text(), 'Save as PDF')] | //div[contains(text(), 'Save as PDF')]")
            ))
            btn.click()
            return True
        except:
            return False

    def _check_no_frame_popup(self) -> bool:
        try:
            time.sleep(1)
            WebDriverWait(self.driver, 3).until(
                EC.presence_of_element_located((By.XPATH, "//*[contains(., 'You need at least 1 visible frame')] | //*[contains(., 'at least 1 visible frame')]"))
            )
            self._dismiss_popups()
            return True
        except:
            return False

    def _select_vector_option(self) -> bool:
        try:
            btn = self.wait_normal.until(EC.element_to_be_clickable(
                (By.XPATH, "//label[contains(., 'Vector')] | //div[contains(text(), 'Vector')]")
            ))
            btn.click()
            return True
        except:
            return False

    def _click_export_button(self) -> bool:
        try:
            btn = self.wait_normal.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'Export')] | //button[contains(@class, 'button') and contains(., 'Export')]")
            ))
            btn.click()
            
            # Smart wait for download button
            try:
                WebDriverWait(self.driver, 5).until(EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Download file')] | //div[contains(text(), 'Download file')]")))
            except:
                time.sleep(2)
            return True
        except:
            return False

    def _wait_for_download(self) -> bool:
        try:
            logger.info("Waiting for PDF generation...")
            btn = self.wait_long.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(., 'Download file')] | //div[contains(text(), 'Download file')]")
            ))
            btn.click()
            time.sleep(5) # Allow download to start
            return True
        except:
            return False

# ==================== Main ====================

def main():
    config = MiroConfig()
    report = CsvReport(config.report_file)
    automator = MiroAutomator(config)

    try:
        automator.start_driver()

        # 1. Load Local Links
        links = []
        if os.path.exists(config.link_file):
            logger.info(f"Reading local links from {config.link_file}...")
            try:
                with open(config.link_file, 'r', encoding='utf-8') as f:
                    links = json.load(f)
            except:
                logger.warning("Failed to read local links.")

        # 2. Scrape New Links
        links = automator.scrape_dashboard(existing_links=links)
        
        # Save updated links
        with open(config.link_file, 'w', encoding='utf-8') as f:
            json.dump(links, f, indent=4, ensure_ascii=False)

        if not links:
            logger.warning("No links found.")
            return

        # 3. Batch Export
        automator.batch_export(links, report)

    except Exception as e:
        logger.critical(f"Main execution failed: {e}")
    finally:
        automator.stop_driver()

if __name__ == "__main__":
    main()