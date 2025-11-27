# Miro Board Export Script / Miro 白板批量导出脚本

A Python script to automatically scrape all Miro board links from your dashboard and export them as Vector PDFs.
一个用于自动抓取 Dashboard 所有 Miro 白板链接并批量导出为矢量 PDF 的 Python 脚本。

## Features / 功能

- **Auto-Scraping / 自动抓取**: Automatically scrolls through the Miro dashboard to capture all board links, handling virtual scrolling. / 自动滚动 Miro 仪表板以捕获所有白板链接，支持虚拟滚动处理。
- **Incremental Update / 增量更新**: Saves links to `miro_board_links.json` and only adds new ones in subsequent runs. / 将链接保存到 `miro_board_links.json`，后续运行仅添加新链接。
- **Batch Export / 批量导出**: Automates the "Export -> Save as PDF -> Vector" flow for each board. / 自动化每个白板的 "Export -> Save as PDF -> Vector" 流程。
- **Smart Retry / 智能重试**: Handles network jitters and UI loading delays with built-in retry logic. / 内置重试逻辑，处理网络抖动和 UI 加载延迟。
- **Reporting / 报告生成**: Generates a CSV report (`miro_export_report.csv`) tracking the status of each export. / 生成 CSV 报告 (`miro_export_report.csv`) 跟踪每个导出的状态。

## Prerequisites / 前置要求

- **OS**: Windows (Script is optimized for Windows paths). / Windows (脚本针对 Windows 路径优化)。
- **Browser**: Microsoft Edge. / Microsoft Edge 浏览器。
- **Python**: Python 3.x installed. / 已安装 Python 3.x。
- **Selenium**: `pip install selenium`

## Installation / 安装

1.  Clone this repository / 克隆此仓库:
    ```bash
    git clone https://github.com/SJYX/Miro_Board_Export.git
    cd Miro_Board_Export
    ```

2.  Install dependencies / 安装依赖:
    ```bash
    pip install selenium
    ```

## Configuration / 配置

Open `Miro_Board_Export.py` and modify the following variables at the top:
打开 `Miro_Board_Export.py` 并修改顶部的以下变量：

```python
# 1. Edge User Data Path (Ensure Edge is completely closed)
# 1. Edge User Data 路径 (请确保 Edge 已完全关闭)
USER_DATA_DIR = r"C:\Users\YourUser\AppData\Local\Microsoft\Edge\User Data"

# 2. Profile Directory, usually "Default"
# 2. 配置文件目录，通常是 "Default"
PROFILE_DIR = "Default"
```

> **Note / 注意**: You must close all Edge browser windows before running the script, as it needs to attach to your user profile to reuse your login session.
> **注意**: 运行脚本前必须关闭所有 Edge 浏览器窗口，因为脚本需要加载您的用户配置文件以复用登录状态。

## Usage / 使用方法

Run the script:
运行脚本：

```bash
python Miro_Board_Export.py
```

The script will:
脚本将：
1.  Open Edge and navigate to your Miro dashboard. / 打开 Edge 并导航到 Miro 仪表板。
2.  Scroll and collect all board links. / 滚动并收集所有白板链接。
3.  Visit each board and export it as a Vector PDF. / 访问每个白板并将其导出为矢量 PDF。
4.  Save the PDF to your default download folder. / 将 PDF 保存到您的默认下载文件夹。

## Troubleshooting / 故障排除

- **Browser fails to start / 浏览器启动失败**: Ensure all Edge windows are closed. Check if `msedgedriver.exe` matches your Edge version (Selenium usually handles this automatically). / 确保所有 Edge 窗口已关闭。检查 `msedgedriver.exe` 是否与您的 Edge 版本匹配（Selenium 通常会自动处理）。
- **Missing Boards / 白板遗漏**: The script uses incremental scrolling. If boards are still missing, try increasing the wait times in the script. / 脚本使用增量滚动。如果仍有白板遗漏，请尝试增加脚本中的等待时间。
- **Export Menu not found / 找不到导出菜单**: Miro's UI might change. The script looks for "Export" in the main menu. / Miro 的 UI 可能会变动。脚本会在主菜单中寻找 "Export"。

## License / 许可证

MIT License
