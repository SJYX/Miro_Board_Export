# Miro Board Export Script

[English Version](#miro-whiteboard-batch-export-script)

A Python script to automatically scrape all Miro board links from your dashboard and export them as Vector PDFs.

## Features

- **Automatic Scraping**: Automatically scrolls the Miro Dashboard to capture all Board links, supporting virtual scrolling.
- **Incremental Update**: Saves links to `miro_board_links.json`, subsequent runs only add new links.
- **Batch Export**: Automates the "Export -> Save as PDF -> Vector" flow for each board.
- **Smart Waits**: Uses dynamic `WebDriverWait` instead of fixed sleeps for faster and more reliable execution.
- **Permission Check**: Automatically checks for the "Share" button to verify permissions before attempting export.
- **Robust Reporting**: Generates a CSV report (`miro_export_report.csv`) that updates existing entries (Upsert) instead of creating duplicates.
- **Popup Handling**: Detects and handles "Need at least 1 visible frame" popups.
- **Resume Capability**: Skips boards that have already been successfully exported.

## Prerequisites

- **OS**: Windows (Script is optimized for Windows paths).
- **Browser**: Microsoft Edge.
- **Python**: Python 3.x installed.
- **Selenium**: `pip install selenium`

## Installation

1.  Clone this repository:
    ```bash
    git clone https://github.com/SJYX/Miro_Board_Export.git
    cd Miro_Board_Export
    ```

2.  Install dependencies:
    ```bash
    pip install selenium
    ```

## Configuration

Open `Miro_Board_Export.py` and modify the following variables at the top:

```python
# 1. Edge User Data Path (Ensure Edge is completely closed)
USER_DATA_DIR = r"C:\Users\YourUser\AppData\Local\Microsoft\Edge\User Data"

# 2. Profile Directory, usually "Default"
PROFILE_DIR = "Default"
```

> **Note**: You must close all Edge browser windows before running the script, as it needs to attach to your user profile to reuse your login session.

## Usage

Run the script:

```bash
python Miro_Board_Export.py
```

The script will:
1.  Open Edge and navigate to your Miro dashboard.
2.  Scroll and collect all board links (incremental).
3.  Visit each board, check permissions, and export it as a Vector PDF.
4.  Save the PDF to your default download folder.
5.  Update the `miro_export_report.csv` with the status of each export.

## Troubleshooting

- **Browser fails to start**: Ensure all Edge windows are closed. Check if `msedgedriver.exe` matches your Edge version (Selenium usually handles this automatically).
- **Insufficient Permissions**: If a board is skipped with this error, it means the account lacks "Share" permissions (Edit/Owner access) for that board.
- **Need at least 1 visible frame**: The board has no frames to export.
- **Missing Boards**: The script uses incremental scrolling. If boards are still missing, try increasing the wait times in the script.

## License

MIT License

---

# Miro Whiteboard Batch Export Script

[中文版本](#miro-board-export-script)

一个用于自动抓取 Dashboard 所有 Miro Board 链接并批量导出为矢量 PDF 的 Python 脚本。

## 功能 (Features)

- **自动抓取**: 自动滚动 Miro Dashboard 以捕获所有 Board 链接，支持虚拟滚动处理。
- **增量更新**: 将链接保存到 `miro_board_links.json`，后续运行仅添加新链接。
- **批量导出**: 自动化每个 Board 的 "Export -> Save as PDF -> Vector" 流程。
- **智能等待**: 使用动态 `WebDriverWait` 替代固定等待，执行更快速、更稳定。
- **权限检查**: 在导出前自动检查 "Share" 按钮以验证权限。
- **智能报告**: 生成 CSV 报告 (`miro_export_report.csv`)，支持 Upsert (更新现有记录)，避免重复数据。
- **弹窗处理**: 自动检测并处理 "Need at least 1 visible frame" 弹窗。
- **断点续传**: 跳过已成功导出的 Board。

## 前置要求 (Prerequisites)

- **OS**: Windows (脚本针对 Windows 路径优化)。
- **Browser**: Microsoft Edge 浏览器。
- **Python**: 已安装 Python 3.x。
- **Selenium**: `pip install selenium`

## 安装 (Installation)

1.  克隆此仓库:
    ```bash
    git clone https://github.com/SJYX/Miro_Board_Export.git
    cd Miro_Board_Export
    ```

2.  安装依赖:
    ```bash
    pip install selenium
    ```

## 配置 (Configuration)

打开 `Miro_Board_Export.py` 并修改顶部的以下变量：

```python
# 1. Edge User Data 路径 (请确保 Edge 已完全关闭)
USER_DATA_DIR = r"C:\Users\YourUser\AppData\Local\Microsoft\Edge\User Data"

# 2. 配置文件目录，通常是 "Default"
PROFILE_DIR = "Default"
```

> **注意**: 运行脚本前必须关闭所有 Edge 浏览器窗口，因为脚本需要加载您的用户配置文件以复用登录状态。

## 使用方法 (Usage)

运行脚本：

```bash
python Miro_Board_Export.py
```

脚本将：
1.  打开 Edge 并导航到 Miro Dashboard。
2.  滚动并收集所有 Board 链接 (增量)。
3.  访问每个 Board，检查权限，并将其导出为矢量 PDF。
4.  将 PDF 保存到您的默认下载文件夹。
5.  更新 `miro_export_report.csv` 记录导出状态。

## 故障排除 (Troubleshooting)

- **浏览器启动失败**: 确保所有 Edge 窗口已关闭。检查 `msedgedriver.exe` 是否与您的 Edge 版本匹配（Selenium 通常会自动处理）。
- **权限不足 (Insufficient Permissions)**: 如果 Board 被跳过并显示此错误，说明当前账户没有该 Board 的 "Share" 权限 (编辑/所有者权限)。
- **需要至少 1 个可见 Frame**: 该 Board 没有 Frame 可供导出。
- **Board 遗漏**: 脚本使用增量滚动。如果仍有 Board 遗漏，请尝试增加脚本中的等待时间。

## 许可证 (License)

MIT License
