# Miro Board Export Script

[English Version](#miro-whiteboard-batch-export-script)

A Python script to automatically scrape all Miro board links from your dashboard and export them as Vector PDFs.
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
2.  Scroll and collect all board links.
3.  Visit each board and export it as a Vector PDF.
4.  Save the PDF to your default download folder.

## Troubleshooting

- **Browser fails to start**: Ensure all Edge windows are closed. Check if `msedgedriver.exe` matches your Edge version (Selenium usually handles this automatically).
- **Missing Boards**: The script uses incremental scrolling. If boards are still missing, try increasing the wait times in the script.
- **Export Menu not found**: Miro's UI might change. The script looks for "Export" in the main menu.

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
- **智能重试**: 内置重试逻辑，处理网络抖动和 UI 加载延迟。
- **断点续传**: 跳过已成功导出的 Board (基于 `miro_export_report.csv`)。
- **报告生成**: 生成 CSV 报告 (`miro_export_report.csv`) 跟踪每个导出的状态。

## 限制 (Limitations)

- **权限**: 如果你没有 Board 的权限，则无法导出。
- **Frame**: 如果 Board 里没有 Frame 也无法导出。

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
2.  滚动并收集所有 Board 链接。
3.  访问每个 Board 并将其导出为矢量 PDF。
4.  将 PDF 保存到您的默认下载文件夹。

## 故障排除 (Troubleshooting)

- **浏览器启动失败**: 确保所有 Edge 窗口已关闭。检查 `msedgedriver.exe` 是否与您的 Edge 版本匹配（Selenium 通常会自动处理）。
- **Board 遗漏**: 脚本使用增量滚动。如果仍有 Board 遗漏，请尝试增加脚本中的等待时间。
- **找不到导出菜单**: Miro 的 UI 可能会变动。脚本会在主菜单中寻找 "Export"。

## 许可证 (License)

MIT License
