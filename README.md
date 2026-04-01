# EZ_Rename - 批量重命名工具 / Bulk Rename Tool

![Python Version](https://img.shields.io/badge/python-3.6%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-lightgrey)

一个轻量、易用的 Windows 批量重命名工具，支持多种重命名模式、拖拽排序、撤销等功能，完全免费开源。

A lightweight and easy-to-use batch renaming tool for Windows, featuring multiple renaming modes, drag-and-drop sorting, undo, and more. Fully free and open source.

---

## 📸 截图 / Screenshot



---

## ✨ 功能特点 / Features

### 中文
- **三种重命名模式**：替换文本、添加文本、格式化（序号 + 日期）
- **实时预览**：输入规则后立即显示新文件名
- **文件列表管理**：支持拖拽添加文件/文件夹、拖拽勾选、分组排序
- **撤销功能**：支持重命名后撤销，也支持刷新、排序等操作的撤销
- **多语言支持**：简体中文 / English（自动跟随系统语言）
- **快捷方式集成**：可一键添加到“发送到”菜单、桌面或开始菜单
- **单实例运行**：双击文件时自动将路径发送到已运行窗口
- **文件占用检测**：重命名时如果文件被占用会给出明确提示
- **合法文件名检测**：自动检测并提示非法字符

### English
- **Three renaming modes**: Replace Text, Add Text, Format (sequence + date)
- **Live preview**: New file names are displayed instantly as you type
- **File list management**: Drag and drop to add files/folders, drag to toggle selection, group sorting
- **Undo support**: Undo renaming as well as sorting and refresh operations
- **Multi-language support**: Simplified Chinese / English (follows system language)
- **Shortcut integration**: One-click add to SendTo, Desktop or Start Menu
- **Single instance**: Dragging files to the exe will pass them to the running window
- **File busy detection**: Clear error message when a file is locked
- **Invalid character detection**: Automatically checks for illegal filename characters

---

## 🚀 下载与安装 / Download & Install

### 方式一：下载可执行文件（推荐） / Option 1: Download Executable (Recommended)

前往 [Releases](https://github.com/zbt00123/EZ_Rename/releases) 页面下载最新版本的 `EZ_Rename.exe`，无需安装，双击即可运行。

Go to the [Releases](https://github.com/your-username/EZ_Rename/releases) page and download the latest `EZ_Rename.exe`. No installation required – just double-click to run.

### 方式二：从源码运行 / Option 2: Run from Source

```bash
git clone https://github.com/zbt00123/EZ_Rename.git
cd EZ_Rename
pip install -r requirements.txt
python EZ_Rename.py
