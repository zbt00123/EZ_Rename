#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
EZ_Rename - 批量重命名工具 (多语言版)
Windows 10/11 桌面应用，单文件可执行 (Python 3 + Tkinter)
支持分组排序、拖拽勾选、整理排序、分隔符、撤销等
"""

import os
import sys
import json
import threading
import queue
import pickle
import socket
import ctypes
import tempfile
import subprocess
import copy
import uuid
import locale
import platform
from datetime import datetime
from tkinter import Tk, messagebox, filedialog
from tkinter import ttk, StringVar, IntVar, BooleanVar, DISABLED, NORMAL
import tkinter as tk

# 尝试导入拖拽支持库
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False
    class TkinterDnD:
        class Tk(Tk):
            pass

# ---------- 全局配置 ----------
CONFIG_FILE = os.path.join(os.environ['USERPROFILE'], '.ez_rename_config.json')
SOCKET_PORT_FILE = os.path.join(tempfile.gettempdir(), '.ez_rename_port')
MUTEX_NAME = "Global\\EZRenameTool_3A4F6B81"

# Windows 非法字符集
INVALID_CHARS = r'\/:*?"<>|'

# ---------- 快捷方式管理函数 ----------
def create_shortcut(target_path, shortcut_path, description=""):
    """创建快捷方式（如果目标路径已存在则跳过）"""
    if os.path.exists(shortcut_path):
        return False
    try:
        from win32com.client import Dispatch
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = target_path
        shortcut.WorkingDirectory = os.path.dirname(target_path)
        shortcut.IconLocation = target_path + ',0'
        if description:
            shortcut.Description = description
        shortcut.save()
        return True
    except:
        return False

def delete_shortcut(shortcut_path):
    """删除快捷方式"""
    try:
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)
            return True
    except:
        pass
    return False

def get_exe_path():
    """获取当前可执行文件的路径（打包后为 exe 路径）"""
    if getattr(sys, 'frozen', False):
        return sys.executable
    else:
        return __file__

def get_sendto_folder():
    """获取 Windows 发送到文件夹路径"""
    try:
        import winshell
        return winshell.folder('sendto')
    except:
        return os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'SendTo')

def get_desktop_folder():
    """获取桌面文件夹路径"""
    try:
        import winshell
        return winshell.desktop()
    except:
        return os.path.join(os.environ['USERPROFILE'], 'Desktop')

def get_startmenu_folder():
    """获取开始菜单程序文件夹路径"""
    return os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs')

def ensure_sendto_shortcut(enable):
    """根据 enable 状态添加或删除发送到快捷方式"""
    sendto = get_sendto_folder()
    shortcut_path = os.path.join(sendto, 'EZ_Rename.lnk')
    exe_path = get_exe_path()
    if enable:
        return create_shortcut(exe_path, shortcut_path, "EZ_Rename")
    else:
        return delete_shortcut(shortcut_path)

def add_desktop_shortcut():
    """添加到桌面快捷方式"""
    desktop = get_desktop_folder()
    shortcut_path = os.path.join(desktop, 'EZ_Rename.lnk')
    exe_path = get_exe_path()
    if create_shortcut(exe_path, shortcut_path, "EZ_Rename"):
        return True, "已添加到桌面"
    else:
        return False, "快捷方式已存在或创建失败"

def add_startmenu_shortcut():
    """固定到开始菜单（添加到开始菜单程序文件夹）"""
    startmenu = get_startmenu_folder()
    shortcut_path = os.path.join(startmenu, 'EZ_Rename.lnk')
    exe_path = get_exe_path()
    if create_shortcut(exe_path, shortcut_path, "EZ_Rename"):
        return True, "已添加到开始菜单"
    else:
        return False, "快捷方式已存在或创建失败"

def delete_startmenu_shortcut():
    """删除开始菜单中的快捷方式"""
    startmenu = get_startmenu_folder()
    shortcut_path = os.path.join(startmenu, 'EZ_Rename.lnk')
    return delete_shortcut(shortcut_path)

def pin_to_startmenu(shortcut_path):
    """将快捷方式固定到开始屏幕（Win10/Win11）"""
    try:
        from win32com.client import Dispatch
        shell = Dispatch('Shell.Application')
        namespace = shell.NameSpace(os.path.dirname(shortcut_path))
        item = namespace.ParseName(os.path.basename(shortcut_path))
        item.InvokeVerb('pintostartscreen')
        return True
    except:
        return False

def unpin_from_startmenu(shortcut_path):
    """尝试取消固定开始屏幕中的快捷方式"""
    try:
        from win32com.client import Dispatch
        shell = Dispatch('Shell.Application')
        namespace = shell.NameSpace(os.path.dirname(shortcut_path))
        item = namespace.ParseName(os.path.basename(shortcut_path))
        # 尝试取消固定动词（可能不存在）
        item.InvokeVerb('unpinfromstart')
        return True
    except:
        # 如果失败，尝试删除固定位置的快捷方式（复杂，此处简化）
        return False

def open_file_location(path):
    """打开文件或文件夹所在位置（Windows）"""
    if os.path.isdir(path):
        subprocess.Popen(['explorer', path])
    else:
        subprocess.Popen(['explorer', '/select,', path])

def is_valid_filename(name):
    """检查文件名是否包含非法字符"""
    return not any(c in INVALID_CHARS for c in name)

def get_invalid_chars_in_name(name):
    """返回文件名中的非法字符集合"""
    return set(c for c in name if c in INVALID_CHARS)

# ---------- 翻译字典 ----------
class Translator:
    def __init__(self):
        self.lang = "auto"  # auto, zh_CN, en_US
        self.strings = {}
        self._load_strings()

    def _load_strings(self):
        self.strings = {
            # 通用
            "app_title": {
                "zh_CN": "EZ_Rename - 批量重命名工具",
                "en_US": "EZ_Rename - Bulk Rename Tool"
            },
            "status_ready": {
                "zh_CN": "就绪",
                "en_US": "Ready"
            },
            "status_renaming": {
                "zh_CN": "正在重命名...",
                "en_US": "Renaming..."
            },
            "status_rename_done": {
                "zh_CN": "重命名完成: 成功 {} 个, 失败 {} 个",
                "en_US": "Rename completed: {} succeeded, {} failed"
            },
            "status_undo_success": {
                "zh_CN": "撤销成功",
                "en_US": "Undo succeeded"
            },
            "status_added_to_sendto": {
                "zh_CN": "已添加到“发送到”文件夹",
                "en_US": "Added to SendTo folder"
            },
            "status_removed_from_sendto": {
                "zh_CN": "已从“发送到”文件夹移除",
                "en_US": "Removed from SendTo folder"
            },
            "status_added_to_desktop": {
                "zh_CN": "已添加到桌面",
                "en_US": "Added to Desktop"
            },
            "status_already_exists": {
                "zh_CN": "快捷方式已存在",
                "en_US": "Shortcut already exists"
            },
            "status_added_to_startmenu": {
                "zh_CN": "已添加到开始菜单",
                "en_US": "Added to Start Menu"
            },
            "status_removed_from_startmenu": {
                "zh_CN": "已从开始菜单移除",
                "en_US": "Removed from Start Menu"
            },
            "status_pinned_to_startmenu": {
                "zh_CN": "已固定到开始屏幕",
                "en_US": "Pinned to Start"
            },
            "status_unpinned_from_startmenu": {
                "zh_CN": "已从开始屏幕取消固定",
                "en_US": "Unpinned from Start"
            },
            "error_add_to_sendto": {
                "zh_CN": "无法创建快捷方式，请检查权限或手动添加。",
                "en_US": "Cannot create shortcut, please check permissions or add manually."
            },

            # 菜单
            "menu_file": {
                "zh_CN": "文件",
                "en_US": "File"
            },
            "menu_import_files": {
                "zh_CN": "导入文件",
                "en_US": "Import Files"
            },
            "menu_import_folder": {
                "zh_CN": "导入文件夹",
                "en_US": "Import Folder"
            },
            "menu_exit": {
                "zh_CN": "退出",
                "en_US": "Exit"
            },
            "menu_edit": {
                "zh_CN": "编辑",
                "en_US": "Edit"
            },
            "menu_undo": {
                "zh_CN": "撤销",
                "en_US": "Undo"
            },
            "menu_select_all": {
                "zh_CN": "全选",
                "en_US": "Select All"
            },
            "menu_invert_selection": {
                "zh_CN": "反选",
                "en_US": "Invert Selection"
            },
            "menu_clear_selection": {
                "zh_CN": "取消选择",
                "en_US": "Clear Selection"
            },
            "menu_refresh": {
                "zh_CN": "刷新",
                "en_US": "Refresh"
            },
            "menu_clear_list": {
                "zh_CN": "清空列表",
                "en_US": "Clear List"
            },
            "menu_remove_selected": {
                "zh_CN": "移除选中",
                "en_US": "Remove Selected"
            },
            "menu_help": {
                "zh_CN": "帮助",
                "en_US": "Help"
            },
            "menu_language": {
                "zh_CN": "语言",
                "en_US": "Language"
            },
            "menu_lang_auto": {
                "zh_CN": "根据系统选择",
                "en_US": "Follow System"
            },
            "menu_lang_zh": {
                "zh_CN": "简体中文",
                "en_US": "Simplified Chinese"
            },
            "menu_lang_en": {
                "zh_CN": "English",
                "en_US": "English"
            },
            "menu_add_to_sendto": {
                "zh_CN": "添加到“发送到”",
                "en_US": "Add to SendTo"
            },
            "menu_add_to_desktop": {
                "zh_CN": "添加到桌面",
                "en_US": "Add to Desktop"
            },
            "menu_add_to_startmenu": {
                "zh_CN": "固定到开始菜单",
                "en_US": "Pin to Start Menu"
            },
            "menu_about": {
                "zh_CN": "关于",
                "en_US": "About"
            },

            # 按钮
            "btn_add_files": {
                "zh_CN": "添加文件",
                "en_US": "Add Files"
            },
            "btn_add_folder": {
                "zh_CN": "添加文件夹",
                "en_US": "Add Folder"
            },
            "btn_clear_list": {
                "zh_CN": "清空列表",
                "en_US": "Clear List"
            },
            "btn_remove_selected": {
                "zh_CN": "移除选中",
                "en_US": "Remove Selected"
            },
            "btn_select_all": {
                "zh_CN": "全选",
                "en_US": "Select All"
            },
            "btn_invert": {
                "zh_CN": "反选",
                "en_US": "Invert"
            },
            "btn_sort": {
                "zh_CN": "整理",
                "en_US": "Sort"
            },
            "btn_refresh": {
                "zh_CN": "刷新",
                "en_US": "Refresh"
            },
            "btn_rename": {
                "zh_CN": "执行重命名",
                "en_US": "Rename"
            },
            "btn_undo": {
                "zh_CN": "撤销",
                "en_US": "Undo"
            },

            # 标签
            "label_mode": {
                "zh_CN": "模式:",
                "en_US": "Mode:"
            },
            "mode_replace": {
                "zh_CN": "替换文本",
                "en_US": "Replace Text"
            },
            "mode_add": {
                "zh_CN": "添加文本",
                "en_US": "Add Text"
            },
            "mode_format": {
                "zh_CN": "格式",
                "en_US": "Format"
            },
            "label_find": {
                "zh_CN": "查找:",
                "en_US": "Find:"
            },
            "label_replace_with": {
                "zh_CN": "替换为:",
                "en_US": "Replace with:"
            },
            "label_add_text": {
                "zh_CN": "添加文本:",
                "en_US": "Add Text:"
            },
            "label_position": {
                "zh_CN": "位置:",
                "en_US": "Position:"
            },
            "position_start": {
                "zh_CN": "开头",
                "en_US": "Start"
            },
            "position_end": {
                "zh_CN": "末尾",
                "en_US": "End"
            },
            "label_base_name": {
                "zh_CN": "基础名称:",
                "en_US": "Base Name:"
            },
            "label_separator": {
                "zh_CN": "分隔符:",
                "en_US": "Separator:"
            },
            "label_start": {
                "zh_CN": "起始:",
                "en_US": "Start:"
            },
            "label_step": {
                "zh_CN": "步长:",
                "en_US": "Step:"
            },
            "label_pad": {
                "zh_CN": "位数:",
                "en_US": "Digits:"
            },
            "label_date_format": {
                "zh_CN": "日期格式:",
                "en_US": "Date Format:"
            },
            "label_seq_position": {
                "zh_CN": "序号位置:",
                "en_US": "Seq Position:"
            },
            "seq_prefix": {
                "zh_CN": "前缀",
                "en_US": "Prefix"
            },
            "seq_suffix": {
                "zh_CN": "后缀",
                "en_US": "Suffix"
            },
            "date_none": {
                "zh_CN": "无",
                "en_US": "None"
            },
            "date_ymd": {
                "zh_CN": "yyyy-mm-dd",
                "en_US": "yyyy-mm-dd"
            },
            "date_yyyymmdd": {
                "zh_CN": "yyyymmdd",
                "en_US": "yyyymmdd"
            },

            # 文件列表列标题
            "col_check": {
                "zh_CN": "✔",
                "en_US": "✔"
            },
            "col_original": {
                "zh_CN": "原文件名",
                "en_US": "Original Name"
            },
            "col_new": {
                "zh_CN": "新文件名",
                "en_US": "New Name"
            },
            "col_status": {
                "zh_CN": "状态",
                "en_US": "Status"
            },
            "col_location": {
                "zh_CN": "位置",
                "en_US": "Location"
            },

            # 状态文本
            "status_modified": {
                "zh_CN": "已修改",
                "en_US": "Modified"
            },
            "status_ready": {
                "zh_CN": "就绪",
                "en_US": "Ready"
            },

            # 消息框
            "msg_no_rename": {
                "zh_CN": "没有需要重命名的勾选文件。",
                "en_US": "No checked files need renaming."
            },
            "msg_invalid_chars_title": {
                "zh_CN": "非法字符",
                "en_US": "Invalid Characters"
            },
            "msg_invalid_chars_body": {
                "zh_CN": "以下文件名包含非法字符，无法重命名：\n{}\n\n请修改后重试。",
                "en_US": "The following filenames contain invalid characters and cannot be renamed:\n{}\n\nPlease modify and try again."
            },
            "msg_undo_no_operation": {
                "zh_CN": "没有可撤销的操作。",
                "en_US": "No operation to undo."
            },
            "msg_undo_no_files": {
                "zh_CN": "没有找到需要恢复的文件。",
                "en_US": "No files to restore found."
            },
            "msg_undo_failed": {
                "zh_CN": "撤销失败，部分文件无法恢复：\n{}",
                "en_US": "Undo failed, some files could not be restored:\n{}"
            },
            "msg_undo_error": {
                "zh_CN": "撤销失败",
                "en_US": "Undo Failed"
            },
            "msg_rename_invalid": {
                "zh_CN": "非法字符",
                "en_US": "Invalid Characters"
            },

            # 关于窗口
            "about_title": {
                "zh_CN": "关于",
                "en_US": "About"
            },
            "about_text": {
                "zh_CN": "软件版本：V 1.0.0\n开发者：ZBT Studio\n代码完全由DeepSeek编写，程序图标使用豆包AI生成。",
                "en_US": "Version: V 1.0.0\nDeveloper: ZBT Studio\nCode entirely written by DeepSeek, program icon generated by Doubao AI."
            },

            # 非法字符提示
            "tooltip_invalid_chars": {
                "zh_CN": "非法字符: {}\n请勿输入 {} 等非法字符",
                "en_US": "Invalid characters: {}\nDo not input {} etc."
            },

            # 底部统计标签
            "total_files_label": {
                "zh_CN": "文件总数:",
                "en_US": "Total:"
            },
            "selected_label": {
                "zh_CN": "已勾选:",
                "en_US": "Selected:"
            },

            # 新增：文件被占用提示
            "msg_file_busy_title": {
                "zh_CN": "文件被占用",
                "en_US": "File Busy"
            },
            "msg_file_busy_body": {
                "zh_CN": "文件“{filename}”被占用，请关闭相关程序后重试。",
                "en_US": "The file '{filename}' is busy. Please close the program that is using it and try again."
            }
        }

    def get(self, key, lang=None):
        """获取当前语言下的字符串"""
        if lang is None:
            lang = self.lang
        if lang == "auto":
            # 获取系统语言
            sys_lang = locale.getdefaultlocale()[0]
            if sys_lang and sys_lang.startswith("zh"):
                lang = "zh_CN"
            else:
                lang = "en_US"
        if key in self.strings:
            trans = self.strings[key]
            if lang in trans:
                return trans[lang]
            elif "zh_CN" in trans:
                return trans["zh_CN"]
            else:
                return key
        return key

    def set_language(self, lang):
        self.lang = lang

# ---------- 主应用类 ----------
class BulkRenameTool:
    def __init__(self):
        # 单实例检测
        self.is_main_instance = self.create_mutex()
        if not self.is_main_instance:
            self.send_files_to_running_instance(sys.argv[1:])
            sys.exit(0)

        # 创建主窗口
        if HAS_DND:
            self.root = TkinterDnD.Tk()
        else:
            self.root = Tk()
        self.root.title("EZ_Rename - Bulk Rename Tool")  # 临时标题，稍后更新
        self.root.geometry("920x650")
        self.root.minsize(920, 650)

        # 加载语言配置
        self.translator = Translator()
        self.load_language()
        self.root.title(self.translator.get("app_title"))

        # ========== 设置窗口图标并保存路径 ==========
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        self.icon_path = os.path.join(base_path, '图标.ico')
        if os.path.exists(self.icon_path):
            try:
                self.root.iconbitmap(self.icon_path)
            except Exception:
                pass
        # ===========================================

        # 变量定义
        self.files = []                     # 列表元素为文件字典或分隔符字典
        self.undo_stack = []                 # 撤销栈，存储操作记录
        # 模式变量
        self.current_mode = StringVar()
        # 设置默认模式（替换文本）
        self.current_mode.set(self.translator.get("mode_replace"))

        self.replace_find = StringVar()
        self.replace_with = StringVar()
        self.add_text = StringVar()
        self.add_position = StringVar()
        self.format_name = StringVar(value="")
        self.format_sep = StringVar(value="_")
        self.format_start = IntVar(value=1)
        self.format_step = IntVar(value=1)
        self.format_pad = IntVar(value=1)
        self.format_date = StringVar()
        self.format_seq_pos = StringVar()
        self.status_text = StringVar(value=self.translator.get("status_ready"))
        self.selected_count = IntVar(value=0)
        self.total_count = IntVar(value=0)

        # 拖拽勾选相关
        self.drag_start = None
        self.drag_anchor_state = None

        # 悬停高亮相关
        self.hovered_item = None

        # 浮动提示窗口字典 {widget: toplevel}
        self.tooltip_windows = {}

        # 加载其他配置
        self.load_config()

        # 创建UI
        self.create_widgets()

        # 应用浅色主题
        self.apply_light_theme()

        # 绑定变量跟踪
        self.trace_variables()

        # 绑定全局快捷键
        self.bind_shortcuts()

        # 启动socket监听
        self.start_socket_listener()

        # 处理命令行参数
        if len(sys.argv) > 1:
            self.add_paths(sys.argv[1:])

        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    # ---------- 语言管理 ----------
    def load_language(self):
        """从配置文件加载语言设置"""
        try:
            with open(CONFIG_FILE, 'r') as f:
                cfg = json.load(f)
                if 'language' in cfg:
                    self.translator.set_language(cfg['language'])
        except:
            pass

    def save_language(self):
        """保存语言设置到配置文件"""
        cfg = {}
        try:
            with open(CONFIG_FILE, 'r') as f:
                cfg = json.load(f)
        except:
            pass
        cfg['language'] = self.translator.lang
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(cfg, f)
        except:
            pass

    def _map_language_values(self, old_lang, new_lang):
        """将语言相关的变量值从旧语言映射到新语言"""
        if old_lang == new_lang:
            return

        # 定义映射表（双向映射）
        position_map = {
            "开头": "Start", "末尾": "End",
            "Start": "开头", "End": "末尾"
        }
        date_map = {
            "无": "None", "yyyy-mm-dd": "yyyy-mm-dd", "yyyymmdd": "yyyymmdd",
            "None": "无", "yyyy-mm-dd": "yyyy-mm-dd", "yyyymmdd": "yyyymmdd"
        }
        seq_pos_map = {
            "前缀": "Prefix", "后缀": "Suffix",
            "Prefix": "前缀", "Suffix": "后缀"
        }

        # 映射 add_position
        old_val = self.add_position.get()
        if old_val in position_map:
            self.add_position.set(position_map[old_val])

        # 映射 format_date
        old_val = self.format_date.get()
        if old_val in date_map:
            self.format_date.set(date_map[old_val])

        # 映射 format_seq_pos
        old_val = self.format_seq_pos.get()
        if old_val in seq_pos_map:
            self.format_seq_pos.set(seq_pos_map[old_val])

    def set_language(self, lang):
        """切换语言并重建界面（保留模式选择）"""
        old_lang = self.translator.lang
        if old_lang == lang:
            return

        # 记录当前模式索引（0:替换文本, 1:添加文本, 2:格式）
        mode_values = [self.translator.get("mode_replace"),
                       self.translator.get("mode_add"),
                       self.translator.get("mode_format")]
        current_mode_text = self.current_mode.get()
        try:
            mode_index = mode_values.index(current_mode_text)
        except ValueError:
            mode_index = 0

        # 映射其他语言相关变量（在重建UI前进行）
        self._map_language_values(old_lang, lang)

        # 设置新语言
        self.translator.set_language(lang)
        self.save_language()

        # 重建界面（不恢复模式，保留文件和撤销栈）
        self.reload_ui()

        # 根据索引设置新语言下的模式文本
        new_mode_values = [self.translator.get("mode_replace"),
                           self.translator.get("mode_add"),
                           self.translator.get("mode_format")]
        self.current_mode.set(new_mode_values[mode_index])

        # 强制更新参数面板和预览
        self.on_mode_change()

        # 刷新预览（确保新文件名显示正确）
        self.refresh_new_names()

    def reload_ui(self):
        """重建整个UI（保留文件列表和撤销栈）"""
        # 保存当前状态
        files_snapshot = copy.deepcopy(self.files)
        undo_snapshot = copy.deepcopy(self.undo_stack)
        # 销毁所有子窗口（除了根窗口）
        for widget in self.root.winfo_children():
            widget.destroy()
        # 重新创建UI
        self.create_widgets()
        # 恢复状态
        self.files = files_snapshot
        self.undo_stack = undo_snapshot
        # 刷新显示
        self.refresh_display()
        self.refresh_new_names()
        # 重新绑定变量跟踪
        self.trace_variables()
        # 重新应用主题
        self.apply_light_theme()

    # ---------- 配置持久化 ----------
    def load_config(self):
        try:
            with open(CONFIG_FILE, 'r') as f:
                cfg = json.load(f)
                if 'geometry' in cfg:
                    self.root.geometry(cfg['geometry'])
                # 加载发送到快捷方式状态
                self.sendto_enabled = cfg.get('sendto_enabled', False)
                # 加载开始菜单快捷方式状态
                self.startmenu_enabled = cfg.get('startmenu_enabled', False)
        except:
            self.sendto_enabled = False
            self.startmenu_enabled = False

    def save_config(self):
        cfg = {}
        try:
            with open(CONFIG_FILE, 'r') as f:
                cfg = json.load(f)
        except:
            pass
        cfg['geometry'] = self.root.geometry()
        cfg['language'] = self.translator.lang
        cfg['sendto_enabled'] = self.sendto_enabled
        cfg['startmenu_enabled'] = self.startmenu_enabled
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(cfg, f)
        except:
            pass

    # ---------- 固定浅色主题 ----------
    def apply_light_theme(self):
        bg = '#f0f0f0'
        fg = '#000000'
        select_bg = '#f0f0f0'          # 选中背景与正常背景一致，避免高亮
        entry_bg = '#ffffff'
        combo_list_bg = '#ffffff'

        style = ttk.Style()
        style.theme_use('vista')
        style.configure('.', background=bg, foreground=fg)
        style.configure('Treeview', background='white', foreground='black', fieldbackground='white')
        style.map('Treeview',
                  background=[('selected', select_bg)],
                  foreground=[('selected', 'black')])
        style.configure('TButton', background=bg)
        style.configure('TLabel', background=bg)
        style.configure('TFrame', background=bg)
        style.configure('TEntry', fieldbackground=entry_bg, foreground='black')
        style.configure('TSpinbox', fieldbackground=entry_bg, foreground='black')
        style.configure('TCombobox', fieldbackground=entry_bg, foreground='black')
        style.map('TCombobox',
                  fieldbackground=[('active', '#ADD8E6'), ('readonly', '#ffffff')],
                  background=[('active', '#ADD8E6')])
        self.root.option_add('*TCombobox*Listbox.background', combo_list_bg)
        self.root.option_add('*TCombobox*Listbox.foreground', fg)

        self.root.configure(bg=bg)
        for child in self.root.winfo_children():
            if not isinstance(child, (ttk.Frame, ttk.LabelFrame)):
                try:
                    child.configure(bg=bg, fg=fg)
                except:
                    pass

        # 配置悬停高亮标签
        self.tree.tag_configure('hover', background='#ADD8E6')
        # 分隔符标签
        self.tree.tag_configure('separator', background='white', foreground='#ADD8E6', font=('TkDefaultFont', 9, 'bold'))

        self.refresh_display()

    # ---------- 创建 UI ----------
    def create_widgets(self):
        # ---------- 菜单栏 ----------
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.translator.get("menu_file"), menu=file_menu)
        file_menu.add_command(label=self.translator.get("menu_import_files"), accelerator="Ctrl+I", command=self.add_files)
        file_menu.add_command(label=self.translator.get("menu_import_folder"), accelerator="Ctrl+Alt+I", command=self.add_folders)
        file_menu.add_separator()
        file_menu.add_command(label=self.translator.get("menu_exit"), accelerator="Ctrl+W / Ctrl+Q", command=self.on_closing)

        # 编辑菜单
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.translator.get("menu_edit"), menu=edit_menu)
        edit_menu.add_command(label=self.translator.get("menu_undo"), accelerator="Ctrl+Z", command=self.undo)
        edit_menu.add_separator()
        edit_menu.add_command(label=self.translator.get("menu_select_all"), accelerator="Ctrl+A", command=self.select_all)
        edit_menu.add_command(label=self.translator.get("menu_invert_selection"), command=self.invert_selection)
        edit_menu.add_command(label=self.translator.get("menu_clear_selection"), accelerator="Ctrl+D", command=self.clear_all_selection)
        edit_menu.add_separator()
        edit_menu.add_command(label=self.translator.get("menu_refresh"), accelerator="F5", command=self.refresh_files)
        edit_menu.add_command(label=self.translator.get("menu_clear_list"), command=self.clear_list)
        edit_menu.add_command(label=self.translator.get("menu_remove_selected"), accelerator="Delete", command=self.remove_selected)

        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=self.translator.get("menu_help"), menu=help_menu)

        # 语言子菜单
        lang_menu = tk.Menu(help_menu, tearoff=0)
        help_menu.add_cascade(label=self.translator.get("menu_language"), menu=lang_menu)
        self.lang_var = tk.StringVar(value=self.translator.lang)
        lang_menu.add_radiobutton(label=self.translator.get("menu_lang_auto"), variable=self.lang_var, value="auto",
                                  command=lambda: self.set_language("auto"))
        lang_menu.add_radiobutton(label=self.translator.get("menu_lang_zh"), variable=self.lang_var, value="zh_CN",
                                  command=lambda: self.set_language("zh_CN"))
        lang_menu.add_radiobutton(label=self.translator.get("menu_lang_en"), variable=self.lang_var, value="en_US",
                                  command=lambda: self.set_language("en_US"))

        # 添加到发送到（勾选项）
        self.sendto_var = tk.BooleanVar(value=self.sendto_enabled)
        help_menu.add_checkbutton(label=self.translator.get("menu_add_to_sendto"),
                                  variable=self.sendto_var,
                                  command=self.toggle_sendto_shortcut)

        # 固定到开始菜单（勾选项）
        self.startmenu_var = tk.BooleanVar(value=self.startmenu_enabled)
        help_menu.add_checkbutton(label=self.translator.get("menu_add_to_startmenu"),
                                  variable=self.startmenu_var,
                                  command=self.toggle_startmenu_shortcut)

        help_menu.add_separator()
        help_menu.add_command(label=self.translator.get("menu_add_to_desktop"), command=self.add_desktop_shortcut)
        help_menu.add_separator()
        help_menu.add_command(label=self.translator.get("menu_about"), command=self.show_about)

        # 顶部按钮栏
        top_frame = ttk.Frame(self.root, padding=5)
        top_frame.pack(fill='x')

        ttk.Button(top_frame, text=self.translator.get("btn_add_files"), command=self.add_files).pack(side='left', padx=2)
        ttk.Button(top_frame, text=self.translator.get("btn_add_folder"), command=self.add_folders).pack(side='left', padx=2)
        ttk.Button(top_frame, text=self.translator.get("btn_clear_list"), command=self.clear_list).pack(side='left', padx=2)
        ttk.Button(top_frame, text=self.translator.get("btn_remove_selected"), command=self.remove_selected).pack(side='left', padx=2)

        # 重命名模式区域
        mode_frame = ttk.LabelFrame(self.root, text="", padding=5)
        mode_frame.pack(fill='x', padx=5, pady=5)

        row1 = ttk.Frame(mode_frame)
        row1.pack(fill='x', pady=2)
        ttk.Label(row1, text=self.translator.get("label_mode")).pack(side='left')
        self.mode_combo = ttk.Combobox(row1, textvariable=self.current_mode,
                                       values=(self.translator.get("mode_replace"),
                                               self.translator.get("mode_add"),
                                               self.translator.get("mode_format")),
                                       state='readonly', width=15)
        self.mode_combo.pack(side='left', padx=5)
        self.mode_combo.bind('<<ComboboxSelected>>', self.on_mode_change)

        # 注意：不在此处设置 current_mode 的初始值，因为已在 __init__ 中设置
        self.param_frame = ttk.Frame(mode_frame)
        self.param_frame.pack(fill='x', pady=2)
        self.update_param_panel()

        # 文件列表区域
        list_frame = ttk.LabelFrame(self.root, text="", padding=5)
        list_frame.pack(fill='both', expand=True, padx=5, pady=5)

        columns = ('checked', 'original', 'new', 'status', 'location')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=12)

        # 设置列标题（使用翻译）
        self.tree.heading('checked', text=self.translator.get("col_check"))
        self.tree.heading('original', text=self.translator.get("col_original"))
        self.tree.heading('new', text=self.translator.get("col_new"))
        self.tree.heading('status', text=self.translator.get("col_status"))
        self.tree.heading('location', text=self.translator.get("col_location"))

        self.tree.column('checked', width=40, anchor='center', stretch=False)
        self.tree.column('original', width=200, stretch=False)
        self.tree.column('new', width=200, stretch=False)
        self.tree.column('status', width=80, anchor='center', stretch=False)
        self.tree.column('location', width=250, stretch=False)

        vsb = ttk.Scrollbar(list_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        hsb = ttk.Scrollbar(list_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        self.tree.bind('<ButtonRelease-1>', self.on_tree_click)
        self.tree.bind('<Motion>', self.on_tree_motion)
        self.tree.bind('<ButtonPress-1>', self.on_drag_start)
        self.tree.bind('<B1-Motion>', self.on_drag_motion)
        self.tree.bind('<ButtonRelease-1>', self.on_drag_end, add='+')

        select_frame = ttk.Frame(list_frame)
        select_frame.grid(row=2, column=0, columnspan=2, pady=5, sticky='w')
        ttk.Button(select_frame, text=self.translator.get("btn_select_all"), command=self.select_all).pack(side='left', padx=5)
        ttk.Button(select_frame, text=self.translator.get("btn_invert"), command=self.invert_selection).pack(side='left', padx=5)
        ttk.Button(select_frame, text=self.translator.get("btn_sort"), command=self.organize_files).pack(side='left', padx=5)
        ttk.Button(select_frame, text=self.translator.get("btn_refresh"), command=self.refresh_files).pack(side='left', padx=5)

        bottom_frame = ttk.Frame(self.root, padding=5)
        bottom_frame.pack(fill='x')

        self.status_label = ttk.Label(bottom_frame, textvariable=self.status_text)
        self.status_label.pack(side='left')

        ttk.Label(bottom_frame, text=self.translator.get("total_files_label")).pack(side='left', padx=(20,0))
        ttk.Label(bottom_frame, textvariable=self.total_count).pack(side='left')
        ttk.Label(bottom_frame, text=self.translator.get("selected_label")).pack(side='left', padx=(10,0))
        ttk.Label(bottom_frame, textvariable=self.selected_count).pack(side='left')

        self.rename_btn = ttk.Button(bottom_frame, text=self.translator.get("btn_rename"), command=self.rename_files)
        self.rename_btn.pack(side='right', padx=5)

        self.undo_btn = ttk.Button(bottom_frame, text=self.translator.get("btn_undo"), command=self.undo)
        self.undo_btn.pack(side='right', padx=5)

        self.refresh_display()

        if HAS_DND:
            self.root.drop_target_register(DND_FILES)
            self.root.dnd_bind('<<Drop>>', self.on_drop)

    # ---------- 快捷方式管理 ----------
    def toggle_sendto_shortcut(self):
        """根据勾选状态添加或删除发送到快捷方式"""
        if self.sendto_var.get():
            # 添加
            if ensure_sendto_shortcut(True):
                self.sendto_enabled = True
                self.status_text.set(self.translator.get("status_added_to_sendto"))
            else:
                # 如果添加失败，恢复勾选状态为 False
                self.sendto_var.set(False)
                self.status_text.set(self.translator.get("error_add_to_sendto"))
        else:
            # 删除
            if ensure_sendto_shortcut(False):
                self.sendto_enabled = False
                self.status_text.set(self.translator.get("status_removed_from_sendto"))
        self.save_config()

    def toggle_startmenu_shortcut(self):
        """根据勾选状态添加或删除开始菜单快捷方式，并在Win11上固定/取消固定"""
        startmenu_folder = get_startmenu_folder()
        shortcut_path = os.path.join(startmenu_folder, 'EZ_Rename.lnk')
        exe_path = get_exe_path()
        win_version = platform.release()

        if self.startmenu_var.get():
            # 添加
            if os.path.exists(shortcut_path):
                self.status_text.set(self.translator.get("status_already_exists"))
                return
            if create_shortcut(exe_path, shortcut_path, "EZ_Rename"):
                self.startmenu_enabled = True
                self.status_text.set(self.translator.get("status_added_to_startmenu"))
                # Win11 额外固定
                if win_version == '11':
                    if pin_to_startmenu(shortcut_path):
                        self.status_text.set(self.translator.get("status_pinned_to_startmenu"))
                    else:
                        self.status_text.set("固定失败，请手动固定")
            else:
                self.startmenu_var.set(False)
                self.status_text.set("无法创建快捷方式")
        else:
            # 删除
            if os.path.exists(shortcut_path):
                # Win11 尝试取消固定
                if win_version == '11':
                    if unpin_from_startmenu(shortcut_path):
                        self.status_text.set(self.translator.get("status_unpinned_from_startmenu"))
                    else:
                        self.status_text.set("取消固定失败，请手动取消")
                if delete_shortcut(shortcut_path):
                    self.startmenu_enabled = False
                    if self.status_text.get() == "":
                        self.status_text.set(self.translator.get("status_removed_from_startmenu"))
                else:
                    self.startmenu_var.set(True)
                    self.status_text.set("删除快捷方式失败")
            else:
                self.startmenu_enabled = False
                self.status_text.set(self.translator.get("status_removed_from_startmenu"))
        self.save_config()

    def add_desktop_shortcut(self):
        success, msg = add_desktop_shortcut()
        self.status_text.set(msg)

    # ---------- 单实例处理 ----------
    def create_mutex(self):
        try:
            kernel32 = ctypes.windll.kernel32
            mutex = kernel32.CreateMutexW(None, False, MUTEX_NAME)
            last_error = kernel32.GetLastError()
            return last_error != 183
        except:
            return True

    def start_socket_listener(self):
        self.socket_queue = queue.Queue()
        self.listener_running = True
        port = self.find_free_port()
        self.listener_port = port
        with open(SOCKET_PORT_FILE, 'w') as f:
            f.write(str(port))

        def server_thread():
            server = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            server.bind(('127.0.0.1', port))
            server.listen(1)
            server.settimeout(1.0)
            while self.listener_running:
                try:
                    conn, addr = server.accept()
                    data = b''
                    while True:
                        chunk = conn.recv(4096)
                        if not chunk:
                            break
                        data += chunk
                    if data:
                        paths = pickle.loads(data)
                        self.socket_queue.put(paths)
                    conn.close()
                except socket.timeout:
                    continue
                except:
                    break
            server.close()

        threading.Thread(target=server_thread, daemon=True).start()
        self.process_socket_queue()

    def process_socket_queue(self):
        try:
            while True:
                paths = self.socket_queue.get_nowait()
                self.add_paths(paths)
        except queue.Empty:
            pass
        self.root.after(500, self.process_socket_queue)

    def send_files_to_running_instance(self, paths):
        if not paths:
            return
        try:
            with open(SOCKET_PORT_FILE, 'r') as f:
                port = int(f.read().strip())
            client = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            client.connect(('127.0.0.1', port))
            data = pickle.dumps(paths)
            client.sendall(data)
            client.close()
        except:
            pass

    # ---------- 参数面板 ----------
    def update_param_panel(self):
        self._clear_all_tooltips()

        for widget in self.param_frame.winfo_children():
            widget.destroy()

        mode = self.current_mode.get()
        if mode == self.translator.get("mode_replace"):
            ttk.Label(self.param_frame, text=self.translator.get("label_find")).pack(side='left')
            entry_find = tk.Entry(self.param_frame, textvariable=self.replace_find, width=15,
                                  bg='#ffffff', fg='black', relief='solid', bd=1)
            entry_find.pack(side='left', padx=5)
            self._bind_entry_validation(entry_find)

            ttk.Label(self.param_frame, text=self.translator.get("label_replace_with")).pack(side='left')
            entry_replace = tk.Entry(self.param_frame, textvariable=self.replace_with, width=15,
                                     bg='#ffffff', fg='black', relief='solid', bd=1)
            entry_replace.pack(side='left', padx=5)
            self._bind_entry_validation(entry_replace)

        elif mode == self.translator.get("mode_add"):
            ttk.Label(self.param_frame, text=self.translator.get("label_add_text")).pack(side='left')
            entry_add = tk.Entry(self.param_frame, textvariable=self.add_text, width=20,
                                 bg='#ffffff', fg='black', relief='solid', bd=1)
            entry_add.pack(side='left', padx=5)
            self._bind_entry_validation(entry_add)

            ttk.Label(self.param_frame, text=self.translator.get("label_position")).pack(side='left')
            pos_combo = ttk.Combobox(self.param_frame, textvariable=self.add_position,
                                      values=(self.translator.get("position_start"),
                                              self.translator.get("position_end")),
                                      state='readonly', width=6)
            pos_combo.pack(side='left', padx=5)
            pos_combo.bind('<<ComboboxSelected>>', lambda e: self.refresh_new_names())
            # 设置初始值（如果没有值，则默认末尾）
            if not self.add_position.get():
                self.add_position.set(self.translator.get("position_end"))

        else:  # 格式模式
            ttk.Label(self.param_frame, text=self.translator.get("label_base_name")).pack(side='left')
            entry_name = tk.Entry(self.param_frame, textvariable=self.format_name, width=10,
                                  bg='#ffffff', fg='black', relief='solid', bd=1)
            entry_name.pack(side='left', padx=2)
            self._bind_entry_validation(entry_name)

            ttk.Label(self.param_frame, text=self.translator.get("label_separator")).pack(side='left')
            entry_sep = tk.Entry(self.param_frame, textvariable=self.format_sep, width=2,
                                 bg='#ffffff', fg='black', relief='solid', bd=1)
            entry_sep.pack(side='left', padx=2)
            self._bind_entry_validation(entry_sep)

            ttk.Label(self.param_frame, text=self.translator.get("label_start")).pack(side='left')
            self.spin_start = ttk.Spinbox(self.param_frame, from_=0, to=9999, textvariable=self.format_start, width=5)
            self.spin_start.pack(side='left', padx=2)
            self.spin_start.bind('<MouseWheel>', self.on_spinwheel)
            ttk.Label(self.param_frame, text=self.translator.get("label_step")).pack(side='left')
            self.spin_step = ttk.Spinbox(self.param_frame, from_=1, to=999, textvariable=self.format_step, width=5)
            self.spin_step.pack(side='left', padx=2)
            self.spin_step.bind('<MouseWheel>', self.on_spinwheel)
            ttk.Label(self.param_frame, text=self.translator.get("label_pad")).pack(side='left')
            self.spin_pad = ttk.Spinbox(self.param_frame, from_=1, to=10, textvariable=self.format_pad, width=5)
            self.spin_pad.pack(side='left', padx=2)
            self.spin_pad.bind('<MouseWheel>', self.on_spinwheel)

            def validate_pad():
                if self.format_pad.get() < 1:
                    self.format_pad.set(1)
            self.format_pad.trace_add('write', lambda *args: validate_pad())

            sub_frame = ttk.Frame(self.param_frame)
            sub_frame.pack(side='left', padx=10)
            ttk.Label(sub_frame, text=self.translator.get("label_date_format")).pack(side='left')
            date_combo = ttk.Combobox(sub_frame, textvariable=self.format_date,
                                       values=(self.translator.get("date_none"),
                                               self.translator.get("date_ymd"),
                                               self.translator.get("date_yyyymmdd")),
                                       state='readonly', width=12)
            date_combo.pack(side='left', padx=5)
            # 设置初始值
            if not self.format_date.get():
                self.format_date.set(self.translator.get("date_none"))

            ttk.Label(sub_frame, text=self.translator.get("label_seq_position")).pack(side='left')
            seq_combo = ttk.Combobox(sub_frame, textvariable=self.format_seq_pos,
                                      values=(self.translator.get("seq_prefix"),
                                              self.translator.get("seq_suffix")),
                                      state='readonly', width=6)
            seq_combo.pack(side='left', padx=5)
            if not self.format_seq_pos.get():
                self.format_seq_pos.set(self.translator.get("seq_suffix"))

            # 绑定事件，确保组合框内容变化时刷新预览
            date_combo.bind('<<ComboboxSelected>>', lambda e: self.refresh_new_names())
            seq_combo.bind('<<ComboboxSelected>>', lambda e: self.refresh_new_names())

        self.bind_param_changes()

    def _bind_entry_validation(self, entry):
        """为输入框绑定实时验证事件，并执行初始验证"""
        entry.bind('<KeyRelease>', self._validate_entry_input)
        self._check_entry_invalid_chars(entry)

    def _validate_entry_input(self, event):
        """事件回调：检查输入框内容中的非法字符，并更新颜色和提示"""
        widget = event.widget
        self._check_entry_invalid_chars(widget)

    def _check_entry_invalid_chars(self, widget):
        """检查指定输入框中的非法字符，更新文字颜色和浮动提示"""
        text = widget.get()
        invalid_chars = get_invalid_chars_in_name(text)
        if invalid_chars:
            widget.config(fg='red')
            chars_str = ', '.join(sorted(invalid_chars))
            tip_text = self.translator.get("tooltip_invalid_chars").format(chars_str, INVALID_CHARS)
            self._show_tooltip(widget, tip_text)
        else:
            widget.config(fg='black')
            self._hide_tooltip(widget)

    def _show_tooltip(self, widget, text):
        """在指定 widget 上方显示浮动提示窗口（无边框，两行文本）"""
        if widget in self.tooltip_windows:
            tip = self.tooltip_windows[widget]
            if tip.winfo_exists():
                for child in tip.winfo_children():
                    if isinstance(child, tk.Label):
                        child.config(text=text)
                        tip.update_idletasks()
                        tip.geometry(f"{child.winfo_reqwidth()+10}x{child.winfo_reqheight()+10}")
                        break
                self._position_tooltip(widget, tip)
                return
            else:
                del self.tooltip_windows[widget]

        tip = tk.Toplevel(self.root)
        tip.wm_overrideredirect(True)
        tip.wm_attributes('-topmost', True)
        tip.configure(bg='#FFFF99')

        label = tk.Label(tip, text=text, bg='#FFFF99', fg='#000000',
                         font=('TkDefaultFont', 9), padx=5, pady=2,
                         justify='left')
        label.pack()

        tip.update_idletasks()
        tip.geometry(f"{label.winfo_reqwidth()+10}x{label.winfo_reqheight()+10}")

        self._position_tooltip(widget, tip)

        self.tooltip_windows[widget] = tip

        def on_destroy(event):
            self._hide_tooltip(widget)
        widget.bind('<Destroy>', on_destroy, add=True)

    def _position_tooltip(self, widget, tip):
        """将提示窗口放置在 widget 的正上方（不遮挡文本框）"""
        widget.update_idletasks()
        x = widget.winfo_rootx()
        y = widget.winfo_rooty()
        tip_y = y - tip.winfo_height() - 5
        if tip_y < 0:
            tip_y = y + widget.winfo_height() + 5
        tip.geometry(f"+{x}+{tip_y}")

    def _hide_tooltip(self, widget):
        """隐藏并销毁指定 widget 的浮动提示"""
        if widget in self.tooltip_windows:
            tip = self.tooltip_windows[widget]
            if tip.winfo_exists():
                tip.destroy()
            del self.tooltip_windows[widget]

    def _clear_all_tooltips(self):
        """清除所有浮动提示窗口"""
        for widget, tip in list(self.tooltip_windows.items()):
            if tip.winfo_exists():
                tip.destroy()
        self.tooltip_windows.clear()

    # ---------- 菜单命令 ----------
    def show_about(self):
        """显示关于窗口（居中对齐主窗口，文字左对齐留边距）"""
        about = tk.Toplevel(self.root)
        about.title(self.translator.get("about_title"))
        # ========== 设置关于窗口图标 ==========
        if hasattr(self, 'icon_path') and self.icon_path and os.path.exists(self.icon_path):
            try:
                about.iconbitmap(self.icon_path)
            except Exception:
                pass
        # =====================================
        about.resizable(False, False)
        about.transient(self.root)
        about.grab_set()

        # 设置窗口大小（根据内容调整）
        text = self.translator.get("about_text")
        frame = ttk.Frame(about, padding=15)
        frame.pack(fill='both', expand=True)
        label = tk.Label(frame, text=text, justify='left', anchor='nw', font=('TkDefaultFont', 10))
        label.pack(fill='both', expand=True)

        about.update_idletasks()
        width = label.winfo_reqwidth() + 30
        height = label.winfo_reqheight() + 30
        about.geometry(f"{width}x{height}")

        # 居中于主窗口
        self.center_window(about)

        def close_about(event=None):
            about.destroy()
        label.bind('<Button-1>', close_about)
        about.bind('<Button-1>', close_about)
        about.bind('<FocusOut>', close_about)

        about.focus_force()

    def center_window(self, win):
        """将窗口 win 居中于主窗口"""
        win.update_idletasks()
        w = win.winfo_width()
        h = win.winfo_height()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (w // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (h // 2)
        win.geometry(f"+{x}+{y}")

    def clear_all_selection(self):
        """全部取消勾选"""
        for f in self.files:
            if f.get('type') == 'file':
                f['checked'] = False
        self.refresh_new_names()

    # ---------- 快捷键绑定 ----------
    def bind_shortcuts(self):
        self.root.bind_all('<Control-i>', lambda e: self.add_files())
        self.root.bind_all('<Control-Alt-i>', lambda e: self.add_folders())
        self.root.bind_all('<Control-w>', lambda e: self.on_closing())
        self.root.bind_all('<Control-q>', lambda e: self.on_closing())
        self.root.bind_all('<Control-z>', lambda e: self.undo())
        self.root.bind_all('<Control-a>', lambda e: self.select_all())
        self.root.bind_all('<Control-d>', lambda e: self.clear_all_selection())
        self.root.bind_all('<F5>', lambda e: self.refresh_files())
        self.root.bind_all('<Delete>', lambda e: self.remove_selected())

    # ---------- 其他原有方法 ----------
    def on_mode_change(self, event=None):
        self.update_param_panel()
        self.refresh_new_names()

    def bind_param_changes(self):
        vars_to_trace = [self.replace_find, self.replace_with, self.add_text,
                         self.add_position, self.format_name, self.format_sep,
                         self.format_start, self.format_step, self.format_pad,
                         self.format_date, self.format_seq_pos]
        for var in vars_to_trace:
            var.trace_add('write', lambda *args: self.refresh_new_names())

    def on_spinwheel(self, event):
        widget = event.widget
        if event.delta > 0:
            widget.invoke('buttonup')
        else:
            widget.invoke('buttondown')
        return "break"

    # ---------- 文件列表操作 ----------
    def add_files(self):
        paths = filedialog.askopenfilenames(title=self.translator.get("menu_import_files"))
        if paths:
            self.add_paths(paths)

    def add_folders(self):
        folder = filedialog.askdirectory(title=self.translator.get("menu_import_folder"))
        if folder:
            self.add_paths([folder])

    def add_paths(self, paths):
        new_files = []
        for p in paths:
            p = os.path.abspath(p)
            if os.path.isdir(p):
                try:
                    for root, dirs, files in os.walk(p):
                        for f in files:
                            full_path = os.path.join(root, f)
                            new_files.append(full_path)
                        break
                except:
                    pass
            else:
                new_files.append(p)

        existing_paths = {f['path'] for f in self.get_all_files()}
        new_files = [f for f in new_files if f not in existing_paths]

        if not new_files:
            return

        file_dicts = []
        for p in new_files:
            name = os.path.basename(p)
            file_dicts.append({
                'type': 'file',
                'path': p,
                'name': name,
                'new_name': name,
                'checked': True,
                'success': False
            })

        all_files = self.get_all_files() + file_dicts
        self.reorganize_files(all_files, sort_key=lambda f: f['name'].lower())

    def get_all_files(self):
        return [f for f in self.files if f.get('type') == 'file']

    def reorganize_files(self, files, sort_key=None):
        if sort_key is None:
            sort_key = lambda f: f['name'].lower()
        groups = {}
        for f in files:
            dir_path = os.path.dirname(f['path'])
            groups.setdefault(dir_path, []).append(f)
        for dir_path in groups:
            groups[dir_path].sort(key=sort_key)
        sorted_dirs = sorted(groups.keys())
        new_list = []
        for i, dir_path in enumerate(sorted_dirs):
            new_list.extend(groups[dir_path])
            if i < len(sorted_dirs) - 1:
                new_list.append({'type': 'separator'})
        self.files = new_list
        self.refresh_new_names()

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        self.add_paths(files)

    def clear_list(self):
        self.files.clear()
        self.undo_stack.clear()
        self.refresh_display()

    def remove_selected(self):
        new_files = [f for f in self.files if f.get('type') != 'file' or not f.get('checked', False)]
        files_only = [f for f in new_files if f.get('type') == 'file']
        self.reorganize_files(files_only)

    def select_all(self):
        for f in self.files:
            if f.get('type') == 'file':
                f['checked'] = True
        self.refresh_new_names()

    def invert_selection(self):
        for f in self.files:
            if f.get('type') == 'file':
                f['checked'] = not f.get('checked', False)
        self.refresh_new_names()

    def organize_files(self):
        self.push_undo_state('reorder')
        files = self.get_all_files()
        def key_func(f):
            return f['name'].lower()
        self.reorganize_files(files, sort_key=key_func)

    def refresh_files(self):
        self.push_undo_state('refresh')
        all_files = self.get_all_files()
        new_file_dicts = []
        for f in all_files:
            path = f['path']
            if os.path.exists(path):
                new_name = os.path.basename(path)
                f['name'] = new_name
                f['new_name'] = new_name
                f['success'] = False
                new_file_dicts.append(f)
        self.reorganize_files(new_file_dicts, sort_key=lambda f: f['name'].lower())

    # ---------- 撤销机制 ----------
    def push_undo_state(self, action_type):
        snapshot = copy.deepcopy(self.files)
        self.undo_stack.append({'type': action_type, 'data': snapshot})

    def undo(self):
        if not self.undo_stack:
            messagebox.showinfo(self.translator.get("msg_undo_no_operation"), self.translator.get("msg_undo_no_operation"))
            return
        last_op = self.undo_stack.pop()
        action_type = last_op['type']
        if action_type == 'rename':
            operations = last_op['data']
            to_restore = []
            for src, dst in operations:
                if os.path.exists(dst):
                    to_restore.append((dst, src))
            if not to_restore:
                messagebox.showinfo(self.translator.get("msg_undo_no_files"), self.translator.get("msg_undo_no_files"))
                return

            temp_prefix = f"__tmp_{uuid.uuid4().hex[:8]}_"
            temp_ops = []
            for idx, (current, original) in enumerate(to_restore):
                ext = os.path.splitext(current)[1]
                temp_name = f"{temp_prefix}{idx}{ext}"
                temp_path = os.path.join(os.path.dirname(current), temp_name)
                temp_ops.append((current, temp_path))

            temp_success = []
            failed = []
            for current, temp in temp_ops:
                try:
                    os.rename(current, temp)
                    temp_success.append((current, temp))
                except Exception as e:
                    failed.append((current, original, str(e)))
                    for c, t in temp_success:
                        try:
                            os.rename(t, c)
                        except:
                            pass
                    break

            if failed:
                msg = "\n".join(f"{f[0]} -> {f[1]}: {f[2]}" for f in failed)
                messagebox.showerror(self.translator.get("msg_undo_failed"), msg)
                return

            restore_success = []
            for idx, (current, original) in enumerate(to_restore):
                temp_path = temp_ops[idx][1]
                try:
                    os.rename(temp_path, original)
                    restore_success.append((original, current))
                except Exception as e:
                    for o, c in restore_success:
                        pass
                    try:
                        os.rename(temp_path, current)
                    except:
                        pass
                    messagebox.showerror(self.translator.get("msg_undo_error"), f"无法恢复文件 {current} -> {original}: {e}")
                    return

            for original, current in restore_success:
                for f in self.files:
                    if f.get('type') == 'file' and f['path'] == current:
                        f['path'] = original
                        f['name'] = os.path.basename(original)
                        f['new_name'] = f['name']
                        f['success'] = False
                        break

            self.refresh_new_names()
            self.status_text.set(self.translator.get("status_undo_success"))

        elif action_type in ('reorder', 'refresh'):
            self.files = copy.deepcopy(last_op['data'])
            self.refresh_new_names()
            self.status_text.set(self.translator.get("status_undo_success"))
        else:
            pass

    # ---------- 鼠标事件 ----------
    def on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == '#1':
                item = self.tree.identify_row(event.y)
                if item:
                    idx = int(item[1:]) - 1
                    if 0 <= idx < len(self.files):
                        f = self.files[idx]
                        if f.get('type') == 'file':
                            f['checked'] = not f.get('checked', False)
                            self.refresh_new_names()
            elif column == '#5':
                item = self.tree.identify_row(event.y)
                if item:
                    idx = int(item[1:]) - 1
                    if 0 <= idx < len(self.files):
                        f = self.files[idx]
                        if f.get('type') == 'file':
                            open_file_location(f['path'])

    def on_drag_start(self, event):
        self.drag_start = self.tree.identify_row(event.y)
        if self.drag_start:
            idx = int(self.drag_start[1:]) - 1
            if 0 <= idx < len(self.files):
                f = self.files[idx]
                if f.get('type') == 'file':
                    self.drag_anchor_state = f.get('checked', False)
                else:
                    self.drag_anchor_state = None
            else:
                self.drag_anchor_state = None
        else:
            self.drag_anchor_state = None

    def on_drag_motion(self, event):
        if self.drag_anchor_state is None:
            return
        current_item = self.tree.identify_row(event.y)
        if not current_item:
            return
        all_items = self.tree.get_children()
        if self.drag_start not in all_items or current_item not in all_items:
            return
        start_idx = all_items.index(self.drag_start)
        end_idx = all_items.index(current_item)
        if start_idx <= end_idx:
            items_range = all_items[start_idx:end_idx+1]
        else:
            items_range = all_items[end_idx:start_idx+1]

        for item in items_range:
            idx = int(item[1:]) - 1
            if 0 <= idx < len(self.files):
                f = self.files[idx]
                if f.get('type') == 'file':
                    new_state = not self.drag_anchor_state
                    if f.get('checked', False) != new_state:
                        f['checked'] = new_state
        self.refresh_display()

    def on_drag_end(self, event):
        self.drag_start = None
        self.drag_anchor_state = None

    def on_tree_motion(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == '#5':
                self.tree.config(cursor="hand2")
            else:
                self.tree.config(cursor="")
        else:
            self.tree.config(cursor="")

        item = self.tree.identify_row(event.y)
        if item != self.hovered_item:
            if self.hovered_item:
                try:
                    tags = self.tree.item(self.hovered_item, 'tags')
                    if 'hover' in tags:
                        new_tags = list(tags)
                        new_tags.remove('hover')
                        self.tree.item(self.hovered_item, tags=tuple(new_tags))
                except:
                    pass
            self.hovered_item = item
            if item:
                idx = int(item[1:]) - 1
                if 0 <= idx < len(self.files):
                    f = self.files[idx]
                    if f.get('type') == 'file':
                        try:
                            tags = self.tree.item(item, 'tags')
                            if 'hover' not in tags:
                                new_tags = list(tags) + ['hover']
                                self.tree.item(item, tags=tuple(new_tags))
                        except:
                            pass

    # ---------- 实时预览 ----------
    def refresh_new_names(self):
        for f in self.files:
            if f.get('type') == 'file':
                f['success'] = False

        mode = self.current_mode.get()
        file_list = [f for f in self.files if f.get('type') == 'file']

        if mode == self.translator.get("mode_format"):
            groups = {}
            for f in file_list:
                dir_path = os.path.dirname(f['path'])
                groups.setdefault(dir_path, []).append(f)
            for dir_path, group in groups.items():
                start = self.format_start.get()
                step = self.format_step.get()
                pad = self.format_pad.get()
                for idx, f in enumerate(group):
                    if f.get('checked', False):
                        num = start + step * idx
                        num_str = str(num).zfill(pad) if pad > 0 else str(num)
                        base_name = self.format_name.get()
                        sep = self.format_sep.get()
                        date_fmt = self.format_date.get()
                        seq_pos = self.format_seq_pos.get()
                        date_part = ""
                        if date_fmt != self.translator.get("date_none"):
                            today = datetime.now()
                            if date_fmt == self.translator.get("date_ymd"):
                                date_part = today.strftime("%Y-%m-%d")
                            else:
                                date_part = today.strftime("%Y%m%d")
                        parts = []
                        if seq_pos == self.translator.get("seq_prefix"):
                            # 顺序：序号 - 基础名称 - 日期（可选）
                            parts.append(num_str)
                            parts.append(base_name)
                            if date_part:
                                parts.append(date_part)
                        else:  # 后缀
                            # 顺序：基础名称 - 序号 - 日期（可选）
                            parts.append(base_name)
                            parts.append(num_str)
                            if date_part:
                                parts.append(date_part)
                        new_base = sep.join(parts)
                        name = f['name']
                        base, ext = os.path.splitext(name)
                        f['new_name'] = new_base + ext if ext else new_base
                    else:
                        f['new_name'] = f['name']
        else:
            for f in file_list:
                if f.get('checked', False):
                    name = f['name']
                    base, ext = os.path.splitext(name)
                    if mode == self.translator.get("mode_replace"):
                        new_base = base.replace(self.replace_find.get(), self.replace_with.get())
                        f['new_name'] = new_base + ext if ext else new_base
                    elif mode == self.translator.get("mode_add"):
                        text = self.add_text.get()
                        pos = self.add_position.get()
                        if pos == self.translator.get("position_start"):
                            f['new_name'] = text + name
                        else:
                            f['new_name'] = base + text + ext if ext else base + text
                else:
                    f['new_name'] = f['name']

        self.refresh_display()

    def refresh_display(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        separator_line = "─" * 40

        for idx, entry in enumerate(self.files):
            if entry.get('type') == 'separator':
                values = (separator_line, separator_line, separator_line, separator_line, separator_line)
                item_id = f'S{idx+1:03d}'
                self.tree.insert('', 'end', iid=item_id, values=values, tags=('separator',))
                continue

            f = entry
            if f['success']:
                status_text = self.translator.get("status_modified")
                tag = 'success'
            elif f['new_name'] != f['name']:
                status_text = self.translator.get("status_ready")
                tag = 'changed'
            else:
                status_text = ""
                tag = 'normal'

            checked = '✔' if f['checked'] else '☐'
            values = (checked, f['name'], f['new_name'], status_text, f['path'])
            item_id = f'I{idx+1:03d}'
            self.tree.insert('', 'end', iid=item_id, values=values, tags=(tag,))

        style = ttk.Style()
        style.configure('success.Treeview', foreground='#008000')
        style.configure('changed.Treeview', foreground='#ff0000')
        style.configure('normal.Treeview', foreground='black')
        self.tree.tag_configure('success', foreground=style.lookup('success.Treeview', 'foreground'))
        self.tree.tag_configure('changed', foreground=style.lookup('changed.Treeview', 'foreground'))
        self.tree.tag_configure('normal', foreground=style.lookup('normal.Treeview', 'foreground'))

        file_count = sum(1 for f in self.files if f.get('type') == 'file')
        selected_count = sum(1 for f in self.files if f.get('type') == 'file' and f.get('checked', False))
        self.total_count.set(file_count)
        self.selected_count.set(selected_count)

    # ---------- 重命名 ----------
    def rename_files(self):
        self.refresh_new_names()

        selected_files = [f for f in self.files if f.get('type') == 'file' and f['checked'] and f['new_name'] != f['name']]
        if not selected_files:
            messagebox.showinfo(self.translator.get("msg_no_rename"), self.translator.get("msg_no_rename"))
            return

        invalid_files = []
        for f in selected_files:
            new_name = f['new_name']
            if not is_valid_filename(new_name):
                invalid_chars = get_invalid_chars_in_name(new_name)
                invalid_files.append((new_name, ', '.join(invalid_chars)))
        if invalid_files:
            msg_body = "\n".join(f"  {name} (非法字符: {chars})" for name, chars in invalid_files)
            msg = self.translator.get("msg_invalid_chars_body").format(msg_body)
            messagebox.showerror(self.translator.get("msg_invalid_chars_title"), msg)
            return

        self.rename_btn.config(state=DISABLED)
        self.undo_btn.config(state=DISABLED)
        self.status_text.set(self.translator.get("status_renaming"))

        final_ops = []
        for f in selected_files:
            src = f['path']
            dst_dir = os.path.dirname(src)
            target_name = f['new_name']
            dst = os.path.join(dst_dir, target_name)
            final_ops.append((src, dst))

        def rename_task():
            temp_prefix = f"__tmp_{uuid.uuid4().hex[:8]}_"
            temp_ops = []
            temp_success = []  # (src, temp_path)
            for idx, (src, dst) in enumerate(final_ops):
                ext = os.path.splitext(dst)[1]
                temp_name = f"{temp_prefix}{idx}{ext}"
                temp_path = os.path.join(os.path.dirname(src), temp_name)
                temp_ops.append((src, temp_path))
                try:
                    os.rename(src, temp_path)
                    temp_success.append((src, temp_path))
                except OSError as e:
                    # 判断是否为文件被占用（共享冲突）: winerror 32
                    if hasattr(e, 'winerror') and e.winerror == 32:
                        error_type = 'busy'
                        error_msg = self.translator.get("msg_file_busy_body").format(filename=os.path.basename(src))
                    else:
                        error_type = 'other'
                        error_msg = str(e)
                    # 回滚所有已成功的临时重命名
                    for s, t in temp_success:
                        try:
                            os.rename(t, s)
                        except:
                            pass
                    self.root.after(0, self.rename_done, [(src, dst, False, error_type, error_msg)], final_ops)
                    return

            final_success = []
            for idx, (src, dst) in enumerate(final_ops):
                temp_path = temp_ops[idx][1]
                try:
                    os.rename(temp_path, dst)
                    final_success.append((src, dst))
                except OSError as e:
                    # 判断是否为文件被占用（共享冲突）
                    if hasattr(e, 'winerror') and e.winerror == 32:
                        error_type = 'busy'
                        error_msg = self.translator.get("msg_file_busy_body").format(filename=os.path.basename(dst))
                    else:
                        error_type = 'other'
                        error_msg = str(e)
                    # 回滚所有已成功的最终重命名（即恢复临时文件）
                    for s, d in final_success:
                        temp_path_undo = temp_ops[final_ops.index((s, d))][1]
                        try:
                            os.rename(d, temp_path_undo)
                        except:
                            pass
                    # 再回滚所有已成功的临时重命名（恢复原始文件）
                    for s, t in temp_success:
                        try:
                            os.rename(t, s)
                        except:
                            pass
                    # 当前失败的临时文件也需要恢复（它还存在）
                    try:
                        os.rename(temp_path, src)
                    except:
                        pass
                    self.root.after(0, self.rename_done, [(src, dst, False, error_type, error_msg)], final_ops)
                    return

            success_results = [(src, dst, True, None, None) for src, dst in final_success]
            self.root.after(0, self.rename_done, success_results, final_ops)

        threading.Thread(target=rename_task, daemon=True).start()

    def rename_done(self, results, final_ops):
        # results: list of (src, dst, ok, error_type, error_msg)
        success_ops = [(src, dst) for src, dst, ok, _, _ in results if ok]
        if success_ops:
            self.undo_stack.append({'type': 'rename', 'data': success_ops})

        success_map = dict(success_ops)
        for f in self.files:
            if f.get('type') == 'file' and f['path'] in success_map:
                f['path'] = success_map[f['path']]
                f['name'] = os.path.basename(f['path'])
                f['new_name'] = f['name']
                f['success'] = True

        success_count = len(success_ops)
        total_selected = len(final_ops)
        fail_count = total_selected - success_count

        self.refresh_display()
        self.rename_btn.config(state=NORMAL)
        self.undo_btn.config(state=NORMAL)
        self.status_text.set(self.translator.get("status_rename_done").format(success_count, fail_count))

        # 检查是否有因文件被占用而失败的文件，如果有，弹出提示框
        busy_files = []
        for src, dst, ok, error_type, error_msg in results:
            if not ok and error_type == 'busy':
                busy_files.append((src, error_msg))
        if busy_files:
            # 仅显示第一个被占用的文件，避免多个弹窗
            src, msg = busy_files[0]
            messagebox.showerror(self.translator.get("msg_file_busy_title"), msg)

    # ---------- 辅助 ----------
    def find_free_port(self):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('127.0.0.1', 0))
            return s.getsockname()[1]

    def trace_variables(self):
        self.bind_param_changes()

    def on_closing(self):
        self.listener_running = False
        self._clear_all_tooltips()
        self.save_config()
        try:
            os.remove(SOCKET_PORT_FILE)
        except:
            pass
        self.root.destroy()

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = BulkRenameTool()
    app.run()