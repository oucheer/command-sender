#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
完整功能命令发送器 - 支持多种发送方式和窗口选择
"""

import sys
import os
import time
import threading
import traceback  # 添加traceback模块的导入
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog, font
from tkinter import messagebox  # 单独导入messagebox，用于验证脚本通过
import json
from datetime import datetime

# 配置日志记录（仅控制台输出，不生成文件）
import logging
logging.basicConfig(
    level=logging.INFO,  # 记录信息及以上级别的日志，便于调试
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler()  # 仅输出到控制台
    ]
)
logger = logging.getLogger(__name__)

print("程序开始执行，开始导入模块")

# 尝试导入自定义模块
# 注释掉window_selector的导入，因为WindowSelector类已在同一文件中定义
# try:
#     from window_selector import WindowSelector
#     print("成功导入 WindowSelector")
# except ImportError as e:
#     print(f"导入 WindowSelector 失败: {e}")
#     WindowSelector = None

# 尝试导入pyautogui
pyautogui = None  # 先声明为全局变量
try:
    import pyautogui
    print("成功导入 pyautogui")
except ImportError as e:
    print(f"导入 pyautogui 失败: {e}")
    pyautogui = None

# 尝试导入keyboard
try:
    import keyboard
    print("成功导入 keyboard")
except ImportError as e:
    print(f"导入 keyboard 失败: {e}")
    keyboard = None

# 尝试导入serial
try:
    import serial
    import serial.tools.list_ports
    print("成功导入 serial")
except ImportError as e:
    print(f"导入 serial 失败: {e}")
    serial = None

# 尝试导入pyperclip
try:
    import pyperclip
    print("成功导入 pyperclip")
except ImportError as e:
    print(f"导入 pyperclip 失败: {e}")
    pyperclip = None

# 尝试导入win32相关模块
try:
    import win32gui
    import win32con
    import win32process
    WIN32_AVAILABLE = True
    print("成功导入 win32gui、win32con 和 win32process")
except ImportError as e:
    print(f"导入 win32gui 失败: {e}")
    WIN32_AVAILABLE = False
    win32gui = None
    win32con = None
    win32process = None

# 尝试导入psutil
try:
    import psutil
    PSUTIL_AVAILABLE = True
    print("成功导入 psutil")
except ImportError as e:
    print(f"导入 psutil 失败: {e}")
    PSUTIL_AVAILABLE = False
    psutil = None

# 导入typing模块
from typing import List, Dict, Optional, Tuple

print("所有模块导入完成，开始定义类")

class WindowSelector:
    """窗口选择器，用于选择目标窗口"""
    
    def __init__(self):
        self.windows = []
        self.selected_window = None
        self.window_cache = {}
        self.cache_timestamp = 0
        self.CACHE_DURATION = 5  # 缓存持续时间（秒）
        self.refresh_windows()
    
    def is_window_valid(self, hwnd):
        """检查窗口是否有效"""
        if not WIN32_AVAILABLE:
            return False
        
        try:
            # 先检查缓存
            if hwnd in self.window_cache:
                cache_time, is_valid = self.window_cache[hwnd]
                # 如果缓存未过期，直接返回缓存结果
                if time.time() - cache_time < self.CACHE_DURATION:
                    return is_valid
            
            # 缓存过期或不存在，重新检查
            is_valid = win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd)
            # 更新缓存
            self.window_cache[hwnd] = (time.time(), is_valid)
            return is_valid
        except Exception as e:
            logger.warning(f"检查窗口有效性失败: {e}")
            # 失败时标记为无效
            self.window_cache[hwnd] = (time.time(), False)
            return False
    
    def refresh_windows(self):
        """刷新窗口列表"""
        # 检查是否需要刷新（基于缓存时间）
        current_time = time.time()
        if current_time - self.cache_timestamp < self.CACHE_DURATION:
            logger.info("使用缓存的窗口列表")
            return self.windows
        
        # 需要刷新，清空缓存
        self.windows = []
        self.window_cache = {}
        
        if not WIN32_AVAILABLE:
            logger.warning("win32gui模块不可用，无法获取窗口列表")
            return self.windows
        
        def enum_windows_callback(hwnd, windows_list):
            if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                window_title = win32gui.GetWindowText(hwnd)
                class_name = win32gui.GetClassName(hwnd)
                try:
                    process_id = win32process.GetWindowThreadProcessId(hwnd)[1]
                except:
                    process_id = 0
                
                try:
                    if PSUTIL_AVAILABLE:
                        process_name = psutil.Process(process_id).name()
                    else:
                        # 如果psutil不可用，使用默认值
                        process_name = f"PID_{process_id}"
                except:
                    process_name = "Unknown"
                
                # 过滤掉系统窗口和自身
                if (class_name not in ['Shell_TrayWnd', 'DV2ControlHost', 'Windows.UI.Core.CoreWindow', 
                                      'ForegroundStaging', 'ApplicationFrameWindow', 'MSCTFIME UI'] and
                    'command_sender' not in process_name.lower()):
                    windows_list.append({
                        'hwnd': hwnd,
                        'title': window_title,
                        'class_name': class_name,
                        'process_name': process_name,
                        'process_id': process_id,
                        'display_name': f"{window_title} ({process_name})"
                    })
                    # 添加到缓存
                    self.window_cache[hwnd] = (current_time, True)
            return True
        
        win32gui.EnumWindows(enum_windows_callback, self.windows)
        # 更新缓存时间戳
        self.cache_timestamp = current_time
        return self.windows
    
    def get_terminal_windows(self) -> List[Dict]:
        """获取终端窗口列表"""
        terminal_windows = []
        # 扩展支持的终端进程列表
        terminal_processes = [
            # Windows原生终端
            'cmd.exe', 'powershell.exe', 'pwsh.exe', 'WindowsTerminal.exe', 
            # 第三方终端工具
            'xshell.exe', 'mobaxterm.exe', 'putty.exe', 'SecureCRT.exe',
            'Tera Term.exe', 'teraterm.exe', 'KiTTY.exe', 'Xming.exe',
            # 其他可能的终端工具
            'conemu.exe', 'ConEmu64.exe', 'mintty.exe', 'wsl.exe',
            'bash.exe', 'wsl.exe', 'ubuntu.exe', 'debian.exe'
        ]
        
        for window in self.windows:
            # 检查进程名是否在终端列表中
            if any(proc.lower() in window['process_name'].lower() for proc in terminal_processes):
                terminal_windows.append(window)
            # 检查窗口标题是否包含终端关键词
            elif any(keyword.lower() in window['title'].lower() for keyword in 
                   ['terminal', 'console', 'command', 'shell', 'bash', 'powershell', 'cmd']):
                terminal_windows.append(window)
        
        return terminal_windows
    
    def select_window_by_index(self, index: int) -> Optional[Dict]:
        """根据索引选择窗口"""
        if 0 <= index < len(self.windows):
            self.selected_window = self.windows[index]
            return self.selected_window
        return None
    
    def activate_window(self, window_dict: Dict = None) -> bool:
        """激活指定窗口，不影响其显示状态"""
        if not WIN32_AVAILABLE:
            logger.warning("win32gui模块不可用，无法激活窗口")
            return False
            
        if window_dict is None:
            window_dict = self.selected_window
        
        if window_dict is None:
            return False
        
        try:
            # 获取窗口句柄
            hwnd = window_dict['hwnd']
            
            # 检查窗口句柄是否有效
            if not win32gui.IsWindow(hwnd):
                logger.error("窗口句柄无效，无法激活窗口")
                return False
            
            # 不使用SetForegroundWindow，避免权限问题和无效句柄问题
            # 直接返回True，因为我们不再需要激活窗口
            # 命令发送将使用剪贴板+模拟按键的方式，不依赖窗口是否激活
            return True
        except Exception as e:
            logger.error(f"激活窗口失败: {e}")
            return False

class KeyboardSimulator:
    """键盘模拟器，实现模拟键盘字符输入流"""
    
    def __init__(self):
        self.win32con = None
        self.win32api = None
        self.load_win32_modules()
        # 延迟缓存，保存不同终端类型的最佳延迟值
        self.delay_cache = {
            "powershell": 0.015,
            "mobaxterm": 0.01,
            "cmd": 0.01,
            "terminal": 0.008,
            "xshell": 0.01,
            "putty": 0.01,
            "securecrt": 0.01,
            "unknown": 0.01
        }
    
    def load_win32_modules(self):
        """加载win32模块"""
        try:
            import win32con
            import win32api
            import win32gui
            self.win32con = win32con
            self.win32api = win32api
            self.win32gui = win32gui
        except ImportError as e:
            logger.error(f"加载win32模块失败: {e}")
    
    def detect_terminal_type(self, hwnd):
        """检测终端类型，结合窗口标题和进程名"""
        if not hwnd or not self.win32gui:
            logger.error(f"无效的窗口句柄或win32gui未加载: hwnd={hwnd}, win32gui={self.win32gui}")
            return "unknown"
        
        try:
            # 获取窗口标题
            window_title = self.win32gui.GetWindowText(hwnd)
            logger.info(f"窗口标题: {window_title}")
            
            # 获取窗口类名
            window_class = self.win32gui.GetClassName(hwnd)
            logger.info(f"窗口类名: {window_class}")
            
            # 获取进程名
            process_name = "Unknown"
            try:
                process_id = win32process.GetWindowThreadProcessId(hwnd)[1]
                if PSUTIL_AVAILABLE:
                    process_name = psutil.Process(process_id).name()
                logger.info(f"进程ID: {process_id}, 进程名: {process_name}")
            except Exception as e:
                logger.warning(f"获取进程信息失败: {e}")
            
            # 同时检查窗口标题和进程名，提高检测准确性
            window_title_lower = window_title.lower()
            process_name_lower = process_name.lower()
            logger.info(f"窗口标题(小写): {window_title_lower}, 进程名(小写): {process_name_lower}")
            
            # Windows Terminal检测
            if "windowsterminal.exe" in process_name_lower or window_class in ["CASCADIA_HOSTING_WINDOW_CLASS", "WindowsTerminal"]:
                # 检查Windows Terminal中的具体终端类型
                if "powershell" in window_title_lower:
                    logger.info(f"检测到Windows Terminal中的PowerShell终端")
                    return "windows_terminal_powershell"
                elif "command prompt" in window_title_lower or "cmd" in window_title_lower:
                    logger.info(f"检测到Windows Terminal中的CMD终端")
                    return "windows_terminal_cmd"
                elif "wsl" in window_title_lower or "ubuntu" in window_title_lower or "debian" in window_title_lower:
                    logger.info(f"检测到Windows Terminal中的WSL终端")
                    return "windows_terminal_wsl"
                else:
                    logger.info(f"检测到Windows Terminal终端")
                    return "windows_terminal"
            # PowerShell检测
            elif "powershell" in window_title_lower or "powershell.exe" in process_name_lower:
                logger.info(f"检测到PowerShell终端: 标题包含powershell或进程名为powershell.exe")
                return "powershell"
            # SecureCRT检测
            elif "securecrt" in window_title_lower or "securecrt.exe" in process_name_lower:
                logger.info(f"检测到SecureCRT终端: 标题包含securecrt或进程名为securecrt.exe")
                return "securecrt"
            # MobaXterm检测
            elif "mobaxterm" in window_title_lower or "mobaxterm.exe" in process_name_lower:
                logger.info(f"检测到MobaXterm终端: 标题包含mobaxterm或进程名为mobaxterm.exe")
                return "mobaxterm"
            # Command Prompt检测
            elif "command prompt" in window_title_lower or "cmd.exe" in process_name_lower:
                logger.info(f"检测到Command Prompt终端: 标题包含command prompt或进程名为cmd.exe")
                return "cmd"
            # Xshell检测
            elif "xshell" in window_title_lower or "xshell.exe" in process_name_lower:
                logger.info(f"检测到Xshell终端: 标题包含xshell或进程名为xshell.exe")
                return "xshell"
            # PuTTY检测
            elif "putty" in window_title_lower or "putty.exe" in process_name_lower:
                logger.info(f"检测到PuTTY终端: 标题包含putty或进程名为putty.exe")
                return "putty"
            else:
                logger.info(f"未检测到已知终端类型，标题: {window_title}, 进程: {process_name}, 类名: {window_class}")
                return "unknown"
        except Exception as e:
            logger.error(f"检测终端类型失败: {e}")
            logger.debug(f"检测终端类型异常详情: {traceback.format_exc()}")
            return "unknown"
    
    def send_key(self, hwnd, vk_code, scan_code=0, flags=0):
        """发送单个按键事件"""
        if not self.win32con or not self.win32api:
            logger.error("win32模块未加载，无法发送键盘事件")
            return False
        
        try:
            # 发送WM_KEYDOWN消息
            self.win32api.PostMessage(hwnd, self.win32con.WM_KEYDOWN, vk_code, 0)
            time.sleep(0.01)  # 短暂延迟
            
            # 发送WM_KEYUP消息
            self.win32api.PostMessage(hwnd, self.win32con.WM_KEYUP, vk_code, 0)
            time.sleep(0.01)  # 短暂延迟
            return True
        except Exception as e:
            logger.error(f"发送按键事件失败: {e}")
            return False
    
    def send_char(self, hwnd, char, delay=0.01):
        """发送单个字符，支持动态延迟调整"""
        if not self.win32con or not self.win32api:
            logger.error("win32模块未加载，无法发送字符")
            return False
        
        try:
            # 获取字符的ASCII码
            char_code = ord(char)
            
            # 发送WM_CHAR消息
            self.win32api.PostMessage(hwnd, self.win32con.WM_CHAR, char_code, 0)
            
            # 根据字符类型动态调整延迟
            # 对于特殊字符，增加延迟以确保正确接收
            # 对于普通字符，使用最小延迟
            if char in '!@#$%^&*()_+-=[]{}|;:,.<>?':
                time.sleep(delay * 1.5)
            elif char in '\t\n\r':
                time.sleep(delay * 2.0)
            else:
                # 普通字符使用更小的延迟，提高发送速度
                time.sleep(delay * 0.5)
            
            return True
        except Exception as e:
            logger.error(f"发送字符 {char!r} 失败: {e}")
            return False
    
    def send_text(self, hwnd, text):
        """发送文本，模拟键盘字符流，支持动态延迟调整"""
        if not self.win32con or not self.win32api:
            logger.error("win32模块未加载，无法发送文本")
            return False
        
        try:
            logger.info(f"向窗口 {hwnd} 发送文本: {text!r}")
            
            # 检测终端类型
            terminal_type = self.detect_terminal_type(hwnd)
            logger.info(f"检测到终端类型: {terminal_type}")
            
            # 基础延迟
            base_delay = 0.01
            
            # 根据终端类型调整发送策略
            if terminal_type in ["windows_terminal", "windows_terminal_powershell", "windows_terminal_cmd", "windows_terminal_wsl"]:
                # Windows Terminal特定发送策略 - 使用剪贴板+模拟按键方式
                return self.send_text_windows_terminal(hwnd, text)
            elif terminal_type == "powershell":
                # PowerShell特定发送策略
                return self.send_text_powershell(hwnd, text)
            elif terminal_type == "mobaxterm":
                # MobaXterm特定发送策略
                return self.send_text_mobaxterm(hwnd, text)
            elif terminal_type == "securecrt":
                # SecureCRT特定发送策略
                return self.send_text_securecrt(hwnd, text)
            elif terminal_type == "xshell":
                # Xshell特定发送策略
                return self.send_text_xshell(hwnd, text)
            else:
                # 通用发送策略
                return self.send_text_generic(hwnd, text)
        except Exception as e:
            logger.error(f"发送文本 {text!r} 失败: {e}")
            return False
    
    def send_text_securecrt(self, hwnd, text):
        """SecureCRT特定文本发送策略 - 增强版，确保可靠性"""
        try:
            terminal_type = "securecrt"
            logger.info(f"开始向SecureCRT窗口发送文本: {hwnd}, 文本: {text!r}")
            
            # 确保win32api和win32con已正确加载
            if not self.win32api or not self.win32con:
                logger.error(f"{terminal_type} win32api或win32con未正确加载，无法发送命令")
                return False
            
            # 增强：确保窗口在前台，提高发送成功率
            try:
                # 检查窗口是否在前台
                if win32gui.GetForegroundWindow() != hwnd:
                    logger.info(f"SecureCRT窗口不在前台，尝试获取焦点")
                    # 先显示窗口，确保可见
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                    time.sleep(0.1)
                    # 先发送WM_SETFOCUS消息
                    win32gui.PostMessage(hwnd, win32con.WM_SETFOCUS, 0, 0)
                    time.sleep(0.1)
                    # 再尝试设置前台窗口
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.3)  # 增加等待时间，确保窗口获得焦点
                    logger.info(f"SecureCRT窗口已获得焦点")
            except Exception as e:
                logger.warning(f"获取窗口焦点失败: {e}")
            
            # 直接使用传入的窗口句柄
            actual_hwnd = hwnd
            logger.info(f"使用SecureCRT窗口: {actual_hwnd}")
            
            # 增强：使用更可靠的延迟设置
            delay = 0.03  # 增加延迟，确保SecureCRT能处理每个字符
            logger.info(f"SecureCRT发送延迟设置: {delay} 秒")
            
            # 发送逻辑
            win32api = self.win32api
            win32con = self.win32con
            post_message = win32api.PostMessage
            wm_char = win32con.WM_CHAR
            wm_keydown = win32con.WM_KEYDOWN
            wm_keyup = win32con.WM_KEYUP
            sleep = time.sleep
            
            # 增强：使用完整的按键事件序列，确保每个字符都被正确处理
            for i, char in enumerate(text):
                char_code = ord(char)
                logger.debug(f"SecureCRT发送第 {i+1}/{len(text)} 个字符: {char!r}, 字符码: {char_code}")
                
                # 发送完整的按键事件序列
                # 1. 发送WM_KEYDOWN消息
                post_message(actual_hwnd, wm_keydown, char_code, 0)
                sleep(delay / 2)
                
                # 2. 发送WM_CHAR消息
                post_message(actual_hwnd, wm_char, char_code, 0)
                sleep(delay / 2)
                
                # 3. 发送WM_KEYUP消息
                post_message(actual_hwnd, wm_keyup, char_code, 0)
                sleep(delay)
            
            # 发送成功，尝试微调延迟
            self.adjust_delay(terminal_type, success=True)
            logger.info(f"SecureCRT文本发送成功: {text!r}")
            return True
        except Exception as e:
            logger.error(f"{terminal_type}文本发送失败: {e}")
            logger.debug(f"SecureCRT发送异常详情: {traceback.format_exc()}")
            # 发送失败，增加延迟
            self.adjust_delay(terminal_type, success=False)
            return False
    
    def send_text_xshell(self, hwnd, text):
        """Xshell特定文本发送策略 - 参考MobaXterm实现"""
        try:
            terminal_type = "xshell"
            logger.info(f"开始向Xshell窗口发送文本: {hwnd}, 文本: {text!r}")
            
            # 确保win32api和win32con已正确加载
            if not self.win32api or not self.win32con:
                logger.error(f"{terminal_type} win32api或win32con未正确加载，无法发送命令")
                return False
            
            # 从延迟缓存中获取当前终端的最佳延迟值
            base_delay = self.delay_cache.get(terminal_type, 0.01)
            logger.info(f"使用{terminal_type}延迟: {base_delay} 秒")
            
            # 参考MobaXterm的发送逻辑
            win32api = self.win32api
            win32con = self.win32con
            post_message = win32api.PostMessage
            wm_char = win32con.WM_CHAR
            sleep = time.sleep
            
            # 使用缓存的延迟
            delay = base_delay
            logger.info(f"Xshell发送延迟设置: {delay} 秒")
            
            for i, char in enumerate(text):
                char_code = ord(char)
                logger.debug(f"Xshell发送第 {i+1}/{len(text)} 个字符: {char!r}, 字符码: {char_code}")
                
                # 发送字符，包括空格，只发送WM_CHAR消息
                post_message(hwnd, wm_char, char_code, 0)
                
                # 根据字符类型动态调整延迟
                if char == ' ':
                    # 空格字符使用适中延迟
                    sleep(delay)
                else:
                    # 普通字符使用标准发送方式，使用适中延迟
                    sleep(delay)
            
            # 发送成功，尝试微调延迟，减少下一次发送的延迟
            self.adjust_delay(terminal_type, success=True)
            logger.info(f"Xshell文本发送成功: {text!r}")
            return True
        except Exception as e:
            logger.error(f"{terminal_type}文本发送失败: {e}")
            logger.debug(f"{terminal_type}发送异常详情: {traceback.format_exc()}")
            # 发送失败，增加延迟，确保下一次发送成功
            self.adjust_delay(terminal_type, success=False)
            return False
    
    def send_text_windows_terminal(self, hwnd, text):
        """Windows Terminal特定文本发送策略 - 使用剪贴板+模拟按键方式，只负责发送文本，不发送Enter键"""
        try:
            terminal_type = "windows_terminal"
            logger.info(f"开始向Windows Terminal窗口发送文本: {hwnd}, 文本: {text!r}")
            
            # 确保pyautogui和pyperclip可用
            if pyautogui is None or pyperclip is None:
                logger.error(f"{terminal_type} pyautogui或pyperclip未正确加载，无法发送命令")
                # 回退到通用发送策略
                return self.send_text_generic(hwnd, text)
            
            # 1. 确保窗口获得焦点
            if win32gui.GetForegroundWindow() != hwnd:
                logger.info(f"Windows Terminal窗口未获得焦点，尝试获取焦点")
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.1)
            
            # 2. 复制命令到剪贴板
            pyperclip.copy(text)
            time.sleep(0.05)
            logger.info(f"已将命令复制到剪贴板: {text!r}")
            
            # 3. 模拟按键Ctrl+V粘贴命令
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.1)
            logger.info(f"已模拟Ctrl+V粘贴命令")
            
            # 注意：不再硬编码发送Enter键，由调用方根据auto_enter设置决定是否发送
            
            # 发送成功，尝试微调延迟，减少下一次发送的延迟
            self.adjust_delay(terminal_type, success=True)
            logger.info(f"Windows Terminal文本发送成功: {text!r}")
            return True
        except Exception as e:
            logger.error(f"{terminal_type}文本发送失败: {e}")
            # 发送失败，增加延迟，确保下一次发送成功
            self.adjust_delay(terminal_type, success=False)
            # 回退到通用发送策略
            logger.info(f"Windows Terminal文本发送失败，回退到通用发送策略")
            return self.send_text_generic(hwnd, text)
    
    def send_text_generic(self, hwnd, text):
        """通用文本发送策略"""
        try:
            # 检测终端类型
            terminal_type = self.detect_terminal_type(hwnd)
            
            # 从延迟缓存中获取当前终端的最佳延迟值
            base_delay = self.delay_cache.get(terminal_type, 0.01)
            logger.info(f"使用延迟: {base_delay} 秒")
            
            # 优化：根据命令长度和复杂度动态调整延迟
            text_length = len(text)
            special_char_count = sum(1 for char in text if char in '!@#$%^&*()_+-=[]{}|;:,.<>?')
            
            # 根据文本长度调整延迟
            delay = base_delay
            if text_length < 10:
                delay = base_delay * 0.8
            elif text_length < 50:
                delay = base_delay * 0.6
            elif text_length < 100:
                delay = base_delay * 0.4
            else:
                delay = base_delay * 0.2
            
            # 如果包含较多特殊字符，适当增加延迟
            if special_char_count > text_length * 0.2:
                delay *= 1.2
            
            # 发送文本，逐个字符发送
            # 优化：使用局部变量减少属性查找
            win32api = self.win32api
            win32con = self.win32con
            post_message = win32api.PostMessage
            wm_char = win32con.WM_CHAR
            sleep = time.sleep
            
            # 批量处理普通字符，减少方法调用开销
            for char in text:
                char_code = ord(char)
                
                # 发送字符，包括空格，只发送WM_CHAR消息，避免重复空格
                post_message(hwnd, wm_char, char_code, 0)
                
                # 根据字符类型动态调整延迟
                if char == ' ':
                    # 空格字符使用适中延迟
                    sleep(delay)
                elif char in '!@#$%^&*()_+-=[]{}|;:,.<>?':
                    # 特殊字符使用较长延迟
                    sleep(delay * 1.5)
                elif char in '\t\n\r':
                    # 控制字符使用较长延迟
                    sleep(delay * 2.0)
                else:
                    # 普通字符使用更小的延迟
                    sleep(delay * 0.5)
            
            # 发送成功，尝试微调延迟，减少下一次发送的延迟
            self.adjust_delay(terminal_type, success=True)
            
            return True
        except Exception as e:
            logger.error(f"通用文本发送失败: {e}")
            # 发送失败，增加延迟，确保下一次发送成功
            self.adjust_delay(terminal_type, success=False)
            return False
    
    def adjust_delay(self, terminal_type, success=True):
        """根据发送结果调整延迟"""
        current_delay = self.delay_cache.get(terminal_type, 0.01)
        
        if success:
            # 发送成功，尝试减少延迟，提高发送速度
            new_delay = max(current_delay * 0.9, 0.005)  # 最小延迟0.005秒
        else:
            # 发送失败，增加延迟，确保下一次发送成功
            new_delay = min(current_delay * 1.5, 0.05)  # 最大延迟0.05秒
        
        if abs(new_delay - current_delay) > 0.001:  # 只有变化超过0.001秒时才更新
            self.delay_cache[terminal_type] = new_delay
            logger.info(f"调整 {terminal_type} 终端延迟: {current_delay:.4f} -> {new_delay:.4f}")
    
    def send_text_powershell(self, hwnd, text):
        """PowerShell特定文本发送策略 - 支持Windows Terminal和传统PowerShell"""
        try:
            terminal_type = "powershell"
            logger.info(f"开始向PowerShell窗口发送文本: {hwnd}, 文本: {text!r}")
            
            # 确保win32api和win32con已正确加载
            if not self.win32api or not self.win32con:
                logger.error(f"{terminal_type} win32api或win32con未正确加载，无法发送命令")
                return False
            
            # 获取实际的终端输入窗口
            # 对于Windows Terminal，需要获取子窗口作为实际输入区域
            actual_hwnd = hwnd
            
            # 检查窗口类名，判断是否为Windows Terminal
            window_class = self.win32gui.GetClassName(hwnd)
            logger.info(f"窗口类名: {window_class}")
            
            if window_class in ["CASCADIA_HOSTING_WINDOW_CLASS", "WindowsTerminal"]:
                logger.info(f"检测到Windows Terminal，正在查找实际输入窗口...")
                
                # 获取所有子窗口
                child_windows = []
                
                def enum_child_callback(child_hwnd, param):
                    param.append(child_hwnd)
                    return True
                
                self.win32gui.EnumChildWindows(hwnd, enum_child_callback, child_windows)
                logger.info(f"找到 {len(child_windows)} 个子窗口")
                
                # 遍历子窗口，查找合适的输入窗口
                for child_hwnd in child_windows:
                    child_class = self.win32gui.GetClassName(child_hwnd)
                    logger.info(f"子窗口 {child_hwnd}: 类名={child_class}")
                    # 选择第一个子窗口作为实际输入区域
                    actual_hwnd = child_hwnd
                    break
                
                logger.info(f"使用子窗口 {actual_hwnd} 作为实际输入区域")
            
            # 从延迟缓存中获取当前终端的最佳延迟值
            base_delay = self.delay_cache.get(terminal_type, 0.01)
            logger.info(f"使用{terminal_type}延迟: {base_delay} 秒")
            
            # 发送逻辑
            win32api = self.win32api
            win32con = self.win32con
            post_message = win32api.PostMessage
            wm_char = win32con.WM_CHAR
            sleep = time.sleep
            
            # 使用缓存的延迟
            delay = base_delay
            logger.info(f"PowerShell发送延迟设置: {delay} 秒")
            
            # 发送字符
            for i, char in enumerate(text):
                char_code = ord(char)
                logger.debug(f"PowerShell发送第 {i+1}/{len(text)} 个字符: {char!r}, 字符码: {char_code}")
                
                # 发送字符，包括空格，只发送WM_CHAR消息
                post_message(actual_hwnd, wm_char, char_code, 0)
                
                # 根据字符类型动态调整延迟
                if char == ' ':
                    # 空格字符使用适中延迟
                    sleep(delay)
                else:
                    # 普通字符使用标准发送方式，使用适中延迟
                    sleep(delay)
            
            # 发送成功，尝试微调延迟，减少下一次发送的延迟
            self.adjust_delay(terminal_type, success=True)
            logger.info(f"PowerShell文本发送成功: {text!r}")
            return True
        except Exception as e:
            logger.error(f"{terminal_type}文本发送失败: {e}")
            logger.debug(f"PowerShell发送异常详情: {traceback.format_exc()}")
            # 发送失败，增加延迟，确保下一次发送成功
            self.adjust_delay(terminal_type, success=False)
            return False
    
    def send_text_mobaxterm(self, hwnd, text):
        """MobaXterm特定文本发送策略"""
        try:
            terminal_type = "mobaxterm"
            logger.info(f"开始向MobaXterm窗口发送文本: {hwnd}, 文本: {text!r}")
            
            # 确保win32api和win32con已正确加载
            if not self.win32api or not self.win32con:
                logger.error(f"{terminal_type} win32api或win32con未正确加载，无法发送命令")
                return False
            
            # 从延迟缓存中获取当前终端的最佳延迟值
            base_delay = self.delay_cache.get(terminal_type, 0.01)
            logger.info(f"使用{terminal_type}延迟: {base_delay} 秒")
            
            # MobaXterm需要特别处理空格和换行符
            win32api = self.win32api
            win32con = self.win32con
            post_message = win32api.PostMessage
            wm_char = win32con.WM_CHAR
            sleep = time.sleep
            
            # 使用缓存的延迟
            delay = base_delay
            logger.info(f"MobaXterm发送延迟设置: {delay} 秒")
            
            for i, char in enumerate(text):
                char_code = ord(char)
                logger.debug(f"MobaXterm发送第 {i+1}/{len(text)} 个字符: {char!r}, 字符码: {char_code}")
                
                # 发送字符，包括空格，只发送WM_CHAR消息，避免重复空格
                post_message(hwnd, wm_char, char_code, 0)
                
                # 根据字符类型动态调整延迟
                if char == ' ':
                    # 空格字符使用适中延迟
                    sleep(delay)
                else:
                    # 普通字符使用标准发送方式，使用适中延迟
                    sleep(delay)
            
            # 发送成功，尝试微调延迟，减少下一次发送的延迟
            self.adjust_delay(terminal_type, success=True)
            logger.info(f"MobaXterm文本发送成功: {text!r}")
            return True
        except Exception as e:
            logger.error(f"MobaXterm文本发送失败: {e}")
            logger.debug(f"MobaXterm发送异常详情: {traceback.format_exc()}")
            # 发送失败，增加延迟，确保下一次发送成功
            self.adjust_delay(terminal_type, success=False)
            return False
    
    def send_enter(self, hwnd):
        """发送回车键，增强对PowerShell、MobaXterm和SecureCRT的兼容性"""
        if not self.win32con or not self.win32api:
            logger.error("win32模块未加载，无法发送回车键")
            return False
        
        try:
            logger.info(f"开始向窗口发送回车键: {hwnd}")
            # 检测终端类型
            terminal_type = self.detect_terminal_type(hwnd)
            logger.info(f"检测到终端类型: {terminal_type}")
            
            # 使用十六进制值0x0D代替win32con.VK_RETURN常量
            VK_RETURN = 0x0D
            logger.info(f"使用回车键值: {VK_RETURN} (0x{VK_RETURN:X})")
            
            # 根据终端类型调整回车键发送策略
            logger.info(f"终端类型: {terminal_type}, 检查 'windows_terminal' 是否在其中: {'windows_terminal' in terminal_type}")
            
            if "windows_terminal" in terminal_type:
                # Windows Terminal特定回车键发送策略
                logger.info(f"匹配到Windows Terminal终端类型，使用专门的回车键发送策略")
                result = self.send_enter_windows_terminal(hwnd)
                logger.info(f"Windows Terminal回车键发送结果: {result}")
                return result
            elif terminal_type == "powershell":
                # PowerShell特定回车键发送策略
                logger.info(f"匹配到PowerShell终端类型，使用专门的回车键发送策略")
                result = self.send_enter_powershell(hwnd)
                logger.info(f"PowerShell回车键发送结果: {result}")
                return result
            elif terminal_type == "mobaxterm":
                # MobaXterm特定回车键发送策略
                logger.info(f"匹配到MobaXterm终端类型，使用专门的回车键发送策略")
                result = self.send_enter_mobaxterm(hwnd)
                logger.info(f"MobaXterm回车键发送结果: {result}")
                return result
            elif terminal_type == "securecrt":
                # SecureCRT特定回车键发送策略
                logger.info(f"匹配到SecureCRT终端类型，使用专门的回车键发送策略")
                result = self.send_enter_securecrt(hwnd)
                logger.info(f"SecureCRT回车键发送结果: {result}")
                return result
            elif terminal_type == "xshell":
                # Xshell特定回车键发送策略
                logger.info(f"匹配到Xshell终端类型，使用专门的回车键发送策略")
                result = self.send_enter_xshell(hwnd)
                logger.info(f"Xshell回车键发送结果: {result}")
                return result
            else:
                # 通用回车键发送策略
                logger.info(f"未匹配到特定终端类型，使用通用回车键发送策略")
                result = self.send_enter_generic(hwnd)
                logger.info(f"通用回车键发送结果: {result}")
                return result
        except Exception as e:
            logger.error(f"发送回车键失败: {e}")
            logger.debug(f"发送回车键异常详情: {traceback.format_exc()}")
            return False
    
    def send_enter_securecrt(self, hwnd):
        """SecureCRT特定回车键发送策略 - 与MobaXterm完全一致"""
        try:
            VK_RETURN = 0x0D
            
            # 与MobaXterm完全一致的回车键发送逻辑
            self.win32api.PostMessage(hwnd, self.win32con.WM_KEYDOWN, VK_RETURN, 0)
            time.sleep(0.01)  # 短暂延迟
            
            # 只发送一个回车键，避免多个空行
            self.win32api.PostMessage(hwnd, self.win32con.WM_CHAR, ord('\r'), 0)
            time.sleep(0.01)
            
            self.win32api.PostMessage(hwnd, self.win32con.WM_KEYUP, VK_RETURN, 0)
            time.sleep(0.01)  # 短暂延迟
            
            return True
        except Exception as e:
            logger.error(f"SecureCRT回车键发送失败: {e}")
            return False
    
    def send_enter_xshell(self, hwnd):
        """Xshell特定回车键发送策略 - 参考MobaXterm实现"""
        try:
            VK_RETURN = 0x0D
            logger.info(f"开始向Xshell窗口发送回车键: {hwnd}")
            
            # 参考MobaXterm的回车键发送逻辑
            self.win32api.PostMessage(hwnd, self.win32con.WM_KEYDOWN, VK_RETURN, 0)
            time.sleep(0.01)  # 短暂延迟
            
            # 只发送一个回车键，避免多个空行
            self.win32api.PostMessage(hwnd, self.win32con.WM_CHAR, ord('\r'), 0)
            time.sleep(0.01)
            
            self.win32api.PostMessage(hwnd, self.win32con.WM_KEYUP, VK_RETURN, 0)
            time.sleep(0.01)  # 短暂延迟
            
            logger.info(f"Xshell回车键发送成功")
            return True
        except Exception as e:
            logger.error(f"Xshell回车键发送失败: {e}")
            logger.debug(f"Xshell回车键发送异常详情: {traceback.format_exc()}")
            return False
    
    def send_enter_generic(self, hwnd):
        """通用回车键发送策略"""
        try:
            VK_RETURN = 0x0D
            
            # 发送回车键的标准事件序列
            
            # 1. 发送WM_KEYDOWN消息
            self.win32api.PostMessage(hwnd, self.win32con.WM_KEYDOWN, VK_RETURN, 0)
            time.sleep(0.01)  # 短暂延迟
            
            # 2. 发送WM_CHAR消息
            self.win32api.PostMessage(hwnd, self.win32con.WM_CHAR, ord('\r'), 0)
            time.sleep(0.01)
            
            # 3. 发送WM_KEYUP消息
            self.win32api.PostMessage(hwnd, self.win32con.WM_KEYUP, VK_RETURN, 0)
            time.sleep(0.01)  # 短暂延迟
            
            return True
        except Exception as e:
            logger.error(f"通用回车键发送失败: {e}")
            return False
    
    def send_enter_windows_terminal(self, hwnd):
        """Windows Terminal特定回车键发送策略 - 确保命令正确执行"""
        try:
            VK_RETURN = 0x0D
            logger.info(f"开始向Windows Terminal窗口发送回车键: {hwnd}")
            
            # Windows Terminal需要可靠的回车键发送
            # 尝试多种方式确保命令执行
            
            # 方式1: 发送Windows API消息
            logger.info(f"使用Windows API方式发送回车键")
            self.win32api.PostMessage(hwnd, self.win32con.WM_KEYDOWN, VK_RETURN, 0)
            time.sleep(0.01)  # 短暂延迟
            
            self.win32api.PostMessage(hwnd, self.win32con.WM_CHAR, ord('\r'), 0)
            time.sleep(0.01)
            
            self.win32api.PostMessage(hwnd, self.win32con.WM_KEYUP, VK_RETURN, 0)
            time.sleep(0.01)  # 短暂延迟
            
            # 方式2: 额外使用pyautogui模拟回车键，确保命令执行
            # 只有当pyautogui可用时才使用此方式
            if pyautogui is not None:
                logger.info(f"使用pyautogui方式发送回车键")
                # 确保窗口获得焦点
                if win32gui.GetForegroundWindow() == hwnd:
                    pyautogui.press('enter')
                    time.sleep(0.05)  # 短暂延迟
                    logger.info(f"pyautogui回车键发送成功")
                else:
                    logger.warning(f"窗口未获得焦点，无法使用pyautogui发送回车键")
            
            logger.info(f"Windows Terminal回车键发送成功")
            return True
        except Exception as e:
            logger.error(f"Windows Terminal回车键发送失败: {e}")
            logger.debug(f"Windows Terminal回车键发送异常详情: {traceback.format_exc()}")
            return False
    
    def send_enter_powershell(self, hwnd):
        """PowerShell特定回车键发送策略 - 确保正确发送回车键"""
        try:
            VK_RETURN = 0x0D
            logger.info(f"开始向PowerShell窗口发送回车键: {hwnd}")
            
            # PowerShell需要可靠的回车键发送
            # 发送回车键的标准事件序列
            self.win32api.PostMessage(hwnd, self.win32con.WM_KEYDOWN, VK_RETURN, 0)
            time.sleep(0.01)  # 短暂延迟
            
            # 发送WM_CHAR消息，确保PowerShell接收到回车键
            self.win32api.PostMessage(hwnd, self.win32con.WM_CHAR, ord('\r'), 0)
            time.sleep(0.01)
            
            self.win32api.PostMessage(hwnd, self.win32con.WM_KEYUP, VK_RETURN, 0)
            time.sleep(0.01)  # 短暂延迟
            
            logger.info(f"PowerShell回车键发送成功")
            return True
        except Exception as e:
            logger.error(f"PowerShell回车键发送失败: {e}")
            logger.debug(f"PowerShell回车键发送异常详情: {traceback.format_exc()}")
            return False
    
    def send_enter_mobaxterm(self, hwnd):
        """MobaXterm特定回车键发送策略 - 只发送一次回车键"""
        try:
            VK_RETURN = 0x0D
            logger.info(f"开始向MobaXterm窗口发送回车键: {hwnd}")
            
            # MobaXterm只需要发送一次回车键，避免重复换行
            
            # 发送回车键的标准事件序列，但只发送一次
            self.win32api.PostMessage(hwnd, self.win32con.WM_KEYDOWN, VK_RETURN, 0)
            time.sleep(0.01)  # 短暂延迟
            
            self.win32api.PostMessage(hwnd, self.win32con.WM_KEYUP, VK_RETURN, 0)
            time.sleep(0.01)  # 短暂延迟
            
            logger.info(f"MobaXterm回车键发送成功")
            return True
        except Exception as e:
            logger.error(f"MobaXterm回车键发送失败: {e}")
            logger.debug(f"MobaXterm回车键发送异常详情: {traceback.format_exc()}")
            return False

class MouseSimulator:
    """鼠标模拟器，实现模拟鼠标选中窗口"""
    
    def __init__(self):
        self.win32con = None
        self.win32api = None
        self.load_win32_modules()
    
    def load_win32_modules(self):
        """加载win32模块"""
        try:
            import win32con
            import win32api
            import win32gui
            self.win32con = win32con
            self.win32api = win32api
            self.win32gui = win32gui
        except ImportError as e:
            logger.error(f"加载win32模块失败: {e}")
    
    def click(self, hwnd, x, y, use_simulated_click=True, saved_mouse_pos=None):
        """模拟鼠标点击，优化鼠标位置处理"""
        if not use_simulated_click:
            logger.info(f"跳过模拟点击: ({x}, {y})")
            return True
        
        if not self.win32con or not self.win32api:
            logger.error("win32模块未加载，无法发送鼠标事件")
            return False
        
        # 保存原始鼠标位置
        original_pos = saved_mouse_pos
        if original_pos is None and pyautogui is not None:
            try:
                original_pos = pyautogui.position()
                logger.info(f"保存原始鼠标位置: {original_pos}")
            except Exception as e:
                logger.warning(f"获取原始鼠标位置失败: {e}")
        
        try:
            # 如果hwnd不为None，使用窗口消息发送点击事件，不影响实际鼠标位置
            if hwnd is not None:
                # 确保窗口句柄有效
                if not self.win32gui or not self.win32gui.IsWindow(hwnd):
                    logger.warning(f"窗口句柄无效，跳过模拟点击: ({x}, {y})")
                    return True
                
                # 发送WM_LBUTTONDOWN消息
                self.win32api.PostMessage(hwnd, self.win32con.WM_LBUTTONDOWN, self.win32con.MK_LBUTTON, self.win32api.MAKELONG(x, y))
                time.sleep(0.01)  # 短暂延迟
                
                # 发送WM_LBUTTONUP消息
                self.win32api.PostMessage(hwnd, self.win32con.WM_LBUTTONUP, 0, self.win32api.MAKELONG(x, y))
                time.sleep(0.01)  # 短暂延迟
                
                logger.info(f"使用窗口消息模拟点击: ({x}, {y})")
                return True
            
            # 只有在hwnd为None且确实需要全局点击时，才使用pyautogui移动鼠标
            # 检查pyautogui是否可用
            if pyautogui is None:
                logger.error("pyautogui模块未安装，无法使用鼠标模拟方式")
                return False
            
            try:
                # 检查当前位置是否已经接近目标位置，避免不必要的移动
                current_x, current_y = pyautogui.position()
                target_x, target_y = x, y
                
                # 只有当鼠标位置与目标位置相差超过5像素时，才移动鼠标
                if abs(current_x - target_x) > 5 or abs(current_y - target_y) > 5:
                    logger.info(f"当前鼠标位置: ({current_x}, {current_y})，目标位置: ({target_x}, {target_y})，需要移动")
                    
                    # 移动到目标位置并点击
                    pyautogui.moveTo(target_x, target_y)
                    time.sleep(0.01)
                    pyautogui.click()
                    time.sleep(0.01)
                    
                    # 恢复原鼠标位置
                    if original_pos is not None:
                        pyautogui.moveTo(original_pos[0], original_pos[1])
                        logger.info(f"恢复鼠标位置: {original_pos}")
                    else:
                        pyautogui.moveTo(current_x, current_y)
                        logger.info(f"恢复鼠标位置: ({current_x}, {current_y})")
                    time.sleep(0.01)
                    
                    logger.info(f"使用pyautogui模拟全局点击并恢复位置: ({target_x}, {target_y})")
                else:
                    # 鼠标位置已经接近目标位置，直接点击，不移动鼠标
                    pyautogui.click()
                    time.sleep(0.01)
                    logger.info(f"使用pyautogui模拟局部点击: ({target_x}, {target_y})")
            except Exception as e:
                logger.error(f"使用pyautogui模拟点击失败: {e}")
                # 无论如何都尝试恢复鼠标位置
                if original_pos is not None:
                    try:
                        pyautogui.moveTo(original_pos[0], original_pos[1])
                        logger.info(f"异常情况下恢复鼠标位置: {original_pos}")
                    except Exception as restore_e:
                        logger.warning(f"异常情况下恢复鼠标位置失败: {restore_e}")
                return False
            
            return True
        except Exception as e:
            logger.error(f"发送鼠标点击事件失败: {e}")
            # 无论如何都尝试恢复鼠标位置
            if original_pos is not None and pyautogui is not None:
                try:
                    pyautogui.moveTo(original_pos[0], original_pos[1])
                    logger.info(f"异常情况下恢复鼠标位置: {original_pos}")
                except Exception as restore_e:
                    logger.warning(f"异常情况下恢复鼠标位置失败: {restore_e}")
            return False

class SerialManager:
    """串口管理器"""
    
    def __init__(self):
        self.serial_port = None
        self.connected = False
        self.current_port = None
        self.baudrate = 9600
        self.timeout = 1
    
    def get_available_ports(self) -> List[Dict]:
        """获取可用串口列表"""
        ports = []
        try:
            for port in serial.tools.list_ports.comports():
                ports.append({
                    'device': port.device,
                    'description': port.description,
                    'hwid': port.hwid,
                    'display_name': f"{port.device} - {port.description}"
                })
        except Exception as e:
            logger.error(f"获取串口列表失败: {e}")
        
        return ports
    
    def connect(self, port: str, baudrate: int = 9600) -> bool:
        """连接串口"""
        try:
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.close()
            
            self.serial_port = serial.Serial(
                port=port,
                baudrate=baudrate,
                timeout=self.timeout
            )
            
            self.connected = True
            self.current_port = port
            self.baudrate = baudrate
            logger.info(f"串口 {port} 连接成功，波特率: {baudrate}")
            return True
        except Exception as e:
            logger.error(f"串口连接失败: {e}")
            self.connected = False
            return False
    
    def disconnect(self):
        """断开串口连接"""
        try:
            if self.serial_port and self.serial_port.is_open:
                self.serial_port.close()
                self.connected = False
                self.current_port = None
                logger.info("串口已断开")
        except Exception as e:
            logger.error(f"断开串口失败: {e}")
    
    def send_command(self, command: str) -> bool:
        """发送命令到串口"""
        if not self.connected or not self.serial_port:
            logger.error("串口未连接")
            return False
        
        try:
            # 确保命令以换行符结尾
            if not command.endswith('\n'):
                command += '\n'
            
            self.serial_port.write(command.encode('utf-8'))
            logger.info(f"命令已发送到串口: {command.strip()}")
            return True
        except Exception as e:
            logger.error(f"串口发送命令失败: {e}")
            return False
    
    def is_connected(self) -> bool:
        """检查串口是否已连接"""
        return self.connected and self.serial_port and self.serial_port.is_open

class ConfigManager:
    """配置管理器"""
    
    def __init__(self, config_file: str = 'command_sender_config.json'):
        self.config_file = config_file
        self.config = self.load_config()
    
    def load_config(self) -> Dict:
        """加载配置"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            logger.error(f"加载配置失败: {e}")
        
        # 返回默认配置
        return {
            'serial_port': '',
            'baudrate': 9600,
            'output_mode': 'clipboard',  # serial, clipboard, keyboard
            'target_window': None,
            'last_file': '',
            'window_geometry': '',
            'always_on_top': False,
            'auto_save': True,
            'recent_files': [],
            # 命令发送配置
            'simulate_keyboard': True,  # 默认使用模拟键盘输入
            'keyboard_delay': 100,  # 模拟键盘输入延迟，单位毫秒
            'auto_enter': True  # 默认自动换行执行命令
        }
    
    def save_config(self):
        """保存配置"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            logger.info("配置保存成功")
        except Exception as e:
            logger.error(f"保存配置失败: {e}")
    
    def get(self, key: str, default=None):
        """获取配置项"""
        return self.config.get(key, default)
    
    def set(self, key: str, value):
        """设置配置项"""
        self.config[key] = value
    
    def add_recent_file(self, file_path: str):
        """添加最近打开的文件"""
        recent_files = self.config.get('recent_files', [])
        
        # 如果文件已在列表中，先移除
        if file_path in recent_files:
            recent_files.remove(file_path)
        
        # 添加到列表开头
        recent_files.insert(0, file_path)
        
        # 限制最近文件数量
        if len(recent_files) > 10:
            recent_files = recent_files[:10]
        
        self.config['recent_files'] = recent_files
        self.save_config()

class CommandSenderApp:
    """命令发送器主应用"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("命令发送器")
        self.root.geometry("800x600")
        
        # 初始化主题颜色（使用最原始的简单配置）
        self.theme_colors = {
            'primary_bg': '#ffffff',
            'secondary_bg': '#ffffff',
            'accent_bg': '#0078d4',
            'text_fg': '#000000',
            'button_bg': '#0078d4',
            'button_fg': '#ffffff',
            'button_hover_bg': '#005a9e',
            'frame_bg': '#ffffff',
            'editor_bg': '#ffffff',  # 纯白色背景
            'editor_fg': '#000000',  # 深黑色文字
            'editor_cursor': '#000000',  # 黑色光标
            'editor_selection': '#c6e0fb',  # 默认蓝色选择背景
            'line_numbers_bg': '#f0f0f0',  # 浅灰色行号背景
            'line_numbers_fg': '#808080',  # 灰色行号
            'current_line_bg': '#f0f0f0',  # 浅灰色当前行背景
            'status_bar_bg': '#f0f0f0',
            'status_bar_fg': '#000000',
            'highlight_bg': '#ffff00',
            'success_color': '#00b250',
            'error_color': '#ff0000',
            'warning_color': '#ff8c00'
        }
        
        # 初始化变量
        self.serial_manager = SerialManager()
        self.config_manager = ConfigManager()
        self.current_file = None
        self.is_modified = False
        self.auto_save_timer = None
        self.drag_start_line = None
        self.always_on_top_var = tk.BooleanVar(value=False)  # 初始化置顶变量
        
        # 文件监控相关变量
        self.current_file_mtime = None  # 当前文件的修改时间
        self.file_monitor_interval = 1000  # 文件监控间隔（毫秒）
        self.file_monitor_timer = None  # 文件监控定时器
        self.is_editing = False  # 标记是否正在编辑文件
        
        # 窗口移动相关变量
        self.drag_data = {"x": 0, "y": 0}
        
        # 初始化UI变量
        self.output_mode = tk.StringVar(value='clipboard')
        self.port_var = tk.StringVar()
        self.baudrate_var = tk.StringVar(value='9600')
        # 提前初始化file_info_var，避免使用时出现属性错误
        self.file_info_var = tk.StringVar(value="未打开文件")
        # 初始化自动回车变量
        self.auto_enter_var = tk.BooleanVar(value=self.config_manager.get('auto_enter', True))
        # 初始化焦点管理配置变量
        self.focus_management_var = tk.StringVar(value=self.config_manager.get('focus_management_strategy', 'aggressive'))
        self.focus_retry_count_var = tk.IntVar(value=self.config_manager.get('focus_retry_count', 3))
        self.focus_retry_delay_var = tk.DoubleVar(value=self.config_manager.get('focus_retry_delay', 0.1))
        self.focus_timeout_var = tk.DoubleVar(value=self.config_manager.get('focus_timeout', 10.0))
        
        # 初始化窗口选择器
        self.window_selector = WindowSelector()
        
        # 创建UI
        try:
            self.create_ui()
            logger.info("UI创建成功")
        except Exception as e:
            logger.error(f"创建UI失败: {e}")
            print(f"创建UI失败: {e}")
        
        # 绑定事件
        try:
            self.bind_events()
            logger.info("事件绑定成功")
        except Exception as e:
            logger.error(f"事件绑定失败: {e}")
            print(f"事件绑定失败: {e}")
        
        # 加载设置
        try:
            self.load_settings()
            logger.info("设置加载成功")
        except Exception as e:
            logger.error(f"加载设置失败: {e}")
            print(f"加载设置失败: {e}")
        
        # 加载上次打开的文件
        try:
            self.load_last_file()
            logger.info("上次文件加载成功")
        except Exception as e:
            logger.error(f"加载上次文件失败: {e}")
            print(f"加载上次文件失败: {e}")
        
        logger.info("CommandSenderApp初始化完成")
    
    def create_ui(self):
        """创建用户界面"""
        print("开始创建UI...")
        # 创建菜单栏
        self.create_menu()
        print("菜单栏创建完成")
        
        # 创建工具栏
        self.create_toolbar()
        print("工具栏创建完成")
        
        # 创建文本编辑器
        self.create_text_editor(self.root)
        print("文本编辑器创建完成")
        
        # 创建目标选择区域
        self.create_target_selection()
        print("目标选择区域创建完成")
        
        # 创建宏记录和回放面板
        self.create_macro_panel()
        print("宏面板创建完成")
        
        # 创建状态栏
        self.create_status_bar()
        print("状态栏创建完成")
        
        # 初始化命令计数器
        self.sent_count = 0
        self.failed_count = 0
        print("UI创建完成")
    
    def create_menu(self):
        """创建菜单栏"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件(F)", menu=file_menu)
        file_menu.add_command(label="新建", command=self.new_file, accelerator="Ctrl+N")
        file_menu.add_command(label="打开", command=self.open_file, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="保存", command=self.save_file, accelerator="Ctrl+S")
        file_menu.add_command(label="另存为", command=self.save_file_as, accelerator="Ctrl+Shift+S")
        file_menu.add_separator()
        
        # 最近文件子菜单
        self.recent_files_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="最近文件", menu=self.recent_files_menu)
        self.update_recent_files_menu()
        
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.on_closing)
        
        # 编辑菜单
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="编辑(E)", menu=edit_menu)
        edit_menu.add_command(label="撤销", command=self.undo, accelerator="Ctrl+Z")
        edit_menu.add_command(label="重做", command=self.redo, accelerator="Ctrl+Y")
        edit_menu.add_separator()
        edit_menu.add_command(label="剪切", command=self.cut, accelerator="Ctrl+X")
        edit_menu.add_command(label="复制", command=self.copy, accelerator="Ctrl+C")
        edit_menu.add_command(label="粘贴", command=self.paste, accelerator="Ctrl+V")
        edit_menu.add_separator()
        edit_menu.add_command(label="查找", command=self.find, accelerator="Ctrl+F")
        edit_menu.add_command(label="替换", command=self.replace, accelerator="Ctrl+H")
        
        # 工具菜单
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="工具(T)", menu=tools_menu)
        tools_menu.add_command(label="刷新窗口列表", command=self.refresh_window_list)
        tools_menu.add_command(label="刷新串口列表", command=self.refresh_serial_list)
        tools_menu.add_separator()
        tools_menu.add_command(label="设置", command=self.show_settings)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助(H)", menu=help_menu)
        help_menu.add_command(label="关于", command=self.show_about)
        
        # 置顶选项 - 直接添加到菜单栏
        self.always_on_top_var = tk.BooleanVar(value=self.config_manager.get('always_on_top', False))
        menubar.add_checkbutton(label="📌置顶", command=self.toggle_always_on_top, variable=self.always_on_top_var)
    
    def create_toolbar(self):
        """创建工具栏"""
        # 使用默认的ttk样式，避免复杂的自定义样式
        toolbar = ttk.Frame(self.root)
        toolbar.pack(fill=tk.X, padx=5, pady=(0, 5))
        
        # 文件操作按钮
        ttk.Button(toolbar, text="新建", command=self.new_file).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(toolbar, text="打开", command=self.open_file).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="保存", command=self.save_file).pack(side=tk.LEFT, padx=2)
        
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 编辑操作按钮
        ttk.Button(toolbar, text="剪切", command=self.cut).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="复制", command=self.copy).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="粘贴", command=self.paste).pack(side=tk.LEFT, padx=2)
        
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 发送按钮
        ttk.Button(toolbar, text="发送当前行", command=self.send_current_line).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="发送选中文本", command=self.send_selected_text).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="发送全部内容", command=self.send_all_content).pack(side=tk.LEFT, padx=2)
    
    def create_send_options(self, parent):
        """创建发送选项面板"""
        # 发送方式选择
        send_frame = ttk.LabelFrame(parent, text="发送方式")
        send_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.output_mode = tk.StringVar(value=self.config_manager.get('output_mode', 'clipboard'))
        
        ttk.Radiobutton(send_frame, text="剪贴板", variable=self.output_mode, 
                       value="clipboard", command=self.on_output_mode_change).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Radiobutton(send_frame, text="终端输入", variable=self.output_mode, 
                       value="terminal", command=self.on_output_mode_change).pack(anchor=tk.W, padx=5, pady=2)
        ttk.Radiobutton(send_frame, text="串口", variable=self.output_mode, 
                       value="serial", command=self.on_output_mode_change).pack(anchor=tk.W, padx=5, pady=2)
        
        # 焦点管理设置
        focus_frame = ttk.LabelFrame(parent, text="焦点管理")
        focus_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 焦点管理策略
        ttk.Label(focus_frame, text="焦点管理策略:").pack(anchor=tk.W, padx=5, pady=2)
        strategy_frame = ttk.Frame(focus_frame)
        strategy_frame.pack(anchor=tk.W, padx=15, pady=2)
        
        ttk.Radiobutton(strategy_frame, text="激进模式", variable=self.focus_management_var, 
                       value="aggressive", command=self.on_focus_strategy_change).pack(anchor=tk.W, padx=5, pady=1)
        ttk.Radiobutton(strategy_frame, text="保守模式", variable=self.focus_management_var, 
                       value="conservative", command=self.on_focus_strategy_change).pack(anchor=tk.W, padx=5, pady=1)
        ttk.Radiobutton(strategy_frame, text="手动模式", variable=self.focus_management_var, 
                       value="manual", command=self.on_focus_strategy_change).pack(anchor=tk.W, padx=5, pady=1)
        
        # 焦点重试设置
        retry_frame = ttk.Frame(focus_frame)
        retry_frame.pack(fill=tk.X, padx=15, pady=5)
        
        ttk.Label(retry_frame, text="重试次数:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        retry_spinbox = ttk.Spinbox(retry_frame, from_=1, to=10, textvariable=self.focus_retry_count_var, width=5, 
                                   command=self.on_focus_retry_change)
        retry_spinbox.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(retry_frame, text="初始重试延迟(秒):").grid(row=0, column=2, sticky=tk.W, padx=15, pady=2)
        retry_delay_entry = ttk.Entry(retry_frame, textvariable=self.focus_retry_delay_var, width=8)
        retry_delay_entry.grid(row=0, column=3, sticky=tk.W, padx=5, pady=2)
        
        ttk.Label(retry_frame, text="超时时间(秒):").grid(row=0, column=4, sticky=tk.W, padx=15, pady=2)
        timeout_entry = ttk.Entry(retry_frame, textvariable=self.focus_timeout_var, width=8)
        timeout_entry.grid(row=0, column=5, sticky=tk.W, padx=5, pady=2)
        
        # 串口设置
        self.serial_frame = ttk.LabelFrame(parent, text="串口设置")
        self.serial_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 串口选择
        ttk.Label(self.serial_frame, text="串口:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(self.serial_frame, textvariable=self.port_var, state="readonly", width=15)
        self.port_combo.grid(row=0, column=1, padx=5, pady=2)
        self.refresh_serial_list()
        
        # 波特率选择
        ttk.Label(self.serial_frame, text="波特率:").grid(row=1, column=0, sticky=tk.W, padx=5, pady=2)
        self.baudrate_var = tk.StringVar(value=str(self.config_manager.get('baudrate', 9600)))
        baudrate_combo = ttk.Combobox(self.serial_frame, textvariable=self.baudrate_var, 
                                     values=['9600', '19200', '38400', '57600', '115200'], 
                                     state="readonly", width=15)
        baudrate_combo.grid(row=1, column=1, padx=5, pady=2)
        
        # 连接按钮
        self.connect_btn = ttk.Button(self.serial_frame, text="连接", command=self.toggle_serial_connection)
        self.connect_btn.grid(row=2, column=0, columnspan=2, pady=5)
        
        # 窗口选择已在create_target_selection方法中处理
        
        # 根据当前输出模式显示/隐藏相关设置
        self.on_output_mode_change()
    
    def create_text_editor(self, parent):
        """创建文本编辑器"""
        # 创建带行号的文本编辑器
        editor_frame = ttk.Frame(parent)
        editor_frame.pack(fill=tk.BOTH, expand=True)
        
        # 应用主题颜色
        colors = self.theme_colors
        
        # 行号显示
        self.line_numbers = tk.Text(editor_frame, width=4, padx=3, takefocus=0,
                                   border=0, state='disabled', wrap='none',
                                   background=colors['line_numbers_bg'],
                                   foreground=colors['line_numbers_fg'],
                                   font=('Consolas', 10))
        self.line_numbers.pack(side=tk.LEFT, fill=tk.Y)
        
        # 新增：发送按钮区域，使用Canvas精确控制按钮位置
        self.send_buttons_canvas = tk.Canvas(editor_frame, width=40, bg=colors['line_numbers_bg'], highlightthickness=0)
        self.send_buttons_canvas.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 5))
        
        # 文本编辑区
        self.text_editor = scrolledtext.ScrolledText(editor_frame, wrap=tk.NONE, undo=True,
                                                   background=colors['editor_bg'],
                                                   foreground=colors['editor_fg'],
                                                   insertbackground=colors['editor_cursor'],
                                                   selectbackground=colors['editor_selection'],
                                                   font=('Consolas', 10))
        self.text_editor.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 配置标签样式
        self.text_editor.tag_configure("current_line", background=colors['current_line_bg'])
        
        # 绑定事件
        self.text_editor.bind("<KeyRelease>", self.on_text_change)
        self.text_editor.bind("<Button-1>", self.on_text_click)
        self.text_editor.bind("<MouseWheel>", self.on_mouse_wheel)
        self.text_editor.bind("<Button-4>", self.on_mouse_wheel)  # Linux
        self.text_editor.bind("<Button-5>", self.on_mouse_wheel)  # Linux
        
        # 同步滚动
        self.text_editor.bind("<MouseWheel>", self.sync_scroll)
        self.text_editor.bind("<Button-4>", self.sync_scroll)  # Linux
        self.text_editor.bind("<Button-5>", self.sync_scroll)  # Linux
        self.text_editor.bind("<Key>", self.sync_scroll)
        self.text_editor.bind("<Button-1>", self.sync_scroll)
        
        # 添加滚动事件监听器，确保拖动滚动条时按钮位置也能更新
        self.text_editor.bind("<Configure>", self.sync_scroll)
        
        # 绑定点击事件，用于显示发送按钮
        self.text_editor.bind("<Button-1>", self.on_text_click, add="+")
        
        # 初始化发送按钮相关变量
        self.send_buttons = {}  # 存储所有发送按钮的字典
        self.current_visible_line = None  # 当前可见的按钮行
        
        # 更新行号
        self.update_line_numbers()
        # 更新发送按钮
        self.update_send_buttons()
        
        # 移除可能导致按钮隐藏的不必要事件绑定
        # 只在点击事件时显示按钮，不需要其他事件触发
    
    def create_target_selection(self):
        """创建目标选择区域"""
        # 使用默认的ttk样式，避免复杂的自定义样式
        target_frame = ttk.Frame(self.root)
        target_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 目标标签
        ttk.Label(target_frame, text="目标:").pack(side=tk.LEFT, padx=5, pady=3)
        
        # 拖拽选择按钮 - 使用🎯图标
        self.drag_select_btn = ttk.Button(
            target_frame, 
            text="拖拽选择", 
            command=self.start_drag_select
        )
        self.drag_select_btn.pack(side=tk.LEFT, padx=5, pady=3)
        
        # 只显示拖拽选择，不显示下拉框
        # 移除或隐藏下拉框
        
        # 添加自动换行选项，放在目标选择区域的右侧，使其更加显眼
        auto_enter_frame = ttk.Frame(target_frame)
        auto_enter_frame.pack(side=tk.RIGHT, padx=5, pady=3)
        
        ttk.Label(auto_enter_frame, text="自动换行执行:").pack(side=tk.LEFT, padx=5, pady=3)
        auto_enter_check = ttk.Checkbutton(auto_enter_frame, variable=self.auto_enter_var, 
                                          command=self.on_auto_enter_change)
        auto_enter_check.pack(side=tk.LEFT, padx=5, pady=3)
        auto_enter_check.bind("<Enter>", lambda e: self.show_tooltip(e, "勾选则命令发送后自动执行，不勾选则需要手动按Enter键执行"))
        auto_enter_check.bind("<Leave>", lambda e: self.hide_tooltip())
        
        return target_frame
    

        

    

    

    

    
    def create_macro_panel(self):
        """创建宏记录和回放面板（MobaXterm风格）"""
        # 使用默认样式，避免过于复杂的自定义样式
        macro_frame = ttk.LabelFrame(self.root, text="宏记录和回放")
        macro_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 初始化宏相关变量
        self.is_recording = False
        self.recorded_macro = []
        self.recording_start_time = 0
        
        # 宏控制按钮
        control_frame = ttk.Frame(macro_frame)
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 开始记录按钮
        self.record_btn = ttk.Button(control_frame, text="开始记录", command=self.start_macro_recording)
        self.record_btn.pack(side=tk.LEFT, padx=5)
        
        # 停止记录按钮
        self.stop_btn = ttk.Button(control_frame, text="停止记录", command=self.stop_macro_recording, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
        # 回放宏按钮
        self.play_btn = ttk.Button(control_frame, text="回放宏", command=self.play_macro, state=tk.DISABLED)
        self.play_btn.pack(side=tk.LEFT, padx=5)
        
        # 保存宏按钮
        self.save_macro_btn = ttk.Button(control_frame, text="保存宏", command=self.save_macro, state=tk.DISABLED)
        self.save_macro_btn.pack(side=tk.LEFT, padx=5)
        
        # 加载宏按钮
        self.load_macro_btn = ttk.Button(control_frame, text="加载宏", command=self.load_macro)
        self.load_macro_btn.pack(side=tk.LEFT, padx=5)
        
        # 宏名称输入
        ttk.Label(control_frame, text="宏名称:").pack(side=tk.LEFT, padx=5)
        self.macro_name_var = tk.StringVar(value="my_macro")
        self.macro_name_entry = ttk.Entry(control_frame, textvariable=self.macro_name_var, width=20)
        self.macro_name_entry.pack(side=tk.LEFT, padx=5)
        
        # 宏记录状态
        self.recording_status_var = tk.StringVar(value="")
        ttk.Label(control_frame, textvariable=self.recording_status_var, foreground="red").pack(side=tk.RIGHT, padx=5)
    
    def create_status_bar(self):
        """创建状态栏"""
        colors = self.theme_colors
        
        self.status_bar = ttk.Frame(self.root)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        
        # 设置状态栏背景色
        self.status_bar.configure(style='StatusBar.TFrame')
        
        # 创建自定义样式
        style = ttk.Style()
        style.configure('StatusBar.TLabel', 
                        background=colors['status_bar_bg'], 
                        foreground=colors['status_bar_fg'],
                        font=('Consolas', 9))
        
        # 定义发送按钮样式
        style.configure('SendButton.TButton', 
                        padding=(2, 2),  # 适当的内边距
                        font=('Arial', 9),  # 清晰的字体和大小
                        foreground=colors['accent_bg'],  # 使用主题强调色
                        background=colors['line_numbers_bg'])  # 与行号区域背景一致
        
        # 状态信息
        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(self.status_bar, textvariable=self.status_var, style='StatusBar.TLabel').pack(side=tk.LEFT, padx=5)
        
        # 分隔符
        ttk.Separator(self.status_bar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 文件信息
        self.file_info_var = tk.StringVar(value="未打开文件")
        ttk.Label(self.status_bar, textvariable=self.file_info_var, style='StatusBar.TLabel').pack(side=tk.LEFT, padx=5)
        
        # 分隔符
        ttk.Separator(self.status_bar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 发送计数（使用成功颜色）
        self.sent_count_var = tk.StringVar(value="已发送: 0")
        ttk.Label(self.status_bar, textvariable=self.sent_count_var, style='StatusBar.TLabel', foreground=colors['success_color']).pack(side=tk.LEFT, padx=5)
        
        # 分隔符
        ttk.Separator(self.status_bar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 失败计数（使用错误颜色）
        self.failed_count_var = tk.StringVar(value="失败: 0")
        ttk.Label(self.status_bar, textvariable=self.failed_count_var, style='StatusBar.TLabel', foreground=colors['error_color']).pack(side=tk.LEFT, padx=5)
        
        # 分隔符
        ttk.Separator(self.status_bar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 连接状态（使用强调颜色）
        self.connection_var = tk.StringVar(value="未连接")
        ttk.Label(self.status_bar, textvariable=self.connection_var, style='StatusBar.TLabel', foreground=colors['accent_bg']).pack(side=tk.RIGHT, padx=5)
    
    def bind_events(self):
        """绑定事件"""
        # 窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 窗口大小改变事件
        self.root.bind("<Configure>", self.on_window_resize)
        
        # 快捷键
        self.root.bind("<Control-n>", lambda e: self.new_file())
        self.root.bind("<Control-o>", lambda e: self.open_file())
        self.root.bind("<Control-s>", lambda e: self.save_file())
        self.root.bind("<Control-Shift-S>", lambda e: self.save_file_as())
        self.root.bind("<Control-f>", lambda e: self.find())
        self.root.bind("<Control-h>", lambda e: self.replace())
        
        # ESC键取消操作
        self.root.bind("<Escape>", self._on_escape)
        
        # 窗口拖动功能 - 点击标题栏位置
        self.root.bind("<Button-1>", self._on_start_drag)
        self.root.bind("<B1-Motion>", self._on_drag)
        self.root.bind("<ButtonRelease-1>", self._on_stop_drag)
    
    def on_window_scale(self, event=None):
        """窗口缩放事件处理，支持自适应缩放"""
        try:
            # 调用现有的窗口大小改变处理
            self.on_window_resize(event)
            logger.debug("窗口缩放事件处理完成")
        except Exception as e:
            logger.error(f"处理窗口缩放事件失败: {str(e)}")
    
    def load_settings(self):
        """加载设置"""
        # 设置窗口位置和大小
        geometry = self.config_manager.get('window_geometry')
        if geometry:
            try:
                self.root.geometry(geometry)
            except:
                pass
        
        # 设置置顶状态
        always_on_top = self.config_manager.get('always_on_top', False)
        self.always_on_top_var.set(always_on_top)
        self.toggle_always_on_top()
        
        # 设置串口
        port = self.config_manager.get('serial_port', '')
        if port:
            self.port_var.set(port)
        
        # 设置波特率
        baudrate = self.config_manager.get('baudrate', 9600)
        self.baudrate_var.set(str(baudrate))
        
        # 清除之前可能保存的无效窗口选择
        # 避免应用程序启动时使用无效的窗口句柄
        self.window_selector.selected_window = None
        
        # 仅当connection_var属性存在时才设置它
        if hasattr(self, 'connection_var'):
            self.connection_var.set("未连接")
        
        # 仅当status_var属性存在时才设置它
        if hasattr(self, 'status_var'):
            self.status_var.set("就绪")
    
    def load_last_file(self):
        """加载上次打开的文件"""
        try:
            # 读取配置中的上次文件路径
            last_file = self.config_manager.get('last_file', '')
            logger.info(f"尝试加载上次文件: {last_file}")
            
            if not last_file:
                logger.info("配置中没有上次打开的文件")
                return
            
            # 检查文件是否存在
            if os.path.exists(last_file):
                # 检查文件是否可读
                if os.access(last_file, os.R_OK):
                    logger.info(f"成功找到上次文件，尝试打开: {last_file}")
                    self.open_file(last_file)
                    logger.info(f"成功打开上次文件: {last_file}")
                else:
                    logger.error(f"上次文件不可读: {last_file}")
            else:
                logger.warning(f"上次文件不存在: {last_file}")
                # 清除不存在的文件路径
                self.config_manager.set('last_file', '')
                self.config_manager.save_config()
        except Exception as e:
            logger.error(f"加载上次文件失败: {str(e)}")
            print(f"Error loading last file: {e}")
            # 清除可能损坏的配置
            self.config_manager.set('last_file', '')
            self.config_manager.save_config()
    
    def on_output_mode_change(self):
        """输出模式改变事件"""
        mode = self.output_mode.get()
        self.config_manager.set('output_mode', mode)
        
        # 根据模式显示/隐藏串口设置
        if mode == 'serial':
            self.serial_frame.pack(fill=tk.X, padx=5, pady=5, after=self.serial_frame.master.winfo_children()[0])
        else:
            self.serial_frame.pack_forget()
    
    def on_auto_enter_change(self):
        """自动换行设置改变事件"""
        auto_enter = self.auto_enter_var.get()
        self.config_manager.set('auto_enter', auto_enter)
        self.config_manager.save_config()
    
    def on_focus_strategy_change(self):
        """焦点管理策略改变事件"""
        strategy = self.focus_management_var.get()
        self.config_manager.set('focus_management_strategy', strategy)
        self.config_manager.save_config()
        logger.info(f"焦点管理策略已更新为: {strategy}")
    
    def on_focus_retry_change(self):
        """焦点重试设置改变事件"""
        retry_count = self.focus_retry_count_var.get()
        retry_delay = self.focus_retry_delay_var.get()
        timeout = self.focus_timeout_var.get()
        
        self.config_manager.set('focus_retry_count', retry_count)
        self.config_manager.set('focus_retry_delay', retry_delay)
        self.config_manager.set('focus_timeout', timeout)
        self.config_manager.save_config()
        
        logger.info(f"焦点重试设置已更新 - 次数: {retry_count}, 延迟: {retry_delay}, 超时: {timeout}")
    
    def refresh_serial_list(self):
        """刷新串口列表"""
        ports = self.serial_manager.get_available_ports()
        port_names = [p['display_name'] for p in ports]
        self.port_combo['values'] = port_names
        
        # 如果当前选择的端口不在列表中，清空选择
        current_port = self.port_var.get()
        if current_port and current_port not in port_names:
            self.port_var.set('')
    
    def refresh_window_list(self):
        """刷新窗口列表"""
        self.window_selector.refresh_windows()
        # 不再需要更新下拉框，因为已经删除了这个UI组件
    
    def toggle_serial_connection(self):
        """切换串口连接状态"""
        if self.serial_manager.is_connected():
            # 断开连接
            self.serial_manager.disconnect()
            self.connect_btn.config(text="连接")
            self.connection_var.set("未连接")
            self.status_var.set("串口已断开")
        else:
            # 连接串口
            port_display = self.port_var.get()
            if not port_display:
                messagebox.showwarning("警告", "请选择串口")
                return
            
            # 从显示名称中提取设备名
            port = port_display.split(' - ')[0]
            baudrate = int(self.baudrate_var.get())
            
            if self.serial_manager.connect(port, baudrate):
                self.connect_btn.config(text="断开")
                self.connection_var.set(f"已连接 {port}")
                self.status_var.set(f"串口 {port} 连接成功")
                
                # 保存设置
                self.config_manager.set('serial_port', port)
                self.config_manager.set('baudrate', baudrate)
                self.config_manager.save_config()
            else:
                messagebox.showerror("错误", "串口连接失败")
    
    def start_drag_select(self):
        """开始拖拽选择窗口"""
        if not WIN32_AVAILABLE:
            messagebox.showwarning("警告", "拖拽功能需要安装pywin32模块")
            return
        
        if not pyautogui:
            messagebox.showwarning("警告", "拖拽功能需要安装pyautogui模块")
            return
            
        # 检查是否正在进行拖拽操作
        if getattr(self, 'is_dragging', False):
            logger.info("拖拽操作已在进行中")
            return
            
        self.status_var.set("请按住鼠标左键选择窗口，松开鼠标时确认选择...")
        logger.info("开始窗口拖拽选择")
        
        # 隐藏主窗口
        self.root.withdraw()
        
        # 等待一小段时间让窗口隐藏
        self.root.after(200, self.start_mouse_drag_selection)  # 减少等待时间以提高响应速度
    
    def start_mouse_drag_selection(self):
        """开始鼠标拖拽选择"""
        try:
            # 标记拖拽状态
            self.is_dragging = True
            self.drag_cancelled = False
            
            import time
            
            # 显示一个可移动的按钮跟随鼠标
            overlay = tk.Toplevel(self.root)
            overlay.overrideredirect(True)  # 无边框窗口
            overlay.attributes('-topmost', True)  # 置顶
            overlay.attributes('-alpha', 0.9)  # 半透明
            
            # 创建可移动按钮
            drag_btn = tk.Button(overlay, text="🎯选择终端",
                               font=("SimHei", 10), bg="#FF6B6B", fg="white",
                               relief=tk.RAISED, padx=10, pady=5)
            drag_btn.pack()
            
            # 设置窗口初始位置为鼠标位置
            if pyautogui is None:
                logger.error("pyautogui模块未安装，无法使用鼠标位置")
                overlay.geometry(f"120x40+100+100")  # 使用默认位置
            else:
                current_x, current_y = pyautogui.position()
                window_width = 120
                window_height = 40
                overlay.geometry(f"{window_width}x{window_height}+{current_x}+{current_y}")
            
            # 简化的鼠标事件处理
            def on_mouse_click(event):
                """鼠标点击事件处理"""
                nonlocal overlay
                try:
                    # 获取当前鼠标位置
                    if pyautogui is None:
                        logger.error("pyautogui模块未安装，无法获取鼠标位置")
                        self._cancel_drag_selection()
                        return
                    click_x, click_y = pyautogui.position()
                    logger.info(f"鼠标点击位置: ({click_x}, {click_y})")
                    
                    # 关闭overlay窗口
                    overlay.destroy()
                    
                    # 执行窗口选择
                    self.root.after(0, lambda x=click_x, y=click_y: self.select_window_at_position(x, y))
                except Exception as e:
                    logger.error(f"处理鼠标点击事件失败: {e}")
                    self._cancel_drag_selection()
            
            def on_esc_press(event):
                """ESC键按下事件处理"""
                nonlocal overlay
                try:
                    # 关闭overlay窗口
                    overlay.destroy()
                    self._cancel_drag_selection()
                except Exception as e:
                    logger.error(f"处理ESC键事件失败: {e}")
            
            # 绑定点击事件
            overlay.bind("<Button-1>", on_mouse_click)
            
            # 绑定ESC键事件
            overlay.bind("<Escape>", on_esc_press)
            
            # 设置窗口移动和更新
            def update_overlay():
                """更新overlay窗口位置"""
                if not overlay.winfo_exists():
                    return
                
                try:
                    # 获取当前鼠标位置
                    if pyautogui is None:
                        logger.error("pyautogui模块未安装，无法更新overlay窗口位置")
                        if overlay.winfo_exists():
                            overlay.destroy()
                        self._cancel_drag_selection()
                        return
                    current_x, current_y = pyautogui.position()
                    # 更新overlay窗口位置
                    overlay.geometry(f"{window_width}x{window_height}+{current_x}+{current_y}")
                    # 继续更新
                    self.root.after(50, update_overlay)  # 每50ms更新一次位置
                except Exception as e:
                    logger.error(f"更新overlay窗口位置失败: {e}")
                    if overlay.winfo_exists():
                        overlay.destroy()
                    self._cancel_drag_selection()
            
            # 启动overlay窗口位置更新
            self.root.after(0, update_overlay)
            
            # 设置超时处理
            def timeout_handler():
                """超时处理"""
                if overlay.winfo_exists():
                    overlay.destroy()
                    self._cancel_drag_selection()
            
            # 30秒超时
            self.root.after(30000, timeout_handler)
            
        except Exception as e:
            logger.error(f"启动鼠标拖拽监控失败: {e}")
            self.status_var.set("启动拖拽选择失败")
            try:
                self.root.deiconify()
            except:
                pass
            self.is_dragging = False
            
    def _cancel_drag_selection(self):
        """取消拖拽选择"""
        # 重置所有相关状态
        self.is_dragging = False
        self.drag_cancelled = True
        self.status_var.set("已取消窗口选择")
        logger.info("用户取消了窗口选择操作")
        
        # 确保主线程窗口可见
        try:
            if self.root.winfo_exists():
                self.root.deiconify()
        except Exception as e:
            logger.error(f"显示主窗口时出错: {e}")
    
    def select_window_at_position(self, x, y):
        """选择指定位置的窗口，确保选中有效终端"""
        try:
            # 1. 获取鼠标位置的窗口句柄，增加异常处理
            hwnd = None
            try:
                hwnd = win32gui.WindowFromPoint((x, y))
                logger.info(f"获取到窗口句柄: {hwnd}")
            except Exception as e:
                logger.warning(f"获取窗口句柄失败: {e}")
                # 不依赖句柄，直接使用模拟点击
                self.status_var.set("已选择终端，将使用模拟键盘输入")
                # 创建一个简化的窗口对象，不包含hwnd
                new_window = {
                    'hwnd': None,
                    'title': "未知终端",
                    'process_id': 0,
                    'process_name': "未知进程",
                    'is_terminal': True,
                    'can_input_commands': True,
                    'display_name': "未知终端",
                    'selected_position': (x, y)  # 保存选择时的鼠标位置
                }
                self.window_selector.selected_window = new_window
                self.connection_var.set(f"已连接: 未知终端")
                self.root.deiconify()
                return
            
            # 验证窗口句柄有效性
            if not hwnd or not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                # 不依赖句柄，直接使用模拟点击
                self.status_var.set("已选择终端，将使用模拟键盘输入")
                new_window = {
                    'hwnd': None,
                    'title': "未知终端",
                    'process_id': 0,
                    'process_name': "未知进程",
                    'is_terminal': True,
                    'can_input_commands': True,
                    'display_name': "未知终端",
                    'selected_position': (x, y)  # 保存选择时的鼠标位置
                }
                self.window_selector.selected_window = new_window
                self.connection_var.set(f"已连接: 未知终端")
                self.root.deiconify()
                return
            
            # 获取顶级窗口，增加异常处理
            try:
                hwnd = win32gui.GetAncestor(hwnd, win32con.GA_ROOT)
                logger.info(f"获取到顶级窗口句柄: {hwnd}")
            except Exception as e:
                logger.warning(f"获取顶级窗口失败: {e}")
            
            # 再次验证顶级窗口句柄有效性
            if hwnd and (not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd)):
                logger.warning(f"顶级窗口句柄无效")
                hwnd = None
            
            # 2. 获取窗口信息，增加异常处理
            window_title = ""
            try:
                if hwnd:
                    window_title = win32gui.GetWindowText(hwnd)
                    logger.info(f"获取到窗口标题: {window_title}")
            except Exception as e:
                logger.warning(f"获取窗口标题失败: {e}")
            
            # 即使窗口标题为空，也尝试继续处理
            if not window_title:
                window_title = "未知终端"
            
            # 获取进程信息，增加异常处理
            process_id = 0
            process_name = "Unknown"
            
            try:
                if hwnd:
                    process_id = win32process.GetWindowThreadProcessId(hwnd)[1]
                    logger.info(f"获取到进程ID: {process_id}")
            except Exception as e:
                logger.warning(f"获取进程ID失败: {e}")
                process_id = 0
            
            try:
                if PSUTIL_AVAILABLE and process_id != 0:
                    process = psutil.Process(process_id)
                    process_name = process.name()
                    logger.info(f"获取到进程名: {process_name}")
            except Exception as e:
                logger.warning(f"获取进程信息失败: {e}")
            
            # 3. 检查是否为可输入命令的终端
            is_terminal = True  # 默认认为是终端，除非明确排除
            can_input_commands = True  # 默认认为可以输入命令
            
            # 特殊检测：如果进程名或窗口标题包含"命令发送器"，则不是终端
            is_command_sender = False
            if '命令发送器' in window_title or 'command_sender' in process_name.lower():
                is_command_sender = True
                is_terminal = False
                can_input_commands = False
                logger.info(f"检测到命令发送器窗口，跳过选择")
            
            # 4. 增强窗口激活逻辑，确保窗口正确获得焦点
            if hwnd:
                # 使用更可靠的窗口激活方法
                try:
                    # 先发送WM_SETFOCUS消息
                    win32gui.PostMessage(hwnd, win32con.WM_SETFOCUS, 0, 0)
                    time.sleep(0.05)
                    
                    # 然后发送WM_ACTIVATE消息
                    win32gui.PostMessage(hwnd, win32con.WM_ACTIVATE, win32con.WA_ACTIVE, 0)
                    time.sleep(0.05)
                    
                    # 最后使用SetForegroundWindow确保窗口获得焦点
                    win32gui.SetForegroundWindow(hwnd)
                    time.sleep(0.1)
                    logger.info(f"成功激活窗口: {window_title}")
                except Exception as e:
                    logger.warning(f"激活窗口失败: {e}")
                    # 备选方案：使用模拟鼠标点击
                    if pyautogui is None:
                        logger.warning("pyautogui模块未安装，无法使用模拟鼠标点击方式")
                    else:
                        original_x, original_y = pyautogui.position()
                        pyautogui.moveTo(x, y)
                        time.sleep(0.1)
                        pyautogui.click()
                        time.sleep(0.1)
                        pyautogui.moveTo(original_x, original_y)
                        time.sleep(0.1)
                        logger.info(f"使用模拟鼠标点击激活窗口: {window_title}")
            else:
                # 备选方案：使用模拟鼠标点击
                if pyautogui is None:
                    logger.warning("pyautogui模块未安装，无法使用模拟鼠标点击方式")
                else:
                    original_x, original_y = pyautogui.position()
                    pyautogui.moveTo(x, y)
                    time.sleep(0.1)
                    pyautogui.click()
                    time.sleep(0.1)
                    pyautogui.moveTo(original_x, original_y)
                    time.sleep(0.1)
                    logger.info("使用模拟鼠标点击激活窗口")
            
            # 5. 创建窗口对象，保存选择位置
            new_window = {
                'hwnd': hwnd,
                'title': window_title,
                'process_id': process_id,
                'process_name': process_name,
                'is_terminal': is_terminal,
                'can_input_commands': can_input_commands,
                'display_name': f"{window_title} ({process_name})",
                'selected_position': (x, y)  # 保存选择时的鼠标位置
            }
            
            # 6. 更新选中窗口
            self.window_selector.selected_window = new_window
            self.config_manager.set('target_window', window_title)
            self.config_manager.save_config()
            
            # 7. 更新状态
            status_text = "已选择终端: "
            status_text += "📟 "  # 终端图标
            status_text += window_title
            self.status_var.set(status_text)
            
            # 更新连接状态
            self.connection_var.set(f"已连接: {window_title}")
            
            logger.info(f"已成功选择终端: {window_title} (PID: {process_id}, 进程: {process_name})")
            
        except Exception as e:
            logger.error(f"选择窗口失败: {e}")
            self.status_var.set(f"错误: 选择窗口时发生错误 - {str(e)}")
        finally:
            # 显示主窗口
            self.root.deiconify()
            # 重置选择状态
            self.is_dragging = False
            self.drag_cancelled = False
    
    def toggle_always_on_top(self):
        """切换窗口置顶状态"""
        try:
            always_on_top = self.always_on_top_var.get()
            self.root.attributes('-topmost', always_on_top)
            try:
                self.config_manager.set('always_on_top', always_on_top)
                self.config_manager.save_config()
            except Exception as e:
                logger.error(f"保存置顶设置失败: {e}")
                # 配置保存失败不影响功能
        except Exception as e:
            logger.error(f"切换窗口置顶状态失败: {e}")
            messagebox.showerror("错误", f"切换窗口置顶状态失败: {str(e)}")
    
    def _on_start_drag(self, event):
        """开始拖动窗口"""
        # 只有点击标题栏区域时才允许拖动
        if event.widget == self.root and event.y < 30:
            self.drag_data["x"] = event.x
            self.drag_data["y"] = event.y
    
    def _on_drag(self, event):
        """拖动窗口"""
        if self.drag_data["x"] != 0 or self.drag_data["y"] != 0:
            x = self.root.winfo_pointerx() - self.drag_data["x"]
            y = self.root.winfo_pointery() - self.drag_data["y"]
            self.root.geometry(f"+{x}+{y}")
    
    def _on_stop_drag(self, event):
        """停止拖动窗口"""
        self.drag_data["x"] = 0
        self.drag_data["y"] = 0
    
    def _on_escape(self, event):
        """处理ESC键事件"""
        # 检查是否有打开的对话框或选择器正在运行
        # 取消窗口选择器
        try:
            if hasattr(self, 'drag_select_thread') and self.drag_select_thread and self.drag_select_thread.is_alive():
                self._cancel_drag_selection()
                return
        except Exception as e:
            logger.error(f"处理ESC键事件失败: {e}")
    
    def open_file(self, file_path=None):
        """打开文件对话框，加载文件内容到编辑器和命令列表"""
        try:
            if not file_path:
                file_path = filedialog.askopenfilename(
                    title="选择命令文件",
                    filetypes=[
                        ("命令文件", "*.txt"),
                        ("批处理文件", "*.bat"),
                        ("Shell脚本", "*.sh"),
                        ("Python文件", "*.py"),
                        ("所有文件", "*.*")
                    ]
                )
            
            if file_path:  # 用户选择了文件
                try:
                    with open(file_path, 'r', encoding='utf-8') as file:
                        content = file.read()
                        # 更新文本编辑器
                        self.text_editor.delete(1.0, tk.END)
                        self.text_editor.insert(1.0, content)
                        
                        # 更新文件信息
                        self.current_file = file_path
                        self.is_modified = False
                        self.update_title()
                        self.update_status(f"已打开文件: {os.path.basename(file_path)}")
                        self.config_manager.add_recent_file(file_path)
                        # 保存当前文件路径到配置，用于下次启动时自动打开
                        self.config_manager.set('last_file', file_path)
                        self.config_manager.save_config()
                        logger.info(f"成功打开文件: {file_path}")
                        
                        # 启动文件监控
                        self.start_file_monitor()
                except Exception as e:
                    messagebox.showerror("错误", f"无法打开文件: {str(e)}")
                    logger.error(f"打开文件失败: {str(e)}")
        except Exception as e:
            logger.error(f"打开文件对话框失败: {str(e)}")
    
    def save_file(self):
        """保存文件"""
        if self.current_file:
            self.save_file_as(self.current_file)
        else:
            self.save_file_as()
    
    def save_file_as(self, file_path=None):
        """另存为文件"""
        try:
            if not file_path:
                file_path = filedialog.asksaveasfilename(
                    title="保存文件",
                    defaultextension=".txt",
                    filetypes=[
                        ("文本文件", "*.txt"),
                        ("Python文件", "*.py"),
                        ("所有文件", "*.*")
                    ]
                )
            
            if file_path:  # 用户选择了保存位置
                content = self.text_editor.get(1.0, tk.END)
                with open(file_path, 'w', encoding='utf-8') as file:
                    file.write(content)
                self.current_file = file_path
                self.is_modified = False
                self.update_title()
                self.update_status(f"已保存文件: {os.path.basename(file_path)}")
                self.config_manager.add_recent_file(file_path)
                # 保存当前文件路径到配置，用于下次启动时自动打开
                self.config_manager.set('last_file', file_path)
                self.config_manager.save_config()
                logger.info(f"成功保存文件: {file_path}")
                
                # 更新文件修改时间，避免误报外部修改
                self.current_file_mtime = self.get_file_mtime(file_path)
                
                # 确保文件监控正在运行
                self.start_file_monitor()
        except Exception as e:
            messagebox.showerror("错误", f"无法保存文件: {str(e)}")
            logger.error(f"保存文件失败: {str(e)}")
    
    def on_window_resize(self, event=None):
        """窗口大小改变事件处理"""
        try:
            # 防抖动处理 - 如果已经有定时器在运行，取消它
            if hasattr(self, 'resize_timer') and self.resize_timer:
                self.root.after_cancel(self.resize_timer)
            
            # 设置新的定时器，延迟100ms执行
            self.resize_timer = self.root.after(100, self._handle_resize)
        except Exception as e:
            logger.error(f"处理窗口大小改变事件失败: {str(e)}")
    
    def _handle_resize(self):
        """实际处理窗口大小改变的函数"""
        try:
            # 更新行号
            self.update_line_numbers()
            logger.debug("窗口大小改变，已更新UI布局")
            
            # 清除定时器引用
            self.resize_timer = None
        except Exception as e:
            logger.error(f"处理窗口大小改变事件失败: {str(e)}")
    
    def get_file_mtime(self, file_path):
        """获取文件的修改时间"""
        try:
            return os.path.getmtime(file_path)
        except Exception:
            return None
    
    def start_file_monitor(self):
        """启动文件监控"""
        if self.current_file and not hasattr(self, '_file_monitor_running'):
            try:
                self._file_monitor_running = True
                self._file_mtime = self.get_file_mtime(self.current_file)
                self.root.after(1000, self.check_file_external_modification)
                logger.info(f"已启动文件监控: {self.current_file}")
            except Exception as e:
                logger.error(f"启动文件监控失败: {str(e)}")
                self._file_monitor_running = False
    
    def stop_file_monitor(self):
        """停止文件监控"""
        if hasattr(self, '_file_monitor_running'):
            self._file_monitor_running = False
            logger.info("已停止文件监控")
    
    def check_file_external_modification(self):
        """检查文件是否被外部修改"""
        if not hasattr(self, '_file_monitor_running') or not self._file_monitor_running or not self.current_file:
            return
        
        try:
            current_mtime = self.get_file_mtime(self.current_file)
            if current_mtime and hasattr(self, '_file_mtime') and current_mtime != self._file_mtime:
                # 文件被外部修改
                self.handle_external_modification()
            
            # 继续监控
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
        except Exception as e:
            logger.error(f"检查文件外部修改失败: {str(e)}")
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
    
    def handle_external_modification(self):
        """处理文件被外部修改的情况"""
        if not self.current_file:
            return
        
        # 如果当前文件未被修改，直接重新加载
        if not self.is_modified:
            self.reload_file()
            return
        
        # 如果当前文件已被修改，询问用户如何处理
        response = messagebox.askyesnocancel(
            "文件已被外部修改",
            "当前文件已被外部程序修改。\n" \
            "是否保存当前更改并重新加载？\n\n" \
            "是: 保存当前更改并重新加载外部修改\n" \
            "否: 放弃当前更改并重新加载外部修改\n" \
            "取消: 保持当前状态不变"
        )
        
        if response is None:  # 用户点击了取消
            return
        elif response:  # 用户点击了是
            # 保存当前更改
            self.save_file()
            # 重新加载文件
            self.reload_file()
        else:  # 用户点击了否
            # 直接重新加载文件，放弃当前更改
            self.reload_file()
    
    def reload_file(self):
        """重新加载当前文件"""
        if not self.current_file:
            return
        
        try:
            with open(self.current_file, 'r', encoding='utf-8') as file:
                content = file.read()
            
            # 更新文本编辑器内容
            self.text_editor.delete(1.0, tk.END)
            self.text_editor.insert(1.0, content)
            
            # 更新文件修改时间和状态
            self._file_mtime = self.get_file_mtime(self.current_file)
            self.is_modified = False
            self.update_title()
            self.update_status(f"已重新加载文件: {os.path.basename(self.current_file)}")
            logger.info(f"已重新加载文件: {self.current_file}")
        except Exception as e:
            messagebox.showerror("错误", f"无法重新加载文件: {str(e)}")
            logger.error(f"重新加载文件失败: {str(e)}")

    def new_file(self):
        """新建文件"""
        # 如果当前内容已修改，提示保存
        if self.is_modified:
            result = messagebox.askyesnocancel("新建文件", "当前文件已修改，是否保存？")
            if result is True:  # 用户选择保存
                self.save_file()
                if not self.current_file:  # 用户取消了保存
                    return
            elif result is None:  # 用户取消
                return
        
        # 清空内容并重置状态
        self.text_editor.delete(1.0, tk.END)
        self.current_file = None
        self.is_modified = False
        self.update_title()
        self.update_status("新建文件")
        # 清除last_file配置，因为用户创建了新文件
        self.config_manager.set('last_file', '')
        self.config_manager.save_config()
        logger.info("已新建文件")
    
    def check_modified(self):
        """检查文件是否已修改"""
        if not hasattr(self, 'current_file'):
            return False
        
        if not self.current_file:
            return False
        
        try:
            current_content = self.text_editor.get(1.0, tk.END)
            with open(self.current_file, 'r', encoding='utf-8') as f:
                file_content = f.read()
            
            return current_content != file_content
        except Exception:
            return False
    
    def update_recent_files(self, file_path):
        """更新最近打开的文件列表"""
        try:
            if file_path in self.recent_files:
                self.recent_files.remove(file_path)
            self.recent_files.insert(0, file_path)
            
            # 限制最近文件数量
            if len(self.recent_files) > 10:
                self.recent_files = self.recent_files[:10]
            
            # 更新菜单
            self.update_recent_files_menu()
            
            # 保存到配置
            self.save_config()
        except Exception as e:
            logger.error(f"更新最近文件列表失败: {str(e)}")
    
    def update_recent_files_menu(self):
        """更新最近文件菜单"""
        try:
            # 清除现有菜单项
            self.recent_files_menu.delete(0, tk.END)
            
            # 获取最近文件列表
            recent_files = self.config_manager.get('recent_files', [])
            
            # 添加最近文件项
            for file_path in recent_files:
                file_name = os.path.basename(file_path)
                self.recent_files_menu.add_command(
                    label=file_name,
                    command=lambda p=file_path: self.open_recent_file(p)
                )
        except Exception as e:
            logger.error(f"更新最近文件菜单失败: {str(e)}")
    
    def open_recent_file(self, file_path):
        """打开最近使用的文件"""
        try:
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                    self.text_editor.delete(1.0, tk.END)
                    self.text_editor.insert(1.0, content)
                    self.current_file = file_path
                    self.is_modified = False
                    self.update_title()
                    self.update_status(f"已打开文件: {os.path.basename(file_path)}")
                    
                    # 添加到最近文件列表（会自动移动到顶部）
                    self.config_manager.add_recent_file(file_path)
                    
                    logger.info(f"成功打开最近文件: {file_path}")
            else:
                messagebox.showerror("错误", f"文件不存在: {file_path}")
                # 从配置中移除不存在的文件
                recent_files = self.config_manager.get('recent_files', [])
                if file_path in recent_files:
                    recent_files.remove(file_path)
                    self.config_manager.set('recent_files', recent_files)
                    self.config_manager.save_config()
                    self.update_recent_files_menu()
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件: {str(e)}")
            logger.error(f"打开最近文件失败: {str(e)}")
    
    def save_config(self):
        """保存配置到文件"""
        try:
            config = {
                'recent_files': self.recent_files,
                'window_geometry': self.root.geometry() if hasattr(self, 'root') else "",
                'font_family': self.font_family if hasattr(self, 'font_family') else "Consolas",
                'font_size': self.font_size if hasattr(self, 'font_size') else 10,
                'theme': self.theme if hasattr(self, 'theme') else "Light"
            }
            
            config_path = os.path.join(os.path.dirname(__file__), 'config.json')
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                
            logger.debug("配置已保存")
        except Exception as e:
            logger.error(f"保存配置失败: {str(e)}")
    
    def load_config(self):
        """从文件加载配置"""
        try:
            config_path = os.path.join(os.path.dirname(__file__), 'config.json')
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                self.recent_files = config.get('recent_files', [])
                self.font_family = config.get('font_family', 'Consolas')
                self.font_size = config.get('font_size', 10)
                self.theme = config.get('theme', 'Light')
                
                # 应用窗口几何配置
                geometry = config.get('window_geometry', '')
                if geometry and hasattr(self, 'root'):
                    try:
                        self.root.geometry(geometry)
                    except:
                        pass
                
                logger.debug("配置已加载")
                return True
        except Exception as e:
            logger.error(f"加载配置失败: {str(e)}")
        
        # 设置默认值
        self.recent_files = []
        self.font_family = "Consolas"
        self.font_size = 10
        self.theme = "Light"
        return False

    def on_closing(self):
        """窗口关闭事件"""
        if self.is_modified:
            result = messagebox.askyesnocancel("保存", "当前文件已修改，是否保存？")
            if result is True:
                if not self.save_file():
                    return  # 保存失败，不关闭
            elif result is None:  # 取消
                return
        
        # 保存窗口几何信息和当前文件路径
        self.config_manager.set('window_geometry', self.root.geometry())
        if self.current_file:
            self.config_manager.set('last_file', self.current_file)
        else:
            self.config_manager.set('last_file', '')
        self.config_manager.save_config()
        
        # 断开串口连接
        if self.serial_manager.is_connected():
            self.serial_manager.disconnect()
        
        # 关闭窗口
        self.root.destroy()
    
    def get_file_mtime(self, file_path):
        """获取文件的修改时间"""
        try:
            return os.path.getmtime(file_path)
        except Exception:
            return None
    
    def start_file_monitor(self):
        """启动文件监控"""
        if self.current_file and not hasattr(self, '_file_monitor_running'):
            try:
                self._file_monitor_running = True
                self._file_mtime = self.get_file_mtime(self.current_file)
                self.root.after(1000, self.check_file_external_modification)
                logger.info(f"已启动文件监控: {self.current_file}")
            except Exception as e:
                logger.error(f"启动文件监控失败: {str(e)}")
                self._file_monitor_running = False
    
    def stop_file_monitor(self):
        """停止文件监控"""
        if hasattr(self, '_file_monitor_running'):
            self._file_monitor_running = False
            logger.info("已停止文件监控")
    
    def check_file_external_modification(self):
        """检查文件是否被外部修改"""
        if not hasattr(self, '_file_monitor_running') or not self._file_monitor_running or not self.current_file:
            return
        
        try:
            current_mtime = self.get_file_mtime(self.current_file)
            if current_mtime and hasattr(self, '_file_mtime') and current_mtime != self._file_mtime:
                # 文件被外部修改
                self.handle_external_modification()
            
            # 继续监控
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
        except Exception as e:
            logger.error(f"检查文件外部修改失败: {str(e)}")
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
    
    def handle_external_modification(self):
        """处理文件被外部修改的情况"""
        if not self.current_file:
            return
        
        # 如果当前文件未被修改，直接重新加载
        if not self.is_modified:
            self.reload_file()
            return
        
        # 如果当前文件已被修改，询问用户如何处理
        response = messagebox.askyesnocancel(
            "文件已被外部修改",
            "当前文件已被外部程序修改。\n" \
            "是否保存当前更改并重新加载？\n\n" \
            "是: 保存当前更改并重新加载外部修改\n" \
            "否: 放弃当前更改并重新加载外部修改\n" \
            "取消: 保持当前状态不变"
        )
        
        if response is None:  # 用户点击了取消
            return
        elif response:  # 用户点击了是
            # 保存当前更改
            self.save_file()
            # 重新加载文件
            self.reload_file()
        else:  # 用户点击了否
            # 直接重新加载文件，放弃当前更改
            self.reload_file()
    
    def reload_file(self):
        """重新加载当前文件"""
        if not self.current_file:
            return
        
        try:
            with open(self.current_file, 'r', encoding='utf-8') as file:
                content = file.read()
            
            # 更新文本编辑器内容
            self.text_editor.delete(1.0, tk.END)
            self.text_editor.insert(1.0, content)
            
            # 更新文件修改时间和状态
            self._file_mtime = self.get_file_mtime(self.current_file)
            self.is_modified = False
            self.update_title()
            self.update_status(f"已重新加载文件: {os.path.basename(self.current_file)}")
            logger.info(f"已重新加载文件: {self.current_file}")
        except Exception as e:
            messagebox.showerror("错误", f"无法重新加载文件: {str(e)}")
            logger.error(f"重新加载文件失败: {str(e)}")
    
    def update_title(self):
        """更新窗口标题"""
        if self.current_file:
            title = f"命令发送器 - {os.path.basename(self.current_file)}"
            if self.is_modified:
                title += " *"
        else:
            title = "命令发送器 - 未命名"
        self.root.title(title)
    
    def update_status(self, message):
        """更新状态栏"""
        self.status_var.set(message)
        
        # 根据消息内容设置不同的状态栏样式
        if "错误" in message and "测试" not in message:
            # 只在真正的错误情况下记录错误日志，排除测试消息
            logger.error(message)
        elif "成功" in message:
            # 成功状态
            logger.info(message)
        elif "警告" in message:
            # 警告状态
            logger.warning(message)
        else:
            # 普通状态
            logger.info(message)
            
        # 更新文件信息
        if self.current_file:
            self.file_info_var.set(f"{os.path.basename(self.current_file)}")
        else:
            self.file_info_var.set("未打开文件")
        
        # 记录命令历史
        if hasattr(self, 'command_history'):
            self.command_history.append(message)
        else:
            self.command_history = [message]
    
    def on_text_change(self, event=None):
        """文本内容改变事件"""
        if not self.is_modified:
            self.is_modified = True
            self.update_title()
        
        # 更新行号
        self.update_line_numbers()
        
        # 高亮当前行
        self.highlight_current_line()
        
        # 安排自动保存
        self.schedule_auto_save()
    
    def on_text_click(self, event=None):
        """文本点击事件"""
        self.highlight_current_line()
    
    def on_mouse_wheel(self, event):
        """鼠标滚轮事件，处理滚动逻辑"""
        # 计算滚动方向和距离
        if hasattr(event, 'delta'):
            # Windows和Mac系统
            delta = event.delta
        else:
            # Linux系统
            delta = -event.num
        
        # 调整滚动量（除以120使滚动速度合适）
        scroll_amount = delta / 120
        
        # 滚动文本编辑器
        self.text_editor.yview_scroll(int(-scroll_amount), "units")
        
        # 同步滚动行号
        self.line_numbers.yview_scroll(int(-scroll_amount), "units")
        
        # 更新行号显示
        self.update_line_numbers()
    
    def sync_scroll(self, event=None):
        """同步滚动条、行号和发送按钮"""
        try:
            # 确保行号和文本区域同步
            self.line_numbers.yview_moveto(self.text_editor.yview()[0])
            
            # 同步发送按钮Canvas滚动，确保按钮跟随文本滚动
            self.send_buttons_canvas.yview_moveto(self.text_editor.yview()[0])
            
            # 更新行号显示
            self.update_line_numbers()
            
            # 如果有可见的发送按钮，更新其位置，确保与对应行对齐
            if self.send_buttons and self.current_visible_line:
                self.show_send_button(self.current_visible_line)
        except Exception as e:
            logger.error(f"同步滚动失败: {e}")
    
    def update_line_numbers(self):
        """更新行号显示"""
        try:
            # 获取文本行数
            line_count = int(self.text_editor.index('end-1c').split('.')[0])
            
            # 生成行号文本
            line_numbers_text = '\n'.join(str(i) for i in range(1, line_count + 1))
            
            # 更新行号显示
            self.line_numbers.config(state='normal')
            self.line_numbers.delete(1.0, tk.END)
            self.line_numbers.insert(1.0, line_numbers_text)
            self.line_numbers.config(state='disabled')
            
            # 只在没有可见按钮时才更新发送按钮
            # 避免在更新行号时意外隐藏当前可见按钮
            if not self.send_buttons:
                self.update_send_buttons()
        except Exception as e:
            logger.error(f"更新行号失败: {e}")
    
    def update_send_buttons(self):
        """更新发送按钮，确保每行都有对应的按钮"""
        try:
            # 获取文本行数
            line_count = int(self.text_editor.index('end-1c').split('.')[0])
            
            # 清除所有现有按钮
            for line_num in list(self.send_buttons.keys()):
                btn = self.send_buttons[line_num]
                btn.destroy()
                del self.send_buttons[line_num]
            
            # 清空Canvas
            self.send_buttons_canvas.delete("all")
            
            # 为当前可见行创建按钮
            # 只在当前行显示按钮，不需要为所有行创建
        except Exception as e:
            logger.error(f"更新发送按钮失败: {e}")
    
    def on_text_click(self, event):
        """文本点击事件，显示对应行的发送按钮"""
        try:
            # 获取点击位置所在行
            click_pos = self.text_editor.index(f"@{event.x},{event.y}")
            current_line = int(click_pos.split('.')[0])
            
            # 显示该行的发送按钮
            self.show_send_button(current_line)
        except Exception as e:
            logger.error(f"处理文本点击失败: {e}")
    
    def show_send_button(self, line_num):
        """显示指定行的发送按钮，隐藏其他行的按钮"""
        try:
            # 只在需要时创建新按钮，避免不必要的销毁和重建
            if self.current_visible_line == line_num and line_num in self.send_buttons:
                # 如果当前行已经有可见按钮，直接返回
                return
            
            # 清除所有现有按钮
            for btn_num in list(self.send_buttons.keys()):
                btn = self.send_buttons[btn_num]
                btn.destroy()
                del self.send_buttons[btn_num]
            
            # 清空Canvas
            self.send_buttons_canvas.delete("all")
            
            # 增强：使用更精确的按钮位置计算方法，直接获取指定行的实际位置
            line_index = f"{line_num}.0"
            line_bbox = self.text_editor.bbox(line_index)
            
            if line_bbox:
                # 行可见，直接使用行的实际Y坐标
                line_x, line_y, line_width, line_height = line_bbox
                x = 2
                # 按钮Y坐标 = 行Y坐标 + 垂直居中调整
                btn_y = line_y + (line_height - 24) / 2  # 24是按钮高度，垂直居中
            else:
                # 行不可见，使用备用方法
                # 获取当前可见区域的起始行
                visible_info = self.text_editor.yview()
                start_fraction = visible_info[0]
                total_lines = int(self.text_editor.index('end-1c').split('.')[0])
                start_line = int(start_fraction * total_lines)
                
                # 使用字体信息获取行高
                try:
                    font = self.text_editor.cget("font")
                    if isinstance(font, str):
                        import tkinter.font as tkfont
                        font_obj = tkfont.Font(family=font, size=10)
                        line_height = font_obj.metrics("linespace")
                    else:
                        line_height = font.metrics("linespace")
                except:
                    line_height = 24  # 默认行高
                
                # 计算按钮位置
                x = 2
                btn_y = (line_num - start_line - 1) * line_height + 2
            
            # 创建按钮，恢复原始样式
            btn = ttk.Button(
                self.send_buttons_canvas,
                text="▶",  # 使用原始三角符号
                width=2,  # 恢复原始宽度
                command=lambda line_num=line_num: self.send_line_command(line_num),
                style='SendButton.TButton'
            )
            
            # 将按钮添加到Canvas，恢复原始尺寸
            btn_window = self.send_buttons_canvas.create_window(x, btn_y, anchor=tk.NW, window=btn, width=24, height=24)
            
            # 保存按钮信息
            self.send_buttons[line_num] = btn
            self.current_visible_line = line_num
        except Exception as e:
            logger.error(f"显示发送按钮失败: {e}")
    
    def send_line_command(self, line_num):
        """发送指定行的命令"""
        try:
            # 获取指定行的内容
            line_start = f"{line_num}.0"
            line_end = f"{line_num}.end"
            line_content = self.text_editor.get(line_start, line_end).strip()
            
            # 发送命令
            if line_content:
                self.update_status(f"正在发送命令: {line_content[:50]}...")
                self.execute_command(line_content)
            else:
                self.update_status("错误: 命令内容为空")
        except Exception as e:
            logger.error(f"发送行命令失败: {e}")
            self.update_status(f"错误: 发送行命令失败 - {str(e)}")
    
    def update_send_buttons_positions(self, event=None):
        """更新发送按钮位置，确保按钮跟随行移动"""
        try:
            # 只有当有可见按钮时才更新位置
            if not self.send_buttons:
                return
            
            # 获取当前可见行
            line_num = list(self.send_buttons.keys())[0]
            
            # 重新显示按钮，确保位置正确
            self.show_send_button(line_num)
        except Exception as e:
            logger.error(f"更新发送按钮位置失败: {e}")
    
    def highlight_current_line(self):
        """高亮当前行"""
        try:
            # 移除之前的高亮
            self.text_editor.tag_remove("current_line", 1.0, tk.END)
            
            # 获取当前光标位置
            cursor_pos = self.text_editor.index(tk.INSERT)
            line_start = cursor_pos.split('.')[0] + ".0"
            line_end = cursor_pos.split('.')[0] + ".end"
            
            # 添加高亮
            self.text_editor.tag_add("current_line", line_start, line_end)
        except Exception as e:
            logger.error(f"高亮当前行失败: {e}")
    
    def schedule_auto_save(self):
        """安排自动保存"""
        if self.auto_save_timer:
            self.root.after_cancel(self.auto_save_timer)
        
        # 5秒后自动保存
        self.auto_save_timer = self.root.after(5000, self.auto_save)
    
    def auto_save(self):
        """自动保存"""
        try:
            if self.is_modified and self.current_file:
                content = self.text_editor.get(1.0, tk.END)
                with open(self.current_file, 'w', encoding='utf-8') as file:
                    file.write(content)
                self.is_modified = False
                self.update_title()
                logger.info(f"自动保存文件: {self.current_file}")
        except Exception as e:
            logger.error(f"自动保存失败: {e}")
    
    def undo(self):
        """撤销操作"""
        try:
            self.text_editor.edit_undo()
        except Exception as e:
            logger.error(f"撤销操作失败: {e}")
    
    def redo(self):
        """重做操作"""
        try:
            self.text_editor.edit_redo()
        except Exception as e:
            logger.error(f"重做操作失败: {e}")
    
    def cut(self):
        """剪切操作"""
        try:
            self.text_editor.event_generate("<<Cut>>")
        except Exception as e:
            logger.error(f"剪切操作失败: {e}")
    
    def copy(self):
        """复制操作"""
        try:
            self.text_editor.event_generate("<<Copy>>")
        except Exception as e:
            logger.error(f"复制操作失败: {e}")
    
    def paste(self):
        """粘贴操作"""
        try:
            self.text_editor.event_generate("<<Paste>>")
        except Exception as e:
            logger.error(f"粘贴操作失败: {e}")
    
    def find(self):
        """查找对话框"""
        try:
            # 创建查找对话框
            dialog = tk.Toplevel(self.root)
            dialog.title("查找")
            dialog.geometry("300x100")
            dialog.resizable(False, False)
            
            # 查找输入框
            ttk.Label(dialog, text="查找内容:").pack(pady=5)
            find_var = tk.StringVar()
            find_entry = ttk.Entry(dialog, textvariable=find_var, width=30)
            find_entry.pack(pady=5)
            find_entry.focus()
            
            # 按钮框架
            button_frame = ttk.Frame(dialog)
            button_frame.pack(pady=10)
            
            def do_find():
                """执行查找"""
                find_text = find_var.get()
                if not find_text:
                    return
                
                # 从当前光标位置开始查找
                pos = self.text_editor.search(find_text, tk.INSERT, tk.END)
                if not pos:
                    # 如果没找到，从头开始查找
                    pos = self.text_editor.search(find_text, 1.0, tk.END)
                
                if pos:
                    # 选中找到的文本
                    end_pos = f"{pos}+{len(find_text)}c"
                    self.text_editor.tag_remove(tk.SEL, 1.0, tk.END)
                    self.text_editor.tag_add(tk.SEL, pos, end_pos)
                    self.text_editor.mark_set(tk.INSERT, end_pos)
                    self.text_editor.see(pos)
                else:
                    messagebox.showinfo("查找", "未找到指定内容")
            
            # 按钮
            ttk.Button(button_frame, text="查找下一个", command=do_find).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
            
            # 绑定回车键
            find_entry.bind("<Return>", lambda e: do_find())
        except Exception as e:
            logger.error(f"显示查找对话框失败: {e}")
    
    def replace(self):
        """替换对话框"""
        try:
            # 创建替换对话框
            dialog = tk.Toplevel(self.root)
            dialog.title("替换")
            dialog.geometry("300x150")
            dialog.resizable(False, False)
            
            # 查找输入框
            ttk.Label(dialog, text="查找内容:").pack(pady=5)
            find_var = tk.StringVar()
            find_entry = ttk.Entry(dialog, textvariable=find_var, width=30)
            find_entry.pack(pady=5)
            find_entry.focus()
            
            # 替换输入框
            ttk.Label(dialog, text="替换为:").pack(pady=5)
            replace_var = tk.StringVar()
            replace_entry = ttk.Entry(dialog, textvariable=replace_var, width=30)
            replace_entry.pack(pady=5)
            
            # 按钮框架
            button_frame = ttk.Frame(dialog)
            button_frame.pack(pady=10)
            
            def do_find():
                """执行查找"""
                find_text = find_var.get()
                if not find_text:
                    return
                
                # 从当前光标位置开始查找
                pos = self.text_editor.search(find_text, tk.INSERT, tk.END)
                if not pos:
                    # 如果没找到，从头开始查找
                    pos = self.text_editor.search(find_text, 1.0, tk.END)
                
                if pos:
                    # 选中找到的文本
                    end_pos = f"{pos}+{len(find_text)}c"
                    self.text_editor.tag_remove(tk.SEL, 1.0, tk.END)
                    self.text_editor.tag_add(tk.SEL, pos, end_pos)
                    self.text_editor.mark_set(tk.INSERT, end_pos)
                    self.text_editor.see(pos)
                else:
                    messagebox.showinfo("查找", "未找到指定内容")
            
            def do_replace():
                """执行替换"""
                find_text = find_var.get()
                replace_text = replace_var.get()
                
                if not find_text:
                    return
                
                # 检查是否有选中的文本
                try:
                    selected = self.text_editor.get(tk.SEL_FIRST, tk.SEL_LAST)
                    if selected == find_text:
                        # 替换选中的文本
                        self.text_editor.delete(tk.SEL_FIRST, tk.SEL_LAST)
                        self.text_editor.insert(tk.INSERT, replace_text)
                        
                        # 查找下一个
                        do_find()
                    else:
                        # 查找第一个匹配项
                        do_find()
                except tk.TclError:
                    # 没有选中文本，查找第一个匹配项
                    do_find()
            
            def do_replace_all():
                """替换全部"""
                find_text = find_var.get()
                replace_text = replace_var.get()
                
                if not find_text:
                    return
                
                # 计数器
                count = 0
                
                # 从头开始查找并替换
                pos = self.text_editor.search(find_text, 1.0, tk.END)
                while pos:
                    end_pos = f"{pos}+{len(find_text)}c"
                    self.text_editor.delete(pos, end_pos)
                    self.text_editor.insert(pos, replace_text)
                    count += 1
                    
                    # 继续查找下一个
                    pos = self.text_editor.search(find_text, end_pos, tk.END)
                
                if count > 0:
                    messagebox.showinfo("替换全部", f"已替换 {count} 处")
                else:
                    messagebox.showinfo("替换全部", "未找到匹配内容")
                
                dialog.destroy()
            
            # 按钮
            ttk.Button(button_frame, text="查找", command=do_find).pack(side=tk.LEFT, padx=2)
            ttk.Button(button_frame, text="替换", command=do_replace).pack(side=tk.LEFT, padx=2)
            ttk.Button(button_frame, text="全部替换", command=do_replace_all).pack(side=tk.LEFT, padx=2)
            ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=2)
            
            # 绑定回车键
            find_entry.bind("<Return>", lambda e: do_find())
        except Exception as e:
            logger.error(f"显示替换对话框失败: {e}")
    
    def show_settings(self):
        """显示设置对话框"""
        try:
            # 创建设置对话框
            dialog = tk.Toplevel(self.root)
            dialog.title("设置")
            dialog.geometry("400x300")
            dialog.resizable(False, False)
            
            # 创建笔记本控件
            notebook = ttk.Notebook(dialog)
            notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # 常规设置选项卡
            general_frame = ttk.Frame(notebook)
            notebook.add(general_frame, text="常规")
            
            # 自动保存设置
            auto_save_var = tk.BooleanVar(value=self.config_manager.get('auto_save', True))
            ttk.Checkbutton(general_frame, text="自动保存", variable=auto_save_var).pack(anchor=tk.W, padx=10, pady=5)
            
            # 字体设置选项卡
            font_frame = ttk.Frame(notebook)
            notebook.add(font_frame, text="字体")
            
            # 字体族
            ttk.Label(font_frame, text="字体:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
            font_family_var = tk.StringVar(value=self.config_manager.get('font_family', 'Consolas'))
            font_family_combo = ttk.Combobox(font_frame, textvariable=font_family_var, width=20)
            font_family_combo['values'] = ['Consolas', 'Courier New', 'Monaco', 'Menlo', 'Lucida Console']
            font_family_combo.grid(row=0, column=1, padx=10, pady=5)
            
            # 字体大小
            ttk.Label(font_frame, text="大小:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
            font_size_var = tk.IntVar(value=self.config_manager.get('font_size', 10))
            font_size_spin = ttk.Spinbox(font_frame, from_=8, to=24, textvariable=font_size_var, width=5)
            font_size_spin.grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)
            
            # 主题设置选项卡
            theme_frame = ttk.Frame(notebook)
            notebook.add(theme_frame, text="主题")
            
            # 主题选择
            ttk.Label(theme_frame, text="主题:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
            theme_var = tk.StringVar(value=self.config_manager.get('theme', 'Light'))
            theme_combo = ttk.Combobox(theme_frame, textvariable=theme_var, width=20)
            theme_combo['values'] = ['Light', 'Dark']
            theme_combo.grid(row=0, column=1, padx=10, pady=5)
            
            # 按钮框架
            button_frame = ttk.Frame(dialog)
            button_frame.pack(pady=10)
            
            def apply_settings():
                """应用设置"""
                # 保存设置
                self.config_manager.set('auto_save', auto_save_var.get())
                self.config_manager.set('font_family', font_family_var.get())
                self.config_manager.set('font_size', font_size_var.get())
                self.config_manager.set('theme', theme_var.get())
                self.config_manager.set('simulate_keyboard', simulate_keyboard_var.get())
                self.config_manager.set('keyboard_delay', keyboard_delay_var.get())
                self.config_manager.save_config()
                
                # 应用字体设置
                self.apply_font_settings(font_family_var.get(), font_size_var.get())
                
                # 应用主题
                self.apply_theme(theme_var.get())
                
                # 移除messagebox，避免SetForegroundWindow错误
                self.update_status("设置已保存并应用")
                dialog.destroy()
            
            # 按钮
            ttk.Button(button_frame, text="应用", command=apply_settings).pack(side=tk.LEFT, padx=5)
            ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        except Exception as e:
            logger.error(f"显示设置对话框失败: {e}")
    
    def show_about(self):
        """显示关于对话框"""
        try:
            about_text = """命令发送器 v1.0

一个用于向终端窗口发送命令的工具。

功能特点：
• 支持多种发送方式（剪贴板、键盘输入、串口）
• 支持窗口选择
• 支持文本编辑
• 支持最近文件
• 支持自动保存
• 支持SecureCRT风格的命令按钮
• 支持MobaXterm风格的宏记录和回放

作者：AI Assistant
日期：2023"""
            
            messagebox.showinfo("关于", about_text)
        except Exception as e:
            logger.error(f"显示关于对话框失败: {e}")
    
    def select_target_window(self):
        """选择目标窗口"""
        try:
            # 获取当前鼠标位置下的窗口
            cursor_pos = win32gui.GetCursorPos()
            window_handle = win32gui.WindowFromPoint(cursor_pos)
            
            # 获取窗口标题
            window_title = win32gui.GetWindowText(window_handle)
            
            # 验证是否为目标窗口类型
            if "SecureCRT" in window_title or "MobaXterm" in window_title:
                self.target_window = window_handle
                self.target_window_title = window_title
                self.status_var.set(f"已选择目标: {window_title}")
                return True
            else:
                messagebox.showwarning("警告", "请选择SecureCRT或MobaXterm窗口")
                return False
        except Exception as e:
            messagebox.showerror("错误", f"选择窗口失败: {e}")
            logger.error(f"选择窗口失败: {e}")
            return False
    
    def execute_macro(self):
        """执行宏命令"""
        try:
            # 检查是否有选中的宏
            if not hasattr(self, 'selected_macro') or not self.selected_macro:
                self.show_warning("请先选择要执行的宏")
                return False
            
            # 执行宏命令
            for command in self.selected_macro['commands']:
                # 发送命令
                self.execute_command(command)
                time.sleep(0.5)  # 命令间延迟
            
            self.show_info(f"宏命令执行完成: {self.selected_macro['name']}")
            logger.info(f"宏命令执行完成: {self.selected_macro['name']}")
            return True
        except Exception as e:
            self.show_error(f"执行宏命令失败: {e}")
            logger.error(f"执行宏命令失败: {e}")
            return False
    
    def show_error(self, message):
        """显示错误消息"""
        messagebox.showerror("错误", message)
        logger.error(message)
    
    def show_warning(self, message):
        """显示警告消息"""
        messagebox.showwarning("警告", message)
        logger.warning(message)
    
    def show_info(self, message):
        """显示信息消息"""
        messagebox.showinfo("信息", message)
        logger.info(message)
    
    def connect_serial(self):
        """连接串口设备"""
        try:
            # 获取选中的端口和波特率
            port = self.port_var.get()
            baudrate = self.baudrate_var.get()
            
            if not port:
                self.show_warning("请先选择串口端口")
                return False
            
            # 转换波特率为整数
            try:
                baudrate = int(baudrate)
            except ValueError:
                self.show_error(f"无效的波特率: {baudrate}")
                return False
            
            # 连接串口
            success = self.serial_manager.connect(port, baudrate)
            if success:
                self.show_info(f"串口连接成功: {port}，波特率: {baudrate}")
                # 更新配置
                self.config_manager.set('serial_port', port)
                self.config_manager.set('baudrate', baudrate)
                self.config_manager.save_config()
            else:
                self.show_error(f"串口连接失败: {port}")
            
            return success
        except Exception as e:
            self.show_error(f"连接串口失败: {e}")
            logger.error(f"连接串口失败: {e}")
            return False
    
    def disconnect_serial(self):
        """断开串口连接"""
        try:
            self.serial_manager.disconnect()
            self.show_info("串口已断开")
            return True
        except Exception as e:
            self.show_error(f"断开串口失败: {e}")
            logger.error(f"断开串口失败: {e}")
            return False
    
    def send_serial_data(self, data):
        """通过串口发送数据"""
        try:
            if not self.serial_manager.is_connected():
                self.show_warning("串口未连接，无法发送数据")
                return False
            
            success = self.serial_manager.send_command(data)
            if success:
                logger.info(f"串口数据发送成功: {data}")
            else:
                self.show_error("串口数据发送失败")
            
            return success
        except Exception as e:
            self.show_error(f"发送串口数据失败: {e}")
            logger.error(f"发送串口数据失败: {e}")
            return False
    
    def show_tooltip(self, event, text):
        """显示命令提示"""
        # 先隐藏旧的tooltip
        self.hide_tooltip()
        
        # 创建临时提示窗口
        self.current_tooltip = tk.Toplevel(self.root)
        self.current_tooltip.overrideredirect(True)  # 无边框
        self.current_tooltip.attributes('-topmost', True)  # 置顶
        self.current_tooltip.attributes('-alpha', 0.9)  # 半透明
        
        # 创建提示标签
        label = ttk.Label(self.current_tooltip, text=text, background="#ffffe0", borderwidth=1, relief="solid")
        label.pack(padx=5, pady=5)
        
        # 计算位置
        x = event.x_root + 10
        y = event.y_root + 10
        
        # 设置位置
        self.current_tooltip.geometry(f"+{x}+{y}")
        
        # 3秒后自动关闭
        self.root.after(3000, self.hide_tooltip)
    
    def hide_tooltip(self, event=None):
        """隐藏命令提示"""
        if hasattr(self, 'current_tooltip') and self.current_tooltip:
            try:
                self.current_tooltip.destroy()
            except:
                pass
            finally:
                self.current_tooltip = None
    

    
    def start_macro_recording(self):
        """开始记录宏"""
        if not keyboard:
            messagebox.showwarning("警告", "缺少keyboard库，无法记录宏")
            return
        
        self.is_recording = True
        self.recorded_macro = []
        self.recording_start_time = time.time()
        
        # 更新按钮状态
        self.record_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.play_btn.config(state=tk.DISABLED)
        self.save_macro_btn.config(state=tk.DISABLED)
        
        # 更新状态信息
        self.recording_status_var.set("正在记录宏...")
        self.status_var.set("正在记录宏...")
        
        logger.info("开始记录宏")
        
        # 开始记录键盘事件
        def record_key_event(e):
            if self.is_recording:
                # 记录按键事件和时间
                elapsed_time = time.time() - self.recording_start_time
                self.recorded_macro.append({
                    "type": "key",
                    "event": e.keysym if hasattr(e, 'keysym') else str(e),
                    "char": e.char if hasattr(e, 'char') else '',
                    "time": elapsed_time
                })
                logger.debug(f"记录按键: {e.keysym}, 时间: {elapsed_time:.3f}s")
        
        # 绑定键盘事件
        self.key_record_bind_id = self.root.bind("<Key>", record_key_event)
    
    def stop_macro_recording(self):
        """停止记录宏"""
        self.is_recording = False
        
        # 更新按钮状态
        self.record_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.play_btn.config(state=tk.NORMAL)
        self.save_macro_btn.config(state=tk.NORMAL)
        
        # 更新状态信息
        self.recording_status_var.set("宏记录已停止")
        self.status_var.set(f"宏记录已停止，共记录 {len(self.recorded_macro)} 个事件")
        
        logger.info(f"停止记录宏，共记录 {len(self.recorded_macro)} 个事件")
        
        # 解绑键盘事件
        if hasattr(self, 'key_record_bind_id'):
            self.root.unbind("<Key>", self.key_record_bind_id)
    
    def play_macro(self):
        """回放宏"""
        if not self.recorded_macro:
            messagebox.showwarning("警告", "没有可回放的宏")
            return
        
        if not pyautogui:
            messagebox.showwarning("警告", "缺少pyautogui库，无法回放宏")
            return
        
        # 更新状态信息
        self.status_var.set("正在回放宏...")
        logger.info("开始回放宏")
        
        # 确保有选中的终端窗口
        if not self.window_selector.selected_window:
            messagebox.showwarning("警告", "请先选择目标终端")
            self.status_var.set("就绪")
            return
        
        try:
            # 等待一小段时间，确保用户准备好
            time.sleep(0.5)
            
            # 回放每个事件
            for i, event_data in enumerate(self.recorded_macro):
                if event_data["type"] == "key":
                    # 模拟按键
                    key = event_data["event"]
                    char = event_data["char"]
                    
                    # 使用更可靠的按键模拟方式
                    if char:
                        # 如果有字符，直接输入字符
                        pyautogui.typewrite(char)
                        logger.debug(f"回放按键: {char}")
                    else:
                        # 否则模拟特殊按键
                        pyautogui.press(key)
                        logger.debug(f"回放按键: {key}")
                    
                    # 添加适当的延迟，模拟真实输入
                    time.sleep(0.05)
            
            self.status_var.set(f"宏回放完成，共执行 {len(self.recorded_macro)} 个事件")
            logger.info(f"宏回放完成")
        except Exception as e:
            logger.error(f"宏回放失败: {e}")
            self.status_var.set(f"宏回放失败: {str(e)}")
    
    def save_macro(self):
        """保存宏到文件"""
        if not self.recorded_macro:
            messagebox.showwarning("警告", "没有可保存的宏")
            return
        
        # 获取宏名称
        macro_name = self.macro_name_var.get().strip()
        if not macro_name:
            macro_name = "my_macro"
        
        # 打开文件对话框，选择保存位置
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("宏文件", "*.json"), ("所有文件", "*.*")],
            initialfile=macro_name
        )
        
        if file_path:
            try:
                # 保存宏数据
                macro_data = {
                    "name": macro_name,
                    "recorded_events": self.recorded_macro,
                    "created_at": time.time(),
                    "version": "1.0"
                }
                
                with open(file_path, "w", encoding="utf-8") as f:
                    json.dump(macro_data, f, indent=2)
                
                self.status_var.set(f"宏保存成功: {file_path}")
                logger.info(f"宏保存成功: {file_path}")
            except Exception as e:
                logger.error(f"保存宏失败: {e}")
                messagebox.showerror("错误", f"保存宏失败: {str(e)}")
    
    def load_macro(self):
        """从文件加载宏"""
        # 打开文件对话框，选择宏文件
        file_path = filedialog.askopenfilename(
            filetypes=[("宏文件", "*.json"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    macro_data = json.load(f)
                
                # 加载宏数据
                self.recorded_macro = macro_data.get("recorded_events", [])
                self.macro_name_var.set(macro_data.get("name", "loaded_macro"))
                
                # 更新按钮状态
                self.play_btn.config(state=tk.NORMAL if self.recorded_macro else tk.DISABLED)
                self.save_macro_btn.config(state=tk.NORMAL if self.recorded_macro else tk.DISABLED)
                
                self.status_var.set(f"宏加载成功: {len(self.recorded_macro)} 个事件")
                logger.info(f"宏加载成功: {file_path}")
            except Exception as e:
                logger.error(f"加载宏失败: {e}")
                messagebox.showerror("错误", f"加载宏失败: {str(e)}")
    
    def apply_font_settings(self, font_family, font_size):
        """应用字体设置"""
        try:
            # 创建字体对象
            custom_font = font.Font(family=font_family, size=font_size)
            
            # 应用到文本编辑器
            self.text_editor.config(font=custom_font)
            
            # 应用到行号显示
            self.line_numbers.config(font=custom_font)
            
            # 更新行号以适应新字体
            self.update_line_numbers()
            
            logger.info(f"已应用字体设置: {font_family} {font_size}")
        except Exception as e:
            logger.error(f"应用字体设置失败: {e}")
    
    def apply_theme(self, theme):
        """应用主题"""
        try:
            if theme == 'Dark':
                # 深色主题
                bg_color = '#2b2b2b'
                fg_color = '#ffffff'
                select_bg = '#404040'
                
                # 应用到文本编辑器
                self.text_editor.config(bg=bg_color, fg=fg_color, selectbackground=select_bg)
                
                # 应用到行号显示
                self.line_numbers.config(bg=bg_color, fg=fg_color)
                
                # 更新当前行高亮颜色
                self.text_editor.tag_configure("current_line", background="#333333")
            else:
                # 浅色主题（默认）
                bg_color = '#ffffff'
                fg_color = '#000000'
                select_bg = '#3399ff'
                
                # 应用到文本编辑器
                self.text_editor.config(bg=bg_color, fg=fg_color, selectbackground=select_bg)
                
                # 应用到行号显示
                self.line_numbers.config(bg=bg_color, fg=fg_color)
                
                # 更新当前行高亮颜色
                self.text_editor.tag_configure("current_line", background="#f0f0f0")
            
            logger.info(f"已应用主题: {theme}")
        except Exception as e:
            logger.error(f"应用主题失败: {e}")
    
    def send_current_line(self):
        """发送当前行"""
        try:
            # 获取当前光标所在行
            cursor_pos = self.text_editor.index(tk.INSERT)
            line_num = cursor_pos.split('.')[0]
            line_start = f"{line_num}.0"
            line_end = f"{line_num}.end"
            
            # 获取当前行内容
            line_content = self.text_editor.get(line_start, line_end).strip()
            
            # 发送命令
            self.execute_command(line_content)
        except Exception as e:
            logger.error(f"发送当前行失败: {e}")
    
    def send_selected_text(self):
        """发送选中文本"""
        try:
            # 获取选中文本
            try:
                selected_text = self.text_editor.get(tk.SEL_FIRST, tk.SEL_LAST)
            except tk.TclError:
                # 没有选中文本，发送当前行
                self.send_current_line()
                return
            
            # 发送命令
            self.execute_command(selected_text)
        except Exception as e:
            logger.error(f"发送选中文本失败: {e}")
    
    def send_all_content(self):
        """发送全部内容"""
        try:
            # 获取全部内容
            all_content = self.text_editor.get(1.0, tk.END).strip()
            
            # 发送命令
            self.execute_command(all_content)
        except Exception as e:
            logger.error(f"发送全部内容失败: {e}")
    
    def _send_to_terminal(self, hwnd, command):
        """向终端窗口发送命令，使用最可靠的方式"""
        # 向终端窗口发送命令，使用最可靠的方式
        # 对于PowerShell、mobaxterm等现代终端，直接使用剪贴板+模拟按键的方式
        # 这种方式最可靠，因为它模拟了用户的实际操作
        
        try:
            # 不强制添加换行符，Enter键的发送由execute_command方法统一控制
            
            # 使用剪贴板+模拟按键的方式
            # 1. 复制命令到剪贴板
            pyperclip.copy(command)
            time.sleep(0.1)
            
            # 2. 模拟按键Ctrl+V粘贴命令
            if pyautogui is None:
                raise Exception("pyautogui模块未安装，无法使用剪贴板+模拟按键方式")
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.1)
            
            # 不强制发送Enter键，由execute_command方法根据auto_enter设置统一控制
        except Exception as e:
            logger.error(f"使用剪贴板+模拟按键发送命令失败: {e}")
            # 如果这种方式失败，尝试使用原始的Windows API方式
            try:
                # 首先检查窗口句柄是否有效
                if not win32gui.IsWindow(hwnd):
                    logger.error(f"尝试向无效窗口句柄 {hwnd} 发送命令")
                    raise Exception(f"无效的窗口句柄: {hwnd}")
                
                # 不强制添加换行符，Enter键的发送由execute_command方法统一控制
                
                # 使用原始的WM_CHAR消息方式
                for char in command:
                    if char == '\n':
                        # 发送回车键执行命令
                        win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                        win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                    else:
                        # 发送普通字符
                        win32gui.PostMessage(hwnd, win32con.WM_CHAR, ord(char), 0)
                    
                    # 短暂延迟，确保命令正确发送
                    time.sleep(0.01)
            except Exception as e2:
                logger.error(f"使用Windows API发送命令失败: {e2}")
                raise e
    
    def _send_to_standard_window(self, hwnd, command):
        """向标准窗口发送命令"""
        # 首先检查窗口句柄是否有效
        if not win32gui.IsWindow(hwnd):
            logger.error(f"尝试向无效窗口句柄 {hwnd} 发送命令")
            raise Exception(f"无效的窗口句柄: {hwnd}")
        
        # 标准窗口使用WM_CHAR消息
        for char in command:
            if char == '\n':
                # 发送回车键
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
            elif char == ' ':
                # 发送空格键
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_SPACE, 0)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_SPACE, 0)
            elif char == '\t':
                # 发送Tab键
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_TAB, 0)
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, win32con.VK_TAB, 0)
            else:
                # 发送普通字符
                for c in char.encode('utf-8'):
                    win32gui.PostMessage(hwnd, win32con.WM_CHAR, c, 0)
        
            # 短暂延迟
            time.sleep(0.01)
        
        # 不强制发送回车键，由调用者决定是否发送
    
    def _fallback_send(self, command):
        """回退发送方法 - 使用剪贴板和键盘模拟"""
        try:
            self.update_status("使用回退方式发送命令...")
            # 使用剪贴板+模拟按键的方式
            time.sleep(0.2)
            pyperclip.copy(command)
            time.sleep(0.2)
            if pyautogui is None:
                raise Exception("pyautogui模块未安装，无法使用剪贴板+键盘模拟方式")
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.2)
            pyautogui.press('enter')
            
            # 更新发送计数
            self.sent_count += 1
            self.sent_count_var.set(f"已发送: {self.sent_count}")
            self.update_status(f"命令通过回退方式发送成功: {command[:50]}...")
        except Exception as e:
            logger.error(f"回退发送方式也失败: {e}")
            self.update_status(f"错误: 命令发送失败 - {str(e)}")
            self.failed_count += 1
            self.failed_count_var.set(f"失败: {self.failed_count}")
    
    def _char_to_vk(self, char):
        """将字符转换为虚拟键码"""
        char_map = {
            'A': 0x41, 'B': 0x42, 'C': 0x43, 'D': 0x44, 'E': 0x45, 'F': 0x46,
            'G': 0x47, 'H': 0x48, 'I': 0x49, 'J': 0x4A, 'K': 0x4B, 'L': 0x4C,
            'M': 0x4D, 'N': 0x4E, 'O': 0x4F, 'P': 0x50, 'Q': 0x51, 'R': 0x52,
            'S': 0x53, 'T': 0x54, 'U': 0x55, 'V': 0x56, 'W': 0x57, 'X': 0x58,
            'Y': 0x59, 'Z': 0x5A,
            '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34,
            '5': 0x35, '6': 0x36, '7': 0x37, '8': 0x38, '9': 0x39,
            '!': 0x31, '@': 0x32, '#': 0x33, '$': 0x34, '%': 0x35,
            '^': 0x36, '&': 0x37, '*': 0x38, '(': 0x39, ')': 0x30,
            '-': 0xBD, '_': 0xBD, '=': 0xBB, '+': 0xBB,
            '[': 0xDB, '{': 0xDB, ']': 0xDD, '}': 0xDD,
            '\\': 0xDC, '|': 0xDC, ';': 0xBA, ':': 0xBA,
            "'": 0xDE, '"': 0xDE, ',': 0xBC, '<': 0xBC,
            '.': 0xBE, '>': 0xBE, '/': 0xBF, '?': 0xBF
        }
        return char_map.get(char, None)
    
    def send_keystroke(self, key_code, hwnd=None):
        """发送单个按键事件到窗口"""
        try:
            if hwnd:
                # 如果有窗口句柄，直接发送到窗口
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, key_code, 0)
                time.sleep(0.01)  # 短暂延迟
                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, key_code, 0)
            else:
                # 否则使用pyautogui发送
                if pyautogui is None:
                    raise Exception("pyautogui模块未安装，无法使用键盘模拟方式")
                pyautogui.press(key_code)
            return True
        except Exception as e:
            logger.error(f"发送按键 {key_code} 失败: {e}")
            return False
    

    
    def send_keyboard_events(self, text):
        """使用Windows API发送键盘事件，模拟键盘字符流"""
        try:
            # 获取目标窗口句柄
            hwnd = None
            
            # 优先使用选中的窗口
            if self.window_selector.selected_window and 'hwnd' in self.window_selector.selected_window:
                hwnd = self.window_selector.selected_window['hwnd']
                # 使用WindowSelector的is_window_valid方法检查窗口有效性
                if not self.window_selector.is_window_valid(hwnd):
                    hwnd = None
            
            # 如果没有选中的窗口或句柄无效，使用当前活动窗口
            if not hwnd:
                hwnd = win32gui.GetForegroundWindow()
            
            if not hwnd or not self.window_selector.is_window_valid(hwnd):
                logger.error("无法获取目标窗口")
                return False
            
            logger.info(f"向窗口 {hwnd} 发送键盘事件: {text!r}")
            
            # 使用内置的KeyboardSimulator类发送键盘事件
            keyboard_sim = KeyboardSimulator()
            
            # 根据文本内容选择不同的发送方式
            if text in ['\n', '\r\n']:
                # 对于换行符，使用send_enter方法
                keyboard_sim.send_enter(hwnd)
            else:
                # 对于普通文本，使用send_text方法
                keyboard_sim.send_text(hwnd, text)
            
            return True
        except Exception as e:
            logger.error(f"发送键盘事件失败: {e}")
            return False
    
    def execute_command(self, command):
        """执行命令，使用多种方式确保命令发送成功"""
        # 记录命令执行开始时间
        start_time = time.time()
        
        # 如果命令以#开头，表示是注释，不执行
        if command.strip().startswith('#'):
            self.update_status(f"注释命令，不执行: {command}")
            logger.info(f"注释命令，不执行: {command}")
            return
        
        # 1. 确保命令格式正确
        command_text = command.strip()
        if not command_text:
            self.update_status("错误: 命令内容为空")
            logger.error("命令内容为空")
            return
        
        # 获取配置，使用UI控件中的变量，确保与用户设置一致
        auto_enter = self.auto_enter_var.get()  # 是否自动换行
        logger.info(f"自动换行配置: {auto_enter}")
        
        # 2. 确保窗口获得焦点（如果有选中的终端）
        focus_success = False
        saved_mouse_pos = None
        
        # 保存鼠标当前位置，避免后续操作导致鼠标跳动
        if pyautogui is not None:
            try:
                saved_mouse_pos = pyautogui.position()
                logger.info(f"保存鼠标位置: {saved_mouse_pos}")
            except Exception as e:
                logger.warning(f"获取鼠标位置失败: {e}")
                saved_mouse_pos = None
        
        # 获取目标窗口句柄
        target_hwnd = None
        if self.window_selector.selected_window:
            target_hwnd = self.window_selector.selected_window.get('hwnd')
        
        # 如果没有选中窗口，使用当前活动窗口
        if not target_hwnd:
            target_hwnd = win32gui.GetForegroundWindow()
            logger.info(f"使用当前活动窗口: {target_hwnd}")
        
        # 增强窗口焦点管理，实现多种焦点获取方法的组合
        def ensure_window_focus(hwnd):
            """确保窗口获得焦点，根据策略调整焦点获取方法"""
            if not hwnd or not win32gui.IsWindow(hwnd):
                logger.error(f"无效的窗口句柄: {hwnd}")
                return False
            
            # 获取窗口详细信息
            window_title = win32gui.GetWindowText(hwnd) if WIN32_AVAILABLE else ""
            window_class = win32gui.GetClassName(hwnd) if WIN32_AVAILABLE else ""
            is_powershell = "PowerShell" in window_title or "Windows PowerShell" in window_title
            
            # 获取当前焦点管理策略
            focus_strategy = self.focus_management_var.get()
            logger.info(f"开始获取窗口焦点 - 句柄: {hwnd}, 标题: '{window_title}', 类名: '{window_class}', 是PowerShell: {is_powershell}, 策略: {focus_strategy}")
            logger.debug(f"当前前台窗口: {win32gui.GetForegroundWindow()}")
            
            # 如果是手动模式，不自动获取焦点
            if focus_strategy == "manual":
                logger.info("使用手动模式，不自动获取焦点")
                return True
            
            # 尝试多种方式获取焦点，根据策略调整方法集
            focus_methods = []
            
            # 针对PowerShell添加特殊处理
            if is_powershell:
                focus_methods.append(("PowerShell专用方法", lambda: _focus_method_powershell(hwnd, saved_mouse_pos)))
            
            # 根据策略选择焦点获取方法
            if focus_strategy == "aggressive":
                # 激进模式：尝试所有焦点获取方法
                focus_methods.extend([
                    ("SetWindowPos+SetForegroundWindow", lambda: _focus_method_1(hwnd)),
                    ("SetFocus", lambda: _focus_method_2(hwnd)),
                    ("Alt+Tab切换", lambda: _focus_method_3(hwnd)),
                    ("模拟鼠标点击", lambda: _focus_method_4(hwnd, saved_mouse_pos)),
                    ("综合方法", lambda: _focus_method_5(hwnd, saved_mouse_pos))
                ])
            elif focus_strategy == "conservative":
                # 保守模式：只尝试温和的焦点获取方法
                focus_methods.extend([
                    ("SetWindowPos+SetForegroundWindow", lambda: _focus_method_1(hwnd)),
                    ("SetFocus", lambda: _focus_method_2(hwnd))
                    # 保守模式下不使用Alt+Tab、模拟点击等强干扰方法
                ])
            
            # 先检查当前窗口是否已经获得焦点
            current_foreground = win32gui.GetForegroundWindow()
            if current_foreground == hwnd:
                logger.info(f"窗口 {hwnd} 已获得焦点，直接返回成功")
                return True
            
            for method_name, focus_func in focus_methods:
                try:
                    logger.debug(f"尝试使用 {method_name} 获取焦点")
                    start_time = time.time()
                    success = focus_func()
                    duration = time.time() - start_time
                    logger.debug(f"{method_name} 执行耗时: {duration:.3f}秒，结果: {success}")
                    
                    # 检查焦点是否实际获得，即使方法返回成功
                    current_foreground = win32gui.GetForegroundWindow()
                    actual_success = success and current_foreground == hwnd
                    
                    if actual_success:
                        logger.info(f"使用 {method_name} 成功获取焦点")
                        logger.debug(f"当前前台窗口: {current_foreground}")
                        return True
                    elif success:
                        # 方法返回成功但实际焦点未获得，记录警告
                        logger.warning(f"{method_name} 返回成功但实际未获得焦点，当前前台窗口: {current_foreground}")
                    else:
                        logger.debug(f"{method_name} 获取焦点失败")
                except Exception as e:
                    logger.warning(f"使用 {method_name} 获取焦点失败: {e}")
                    logger.debug(f"{method_name} 异常详情: {traceback.format_exc()}")
            
            # 最后再次检查焦点，即使所有方法都失败
            final_foreground = win32gui.GetForegroundWindow()
            if final_foreground == hwnd:
                logger.info(f"最终检查发现窗口 {hwnd} 已获得焦点，返回成功")
                return True
            
            logger.warning(f"所有焦点获取方法均失败，最终前台窗口: {final_foreground}")
            logger.debug(f"目标窗口: {hwnd}，标题: {win32gui.GetWindowText(hwnd)}")
            return False
        
        def _focus_method_powershell(hwnd, saved_mouse_pos):
            """PowerShell专用焦点获取方法，简化版"""
            logger.info(f"使用PowerShell专用焦点获取方法")
            
            try:
                # 1. 验证窗口有效性
                if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
                    logger.warning(f"PowerShell窗口 {hwnd} 无效或不可见")
                    return False
                
                # 2. 先将窗口置顶，确保可见
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
                time.sleep(0.05)
                
                # 3. 设置为前台窗口
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.05)
                
                # 4. 直接返回True，由上层逻辑检查实际焦点状态
                # 这样可以避免在方法内部进行复杂的焦点检查，提高成功率
                return True
            except Exception as e:
                logger.error(f"PowerShell焦点获取失败: {e}")
                logger.debug(f"PowerShell焦点获取异常详情: {traceback.format_exc()}")
                return False
        
        def _focus_method_1(hwnd):
            """方法1: 先置顶，再设置前台窗口"""
            # 1. 先将窗口置顶
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            time.sleep(0.05)
            
            # 2. 然后设置为前台窗口
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.05)
            
            # 检查是否获得焦点
            current_foreground = win32gui.GetForegroundWindow()
            return current_foreground == hwnd
        
        def _focus_method_2(hwnd):
            """方法2: 使用SetFocus设置输入焦点"""
            # 获取窗口的子窗口（通常是实际的编辑区域）
            child_hwnd = win32gui.GetWindow(hwnd, win32con.GW_CHILD)
            if child_hwnd:
                try:
                    win32gui.SetFocus(child_hwnd)
                    time.sleep(0.05)
                    return True
                except Exception as e:
                    logger.warning(f"设置子窗口焦点失败: {e}")
            
            # 如果没有子窗口，直接设置主窗口焦点
            try:
                win32gui.SetFocus(hwnd)
                time.sleep(0.05)
                return True
            except Exception as e:
                logger.warning(f"设置主窗口焦点失败: {e}")
                return False
        
        def _focus_method_3(hwnd):
            """方法3: 模拟Alt+Tab切换窗口"""
            if pyautogui is None:
                logger.warning("pyautogui不可用，无法使用Alt+Tab切换方法")
                return False
            
            # 保存当前窗口标题，用于验证切换是否成功
            original_title = win32gui.GetWindowText(hwnd)
            
            # 模拟Alt+Tab
            pyautogui.keyDown('alt')
            pyautogui.press('tab')
            time.sleep(0.1)
            pyautogui.keyUp('alt')
            time.sleep(0.1)
            
            # 检查当前前台窗口是否是目标窗口
            current_foreground = win32gui.GetForegroundWindow()
            current_title = win32gui.GetWindowText(current_foreground)
            
            return current_foreground == hwnd or original_title in current_title
        
        def _focus_method_4(hwnd, saved_mouse_pos):
            """方法4: 模拟鼠标点击窗口"""
            # 获取窗口位置
            rect = win32gui.GetWindowRect(hwnd)
            if rect[0] == 0 and rect[1] == 0 and rect[2] == 0 and rect[3] == 0:
                logger.warning(f"窗口 {hwnd} 位置无效: {rect}")
                return False
            
            # 点击窗口中心位置
            center_x = (rect[0] + rect[2]) // 2
            center_y = (rect[1] + rect[3]) // 2
            
            # 使用MouseSimulator模拟点击
            mouse_sim = MouseSimulator()
            mouse_sim.click(hwnd, center_x - rect[0], center_y - rect[1], use_simulated_click=True, saved_mouse_pos=saved_mouse_pos)
            time.sleep(0.1)
            
            # 检查是否获得焦点
            current_foreground = win32gui.GetForegroundWindow()
            return current_foreground == hwnd
        
        def _focus_method_5(hwnd, saved_mouse_pos):
            """方法5: 综合方法，组合多种技术"""
            # 1. 先将窗口置顶
            win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, 0, 0, win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
            time.sleep(0.05)
            
            # 2. 设置为前台窗口
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.05)
            
            # 3. 设置焦点
            win32gui.SetFocus(hwnd)
            time.sleep(0.05)
            
            # 4. 模拟鼠标点击
            rect = win32gui.GetWindowRect(hwnd)
            if rect[0] != 0 or rect[1] != 0:
                center_x = (rect[0] + rect[2]) // 2
                center_y = (rect[1] + rect[3]) // 2
                mouse_sim = MouseSimulator()
                mouse_sim.click(hwnd, center_x - rect[0], center_y - rect[1], use_simulated_click=True, saved_mouse_pos=saved_mouse_pos)
                time.sleep(0.1)
            
            # 5. 检查是否获得焦点
            current_foreground = win32gui.GetForegroundWindow()
            return current_foreground == hwnd
        
        def _check_focus_stability(hwnd, stability_time=0.2):
            """检查焦点是否稳定，在指定时间内保持不变"""
            start_time = time.time()
            while time.time() - start_time < stability_time:
                current_foreground = win32gui.GetForegroundWindow()
                if current_foreground != hwnd:
                    logger.warning(f"焦点在 {time.time() - start_time:.2f} 秒内丢失")
                    return False
                time.sleep(0.05)  # 每50ms检查一次
            return True
        
        # 获取目标窗口句柄
        target_hwnd = None
        if self.window_selector.selected_window:
            target_hwnd = self.window_selector.selected_window.get('hwnd')
            logger.info(f"从选中窗口获取到句柄: {target_hwnd}")
        
        # 如果没有选中窗口或窗口句柄无效，使用当前活动窗口
        if not target_hwnd or not win32gui.IsWindow(target_hwnd):
            original_hwnd = target_hwnd
            target_hwnd = win32gui.GetForegroundWindow()
            window_title = win32gui.GetWindowText(target_hwnd)
            logger.info(f"使用当前活动窗口: 句柄={target_hwnd}, 标题={window_title}, 原句柄={original_hwnd}")
        else:
            window_title = win32gui.GetWindowText(target_hwnd)
            logger.info(f"使用选中的窗口: 句柄={target_hwnd}, 标题={window_title}")
        
        # 确保窗口获得焦点
        logger.info(f"开始确保窗口获得焦点，目标窗口: {target_hwnd}, 标题: {window_title}")
        focus_success = ensure_window_focus(target_hwnd)
        logger.info(f"窗口焦点获取结果: {focus_success}, 最终前台窗口: {win32gui.GetForegroundWindow()}")
        
        # 记录发送尝试次数
        send_attempts = 0
        max_attempts = self.focus_retry_count_var.get()  # 从配置获取重试次数
        retry_delay = self.focus_retry_delay_var.get()  # 从配置获取初始重试延迟
        max_retry_delay = 1.0  # 最大重试延迟（固定值）
        start_time = time.time()
        timeout = self.focus_timeout_var.get()  # 从配置获取超时时间
        logger.info(f"命令发送配置: 最大尝试次数={max_attempts}, 初始重试延迟={retry_delay}, 最大重试延迟={max_retry_delay}, 超时时间={timeout}")
        
        # 记录初始焦点状态
        initial_focus = win32gui.GetForegroundWindow()
        logger.info(f"初始焦点状态 - 目标窗口: {target_hwnd}, 当前焦点窗口: {initial_focus}, 焦点一致: {initial_focus == target_hwnd}")
        
        # 使用内置的KeyboardSimulator类检测终端类型
        keyboard_sim = KeyboardSimulator()
        terminal_type = keyboard_sim.detect_terminal_type(target_hwnd)
        logger.info(f"检测到终端类型: {terminal_type}")
        
        # 获取窗口标题，用于更准确的终端类型判断
        window_title = win32gui.GetWindowText(target_hwnd) if WIN32_AVAILABLE else ""
        logger.info(f"窗口标题: {window_title}")
        
        # 标记是否发送成功
        send_success = False
        final_error = ""
        
        # 判断是否为SecureCRT或PowerShell或Windows Terminal
        is_securecrt = terminal_type == "securecrt" or "SecureCRT" in window_title
        is_powershell = terminal_type == "powershell" or "PowerShell" in window_title or "Windows PowerShell" in window_title
        is_windows_terminal = "windows_terminal" in terminal_type
        
        # 确定终端名称
        terminal_name = "Windows Terminal" if is_windows_terminal else "PowerShell" if is_powershell else "SecureCRT" if is_securecrt else "Unknown"
        logger.info(f"最终确定终端类型: {terminal_name}")
        
        # 智能重试机制：根据焦点状态和发送结果进行重试
        while send_attempts < max_attempts and not send_success:
            # 检查超时
            elapsed_time = time.time() - start_time
            if elapsed_time > timeout:
                logger.error(f"命令发送超时，总耗时: {elapsed_time:.2f}秒，重试次数: {send_attempts}/{max_attempts}")
                final_error = f"命令发送超时，超过{timeout}秒"
                break
            
            # 指数退避算法
            if send_attempts > 0:  # 第一次尝试不需要延迟
                logger.debug(f"重试延迟: {retry_delay:.3f}秒")
                time.sleep(retry_delay)
                retry_delay = min(retry_delay * 2, max_retry_delay)  # 每次重试延迟翻倍，直到最大延迟
                logger.debug(f"下一次重试延迟: {retry_delay:.3f}秒")
            
            send_attempts += 1
            # 记录当前焦点状态
            current_focus = win32gui.GetForegroundWindow()
            logger.info(f"尝试 {send_attempts}/{max_attempts}: 使用Windows API发送命令: {command_text!r}, 当前焦点窗口: {current_focus}, 焦点一致: {current_focus == target_hwnd}")
            
            # 重置焦点状态，每次尝试重新获取
            focus_success = False
            
            try:
                # 1. 确保窗口获得焦点
                focus_success = ensure_window_focus(target_hwnd)
                if not focus_success:
                    logger.warning(f"尝试 {send_attempts}/{max_attempts}: 无法获取窗口焦点")
                    continue
                
                # 2. 检查焦点是否稳定
                if win32gui.GetForegroundWindow() != target_hwnd:
                    logger.warning(f"尝试 {send_attempts}/{max_attempts}: 窗口焦点不稳定，重新获取")
                    focus_success = ensure_window_focus(target_hwnd)
                    if not focus_success:
                        continue
                
                # 3. 发送命令文本（添加焦点保护）
                text_send_result = keyboard_sim.send_text(target_hwnd, command_text)
                logger.info(f"命令文本发送结果: {text_send_result}")
                
                # 4. 发送命令后立即检查焦点
                if win32gui.GetForegroundWindow() != target_hwnd:
                    logger.warning(f"{terminal_name}窗口在发送命令文本后失去焦点")
                    # 尝试重新获取焦点
                    focus_recovered = ensure_window_focus(target_hwnd)
                    if not focus_recovered:
                        logger.warning(f"无法恢复{terminal_name}窗口焦点")
                
                # 5. 发送Enter键（如果需要），添加焦点保护
                enter_send_result = True
                # 5.5 根据用户设置发送Enter键
                if text_send_result and auto_enter:
                    logger.info(f"根据用户设置，发送Enter键来执行命令")
                    current_focus = win32gui.GetForegroundWindow()
                    if current_focus == target_hwnd:
                        enter_send_result = keyboard_sim.send_enter(target_hwnd)
                        logger.info(f"Enter键发送结果: {enter_send_result}")
                        
                        # 发送Enter后立即检查焦点
                        if win32gui.GetForegroundWindow() != target_hwnd:
                            logger.warning(f"{terminal_name}窗口在发送Enter键后失去焦点")
                    else:
                        enter_send_result = False
                        logger.warning(f"{terminal_name}窗口未获得焦点，跳过发送Enter键，当前焦点窗口: {current_focus}")
                
                # 6. 最终焦点检查，确保命令发送完成后焦点仍然存在
                final_foreground = win32gui.GetForegroundWindow()
                if final_foreground != target_hwnd:
                    logger.warning(f"{terminal_name}窗口最终失去焦点，命令可能未被完全接收")
                else:
                    logger.info(f"{terminal_name}窗口在整个命令发送过程中保持了稳定焦点")
                
                # 7. 额外的焦点稳定性检查，确保命令已被完全处理
                stability_check = _check_focus_stability(target_hwnd, stability_time=0.1)
                if not stability_check:
                    logger.warning(f"{terminal_name}窗口在命令发送完成后焦点不稳定")
                
                # 7. 检查发送结果
                if text_send_result and (not auto_enter or enter_send_result):
                    logger.info(f"{terminal_name}命令发送成功: {command_text} (自动换行: {auto_enter})")
                    send_success = True
                    break
                else:
                    logger.warning(f"尝试 {send_attempts}/{max_attempts}: 命令发送失败：文本发送或Enter发送失败")
                    final_error = "命令文本发送或Enter发送失败"
            except Exception as e:
                logger.error(f"尝试 {send_attempts}/{max_attempts}: Windows API发送命令失败: {e}")
                final_error = str(e)
        
        # 如果Windows API方式失败，尝试使用剪贴板+键盘模拟方式作为最后的备选
        if not send_success and send_attempts < max_attempts:
            try:
                send_attempts += 1
                logger.info(f"尝试 {send_attempts}/{max_attempts}: 使用剪贴板+键盘模拟方式发送命令: {command_text!r}")
                self.update_status("正在使用剪贴板+键盘模拟方式发送命令...")
                
                # 确保窗口获得焦点
                if not ensure_window_focus(target_hwnd):
                    logger.warning(f"无法获取窗口焦点，跳过剪贴板方式")
                    final_error = "无法获取窗口焦点，命令发送失败"
                    # 无法获取焦点，跳过剪贴板方式，继续执行后面的代码
                    pass
                
                # 1. 复制命令到剪贴板
                pyperclip.copy(command_text)
                time.sleep(0.05)
                logger.info(f"已将命令复制到剪贴板: {command_text!r}")
                
                # 2. 模拟按键Ctrl+V粘贴命令
                if pyautogui is None:
                    raise Exception("pyautogui模块未安装，无法使用剪贴板+键盘模拟方式")
                
                # 检查窗口焦点状态
                if not focus_success:
                    # 再次尝试获取焦点
                    try:
                        current_x, current_y = pyautogui.position()
                        pyautogui.click()
                        time.sleep(0.1)
                        focus_success = True
                        logger.info(f"已再次尝试获取窗口焦点")
                    except Exception as focus_e:
                        logger.warning(f"再次尝试获取焦点失败: {focus_e}")
                
                # 模拟Ctrl+V粘贴
                try:
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(0.05)  # 进一步减少延迟，从0.2改为0.05
                    logger.info(f"已模拟Ctrl+V粘贴命令")
                    
                    # 3. 模拟按键Enter执行命令（如果需要）
                    if auto_enter:
                        pyautogui.press('enter')
                        time.sleep(0.05)  # 进一步减少延迟，从0.1改为0.05
                        logger.info(f"已模拟Enter键执行命令")
                except Exception as e:
                    logger.error(f"模拟键盘操作失败: {e}")
                    raise
                
                # 检查目标窗口是否仍然是有效窗口
                if WIN32_AVAILABLE and win32gui.IsWindow(target_hwnd):
                    # 对于SecureCRT，添加更严格的验证
                    if is_securecrt:
                        logger.info(f"SecureCRT剪贴板+键盘模拟发送命令完成")
                        send_success = True
                    # 对于PowerShell，添加更严格的验证
                    elif is_powershell:
                        logger.info(f"PowerShell剪贴板+键盘模拟发送命令完成")
                        send_success = True
                    else:
                        logger.info(f"剪贴板+键盘模拟发送命令完成，窗口焦点有效")
                        send_success = True
                else:
                    logger.warning(f"剪贴板+键盘模拟方式可能发送失败，目标窗口无效")
                    final_error = "剪贴板+键盘模拟方式发送失败：目标窗口无效"
            except Exception as e:
                logger.error(f"剪贴板+键盘模拟发送命令失败: {e}")
                final_error = str(e)
        
        # 如果前两种方式失败，尝试使用pyautogui直接输入命令
        if not send_success and send_attempts < max_attempts:
            try:
                send_attempts += 1
                logger.info(f"尝试 {send_attempts}/{max_attempts}: 使用pyautogui直接输入命令: {command_text!r}")
                self.update_status("正在使用pyautogui直接输入方式发送命令...")
                
                # 检查pyautogui是否可用
                if pyautogui is None:
                    raise Exception("pyautogui模块未安装，无法使用直接输入方式")
                
                # 检查窗口焦点状态
                if not focus_success:
                    # 再次尝试获取焦点
                    try:
                        current_x, current_y = pyautogui.position()
                        pyautogui.click()
                        time.sleep(0.1)
                        focus_success = True
                        logger.info(f"已再次尝试获取窗口焦点")
                    except Exception as focus_e:
                        logger.warning(f"再次尝试获取焦点失败: {focus_e}")
                
                # 直接输入命令
                try:
                    pyautogui.typewrite(command_text, interval=0.02)  # 减少输入间隔，从0.05改为0.02
                    time.sleep(0.05)  # 减少延迟，从0.1改为0.05
                    logger.info(f"已使用pyautogui直接输入命令文本")
                    
                    # 模拟按键Enter执行命令（如果需要）
                    if auto_enter:
                        pyautogui.press('enter')
                        time.sleep(0.05)  # 减少延迟，从0.1改为0.05
                        logger.info(f"已模拟Enter键执行命令")
                except Exception as e:
                    logger.error(f"模拟键盘输入失败: {e}")
                    raise
                
                # 检查目标窗口是否仍然是有效窗口
                if WIN32_AVAILABLE and win32gui.IsWindow(target_hwnd):
                    # 对于SecureCRT，添加更严格的验证
                    if is_securecrt:
                        logger.info(f"SecureCRT pyautogui直接输入命令完成")
                        send_success = True
                    # 对于PowerShell，添加更严格的验证
                    elif is_powershell:
                        logger.info(f"PowerShell pyautogui直接输入命令完成")
                        send_success = True
                    else:
                        logger.info(f"pyautogui直接输入命令完成，窗口焦点有效")
                        send_success = True
                else:
                    logger.warning(f"pyautogui直接输入方式可能发送失败，目标窗口无效")
                    final_error = "pyautogui直接输入方式发送失败：目标窗口无效"
            except Exception as e:
                logger.error(f"pyautogui直接输入命令失败: {e}")
                final_error = str(e)
        
        # 根据发送结果更新状态
        if send_success:
            self.update_status(f"命令发送成功: {command[:50]}...")
            self.sent_count += 1
            self.sent_count_var.set(f"已发送: {self.sent_count}")
            logger.info(f"命令发送成功，总耗时: {time.time() - start_time:.2f}秒")
        else:
            # 所有方式都失败的情况
            error_msg = f"错误: 命令发送失败，所有发送方式均尝试过"
            if final_error:
                error_msg += f" - {final_error}"
            self.update_status(error_msg)
            self.failed_count += 1
            self.failed_count_var.set(f"失败: {self.failed_count}")
            logger.error(f"命令发送最终失败，总耗时: {time.time() - start_time:.2f}秒, 错误: {final_error}")
        
        return

    def send_command(self, command):
        """发送命令到选定窗口或串口，根据output_mode处理不同发送方式"""
        try:
            output_mode = self.output_mode.get()
            
            if output_mode == 'serial':
                # 串口发送模式
                if not self.serial_manager:
                    self.update_status("错误: 串口管理器未初始化")
                    self.failed_count += 1
                    self.failed_count_var.set(f"失败: {self.failed_count}")
                    return
                
                if not self.serial_manager.is_connected():
                    self.update_status("警告: 请先连接串口")
                    return
                
                # 优化命令格式，确保以换行符结尾
                command_to_send = command.strip()
                if not command_to_send.endswith('\n'):
                    command_to_send += '\n'
                
                # 发送命令到串口
                if self.serial_manager.send_command(command_to_send):
                    self.update_status(f"已通过串口发送命令: {command[:50]}...")
                    self.sent_count += 1
                    self.sent_count_var.set(f"已发送: {self.sent_count}")
                else:
                    self.update_status("错误: 串口发送命令失败")
                    self.failed_count += 1
                    self.failed_count_var.set(f"失败: {self.failed_count}")
            elif output_mode == 'clipboard':
                # 剪贴板模式，只复制到剪贴板
                pyperclip.copy(command.strip())
                self.update_status(f"命令已复制到剪贴板: {command[:50]}...")
                self.sent_count += 1
                self.sent_count_var.set(f"已发送: {self.sent_count}")
            else:  # terminal模式
                # 终端发送模式，调用execute_command
                self.execute_command(command)
        except Exception as e:
            logger.error(f"发送命令失败: {e}")
            error_msg = str(e)
            
            # 更新状态和失败计数
            self.update_status(f"错误: 命令发送失败 - {error_msg}")
            self.failed_count += 1
            self.failed_count_var.set(f"失败: {self.failed_count}")
    
    def get_file_mtime(self, file_path):
        """获取文件的修改时间"""
        try:
            return os.path.getmtime(file_path)
        except Exception:
            return None
    
    def start_file_monitor(self):
        """启动文件监控"""
        if self.current_file and not hasattr(self, '_file_monitor_running'):
            try:
                self._file_monitor_running = True
                self._file_mtime = self.get_file_mtime(self.current_file)
                self.root.after(1000, self.check_file_external_modification)
                logger.info(f"已启动文件监控: {self.current_file}")
            except Exception as e:
                logger.error(f"启动文件监控失败: {str(e)}")
                self._file_monitor_running = False
    
    def stop_file_monitor(self):
        """停止文件监控"""
        if hasattr(self, '_file_monitor_running'):
            self._file_monitor_running = False
            logger.info("已停止文件监控")
    
    def check_file_external_modification(self):
        """检查文件是否被外部修改"""
        if not hasattr(self, '_file_monitor_running') or not self._file_monitor_running or not self.current_file:
            return
        
        try:
            current_mtime = self.get_file_mtime(self.current_file)
            if current_mtime and hasattr(self, '_file_mtime') and current_mtime != self._file_mtime:
                # 文件被外部修改
                self.handle_external_modification()
            
            # 继续监控
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
        except Exception as e:
            logger.error(f"检查文件外部修改失败: {str(e)}")
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
    
    def handle_external_modification(self):
        """处理文件被外部修改的情况"""
        if not self.current_file:
            return
        
        # 如果当前文件未被修改，直接重新加载
        if not self.is_modified:
            self.reload_file()
            return
        
        # 如果当前文件已被修改，询问用户如何处理
        response = messagebox.askyesnocancel(
            "文件已被外部修改",
            "当前文件已被外部程序修改。\n" \
            "是否保存当前更改并重新加载？\n\n" \
            "是: 保存当前更改并重新加载外部修改\n" \
            "否: 放弃当前更改并重新加载外部修改\n" \
            "取消: 保持当前状态不变"
        )
        
        if response is None:  # 用户点击了取消
            return
        elif response:  # 用户点击了是
            # 保存当前更改
            self.save_file()
            # 重新加载文件
            self.reload_file()
        else:  # 用户点击了否
            # 直接重新加载文件，放弃当前更改
            self.reload_file()
    
    def reload_file(self):
        """重新加载当前文件"""
        if not self.current_file:
            return
        
        try:
            with open(self.current_file, 'r', encoding='utf-8') as file:
                content = file.read()
            
            # 更新文本编辑器内容
            self.text_editor.delete(1.0, tk.END)
            self.text_editor.insert(1.0, content)
            
            # 更新文件修改时间和状态
            self._file_mtime = self.get_file_mtime(self.current_file)
            self.is_modified = False
            self.update_title()
            self.update_status(f"已重新加载文件: {os.path.basename(self.current_file)}")
            logger.info(f"已重新加载文件: {self.current_file}")
        except Exception as e:
            messagebox.showerror("错误", f"无法重新加载文件: {str(e)}")
            logger.error(f"重新加载文件失败: {str(e)}")
    
    def get_file_mtime(self, file_path):
        """获取文件的修改时间"""
        try:
            return os.path.getmtime(file_path)
        except Exception:
            return None
    
    def start_file_monitor(self):
        """启动文件监控"""
        if self.current_file and not hasattr(self, '_file_monitor_running'):
            try:
                self._file_monitor_running = True
                self._file_mtime = self.get_file_mtime(self.current_file)
                self.root.after(1000, self.check_file_external_modification)
                logger.info(f"已启动文件监控: {self.current_file}")
            except Exception as e:
                logger.error(f"启动文件监控失败: {str(e)}")
                self._file_monitor_running = False
    
    def stop_file_monitor(self):
        """停止文件监控"""
        if hasattr(self, '_file_monitor_running'):
            self._file_monitor_running = False
            logger.info("已停止文件监控")
    
    def check_file_external_modification(self):
        """检查文件是否被外部修改"""
        if not hasattr(self, '_file_monitor_running') or not self._file_monitor_running or not self.current_file:
            return
        
        try:
            current_mtime = self.get_file_mtime(self.current_file)
            if current_mtime and hasattr(self, '_file_mtime') and current_mtime != self._file_mtime:
                # 文件被外部修改
                self.handle_external_modification()
            
            # 继续监控
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
        except Exception as e:
            logger.error(f"检查文件外部修改失败: {str(e)}")
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
    
    def handle_external_modification(self):
        """处理文件被外部修改的情况"""
        if not self.current_file:
            return
        
        # 如果当前文件未被修改，直接重新加载
        if not self.is_modified:
            self.reload_file()
            return
        
        # 如果当前文件已被修改，询问用户如何处理
        response = messagebox.askyesnocancel(
            "文件已被外部修改",
            "当前文件已被外部程序修改。\n" \
            "是否保存当前更改并重新加载？\n\n" \
            "是: 保存当前更改并重新加载外部修改\n" \
            "否: 放弃当前更改并重新加载外部修改\n" \
            "取消: 保持当前状态不变"
        )
        
        if response is None:  # 用户点击了取消
            return
        elif response:  # 用户点击了是
            # 保存当前更改
            self.save_file()
            # 重新加载文件
            self.reload_file()
        else:  # 用户点击了否
            # 直接重新加载文件，放弃当前更改
            self.reload_file()
    
    def reload_file(self):
        """重新加载当前文件"""
        if not self.current_file:
            return
        
        try:
            with open(self.current_file, 'r', encoding='utf-8') as file:
                content = file.read()
            
            # 更新文本编辑器内容
            self.text_editor.delete(1.0, tk.END)
            self.text_editor.insert(1.0, content)
            
            # 更新文件修改时间和状态
            self._file_mtime = self.get_file_mtime(self.current_file)
            self.is_modified = False
            self.update_title()
            self.update_status(f"已重新加载文件: {os.path.basename(self.current_file)}")
            logger.info(f"已重新加载文件: {self.current_file}")
        except Exception as e:
            messagebox.showerror("错误", f"无法重新加载文件: {str(e)}")
            logger.error(f"重新加载文件失败: {str(e)}")
    
    def get_file_mtime(self, file_path):
        """获取文件的修改时间"""
        try:
            return os.path.getmtime(file_path)
        except Exception:
            return None
    
    def start_file_monitor(self):
        """启动文件监控"""
        if self.current_file and not hasattr(self, '_file_monitor_running'):
            try:
                self._file_monitor_running = True
                self._file_mtime = self.get_file_mtime(self.current_file)
                self.root.after(1000, self.check_file_external_modification)
                logger.info(f"已启动文件监控: {self.current_file}")
            except Exception as e:
                logger.error(f"启动文件监控失败: {str(e)}")
                self._file_monitor_running = False
    
    def stop_file_monitor(self):
        """停止文件监控"""
        if hasattr(self, '_file_monitor_running'):
            self._file_monitor_running = False
            logger.info("已停止文件监控")
    
    def check_file_external_modification(self):
        """检查文件是否被外部修改"""
        if not hasattr(self, '_file_monitor_running') or not self._file_monitor_running or not self.current_file:
            return
        
        try:
            current_mtime = self.get_file_mtime(self.current_file)
            if current_mtime and hasattr(self, '_file_mtime') and current_mtime != self._file_mtime:
                # 文件被外部修改
                self.handle_external_modification()
            
            # 继续监控
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
        except Exception as e:
            logger.error(f"检查文件外部修改失败: {str(e)}")
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
    
    def handle_external_modification(self):
        """处理文件被外部修改的情况"""
        if not self.current_file:
            return
        
        # 如果当前文件未被修改，直接重新加载
        if not self.is_modified:
            self.reload_file()
            return
        
        # 如果当前文件已被修改，询问用户如何处理
        response = messagebox.askyesnocancel(
            "文件已被外部修改",
            "当前文件已被外部程序修改。\n" \
            "是否保存当前更改并重新加载？\n\n" \
            "是: 保存当前更改并重新加载外部修改\n" \
            "否: 放弃当前更改并重新加载外部修改\n" \
            "取消: 保持当前状态不变"
        )
        
        if response is None:  # 用户点击了取消
            return
        elif response:  # 用户点击了是
            # 保存当前更改
            self.save_file()
            # 重新加载文件
            self.reload_file()
        else:  # 用户点击了否
            # 直接重新加载文件，放弃当前更改
            self.reload_file()
    
    def reload_file(self):
        """重新加载当前文件"""
        if not self.current_file:
            return
        
        try:
            with open(self.current_file, 'r', encoding='utf-8') as file:
                content = file.read()
            
            # 更新文本编辑器内容
            self.text_editor.delete(1.0, tk.END)
            self.text_editor.insert(1.0, content)
            
            # 更新文件修改时间和状态
            self._file_mtime = self.get_file_mtime(self.current_file)
            self.is_modified = False
            self.update_title()
            self.update_status(f"已重新加载文件: {os.path.basename(self.current_file)}")
            logger.info(f"已重新加载文件: {self.current_file}")
        except Exception as e:
            messagebox.showerror("错误", f"无法重新加载文件: {str(e)}")
            logger.error(f"重新加载文件失败: {str(e)}")
    
    def get_file_mtime(self, file_path):
        """获取文件的修改时间"""
        try:
            return os.path.getmtime(file_path)
        except Exception:
            return None
    
    def start_file_monitor(self):
        """启动文件监控"""
        if self.current_file and not hasattr(self, '_file_monitor_running'):
            try:
                self._file_monitor_running = True
                self._file_mtime = self.get_file_mtime(self.current_file)
                self.root.after(1000, self.check_file_external_modification)
                logger.info(f"已启动文件监控: {self.current_file}")
            except Exception as e:
                logger.error(f"启动文件监控失败: {str(e)}")
                self._file_monitor_running = False
    
    def stop_file_monitor(self):
        """停止文件监控"""
        if hasattr(self, '_file_monitor_running'):
            self._file_monitor_running = False
            logger.info("已停止文件监控")
    
    def check_file_external_modification(self):
        """检查文件是否被外部修改"""
        if not hasattr(self, '_file_monitor_running') or not self._file_monitor_running or not self.current_file:
            return
        
        try:
            current_mtime = self.get_file_mtime(self.current_file)
            if current_mtime and hasattr(self, '_file_mtime') and current_mtime != self._file_mtime:
                # 文件被外部修改
                self.handle_external_modification()
            
            # 继续监控
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
        except Exception as e:
            logger.error(f"检查文件外部修改失败: {str(e)}")
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
    
    def handle_external_modification(self):
        """处理文件被外部修改的情况"""
        if not self.current_file:
            return
        
        # 如果当前文件未被修改，直接重新加载
        if not self.is_modified:
            self.reload_file()
            return
        
        # 如果当前文件已被修改，询问用户如何处理
        response = messagebox.askyesnocancel(
            "文件已被外部修改",
            "当前文件已被外部程序修改。\n" \
            "是否保存当前更改并重新加载？\n\n" \
            "是: 保存当前更改并重新加载外部修改\n" \
            "否: 放弃当前更改并重新加载外部修改\n" \
            "取消: 保持当前状态不变"
        )
        
        if response is None:  # 用户点击了取消
            return
        elif response:  # 用户点击了是
            # 保存当前更改
            self.save_file()
            # 重新加载文件
            self.reload_file()
        else:  # 用户点击了否
            # 直接重新加载文件，放弃当前更改
            self.reload_file()
    
    def reload_file(self):
        """重新加载当前文件"""
        if not self.current_file:
            return
        
        try:
            with open(self.current_file, 'r', encoding='utf-8') as file:
                content = file.read()
            
            # 更新文本编辑器内容
            self.text_editor.delete(1.0, tk.END)
            self.text_editor.insert(1.0, content)
            
            # 更新文件修改时间和状态
            self._file_mtime = self.get_file_mtime(self.current_file)
            self.is_modified = False
            self.update_title()
            self.update_status(f"已重新加载文件: {os.path.basename(self.current_file)}")
            logger.info(f"已重新加载文件: {self.current_file}")
        except Exception as e:
            messagebox.showerror("错误", f"无法重新加载文件: {str(e)}")
            logger.error(f"重新加载文件失败: {str(e)}")
    
    def start_file_monitor(self):
        """启动文件监控"""
        if self.current_file and not hasattr(self, '_file_monitor_running'):
            try:
                self._file_monitor_running = True
                self._file_mtime = self.get_file_mtime(self.current_file)
                self.root.after(1000, self.check_file_external_modification)
                logger.info(f"已启动文件监控: {self.current_file}")
            except Exception as e:
                logger.error(f"启动文件监控失败: {str(e)}")
                self._file_monitor_running = False
    
    def stop_file_monitor(self):
        """停止文件监控"""
        if hasattr(self, '_file_monitor_running'):
            self._file_monitor_running = False
            logger.info("已停止文件监控")
    
    def check_file_external_modification(self):
        """检查文件是否被外部修改"""
        if not hasattr(self, '_file_monitor_running') or not self._file_monitor_running or not self.current_file:
            return
        
        try:
            current_mtime = self.get_file_mtime(self.current_file)
            if current_mtime and hasattr(self, '_file_mtime') and current_mtime != self._file_mtime:
                # 文件被外部修改
                self.handle_external_modification()
            
            # 继续监控
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
        except Exception as e:
            logger.error(f"检查文件外部修改失败: {str(e)}")
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
    
    def handle_external_modification(self):
        """处理文件被外部修改的情况"""
        if not self.current_file:
            return
        
        # 如果当前文件未被修改，直接重新加载
        if not self.is_modified:
            self.reload_file()
            return
        
        # 如果当前文件已被修改，询问用户如何处理
        response = messagebox.askyesnocancel(
            "文件已被外部修改",
            "当前文件已被外部程序修改。\n" \
            "是否保存当前更改并重新加载？\n\n" \
            "是: 保存当前更改并重新加载外部修改\n" \
            "否: 放弃当前更改并重新加载外部修改\n" \
            "取消: 保持当前状态不变"
        )
        
        if response is None:  # 用户点击了取消
            return
        elif response:  # 用户点击了是
            # 保存当前更改
            self.save_file()
            # 重新加载文件
            self.reload_file()
        else:  # 用户点击了否
            # 直接重新加载文件，放弃当前更改
            self.reload_file()
    
    def reload_file(self):
        """重新加载当前文件"""
        if not self.current_file:
            return
        
        try:
            with open(self.current_file, 'r', encoding='utf-8') as file:
                content = file.read()
            
            # 更新文本编辑器内容
            self.text_editor.delete(1.0, tk.END)
            self.text_editor.insert(1.0, content)
            
            # 更新文件修改时间和状态
            self._file_mtime = self.get_file_mtime(self.current_file)
            self.is_modified = False
            self.update_title()
            self.update_status(f"已重新加载文件: {os.path.basename(self.current_file)}")
            logger.info(f"已重新加载文件: {self.current_file}")
        except Exception as e:
            messagebox.showerror("错误", f"无法重新加载文件: {str(e)}")
            logger.error(f"重新加载文件失败: {str(e)}")
    
    def get_file_mtime(self, file_path):
        """获取文件的修改时间"""
        try:
            return os.path.getmtime(file_path)
        except Exception:
            return None
    
    def start_file_monitor(self):
        """启动文件监控"""
        if self.current_file and not hasattr(self, '_file_monitor_running'):
            try:
                self._file_monitor_running = True
                self._file_mtime = self.get_file_mtime(self.current_file)
                self.root.after(1000, self.check_file_external_modification)
                logger.info(f"已启动文件监控: {self.current_file}")
            except Exception as e:
                logger.error(f"启动文件监控失败: {str(e)}")
                self._file_monitor_running = False
    
    def stop_file_monitor(self):
        """停止文件监控"""
        if hasattr(self, '_file_monitor_running'):
            self._file_monitor_running = False
            logger.info("已停止文件监控")
    
    def check_file_external_modification(self):
        """检查文件是否被外部修改"""
        if not hasattr(self, '_file_monitor_running') or not self._file_monitor_running or not self.current_file:
            return
        
        try:
            current_mtime = self.get_file_mtime(self.current_file)
            if current_mtime and hasattr(self, '_file_mtime') and current_mtime != self._file_mtime:
                # 文件被外部修改
                self.handle_external_modification()
            
            # 继续监控
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
        except Exception as e:
            logger.error(f"检查文件外部修改失败: {str(e)}")
            if self._file_monitor_running:
                self.root.after(1000, self.check_file_external_modification)
    
    def handle_external_modification(self):
        """处理文件被外部修改的情况"""
        if not self.current_file:
            return
        
        # 如果当前文件未被修改，直接重新加载
        if not self.is_modified:
            self.reload_file()
            return
        
        # 如果当前文件已被修改，询问用户如何处理
        response = messagebox.askyesnocancel(
            "文件已被外部修改",
            "当前文件已被外部程序修改。\n" \
            "是否保存当前更改并重新加载？\n\n" \
            "是: 保存当前更改并重新加载外部修改\n" \
            "否: 放弃当前更改并重新加载外部修改\n" \
            "取消: 保持当前状态不变"
        )
        
        if response is None:  # 用户点击了取消
            return
        elif response:  # 用户点击了是
            # 保存当前更改
            self.save_file()
            # 重新加载文件
            self.reload_file()
        else:  # 用户点击了否
            # 直接重新加载文件，放弃当前更改
            self.reload_file()
    
    def reload_file(self):
        """重新加载当前文件"""
        if not self.current_file:
            return
        
        try:
            with open(self.current_file, 'r', encoding='utf-8') as file:
                content = file.read()
            
            # 更新文本编辑器内容
            self.text_editor.delete(1.0, tk.END)
            self.text_editor.insert(1.0, content)
            
            # 更新文件修改时间和状态
            self._file_mtime = self.get_file_mtime(self.current_file)
            self.is_modified = False
            self.update_title()
            self.update_status(f"已重新加载文件: {os.path.basename(self.current_file)}")
            logger.info(f"已重新加载文件: {self.current_file}")
        except Exception as e:
            messagebox.showerror("错误", f"无法重新加载文件: {str(e)}")
            logger.error(f"重新加载文件失败: {str(e)}")

def main():
    """主函数"""
    print("进入main函数")
    try:
        print("创建Tk根窗口")
        root = tk.Tk()
        
        # 设置窗口图标
        icon_path = "cmd_sender.ico"
        if os.path.exists(icon_path):
            try:
                root.iconbitmap(icon_path)
                print(f"成功加载图标: {icon_path}")
            except Exception as e:
                print(f"加载图标失败: {e}")
        else:
            print(f"未找到图标文件: {icon_path}")
        
        print("创建CommandSenderApp")
        app = CommandSenderApp(root)
        print("启动主循环")
        root.mainloop()
        print("程序正常退出")
    except Exception as e:
        print(f"程序运行出错: {e}")
        import traceback
        traceback.print_exc()
        logger.error(f"应用程序运行出错: {e}")
        # 移除messagebox调用，避免SetForegroundWindow错误
        print(f"应用程序运行出错: {e}")

if __name__ == "__main__":
    print("开始执行main函数")
    main()
