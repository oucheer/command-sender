# 命令发送器 (Command Sender)

<div align="center">
  <p>📤 向多种终端发送命令的高效工具</p>
  <p>
    <a href="https://github.com/">
      <img src="https://img.shields.io/badge/GitHub-开源项目-blue?style=flat-square" alt="GitHub">
    </a>
    <a href="LICENSE">
      <img src="https://img.shields.io/badge/LICENSE-MIT-green?style=flat-square" alt="License">
    </a>
    <a href="requirements.txt">
      <img src="https://img.shields.io/badge/Python-3.8%2B-blue?style=flat-square" alt="Python Version">
    </a>
  </p>
</div>

## 📝 项目介绍

命令发送器是一个功能强大的工具，用于向各种终端窗口发送命令。它通过模拟键盘输入的方式，将命令准确地发送到指定的终端窗口，支持多种终端类型，帮助用户提高工作效率。

## ✨ 主要功能

- 🎯 **多种终端支持**：兼容PowerShell、MobaXterm、SecureCRT、Xshell、PuTTY、Windows Terminal等
- ⌨️ **模拟键盘输入**：通过Windows API实现可靠的键盘事件模拟
- 📋 **剪贴板集成**：支持剪贴板方式发送命令
- 🔌 **串口通信**：支持通过串口发送命令
- 🎨 **友好的GUI界面**：基于Tkinter构建的直观用户界面
- ⚡ **高效的命令发送**：支持单行发送、选中文本发送和全部内容发送
- 📦 **简单的编译过程**：使用PyInstaller快速编译为可执行文件
- 🎯 **智能终端识别**：自动识别终端类型并采用最佳发送策略
- 🔍 **可靠的焦点管理**：确保命令发送到正确的窗口

## 🚀 快速开始

### 环境要求

- Python 3.8 或更高版本
- Windows 操作系统

### 安装依赖

```bash
pip install -r requirements.txt
```

### 运行程序

```bash
python complete_command_sender.py
```

### 基本使用

1. 启动应用程序
2. 点击「拖拽选择」按钮选择目标终端窗口
3. 在文本编辑器中输入要发送的命令
4. 点击发送按钮或使用快捷键发送命令

## 🔧 编译指南

### 编译环境依赖

- Python 3.8+ 
- PyInstaller 6.0+ 

### 编译步骤

1. **安装PyInstaller**：
   ```bash
   pip install pyinstaller
   ```

2. **编译可执行文件**（两种方式任选其一）：
   
   **方式一：使用spec文件（推荐）**
   ```bash
   pyinstaller cmd_sender.spec
   ```
   
   **方式二：直接使用命令行参数**
   ```bash
   python -m PyInstaller --onefile --windowed --name "cmd_sender" complete_command_sender.py
   ```

3. **获取编译结果**：
   - 编译成功后，可执行文件将生成在 `dist` 目录下
   - 生成的可执行文件：`dist/cmd_sender.exe`

## 📖 使用指导

### 1. 选择目标终端

- 点击「拖拽选择」按钮
- 将鼠标移动到目标终端窗口上并点击
- 应用程序会自动识别终端类型

### 2. 编写命令

- 在文本编辑器中输入要发送的命令
- 支持多行命令，每行一个命令
- 可以从文件中加载命令（通过「文件」菜单）

### 3. 发送命令

- **发送当前行**：点击当前行左侧的发送按钮
- **发送选中文本**：选中要发送的文本，点击工具栏中的「发送选中文本」按钮
- **发送全部内容**：点击工具栏中的「发送全部内容」按钮

### 4. 发送方式选择

- **剪贴板**：只复制命令到剪贴板
- **终端输入**：发送命令到终端
- **串口**：发送命令到串口设备

### 5. 自动换行设置

- 勾选「自动换行执行」选项，命令发送后会自动执行
- 取消勾选则需要手动按Enter键执行命令

## 🔧 技术架构

| 技术点 | 描述 | 库/工具 |
|-------|------|--------|
| Python GUI开发 | 构建图形界面 | Tkinter |
| Windows API调用 | 窗口管理和消息发送 | pywin32 |
| 键盘模拟 | 模拟键盘事件 | pyautogui, keyboard |
| 剪贴板操作 | 管理剪贴板内容 | pyperclip |
| 串口通信 | 串口命令发送 | serial |
| 进程管理 | 获取进程信息 | psutil |
| 终端类型识别 | 识别终端类型 | 自定义算法 |
| 焦点管理 | 获取窗口焦点 | 自定义算法 |

## 🤝 贡献指南

欢迎您参与到项目的开发中来！以下是贡献指南：

### 贡献流程

1. Fork 本仓库
2. 创建您的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启一个 Pull Request

### 代码规范

- 遵循 PEP 8 代码风格规范
- 保持代码简洁明了
- 添加适当的注释
- 确保代码可以正常运行

### 提交信息规范

- 使用英文撰写提交信息
- 简短描述更改内容（不超过50个字符）
- 对于复杂更改，添加详细描述

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

## 📞 支持与联系方式

如果您在使用过程中遇到问题，或者有任何建议，欢迎通过以下方式联系我们：

- 提交 [Issue](https://github.com/issues) 反馈问题
- 提交 [Pull Request](https://github.com/pulls) 贡献代码

## 🙏 致谢

感谢所有为项目做出贡献的开发者和用户！

---

<div align="center">
  <p>Made with ❤️ by Command Sender Team</p>
</div>