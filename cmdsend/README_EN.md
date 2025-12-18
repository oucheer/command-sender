# Command Sender

<div align="center">
  <p>📤 An efficient tool for sending commands to various terminal windows</p>
  <p>
    <a href="https://github.com/">
      <img src="https://img.shields.io/badge/GitHub-Open%20Source-blue?style=flat-square" alt="GitHub">
    </a>
    <a href="LICENSE">
      <img src="https://img.shields.io/badge/LICENSE-MIT-green?style=flat-square" alt="License">
    </a>
    <a href="requirements.txt">
      <img src="https://img.shields.io/badge/Python-3.8%2B-blue?style=flat-square" alt="Python Version">
    </a>
  </p>
</div>

## 📝 Project Introduction

Command Sender is a powerful tool for sending commands to various terminal windows. It sends commands accurately to specified terminal windows by simulating keyboard input, supporting multiple terminal types, and helping users improve work efficiency.

## ✨ Main Features

- 🎯 **Multiple Terminal Support**: Compatible with PowerShell, MobaXterm, SecureCRT, Xshell, PuTTY, Windows Terminal, etc.
- ⌨️ **Simulated Keyboard Input**: Reliable keyboard event simulation via Windows API
- 📋 **Clipboard Integration**: Support for sending commands via clipboard
- 🔌 **Serial Communication**: Support for sending commands via serial port
- 🎨 **User-Friendly GUI**: Intuitive user interface built with Tkinter
- ⚡ **Efficient Command Sending**: Support for single line sending, selected text sending, and full content sending
- 📦 **Simple Compilation Process**: Quick compilation to executable file using PyInstaller
- 🎯 **Intelligent Terminal Recognition**: Automatically recognizes terminal types and adopts optimal sending strategies
- 🔍 **Reliable Focus Management**: Ensures commands are sent to the correct window

## 🚀 Quick Start

### Environment Requirements

- Python 3.8 or higher
- Windows operating system

### Install Dependencies

```bash
pip install -r requirements.txt
```

### Run the Program

```bash
python complete_command_sender.py
```

### Basic Usage

1. Launch the application
2. Click the "拖拽选择" (Drag Select) button to select the target terminal window
3. Enter the commands to be sent in the text editor
4. Click the send button or use keyboard shortcuts to send commands

## 🔧 Compilation Guide

### Compilation Environment Dependencies

- Python 3.8+ 
- PyInstaller 6.0+ 

### Compilation Steps

1. **Install PyInstaller**:
   ```bash
   pip install pyinstaller
   ```

2. **Compile the executable file** (Choose either method):
   
   **Method 1: Using spec file (Recommended)**
   ```bash
   pyinstaller cmd_sender.spec
   ```
   
   **Method 2: Using command line parameters directly**
   ```bash
   python -m PyInstaller --onefile --windowed --name "cmd_sender" complete_command_sender.py
   ```

3. **Get compilation results**:
   - After successful compilation, the executable file will be generated in the `dist` directory
   - Generated executable file: `dist/cmd_sender.exe`

## 📖 Usage Instructions

### 1. Select Target Terminal

- Click the "拖拽选择" (Drag Select) button
- Move the mouse to the target terminal window and click
- The application will automatically identify the terminal type

### 2. Write Commands

- Enter the commands to be sent in the text editor
- Support for multiple lines of commands, one command per line
- Commands can be loaded from files (via the "File" menu)

### 3. Send Commands

- **Send current line**: Click the send button on the left side of the current line
- **Send selected text**: Select the text to be sent, click the "发送选中文本" (Send Selected Text) button in the toolbar
- **Send all content**: Click the "发送全部内容" (Send All Content) button in the toolbar

### 4. Select Sending Mode

- **Clipboard**: Only copy commands to clipboard
- **Terminal Input**: Send commands to terminal
- **Serial**: Send commands to serial device

### 5. Auto-Enter Setting

- Check the "自动换行执行" (Auto Enter) option, commands will be executed automatically after sending
- Uncheck it if you need to press Enter manually to execute commands

## 🔧 Technical Architecture

| Technical Point | Description | Library/Tool |
|----------------|-------------|--------------|
| Python GUI Development | Build graphical interface | Tkinter |
| Windows API Calls | Window management and message sending | pywin32 |
| Keyboard Simulation | Simulate keyboard events | pyautogui, keyboard |
| Clipboard Operations | Manage clipboard content | pyperclip |
| Serial Communication | Serial command sending | serial |
| Process Management | Get process information | psutil |
| Terminal Type Recognition | Identify terminal types | Custom algorithm |
| Focus Management | Get window focus | Custom algorithm |

## 🤝 Contribution Guide

Welcome to participate in the development of the project! Here is the contribution guide:

### Contribution Process

1. Fork this repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

### Code Guidelines

- Follow PEP 8 code style guidelines
- Keep code concise and clear
- Add appropriate comments
- Ensure code can run normally

### Commit Message Guidelines

- Write commit messages in English
- Briefly describe the changes (no more than 50 characters)
- For complex changes, add detailed descriptions

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details

## 📞 Support and Contact

If you encounter any problems during use, or have any suggestions, please contact us through the following ways:

- Submit an [Issue](https://github.com/issues) to report problems
- Submit a [Pull Request](https://github.com/pulls) to contribute code

## 🙏 Acknowledgments

Thanks to all developers and users who have contributed to the project!

---

<div align="center">
  <p>Made with ❤️ by Command Sender Team</p>
</div>