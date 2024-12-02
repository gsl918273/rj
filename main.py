import winreg
import re
import tkinter as tk
from tkinter import messagebox
import os
import shutil
import stat


def escape_regex(s):
    """将输入字符串转义为正则表达式安全的字符串"""
    return re.escape(s)


def check_software_installed(software_name):
    """检查软件是否安装"""
    def search_in_registry(key_path, root_key):
        try:
            key = winreg.OpenKey(root_key, key_path)
            subkey_count = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(key, subkey_count)
                    subkey_count += 1
                    subkey = winreg.OpenKey(key, subkey_name)
                    try:
                        display_name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                        if software_name.lower() in display_name.lower():  # 宽松匹配
                            try:
                                install_location = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                                if not install_location:
                                    raise FileNotFoundError
                            except FileNotFoundError:
                                uninstall_string = winreg.QueryValueEx(subkey, "UninstallString")[0]
                                install_location = parse_install_path_from_uninstall(uninstall_string)
                            return True, install_location
                    except FileNotFoundError:
                        continue
                except OSError:
                    break
        except OSError:
            return False, "未找到安装路径"
        return False, "未找到安装路径"

    def parse_install_path_from_uninstall(uninstall_string):
        """尝试从卸载字符串中提取安装路径"""
        match = re.search(r"\"(.*?)\"", uninstall_string)
        if match:
            potential_path = match.group(1)
            if os.path.exists(potential_path):
                return os.path.dirname(potential_path)
        return "路径信息不可用"

    registry_paths = [
        (r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", winreg.HKEY_LOCAL_MACHINE),
        (r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", winreg.HKEY_LOCAL_MACHINE),
        (r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", winreg.HKEY_CURRENT_USER),
    ]

    for path, root_key in registry_paths:
        installed, location = search_in_registry(path, root_key)
        if installed:
            return True, location
    return False, "未找到安装路径"


def search_in_filesystem(software_name):
    """在文件系统中查找软件"""
    potential_paths = [
        os.environ.get("PROGRAMFILES", ""),
        os.environ.get("PROGRAMFILES(X86)", ""),
        os.environ.get("LOCALAPPDATA", ""),
    ]
    for path in potential_paths:
        if path and os.path.exists(path):
            for root, dirs, files in os.walk(path):
                for file in files:
                    if software_name.lower() in file.lower():
                        return True, root
    return False, "未找到文件路径"


def delete_software(software_name, install_path):
    """删除软件及其安装路径"""
    def make_writable(func, path, _):
        os.chmod(path, stat.S_IWRITE)
        func(path)

    try:
        if install_path and os.path.exists(install_path):
            shutil.rmtree(install_path, onerror=make_writable)
    except Exception as e:
        raise e


def delete_all_installed_software():
    """一键删除所有指定的软件"""
    software_names = text_area.get("1.0", "end").splitlines()
    for software_name in software_names:
        if software_name.strip():
            installed, install_path = check_software_installed(software_name.strip())
            if not installed:
                installed, install_path = search_in_filesystem(software_name.strip())
            if installed:
                try:
                    delete_software(software_name.strip(), install_path)
                except Exception as e:
                    messagebox.showerror("错误", f"删除 {software_name} 时出现错误：{str(e)}")
            else:
                print(f"软件 \"{software_name}\" 未安装，跳过删除")
    messagebox.showinfo("提示", "所有已安装软件已处理完成。")


def start_search():
    """开始查询软件是否安装"""
    software_names = text_area.get("1.0", "end").splitlines()
    result_label.config(text="")
    results = ""
    for software_name in software_names:
        if software_name.strip():
            installed, install_path = check_software_installed(software_name.strip())
            if not installed:
                installed, install_path = search_in_filesystem(software_name.strip())
            status = "已安装" if installed else "未安装"
            results += f"{software_name}: {status}"
            if installed:
                results += f"，安装路径：{install_path}"
            results += "\n"
    result_label.config(text=results)


# GUI 构建
root = tk.Tk()
root.title("软件安装查询工具")
root.geometry("800x700")
root.resizable(True, True)

# 使用滚动条和框架
frame = tk.Frame(root)
frame.pack(fill=tk.BOTH, expand=True)

scrollbar = tk.Scrollbar(frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

canvas = tk.Canvas(frame, yscrollcommand=scrollbar.set)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar.config(command=canvas.yview)

inner_frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=inner_frame, anchor="nw")

text_area = tk.Text(inner_frame, height=10, width=80)
text_area.pack(pady=10)

result_label = tk.Label(inner_frame, text="", wraplength=700, height=20, anchor="nw", justify="left", bg="#f0f0f0")
result_label.pack(pady=10)

search_button = tk.Button(inner_frame, text="开始查询", command=start_search)
search_button.pack(pady=10)

delete_all_button = tk.Button(inner_frame, text="一键删除", command=delete_all_installed_software)
delete_all_button.pack(pady=10)

# 更新滚动区域
inner_frame.update_idletasks()
canvas.config(scrollregion=canvas.bbox("all"))

root.mainloop()