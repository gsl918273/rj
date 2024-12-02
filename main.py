import winreg
import re
import tkinter as tk
from tkinter import messagebox
import time
import os
import shutil


def escape_regex(s):
    """将输入字符串转义为正则表达式安全的字符串"""
    return re.escape(s)


def check_software_installed(software_name):
    """
    检查软件是否安装，支持模糊查询，覆盖32位和64位注册表。
    :param software_name: 要检查的软件名称
    :return: (bool, str) 是否安装，安装路径（可能为空）
    """
    def search_in_registry(key_path):
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)
            subkey_count = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(key, subkey_count)
                    subkey_count += 1
                    subkey = winreg.OpenKey(key, subkey_name)
                    try:
                        display_name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                        if re.search(escape_regex(software_name), display_name, re.IGNORECASE):
                            try:
                                install_location = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                            except FileNotFoundError:
                                install_location = "路径信息不可用"
                            return True, install_location
                    except FileNotFoundError:
                        continue
                except OSError:
                    break
        except OSError:
            return False, "未找到安装路径"
        return False, "未找到安装路径"

    registry_paths = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    ]
    for path in registry_paths:
        installed, location = search_in_registry(path)
        if installed:
            return True, location

    return False, "未找到安装路径"


def delete_software(software_name, install_path):
    """删除软件及其安装路径"""
    try:
        registry_paths = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        ]
        for key_path in registry_paths:
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_ALL_ACCESS)
                subkey_count = 0
                while True:
                    try:
                        subkey_name = winreg.EnumKey(key, subkey_count)
                        subkey_count += 1
                        subkey = winreg.OpenKey(key, subkey_name)
                        display_name, _ = winreg.QueryValueEx(subkey, "DisplayName")
                        if re.search(escape_regex(software_name), display_name, re.IGNORECASE):
                            winreg.DeleteKey(key, subkey_name)
                            break
                    except OSError:
                        break
                winreg.CloseKey(key)
            except OSError:
                continue

        if install_path and os.path.exists(install_path):
            shutil.rmtree(install_path)
    except Exception as e:
        raise e


def delete_all_installed_software():
    """一键删除所有指定的软件"""
    software_names = text_area.get("1.0", "end").splitlines()
    for software_name in software_names:
        if software_name.strip():
            installed, install_path = check_software_installed(software_name.strip())
            if installed:
                try:
                    delete_software(software_name.strip(), install_path)
                except Exception as e:
                    messagebox.showerror("错误", f"删除 {software_name} 时出现错误：{str(e)}")
    messagebox.showinfo("提示", "所有已安装软件已删除。")


def start_search():
    """开始查询软件是否安装"""
    software_names = text_area.get("1.0", "end").splitlines()
    waiting_label = tk.Label(root, text="正在查询...")
    waiting_label.pack(pady=20)
    root.update()
    time.sleep(2)
    waiting_label.destroy()
    result_label.config(text="")
    results = ""
    for software_name in software_names:
        if software_name.strip():
            installed, install_path = check_software_installed(software_name.strip())
            status = "已安装" if installed else "未安装"
            results += f"{software_name}: {status}"
            if installed:
                results += f"，安装路径：{install_path}"
            results += "\n"
    result_label.config(text=results)


# GUI 构建
root = tk.Tk()
root.title("软件安装查询工具")
root.geometry("800x800")

# 用户手动输入软件名称
text_area = tk.Text(root, height=15, width=100)
text_area.pack(pady=10)

result_label = tk.Label(root, text="", wraplength=550, height=15, width=100, anchor="w", justify="left")
result_label.pack(pady=10)

search_button = tk.Button(root, text="开始查询", command=start_search)
search_button.pack(pady=10)

delete_all_button = tk.Button(root, text="一键删除", command=delete_all_installed_software)
delete_all_button.pack(pady=10)

root.mainloop()