import winreg
import re
import tkinter as tk
from tkinter import messagebox
import os
import subprocess


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
                                uninstall_string = winreg.QueryValueEx(subkey, "UninstallString")[0]
                                return True, uninstall_string
                            except FileNotFoundError:
                                return True, None
                    except FileNotFoundError:
                        continue
                except OSError:
                    break
        except OSError:
            return False, None
        return False, None

    registry_paths = [
        (r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", winreg.HKEY_LOCAL_MACHINE),
        (r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", winreg.HKEY_LOCAL_MACHINE),
        (r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", winreg.HKEY_CURRENT_USER),
    ]

    for path, root_key in registry_paths:
        installed, uninstall_cmd = search_in_registry(path, root_key)
        if installed:
            return True, uninstall_cmd
    return False, None


def run_uninstall_command(uninstall_cmd):
    """运行卸载命令"""
    try:
        print(f"执行卸载命令: {uninstall_cmd}")
        subprocess.run(uninstall_cmd, shell=True, check=True)
        return True
    except subprocess.CalledProcessError as e:
        print(f"卸载失败: {e}")
        return False


def delete_all_installed_software():
    """一键删除所有指定的软件"""
    software_names = text_area.get("1.0", "end").splitlines()
    for software_name in software_names:
        if software_name.strip():
            installed, uninstall_cmd = check_software_installed(software_name.strip())
            if installed and uninstall_cmd:
                if messagebox.askyesno("确认卸载", f"确定要卸载 {software_name.strip()} 吗？"):
                    success = run_uninstall_command(uninstall_cmd)
                    if success:
                        messagebox.showinfo("卸载成功", f"{software_name.strip()} 已成功卸载。")
                    else:
                        messagebox.showerror("卸载失败", f"{software_name.strip()} 卸载失败。")
            else:
                print(f"软件 \"{software_name}\" 未安装或无法卸载，跳过处理。")
    messagebox.showinfo("提示", "所有指定的软件处理完成。")


def start_search():
    """开始查询软件是否安装"""
    software_names = text_area.get("1.0", "end").splitlines()
    result_label.config(text="")
    results = ""
    for software_name in software_names:
        if software_name.strip():
            installed, uninstall_cmd = check_software_installed(software_name.strip())
            status = "已安装" if installed else "未安装"
            results += f"{software_name}: {status}"
            if installed:
                results += f"，卸载命令：{uninstall_cmd}" if uninstall_cmd else "（无卸载命令）"
            results += "\n"
    result_label.config(text=results)


# GUI 构建
root = tk.Tk()
root.title("软件安装查询工具")
root.geometry("800x700")
root.resizable(True, True)

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

delete_all_button = tk.Button(inner_frame, text="一键卸载", command=delete_all_installed_software)
delete_all_button.pack(pady=10)

# 更新滚动区域
inner_frame.update_idletasks()
canvas.config(scrollregion=canvas.bbox("all"))

root.mainloop()