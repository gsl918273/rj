import winreg
import re
import tkinter as tk
from tkinter import messagebox
import os
import shutil


def escape_regex(s):
    """将输入字符串转义为正则表达式安全的字符串"""
    return re.escape(s)


def check_software_installed(software_name):
    """
    检查软件是否安装，支持模糊查询，覆盖32位和64位注册表及用户级注册表。
    :param software_name: 要检查的软件名称
    :return: (bool, str) 是否安装，安装路径（可能为空）
    """

    def search_in_registry(key_path, root_key):
        try:
            print(f"正在搜索注册表路径: {key_path}")
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
                                if not install_location:  # 如果 InstallLocation 为空
                                    raise FileNotFoundError
                            except FileNotFoundError:
                                # 尝试使用 UninstallString
                                uninstall_string = winreg.QueryValueEx(subkey, "UninstallString")[0]
                                install_location = parse_install_path_from_uninstall(uninstall_string)
                            print(f"找到匹配的软件: {display_name}, 安装路径: {install_location}")
                            return True, install_location
                    except FileNotFoundError:
                        continue
                except OSError:
                    break
        except OSError as e:
            print(f"未找到注册表路径 {key_path}，错误信息: {e}")
            return False, "未找到安装路径"
        return False, "未找到安装路径"

    def parse_install_path_from_uninstall(uninstall_string):
        """尝试从卸载字符串中提取安装路径"""
        if uninstall_string and os.path.exists(uninstall_string):
            return os.path.dirname(uninstall_string)
        # 如果卸载字符串是命令路径，提取目录部分
        match = re.search(r"\"(.*?)\"", uninstall_string)
        if match:
            potential_path = match.group(1)
            if os.path.exists(potential_path):
                return os.path.dirname(potential_path)
        return "路径信息不可用"

    # 搜索注册表路径，包括系统和用户级安装
    registry_paths = [
        (r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", winreg.HKEY_LOCAL_MACHINE),
        (r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", winreg.HKEY_LOCAL_MACHINE),
        (r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", winreg.HKEY_CURRENT_USER),
        (r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", winreg.HKEY_CURRENT_USER),
    ]

    for path, root_key in registry_paths:
        installed, location = search_in_registry(path, root_key)
        if installed:
            return True, location

    print(f"未找到与 \"{software_name}\" 匹配的软件")
    return False, "未找到安装路径"


def search_in_filesystem(software_name):
    """在文件系统中查找软件"""
    potential_paths = [
        os.environ.get("PROGRAMFILES", ""),  # 默认安装路径
        os.environ.get("PROGRAMFILES(X86)", ""),  # 32位安装路径
        os.environ.get("LOCALAPPDATA", ""),  # 用户本地应用路径
        os.environ.get("APPDATA", ""),  # 用户漫游应用路径
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
    try:
        print(f"开始删除软件: {software_name}")
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
                        if software_name.lower() in display_name.lower():
                            winreg.DeleteKey(key, subkey_name)
                            print(f"成功删除注册表项: {subkey_name}")
                            break
                    except OSError:
                        break
                winreg.CloseKey(key)
            except OSError as e:
                print(f"无法删除注册表路径 {key_path}, 错误: {e}")
                continue

        if install_path and os.path.exists(install_path):
            shutil.rmtree(install_path)
            print(f"成功删除安装路径: {install_path}")
    except Exception as e:
        print(f"删除软件 \"{software_name}\" 时出现错误: {e}")
        raise e


def delete_all_installed_software():
    """一键删除所有指定的软件"""
    software_names = text_area.get("1.0", "end").splitlines()
    for software_name in software_names:
        if software_name.strip():
            installed, install_path = check_software_installed(software_name.strip())
            if not installed:
                # 如果未在注册表中找到，尝试文件系统搜索
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
    result_label.config(text="")  # 清空结果
    results = ""
    for software_name in software_names:
        if software_name.strip():
            installed, install_path = check_software_installed(software_name.strip())
            if not installed:
                # 如果未在注册表中找到，尝试文件系统搜索
                installed, install_path = search_in_filesystem(software_name.strip())
            status = "已安装" if installed else "未安装"
            results += f"{software_name}: {status}"
            if installed:
                results += f"，安装路径：{install_path}"
            results += "\n"
    result_label.config(text=results)
    print("查询结果如下:")
    print(results)


# GUI 构建
root = tk.Tk()
root.title("软件安装查询工具")
root.geometry("800x600")

# 用户手动输入软件名称
text_area = tk.Text(root, height=10, width=80)
text_area.pack(pady=10)

result_label = tk.Label(root, text="", wraplength=700, height=20, anchor="nw", justify="left", bg="#f0f0f0")
result_label.pack(pady=10)

search_button = tk.Button(root, text="开始查询", command=start_search)
search_button.pack(pady=10)

delete_all_button = tk.Button(root, text="一键删除", command=delete_all_installed_software)
delete_all_button.pack(pady=10)

root.mainloop()