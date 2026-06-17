"""
大圣.快递物流出港帐单结算系统 - 卸载程序
"""
import os
import sys
import shutil
import ctypes
import tkinter as tk
from tkinter import messagebox

APP_NAME = "大圣快递帐单结算系统"
UNINSTALL_FILENAME = "uninstall.txt"


def find_install_dir():
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    candidate = os.path.join(base, UNINSTALL_FILENAME)
    if os.path.exists(candidate):
        return base
    # 往上找
    for _ in range(3):
        parent = os.path.dirname(base)
        if parent == base:
            break
        base = parent
        if os.path.exists(os.path.join(base, UNINSTALL_FILENAME)):
            return base
    return None


def delete_shortcuts():
    # 桌面
    desktop_candidates = [
        os.path.join(os.path.expanduser("~"), "Desktop"),
        os.path.join(os.path.expanduser("~"), "桌面"),
    ]
    public = os.environ.get("PUBLIC")
    if public:
        desktop_candidates.append(os.path.join(public, "Desktop"))
    for desk in desktop_candidates:
        if os.path.isdir(desk):
            for f in os.listdir(desk):
                if APP_NAME in f and f.endswith((".lnk", ".bat")):
                    try:
                        os.remove(os.path.join(desk, f))
                    except Exception:
                        pass
    # 开始菜单
    start_menus = [
        os.path.join(os.environ.get("APPDATA", ""),
                     "Microsoft", "Windows", "Start Menu", "Programs", APP_NAME),
        os.path.join(os.environ.get("ProgramData", ""),
                     "Microsoft", "Windows", "Start Menu", "Programs", APP_NAME),
    ]
    for sm in start_menus:
        if os.path.isdir(sm):
            try:
                shutil.rmtree(sm, ignore_errors=True)
            except Exception:
                pass


class UninstallApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME + " - 卸载程序")
        self.geometry("520x260")
        self.resizable(False, False)
        self.configure(bg="#ffffff")
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

        self.install_dir = find_install_dir()

        # 头部蓝色条
        header = tk.Frame(self, bg="#c0392b", height=70)
        header.pack(fill=tk.X)
        tk.Label(header, text="卸载 " + APP_NAME,
                 bg="#c0392b", fg="white",
                 font=("Microsoft YaHei", 15, "bold")).pack(anchor=tk.W, padx=28, pady=(18, 0))
        tk.Label(header, text="从此计算机中移除此软件",
                 bg="#c0392b", fg="#ffd6cc",
                 font=("Microsoft YaHei", 9)).pack(anchor=tk.W, padx=28, pady=(2, 0))

        # 内容
        body = tk.Frame(self, bg="#ffffff")
        body.pack(fill=tk.BOTH, expand=True, padx=28, pady=16)

        if not self.install_dir:
            tk.Label(body, text="⚠ 找不到安装信息。",
                     bg="#ffffff", fg="red", font=("Microsoft YaHei", 12, "bold")).pack(anchor=tk.W, pady=8)
            tk.Label(body, text="请在 " + APP_NAME + " 的安装目录中运行此程序。",
                     bg="#ffffff", fg="#333333", font=("Microsoft YaHei", 10)).pack(anchor=tk.W, pady=2)
        else:
            tk.Label(body, text="\n确定要从您的计算机中完全移除此软件吗？",
                     bg="#ffffff", fg="#333333", font=("Microsoft YaHei", 11)).pack(anchor=tk.W, pady=(0, 4))
            tk.Label(body, text="安装位置：" + self.install_dir,
                     bg="#ffffff", fg="#666666", font=("Microsoft YaHei", 9)).pack(anchor=tk.W, pady=4)
            tk.Label(body,
                     text="\n此操作将删除：\n   • 程序主体文件\n   • 桌面和开始菜单的快捷方式\n   • 用户配置信息",
                     bg="#ffffff", fg="#555555", font=("Microsoft YaHei", 10),
                     justify=tk.LEFT).pack(anchor=tk.W, pady=(8, 0))

        # 底部按钮栏
        bottom = tk.Frame(self, bg="#f4f4f4", height=56)
        bottom.pack(side=tk.BOTTOM, fill=tk.X)
        tk.Frame(self, bg="#e0e0e0", height=1).pack(side=tk.BOTTOM, fill=tk.X)

        if self.install_dir:
            tk.Button(bottom, text="取消", width=10,
                      font=("Microsoft YaHei", 10),
                      command=self.destroy).pack(side=tk.RIGHT, padx=(8, 24), pady=12)
            tk.Button(bottom, text="卸载", width=10,
                      font=("Microsoft YaHei", 10, "bold"),
                      bg="#c0392b", fg="white",
                      activebackground="#922b21", activeforeground="white",
                      command=self.do_uninstall).pack(side=tk.RIGHT, padx=8, pady=12)
        else:
            tk.Button(bottom, text="关闭", width=10,
                      font=("Microsoft YaHei", 10, "bold"),
                      bg="#2b78e4", fg="white",
                      command=self.destroy).pack(side=tk.RIGHT, padx=24, pady=12)

    def do_uninstall(self):
        if not messagebox.askyesno("确认卸载",
                                   '真的要卸载 "' + APP_NAME + '" 吗？\n\n此操作无法撤销。',
                                   parent=self):
            return
        try:
            delete_shortcuts()
            target = self.install_dir
            if target:
                # 清空安装目录内容（不删除目录本身，防止删除自身 exe 时出错）
                for item in os.listdir(target):
                    path = os.path.join(target, item)
                    try:
                        if os.path.isdir(path):
                            shutil.rmtree(path, ignore_errors=True)
                        else:
                            try:
                                os.remove(path)
                            except Exception:
                                pass
                    except Exception:
                        pass
            messagebox.showinfo("卸载完成",
                                APP_NAME + " 已成功从您的计算机中移除。\n\n感谢您的使用！",
                                parent=self)
        except Exception as e:
            messagebox.showerror("卸载失败", str(e), parent=self)
        self.destroy()


def main():
    app = UninstallApp()
    app.mainloop()


if __name__ == "__main__":
    main()
