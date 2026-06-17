"""
大圣.快递物流出港帐单结算系统 - 安装程序
多步骤向导：欢迎 → 选择目录 → 确认 → 安装中 → 完成
"""
import os
import sys
import shutil
import ctypes
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_NAME = "大圣快递帐单结算系统"
APP_DIR_NAME = "SFeeSystem"
VERSION = "1.0"


def get_source_dir():
    """获取程序源文件目录（setup.exe 同目录下的 SFeeSystem/）"""
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    source = os.path.join(base, APP_DIR_NAME)
    if not os.path.isdir(source):
        parent = os.path.join(os.path.dirname(base), APP_DIR_NAME)
        if os.path.isdir(parent):
            return parent
    return source


def default_install_dir():
    """推荐安装目录：优先 Program Files，否则 LocalAppData"""
    program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
    candidate = os.path.join(program_files, APP_NAME)
    try:
        test_path = os.path.join(program_files, "_sf_test_" + str(os.getpid()))
        os.makedirs(test_path, exist_ok=True)
        os.rmdir(test_path)
        return candidate
    except Exception:
        local = os.environ.get("LOCALAPPDATA", os.path.expanduser("~"))
        return os.path.join(local, "Programs", APP_NAME)


def create_shortcut(target_path, shortcut_dir, shortcut_name):
    """创建 .lnk 快捷方式"""
    try:
        from win32com.client import Dispatch
        shell = Dispatch("WScript.Shell")
        lnk_path = os.path.join(shortcut_dir, shortcut_name + ".lnk")
        shortcut = shell.CreateShortCut(lnk_path)
        shortcut.Targetpath = target_path
        shortcut.WorkingDirectory = os.path.dirname(target_path)
        shortcut.IconLocation = target_path
        shortcut.WindowStyle = 1
        shortcut.Description = APP_NAME
        shortcut.Save()
        return lnk_path
    except Exception:
        try:
            # 回落方案：生成 .bat 启动脚本
            bat_path = os.path.join(shortcut_dir, shortcut_name + ".bat")
            with open(bat_path, "w", encoding="utf-8") as f:
                f.write('@echo off\r\nstart "" "%s"\r\n' % target_path)
            return bat_path
        except Exception:
            return None


class InstallerWizard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME + " " + VERSION + " - 安装向导")
        self.geometry("640x520")
        self.resizable(False, False)
        self.configure(bg="#ffffff")

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

        # 数据
        self.source_dir = get_source_dir()
        self.install_dir = tk.StringVar(value=default_install_dir())
        self.desktop_shortcut = tk.BooleanVar(value=True)
        self.start_menu_shortcut = tk.BooleanVar(value=True)
        self.auto_run = tk.BooleanVar(value=True)
        self.source_ok = os.path.isdir(self.source_dir) and os.path.exists(
            os.path.join(self.source_dir, APP_DIR_NAME + ".exe")
        ) or os.path.exists(os.path.join(self.source_dir, "SFeeSystem.exe"))

        # 底部按钮栏先打包（确保始终可见）
        self.bottom_bar = tk.Frame(self, bg="#f4f4f4", height=64)
        self.bottom_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 分隔线
        tk.Frame(self, bg="#e0e0e0", height=1).pack(side=tk.BOTTOM, fill=tk.X)

        # 主容器后打包（填充剩余空间）
        self.content_frame = tk.Frame(self, bg="#ffffff")
        self.content_frame.pack(fill=tk.BOTH, expand=True)

        self.btn_back = tk.Button(
            self.bottom_bar, text="上一步", width=10, font=("Microsoft YaHei", 10),
            command=self.go_back, state=tk.DISABLED
        )
        self.btn_back.pack(side=tk.LEFT, padx=(24, 8), pady=12)

        self.btn_next = tk.Button(
            self.bottom_bar, text="下一步 >", width=12, font=("Microsoft YaHei", 10, "bold"),
            bg="#2b78e4", fg="white", activebackground="#1e5fa8", activeforeground="white",
            command=self.go_next
        )
        self.btn_next.pack(side=tk.LEFT, padx=8, pady=12)

        self.btn_cancel = tk.Button(
            self.bottom_bar, text="取消", width=10, font=("Microsoft YaHei", 10),
            command=self.quit_app
        )
        self.btn_cancel.pack(side=tk.RIGHT, padx=(8, 24), pady=12)

        # 页面管理
        self.current_step = 1
        self.total_steps = 4  # 欢迎、选择目录、安装中、完成
        self.show_step_1()

    def clear_content(self):
        for w in self.content_frame.winfo_children():
            w.destroy()

    def show_header(self, title, subtitle=""):
        header = tk.Frame(self.content_frame, bg="#1e5fa8", height=80)
        header.pack(fill=tk.X)
        tk.Label(
            header, text=title, bg="#1e5fa8", fg="white",
            font=("Microsoft YaHei", 16, "bold")
        ).pack(anchor=tk.W, padx=28, pady=(18, 0))
        if subtitle:
            tk.Label(
                header, text=subtitle, bg="#1e5fa8", fg="#cce3ff",
                font=("Microsoft YaHei", 9)
            ).pack(anchor=tk.W, padx=28, pady=(2, 0))
        tk.Frame(self.content_frame, bg="#ffffff").pack(fill=tk.BOTH, expand=True, pady=(0, 0))

    # ============ 步骤1: 欢迎 ============
    def show_step_1(self):
        self.clear_content()
        self.current_step = 1
        self.btn_back.config(state=tk.DISABLED)
        self.btn_next.config(text="下一步 >", state=tk.NORMAL, bg="#2b78e4", fg="white")

        self.show_header("欢迎使用 " + APP_NAME + " 安装向导",
                         "版本 " + VERSION)

        body = tk.Frame(self.content_frame, bg="#ffffff")
        body.pack(fill=tk.BOTH, expand=True, padx=28, pady=20)

        if not self.source_ok:
            tk.Label(body, text="⚠ 错误：找不到程序安装源文件！",
                     bg="#ffffff", fg="red", font=("Microsoft YaHei", 12, "bold")).pack(anchor=tk.W, pady=8)
            tk.Label(body, text="安装程序需要与 " + APP_DIR_NAME + " 文件夹放在同一目录。",
                     bg="#ffffff", fg="#333333", font=("Microsoft YaHei", 10), justify=tk.LEFT, wraplength=560).pack(anchor=tk.W, pady=4)
            self.btn_next.config(state=tk.DISABLED, bg="#cccccc")
            return

        info = (
            "本向导将在您的电脑上安装 " + APP_NAME + "。\n\n"
            "安装前建议关闭其他正在运行的同类应用程序。\n\n"
            "点击「下一步」继续安装，或点击「取消」退出安装向导。"
        )
        tk.Label(body, text=info, bg="#ffffff", fg="#333333",
                 font=("Microsoft YaHei", 10), justify=tk.LEFT, wraplength=560).pack(anchor=tk.W, pady=(10, 0))

        tk.Label(body, text="\n占用空间：约 150 MB\n支持系统：Windows 7 / 8 / 10 / 11",
                 bg="#ffffff", fg="#666666", font=("Microsoft YaHei", 9), justify=tk.LEFT).pack(anchor=tk.W, pady=(10, 0))

    # ============ 步骤2: 选择目录 ============
    def show_step_2(self):
        self.clear_content()
        self.current_step = 2
        self.btn_back.config(state=tk.NORMAL)
        self.btn_next.config(text="安装 >", bg="#2b78e4", fg="white")

        self.show_header("选择安装目录",
                         "选择要安装 " + APP_NAME + " 的位置")

        body = tk.Frame(self.content_frame, bg="#ffffff")
        body.pack(fill=tk.BOTH, expand=True, padx=28, pady=20)

        tk.Label(body, text="\n要在哪个文件夹安装该程序？",
                 bg="#ffffff", fg="#333333", font=("Microsoft YaHei", 10, "bold")).pack(anchor=tk.W)

        tk.Label(body, text="\n目标文件夹：",
                 bg="#ffffff", fg="#333333", font=("Microsoft YaHei", 10)).pack(anchor=tk.W, pady=(6, 2))

        dir_row = tk.Frame(body, bg="#ffffff")
        dir_row.pack(fill=tk.X, pady=4)
        entry = tk.Entry(dir_row, textvariable=self.install_dir,
                         font=("Microsoft YaHei", 10), relief=tk.SOLID, bd=1)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=4, padx=(0, 8))

        tk.Button(dir_row, text="浏览...", width=10,
                  font=("Microsoft YaHei", 10), command=self.browse_dir).pack(side=tk.LEFT, ipady=4)

        tk.Label(body, text="\n附加选项：",
                 bg="#ffffff", fg="#333333", font=("Microsoft YaHei", 10, "bold")).pack(anchor=tk.W, pady=(16, 4))

        option_frame = tk.Frame(body, bg="#ffffff")
        option_frame.pack(anchor=tk.W)

        tk.Checkbutton(option_frame, text="创建桌面快捷方式",
                       variable=self.desktop_shortcut, bg="#ffffff",
                       font=("Microsoft YaHei", 10)).pack(anchor=tk.W, pady=2)
        tk.Checkbutton(option_frame, text="创建开始菜单快捷方式",
                       variable=self.start_menu_shortcut, bg="#ffffff",
                       font=("Microsoft YaHei", 10)).pack(anchor=tk.W, pady=2)

        tk.Label(body,
                 text="\n\n点击「安装」继续。如果要查看或更改安装设置，点击「上一步」。",
                 bg="#ffffff", fg="#666666", font=("Microsoft YaHei", 9)).pack(anchor=tk.W, pady=(20, 0))

    def browse_dir(self):
        d = filedialog.askdirectory(title="选择安装目录",
                                    initialdir=self.install_dir.get() or "/")
        if d:
            self.install_dir.set(d)

    # ============ 步骤3: 安装中 ============
    def show_step_3(self):
        self.clear_content()
        self.current_step = 3
        self.btn_back.config(state=tk.DISABLED)
        self.btn_next.config(state=tk.DISABLED, bg="#cccccc")
        self.btn_cancel.config(state=tk.DISABLED)

        self.show_header("正在安装 " + APP_NAME,
                         "请稍候，正在复制文件到您的电脑...")

        body = tk.Frame(self.content_frame, bg="#ffffff")
        body.pack(fill=tk.BOTH, expand=True, padx=28, pady=20)

        tk.Label(body, text="正在复制文件，请稍候...",
                 bg="#ffffff", fg="#333333", font=("Microsoft YaHei", 11)).pack(anchor=tk.W, pady=(20, 12))

        self.progress = ttk.Progressbar(body, mode="determinate", length=560)
        self.progress.pack(fill=tk.X, pady=4)

        self.status_label = tk.Label(body, text="准备中...",
                                     bg="#ffffff", fg="#2b78e4", font=("Microsoft YaHei", 10))
        self.status_label.pack(anchor=tk.W, pady=(6, 0))

        self.log_label = tk.Label(body, text="", bg="#ffffff", fg="#999999",
                                  font=("Microsoft YaHei", 8), justify=tk.LEFT)
        self.log_label.pack(anchor=tk.W, pady=(16, 0))

        # 启动安装（延迟执行，让 UI 先绘制）
        self.after(300, self.do_install)

    def do_install(self):
        try:
            target = self.install_dir.get().strip()
            if not target:
                raise Exception("请选择安装目录")

            # 确保目录可写
            if os.path.exists(target) and os.path.isdir(target):
                if os.listdir(target):
                    if not messagebox.askyesno(
                        "目录已存在",
                        '目录 "' + target + '" 已存在且不为空。\n是否在该目录上覆盖安装？',
                        parent=self
                    ):
                        self.show_step_2()
                        return
            os.makedirs(target, exist_ok=True)

            # 1) 收集要复制的文件总数
            all_files = []
            for root, dirs, files in os.walk(self.source_dir):
                for f in files:
                    all_files.append((os.path.join(root, f),
                                      os.path.relpath(os.path.join(root, f), self.source_dir)))
            total = len(all_files)

            # 2) 复制文件（进度更新）
            copied = 0
            log_lines = []
            for src, rel in all_files:
                dst = os.path.join(target, rel)
                try:
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    shutil.copy2(src, dst)
                except Exception:
                    pass  # 跳过无法复制的文件
                copied += 1

                if copied % 50 == 0 or copied == total:
                    percent = int(copied / total * 90)
                    self.progress.config(value=percent)
                    self.status_label.config(text="复制文件... " + str(copied) + " / " + str(total))
                    if len(log_lines) < 3:
                        log_lines.append("  → " + rel)
                        self.log_label.config(text="\n".join(log_lines))
                    self.update()

            # 3) 写卸载记录
            self.status_label.config(text="写入卸载信息...")
            self.progress.config(value=93)
            self.update()
            try:
                with open(os.path.join(target, "uninstall.txt"), "w", encoding="utf-8") as f:
                    f.write("[InstallInfo]\ninstall_dir=" + target + "\nversion=" + VERSION + "\n")
            except Exception:
                pass

            # 4) 复制 uninstall.exe（如果安装源里没有）
            source_uninst = os.path.join(self.source_dir, "uninstall.exe")
            if not os.path.exists(source_uninst):
                alt = os.path.join(os.path.dirname(self.source_dir), "uninstall.exe")
                if os.path.exists(alt):
                    try:
                        shutil.copy2(alt, os.path.join(target, "uninstall.exe"))
                    except Exception:
                        pass
            else:
                try:
                    shutil.copy2(source_uninst, os.path.join(target, "uninstall.exe"))
                except Exception:
                    pass

            # 5) 创建快捷方式
            target_exe = os.path.join(target, "SFeeSystem.exe")
            if not os.path.exists(target_exe):
                # 兼容其他命名
                for name in os.listdir(target):
                    if name.lower().endswith(".exe") and name.lower() != "uninstall.exe":
                        target_exe = os.path.join(target, name)
                        break

            if self.desktop_shortcut.get():
                self.progress.config(value=96)
                self.status_label.config(text="创建桌面快捷方式...")
                self.update()
                desktop = os.path.join(os.path.expanduser("~"), "Desktop")
                if not os.path.isdir(desktop):
                    zh = os.path.join(os.path.expanduser("~"), "桌面")
                    if os.path.isdir(zh):
                        desktop = zh
                create_shortcut(target_exe, desktop, APP_NAME)

            if self.start_menu_shortcut.get():
                self.progress.config(value=98)
                self.status_label.config(text="创建开始菜单快捷方式...")
                self.update()
                start_menu = os.path.join(
                    os.environ.get("APPDATA", os.path.expanduser("~")),
                    "Microsoft", "Windows", "Start Menu", "Programs", APP_NAME
                )
                os.makedirs(start_menu, exist_ok=True)
                create_shortcut(target_exe, start_menu, APP_NAME)

            self.progress.config(value=100)
            self.status_label.config(text="安装完成！")
            self.update()
            self.after(500, lambda: self.show_step_4(target, target_exe))

        except Exception as e:
            messagebox.showerror("安装失败",
                                 "安装过程中出现错误：\n\n" + str(e),
                                 parent=self)
            self.btn_cancel.config(state=tk.NORMAL, text="关闭")
            self.btn_cancel.config(command=self.destroy)

    # ============ 步骤4: 完成 ============
    def show_step_4(self, target, target_exe):
        self.clear_content()
        self.current_step = 4
        self.btn_back.pack_forget()
        self.btn_next.config(state=tk.NORMAL, text="完成", bg="#2b78e4", fg="white",
                             command=lambda: self.finish(target_exe))
        self.btn_cancel.pack_forget()

        self.show_header("安装完成",
                         APP_NAME + " 已成功安装到您的电脑")

        body = tk.Frame(self.content_frame, bg="#ffffff")
        body.pack(fill=tk.BOTH, expand=True, padx=28, pady=20)

        tk.Label(body, text="✓ " + APP_NAME + " 安装成功！",
                 bg="#ffffff", fg="#2b78e4", font=("Microsoft YaHei", 13, "bold")).pack(anchor=tk.W, pady=(16, 8))

        tk.Label(body, text="安装位置：" + target,
                 bg="#ffffff", fg="#333333", font=("Microsoft YaHei", 10)).pack(anchor=tk.W, pady=4)

        tk.Label(body, text="\n点击「完成」关闭此向导。" + ("" if self.auto_run.get() else ""),
                 bg="#ffffff", fg="#333333", font=("Microsoft YaHei", 10)).pack(anchor=tk.W, pady=(8, 0))

        tk.Checkbutton(body, text="安装完成后立即运行 " + APP_NAME,
                       variable=self.auto_run, bg="#ffffff",
                       font=("Microsoft YaHei", 10)).pack(anchor=tk.W, pady=(12, 0))

    def finish(self, target_exe):
        if self.auto_run.get() and target_exe and os.path.exists(target_exe):
            try:
                os.startfile(target_exe)
            except Exception:
                pass
        self.destroy()

    # ============ 导航 ============
    def go_next(self):
        if self.current_step == 1:
            self.show_step_2()
        elif self.current_step == 2:
            self.show_step_3()
        # step 4 通过 finish 关闭

    def go_back(self):
        if self.current_step == 2:
            self.show_step_1()
        # 安装中和完成界面不允许回退

    def quit_app(self):
        if messagebox.askyesno("退出安装",
                               "确定要退出 " + APP_NAME + " 安装向导吗？\n程序尚未安装。",
                               parent=self):
            self.destroy()


def main():
    app = InstallerWizard()
    app.mainloop()


if __name__ == "__main__":
    main()
