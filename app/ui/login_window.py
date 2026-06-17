"""
登录窗口 + 修改密码窗口

设计要点：
  • 顶部小 logo（猴子图标）+ 标题（紧凑布局）
  • 白色圆角卡片 + 棕色主题色
  • 内置账号：user/user（3个月）  admin/admin（6个月）
  • 到期后提示"软件已到期，请联系管理员"
  • 窗口图标 + 主窗口图标：dasheng.ico
  • 窗口：460 × 520（紧凑、可居中、不遮挡屏幕）
"""
import os
import sys
from typing import Optional

from PyQt5.QtWidgets import (
    QLabel, QLineEdit, QPushButton, QVBoxLayout,
    QHBoxLayout, QMessageBox, QDialog,
    QFrame, QGraphicsDropShadowEffect,
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QIcon, QPixmap, QColor

from app.models.user import (
    UserService,
    clear_remember_token,
    get_remaining_days,
    get_expire_date,
)
from app.models.path_config import get_resource_path

APP_TITLE = "大圣快递账单结算系统"

COLOR_BG_TOP = "#6d4c2e"
COLOR_BG_BOTTOM = "#8b6239"


def _icon_path():
    """登录窗口图标（.ico 格式）"""
    p = os.path.join(get_resource_path("data", "icons"), "dasheng.ico")
    if os.path.exists(p):
        return p
    return ""


def _logo_png_path():
    """猴子 logo（PNG 格式）"""
    p = os.path.join(get_resource_path("data", "icons"), "dasheng_200.png")
    if os.path.exists(p):
        return p
    p2 = os.path.join(get_resource_path("data", "icons"), "monkey-icon.png")
    if os.path.exists(p2):
        return p2
    return ""


# ===================== 登录窗口 =====================

class LoginDialog(QDialog):

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_TITLE} - 用户登录")
        self.setFixedSize(460, 520)             # 紧凑尺寸：不会溢出屏幕
        self.setModal(True)
        self.logged_in_username = None

        ip = _icon_path()
        if ip:
            self.setWindowIcon(QIcon(ip))

        self.setStyleSheet(f"""
            QDialog {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {COLOR_BG_TOP}, stop:1 {COLOR_BG_BOTTOM});
            }}
        """)

        self._build()

    def _build(self):
        main = QVBoxLayout(self)
        main.setContentsMargins(24, 20, 24, 20)
        main.setSpacing(14)

        # ===== 顶部：小 logo + 标题
        logo_label = QLabel()
        logo_path = _logo_png_path()
        if logo_path:
            pix = QPixmap(logo_path)
            logo_label.setPixmap(pix.scaled(
                72, 72, Qt.KeepAspectRatio, Qt.SmoothTransformation
            ))
        logo_label.setAlignment(Qt.AlignCenter)
        logo_label.setFixedHeight(80)

        title_main = QLabel(APP_TITLE)
        f1 = QFont()
        f1.setPointSize(15)
        f1.setBold(True)
        title_main.setFont(f1)
        title_main.setAlignment(Qt.AlignCenter)
        title_main.setStyleSheet("color:#ffffff;")

        title_sub = QLabel("用户登录")
        title_sub.setAlignment(Qt.AlignCenter)
        title_sub.setStyleSheet("color:#ffffff; font-size:12px;")

        main.addWidget(logo_label)
        main.addWidget(title_main)
        main.addWidget(title_sub)
        main.addSpacing(8)

        # ===== 白色卡片
        card = QFrame(self)
        card.setObjectName("login_card")
        card.setStyleSheet("""
            QFrame#login_card {
                background-color: #ffffff;
                border-radius: 12px;
            }
        """)
        shadow = QGraphicsDropShadowEffect(card)
        shadow.setBlurRadius(30)
        shadow.setYOffset(6)
        shadow.setColor(QColor(0, 0, 0, 80))
        card.setGraphicsEffect(shadow)

        cl = QVBoxLayout(card)
        cl.setContentsMargins(28, 20, 28, 18)
        cl.setSpacing(10)

        # ---- 用户名
        lbl1 = QLabel("用户名")
        lbl1.setStyleSheet("color:#555; font-size:12px; font-weight:bold;")
        cl.addWidget(lbl1)

        self.username_edit = QLineEdit()
        self.username_edit.setPlaceholderText("请输入用户名（user 或 admin）")
        self.username_edit.setFixedHeight(36)
        self.username_edit.setStyleSheet("""
            QLineEdit {
                padding: 6px 12px;
                border: 1px solid #d7d7d7;
                border-radius: 6px;
                font-size: 13px;
                background: #fafafa;
            }
            QLineEdit:focus { border: 1px solid #6d4c2e; background: #ffffff; }
        """)
        cl.addWidget(self.username_edit)

        # ---- 密码行
        pwd_header = QHBoxLayout()
        pwd_header.setContentsMargins(0, 0, 0, 0)
        lbl2 = QLabel("密码")
        lbl2.setStyleSheet("color:#555; font-size:12px; font-weight:bold;")

        self.toggle_btn = QPushButton("显示")
        self.toggle_btn.setCursor(Qt.PointingHandCursor)
        self.toggle_btn.setCheckable(True)
        self.toggle_btn.setStyleSheet("""
            QPushButton {
                background: transparent;
                color: #8b6239;
                border: none;
                padding: 0;
                font-size: 12px;
            }
            QPushButton:hover { text-decoration: underline; }
        """)
        self.toggle_btn.toggled.connect(self._toggle_password)

        pwd_header.addWidget(lbl2)
        pwd_header.addStretch(1)
        pwd_header.addWidget(self.toggle_btn)
        cl.addLayout(pwd_header)

        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.password_edit.setPlaceholderText("请输入密码")
        self.password_edit.setFixedHeight(36)
        self.password_edit.setStyleSheet(self.username_edit.styleSheet())
        cl.addWidget(self.password_edit)

        cl.addSpacing(4)

        # ---- 登录按钮
        self.btn_login = QPushButton("登 录")
        self.btn_login.setFixedHeight(40)
        self.btn_login.setCursor(Qt.PointingHandCursor)
        self.btn_login.setStyleSheet("""
            QPushButton {
                background-color: #6d4c2e;
                color: white;
                border: none;
                border-radius: 6px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #8b6239; }
            QPushButton:pressed { background-color: #5a3d24; }
        """)
        self.btn_login.clicked.connect(self._on_login)
        self.btn_login.setDefault(True)
        cl.addWidget(self.btn_login)

        # ---- 退出按钮
        self.btn_quit = QPushButton("退 出")
        self.btn_quit.setFixedHeight(34)
        self.btn_quit.setCursor(Qt.PointingHandCursor)
        self.btn_quit.setStyleSheet("""
            QPushButton {
                background-color: #eeeeee;
                color: #555;
                border: none;
                border-radius: 6px;
                font-size: 12px;
            }
            QPushButton:hover { background-color: #e0e0e0; }
        """)
        self.btn_quit.clicked.connect(self.reject)
        cl.addWidget(self.btn_quit)

        # ---- 版权（放在卡片底部居中）
        footer = QLabel("© 杭州喵喵至家网络有限公司")
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet("color:#999; font-size:10px;")
        cl.addSpacing(4)
        cl.addWidget(footer)

        main.addWidget(card, 1)

    def _toggle_password(self, checked):
        if checked:
            self.password_edit.setEchoMode(QLineEdit.Normal)
            self.toggle_btn.setText("隐藏")
        else:
            self.password_edit.setEchoMode(QLineEdit.Password)
            self.toggle_btn.setText("显示")

    def _on_login(self):
        username = self.username_edit.text().strip()
        password = self.password_edit.text()

        if not username or not password:
            QMessageBox.warning(self, "提示", "请输入用户名和密码")
            return

        ok, msg, user = UserService.verify_login(username, password)
        if not ok:
            if "已到期" in msg:
                QMessageBox.critical(self, "软件已到期", msg)
            else:
                QMessageBox.warning(self, "登录失败", msg)
            return

        if user is not None:
            remaining = get_remaining_days(user)
            exp_date = get_expire_date(user)
            if remaining <= 30:
                tip = f"欢迎，{username}\n\n有效期至：{exp_date.strftime('%Y-%m-%d')}\n剩余 {remaining} 天"
                if remaining <= 7:
                    tip += "\n\n如需续费，请联系管理员"
                QMessageBox.information(self, "登录成功", tip)

        self.logged_in_username = username
        self.accept()


# ===================== 修改密码窗口 =====================

class ChangePasswordDialog(QDialog):

    def __init__(self, username, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"{APP_TITLE} - 修改密码")
        self.setFixedSize(420, 380)
        self.setModal(True)
        self._username = username

        ip = _icon_path()
        if ip:
            self.setWindowIcon(QIcon(ip))

        self.setStyleSheet(f"""
            QDialog {{
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 {COLOR_BG_TOP}, stop:1 {COLOR_BG_BOTTOM});
            }}
        """)
        self._build()

    def _build(self):
        main = QVBoxLayout(self)
        main.setContentsMargins(24, 18, 24, 18)
        main.setSpacing(12)

        title = QLabel("修改密码")
        f = QFont()
        f.setPointSize(15)
        f.setBold(True)
        title.setFont(f)
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color:#ffffff;")

        sub = QLabel(f"当前用户：{self._username}")
        sub.setAlignment(Qt.AlignCenter)
        sub.setStyleSheet("color:#ffffff; font-size:12px;")

        main.addWidget(title)
        main.addWidget(sub)
        main.addSpacing(8)

        card = QFrame(self)
        card.setObjectName("pwd_card")
        card.setStyleSheet("""
            QFrame#pwd_card {
                background-color: #ffffff;
                border-radius: 12px;
            }
        """)
        shadow = QGraphicsDropShadowEffect(card)
        shadow.setBlurRadius(30)
        shadow.setYOffset(6)
        shadow.setColor(QColor(0, 0, 0, 80))
        card.setGraphicsEffect(shadow)

        cl = QVBoxLayout(card)
        cl.setContentsMargins(28, 18, 28, 16)
        cl.setSpacing(10)

        self.old_edit = self._add_pwd_row(cl, "当前密码")
        self.new_edit = self._add_pwd_row(cl, "新密码（至少 6 位）")
        self.confirm_edit = self._add_pwd_row(cl, "确认新密码")

        cl.addSpacing(4)

        self.btn_ok = QPushButton("确认修改")
        self.btn_ok.setFixedHeight(38)
        self.btn_ok.setCursor(Qt.PointingHandCursor)
        self.btn_ok.setStyleSheet("""
            QPushButton {
                background-color: #6d4c2e;
                color: white; border: none;
                border-radius: 6px; font-size: 13px; font-weight: bold;
            }
            QPushButton:hover { background-color: #8b6239; }
        """)
        self.btn_ok.clicked.connect(self._on_ok)
        cl.addWidget(self.btn_ok)

        self.btn_cancel = QPushButton("取 消")
        self.btn_cancel.setFixedHeight(32)
        self.btn_cancel.setCursor(Qt.PointingHandCursor)
        self.btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #eeeeee;
                color: #555; border: none;
                border-radius: 6px; font-size: 12px;
            }
            QPushButton:hover { background-color: #e0e0e0; }
        """)
        self.btn_cancel.clicked.connect(self.reject)
        cl.addWidget(self.btn_cancel)

        main.addWidget(card, 1)

    def _add_pwd_row(self, parent_layout, label_text):
        lbl = QLabel(label_text)
        lbl.setStyleSheet("color:#555; font-size:12px; font-weight:bold;")
        parent_layout.addWidget(lbl)

        edit = QLineEdit()
        edit.setEchoMode(QLineEdit.Password)
        edit.setFixedHeight(36)
        edit.setPlaceholderText("请输入")
        edit.setStyleSheet("""
            QLineEdit {
                padding: 6px 12px;
                border: 1px solid #d7d7d7;
                border-radius: 6px;
                font-size: 13px;
                background: #fafafa;
            }
            QLineEdit:focus { border: 1px solid #6d4c2e; background: #ffffff; }
        """)
        parent_layout.addWidget(edit)
        return edit

    def _on_ok(self):
        old = self.old_edit.text()
        new = self.new_edit.text()
        confirm = self.confirm_edit.text()

        if not old or not new:
            QMessageBox.warning(self, "提示", "请输入密码")
            return
        if len(new) < 6:
            QMessageBox.warning(self, "提示", "新密码至少 6 位")
            return
        if new == old:
            QMessageBox.warning(self, "提示", "新密码不能与当前密码相同")
            return
        if new != confirm:
            QMessageBox.warning(self, "提示", "两次输入的新密码不一致")
            return

        ok, msg = UserService.change_password(self._username, old, new)
        if not ok:
            QMessageBox.critical(self, "修改失败", msg)
            return

        clear_remember_token()
        QMessageBox.information(self, "修改成功", "密码已更新，请使用新密码登录")
        self.accept()


# ===================== 对外入口 =====================

def show_login_flow() -> Optional[str]:
    """启动登录流程，返回登录成功的用户名；失败/退出返回 None"""
    from PyQt5.QtWidgets import QApplication

    UserService.ensure_builtin_accounts()

    dlg = LoginDialog()
    if dlg.exec_() == QDialog.Accepted and dlg.logged_in_username:
        return dlg.logged_in_username
    return None
