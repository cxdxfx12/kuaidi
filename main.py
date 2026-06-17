"""
派费计算系统 - 桌面应用（支持300万行大数据 + 登录认证）
杭州喵喵至家网络有限公司
"""
import sys
import os
import traceback
import logging

# PyInstaller 打包后，多进程子进程需要调用 freeze_support()
if hasattr(sys, 'frozen'):
    try:
        import multiprocessing
        multiprocessing.freeze_support()
    except Exception:
        pass

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 配置日志
log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
os.makedirs(log_dir, exist_ok=True)
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(log_dir, 'app.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)

from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QIcon


def main():
    """应用入口（先初始化数据库 → 登录验证 → 主窗口）"""
    try:
        app = QApplication(sys.argv)
        app.setApplicationName("大圣.快递物流出港帐单结算系统")
        app.setOrganizationName("杭州喵喵至家网络有限公司")
        app.setStyle("Fusion")
        logging.info("Qt应用初始化成功")

        # 1) 先确保数据库存在，并触发建表（包含 users / login_tokens）
        logging.info("正在初始化数据库...")
        from app.models.database import get_session
        try:
            _ = get_session()
            logging.info("数据库初始化成功")
        except Exception as e:
            logging.error(f"数据库初始化失败: {e}")
            raise

        # 2) 登录流程（首次启动 → 创建管理员；否则 → 登录窗口或记住我）
        logging.info("正在显示登录窗口...")
        from app.ui.login_window import show_login_flow
        username = show_login_flow()
        logging.info(f"登录结果: username={username}")
        if not username:
            logging.info("用户取消登录")
            sys.exit(0)

        # 3) 打开主窗口（传入当前用户名）
        logging.info(f"正在打开主窗口，用户: {username}")
        from app.ui.main_window import MainWindow
        window = MainWindow(current_user=username)
        window.show()
        logging.info("主窗口显示成功")

        sys.exit(app.exec_())
        
    except Exception as e:
        logging.error(f"应用崩溃: {e}")
        logging.error(traceback.format_exc())
        # 显示错误对话框
        try:
            from PyQt5.QtWidgets import QMessageBox
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("应用错误")
            msg.setText(f"应用启动失败:\n\n{e}\n\n详细日志已保存到 logs/app.log")
            msg.exec_()
        except Exception:
            pass
        sys.exit(1)


if __name__ == "__main__":
    main()
