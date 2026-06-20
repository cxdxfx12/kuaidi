"""
主窗口 - 运费计算系统
"""
import os
import sys
import json
from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QPushButton, QLabel, QFileDialog,
    QTableWidget, QTableWidgetItem, QProgressBar,
    QMessageBox, QStatusBar, QTextEdit, QGroupBox,
    QHeaderView, QAbstractItemView, QLineEdit, QSizePolicy,
    QComboBox, QInputDialog, QRadioButton, QScrollArea,
    QFrame, QDialog, QProgressDialog, QApplication
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QIcon

from app.models.database import init_db, get_session
from app.models.path_config import get_config_file, get_data_dir, get_user_data_dir, get_resource_path
from app.models.fee_record import FeeRecord
from app.models.fee_detail import FeeDetail
from app.services.calculate_service import CalculateService
from app.services.export_service import ExportService
from app.services.column_matcher import ColumnMatcher
from app.services.rule_service import RuleService, Rule
from app.core.settlement import SettlementEngine


class CalculateWorker(QThread):
    """计算后台线程，避免UI卡死 + 支持进度反馈 + 多文件批量处理"""
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)
    progress = pyqtSignal(int, str)  # (percent, stage_text)  percent: 0-100
    file_progress = pyqtSignal(int, int, str)  # (file_index, total_files, file_name)

    def __init__(self, file_paths, sheet_name=None):
        super().__init__()
        # 兼容单文件和多文件
        if isinstance(file_paths, (list, tuple)):
            self.file_paths = list(file_paths)
        else:
            self.file_paths = [file_paths]
        self.sheet_name = sheet_name

    def run(self):
        try:
            service = CalculateService()
            total_files = len(self.file_paths)
            all_results = []
            total_rows_all = 0
            total_fee_all = 0.0
            total_success_all = 0
            total_exception_all = 0
            failed_files = []

            for file_idx, file_path in enumerate(self.file_paths):
                file_name = os.path.basename(file_path)
                self.file_progress.emit(file_idx + 1, total_files, file_name)

                # 为每个文件定义进度回调，映射到总体进度
                def make_cb(fidx, total):
                    def cb(percent, stage_text):
                        # 每个文件占 100/total % 的总体进度
                        overall = int((fidx * 100 + percent) / total)
                        overall = min(max(overall, 0), 100)
                        self.progress.emit(overall, f"[第{fidx+1}/{total}个] {stage_text}")
                    return cb

                cb = make_cb(file_idx, total_files)

                try:
                    result = service.import_and_calculate(file_path, self.sheet_name, progress_callback=cb)
                    all_results.append({
                        "file_name": file_name,
                        "record_id": result["record_id"],
                        "total_fee": result["total_fee"],
                        "success_count": result["success_count"],
                        "exception_count": result["exception_count"],
                        "total_rows": result["total_rows"]
                    })
                    total_rows_all += result.get("total_rows", 0)
                    total_fee_all += result.get("total_fee", 0.0)
                    total_success_all += result.get("success_count", 0)
                    total_exception_all += result.get("exception_count", 0)
                except Exception as file_err:
                    failed_files.append(f"{file_name}: {file_err}")
                    continue

            # 发送100%完成信号
            self.progress.emit(100, "全部文件处理完成")

            # 返回汇总结果（第一个记录的id用于切换到结果tab）
            final_result = {
                "record_id": all_results[0]["record_id"] if all_results else None,
                "total_fee": total_fee_all,
                "success_count": total_success_all,
                "exception_count": total_exception_all,
                "total_rows": total_rows_all,
                "file_count": total_files,
                "files": all_results,
                "failed_files": failed_files
            }
            self.finished.emit(final_result)
        except Exception as e:
            self.error.emit(str(e))


class FilePreviewWorker(QThread):
    """文件预览后台线程 - 解决加载大Excel时界面无响应问题"""
    finished = pyqtSignal(dict)  # (parse_result)
    error = pyqtSignal(str)
    progress = pyqtSignal(int, str)  # 读取过程中的进度反馈

    def __init__(self, file_path: str):
        super().__init__()
        self.file_path = file_path

    def run(self):
        try:
            self.progress.emit(5, "开始读取文件...")
            from app.services.excel_parser import ExcelParser
            parser = ExcelParser()

            def cb(percent):
                self.progress.emit(int(percent), f"读取文件中... {percent}%")

            parse_result = parser.parse(self.file_path, row_callback=cb)
            self.finished.emit(parse_result)
        except Exception as e:
            self.error.emit(str(e))


class ResultLoadWorker(QThread):
    """结果数据加载后台线程 - 仅加载显示所需数据，避免UI卡死和内存溢出"""
    finished = pyqtSignal(dict)  # {summary: {...}, prepared: [rows], total_rows, display_count}
    error = pyqtSignal(str)

    def __init__(self, record_id, max_display=20000):
        super().__init__()
        self.record_id = record_id
        self.max_display = max_display

    def run(self):
        try:
            from app.models.database import get_session
            from app.models.fee_detail import FeeDetail
            from app.models.fee_record import FeeRecord
            import json
            import re as _re

            session = get_session()
            try:
                # 优先从 FeeRecord 中直接读取已存储的汇总值（更高效）
                record = session.query(FeeRecord).filter(
                    FeeRecord.id == self.record_id
                ).first()

                if record:
                    total_rows = int(record.total_rows or 0)
                    total_fee = float(getattr(record, "total_fee", 0) or 0)
                    success_count = int(record.success_rows or 0)
                    exception_count = int(record.error_rows or 0)
                else:
                    # 回退：实际查询统计
                    total_rows = session.query(FeeDetail).filter(
                        FeeDetail.record_id == self.record_id
                    ).count()
                    total_fee = 0.0
                    exception_count = 0
                    # 分批统计以避免内存溢出
                    batch_size = 5000
                    offset = 0
                    while True:
                        batch = session.query(FeeDetail).filter(
                            FeeDetail.record_id == self.record_id
                        ).order_by(FeeDetail.id).limit(batch_size).offset(offset).all()
                        if not batch:
                            break
                        for d in batch:
                            total_fee += float(d.calculated_fee or 0)
                            if d.is_exception:
                                exception_count += 1
                        offset += batch_size
                    success_count = total_rows - exception_count

                summary = {
                    "total_rows": total_rows,
                    "total_fee": total_fee,
                    "success_count": success_count,
                    "exception_count": exception_count,
                }

                # 仅取前 max_display 行用于显示
                display_details = session.query(FeeDetail).filter(
                    FeeDetail.record_id == self.record_id
                ).order_by(FeeDetail.row_index).limit(self.max_display).all()
            finally:
                session.close()

            display_count = len(display_details)

            # 预解析为纯Python元组（避免传递SQLAlchemy对象）
            prepared = []
            for d in display_details:
                business_date = ""
                customer_name = ""
                try:
                    if d.original_data:
                        od = d.original_data if isinstance(d.original_data, dict) else json.loads(d.original_data)
                        raw_date = od.get("business_date", "")
                        if raw_date:
                            s = str(raw_date)
                            digits = _re.findall(r"\d+", s)
                            if len(digits) >= 3:
                                business_date = f"{int(digits[0]):04d}/{int(digits[1]):02d}/{int(digits[2]):02d}"
                            else:
                                business_date = s
                        customer_name = str(od.get("customer_name", "") or "")
                except Exception:
                    pass

                prepared.append((
                    str(d.row_index),
                    business_date,
                    d.tracking_no or "",
                    f"{d.station_code or ''} {d.station_name or ''}",
                    d.region_name or "",
                    f"{float(d.weight or 0):.3f}",
                    customer_name,
                    str(d.quantity or 0),
                    f"{float(d.calculated_fee or 0):.2f}",
                    d.rule_name or "",
                    "⚠️ 异常" if d.is_exception else "✓"
                ))

            # 重要：不再传递完整details，仅传递显示所需数据
            # 结算和导出功能会按需独立加载数据
            self.finished.emit({
                "summary": summary,
                "prepared": prepared,
                "total_rows": summary["total_rows"],
                "display_count": display_count,
            })
        except Exception as e:
            import traceback
            self.error.emit(str(e) + "\n" + traceback.format_exc())


class ExportWorker(QThread):
    """导出 Excel 后台线程 - 避免大数据量导出时 UI 未响应"""
    finished = pyqtSignal(str)  # 成功：返回文件路径
    error = pyqtSignal(str)
    progress = pyqtSignal(int, str)  # 进度：百分比 + 消息

    def __init__(self, record_id: int, target_file_path: str):
        super().__init__()
        self.record_id = record_id
        self.target_file_path = target_file_path

    def run(self):
        try:
            from app.services.export_service import ExportService
            service = ExportService()
            # 进度回调：发射 progress 信号，由 UI 线程更新浮层
            def on_progress(pct, msg):
                self.progress.emit(pct, msg)

            file_path = service.export_details(
                self.record_id, self.target_file_path, progress_callback=on_progress
            )
            self.finished.emit(file_path)
        except Exception as e:
            import traceback
            self.error.emit(str(e) + "\n" + traceback.format_exc())


class ExportSettlementWorker(QThread):
    """结算单导出后台线程"""
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, details, settlement_type, group_key, target_file_path, record_id=None):
        super().__init__()
        self.details = details
        self.settlement_type = settlement_type
        self.group_key = group_key
        self.target_file_path = target_file_path
        self.record_id = record_id

    def run(self):
        try:
            from app.services.export_service import ExportService
            # export_settlement_details 内部会自己决定目录，但传入的 details 已经在内存中
            # 为了保证写到用户指定文件，这里直接写入
            from openpyxl import Workbook

            type_names = {"station": "网点", "contract": "承包区", "monthly": "月结客户"}
            type_name = type_names.get(self.settlement_type, "结算")

            filtered = []
            for d in self.details:
                if self.settlement_type == "station" and d.station_code == self.group_key:
                    filtered.append(d)
                elif self.settlement_type == "contract":
                    sc = d.station_code or ""
                    if len(sc) >= 3 and sc[:3] == self.group_key:
                        filtered.append(d)
                elif self.settlement_type == "monthly":
                    original = {}
                    try:
                        if d.original_data:
                            import json as _json
                            original = d.original_data if isinstance(d.original_data, dict) \
                                else _json.loads(d.original_data)
                    except Exception:
                        pass
                    cc = original.get("客户编码", original.get("客户代码", ""))
                    if str(cc).strip() == self.group_key:
                        filtered.append(d)

            wb = Workbook(write_only=True)
            ws = wb.create_sheet(type_name)
            ws.append(["行号", "快递单号", "网点编码", "网点名称", "区域",
                       "重量(kg)", "件数", "运费(元)", "应用规则", "是否异常", "备注"])
            for d in filtered:
                ws.append([
                    str(d.row_index), d.tracking_no or "", d.station_code or "",
                    d.station_name or "", d.region_name or "",
                    float(d.weight or 0), int(d.quantity or 0),
                    float(d.calculated_fee or 0), d.rule_name or "",
                    "是" if d.is_exception else "否", d.remark or "",
                ])
            wb.save(self.target_file_path)
            self.finished.emit(self.target_file_path)
        except Exception as e:
            import traceback
            self.error.emit(str(e) + "\n" + traceback.format_exc())


class ExportMultiWorker(QThread):
    """多文件导出后台线程"""
    finished = pyqtSignal(list)  # 成功：返回文件路径列表
    error = pyqtSignal(str)
    progress = pyqtSignal(int, str)

    def __init__(self, record_ids: list, export_dir: str):
        super().__init__()
        self.record_ids = record_ids
        self.export_dir = export_dir

    def run(self):
        try:
            from app.services.export_service import ExportService
            service = ExportService()
            def on_progress(pct, msg):
                self.progress.emit(pct, msg)
            exported_files = service.export_multiple_records(
                self.record_ids, self.export_dir, progress_callback=on_progress
            )
            self.finished.emit(exported_files)
        except Exception as e:
            import traceback
            self.error.emit(str(e) + "\n" + traceback.format_exc())


class ExportMergedWorker(QThread):
    """合并导出后台线程"""
    finished = pyqtSignal(list)  # 成功：返回文件路径列表
    error = pyqtSignal(str)
    progress = pyqtSignal(int, str)

    def __init__(self, record_ids: list, export_dir: str):
        super().__init__()
        self.record_ids = record_ids
        self.export_dir = export_dir

    def run(self):
        try:
            from app.services.export_service import ExportService
            service = ExportService()
            def on_progress(pct, msg):
                self.progress.emit(pct, msg)
            exported_files = service.export_merged_records(
                self.record_ids, self.export_dir, progress_callback=on_progress
            )
            self.finished.emit(exported_files)
        except Exception as e:
            import traceback
            self.error.emit(str(e) + "\n" + traceback.format_exc())


# 34省分组定义（用于简化客户规则配置）
# 格式: (分组名, 省份列表, 默认首重, 默认续重, 默认保底, 默认续重单位, 默认进位模式)
# 共11个分组，覆盖34省+港澳台
PROVINCE_GROUPS = [
    ("一区",
     ["浙江"],
     2.5, 1.0, 2.5, "kg", "actual"),

    ("二区",
     ["江苏", "安徽"],
     3.1, 1.0, 3.1, "kg", "actual"),

    ("三区",
     ["天津", "河北", "山东", "山西", "河南", "湖北", "湖南", "江西", "福建", "广东"],
     3.1, 1.5, 3.1, "kg", "actual"),

    ("四区",
     ["上海"],
     3.1, 1.5, 3.1, "kg", "actual"),

    ("五区",
     ["北京"],
     3.1, 1.5, 3.1, "kg", "actual"),

    ("六区",
     ["重庆"],
     3.6, 2.0, 3.6, "kg", "actual"),

    ("七区",
     ["广西"],
     3.1, 2.0, 3.1, "kg", "actual"),

    ("八区",
     ["黑龙江", "吉林", "辽宁", "陕西", "内蒙古"],
     3.6, 1.5, 3.6, "kg", "actual"),

    ("九区",
     ["四川", "贵州", "云南"],
     3.6, 2.0, 3.6, "kg", "actual"),

    ("十区",
     ["甘肃", "宁夏", "青海"],
     3.6, 2.0, 3.6, "kg", "actual"),

    ("十一区",
     ["新疆"],
     3.6, 2.0, 3.6, "kg", "actual"),

    ("十二区",
     ["西藏"],
     5.0, 2.5, 5.0, "kg", "actual"),

    ("十三区",
     ["海南"],
     4.0, 1.8, 4.0, "kg", "actual"),

    ("十四区",
     ["香港", "澳门", "台湾"],
     30.0, 20.0, 30.0, "kg", "actual"),
]


class MainWindow(QMainWindow):
    """主窗口"""

    def __init__(self, current_user=None):
        super().__init__()
        self._current_user = current_user or ""
        if self._current_user:
            self.setWindowTitle(
                f"大圣.快递物流出港帐单结算系统  -  用户：{self._current_user}  ·  杭州喵喵至家网络有限公司  ·  大圣智慧软件  17771300068 / 19171045360"
            )
        else:
            self.setWindowTitle("大圣.快递物流出港帐单结算系统  -  杭州喵喵至家网络有限公司  ·  大圣智慧软件  17771300068 / 19171045360")
        self.resize(1180, 640)
        self.setMinimumSize(960, 540)

        # 窗口图标（猴子图标）
        try:
            icon_file = os.path.join(get_resource_path("data", "icons"), "monkey-icon.png")
            if os.path.exists(icon_file):
                self.setWindowIcon(QIcon(icon_file))
        except Exception:
            pass

        self.current_record_id = None
        self.all_record_ids = []  # 保存所有导入文件的记录ID列表
        self.current_details = []
        self._settlement_cache = {}
        self._export_overlay = None  # 导出进度提示浮层

        self._init_database()

        self._init_ui()

        self._init_menu_bar()

        self._load_default_settings()

    def resizeEvent(self, event):
        """窗口大小变化时，重新居中导出浮层"""
        super().resizeEvent(event)
        if self._export_overlay is not None and self._export_overlay.isVisible():
            geo = self.geometry()
            cx = geo.x() + geo.width() // 2
            cy = geo.y() + geo.height() // 2
            self._export_overlay.move(cx - 220, cy - 130)

    def moveEvent(self, event):
        """窗口移动时，导出浮层跟随居中"""
        super().moveEvent(event)
        if self._export_overlay is not None and self._export_overlay.isVisible():
            geo = self.geometry()
            cx = geo.x() + geo.width() // 2
            cy = geo.y() + geo.height() // 2
            self._export_overlay.move(cx - 220, cy - 130)

    def _init_menu_bar(self):
        """顶部菜单栏（账号相关）：修改密码 / 退出登录 / 退出软件"""
        try:
            from app.models.user import clear_remember_token
            from app.ui.login_window import ChangePasswordDialog, show_login_flow
        except Exception:
            return
        if not self._current_user:
            return
        menubar = self.menuBar()
        account_menu = menubar.addMenu("账号(&A)")
        act_pwd = account_menu.addAction("修改密码...")
        act_logout = account_menu.addAction("退出登录")
        act_exit = account_menu.addAction("退出软件")

        def _change_password():
            dlg = ChangePasswordDialog(self._current_user, self)
            dlg.exec_()

        def _logout():
            from PyQt5.QtWidgets import QApplication
            reply = QMessageBox.question(self, "退出登录", "确认退出当前账号？",
                                         QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                clear_remember_token()
                self.close()
                new_username = show_login_flow()
                if new_username:
                    new_win = MainWindow(current_user=new_username)
                    new_win.show()
                    QApplication.instance()._new_main_window_ref = new_win
                else:
                    QApplication.instance().quit()

        def _exit_app():
            from PyQt5.QtWidgets import QApplication
            QApplication.instance().quit()

        act_pwd.triggered.connect(_change_password)
        act_logout.triggered.connect(_logout)
        act_exit.triggered.connect(_exit_app)

    def _init_database(self):
        """初始化数据库"""
        try:
            init_db()
            # 初始化列名映射
            matcher = ColumnMatcher()
            matcher.init_default_mappings()
        except Exception as e:
            QMessageBox.critical(self, "数据库错误", f"数据库初始化失败：{e}")

    def _init_ui(self):
        """初始化界面 - 液态玻璃 (Liquid Glass) 风格"""
        self.setStyleSheet("""
            /* ========== 全局 ========== */
            QMainWindow { background: #eef2f7; }
            QWidget {
                background: #eef2f7;
                color: #1e293b;
                font-family: "Microsoft YaHei", "PingFang SC", "Segoe UI";
                font-size: 11px;
            }

            /* ========== 顶栏品牌条：液态玻璃渐变 ========== */
            #brand_bar {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #4f46e5, stop:0.5 #6366f1, stop:1 #4f46e5);
                border-bottom: 1px solid #c7d2fe;
            }
            #brand_title {
                color: white;
                font-size: 16px;
                font-weight: 500;
                padding-left: 4px;
            }
            #brand_sub {
                color: #e0e7ff;
                font-size: 12px;
            }

            /* ========== Tab页：液态玻璃胶囊 ========== */
            QTabWidget::pane {
                border: none;
                background: #eef2f7;
                margin-top: 0px;
            }
            QTabBar::tab {
                background: #f1f5f9;
                color: #64748b;
                padding: 8px 20px;
                margin-right: 4px;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                font-size: 11px;
                font-weight: 500;
            }
            QTabBar::tab:selected {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #ffffff, stop:1 #eef2ff);
                color: #4f46e5;
                border: 1px solid #c7d2fe;
            }
            QTabBar::tab:hover:!selected {
                background: #ffffff;
                color: #4f46e5;
                border: 1px solid #cbd5e1;
            }

            /* ========== 卡片容器：玻璃卡片 ========== */
            QGroupBox {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                padding: 10px 12px;
                margin-top: 10px;
                font-size: 11px;
                font-weight: 500;
                color: #334155;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 12px;
                padding: 0 10px;
                background: #ffffff;
                color: #4f46e5;
            }

            /* ========== 主按钮：玻璃态胶囊 ========== */
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #6366f1, stop:1 #4f46e5);
                color: white;
                border: none;
                padding: 6px 14px;
                border-radius: 8px;
                font-size: 11px;
                font-weight: 500;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #818cf8, stop:1 #6366f1);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #4f46e5, stop:1 #4338ca);
            }
            QPushButton:disabled {
                background: #cbd5e1;
                color: #94a3b8;
            }

            /* ========== 绿色主按钮（开始计算） ========== */
            QPushButton#primary_green {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #34d399, stop:1 #10b981);
                padding: 10px;
                font-size: 12px;
                font-weight: 500;
                border-radius: 8px;
            }
            QPushButton#primary_green:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #6ee7b7, stop:1 #34d399);
            }

            /* ========== 次要按钮（描边） ========== */
            QPushButton#secondary {
                background: #ffffff;
                color: #475569;
                border: 1px solid #cbd5e1;
                padding: 6px 14px;
                border-radius: 8px;
                font-size: 11px;
            }
            QPushButton#secondary:hover {
                background: #eef2ff;
                color: #4f46e5;
                border-color: #a5b4fc;
            }

            /* ========== 红色按钮 ========== */
            QPushButton#danger {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f87171, stop:1 #ef4444);
                color: white;
                padding: 6px 14px;
                border-radius: 8px;
                font-size: 11px;
            }
            QPushButton#danger:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #fca5a5, stop:1 #f87171);
            }

            /* ========== 标签 ========== */
            QLabel { color: #334155; font-size: 11px; }
            QLabel#hint {
                color: #64748b;
                padding: 10px 12px;
                background: #f8fafc;
                border: 1px dashed #cbd5e1;
                border-radius: 8px;
                font-size: 11px;
            }
            QLabel#success_hint {
                color: #047857;
                padding: 10px 12px;
                background: #ecfdf5;
                border: 1px solid #a7f3d0;
                border-radius: 8px;
                font-size: 11px;
            }

            /* ========== 输入框 / 表格 ========== */
            QLineEdit, QTextEdit, QTableWidget {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                color: #1e293b;
                padding: 3px 8px;
                selection-background-color: #e0e7ff;
                selection-color: #1e293b;
                font-size: 11px;
                min-height: 22px;
            }
            QLineEdit:focus, QTextEdit:focus, QTableWidget:focus {
                border: 1px solid #a5b4fc;
            }
            QTableWidget {
                gridline-color: #e2e8f0;
                padding: 0px;
                selection-background-color: #e0e7ff;
                selection-color: #1e293b;
                alternate-background-color: #f1f5f9;
            }
            QTableWidget::item { padding: 4px 8px; }
            QTableWidget::item:selected {
                background: #e0e7ff;
                color: #1e293b;
            }

            /* ========== 表头：玻璃亮白 ========== */
            QHeaderView::section {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #ffffff, stop:1 #f1f5f9);
                color: #475569;
                padding: 6px 8px;
                border: none;
                border-right: 1px solid #e2e8f0;
                border-bottom: 1px solid #e2e8f0;
                font-size: 11px;
                font-weight: 600;
            }
            QHeaderView::section:last { border-right: none; }

            /* ========== 进度条 ========== */
            QProgressBar {
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                background: #f8fafc;
                text-align: center;
                color: #334155;
                height: 18px;
                font-size: 11px;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #818cf8, stop:1 #4f46e5);
                border-radius: 7px;
            }

            /* ========== 下拉框 ========== */
            QComboBox {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 3px 12px;
                color: #1e293b;
                min-width: 80px;
                min-height: 24px;
                font-size: 11px;
            }
            QComboBox:hover { border-color: #a5b4fc; }
            QComboBox::drop-down { border: none; width: 22px; }
            QComboBox::down-arrow {
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 5px solid #64748b;
            }
            QComboBox QAbstractItemView {
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                selection-background-color: #c7d2fe;
                selection-color: #1e293b;
                padding: 4px;
                background: white;
            }

            /* ========== 滚动条：细线玻璃 ========== */
            QScrollBar:vertical {
                background: transparent;
                width: 8px;
                border-radius: 4px;
            }
            QScrollBar::handle:vertical {
                background: #cbd5e1;
                border-radius: 4px;
                min-height: 30px;
            }
            QScrollBar::handle:vertical:hover { background: #94a3b8; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
            QScrollBar:horizontal {
                background: transparent;
                height: 8px;
                border-radius: 4px;
            }
            QScrollBar::handle:horizontal {
                background: #cbd5e1;
                border-radius: 4px;
                min-width: 30px;
            }
            QScrollBar::handle:horizontal:hover { background: #94a3b8; }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0px; }

            /* ========== 状态栏 ========== */
            QStatusBar {
                background: #ffffff;
                color: #64748b;
                border-top: 1px solid #e2e8f0;
                font-size: 11px;
            }
            QStatusBar QLabel { color: #64748b; font-size: 11px; }

            /* ========== 菜单 ========== */
            QMenu {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 4px;
            }
            QMenu::item {
                padding: 6px 18px;
                border-radius: 6px;
                color: #334155;
                font-size: 11px;
            }
            QMenu::item:selected { background: #eef2ff; color: #4f46e5; }

            /* ========== 单选框 ========== */
            QRadioButton {
                color: #475569;
                spacing: 6px;
                font-size: 11px;
            }
            QRadioButton::indicator {
                width: 14px;
                height: 14px;
            }
            QRadioButton::indicator:unchecked {
                border: 1px solid #cbd5e1;
                border-radius: 7px;
                background: #ffffff;
            }
            QRadioButton::indicator:checked {
                border: 4px solid #6366f1;
                border-radius: 7px;
                background: #ffffff;
            }

            /* ========== 复选框 ========== */
            QCheckBox {
                color: #475569;
                spacing: 6px;
                font-size: 11px;
            }
            QCheckBox::indicator {
                width: 14px;
                height: 14px;
            }
            QCheckBox::indicator:unchecked {
                border: 1px solid #cbd5e1;
                border-radius: 3px;
                background: #ffffff;
            }
            QCheckBox::indicator:checked {
                border: 1px solid #6366f1;
                border-radius: 3px;
                background: #6366f1;
            }
        """)

        # 状态栏
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("✨ 系统就绪")

        # 中央Tab容器
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 设置窗口图标
        icon_path = os.path.join(get_resource_path("data", "icons"), "monkey-icon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # ========== 品牌顶栏（液态玻璃渐变） ==========
        brand_bar = QWidget()
        brand_bar.setObjectName("brand_bar")
        brand_bar.setStyleSheet("""
            QWidget#brand_bar {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #4f46e5, stop:0.5 #6366f1, stop:1 #4f46e5);
                border-bottom: 1px solid #c7d2fe;
            }
            QLabel { background: transparent; }
        """)
        brand_layout = QHBoxLayout(brand_bar)
        brand_layout.setContentsMargins(20, 10, 20, 10)

        # 左侧：图标 + 标题
        left_layout = QHBoxLayout()
        if os.path.exists(icon_path):
            icon_label = QLabel()
            icon_pixmap = QIcon(icon_path).pixmap(32, 32)
            icon_label.setPixmap(icon_pixmap)
            left_layout.addWidget(icon_label)

        title = QLabel("大圣 · 快递物流出港账单结算系统")
        title.setObjectName("brand_title")
        title.setFont(QFont("Microsoft YaHei", 14, QFont.DemiBold))
        title.setStyleSheet("color: white; padding-left: 8px;")
        left_layout.addWidget(title)

        brand_layout.addLayout(left_layout)
        brand_layout.addStretch()

        # 右侧：小信息
        right_label = QLabel("高效 · 精准 · 智能")
        right_label.setStyleSheet("color: #e0e7ff; font-size: 12px; padding-right: 6px;")
        brand_layout.addWidget(right_label)

        main_layout.addWidget(brand_bar)

        # ========== 内容区（内边距） ==========
        content_container = QWidget()
        content_layout = QVBoxLayout(content_container)
        content_layout.setContentsMargins(18, 12, 18, 8)
        content_layout.setSpacing(8)

        # Tab页
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.setMovable(False)
        content_layout.addWidget(self.tabs)

        # 底部信息栏
        bottom_bar = QWidget()
        bottom_bar.setStyleSheet("""
            background: #f8fafc;
            border-top: 1px solid #e2e8f0;
            padding: 6px 0px;
        """)
        bottom_layout = QHBoxLayout(bottom_bar)
        bottom_layout.setContentsMargins(20, 8, 20, 8)
        bottom_layout.setAlignment(Qt.AlignCenter)

        if os.path.exists(icon_path):
            bottom_icon = QLabel()
            bottom_pixmap = QIcon(icon_path).pixmap(18, 18)
            bottom_icon.setPixmap(bottom_pixmap)
            bottom_layout.addWidget(bottom_icon)

        sub_title = QLabel("杭州喵喵至家网络有限公司  ·  大圣智慧软件  ·  联系电话：17771300068 / 19171045360")
        sub_title.setFont(QFont("Microsoft YaHei", 10))
        sub_title.setStyleSheet("color: #64748b;")
        bottom_layout.addWidget(sub_title)

        content_layout.addWidget(bottom_bar)

        main_layout.addWidget(content_container)

        # 各Tab页
        self.tabs.addTab(self._create_import_tab(), "📁 导入计算")
        self.tabs.addTab(self._create_result_tab(), "📊 计算结果")
        self.tabs.addTab(self._create_settlement_tab(), "💰 多级结算")
        self.tabs.addTab(self._create_history_tab(), "📜 历史记录")
        self.tabs.addTab(self._create_rule_tab(), "⚙️ 规则配置")

    def _create_import_tab(self):
        """导入Tab - 支持最多5个文件批量导入"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # 选择文件区
        file_group = QGroupBox("📂 第一步：选择Excel文件（最多5个）")
        file_layout = QVBoxLayout(file_group)

        # 文件列表
        file_row = QHBoxLayout()
        self.file_list_label = QLabel("未选择文件（提示：可一次选择最多5个文件分批计算）")
        self.file_list_label.setStyleSheet("""
            color: #64748b;
            padding: 10px 12px;
            background: #f8fafc;
            border: 1px dashed #cbd5e1;
            border-radius: 8px;
            font-size: 11px;
        """)
        self.file_list_label.setWordWrap(True)
        self.file_list_label.setMinimumHeight(36)
        file_row.addWidget(self.file_list_label, 1)

        select_btn = QPushButton("📂 选择Excel文件")
        select_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #6366f1, stop:1 #4f46e5);
                color: white;
                border: none;
                padding: 6px 16px;
                border-radius: 8px;
                font-size: 11px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #818cf8, stop:1 #6366f1);
            }
        """)
        select_btn.clicked.connect(self._select_file)
        file_row.addWidget(select_btn)

        file_layout.addLayout(file_row)
        layout.addWidget(file_group)

        # 列名预览区
        preview_group = QGroupBox("🔍 第二步：列名自动匹配结果")
        preview_layout = QVBoxLayout(preview_group)

        self.match_label = QLabel("请先选择Excel文件（按第一个文件自动匹配列名）")
        self.match_label.setWordWrap(True)
        self.match_label.setStyleSheet("""
            padding: 10px 12px;
            color: #334155;
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            font-size: 11px;
        """)
        preview_layout.addWidget(self.match_label)
        layout.addWidget(preview_group)

        # 计算按钮
        self.calc_btn = QPushButton("🚀 开始计算")
        self.calc_btn.setObjectName("primary_green")
        self.calc_btn.clicked.connect(self._start_calculate)
        layout.addWidget(self.calc_btn)

        # 进度条区域
        progress_group = QGroupBox("📊 处理进度")
        progress_layout = QVBoxLayout(progress_group)

        self.progress_stage_label = QLabel("等待开始...")
        self.progress_stage_label.setStyleSheet("color: #4f46e5; font-weight: 500; padding: 3px; font-size: 11px;")
        progress_layout.addWidget(self.progress_stage_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0%")
        self.progress_bar.setVisible(False)
        progress_layout.addWidget(self.progress_bar)

        layout.addWidget(progress_group)

        # 日志
        log_group = QGroupBox("📝 处理日志")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(180)
        self.log_text.setStyleSheet("""
            QTextEdit {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 6px;
                color: #334155;
                font-size: 11px;
            }
        """)
        log_layout.addWidget(self.log_text)
        layout.addWidget(log_group)

        layout.addStretch()

        # Worker引用（防止被GC）
        self.preview_worker = None
        self.calc_worker = None
        # 已选择的文件列表
        self.selected_files = []
        return widget

    def _create_result_tab(self):
        """结果Tab - 液态玻璃风格"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # 顶部工具栏：文件选择 + 汇总信息卡片
        top_bar = QHBoxLayout()

        # 文件选择下拉框
        file_label = QLabel("📁 选择文件:")
        file_label.setStyleSheet("font-weight: 500; color: #4f46e5; font-size: 11px;")
        top_bar.addWidget(file_label)

        self.file_combo = QComboBox()
        self.file_combo.setMinimumWidth(240)
        self.file_combo.currentIndexChanged.connect(self._on_file_combo_changed)
        top_bar.addWidget(self.file_combo)

        top_bar.addStretch()

        # 汇总信息卡片
        info_group = QGroupBox("📊 汇总信息")
        info_group.setStyleSheet("""
            QGroupBox {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 10px;
                padding: 8px;
                color: #334155;
                font-weight: 500;
                font-size: 11px;
            }
        """)
        info_layout = QHBoxLayout(info_group)
        self.summary_labels = {}
        colors = ["#4f46e5", "#10b981", "#ef4444", "#f59e0b"]
        for i, key in enumerate(["总行数", "成功", "异常", "运费总额"]):
            col = QVBoxLayout()
            col.setSpacing(1)
            label = QLabel(key)
            label.setStyleSheet("color: #64748b; font-size: 10px;")
            label.setAlignment(Qt.AlignCenter)
            value = QLabel("0")
            value.setStyleSheet(f"color: {colors[i]}; font-size: 18px; font-weight: 600; padding-top: 1px;")
            value.setAlignment(Qt.AlignCenter)
            self.summary_labels[key] = value
            col.addWidget(label)
            col.addWidget(value)
            info_layout.addLayout(col)
        top_bar.addWidget(info_group)

        layout.addLayout(top_bar)

        # 表格
        self.result_table = QTableWidget()
        self.result_table.setColumnCount(11)
        self.result_table.setHorizontalHeaderLabels([
            "行号", "业务日期", "快递单号", "网点", "区域",
            "重量(kg)", "客户名称", "件数", "运费(元)", "规则", "异常"
        ])
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.result_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.result_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.result_table.setAlternatingRowColors(True)
        self.result_table.verticalHeader().setDefaultSectionSize(30)
        self.result_table.setStyleSheet("""
            QTableWidget {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                gridline-color: #e2e8f0;
                alternate-background-color: #f1f5f9;
            }
            QTableWidget::item { padding: 4px 8px; }
            QTableWidget::item:selected {
                background: #e0e7ff;
                color: #1e293b;
            }
        """)
        layout.addWidget(self.result_table)

        # 导出按钮区域
        export_row = QHBoxLayout()

        # 主导出按钮（带下拉菜单）
        export_btn = QPushButton("📥 导出Excel")
        export_row.addWidget(export_btn)

        # 创建下拉菜单
        from PyQt5.QtWidgets import QMenu
        export_menu = QMenu(export_btn)
        
        act_export_current = export_menu.addAction("导出当前文件")
        act_export_current.triggered.connect(lambda: self._export_details())
        
        act_export_all = export_menu.addAction("分别导出所有文件")
        act_export_all.triggered.connect(lambda: self._export_all_separately())
        
        act_export_merged = export_menu.addAction("合并导出（自动拆分）")
        act_export_merged.triggered.connect(lambda: self._export_merged())
        
        export_btn.setMenu(export_menu)

        export_row.addStretch()
        layout.addLayout(export_row)

        return widget

    def _create_settlement_tab(self):
        """多级结算Tab - 液态玻璃风格"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # 切换按钮（3种结算类型）- 胶囊式
        switch_container = QWidget()
        switch_container.setStyleSheet("""
            background: #f1f5f9;
            border-radius: 10px;
            padding: 6px;
            border: 1px solid #e2e8f0;
        """)
        switch_row = QHBoxLayout(switch_container)

        # 激活按钮样式
        active_style = """
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #6366f1, stop:1 #4f46e5);
                color: white;
                border: none;
                padding: 7px 20px;
                border-radius: 8px;
                font-weight: 500;
                font-size: 11px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #818cf8, stop:1 #6366f1);
            }
        """
        # 未激活样式
        inactive_style = """
            QPushButton {
                background: #ffffff;
                color: #64748b;
                border: 1px solid #e2e8f0;
                padding: 7px 20px;
                border-radius: 8px;
                font-weight: 500;
                font-size: 11px;
            }
            QPushButton:hover {
                background: #eef2ff;
                color: #4f46e5;
                border-color: #c7d2fe;
            }
        """

        self.station_btn = QPushButton("📍 网点结算")
        self.station_btn.setCheckable(True)
        self.station_btn.setChecked(True)
        self.station_btn.setStyleSheet(active_style)
        self.station_btn.clicked.connect(lambda: self._switch_settlement("station"))
        switch_row.addWidget(self.station_btn)

        self.contract_btn = QPushButton("🏢 承包区结算")
        self.contract_btn.setCheckable(True)
        self.contract_btn.setStyleSheet(inactive_style)
        self.contract_btn.clicked.connect(lambda: self._switch_settlement("contract"))
        switch_row.addWidget(self.contract_btn)

        self.monthly_btn = QPushButton("📋 月结客户结算")
        self.monthly_btn.setCheckable(True)
        self.monthly_btn.setStyleSheet(inactive_style)
        self.monthly_btn.clicked.connect(lambda: self._switch_settlement("monthly"))
        switch_row.addWidget(self.monthly_btn)

        switch_row.addStretch()
        layout.addWidget(switch_container)

        # 结算表格
        self.settlement_table = QTableWidget()
        self.settlement_table.setColumnCount(8)
        self.settlement_table.setHorizontalHeaderLabels([
            "编码", "名称", "订单数", "总重量(kg)", "运费(元)",
            "分成比例", "结算金额(元)", "公司收入(元)"
        ])
        self.settlement_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.settlement_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.settlement_table.cellDoubleClicked.connect(self._export_settlement_detail)
        self.settlement_table.setAlternatingRowColors(True)
        self.settlement_table.verticalHeader().setDefaultSectionSize(30)
        self.settlement_table.setStyleSheet("""
            QTableWidget {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                gridline-color: #e2e8f0;
                alternate-background-color: #f1f5f9;
            }
            QTableWidget::item { padding: 4px 8px; }
            QTableWidget::item:selected {
                background: #e0e7ff;
                color: #1e293b;
            }
        """)
        layout.addWidget(self.settlement_table)

        # 导出按钮行
        export_row = QHBoxLayout()
        export_settlement_btn = QPushButton("📥 导出结算汇总")
        export_settlement_btn.clicked.connect(lambda checked=False: self._export_settlement())
        export_row.addWidget(export_settlement_btn)

        export_detail_btn = QPushButton("📄 导出选中明细")
        export_detail_btn.setStyleSheet("""
            QPushButton {
                background: #ffffff;
                color: #475569;
                border: 1px solid #cbd5e1;
                padding: 6px 14px;
                border-radius: 8px;
                font-weight: 500;
                font-size: 11px;
            }
            QPushButton:hover {
                background: #eef2ff;
                color: #4f46e5;
                border-color: #c7d2fe;
            }
        """)
        export_detail_btn.clicked.connect(self._export_settlement_detail)
        export_row.addWidget(export_detail_btn)

        export_row.addStretch()
        layout.addLayout(export_row)
        return widget

    def _create_history_tab(self):
        """历史记录Tab - 液态玻璃风格"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # 按钮行
        btn_row = QHBoxLayout()

        refresh_btn = QPushButton("🔄 刷新历史")
        refresh_btn.setStyleSheet("""
            QPushButton {
                background: #ffffff;
                color: #475569;
                border: 1px solid #cbd5e1;
                padding: 6px 14px;
                border-radius: 8px;
                font-weight: 500;
                font-size: 11px;
            }
            QPushButton:hover {
                background: #eef2ff;
                color: #4f46e5;
                border-color: #c7d2fe;
            }
        """)
        refresh_btn.clicked.connect(self._load_history)
        btn_row.addWidget(refresh_btn)

        clear_btn = QPushButton("🗑️ 清空历史")
        clear_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #f87171, stop:1 #ef4444);
                color: white;
                border: none;
                padding: 6px 14px;
                border-radius: 8px;
                font-weight: 500;
                font-size: 11px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #fca5a5, stop:1 #f87171);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #ef4444, stop:1 #b91c1c);
            }
        """)
        clear_btn.clicked.connect(self._clear_history)
        btn_row.addWidget(clear_btn)

        btn_row.addStretch()
        layout.addLayout(btn_row)

        self.history_table = QTableWidget()
        self.history_table.setColumnCount(6)
        self.history_table.setHorizontalHeaderLabels([
            "ID", "文件名", "总行数", "运费总额", "状态", "创建时间"
        ])
        self.history_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.history_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.history_table.cellDoubleClicked.connect(self._load_history_detail)
        self.history_table.setAlternatingRowColors(True)
        self.history_table.verticalHeader().setDefaultSectionSize(30)
        self.history_table.setStyleSheet("""
            QTableWidget {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                gridline-color: #e2e8f0;
                alternate-background-color: #f1f5f9;
            }
            QTableWidget::item { padding: 4px 8px; }
            QTableWidget::item:selected {
                background: #e0e7ff;
                color: #1e293b;
            }
        """)
        layout.addWidget(self.history_table)

        return widget

    def _create_rule_tab(self):
        """规则配置Tab - 可视化编辑运费规则"""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        # 全局默认设置区
        default_group = QGroupBox("🌐 全局默认设置")
        default_layout = QHBoxLayout(default_group)

        # 默认首重
        default_layout.addWidget(QLabel("默认首重(kg):"))
        self.default_first_weight = QLineEdit("1.0")
        self.default_first_weight.setFixedWidth(80)
        default_layout.addWidget(self.default_first_weight)

        # 默认续重单价
        default_layout.addWidget(QLabel("默认续重单价(元/kg):"))
        self.default_continued_fee = QLineEdit("2.0")
        self.default_continued_fee.setFixedWidth(80)
        default_layout.addWidget(self.default_continued_fee)

        # 默认保底价
        default_layout.addWidget(QLabel("默认保底价(元):"))
        self.default_min_fee = QLineEdit("5.0")
        self.default_min_fee.setFixedWidth(80)
        default_layout.addWidget(self.default_min_fee)

        # 无重量默认价格（订单无重量时的保底结算价）
        default_layout.addWidget(QLabel("无重量默认价(元):"))
        self.default_empty_weight_fee = QLineEdit("3.0")
        self.default_empty_weight_fee.setFixedWidth(80)
        default_layout.addWidget(self.default_empty_weight_fee)

        btn_save_default = QPushButton("💾 保存默认设置")
        btn_save_default.clicked.connect(self._save_default_settings)
        default_layout.addWidget(btn_save_default)

        layout.addWidget(default_group)

        # 主布局：左侧客户列表 + 右侧规则配置
        main_split = QHBoxLayout()

        # 左侧：客户列表
        station_group = QGroupBox("👥 客户列表")
        station_layout = QVBoxLayout(station_group)

        # 模式选择器
        mode_layout = QHBoxLayout()
        mode_layout.addWidget(QLabel("查看模式:"))
        self.rule_mode_combo = QComboBox()
        self.rule_mode_combo.addItems(["客户规则", "区域规则", "全局规则"])
        self.rule_mode_combo.currentIndexChanged.connect(self._on_rule_mode_changed)
        mode_layout.addWidget(self.rule_mode_combo, 1)
        station_layout.addLayout(mode_layout)

        # 客户列表容器
        self.station_list_widget = QWidget()
        station_list_layout = QVBoxLayout(self.station_list_widget)
        station_list_layout.setContentsMargins(0, 0, 0, 0)

        # 客户操作按钮
        station_btn_layout = QHBoxLayout()
        btn_add_station = QPushButton("➕ 新增客户")
        btn_add_station.clicked.connect(self._add_station)
        station_btn_layout.addWidget(btn_add_station)

        btn_del_station = QPushButton("🗑️ 删除客户")
        btn_del_station.clicked.connect(self._delete_station)
        station_btn_layout.addWidget(btn_del_station)

        btn_import_station = QPushButton("📥 批量导入客户+规则")
        btn_import_station.clicked.connect(self._import_stations_and_rules)
        station_btn_layout.addWidget(btn_import_station)

        btn_download_template = QPushButton("📄 下载模板")
        btn_download_template.clicked.connect(self._download_import_template)
        station_btn_layout.addWidget(btn_download_template)
        station_list_layout.addLayout(station_btn_layout)

        # 客户表格
        self.station_table = QTableWidget()
        self.station_table.setColumnCount(3)
        self.station_table.setHorizontalHeaderLabels(["客户编码", "客户名称", "规则类型"])
        self.station_table.horizontalHeader().setStretchLastSection(True)
        self.station_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.station_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.station_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.station_table.setEditTriggers(QAbstractItemView.DoubleClicked)  # 双击才进入编辑
        self.station_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.station_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.station_table.setAlternatingRowColors(True)
        self.station_table.verticalHeader().setDefaultSectionSize(30)
        self.station_table.setStyleSheet("""
            QTableWidget {
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                gridline-color: #e2e8f0;
                alternate-background-color: #f1f5f9;
                font-size: 11px;
            }
            QTableWidget::item { padding: 5px 8px; }
            QTableWidget::item:selected { background: #e0e7ff; color: #1e293b; }
        """)
        self.station_table.cellClicked.connect(self._on_station_selected)
        self.station_table.itemSelectionChanged.connect(
            lambda: self._on_station_selected(self.station_table.currentRow(), 0))
        self.station_table.setMinimumHeight(120)
        station_list_layout.addWidget(self.station_table)
        station_layout.addWidget(self.station_list_widget)

        # 区域列表容器（在区域模式显示）
        self.region_list_widget = QWidget()
        region_list_layout = QVBoxLayout(self.region_list_widget)
        region_list_layout.setContentsMargins(0, 0, 0, 0)

        self.region_table = QTableWidget()
        self.region_table.setColumnCount(2)
        self.region_table.setHorizontalHeaderLabels(["规则名称", "适用区域"])
        self.region_table.horizontalHeader().setStretchLastSection(True)
        self.region_table.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.region_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.region_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.region_table.cellClicked.connect(self._on_region_selected)
        self.region_table.setAlternatingRowColors(True)
        self.region_table.verticalHeader().setDefaultSectionSize(30)
        self.region_table.setMinimumHeight(120)
        region_list_layout.addWidget(self.region_table)
        station_layout.addWidget(self.region_list_widget)
        self.region_list_widget.hide()

        # 全局规则容器
        self.global_widget = QWidget()
        global_layout = QVBoxLayout(self.global_widget)
        global_layout.setContentsMargins(0, 0, 0, 0)
        global_info = QLabel("已选中：全局规则（兜底规则，所有未匹配规则的派件都使用此规则）")
        global_info.setStyleSheet("color: #4f46e5; font-weight: 500; padding: 8px; font-size: 11px;")
        global_info.setWordWrap(True)
        global_layout.addWidget(global_info)
        station_layout.addWidget(self.global_widget)
        self.global_widget.hide()

        main_split.addWidget(station_group, 1)

        # 右侧：选中项的规则配置（整体支持滚动，14区表格完整可见）
        right_scroll = QScrollArea()
        right_scroll.setWidgetResizable(True)
        right_scroll.setFrameShape(QFrame.NoFrame)
        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(2, 2, 2, 2)

        self.config_group = QGroupBox("⚙️ 运费规则配置 - 当前客户：请选择")
        config_layout = QVBoxLayout(self.config_group)

        # 选中项信息（醒目显示当前客户）
        self.selected_info_label = QLabel("👆 请在左侧点击选择一个客户，右侧即显示该客户的运费规则")
        self.selected_info_label.setStyleSheet("""
            QLabel {
                color: #4f46e5;
                background-color: #eef2ff;
                font-weight: 500;
                font-size: 11px;
                padding: 8px 12px;
                border: 1px solid #c7d2fe;
                border-radius: 8px;
            }
        """)
        self.selected_info_label.setWordWrap(True)
        self.selected_info_label.setAlignment(Qt.AlignCenter)
        config_layout.addWidget(self.selected_info_label)

        # 客户规则采用简化分组表（14个组覆盖34省+港澳台）
        self.station_rule_table = QTableWidget()
        self.station_rule_table.setColumnCount(8)
        self.station_rule_table.setHorizontalHeaderLabels(
            ["分组名称", "涵盖省份", "首重费(元)", "续重费(元)", "保底费(元)", "续重单位", "重量进位", "计泡系数"]
        )
        self.station_rule_table.horizontalHeader().setStretchLastSection(False)
        # 为每列设置最小宽度，避免右侧文字被截断
        self.station_rule_table.setColumnWidth(0, 100)
        self.station_rule_table.setColumnWidth(1, 260)
        self.station_rule_table.setColumnWidth(2, 80)
        self.station_rule_table.setColumnWidth(3, 80)
        self.station_rule_table.setColumnWidth(4, 80)
        self.station_rule_table.setColumnWidth(5, 100)
        self.station_rule_table.setColumnWidth(6, 110)
        self.station_rule_table.setColumnWidth(7, 140)
        self.station_rule_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Interactive)
        self.station_rule_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.station_rule_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)
        self.station_rule_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Interactive)
        self.station_rule_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Interactive)
        self.station_rule_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.Interactive)
        self.station_rule_table.horizontalHeader().setSectionResizeMode(6, QHeaderView.Interactive)
        # 最小总宽度，确保所有列完整可见，超出时自动显示横向滚动条
        self.station_rule_table.setMinimumWidth(990)
        self.station_rule_table.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.station_rule_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.station_rule_table.verticalHeader().setDefaultSectionSize(32)
        self.station_rule_table.verticalHeader().setVisible(False)
        self.station_rule_table.setMinimumHeight(440)
        self.station_rule_table.setSizePolicy(
            QSizePolicy.Expanding, QSizePolicy.Expanding
        )
        self.station_rule_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerItem)
        self.station_rule_table.setStyleSheet("""
            QTableWidget {
                alternate-background-color: #f1f5f9;
                background: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                gridline-color: #e2e8f0;
                font-size: 11px;
            }
            QTableWidget::item { padding: 5px 8px; }
            QTableWidget::item:selected { background: #e0e7ff; color: #1e293b; }
        """)

        # 客户规则模式选择（仅客户模式显示）
        self.inherit_mode_row = QHBoxLayout()
        self.btn_use_inherit = QRadioButton("继承区域规则（使用全局区域配置）")
        self.inherit_mode_row.addWidget(self.btn_use_inherit)

        self.btn_use_custom = QRadioButton("使用专属规则（此客户有独立运费标准）")
        self.inherit_mode_row.addWidget(self.btn_use_custom)
        config_layout.addLayout(self.inherit_mode_row)

        # ── 快捷工具栏 ──────────────────────────────────────────
        quick_group = QGroupBox("⚡ 快捷操作（选中行后批量修改，或一键设置所有省份）")
        quick_layout = QHBoxLayout(quick_group)

        quick_layout.addWidget(QLabel("首重:"))
        self.quick_first_fee = QLineEdit()
        self.quick_first_fee.setFixedWidth(60)
        self.quick_first_fee.setPlaceholderText("如 5.0")
        quick_layout.addWidget(self.quick_first_fee)

        quick_layout.addWidget(QLabel("续重:"))
        self.quick_continued_fee = QLineEdit()
        self.quick_continued_fee.setFixedWidth(60)
        self.quick_continued_fee.setPlaceholderText("如 2.0")
        quick_layout.addWidget(self.quick_continued_fee)

        quick_layout.addWidget(QLabel("保底:"))
        self.quick_min_fee = QLineEdit()
        self.quick_min_fee.setFixedWidth(60)
        self.quick_min_fee.setPlaceholderText("如 5.0")
        quick_layout.addWidget(self.quick_min_fee)

        quick_layout.addWidget(QLabel("续重:"))
        self.quick_unit = QComboBox()
        self.quick_unit.addItems(["全续", "百克续"])
        quick_layout.addWidget(self.quick_unit)

        quick_layout.addWidget(QLabel("进位:"))
        self.quick_rounding = QComboBox()
        self.quick_rounding.addItems(["实际重量", "0.5进位", "四舍五入", "向上取整", "分段进位", "进位舍位"])
        quick_layout.addWidget(self.quick_rounding)

        btn_apply_selected = QPushButton("📝 应用到选中行")
        btn_apply_selected.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #6366f1, stop:1 #4f46e5);
                color: white;
                border: none;
                padding: 6px 14px;
                border-radius: 8px;
                font-size: 11px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #818cf8, stop:1 #6366f1);
            }
        """)
        btn_apply_selected.clicked.connect(self._quick_apply_selected)
        quick_layout.addWidget(btn_apply_selected)

        btn_apply_all = QPushButton("🗂️ 应用到所有34省")
        btn_apply_all.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #34d399, stop:1 #10b981);
                color: white;
                border: none;
                padding: 6px 14px;
                border-radius: 8px;
                font-size: 11px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #6ee7b7, stop:1 #34d399);
            }
        """)
        btn_apply_all.clicked.connect(self._quick_apply_all)
        quick_layout.addWidget(btn_apply_all)

        btn_copy_region = QPushButton("📋 从区域规则复制")
        btn_copy_region.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #38bdf8, stop:1 #0284c7);
                color: white;
                border: none;
                padding: 6px 14px;
                border-radius: 8px;
                font-size: 11px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #7dd3fc, stop:1 #38bdf8);
            }
        """)
        btn_copy_region.clicked.connect(self._quick_copy_from_region)
        quick_layout.addWidget(btn_copy_region)

        btn_clear_inputs = QPushButton("清空输入")
        btn_clear_inputs.setStyleSheet("""
            QPushButton {
                background: #ffffff;
                color: #475569;
                border: 1px solid #cbd5e1;
                padding: 6px 14px;
                border-radius: 8px;
                font-size: 11px;
            }
            QPushButton:hover {
                background: #f1f5f9;
                border-color: #94a3b8;
            }
        """)
        btn_clear_inputs.clicked.connect(self._quick_clear_inputs)
        quick_layout.addWidget(btn_clear_inputs)

        quick_layout.addStretch(1)
        config_layout.addWidget(quick_group)

        # 表格加入布局
        config_layout.addWidget(self.station_rule_table, 1)

        # 表格初始化完成后再设置默认选中（此时连接信号，避免提前触发）
        self.btn_use_inherit.setChecked(True)
        self.btn_use_inherit.toggled.connect(self._on_inherit_mode_changed)
        self.btn_use_custom.toggled.connect(self._on_inherit_mode_changed)

        right_layout.addWidget(self.config_group, 2)
        main_split.addWidget(right_scroll, 2)

        layout.addLayout(main_split, 1)

        # 保存按钮
        bottom_btn_layout = QHBoxLayout()
        btn_save_all = QPushButton("💾 保存所有规则")
        btn_save_all.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #6366f1, stop:1 #4f46e5);
                color: white;
                border: none;
                padding: 8px 22px;
                border-radius: 8px;
                font-weight: 500;
                font-size: 11px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #818cf8, stop:1 #6366f1);
            }
        """)
        btn_save_all.clicked.connect(self._save_all_rules)
        bottom_btn_layout.addWidget(btn_save_all)

        btn_reload_all = QPushButton("🔄 重新加载")
        btn_reload_all.setStyleSheet("""
            QPushButton {
                background: #ffffff;
                color: #475569;
                border: 1px solid #cbd5e1;
                padding: 8px 22px;
                border-radius: 8px;
                font-weight: 500;
                font-size: 11px;
            }
            QPushButton:hover {
                background: #f1f5f9;
                border-color: #94a3b8;
            }
        """)
        btn_reload_all.clicked.connect(self._reload_all_rules)
        bottom_btn_layout.addWidget(btn_reload_all)

        right_layout.addLayout(bottom_btn_layout)

        # 活动加价规则（双十一、618等大促期间加价）
        promotion_group = QGroupBox("🎁 活动加价规则")
        promo_layout = QVBoxLayout(promotion_group)
        promo_info = QLabel("支持设置多条活动加价规则。计算时自动检测当前日期是否在活动期间，按规则加价。")
        promo_info.setStyleSheet("color: #64748b; padding: 2px; font-size: 11px;")
        promo_layout.addWidget(promo_info)

        # 活动规则表格
        self.promotion_table = QTableWidget()
        self.promotion_table.setColumnCount(6)
        self.promotion_table.setHorizontalHeaderLabels(["活动名称", "开始日期", "结束日期", "加价类型", "加价值", "限定省份(逗号分隔)"])
        self.promotion_table.horizontalHeader().setStretchLastSection(True)
        self.promotion_table.setEditTriggers(QAbstractItemView.AllEditTriggers)
        self.promotion_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.promotion_table.setMinimumHeight(80)
        self.promotion_table.setMaximumHeight(130)
        self.promotion_table.verticalHeader().setDefaultSectionSize(28)
        self.promotion_table.horizontalHeader().setFixedHeight(26)
        promo_layout.addWidget(self.promotion_table)

        # 活动操作按钮
        promo_btn_layout = QHBoxLayout()
        btn_add_promo = QPushButton("➕ 添加活动")
        btn_add_promo.clicked.connect(self._add_promotion)
        promo_btn_layout.addWidget(btn_add_promo)
        btn_del_promo = QPushButton("🗑️ 删除活动")
        btn_del_promo.clicked.connect(self._delete_promotion)
        promo_btn_layout.addWidget(btn_del_promo)
        promo_hint = QLabel("加价类型可选：fixed(每单加X元) / weight(每kg加X元) / percent(加X%)")
        promo_hint.setStyleSheet("color: #64748b; font-size: 11px;")
        promo_btn_layout.addWidget(promo_hint, 1)
        promo_layout.addLayout(promo_btn_layout)
        right_layout.addWidget(promotion_group)

        # 测试区
        test_group = QGroupBox("🧪 快速测试")
        test_layout = QHBoxLayout(test_group)
        test_layout.addWidget(QLabel("重量(kg):"))
        self.test_weight = QLineEdit("1.5")
        self.test_weight.setFixedWidth(70)
        test_layout.addWidget(self.test_weight)
        test_layout.addWidget(QLabel("客户编码:"))
        self.test_station = QLineEdit("")
        self.test_station.setPlaceholderText("如 C001")
        self.test_station.setFixedWidth(100)
        test_layout.addWidget(self.test_station)
        test_layout.addWidget(QLabel("区域:"))
        self.test_region = QComboBox()
        # 填充所有省份和港澳台（按分组顺序）
        for group_name, provinces, *_ in PROVINCE_GROUPS:
            for p in provinces:
                self.test_region.addItem(p)
        self.test_region.setEditable(True)  # 允许用户手动输入
        self.test_region.setFixedWidth(120)
        test_layout.addWidget(self.test_region, 1)
        btn_test = QPushButton("计算")
        btn_test.clicked.connect(self._rule_test)
        test_layout.addWidget(btn_test)
        self.test_result = QLabel("结果：—")
        self.test_result.setStyleSheet("color: #10b981; font-weight: 500; font-size: 11px;")
        test_layout.addWidget(self.test_result, 1)
        right_layout.addWidget(test_group)

        right_layout.addStretch(1)
        right_scroll.setWidget(right_container)

        self._reload_all_rules()
        return widget

    # --- 网点规则配置相关函数 ---

    # 34个省份默认数据 (省份, 首重费, 续重费, 保底费, 所属区域)
    def _get_province_defaults(self):
        return [
            ("北京", 4.0, 1.8, 4.0, "华北地区"),
            ("天津", 4.0, 1.8, 4.0, "华北地区"),
            ("上海", 3.5, 1.5, 3.5, "华东地区"),
            ("重庆", 5.0, 2.5, 5.0, "西南地区"),
            ("江苏", 3.5, 1.5, 3.5, "华东地区"),
            ("浙江", 3.5, 1.5, 3.5, "华东地区"),
            ("安徽", 3.5, 1.5, 3.5, "华东地区"),
            ("福建", 3.5, 1.5, 3.5, "华东地区"),
            ("江西", 3.5, 1.5, 3.5, "华东地区"),
            ("山东", 3.5, 1.5, 3.5, "华东地区"),
            ("河北", 4.0, 1.8, 4.0, "华北地区"),
            ("山西", 4.0, 1.8, 4.0, "华北地区"),
            ("内蒙古", 4.0, 1.8, 4.0, "华北地区"),
            ("广东", 4.0, 1.8, 4.0, "华南地区"),
            ("广西", 4.0, 1.8, 4.0, "华南地区"),
            ("海南", 4.0, 1.8, 4.0, "华南地区"),
            ("河南", 4.5, 2.0, 4.5, "华中地区"),
            ("湖北", 4.5, 2.0, 4.5, "华中地区"),
            ("湖南", 4.5, 2.0, 4.5, "华中地区"),
            ("四川", 5.0, 2.5, 5.0, "西南地区"),
            ("贵州", 5.0, 2.5, 5.0, "西南地区"),
            ("云南", 5.0, 2.5, 5.0, "西南地区"),
            ("西藏", 5.0, 2.5, 5.0, "西南地区"),
            ("陕西", 8.0, 4.0, 8.0, "西北地区"),
            ("甘肃", 8.0, 4.0, 8.0, "西北地区"),
            ("青海", 8.0, 4.0, 8.0, "西北地区"),
            ("宁夏", 8.0, 4.0, 8.0, "西北地区"),
            ("新疆", 8.0, 4.0, 8.0, "西北地区"),
            ("辽宁", 5.0, 2.5, 5.0, "东北地区"),
            ("吉林", 5.0, 2.5, 5.0, "东北地区"),
            ("黑龙江", 5.0, 2.5, 5.0, "东北地区"),
            ("香港", 30.0, 20.0, 30.0, "港澳台"),
            ("澳门", 30.0, 20.0, 30.0, "港澳台"),
            ("台湾", 30.0, 20.0, 30.0, "港澳台"),
        ]

    def _reload_all_rules(self):
        """加载所有规则并初始化界面"""
        try:
            service = RuleService()
            rules = service.load_rules()

            station_rules = [r for r in rules if r.rule_type == "station"]
            region_rules = [r for r in rules if r.rule_type == "region"]
            global_rules = [r for r in rules if r.rule_type == "global"]

            self._populate_station_table(station_rules)
            self._populate_region_table(region_rules)
            self._current_global_rules = global_rules

            # 构建 region 费用映射：{province: (first_fee, continued_fee, min_fee, region_name, continued_unit, weight_rounding)}
            self._region_fee_map = {}
            for r in region_rules:
                province = r.regions.strip() if r.regions else ""
                if province:
                    region_name = ""
                    if " - " in r.name:
                        parts = r.name.split(" - ")
                        if len(parts) >= 2:
                            region_name = parts[-1].strip()
                    self._region_fee_map[province] = (
                        r.first_fee, r.continued_fee, r.min_fee, region_name, r.continued_unit, r.weight_rounding
                    )

            # 构建客户专属规则缓存：{customer_code: {province: (first_fee, continued_fee, min_fee, continued_unit, weight_rounding)}}
            self._station_province_cache = {}
            for r in station_rules:
                if r.stations and r.stations.strip():
                    stations_str = r.stations.strip()
                    code = stations_str.split(",")[0].strip()
                    province = r.regions.strip() if r.regions else ""
                    if code not in self._station_province_cache:
                        self._station_province_cache[code] = {}
                    if province:
                        self._station_province_cache[code][province] = (
                            r.first_fee, r.continued_fee, r.min_fee, r.continued_unit, r.weight_rounding,
                            getattr(r, '计泡系数', 6000.0)
                        )

            # 判断每个客户是"继承区域"还是"专属规则"
            for row in range(self.station_table.rowCount()):
                code_item = self.station_table.item(row, 0)
                if code_item:
                    code = code_item.text().strip()
                    if code in self._station_province_cache and self._station_province_cache[code]:
                        self.station_table.item(row, 2).setText("专属规则")
                    else:
                        self.station_table.item(row, 2).setText("继承区域")

            if self.station_table.rowCount() > 0:
                self.station_table.selectRow(0)
                self._on_station_selected(0, 0)

            self.rule_mode_combo.setCurrentIndex(0)

            # 加载活动加价规则
            self.promotion_table.setRowCount(0)
            try:
                config_file = get_config_file("fee_rules.json")
                if os.path.exists(config_file):
                    with open(config_file, "r", encoding="utf-8") as f:
                        data = json.load(f)
                    promo_rules = data.get("promotion_rules", [])
                    for pr in promo_rules:
                        row = self.promotion_table.rowCount()
                        self.promotion_table.insertRow(row)
                        self.promotion_table.setItem(row, 0, QTableWidgetItem(str(pr.get("name", ""))))
                        self.promotion_table.setItem(row, 1, QTableWidgetItem(str(pr.get("start_date", ""))))
                        self.promotion_table.setItem(row, 2, QTableWidgetItem(str(pr.get("end_date", ""))))
                        self.promotion_table.setItem(row, 3, QTableWidgetItem(str(pr.get("markup_type", "percent"))))
                        self.promotion_table.setItem(row, 4, QTableWidgetItem(str(pr.get("markup_value", "0"))))
                        self.promotion_table.setItem(row, 5, QTableWidgetItem(str(pr.get("regions", ""))))
            except Exception:
                pass
        except Exception as e:
            import traceback
            traceback.print_exc()
            QMessageBox.warning(self, "加载失败", f"加载规则失败：{e}")

    def _populate_station_table(self, station_rules):
        """填充客户表格"""
        default_stations = [
            ("C001", "默认客户"),
            ("C002", "电商VIP客户"),
            ("C003", "月结客户"),
        ]

        # 合并已有规则和默认客户（解析stations="编码,名称"格式）
        all_stations = {}
        for r in station_rules:
            if r.stations and r.stations.strip():
                parts = [p.strip() for p in r.stations.split(",") if p.strip()]
                code = parts[0]
                name = parts[1] if len(parts) > 1 else (r.name or code)
                all_stations[code] = name

        for code, name in default_stations:
            if code not in all_stations:
                all_stations[code] = name

        if not all_stations:
            all_stations["C001"] = "示例客户"

        self.station_table.setRowCount(0)
        for code in sorted(all_stations.keys()):
            # 获取客户名称：从已解析的all_stations取
            name = all_stations[code]
            if not name:
                for r in station_rules:
                    parts = [p.strip() for p in (r.stations or "").split(",") if p and p.strip()]
                    if parts and parts[0] == code and r.name and code not in dict(default_stations):
                        name = r.name
                        break
            if not name:
                name_map = dict(default_stations)
                name = name_map.get(code, code)

            row = self.station_table.rowCount()
            self.station_table.insertRow(row)
            self.station_table.setItem(row, 0, QTableWidgetItem(code))
            self.station_table.setItem(row, 1, QTableWidgetItem(name))
            self.station_table.setItem(row, 2, QTableWidgetItem("继承区域"))

    def _populate_region_table(self, region_rules):
        """填充区域规则表格"""
        self.region_table.setRowCount(len(region_rules))
        for i, r in enumerate(region_rules):
            self.region_table.setItem(i, 0, QTableWidgetItem(r.name or "未命名"))
            self.region_table.setItem(i, 1, QTableWidgetItem(r.regions or ""))
            self.region_table.item(i, 0).setData(Qt.UserRole, r)

    def _on_rule_mode_changed(self, index):
        """切换查看模式"""
        mode = self.rule_mode_combo.currentText()

        if mode == "客户规则":
            self.station_list_widget.show()
            self.region_list_widget.hide()
            self.global_widget.hide()
            # 显示继承模式选择行
            for i in range(self.inherit_mode_row.count()):
                item = self.inherit_mode_row.itemAt(i)
                if item and item.widget():
                    item.widget().setVisible(True)
            if self.station_table.rowCount() > 0:
                self.station_table.selectRow(0)
                self._on_station_selected(0, 0)
        elif mode == "区域规则":
            self.station_list_widget.hide()
            self.region_list_widget.show()
            self.global_widget.hide()
            for i in range(self.inherit_mode_row.count()):
                item = self.inherit_mode_row.itemAt(i)
                if item and item.widget():
                    item.widget().setVisible(False)
            if self.region_table.rowCount() > 0:
                self.region_table.selectRow(0)
                self._on_region_selected(0, 0)
        elif mode == "全局规则":
            self.station_list_widget.hide()
            self.region_list_widget.hide()
            self.global_widget.show()
            for i in range(self.inherit_mode_row.count()):
                item = self.inherit_mode_row.itemAt(i)
                if item and item.widget():
                    item.widget().setVisible(False)
            self._load_global_rule()

    def _on_station_selected(self, row, col):
        """选中网点后加载所有省份规则（切换前自动保存当前客户编辑 + 从文件重新加载最新规则）"""
        try:
            if row < 0:
                return

            code_item = self.station_table.item(row, 0)
            name_item = self.station_table.item(row, 1)
            type_item = self.station_table.item(row, 2)
            if not code_item:
                return

            station_code = code_item.text().strip()
            station_name = name_item.text().strip() if name_item else station_code
            rule_type = type_item.text().strip() if type_item else "继承区域"

            # 切换客户前，先保存当前客户的编辑到缓存（避免编辑丢失）
            if hasattr(self, '_current_station_code_editing') and self._current_station_code_editing:
                if self._current_station_code_editing != station_code:
                    try:
                        self._save_current_station_to_cache()
                    except Exception as save_err:
                        import logging
                        logging.warning(f"保存当前客户编辑到缓存失败: {save_err}")

            self._current_station_code_editing = station_code

            # 从 fee_rules.json 重新加载最新规则（确保切换时看到的数据是最新的）
            try:
                self._reload_rules_from_file_only()
            except Exception as reload_err:
                import logging
                logging.warning(f"从文件重新加载规则时出错（不影响使用，继续使用内存缓存）: {reload_err}")

            # 根据最新缓存决定规则类型显示
            if station_code in self._station_province_cache and self._station_province_cache[station_code]:
                rule_type = "专属规则"
                if type_item:
                    type_item.setText("专属规则")
            else:
                rule_type = "继承区域"
                if type_item:
                    type_item.setText("继承区域")

            # 更新右侧标题栏和选中信息标签，让用户清楚知道当前编辑的是哪个客户
            try:
                self.config_group.setTitle(f"⚙️ 运费规则配置 - 当前客户：{station_name}（{station_code}）")
                mode_icon = "⭐" if rule_type == "专属规则" else "📋"
                self.selected_info_label.setText(f"{mode_icon} 正在编辑：{station_name}（{station_code}） | 规则模式：{rule_type}")
            except Exception:
                pass

            # 设置单选按钮
            try:
                if rule_type == "专属规则":
                    self.btn_use_custom.setChecked(True)
                else:
                    self.btn_use_inherit.setChecked(True)
            except Exception:
                pass

            # 填充14个分组规则表格（覆盖34省+港澳台）
            self._populate_province_table(station_code)

        except Exception as e:
            import logging
            import traceback
            logging.error(f"切换客户时出错: {e}\n{traceback.format_exc()}")
            # 出错时不要崩溃，尝试重置为默认状态
            try:
                if hasattr(self, 'station_rule_table'):
                    self.station_rule_table.setRowCount(0)
                from PyQt5.QtWidgets import QMessageBox
                QMessageBox.warning(self, "提示", f"加载客户规则时出错: {e}")
            except Exception:
                pass

    def _reload_rules_from_file_only(self):
        """仅从 fee_rules.json 文件重新加载规则数据到内存缓存（不影响当前选中行，也不重置UI）"""
        try:
            # 强制刷新 calculate_service 中的全局规则缓存（包括活动加价规则）
            from app.services.calculate_service import _build_rule_indexes
            _build_rule_indexes(force_reload=True)

            service = RuleService()
            rules = service.load_rules()

            # 构建最新的 region_fee_map（用于默认值）
            new_region_map = {}
            for r in rules:
                if r.rule_type == "region":
                    province = r.regions.strip() if r.regions else ""
                    if province:
                        region_name = ""
                        if " - " in r.name:
                            parts = r.name.split(" - ")
                            if len(parts) >= 2:
                                region_name = parts[-1].strip()
                        # 兼容旧数据（可能没有continued_unit/rounding字段）
                        continued_unit = getattr(r, 'continued_unit', 'kg') or 'kg'
                        weight_rounding = getattr(r, 'weight_rounding', 'actual') or 'actual'
                        new_region_map[province] = (
                            r.first_fee, r.continued_fee, r.min_fee,
                            region_name, continued_unit, weight_rounding
                        )
            self._region_fee_map = new_region_map

            # 构建最新的 station_province_cache（仅包含"专属规则"客户）
            new_station_cache = {}
            for r in rules:
                if r.rule_type == "station" and r.stations and r.stations.strip():
                    stations_str = r.stations.strip()
                    # stations 格式为 "编码,名称"
                    parts = [p.strip() for p in stations_str.split(",") if p.strip()]
                    code = parts[0] if parts else ""
                    province = r.regions.strip() if r.regions else ""
                    if not code or not province:
                        continue
                    if r.first_fee <= 0 and r.continued_fee <= 0 and r.min_fee <= 0:
                        # 这是继承区域的标记规则（空费用值），跳过
                        continue
                    continued_unit = getattr(r, 'continued_unit', 'kg') or 'kg'
                    weight_rounding = getattr(r, 'weight_rounding', 'actual') or 'actual'
                    if code not in new_station_cache:
                        new_station_cache[code] = {}
                    new_station_cache[code][province] = (
                        r.first_fee, r.continued_fee, r.min_fee,
                        continued_unit, weight_rounding
                    )
            self._station_province_cache = new_station_cache

            # 重新加载活动加价规则到 UI 表格（包含省份限定）
            try:
                promo_rules = service.load_promotion_rules()
                if hasattr(self, 'promotion_table'):
                    self.promotion_table.setRowCount(0)
                    for pr in promo_rules:
                        row = self.promotion_table.rowCount()
                        self.promotion_table.insertRow(row)
                        self.promotion_table.setItem(row, 0, QTableWidgetItem(str(pr.get("name", ""))))
                        self.promotion_table.setItem(row, 1, QTableWidgetItem(str(pr.get("start_date", ""))))
                        self.promotion_table.setItem(row, 2, QTableWidgetItem(str(pr.get("end_date", ""))))
                        self.promotion_table.setItem(row, 3, QTableWidgetItem(str(pr.get("markup_type", "percent"))))
                        self.promotion_table.setItem(row, 4, QTableWidgetItem(str(pr.get("markup_value", "0"))))
                        self.promotion_table.setItem(row, 5, QTableWidgetItem(str(pr.get("regions", ""))))
            except Exception:
                pass

        except Exception as e:
            import logging
            import traceback
            logging.warning(f"_reload_rules_from_file_only 异常（不影响使用）: {e}\n{traceback.format_exc()}")
            # 出错时至少确保缓存存在且可用
            if not hasattr(self, '_station_province_cache') or self._station_province_cache is None:
                self._station_province_cache = {}
            if not hasattr(self, '_region_fee_map') or self._region_fee_map is None:
                self._region_fee_map = {}

    ROUNDING_MODE_MAP = {
        "实际重量": "actual",
        "0.5进位": "round_05",
        "四舍五入": "round_1",
        "向上取整": "ceil_1kg",
        "分段进位": "segmented",
        "进位舍位": "round_trunc"
    }
    ROUNDING_MODE_REVERSE = {v: k for k, v in ROUNDING_MODE_MAP.items()}

    def _populate_province_table(self, station_code):
        """填充选中客户的分组规则表格（14组覆盖34省+港澳台）- 带完整异常保护"""
        try:
            if not hasattr(self, '_station_province_cache') or self._station_province_cache is None:
                self._station_province_cache = {}
            cached_data = self._station_province_cache.get(station_code, {})
            if not hasattr(self, '_region_fee_map') or self._region_fee_map is None:
                self._region_fee_map = {}

            self.station_rule_table.setRowCount(0)
        except Exception as init_err:
            import logging
            logging.warning(f"_populate_province_table 初始化异常: {init_err}")
            # 出错时至少确保表格是空的
            try:
                self.station_rule_table.setRowCount(0)
            except Exception:
                pass
            return

        for (gname, provinces, def_first, def_continued, def_min, def_unit, def_rounding) in PROVINCE_GROUPS:
            try:
                # 取该分组内省份的已有客户专属值（如果存在）
                first_fee = def_first
                continued_fee = def_continued
                min_fee = def_min
                continued_unit = def_unit
                weight_rounding = def_rounding

                # 优先使用客户已有的规则（取该组内第一个省份的值为代表）
                for province in provinces:
                    if province in cached_data:
                        data = cached_data[province]
                        if len(data) >= 5:
                            first_fee, continued_fee, min_fee, continued_unit, weight_rounding = data
                        elif len(data) == 4:
                            first_fee, continued_fee, min_fee, continued_unit = data
                        else:
                            first_fee, continued_fee, min_fee = data
                        break
                else:
                    # 组内省份都没有客户专属值，尝试取区域规则值
                    for province in provinces:
                        if province in self._region_fee_map:
                            r_data = self._region_fee_map[province]
                            if len(r_data) >= 6:
                                r_first, r_continued, r_min, _, r_unit, r_rounding = r_data
                            elif len(r_data) >= 5:
                                r_first, r_continued, r_min, _, r_unit = r_data
                                r_rounding = "actual"
                            else:
                                r_first, r_continued, r_min = r_data[:3]
                                r_unit = "kg"
                                r_rounding = "actual"
                            first_fee, continued_fee, min_fee = r_first, r_continued, r_min
                            continued_unit = r_unit
                            weight_rounding = r_rounding
                            break

                if continued_unit == "100g":
                    weight_rounding = "actual"

                # 确保数值有效
                try:
                    first_fee = float(first_fee)
                    continued_fee = float(continued_fee)
                    min_fee = float(min_fee)
                except (ValueError, TypeError):
                    first_fee = float(def_first)
                    continued_fee = float(def_continued)
                    min_fee = float(def_min)

                row = self.station_rule_table.rowCount()
                self.station_rule_table.insertRow(row)

                # 分组名称（只读）
                name_item = QTableWidgetItem(gname)
                name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
                name_item.setToolTip("、".join(provinces))
                self.station_rule_table.setItem(row, 0, name_item)

                # 涵盖省份（只读）
                prov_text = "、".join(provinces)
                prov_item = QTableWidgetItem(prov_text)
                prov_item.setFlags(prov_item.flags() & ~Qt.ItemIsEditable)
                prov_item.setToolTip("、".join(provinces))
                self.station_rule_table.setItem(row, 1, prov_item)

                self.station_rule_table.setItem(row, 2, QTableWidgetItem(f"{first_fee:.2f}"))
                self.station_rule_table.setItem(row, 3, QTableWidgetItem(f"{continued_fee:.2f}"))
                self.station_rule_table.setItem(row, 4, QTableWidgetItem(f"{min_fee:.2f}"))

                unit_combo = QComboBox()
                unit_combo.addItems(["全续", "百克续"])
                unit_combo.setMinimumHeight(26)
                if continued_unit == "100g":
                    unit_combo.setCurrentText("百克续")
                else:
                    unit_combo.setCurrentText("全续")
                unit_combo.currentTextChanged.connect(lambda text, r=row: self._on_unit_changed(r, text))
                self.station_rule_table.setCellWidget(row, 5, unit_combo)

                rounding_combo = QComboBox()
                rounding_combo.addItems(list(self.ROUNDING_MODE_MAP.keys()))
                rounding_combo.setMinimumHeight(26)
                rounding_text = self.ROUNDING_MODE_REVERSE.get(weight_rounding, "实际重量")
                rounding_combo.setCurrentText(rounding_text)
                rounding_combo.currentTextChanged.connect(lambda text, r=row: self._on_rounding_changed(r, text))
                if continued_unit == "100g":
                    rounding_combo.setEnabled(False)
                self.station_rule_table.setCellWidget(row, 6, rounding_combo)

                # 计泡系统 QComboBox（列7）
                vol_div_combo = QComboBox()
                vol_div_combo.addItems(["6000（顺丰/京东/德邦）", "8000（圆通/中通/韵达）", "5000（EMS）", "12000（大件轻抛）"])
                vol_div_combo.setMinimumHeight(26)
                # 从缓存读取计泡系数，找不到则默认6000
                vol_div_value = 6000
                for province in provinces:
                    if province in cached_data:
                        data = cached_data[province]
                        if len(data) >= 6:
                            vol_div_value = data[5]
                            break
                vol_div_text = f"{int(vol_div_value)}（" + ["6000（顺丰/京东/德邦）", "8000（圆通/中通/韵达）", "5000（EMS）", "12000（大件轻抛）"][
                    {6000: 0, 8000: 1, 5000: 2, 12000: 3}.get(vol_div_value, 0)]
                vol_div_combo.setCurrentText(vol_div_text)
                self.station_rule_table.setCellWidget(row, 7, vol_div_combo)

            except Exception as row_err:
                import logging
                logging.warning(f"填充分组 [{gname}] 时出错: {row_err}")
                # 跳过当前分组，继续处理下一个
                continue

        try:
            self._on_inherit_mode_changed()
        except Exception:
            pass

    def _on_unit_changed(self, row, unit_text):
        """续重单位变化时的联动逻辑：百克续→强制实际重量"""
        # 分组表格：列5=续重单位(ComboBox), 列6=重量进位(ComboBox)
        rounding_combo = self.station_rule_table.cellWidget(row, 6)
        if rounding_combo and isinstance(rounding_combo, QComboBox):
            if unit_text == "百克续":
                rounding_combo.setCurrentText("实际重量")
                rounding_combo.setEnabled(False)
            else:
                rounding_combo.setEnabled(True)

    def _on_rounding_changed(self, row, rounding_text):
        """重量进位变化时的联动逻辑：向上取整/进位舍位→强制全续"""
        unit_combo = self.station_rule_table.cellWidget(row, 5)
        if unit_combo and isinstance(unit_combo, QComboBox):
            rounding_code = self.ROUNDING_MODE_MAP.get(rounding_text, "actual")
            if rounding_code in ["ceil_1kg", "round_trunc"]:
                unit_combo.setCurrentText("全续")

    # ── 快捷操作方法 ──────────────────────────────────────────────
    def _quick_clear_inputs(self):
        """清空快捷工具栏的输入框"""
        self.quick_first_fee.clear()
        self.quick_continued_fee.clear()
        self.quick_min_fee.clear()

    def _apply_quick_values_to_row(self, row):
        """把快捷工具栏输入的值应用到指定行（只应用非空的字段）"""
        # 分组表格：0=分组名 1=涵盖省 2=首重 3=续重 4=保底 5=续重单位(ComboBox) 6=重量进位(ComboBox)
        first_val = self.quick_first_fee.text().strip()
        if first_val:
            try:
                float(first_val)
                self.station_rule_table.setItem(row, 2, QTableWidgetItem(f"{float(first_val):.2f}"))
            except ValueError:
                pass
        continued_val = self.quick_continued_fee.text().strip()
        if continued_val:
            try:
                float(continued_val)
                self.station_rule_table.setItem(row, 3, QTableWidgetItem(f"{float(continued_val):.2f}"))
            except ValueError:
                pass
        min_val = self.quick_min_fee.text().strip()
        if min_val:
            try:
                float(min_val)
                self.station_rule_table.setItem(row, 4, QTableWidgetItem(f"{float(min_val):.2f}"))
            except ValueError:
                pass
        # 续重单位
        unit_combo = self.station_rule_table.cellWidget(row, 5)
        if unit_combo and isinstance(unit_combo, QComboBox):
            unit_combo.setCurrentText(self.quick_unit.currentText())
        # 重量进位
        rounding_combo = self.station_rule_table.cellWidget(row, 6)
        if rounding_combo and isinstance(rounding_combo, QComboBox):
            rounding_combo.setCurrentText(self.quick_rounding.currentText())
        # 联动：百克续时禁用进位
        if self.quick_unit.currentText() == "百克续":
            r_combo = self.station_rule_table.cellWidget(row, 6)
            if r_combo and isinstance(r_combo, QComboBox):
                r_combo.setCurrentText("实际重量")
                r_combo.setEnabled(False)
        else:
            r_combo = self.station_rule_table.cellWidget(row, 6)
            if r_combo and isinstance(r_combo, QComboBox):
                r_combo.setEnabled(True)
        # 联动：向上取整/进位舍位时用全续
        rounding_code = self.ROUNDING_MODE_MAP.get(self.quick_rounding.currentText(), "actual")
        if rounding_code in ["ceil_1kg", "round_trunc"]:
            u_combo = self.station_rule_table.cellWidget(row, 5)
            if u_combo and isinstance(u_combo, QComboBox):
                u_combo.setCurrentText("全续")

    def _quick_apply_selected(self):
        """应用到选中行（支持Ctrl+点击多选）"""
        rows = set()
        for item in self.station_rule_table.selectedItems():
            rows.add(item.row())
        if not rows:
            QMessageBox.information(self, "提示", "请先在下方表格中选中要修改的省份行（按住Ctrl可多选）")
            return
        for row in rows:
            self._apply_quick_values_to_row(row)

    def _quick_apply_all(self):
        """应用到所有34个省份"""
        if self.station_rule_table.rowCount() == 0:
            QMessageBox.information(self, "提示", "请先选择一个客户")
            return
        for row in range(self.station_rule_table.rowCount()):
            self._apply_quick_values_to_row(row)

    def _quick_copy_from_region(self):
        """从区域规则复制到当前客户所有省份（按分组取区域内省份的区域规则值）"""
        if not hasattr(self, '_region_fee_map') or not self._region_fee_map:
            QMessageBox.information(self, "提示", "还没有配置区域规则")
            return
        if self.station_rule_table.rowCount() == 0:
            QMessageBox.information(self, "提示", "请先选择一个客户")
            return
        count = 0
        for row in range(self.station_rule_table.rowCount()):
            gname_item = self.station_rule_table.item(row, 0)
            if not gname_item:
                continue
            group_name = gname_item.text()
            # 找该组对应的省份列表
            target_provinces = []
            for (gn, provs, *_) in PROVINCE_GROUPS:
                if gn == group_name:
                    target_provinces = provs
                    break
            # 在组内省份里找第一个有区域规则的省份作为代表
            src_first = src_continued = src_min = None
            src_unit = "kg"
            src_rounding = "actual"
            for province in target_provinces:
                if province in self._region_fee_map:
                    r_data = self._region_fee_map[province]
                    if len(r_data) >= 6:
                        r_first, r_continued, r_min, _, r_unit, r_rounding = r_data
                    elif len(r_data) >= 5:
                        r_first, r_continued, r_min, _, r_unit = r_data
                        r_rounding = "actual"
                    else:
                        r_first, r_continued, r_min = r_data[:3]
                        r_unit = "kg"
                        r_rounding = "actual"
                    src_first, src_continued, src_min = r_first, r_continued, r_min
                    src_unit, src_rounding = r_unit, r_rounding
                    break
            if src_first is None:
                continue
            self.station_rule_table.setItem(row, 2, QTableWidgetItem(f"{src_first:.2f}"))
            self.station_rule_table.setItem(row, 3, QTableWidgetItem(f"{src_continued:.2f}"))
            self.station_rule_table.setItem(row, 4, QTableWidgetItem(f"{src_min:.2f}"))
            u_combo = self.station_rule_table.cellWidget(row, 5)
            if u_combo and isinstance(u_combo, QComboBox):
                u_combo.setCurrentText("百克续" if src_unit == "100g" else "全续")
            r_combo = self.station_rule_table.cellWidget(row, 6)
            if r_combo and isinstance(r_combo, QComboBox):
                r_combo.setCurrentText(self.ROUNDING_MODE_REVERSE.get(src_rounding, "实际重量"))
            count += 1
        QMessageBox.information(self, "完成", f"已从区域规则复制 {count} 个分组的值")

    def _on_inherit_mode_changed(self):
        """继承模式变化 - 控制表格是否可编辑"""
        if not hasattr(self, 'station_rule_table') or self.station_rule_table is None:
            return
        if not hasattr(self, 'btn_use_custom') or not hasattr(self, 'btn_use_inherit'):
            return
        if self.btn_use_custom.isChecked():
            # 专属规则 - 允许编辑费用列
            for row in range(self.station_rule_table.rowCount()):
                for col in [1, 2, 3]:
                    item = self.station_rule_table.item(row, col)
                    if item:
                        item.setFlags(item.flags() | Qt.ItemIsEditable)
        else:
            # 继承区域 - 费用列只读，显示浅灰色
            for row in range(self.station_rule_table.rowCount()):
                for col in [1, 2, 3]:
                    item = self.station_rule_table.item(row, col)
                    if item:
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)

    def _on_region_selected(self, row, col):
        """选中区域规则"""
        if row < 0:
            return
        name_item = self.region_table.item(row, 0)
        if not name_item:
            return
        self.config_group.setTitle(f"📋 运费规则配置 - 当前区域规则：{name_item.text()}")
        self.selected_info_label.setText(f"📋 正在编辑区域规则：{name_item.text()}")
        rule = name_item.data(Qt.UserRole)

        self.station_rule_table.setRowCount(0)
        row_idx = self.station_rule_table.rowCount()
        self.station_rule_table.insertRow(row_idx)

        item1 = QTableWidgetItem(rule.regions or "区域")
        item1.setFlags(item1.flags() & ~Qt.ItemIsEditable)
        self.station_rule_table.setItem(row_idx, 0, item1)

        self.station_rule_table.setItem(row_idx, 1, QTableWidgetItem(f"{rule.first_fee:.2f}"))
        self.station_rule_table.setItem(row_idx, 2, QTableWidgetItem(f"{rule.continued_fee:.2f}"))
        self.station_rule_table.setItem(row_idx, 3, QTableWidgetItem(f"{rule.min_fee:.2f}"))

        unit_combo = QComboBox()
        unit_combo.addItems(["全续", "百克续"])
        unit_combo.setMinimumHeight(26)
        if rule.continued_unit == "100g":
            unit_combo.setCurrentText("百克续")
        else:
            unit_combo.setCurrentText("全续")
        unit_combo.currentTextChanged.connect(lambda text, r=row_idx: self._on_unit_changed(r, text))
        self.station_rule_table.setCellWidget(row_idx, 4, unit_combo)

        rounding_combo = QComboBox()
        rounding_combo.addItems(list(self.ROUNDING_MODE_MAP.keys()))
        rounding_combo.setMinimumHeight(26)
        rounding_text = self.ROUNDING_MODE_REVERSE.get(rule.weight_rounding, "实际重量")
        rounding_combo.setCurrentText(rounding_text)
        rounding_combo.currentTextChanged.connect(lambda text, r=row_idx: self._on_rounding_changed(r, text))
        if rule.continued_unit == "100g":
            rounding_combo.setEnabled(False)
        self.station_rule_table.setCellWidget(row_idx, 5, rounding_combo)

        region_item = QTableWidgetItem(rule.name or "")
        region_item.setFlags(region_item.flags() & ~Qt.ItemIsEditable)
        self.station_rule_table.setItem(row_idx, 6, region_item)

        for c in [1, 2, 3]:
            item = self.station_rule_table.item(row_idx, c)
            if item:
                item.setFlags(item.flags() | Qt.ItemIsEditable)

    def _load_global_rule(self):
        """加载全局规则"""
        self.config_group.setTitle("🌍 运费规则配置 - 全局兜底规则")
        self.selected_info_label.setText("🌍 正在编辑：全局兜底规则（所有客户/区域都匹配不到时使用）")
        service = RuleService()
        rules = service.load_rules()
        global_rule = None
        for r in rules:
            if r.rule_type == "global":
                global_rule = r
                break

        if not global_rule:
            global_rule = Rule("全局规则", "", "", 0, 999, 6.0, 3.0, 6.0, "global")

        self.station_rule_table.setRowCount(0)
        row = self.station_rule_table.rowCount()
        self.station_rule_table.insertRow(row)

        item0 = QTableWidgetItem("全局兜底")
        item0.setFlags(item0.flags() & ~Qt.ItemIsEditable)
        self.station_rule_table.setItem(row, 0, item0)

        self.station_rule_table.setItem(row, 1, QTableWidgetItem(f"{global_rule.first_fee:.2f}"))
        self.station_rule_table.setItem(row, 2, QTableWidgetItem(f"{global_rule.continued_fee:.2f}"))
        self.station_rule_table.setItem(row, 3, QTableWidgetItem(f"{global_rule.min_fee:.2f}"))

        unit_combo = QComboBox()
        unit_combo.addItems(["全续", "百克续"])
        unit_combo.setMinimumHeight(26)
        if global_rule.continued_unit == "100g":
            unit_combo.setCurrentText("百克续")
        else:
            unit_combo.setCurrentText("全续")
        unit_combo.currentTextChanged.connect(lambda text, r=row: self._on_unit_changed(r, text))
        self.station_rule_table.setCellWidget(row, 4, unit_combo)

        rounding_combo = QComboBox()
        rounding_combo.addItems(list(self.ROUNDING_MODE_MAP.keys()))
        rounding_combo.setMinimumHeight(26)
        rounding_text = self.ROUNDING_MODE_REVERSE.get(global_rule.weight_rounding, "实际重量")
        rounding_combo.setCurrentText(rounding_text)
        rounding_combo.currentTextChanged.connect(lambda text, r=row: self._on_rounding_changed(r, text))
        if global_rule.continued_unit == "100g":
            rounding_combo.setEnabled(False)
        self.station_rule_table.setCellWidget(row, 5, rounding_combo)

        item6 = QTableWidgetItem("全国通用")
        item6.setFlags(item6.flags() & ~Qt.ItemIsEditable)
        self.station_rule_table.setItem(row, 6, item6)

        for c in [1, 2, 3]:
            item = self.station_rule_table.item(row, c)
            if item:
                item.setFlags(item.flags() | Qt.ItemIsEditable)

        # 计泡系数（列7）- 全局规则只读显示默认值6000
        vol_div_item = QTableWidgetItem("6000（默认）")
        vol_div_item.setFlags(vol_div_item.flags() & ~Qt.ItemIsEditable)
        self.station_rule_table.setItem(row, 7, vol_div_item)

    def _add_station(self):
        """新增客户"""
        code, ok = QInputDialog.getText(self, "新增客户", "请输入客户编码:")
        if not ok or not code.strip():
            return
        name, ok = QInputDialog.getText(self, "新增客户", "请输入客户名称:")
        if not ok:
            return

        row = self.station_table.rowCount()
        self.station_table.insertRow(row)
        self.station_table.setItem(row, 0, QTableWidgetItem(code.strip()))
        self.station_table.setItem(row, 1, QTableWidgetItem(name.strip() or code.strip()))
        self.station_table.setItem(row, 2, QTableWidgetItem("继承区域"))
        self.station_table.selectRow(row)
        self._on_station_selected(row, 0)

    def _delete_station(self):
        """删除选中的客户"""
        rows = sorted({i.row() for i in self.station_table.selectedIndexes()}, reverse=True)
        if not rows:
            QMessageBox.information(self, "提示", "请先在表格中选中要删除的客户")
            return
        for r in rows:
            code_item = self.station_table.item(r, 0)
            if code_item and code_item.text().strip() in self._station_province_cache:
                del self._station_province_cache[code_item.text().strip()]
            self.station_table.removeRow(r)
        self.station_rule_table.setRowCount(0)
        self.selected_info_label.setText("请在左侧选择客户或区域")

    def _download_import_template(self):
        """下载 Excel 导入模板"""
        from PyQt5.QtWidgets import QFileDialog
        import os

        default_name = "客户+规则导入模板.xlsx"
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存模板", default_name, "Excel 文件 (*.xlsx)"
        )
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        try:
            service = RuleService()
            ok = service.generate_import_template(file_path)
            if ok:
                QMessageBox.information(self, "成功", f"模板已保存到：\n{file_path}\n\nSheet 1: 客户档案（客户编码/客户名称必填）\nSheet 2: 客户专属规则（按分区填写运费）")
            else:
                QMessageBox.warning(self, "失败", "模板生成失败。请检查是否有 Excel 写入权限。")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"生成模板失败：{str(e)}")

    def _import_stations_and_rules(self):
        """批量导入客户 + 专属规则（双 Sheet Excel）"""
        from PyQt5.QtWidgets import QFileDialog, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QTableWidget, QTableWidgetItem, QDialogButtonBox, QGroupBox
        from PyQt5.QtCore import Qt
        import os

        # ========= 1. 选择文件 =========
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择客户+规则导入文件",
            "", "Excel 文件 (*.xlsx *.xls)"
        )
        if not file_path:
            return

        # ========= 2. 解析文件 =========
        try:
            service = RuleService()
            parsed = service.parse_import_excel(file_path)
        except Exception as e:
            QMessageBox.critical(self, "解析失败", f"解析 Excel 失败：{str(e)}")
            return

        # ========= 3. 预览对话框 =========
        stations_count = len(parsed["stations"])
        rules_count = len(parsed["rules"])
        errors_count = len(parsed["errors"])
        warnings_count = len(parsed["warnings"])

        dialog = QDialog(self)
        dialog.setWindowTitle(f"导入预览 - {os.path.basename(file_path)}")
        dialog.resize(900, 600)
        layout = QVBoxLayout(dialog)

        # 顶部信息
        info_label = QLabel(
            f"📋 解析结果：<b>{stations_count}</b> 个客户，<b>{rules_count}</b> 条规则"
            + (f"，<b style='color:red'>{errors_count}</b> 个错误" if errors_count else "")
            + (f"，<b style='color:orange'>{warnings_count}</b> 个警告" if warnings_count else "")
        )
        info_label.setTextFormat(Qt.RichText)
        layout.addWidget(info_label)

        # 冲突模式选择
        conflict_group = QGroupBox("冲突处理")
        conflict_layout = QHBoxLayout(conflict_group)
        mode_combo = QComboBox()
        mode_combo.addItems([
            "覆盖模式（已存在客户和规则全部替换为新内容）",
            "跳过模式（已存在客户不更新，仅新增未存在的客户）",
            "追加模式（在已有规则基础上添加新的省份规则，不删除原有规则）"
        ])
        mode_combo.setCurrentIndex(1)  # 默认跳过
        conflict_layout.addWidget(QLabel("冲突处理："))
        conflict_layout.addWidget(mode_combo, 1)
        layout.addWidget(conflict_group)

        # 客户档案预览
        if parsed["stations"]:
            customer_group = QGroupBox(f"客户档案（共 {stations_count} 条）")
            c_layout = QVBoxLayout(customer_group)
            c_table = QTableWidget()
            c_table.setColumnCount(3)
            c_table.setHorizontalHeaderLabels(["客户编码", "客户名称", "联系电话"])
            c_table.horizontalHeader().setStretchLastSection(True)
            c_table.setRowCount(min(stations_count, 200))
            for i, s in enumerate(parsed["stations"][:200]):
                c_table.setItem(i, 0, QTableWidgetItem(s["code"]))
                c_table.setItem(i, 1, QTableWidgetItem(s["name"]))
                c_table.setItem(i, 2, QTableWidgetItem(s.get("phone", "") or ""))
            c_table.setEditTriggers(QTableWidget.NoEditTriggers)
            c_table.setSelectionBehavior(QTableWidget.SelectRows)
            c_layout.addWidget(c_table)
            layout.addWidget(customer_group, 2)

        # 规则预览
        if parsed["rules"]:
            rules_group = QGroupBox(f"专属规则（共 {rules_count} 条，仅显示前200条）")
            r_layout = QVBoxLayout(rules_group)
            r_table = QTableWidget()
            r_table.setColumnCount(5)
            r_table.setHorizontalHeaderLabels(["客户编码", "省份", "首重费", "续重费", "保底费"])
            r_table.horizontalHeader().setStretchLastSection(True)
            r_table.setRowCount(min(rules_count, 200))
            for i, r in enumerate(parsed["rules"][:200]):
                r_table.setItem(i, 0, QTableWidgetItem(r["code"]))
                r_table.setItem(i, 1, QTableWidgetItem(r["province"]))
                r_table.setItem(i, 2, QTableWidgetItem(str(r["first_fee"])))
                r_table.setItem(i, 3, QTableWidgetItem(str(r["continued_fee"])))
                r_table.setItem(i, 4, QTableWidgetItem(str(r["min_fee"])))
            r_table.setEditTriggers(QTableWidget.NoEditTriggers)
            r_table.setSelectionBehavior(QTableWidget.SelectRows)
            r_layout.addWidget(r_table)
            layout.addWidget(rules_group, 3)

        # 错误信息
        if parsed["errors"]:
            error_group = QGroupBox(f"⚠️ 错误信息（共 {errors_count} 条）")
            error_layout = QVBoxLayout(error_group)
            error_label = QLabel("\n".join(parsed["errors"][:20]))
            error_label.setStyleSheet("color: #ef4444; font-size: 11px;")
            error_label.setWordWrap(True)
            error_layout.addWidget(error_label)
            layout.addWidget(error_group)

        # 底部按钮
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.button(QDialogButtonBox.Ok).setText("✅ 确认导入")
        btn_box.button(QDialogButtonBox.Cancel).setText("取消")
        btn_box.accepted.connect(dialog.accept)
        btn_box.rejected.connect(dialog.reject)
        layout.addWidget(btn_box)

        if dialog.exec_() != QDialog.Accepted:
            return

        # ========= 4. 确认导入 =========
        mode_idx = mode_combo.currentIndex()
        conflict_mode = "overwrite" if mode_idx == 0 else ("skip" if mode_idx == 1 else "append")

        try:
            result = service.save_import_result(parsed, conflict_mode=conflict_mode)
            if result["success"]:
                stats = result["stats"]
                QMessageBox.information(
                    self, "导入成功",
                    f"✅ 导入完成！\n\n"
                    f"📦 客户: 新增 {stats['inserted_customers']} 个 / 更新 {stats['updated_customers']} 个 / 跳过 {stats['skipped_customers']} 个\n"
                    f"📋 规则: 新增 {stats['inserted_rules']} 条 / 更新 {stats['updated_rules']} 条 / 跳过 {stats['skipped_rules']} 条\n\n"
                    f"提示：数据已写入 fee_rules.json，界面将自动刷新。"
                )
                # 刷新客户列表和规则表格
                self._reload_all_after_import()
            else:
                QMessageBox.warning(self, "导入失败", result["message"])
        except Exception as e:
            QMessageBox.critical(self, "导入失败", f"保存数据时出错：{str(e)}")

    def _reload_all_after_import(self):
        """导入完成后刷新客户列表和规则缓存"""
        try:
            # 1. 重新从 fee_rules.json 加载规则到内存
            service = RuleService()
            all_rules = service.load_rules()

            # 2. 刷新客户列表
            # 先从数据库获取客户档案
            from app.models.database import get_session
            from app.models.station import Station
            session = get_session()
            try:
                stations = session.query(Station).all()
                station_map = {s.station_code: s.station_name for s in stations}
            finally:
                session.close()

            # 3. 识别每个客户的规则类型
            station_codes = set(station_map.keys())
            for r in all_rules:
                if r.rule_type == "station" and r.stations and r.stations.strip():
                    for sc in [s.strip() for s in r.stations.split(",") if s.strip()]:
                        if sc in station_codes and (r.first_fee > 0 or r.continued_fee > 0 or r.min_fee > 0):
                            # 有专属规则
                            station_map[sc] = station_map.get(sc, sc)
                            # 在后续刷新 UI 时用专属规则标识
                            pass

            # 4. 重建左侧客户列表
            self.station_table.setRowCount(0)
            # 找出有专属规则的客户
            station_with_rules = set()
            for r in all_rules:
                if r.rule_type == "station" and r.stations and r.stations.strip():
                    sc_val_list = [s.strip() for s in r.stations.split(",") if s.strip()]
                    for sc_v in sc_val_list:
                        if r.first_fee > 0 or r.continued_fee > 0 or r.min_fee > 0:
                            if r.regions and r.regions.strip():
                                station_with_rules.add(sc_v)

            # 按字母/编码顺序显示客户
            sorted_codes = sorted(station_map.keys())
            for code in sorted_codes:
                row = self.station_table.rowCount()
                self.station_table.insertRow(row)
                self.station_table.setItem(row, 0, QTableWidgetItem(code))
                self.station_table.setItem(row, 1, QTableWidgetItem(station_map[code]))
                rtype = "专属规则" if code in station_with_rules else "继承区域"
                self.station_table.setItem(row, 2, QTableWidgetItem(rtype))

            # 5. 重建缓存
            self._station_province_cache = {}
            for r in all_rules:
                if r.rule_type == "station" and r.stations and r.stations.strip():
                    sc_val_list = [s.strip() for s in r.stations.split(",") if s.strip()]
                    province_val = (r.regions or "").strip()
                    if r.first_fee > 0 or r.continued_fee > 0 or r.min_fee > 0:
                        for sc_v in sc_val_list:
                            if sc_v not in self._station_province_cache:
                                self._station_province_cache[sc_v] = {}
                            if province_val and province_val not in self._station_province_cache[sc_v]:
                                self._station_province_cache[sc_v][province_val] = [
                                    province_val,
                                    r.first_fee, r.continued_fee, r.min_fee,
                                    r.continued_unit or "kg",
                                    r.weight_rounding or "actual",
                                    getattr(r, '计泡系数', 6000.0),
                                ]

            # 6. 自动选中第一个客户并刷新右侧
            if self.station_table.rowCount() > 0:
                self.station_table.selectRow(0)
                self._on_station_selected(0, 0)
        except Exception as e:
            QMessageBox.warning(self, "刷新失败", f"导入成功，但刷新界面时出错：{str(e)}")

    def _import_stations(self):
        """批量导入客户（支持 Excel / CSV）—— 保留旧逻辑作为兼容"""
        self._import_stations_and_rules()

    def _save_all_rules(self):
        """保存所有规则"""
        try:
            service = RuleService()
            all_rules = []

            mode = self.rule_mode_combo.currentText()

            # 先收集客户列表
            stations_data = []
            for row in range(self.station_table.rowCount()):
                def get_val(col):
                    item = self.station_table.item(row, col)
                    return item.text().strip() if item else ""
                code = get_val(0)
                name = get_val(1)
                rtype = get_val(2)
                if code and name:
                    stations_data.append((code, name, rtype))

            # 如果当前是客户规则模式，先保存当前编辑
            if mode == "客户规则" and self.station_table.currentRow() >= 0:
                self._save_current_station_to_cache()

            # 生成客户规则（stations同时包含编码和名称，便于Excel用客户名称匹配）
            for code, name, rtype in stations_data:
                if code in self._station_province_cache and self._station_province_cache[code]:
                    # 专属规则 - 每个省份一条
                    for province, data in self._station_province_cache[code].items():
                        if len(data) >= 6:
                            first_fee, continued_fee, min_fee, continued_unit, weight_rounding, vol_div = data
                        elif len(data) >= 5:
                            first_fee, continued_fee, min_fee, continued_unit, weight_rounding = data
                            vol_div = 6000.0
                        else:
                            first_fee, continued_fee, min_fee = data[:3]
                            continued_unit = data[3] if len(data) > 3 else "kg"
                            weight_rounding = "actual"
                            vol_div = 6000.0
                        rule = Rule(
                            name=f"{name} - {province}",
                            regions=province,
                            stations=f"{code},{name}",
                            min_weight=0.0,
                            max_weight=999.0,
                            first_fee=float(first_fee),
                            continued_fee=float(continued_fee),
                            min_fee=float(min_fee),
                            rule_type="station",
                            continued_unit=continued_unit,
                            weight_rounding=weight_rounding,
                            计泡系数=vol_div,
                        )
                        all_rules.append(rule)
                else:
                    # 继承区域 - 只保存一条标记规则
                    rule = Rule(
                        name=name,
                        regions="",
                        stations=f"{code},{name}",  # 同时保存编码和名称
                        min_weight=0.0,
                        max_weight=999.0,
                        first_fee=0.0,
                        continued_fee=0.0,
                        min_fee=0.0,
                        rule_type="station"
                    )
                    all_rules.append(rule)

            # 读取旧的区域规则和全局规则
            old_rules = service.load_rules()
            for r in old_rules:
                if r.rule_type == "region":
                    all_rules.append(r)
                elif r.rule_type == "global":
                    all_rules.append(r)

            # 如果没有全局规则，添加默认
            has_global = any(r.rule_type == "global" for r in all_rules)
            if not has_global:
                all_rules.append(Rule("全局规则", "", "", 0, 999, 6.0, 3.0, 6.0, "global"))

            # 收集活动加价规则（含省份限定）
            promo_list = []
            for row in range(self.promotion_table.rowCount()):
                def get_promo_val(col):
                    item = self.promotion_table.item(row, col)
                    return item.text().strip() if item else ""
                name = get_promo_val(0)
                start = get_promo_val(1)
                end = get_promo_val(2)
                mtype = get_promo_val(3)
                mval = get_promo_val(4)
                regions = get_promo_val(5)  # 省份限定（逗号分隔），空=所有省份
                if name and start and end:
                    promo_list.append({
                        "name": name,
                        "start_date": start,
                        "end_date": end,
                        "markup_type": mtype or "percent",
                        "markup_value": float(mval) if mval else 0,
                        "regions": regions or ""
                    })

            if service.save_rules(all_rules, promo_list):
                # 保存成功后，从文件重新加载一次（确保内存缓存与磁盘文件一致）
                try:
                    self._reload_rules_from_file_only()
                except Exception:
                    pass

                # 更新客户表格的规则类型显示
                for row in range(self.station_table.rowCount()):
                    code_item = self.station_table.item(row, 0)
                    type_item = self.station_table.item(row, 2)
                    if code_item and type_item:
                        code = code_item.text().strip()
                        if code in self._station_province_cache and self._station_province_cache[code]:
                            type_item.setText("专属规则")
                        else:
                            type_item.setText("继承区域")

                # 如果当前有选中的客户，刷新右侧规则表格
                current_row = self.station_table.currentRow()
                if current_row >= 0:
                    try:
                        self._on_station_selected(current_row, 0)
                    except Exception:
                        pass

                QMessageBox.information(self, "成功", f"已保存 {len(all_rules)} 条运费规则 + {len(promo_list)} 条活动规则！")
            else:
                QMessageBox.warning(self, "失败", "保存失败")
        except Exception as e:
            import traceback
            traceback.print_exc()
            QMessageBox.warning(self, "保存失败", str(e))

    def _save_current_station_to_cache(self):
        """将当前选中客户的分组表格数据保存到缓存（按组内省份展开）"""
        row = self.station_table.currentRow()
        if row < 0:
            return
        code_item = self.station_table.item(row, 0)
        if not code_item:
            return
        station_code = code_item.text().strip()

        use_custom = self.btn_use_custom.isChecked()

        if not use_custom:
            # 继承区域模式 - 清空缓存
            if station_code in self._station_province_cache:
                del self._station_province_cache[station_code]
            return

        # 专属规则 - 从分组表提取费用数据，并按组内省份展开
        province_data = {}
        for table_row in range(self.station_rule_table.rowCount()):
            def get_val(col):
                item = self.station_rule_table.item(table_row, col)
                return item.text().strip() if item else ""

            group_name = get_val(0)
            provinces_text = get_val(1)
            try:
                first_fee = float(get_val(2) or 0)
                continued_fee = float(get_val(3) or 0)
                min_fee = float(get_val(4) or 0)
            except ValueError:
                continue

            unit_widget = self.station_rule_table.cellWidget(table_row, 5)
            if unit_widget and isinstance(unit_widget, QComboBox):
                continued_unit_text = unit_widget.currentText()
            else:
                continued_unit_text = "全续"
            unit_code = "100g" if continued_unit_text == "百克续" else "kg"

            rounding_widget = self.station_rule_table.cellWidget(table_row, 6)
            if rounding_widget and isinstance(rounding_widget, QComboBox):
                rounding_text = rounding_widget.currentText()
                rounding_code = self.ROUNDING_MODE_MAP.get(rounding_text, "actual")
            else:
                rounding_code = "actual"

            # 计泡系数（列7）
            vol_div_widget = self.station_rule_table.cellWidget(table_row, 7)
            if vol_div_widget and isinstance(vol_div_widget, QComboBox):
                vol_div_text = vol_div_widget.currentText()
                # 从文本中提取数字，如 "6000（顺丰/京东/德邦）" -> 6000
                import re
                m = re.search(r'\d+', vol_div_text)
                vol_div_code = float(m.group()) if m else 6000.0
            else:
                vol_div_code = 6000.0

            # 找到该分组对应的省份列表
            target_provinces = []
            for (gn, provs, *_) in PROVINCE_GROUPS:
                if gn == group_name:
                    target_provinces = provs
                    break
            if not target_provinces and provinces_text and provinces_text != "N/A":
                # 兼容旧数据：尝试按"省"字解析
                target_provinces = [p for p in provinces_text.replace("、", " ").split() if p]

            for province in target_provinces:
                province_data[province] = (first_fee, continued_fee, min_fee, unit_code, rounding_code, vol_div_code)

        if province_data:
            self._station_province_cache[station_code] = province_data

    def _add_promotion(self):
        """添加活动加价规则（省份留空=对所有省份生效）"""
        row = self.promotion_table.rowCount()
        self.promotion_table.insertRow(row)
        self.promotion_table.setItem(row, 0, QTableWidgetItem("双十一大促"))
        self.promotion_table.setItem(row, 1, QTableWidgetItem("2026-11-01"))
        self.promotion_table.setItem(row, 2, QTableWidgetItem("2026-11-15"))
        self.promotion_table.setItem(row, 3, QTableWidgetItem("percent"))
        self.promotion_table.setItem(row, 4, QTableWidgetItem("20"))
        self.promotion_table.setItem(row, 5, QTableWidgetItem(""))  # 省份限定（留空=所有省份）

    def _delete_promotion(self):
        """删除选中的活动规则"""
        rows = sorted({i.row() for i in self.promotion_table.selectedIndexes()}, reverse=True)
        if not rows:
            return
        for r in rows:
            self.promotion_table.removeRow(r)

    def _rule_test(self):
        """测试计算"""
        try:
            w = float(self.test_weight.text() or 0)
        except ValueError:
            QMessageBox.warning(self, "格式错误", "重量必须是数字")
            return
        if hasattr(self.test_region, "currentText"):
            region = self.test_region.currentText()
        else:
            region = self.test_region.text()
        station_code = self.test_station.text()

        service = RuleService()
        result = service.calculate_fee(w, region, station_code)

        if result["is_exception"]:
            status = "异常 (" + result["remark"] + ")"
        else:
            status = f"¥{result['fee']:.2f}（{result['rule_name']}）"
            # 额外提示活动规则日期信息
            try:
                promo_rules = service.load_promotion_rules()
                if promo_rules:
                    from datetime import datetime
                    today_str = datetime.now().strftime("%Y-%m-%d")
                    today = datetime.strptime(today_str, "%Y-%m-%d")
                    active_count = 0
                    upcoming_count = 0
                    for pr in promo_rules:
                        try:
                            start = datetime.strptime(str(pr.get("start_date", "")), "%Y-%m-%d")
                            end = datetime.strptime(str(pr.get("end_date", "")), "%Y-%m-%d")
                            if start <= today <= end:
                                active_count += 1
                            elif today < start:
                                upcoming_count += 1
                        except Exception:
                            continue
                    if active_count == 0 and upcoming_count > 0:
                        status += f" [注意:{upcoming_count}条活动未开始]"
            except Exception:
                pass

        self.test_result.setText("结果：" + status)

    def _load_default_settings(self):
        """启动时加载已保存的默认设置到UI"""
        try:
            config_file = get_config_file("default_settings.json")
            if not os.path.exists(config_file):
                return
            with open(config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
            # 把已保存的值回填到 UI
            if "first_weight" in data:
                self.default_first_weight.setText(str(data["first_weight"]))
            if "continued_fee" in data:
                self.default_continued_fee.setText(str(data["continued_fee"]))
            if "min_fee" in data:
                self.default_min_fee.setText(str(data["min_fee"]))
            if "empty_weight_fee" in data:
                self.default_empty_weight_fee.setText(str(data["empty_weight_fee"]))
        except Exception:
            pass

    def _save_default_settings(self):
        """保存全局默认设置"""
        import json
        import os
        try:
            # 验证输入
            first_weight = float(self.default_first_weight.text() or 1.0)
            continued_fee = float(self.default_continued_fee.text() or 2.0)
            min_fee = float(self.default_min_fee.text() or 5.0)
            empty_weight_fee = float(self.default_empty_weight_fee.text() or 3.0)

            # 保存到配置文件
            config_file = get_config_file("default_settings.json")

            config = {
                "first_weight": first_weight,
                "continued_fee": continued_fee,
                "min_fee": min_fee,
                "empty_weight_fee": empty_weight_fee
            }
            with open(config_file, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)

            QMessageBox.information(self, "成功", "默认设置已保存！")
        except ValueError as e:
            QMessageBox.warning(self, "格式错误", str(e))
        except Exception as e:
            QMessageBox.warning(self, "保存失败", str(e))

    # ==================== 事件处理 ====================

    def _select_file(self):
        """选择Excel文件 - 支持最多5个文件批量导入"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self, "选择Excel文件（最多5个）", "",
            "Excel Files (*.xlsx *.xls);;CSV Files (*.csv);;All Files (*)"
        )
        if file_paths:
            # 限制最多5个
            if len(file_paths) > 5:
                QMessageBox.warning(self, "提示", f"最多支持同时处理5个文件，已自动截取前5个（共选择了{len(file_paths)}个）")
                file_paths = file_paths[:5]

            self.selected_files = file_paths

            # 显示文件列表
            file_list_text = f"已选择 {len(file_paths)} 个文件：\n"
            for i, fp in enumerate(file_paths):
                fsize = os.path.getsize(fp) / (1024 * 1024)
                file_list_text += f"  {i + 1}. {os.path.basename(fp)}  ({fsize:.1f} MB)\n"
            self.file_list_label.setText(file_list_text.strip())
            self.file_list_label.setStyleSheet("color: #4f46e5; padding: 8px 12px; background: #eef2ff; border: 1px solid #c7d2fe; border-radius: 8px; font-size: 11px;")

            # 用第一个文件预览列名
            self._preview_columns_async(file_paths[0])

    def _preview_columns_async(self, file_path):
        """异步预览列名和匹配结果 - 不会阻塞UI主线程"""
        self.match_label.setText("⏳ 正在读取文件，请稍候...")
        self.match_label.setStyleSheet("padding: 8px 12px; color: #4f46e5; font-weight: 500; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; font-size: 11px;")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(5)
        self.progress_bar.setFormat("读取文件中... %p%")
        self.progress_stage_label.setText("正在读取文件以获取列名...")

        # 使用后台线程读取Excel
        self.preview_worker = FilePreviewWorker(file_path)
        self.preview_worker.progress.connect(self._on_preview_progress)
        self.preview_worker.finished.connect(self._on_preview_finished)
        self.preview_worker.error.connect(self._on_preview_error)
        self.preview_worker.start()

    def _on_preview_progress(self, percent, stage_text):
        """预览阶段进度更新"""
        self.progress_bar.setValue(percent)
        self.progress_stage_label.setText(stage_text)

    def _on_preview_finished(self, parse_result):
        """预览完成，显示列名匹配结果"""
        columns = parse_result["columns"]
        row_count = parse_result["row_count"]
        matched = parse_result["matched"]

        text_lines = []
        text_lines.append(f"✅ 检测到 {len(columns)} 列，{row_count} 行数据\n")
        text_lines.append("已匹配的列：")
        for std, actual in matched.get("matched", {}).items():
            text_lines.append(f"  {std}  <--  {actual}")

        if matched.get("unmatched"):
            text_lines.append("")
            text_lines.append(f"未匹配的列：{', '.join(matched['unmatched'])}")
            text_lines.append("（未匹配列不影响计算，但不会参与运费逻辑）")

        self.match_label.setStyleSheet("padding: 8px 12px; color: #334155; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; font-size: 11px;")
        self.match_label.setText("\n".join(text_lines))

        # 完成后隐藏进度条，重置状态
        self.progress_bar.setValue(100)
        self.progress_stage_label.setText("文件读取完成，可以开始计算")

    def _on_preview_error(self, error):
        """预览失败"""
        self.match_label.setStyleSheet("padding: 8px 12px; color: #ef4444; font-weight: 500; background: #fef2f2; border: 1px solid #fecaca; border-radius: 8px; font-size: 11px;")
        self.match_label.setText(f"❌ 解析失败：{error}")
        self.progress_bar.setVisible(False)
        self.progress_stage_label.setText("读取失败")

    def _start_calculate(self):
        """开始计算 - 支持单文件/多文件批量处理"""
        if not hasattr(self, "selected_files") or not self.selected_files:
            QMessageBox.warning(self, "提示", "请先选择Excel文件")
            return

        # 显示进度UI
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        self.progress_stage_label.setText(f"准备开始... 共 {len(self.selected_files)} 个文件")
        self.progress_stage_label.setStyleSheet("color: #4f46e5; font-weight: 500; padding: 3px; font-size: 11px;")

        # 禁用按钮，防止重复点击
        self.calc_btn.setEnabled(False)

        # 日志
        for fp in self.selected_files:
            self.log_text.append(f"▶ 加入待处理：{os.path.basename(fp)}")
        self.statusBar.showMessage("计算中...")

        # 后台线程批量计算
        self.calc_worker = CalculateWorker(self.selected_files)
        self.calc_worker.progress.connect(self._on_calculate_progress)
        self.calc_worker.file_progress.connect(self._on_file_progress)
        self.calc_worker.finished.connect(self._on_calculate_finished)
        self.calc_worker.error.connect(self._on_calculate_error)
        self.calc_worker.start()

    def _on_file_progress(self, file_idx, total_files, file_name):
        """新文件开始处理时的提示"""
        self.log_text.append(f"\n--- 正在处理第 {file_idx}/{total_files} 个文件：{file_name} ---")

    def _on_calculate_progress(self, percent, stage_text):
        """计算阶段进度更新 - 直接在UI线程"""
        self.progress_bar.setValue(percent)
        self.progress_bar.setFormat(f"{percent}%")
        self.progress_stage_label.setText(stage_text)
        # 状态栏同步更新
        self.statusBar.showMessage(f"计算中... {percent}%")

    def _on_calculate_finished(self, result):
        """计算完成"""
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat("100%")
        self.progress_stage_label.setText("✅ 计算完成！")
        self.progress_stage_label.setStyleSheet("color: #10b981; font-weight: 500; padding: 3px; font-size: 11px;")
        self.calc_btn.setEnabled(True)

        self.current_record_id = result["record_id"]
        # 保存所有文件的记录ID列表
        self.all_record_ids = [f["record_id"] for f in result.get("files", [])]
        self.statusBar.showMessage("计算完成")

        # 填充文件选择下拉框
        self.file_combo.clear()
        file_results = result.get("files", [])
        for i, fr in enumerate(file_results):
            display_text = f"{i+1}. {fr['file_name']} ({fr['total_rows']:,}行)"
            self.file_combo.addItem(display_text, fr["record_id"])

        # 多文件处理时显示汇总
        file_count = result.get("file_count", 1)
        if file_count > 1:
            # 多文件汇总日志
            file_results = result.get("files", [])
            for fr in file_results:
                self.log_text.append(
                    f"✅ {fr['file_name']}：¥{fr['total_fee']:.2f}，成功 {fr['success_count']} 条，异常 {fr['exception_count']} 条"
                )
            failed_files = result.get("failed_files", [])
            if failed_files:
                for ff in failed_files:
                    self.log_text.append(f"❌ {ff}")

            self.log_text.append(
                f"\n=== 全部{file_count}个文件处理完成：运费总额 ¥{result['total_fee']:.2f}，成功 {result['success_count']} 条，异常 {result['exception_count']} 条 ==="
            )

            # 多文件结果弹窗
            msg = f"全部{file_count}个文件处理完成！\n\n运费总额：¥{result['total_fee']:.2f}\n成功：{result['success_count']} 条\n异常：{result['exception_count']} 条"
            if failed_files:
                msg += f"\n失败：{len(failed_files)} 个"
            QMessageBox.information(self, "完成", msg)
        else:
            self.log_text.append(f"✅ 计算完成：总额 ¥{result['total_fee']:.2f}，成功 {result['success_count']} 条，异常 {result['exception_count']} 条")
            QMessageBox.information(self, "完成",
                f"计算完成！\n\n运费总额：¥{result['total_fee']:.2f}\n成功：{result['success_count']} 条\n异常：{result['exception_count']} 条")

        # 切换到结果Tab
        self.tabs.setCurrentIndex(1)
        self._load_result(result["record_id"])

    def _on_calculate_error(self, error):
        """计算失败"""
        try:
            self.progress_bar.setVisible(False)
            self.progress_stage_label.setText(f"❌ 计算失败：{error}")
            self.progress_stage_label.setStyleSheet("color: #ef4444; font-weight: 500; padding: 3px; font-size: 11px;")
            self.calc_btn.setEnabled(True)
            self.statusBar.showMessage("计算失败")
            self.log_text.append(f"❌ 计算失败：{error}")
        except Exception as inner_err:
            print(f"UI更新失败: {inner_err}")
        try:
            QMessageBox.critical(self, "失败", f"计算失败：{error}")
        except Exception as inner_err:
            print(f"弹窗失败: {inner_err}")

    def _load_result(self, record_id):
        """加载结果 - 后台线程加载数据避免UI卡死"""
        try:
            self.statusBar.showMessage("正在加载计算结果...")
            # 先切换到结果页
            self.tabs.setCurrentIndex(1)

            # 清除现有内容，显示加载状态
            self.result_table.setRowCount(0)
            # 在表格区域临时显示"加载中"提示（通过设置第一行文本）
            self.result_table.setRowCount(1)
            self.result_table.setItem(0, 0, QTableWidgetItem("正在加载数据，请稍候..."))
            for col in range(11):
                self.result_table.setItem(0, col, QTableWidgetItem(""))
            self.result_table.item(0, 0).setText("正在加载数据，请稍候...")

            # 启动后台线程加载数据
            self.result_worker = ResultLoadWorker(record_id, max_display=20000)
            self.result_worker.finished.connect(self._on_result_loaded)
            self.result_worker.error.connect(self._on_result_load_error)
            self.result_worker.start()
        except Exception as e:
            import traceback
            print(f"_load_result启动异常: {e}")
            traceback.print_exc()
            QMessageBox.warning(self, "提示", f"加载结果失败: {e}")

    def _on_file_combo_changed(self, index):
        """文件选择下拉框切换时加载对应文件的结果"""
        record_id = self.file_combo.currentData()
        if record_id and record_id != self.current_record_id:
            self.current_record_id = record_id
            self._load_result(record_id)

    def _on_result_loaded(self, payload):
        """后台线程完成数据加载后，在UI线程填充表格"""
        try:
            summary = payload["summary"]
            prepared = payload["prepared"]
            total_rows = payload["total_rows"]
            display_count = payload["display_count"]

            # 缓存完整details，用于结算/导出；同时清空结算缓存，下次切换时重新计算
            self.current_details = payload.get("details", [])
            self._settlement_cache = {}  # 切换按钮时按需重新结算

            # 更新汇总信息
            self.summary_labels["总行数"].setText(str(summary["total_rows"]))
            self.summary_labels["成功"].setText(str(summary["success_count"]))
            self.summary_labels["异常"].setText(str(summary["exception_count"]))
            self.summary_labels["运费总额"].setText(f"¥{summary['total_fee']:.2f}")

            # 填充表格 - 禁用刷新/信号/排序，一次性构建完恢复
            self.result_table.setSortingEnabled(False)
            self.result_table.blockSignals(True)
            self.result_table.setUpdatesEnabled(False)

            try:
                self.result_table.setRowCount(display_count)
                for i, row in enumerate(prepared):
                    for col in range(11):
                        self.result_table.setItem(i, col, QTableWidgetItem(row[col]))
            finally:
                self.result_table.setUpdatesEnabled(True)
                self.result_table.blockSignals(False)

            # 状态栏提示
            if total_rows > display_count:
                self.statusBar.showMessage(f"表格显示前{display_count:,}行（共{total_rows:,}行），导出Excel为完整数据")
            else:
                self.statusBar.showMessage(f"已加载{total_rows:,}行结果")
        except Exception as e:
            import traceback
            print(f"_on_result_loaded填充表格失败: {e}")
            traceback.print_exc()
            self.statusBar.showMessage(f"结果填充失败: {e}")

    def _ensure_details_loaded(self) -> bool:
        """确保完整数据已加载（结算/导出前调用）- 从DB按需加载"""
        if self.current_details:
            return True
        if not self.current_record_id:
            return False

        try:
            self.statusBar.showMessage("正在加载完整数据用于结算/导出...")
            from app.models.database import get_session
            from app.models.fee_detail import FeeDetail

            session = get_session()
            try:
                self.current_details = session.query(FeeDetail).filter(
                    FeeDetail.record_id == self.current_record_id
                ).all()
            finally:
                session.close()

            self.statusBar.showMessage(f"完整数据已加载（{len(self.current_details):,}行）")
            return bool(self.current_details)
        except Exception as e:
            import traceback
            print(f"_ensure_details_loaded加载失败: {e}")
            traceback.print_exc()
            self.statusBar.showMessage(f"加载失败: {e}")
            return False

    def _on_result_load_error(self, error):
        """结果数据加载失败"""
        print(f"加载结果失败: {error}")
        self.statusBar.showMessage(f"加载结果失败")
        try:
            self.result_table.setRowCount(1)
            self.result_table.setItem(0, 0, QTableWidgetItem(f"加载失败: {str(error)[:200]}"))
        except Exception:
            pass

    def _switch_settlement(self, type_):
        """切换结算视图 - 网点/承包区/月结客户"""
        active_style = """
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #6366f1, stop:1 #4f46e5);
                color: white;
                border: none;
                padding: 7px 20px;
                border-radius: 8px;
                font-weight: 500;
                font-size: 11px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #818cf8, stop:1 #6366f1);
            }
        """
        inactive_style = """
            QPushButton {
                background: #ffffff;
                color: #64748b;
                border: 1px solid #e2e8f0;
                padding: 7px 20px;
                border-radius: 8px;
                font-weight: 500;
                font-size: 11px;
            }
            QPushButton:hover {
                background: #eef2ff;
                color: #4f46e5;
                border-color: #c7d2fe;
            }
        """

        self.station_btn.setChecked(type_ == "station")
        self.contract_btn.setChecked(type_ == "contract")
        self.monthly_btn.setChecked(type_ == "monthly")

        self.station_btn.setStyleSheet(active_style if type_ == "station" else inactive_style)
        self.contract_btn.setStyleSheet(active_style if type_ == "contract" else inactive_style)
        self.monthly_btn.setStyleSheet(active_style if type_ == "monthly" else inactive_style)

        # 确保完整数据已加载
        if not self._ensure_details_loaded():
            QMessageBox.warning(self, "提示", "暂无结算数据，请先完成计算")
            return

        # 1) 结算缓存：只在第一次切换时做全量遍历，之后直接取缓存
        if not hasattr(self, '_settlement_cache'):
            self._settlement_cache = {}

        if type_ not in self._settlement_cache:
            engine = SettlementEngine()
            if type_ == "station":
                data = engine.calculate_station_settlement(self.current_details)
                headers = ["网点编码", "网点名称", "订单数", "总重量(kg)", "总运费(元)", "分成比例", "网点收入(元)", "公司收入(元)"]
                keys = ["station_code", "station_name", "order_count", "total_weight", "total_fee", "commission_rate", "station_income", "company_income"]
            elif type_ == "contract":
                data = engine.calculate_contract_settlement(self.current_details)
                headers = ["承包区编码", "承包区名称", "网点", "订单数", "总重量(kg)", "总运费(元)", "分成比例", "承包区收入(元)"]
                keys = ["contract_code", "contract_name", "station_code", "order_count", "total_weight", "total_fee", "commission_rate", "contract_income"]
            else:  # monthly
                data = engine.calculate_monthly_settlement(self.current_details)
                headers = ["客户编码", "客户名称", "订单数", "总重量(kg)", "总运费(元)", "代收货款(元)", "账单状态", "应收金额(元)"]
                keys = ["customer_code", "customer_name", "order_count", "total_weight", "total_fee", "cod_amount", "status", "receivable"]
            self._settlement_cache[type_] = {"data": data, "headers": headers, "keys": keys}

        cache = self._settlement_cache[type_]
        data = cache["data"]
        headers = cache["headers"]
        keys = cache["keys"]

        # 2) 表格填充性能优化：批量禁用刷新/信号/排序，一次性填充完再恢复
        self.settlement_table.setSortingEnabled(False)
        self.settlement_table.blockSignals(True)
        self.settlement_table.setUpdatesEnabled(False)

        self.settlement_table.setColumnCount(len(headers))
        self.settlement_table.setHorizontalHeaderLabels(headers)
        self.settlement_table.setRowCount(len(data))

        col_count = len(keys)
        for i, item in enumerate(data):
            for j in range(col_count):
                val = item.get(keys[j], "")
                if isinstance(val, float):
                    val = f"{val:.2f}"
                self.settlement_table.setItem(i, j, QTableWidgetItem(str(val)))

        self.settlement_table.setUpdatesEnabled(True)
        self.settlement_table.blockSignals(False)
        self.settlement_table.setSortingEnabled(True)

    def _show_export_overlay(self):
        """显示导出进度（使用 QProgressDialog，PyQt5 标准组件，不会崩溃）"""
        self._hide_export_overlay()

        # QProgressDialog 是 PyQt5 标准组件，父窗口设为 self，自动居中
        dlg = QProgressDialog(
            "正在准备写入 Excel...",          # 正文
            None,                              # 取消按钮文本（None=不显示取消）
            0,                                 # 最小值
            100,                               # 最大值
            self                               # 父窗口
        )
        dlg.setWindowTitle("数据导出中")
        dlg.setWindowModality(Qt.WindowModal)  # 窗口级模态
        dlg.setMinimumDuration(0)              # 立即显示（不等待）
        dlg.setFixedWidth(420)
        dlg.setAutoClose(False)                # 100% 后不自动关闭
        dlg.setAutoReset(False)                # 100% 后不自动重置
        dlg.setCancelButton(None)              # 禁用取消按钮

        # 设置窗口图标（猴子图标）
        try:
            icon_path = os.path.join(get_resource_path("data", "icons"), "monkey-icon.png")
            if os.path.exists(icon_path):
                dlg.setWindowIcon(QIcon(icon_path))
        except Exception:
            pass

        # 显示
        dlg.show()
        dlg.setValue(0)
        dlg.setLabelText("正在准备写入 Excel，请稍候...")
        # 主动处理事件让窗口先渲染出来
        from PyQt5.QtWidgets import QApplication as _QApp
        _QApp.processEvents()

        self._export_overlay = dlg

    def _update_export_progress(self, value, message=""):
        """更新导出进度（由后台线程的 progress 信号触发）"""
        if self._export_overlay is None:
            return
        try:
            value = int(max(0, min(100, value)))
            if message:
                self._export_overlay.setLabelText(f"{value:3d}%  {message}")
            else:
                self._export_overlay.setLabelText(f"{value:3d}%")
            self._export_overlay.setValue(value)
            # 处理事件，刷新 UI
            from PyQt5.QtWidgets import QApplication as _QApp
            _QApp.processEvents()
        except Exception:
            pass

    def _hide_export_overlay(self):
        """关闭导出进度"""
        if self._export_overlay is None:
            return
        dlg = self._export_overlay
        self._export_overlay = None
        try:
            dlg.close()
            dlg.deleteLater()
        except Exception:
            pass

    def _export_details(self):
        """导出明细 - 先弹保存对话框，再后台线程导出，避免未响应"""
        if not self.current_record_id:
            QMessageBox.warning(self, "提示", "暂无数据可导出")
            return

        # 先查一次获取源文件名 & 业务日期
        original_file_name = "运费明细"
        business_date_str = ""
        try:
            from app.models.database import get_session
            from app.models.fee_record import FeeRecord
            from app.models.fee_detail import FeeDetail
            import json as _json
            import re as _re

            session = get_session()
            try:
                record = session.query(FeeRecord).filter(FeeRecord.id == self.current_record_id).first()
                if record and record.file_name:
                    original_file_name = os.path.splitext(os.path.basename(record.file_name))[0]

                first_d = session.query(FeeDetail).filter(
                    FeeDetail.record_id == self.current_record_id
                ).order_by(FeeDetail.id).first()
                if first_d and first_d.original_data:
                    try:
                        od = first_d.original_data if isinstance(first_d.original_data, dict) \
                            else _json.loads(first_d.original_data)
                        raw = od.get("business_date", "")
                        if raw:
                            digits = _re.findall(r"\d+", str(raw))
                            if len(digits) >= 3:
                                business_date_str = f"{int(digits[0]):04d}{int(digits[1]):02d}{int(digits[2]):02d}"
                    except Exception:
                        pass
            finally:
                session.close()
        except Exception:
            pass

        if not business_date_str:
            from datetime import datetime as _dt
            business_date_str = _dt.now().strftime("%Y%m%d")

        default_name = f"{original_file_name}-帐单已结算{business_date_str}.xlsx"

        # 按优先级测试可写目录，默认给到一个肯定能写的目录
        # 开发模式：__file__ 在源代码位置；打包模式（PyInstaller）：sys.executable 是 exe 所在目录
        import sys as _sys
        if getattr(sys, 'frozen', False):
            # PyInstaller 打包模式，程序目录 = exe 所在目录
            app_root = os.path.dirname(sys.executable)
        else:
            # 开发模式
            app_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

        candidate_dirs = [
            get_data_dir("exports"),                                    # 用户数据目录/data/exports（安装后模式：APPDATA下）
            os.path.expanduser("~/Desktop"),                            # 桌面
            get_user_data_dir(),                                         # 用户数据根目录
            os.path.expanduser("~"),                                     # 用户目录
            os.getcwd(),                                                 # 当前目录
        ]
        default_dir = None
        for d in candidate_dirs:
            try:
                os.makedirs(d, exist_ok=True)
                test_file = os.path.join(d, "_write_test.tmp")
                with open(test_file, "w", encoding="utf-8") as f:
                    f.write("ok")
                try:
                    os.remove(test_file)
                except Exception:
                    pass
                default_dir = d
                break
            except Exception:
                continue
        if not default_dir:
            QMessageBox.critical(self, "失败", "无法找到可写目录，请检查磁盘权限")
            return

        # 弹出保存文件对话框
        file_path, _ = QFileDialog.getSaveFileName(
            self, "导出 Excel", os.path.join(default_dir, default_name),
            "Excel 文件 (*.xlsx)"
        )
        if not file_path:
            return
        if not file_path.lower().endswith(".xlsx"):
            file_path += ".xlsx"

        # 启动后台线程导出
        self.statusBar.showMessage("正在导出 Excel，请稍候...")
        self._show_export_overlay()
        try:
            self.export_worker = ExportWorker(self.current_record_id, file_path)
            self.export_worker.finished.connect(self._on_export_done)
            self.export_worker.error.connect(self._on_export_error)
            self.export_worker.progress.connect(self._update_export_progress)
            self.export_worker.start()
        except Exception as e:
            self._hide_export_overlay()
            QMessageBox.critical(self, "失败", f"导出启动失败：{e}")

    def _on_export_done(self, file_path: str):
        self._update_export_progress(100, "写入完成，正在关闭文件...")
        QTimer.singleShot(400, lambda: self._hide_export_overlay())
        QTimer.singleShot(500, lambda: self.statusBar.showMessage(f"导出完成：{file_path}"))
        QTimer.singleShot(500, lambda: QMessageBox.information(
            self, "成功", f"已导出到：\n{file_path}"))

    def _on_export_error(self, err: str):
        self._hide_export_overlay()
        self.statusBar.showMessage("导出失败")
        QMessageBox.critical(self, "失败", f"导出失败：\n{err[:500]}")

    def _export_all_separately(self):
        """分别导出所有导入的文件 - 使用后台线程避免UI未响应"""
        if not self.all_record_ids:
            QMessageBox.warning(self, "提示", "暂无数据可导出")
            return

        if len(self.all_record_ids) == 1:
            self._export_details()
            return

        export_dir = QFileDialog.getExistingDirectory(
            self, "选择导出目录", os.path.expanduser("~/Desktop")
        )
        if not export_dir:
            return

        from app.services.export_service import ExportService
        service = ExportService()
        records_info = service.get_record_info(self.all_record_ids)

        info_text = f"将分别导出 {len(self.all_record_ids)} 个文件到：\n{export_dir}\n\n"
        for info in records_info:
            info_text += f"• {info['file_name']}: {info['total_rows']} 行, ¥{info['total_fee']:.2f}\n"

        reply = QMessageBox.question(self, "确认导出", info_text,
                                     QMessageBox.Yes | QMessageBox.No)
        if reply != QMessageBox.Yes:
            return

        self.statusBar.showMessage("正在导出多个文件...")
        self._show_export_overlay()
        try:
            self._multi_export_dir = export_dir
            self.multi_export_worker = ExportMultiWorker(self.all_record_ids, export_dir)
            self.multi_export_worker.finished.connect(self._on_multi_export_done)
            self.multi_export_worker.error.connect(self._on_multi_export_error)
            self.multi_export_worker.progress.connect(self._update_export_progress)
            self.multi_export_worker.start()
        except Exception as e:
            self._hide_export_overlay()
            QMessageBox.critical(self, "失败", f"导出启动失败：{e}")

    def _on_multi_export_done(self, exported_files: list):
        self._update_export_progress(100, "写入完成，正在关闭文件...")
        QTimer.singleShot(400, lambda: self._hide_export_overlay())
        export_dir = getattr(self, '_multi_export_dir', "")
        if export_dir and len(exported_files) > 0:
            first_dir = os.path.dirname(exported_files[0])
        else:
            first_dir = export_dir
        QTimer.singleShot(500, lambda: self.statusBar.showMessage(
            f"导出完成：{first_dir}  共 {len(exported_files)} 个文件"))
        QTimer.singleShot(500, lambda: QMessageBox.information(
            self, "成功", f"已导出 {len(exported_files)} 个文件到：\n{first_dir}"))

    def _on_multi_export_error(self, err: str):
        self._hide_export_overlay()
        self.statusBar.showMessage("导出失败")
        QMessageBox.critical(self, "失败", f"导出失败：\n{err[:500]}")

    def _export_merged(self):
        """合并导出所有文件（自动拆分避免超104万行）- 使用后台线程"""
        if not self.all_record_ids:
            QMessageBox.warning(self, "提示", "暂无数据可导出")
            return

        export_dir = QFileDialog.getExistingDirectory(
            self, "选择导出目录", os.path.expanduser("~/Desktop")
        )
        if not export_dir:
            return

        from app.services.export_service import ExportService
        service = ExportService()
        records_info = service.get_record_info(self.all_record_ids)
        total_rows = sum(info['total_rows'] for info in records_info)

        estimated_files = max(1, (total_rows + 999999) // 1000000)

        info_text = f"总数据：{total_rows:,} 行\n"
        info_text += f"预计导出：{estimated_files} 个Excel文件\n"
        info_text += f"导出目录：{export_dir}\n\n"
        info_text += "注意：合并导出会新增\"来源文件\"列，标识每行数据的原始文件名。\n"
        if estimated_files > 1:
            info_text += f"\n⚠️ 数据超过100万行，将自动拆分为 {estimated_files} 个文件。"

        reply = QMessageBox.question(self, "确认合并导出", info_text,
                                     QMessageBox.Yes | QMessageBox.No)
        if reply != QMessageBox.Yes:
            return

        self.statusBar.showMessage("正在合并导出...")
        self._show_export_overlay()
        try:
            self._merged_export_dir = export_dir
            self.merged_export_worker = ExportMergedWorker(self.all_record_ids, export_dir)
            self.merged_export_worker.finished.connect(self._on_merged_export_done)
            self.merged_export_worker.error.connect(self._on_merged_export_error)
            self.merged_export_worker.progress.connect(self._update_export_progress)
            self.merged_export_worker.start()
        except Exception as e:
            self._hide_export_overlay()
            QMessageBox.critical(self, "失败", f"导出启动失败：{e}")

    def _on_merged_export_done(self, exported_files: list):
        self._update_export_progress(100, "写入完成，正在关闭文件...")
        QTimer.singleShot(400, lambda: self._hide_export_overlay())
        export_dir = getattr(self, '_merged_export_dir', "")
        if len(exported_files) > 1:
            QTimer.singleShot(500, lambda: self.statusBar.showMessage(
                f"导出完成：{export_dir}  共 {len(exported_files)} 个文件"))
            QTimer.singleShot(500, lambda: QMessageBox.information(
                self, "成功",
                f"数据已拆分导出为 {len(exported_files)} 个文件：\n\n" +
                "\n".join([os.path.basename(f) for f in exported_files]) +
                f"\n\n保存目录：{export_dir}"))
        else:
            QTimer.singleShot(500, lambda: self.statusBar.showMessage(
                f"导出完成：{exported_files[0]}"))
            QTimer.singleShot(500, lambda: QMessageBox.information(
                self, "成功", f"已导出到：\n{exported_files[0]}"))

    def _on_merged_export_error(self, err: str):
        self._hide_export_overlay()
        self.statusBar.showMessage("导出失败")
        QMessageBox.critical(self, "失败", f"导出失败：\n{err[:500]}")

    def _export_settlement(self):
        """导出结算单汇总"""
        # 确保完整数据已加载
        if not self._ensure_details_loaded():
            QMessageBox.warning(self, "提示", "暂无数据，请先完成计算")
            return

        # 找一个可写目录作为默认（打包后用 exe 所在目录）
        if getattr(sys, 'frozen', False):
            app_root = os.path.dirname(sys.executable)
        else:
            app_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        default_settle_dir = os.path.join(app_root, "data", "exports")
        try:
            os.makedirs(default_settle_dir, exist_ok=True)
        except Exception:
            default_settle_dir = os.getcwd()

        # 选择导出目录
        export_dir = QFileDialog.getExistingDirectory(
            self, "选择导出目录", default_settle_dir
        )
        if not export_dir:
            return

        try:
            engine = SettlementEngine()
            type_ = self._get_current_settlement_type()
            self._show_export_overlay()
            if type_ == "station":
                data = engine.calculate_station_settlement(self.current_details)
            elif type_ == "contract":
                data = engine.calculate_contract_settlement(self.current_details)
            else:  # monthly
                data = engine.calculate_monthly_settlement(self.current_details)

            service = ExportService(export_dir=export_dir)
            file_path = service.export_settlement(data, type_, record_id=self.current_record_id)
            self._hide_export_overlay()
            self.statusBar.showMessage(f"导出完成：{file_path}")
            QMessageBox.information(self, "成功", f"已导出到：\n{file_path}")
        except Exception as e:
            self._hide_export_overlay()
            QMessageBox.critical(self, "失败", f"导出失败：{e}")

    def _export_settlement_detail(self):
        """导出选中行的结算明细"""
        # 确保完整数据已加载
        if not self._ensure_details_loaded():
            QMessageBox.warning(self, "提示", "暂无数据，请先完成计算")
            return

        # 获取选中的行
        selected_row = self.settlement_table.currentRow()
        if selected_row < 0:
            # 如果没有选中，尝试获取第一行
            if self.settlement_table.rowCount() > 0:
                selected_row = 0
            else:
                QMessageBox.warning(self, "提示", "请先选择要导出的结算对象")
                return

        # 获取编码（第一列）
        item = self.settlement_table.item(selected_row, 0)
        if not item:
            QMessageBox.warning(self, "提示", "无法获取结算对象编码")
            return
        group_key = item.text()

        # 选择导出目录
        export_dir = QFileDialog.getExistingDirectory(
            self, "选择导出目录", os.path.expanduser("~/Desktop")
        )
        if not export_dir:
            return

        try:
            service = ExportService(export_dir=export_dir)
            type_ = self._get_current_settlement_type()
            self._show_export_overlay()

            file_path = service.export_settlement_details(
                self.current_details, type_, group_key, export_dir, self.current_record_id
            )
            self._hide_export_overlay()
            self.statusBar.showMessage(f"导出完成：{file_path}")
            QMessageBox.information(self, "成功", f"已导出到：\n{file_path}")
        except Exception as e:
            self._hide_export_overlay()
            QMessageBox.critical(self, "失败", f"导出失败：{e}")

    def _get_current_settlement_type(self) -> str:
        """获取当前结算类型 - 网点/承包区/月结客户"""
        if self.station_btn.isChecked():
            return "station"
        elif self.contract_btn.isChecked():
            return "contract"
        else:  # monthly
            return "monthly"

    def _load_history(self):
        """加载历史记录"""
        session = get_session()
        try:
            records = session.query(FeeRecord).order_by(FeeRecord.id.desc()).limit(100).all()
            self.history_table.setRowCount(len(records))
            for i, r in enumerate(records):
                self.history_table.setItem(i, 0, QTableWidgetItem(str(r.id)))
                self.history_table.setItem(i, 1, QTableWidgetItem(r.file_name))
                self.history_table.setItem(i, 2, QTableWidgetItem(str(r.total_rows)))
                self.history_table.setItem(i, 3, QTableWidgetItem(f"{float(r.total_fee or 0):.2f}"))
                self.history_table.setItem(i, 4, QTableWidgetItem(r.status))
                self.history_table.setItem(i, 5, QTableWidgetItem(r.created_at.strftime("%Y-%m-%d %H:%M") if r.created_at else ""))
        finally:
            session.close()

    def _load_history_detail(self, row, col):
        """双击加载历史详情"""
        record_id = int(self.history_table.item(row, 0).text())
        self.current_record_id = record_id
        self.tabs.setCurrentIndex(1)
        self._load_result(record_id)

    def _clear_history(self):
        """清空所有历史记录（场景3：一键重置，不碰规则配置）"""
        from sqlalchemy import text

        session = get_session()
        try:
            # 先统计两条数据量，给用户明确提示
            conn = session.connection().connection
            cur = conn.cursor()
            cur.execute("SELECT COUNT(*) FROM fee_details")
            detail_count = cur.fetchone()[0]
            cur.close()

            record_count = session.query(FeeRecord).count()

            if record_count == 0 and detail_count == 0:
                QMessageBox.information(self, "提示", "没有历史记录可清空")
                return

            reply = QMessageBox.question(
                self,
                "确认清空",
                f"将删除以下数据（用原生SQL直接删除，速度快）：\n\n"
                f"  • 历史记录概要：{record_count:,} 条（fee_records 表）\n"
                f"  • 计算明细：{detail_count:,} 条（fee_details 表）\n\n"
                f"✅ 保留：客户/网点/区域/全局 计费规则配置（fee_rules）\n"
                f"✅ 保留：其他系统配置\n\n"
                f"⚠️ 此操作不可恢复，确定继续吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

            # 两步删除：先清明细（大表），再清记录。用原生SQL，跳过ORM对象级联，快得多
            cursor = conn.cursor()
            try:
                cursor.execute("DELETE FROM fee_details")
                deleted_detail = cursor.rowcount
                cursor.execute("DELETE FROM fee_records")
                deleted_record = cursor.rowcount
                conn.commit()
            finally:
                cursor.close()

            # 清空内存缓存
            self.all_record_ids = []
            if hasattr(self, "_details_cache"):
                self._details_cache = {}

            # 刷新界面
            self.history_table.setRowCount(0)
            if hasattr(self, "result_table"):
                self.result_table.setRowCount(0)
            if hasattr(self, "summary_labels"):
                for key in self.summary_labels:
                    self.summary_labels[key].setText("0")

            self.statusBar.showMessage(f"已清空：{deleted_record:,} 条记录 / {deleted_detail:,} 条明细")
            QMessageBox.information(
                self,
                "成功",
                f"已清空历史数据：\n\n"
                f"  • 记录：{deleted_record:,} 条\n"
                f"  • 明细：{deleted_detail:,} 条\n\n"
                f"✅ 规则配置未受影响"
            )
        except Exception as e:
            try:
                session.rollback()
            except Exception:
                pass
            self.statusBar.showMessage("清空失败")
            QMessageBox.critical(self, "失败", f"清空失败：{e}")
        finally:
            session.close()
