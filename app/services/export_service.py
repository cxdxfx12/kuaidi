"""
数据导出服务 - 优化版
核心优化：
1. xlsxwriter 替代 openpyxl（写入快 5-10 倍）
2. 原生 SQL 批量读取（跳过 ORM 对象创建开销）
3. 减少 session 开闭次数
4. 优化 JSON 解析和字符串清理
"""
import os
import re
import json
import xlsxwriter  # 放顶部确保 PyInstaller 静态分析能发现
from datetime import datetime
from typing import List, Dict, Optional
from app.models.database import get_session
from app.models.fee_detail import FeeDetail

MAX_ROWS_PER_SHEET = 1000000


def _find_writable_dir(preferred: Optional[str] = None, record_src_dir: Optional[str] = None) -> str:
    candidates = []
    if preferred:
        candidates.append(preferred)
    candidates += [
        os.path.expanduser("~/Desktop"),
        os.path.expanduser("~/桌面"),
        os.path.expanduser("~/Documents"),
        os.path.expanduser("~/文档"),
    ]
    if record_src_dir:
        candidates.append(record_src_dir)
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    candidates += [
        os.path.join(project_root, "data", "exports"),
        os.path.expanduser("~"),
        os.getcwd(),
    ]
    for d in candidates:
        if not d:
            continue
        try:
            os.makedirs(d, exist_ok=True)
            test_file = os.path.join(d, f"_test_write_{os.getpid()}.tmp")
            with open(test_file, "w", encoding="utf-8") as f:
                f.write("ok")
            try:
                os.remove(test_file)
            except Exception:
                pass
            return d
        except Exception:
            continue
    raise RuntimeError("无法找到可写入的导出目录，请检查磁盘权限")


# 预编译正则，避免重复编译开销
_RE_DATE = re.compile(r"\d+")
# 控制字符表（用于字符串清理）
_BAD_CHARS = frozenset(
    i for i in list(range(0, 9)) + list(range(11, 13)) + list(range(14, 32)) + list(range(127, 128))
)


def _clean_str(val) -> str:
    """清理字符串中的控制字符"""
    if val is None:
        return ""
    s = str(val)
    return s.translate({c: None for c in _BAD_CHARS})


class ExportService:
    def __init__(self, export_dir: Optional[str] = None):
        self.preferred_dir = export_dir

    def export_details(self, record_id: int, target_file_path: Optional[str] = None) -> str:
        from app.models.fee_record import FeeRecord

        # 获取源文件信息
        business_date_str = "unknown"
        original_file_name = f"record_{record_id}"
        session = get_session()
        try:
            record = session.query(FeeRecord).filter(FeeRecord.id == record_id).first()
            if record:
                if record.file_name:
                    original_file_name = os.path.splitext(os.path.basename(record.file_name))[0]
                if record.file_path:
                    src_dir = os.path.dirname(record.file_path)
                else:
                    src_dir = None
            else:
                src_dir = None
        finally:
            session.close()

        # 确定路径
        if target_file_path:
            file_path = target_file_path
            os.makedirs(os.path.dirname(file_path) or ".", exist_ok=True)
        else:
            final_dir = _find_writable_dir(self.preferred_dir, src_dir)
            filename = f"{_clean_str(original_file_name)}-帐单已结算{business_date_str}.xlsx"
            file_path = os.path.join(final_dir, filename)

        # 使用 xlsxwriter 高性能写入
        self._write_details_xlsxwriter(record_id, file_path)
        return file_path

    def _write_details_xlsxwriter(self, record_id: int, file_path: str):
        """用 xlsxwriter 高速写入明细"""
        headers = [
            "行号", "业务日期", "快递单号", "网点编码", "网点名称",
            "区域", "重量(kg)", "客户名称", "件数", "运费(元)",
            "应用规则", "是否异常", "备注",
        ]

        wb = xlsxwriter.Workbook(file_path, {"constant_memory": True})
        ws = wb.add_worksheet("明细")

        # 写入表头
        fmt_header = wb.add_format({"bold": True, "bg_color": "#D9E1F2"})
        for col, h in enumerate(headers):
            ws.write(0, col, h, fmt_header)

        # 原生 SQL 批量读取（一次性取完，按 25000 切片避免内存问题）
        session = get_session()
        try:
            # 获取总行数
            total = session.query(FeeDetail).filter(
                FeeDetail.record_id == record_id
            ).count()
        finally:
            session.close()

        # 分批读取，每批 25000 行（用原生 SQL 跳过 ORM 开销）
        BATCH = 25000
        row_num = 1  # Excel 从第1行开始（第0行是表头）

        for offset in range(0, total, BATCH):
            rows = self._fetch_batch_native(record_id, offset, BATCH)
            for r in rows:
                business_date = ""
                customer_name = ""
                original_data = r[13]  # original_data 字段
                if original_data:
                    try:
                        od = json.loads(original_data) if isinstance(original_data, str) else original_data
                        raw_date = od.get("business_date", "")
                        if raw_date:
                            digits = _RE_DATE.findall(str(raw_date))
                            if len(digits) >= 3:
                                business_date = f"{int(digits[0]):04d}/{int(digits[1]):02d}/{int(digits[2]):02d}"
                            else:
                                business_date = str(raw_date)
                        customer_name = str(od.get("customer_name", "") or "")
                    except Exception:
                        pass

                is_exception = "是" if r[11] else "否"
                weight = r[6]
                quantity = r[8]
                fee = r[10]

                ws.write(row_num, 0, r[2])   # row_index
                ws.write(row_num, 1, business_date)
                ws.write(row_num, 2, _clean_str(r[3]))   # tracking_no
                ws.write(row_num, 3, _clean_str(r[4]))   # station_code
                ws.write(row_num, 4, _clean_str(r[5]))   # station_name
                ws.write(row_num, 5, _clean_str(r[7]))   # region_name
                ws.write(row_num, 6, float(weight) if weight else 0.0)
                ws.write(row_num, 7, customer_name)
                ws.write(row_num, 8, int(quantity) if quantity else 0)
                ws.write(row_num, 9, float(fee) if fee else 0.0)
                ws.write(row_num, 10, _clean_str(r[9]))   # rule_name
                ws.write(row_num, 11, is_exception)
                ws.write(row_num, 12, _clean_str(r[12] or ""))
                row_num += 1
            # 每批处理完后主动释放
            del rows

        wb.close()

    def _fetch_batch_native(self, record_id: int, offset: int, limit: int) -> List[tuple]:
        """用原生 SQL 高效批量读取，跳过 ORM 对象创建"""
        from sqlalchemy import text
        session = get_session()
        try:
            # 直接用原始 SQL，按字段索引读取（避免 ORM 对象创建）
            # FeeDetail 字段顺序：id, record_id, row_index, tracking_no, station_code, station_name,
            #                       weight, region_name, quantity, rule_name, calculated_fee,
            #                       is_exception, remark, original_data
            sql = text(f"""
                SELECT id, record_id, row_index, tracking_no, station_code, station_name,
                       weight, region_name, quantity, rule_name, calculated_fee,
                       is_exception, remark, original_data
                FROM fee_details
                WHERE record_id = :rid
                ORDER BY id
                LIMIT :lim OFFSET :off
            """)
            result = session.execute(sql, {"rid": record_id, "lim": limit, "off": offset})
            rows = result.fetchall()
            return rows
        finally:
            session.close()

    def export_settlement(self, settlement_data: List[Dict], settlement_type: str = "station",
                          export_dir: Optional[str] = None, record_id: Optional[int] = None) -> str:
        final_dir = _find_writable_dir(export_dir or self.preferred_dir)
        type_names = {"station": "网点", "contract": "承包区", "monthly": "月结客户"}
        type_name = type_names.get(settlement_type, "结算")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        record_suffix = f"_{record_id}" if record_id else ""
        filename = f"{type_name}结算单{record_suffix}_{timestamp}.xlsx"
        file_path = os.path.join(final_dir, filename)

        wb = xlsxwriter.Workbook(file_path)
        ws = wb.add_worksheet(_clean_str(type_name))
        fmt_header = wb.add_format({"bold": True, "bg_color": "#D9E1F2"})

        if settlement_data:
            headers = list(settlement_data[0].keys())
            for col, h in enumerate(headers):
                ws.write(0, col, _clean_str(h), fmt_header)
            for row_idx, row in enumerate(settlement_data):
                for col_idx, h in enumerate(headers):
                    v = row.get(h, "")
                    if isinstance(v, float):
                        ws.write(row_idx + 1, col_idx, v)
                    elif isinstance(v, int):
                        ws.write(row_idx + 1, col_idx, v)
                    else:
                        ws.write(row_idx + 1, col_idx, _clean_str(v))

        wb.close()
        return file_path

    def export_settlement_details(self, details: List[FeeDetail], settlement_type: str,
                                  group_key: str, export_dir: Optional[str] = None,
                                  record_id: Optional[int] = None) -> str:
        final_dir = _find_writable_dir(export_dir or self.preferred_dir)
        type_names = {"station": "网点", "contract": "承包区", "monthly": "月结客户"}
        type_name = type_names.get(settlement_type, "结算")

        filtered = []
        for d in details:
            if settlement_type == "station" and d.station_code == group_key:
                filtered.append(d)
            elif settlement_type == "contract":
                station_code = d.station_code or ""
                if len(station_code) >= 3 and station_code[:3] == group_key:
                    filtered.append(d)
            elif settlement_type == "monthly":
                original = {}
                try:
                    if d.original_data:
                        original = d.original_data if isinstance(d.original_data, dict) \
                            else json.loads(d.original_data)
                except Exception:
                    pass
                customer_code = original.get("客户编码", original.get("客户代码", ""))
                if str(customer_code).strip() == group_key:
                    filtered.append(d)

        if not filtered:
            raise ValueError("没有找到匹配的明细数据")

        wb = xlsxwriter.Workbook(os.path.join(final_dir, "temp.xlsx"), {"constant_memory": True})
        ws = wb.add_worksheet(_clean_str(type_name))
        fmt_header = wb.add_format({"bold": True, "bg_color": "#D9E1F2"})

        headers = [
            "行号", "快递单号", "网点编码", "网点名称",
            "区域", "重量(kg)", "件数", "运费(元)",
            "应用规则", "是否异常", "备注",
        ]
        for col, h in enumerate(headers):
            ws.write(0, col, h, fmt_header)

        for row_idx, d in enumerate(filtered):
            ws.write(row_idx + 1, 0, _clean_str(d.row_index))
            ws.write(row_idx + 1, 1, _clean_str(d.tracking_no))
            ws.write(row_idx + 1, 2, _clean_str(d.station_code))
            ws.write(row_idx + 1, 3, _clean_str(d.station_name))
            ws.write(row_idx + 1, 4, _clean_str(d.region_name))
            ws.write(row_idx + 1, 5, float(d.weight or 0))
            ws.write(row_idx + 1, 6, int(d.quantity or 0))
            ws.write(row_idx + 1, 7, float(d.calculated_fee or 0))
            ws.write(row_idx + 1, 8, _clean_str(d.rule_name))
            ws.write(row_idx + 1, 9, "是" if d.is_exception else "否")
            ws.write(row_idx + 1, 10, _clean_str(d.remark or ""))

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        record_suffix = f"_{record_id}" if record_id else ""
        safe_key = str(group_key).replace("/", "_").replace("\\", "_")
        filename = f"{type_name}_{_clean_str(safe_key)}_明细{record_suffix}_{timestamp}.xlsx"
        file_path = os.path.join(final_dir, filename)
        wb.close()

        # 重命名临时文件
        temp_path = os.path.join(final_dir, "temp.xlsx")
        if os.path.exists(temp_path):
            os.replace(temp_path, file_path)
        return file_path

    # ============================================
    # 多记录导出功能（分别导出、合并导出）
    # ============================================

    def export_multiple_records(self, record_ids: List[int], export_dir: Optional[str] = None,
                                progress_callback=None) -> List[str]:
        from app.models.fee_record import FeeRecord

        final_dir = _find_writable_dir(export_dir or self.preferred_dir)
        exported_files = []

        for i, record_id in enumerate(record_ids):
            if progress_callback:
                progress = int((i + 1) / len(record_ids) * 100)
                progress_callback(progress, f"正在导出第 {i+1}/{len(record_ids)} 个文件...")

            session = get_session()
            try:
                record = session.query(FeeRecord).filter(FeeRecord.id == record_id).first()
                if record and record.file_name:
                    original_name = os.path.splitext(record.file_name)[0]
                else:
                    original_name = f"record_{record_id}"
            finally:
                session.close()

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"{_clean_str(original_name)}-帐单已结算_{timestamp}.xlsx"
            file_path = os.path.join(final_dir, filename)

            self._write_details_xlsxwriter(record_id, file_path)
            exported_files.append(file_path)

        return exported_files

    def export_merged_records(self, record_ids: List[int], export_dir: Optional[str] = None,
                              base_name: str = "合并结算", progress_callback=None) -> List[str]:
        final_dir = _find_writable_dir(export_dir or self.preferred_dir)

        # 统计总行数（用原生 count）
        from sqlalchemy import text
        session = get_session()
        try:
            total = session.execute(
                text(f"SELECT COUNT(*) FROM fee_details WHERE record_id IN ({','.join(map(str, record_ids))})")
            ).scalar()
        finally:
            session.close()

        if total == 0:
            raise ValueError("没有数据可导出")

        num_files = (total + MAX_ROWS_PER_SHEET - 1) // MAX_ROWS_PER_SHEET
        if progress_callback:
            progress_callback(0, f"总数据 {total:,} 行，将拆分为 {num_files} 个文件")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        headers = [
            "行号", "业务日期", "快递单号", "网点编码", "网点名称",
            "区域", "重量(kg)", "客户名称", "件数", "运费(元)",
            "应用规则", "是否异常", "备注", "来源文件",
        ]

        # 一次只打开一个文件（用 constant_memory 模式，写完立即关闭缓冲区）
        def _new_wb():
            wb = xlsxwriter.Workbook(
                os.path.join(final_dir, f"{_clean_str(base_name)}_writing_{len(exported_files)}.xlsx"),
                {"constant_memory": True}
            )
            ws = wb.add_worksheet("明细")
            fmt_header = wb.add_format({"bold": True, "bg_color": "#D9E1F2"})
            for col, h in enumerate(headers):
                ws.write(0, col, h, fmt_header)
            return wb, ws

        def _close_wb(wb, wb_idx):
            fname = f"{_clean_str(base_name)}_{wb_idx}_{timestamp}.xlsx"
            fpath = os.path.join(final_dir, fname)
            tmp_path = wb.filename
            wb.close()
            # xlsxwriter 无法直接写入目标路径，用 rename 策略
            os.replace(tmp_path, fpath)
            exported_files.append(fpath)

        wb, ws = _new_wb()
        wb_idx = 1
        row_in_ws = 0
        processed = 0

        for record_id in record_ids:
            # 获取来源文件名
            from app.models.fee_record import FeeRecord
            session = get_session()
            try:
                record = session.query(FeeRecord).filter(FeeRecord.id == record_id).first()
                source_file = record.file_name if record else f"record_{record_id}"
            finally:
                session.close()

            # 获取该记录的行数
            from sqlalchemy import text as _text
            session = get_session()
            try:
                total_rec = session.execute(
                    _text(f"SELECT COUNT(*) FROM fee_details WHERE record_id = {record_id}")
                ).scalar()
            finally:
                session.close()
            if not total_rec:
                continue

            BATCH = 25000
            for offset in range(0, total_rec, BATCH):
                rows = self._fetch_batch_native(record_id, offset, BATCH)
                for r in rows:
                    if row_in_ws >= MAX_ROWS_PER_SHEET:
                        _close_wb(wb, wb_idx)
                        wb_idx += 1
                        wb, ws = _new_wb()
                        row_in_ws = 0

                    business_date = ""
                    customer_name = ""
                    original_data = r[13]
                    if original_data:
                        try:
                            od = json.loads(original_data) if isinstance(original_data, str) else original_data
                            raw_date = od.get("business_date", "")
                            if raw_date:
                                digits = _RE_DATE.findall(str(raw_date))
                                if len(digits) >= 3:
                                    business_date = f"{int(digits[0]):04d}/{int(digits[1]):02d}/{int(digits[2]):02d}"
                                else:
                                    business_date = str(raw_date)
                            customer_name = str(od.get("customer_name", "") or "")
                        except Exception:
                            pass

                    is_exception = "是" if r[11] else "否"
                    excel_row = row_in_ws + 1

                    ws.write(excel_row, 0, r[2])   # row_index
                    ws.write(excel_row, 1, business_date)
                    ws.write(excel_row, 2, _clean_str(r[3]))   # tracking_no
                    ws.write(excel_row, 3, _clean_str(r[4]))   # station_code
                    ws.write(excel_row, 4, _clean_str(r[5]))   # station_name
                    ws.write(excel_row, 5, _clean_str(r[7]))   # region_name
                    ws.write(excel_row, 6, float(r[6]) if r[6] else 0.0)  # weight
                    ws.write(excel_row, 7, customer_name)
                    ws.write(excel_row, 8, int(r[8]) if r[8] else 0)   # quantity
                    ws.write(excel_row, 9, float(r[10]) if r[10] else 0.0)  # calculated_fee
                    ws.write(excel_row, 10, _clean_str(r[9]))   # rule_name
                    ws.write(excel_row, 11, is_exception)
                    ws.write(excel_row, 12, _clean_str(r[12] or ""))
                    ws.write(excel_row, 13, _clean_str(source_file))

                    row_in_ws += 1
                    processed += 1

                    if progress_callback and processed % 10000 == 0:
                        pct = int(processed / total * 100)
                        progress_callback(pct, f"导出中... {processed:,}/{total:,} 行")

                del rows

        _close_wb(wb, wb_idx)

        if progress_callback:
            progress_callback(100, f"导出完成，共 {len(exported_files)} 个文件")

        return exported_files

    def get_record_info(self, record_ids: List[int]) -> List[Dict]:
        from app.models.fee_record import FeeRecord

        session = get_session()
        try:
            result = []
            for record_id in record_ids:
                record = session.query(FeeRecord).filter(FeeRecord.id == record_id).first()
                if record:
                    result.append({
                        "id": record_id,
                        "file_name": record.file_name or f"record_{record_id}",
                        "total_rows": record.total_rows or 0,
                        "success_rows": record.success_rows or 0,
                        "error_rows": record.error_rows or 0,
                        "total_fee": float(record.total_fee or 0),
                    })
            return result
        finally:
            session.close()
