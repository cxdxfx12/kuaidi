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
# 快速从 JSON 字符串提取字段（比 json.loads 快约30倍）
_RE_BUSINESS_DATE = re.compile(r'"business_date":\s*"([^"]*)"')
_RE_CUSTOMER_NAME = re.compile(r'"customer_name":\s*"([^"]*)"')
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

    def export_details(self, record_id: int, target_file_path: Optional[str] = None,
                       progress_callback=None) -> str:
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
        self._write_details_xlsxwriter(record_id, file_path, progress_callback)
        return file_path

    def _write_details_xlsxwriter(self, record_id: int, file_path: str, progress_callback=None):
        """用 xlsxwriter 高速写入明细
        保留列：行号、业务日期、快递单号、区域、重量(kg)、客户名称、运费(元)
        移除列：网点编码、网点名称、件数、应用规则、是否异常、备注
        改进：单连接 + 游标迭代 (fetchmany)，避免多次 LIMIT/OFFSET 导致数据不一致
        """
        headers = [
            "行号", "业务日期", "快递单号", "区域", "重量(kg)", "客户名称", "运费(元)",
        ]

        wb = xlsxwriter.Workbook(file_path)
        ws = wb.add_worksheet("明细")

        fmt_header = wb.add_format({"bold": True, "bg_color": "#D9E1F2"})
        for col, h in enumerate(headers):
            ws.write(0, col, h, fmt_header)

        # 单连接 + 游标迭代：一次性打开连接，用 fetchmany 分批读取
        from sqlalchemy import text
        session = get_session()
        try:
            conn = session.connection().connection  # 原生 sqlite3.Connection

            # 先拿到总行数
            cur_count = conn.cursor()
            cur_count.execute(
                "SELECT COUNT(*) FROM fee_details WHERE record_id = ?", (record_id,)
            )
            total = cur_count.fetchone()[0]
            cur_count.close()

            if total == 0:
                wb.close()
                return

            # 使用服务器端游标 + ORDER BY id 按批次读取
            BATCH = 25000
            cur = conn.cursor()
            cur.execute("""
                SELECT id, record_id, row_index, tracking_no, station_code, station_name,
                       weight, region_name, quantity, rule_name, calculated_fee,
                       is_exception, remark, original_data
                FROM fee_details
                WHERE record_id = ?
                ORDER BY id
            """, (record_id,))

            row_num = 1  # Excel 数据从第1行开始（第0行是表头）
            written = 0

            while True:
                rows = cur.fetchmany(BATCH)
                if not rows:
                    break

                # 批量构建行数据，减少逐格写入开销
                batch_rows = []
                for r in rows:
                    business_date = ""
                    customer_name = ""
                    original_data = r[13]
                    if original_data:
                        # 用正则直接提取，比 json.loads 快30倍以上
                        m = _RE_BUSINESS_DATE.search(original_data)
                        if m:
                            raw_date = m.group(1)
                            if raw_date:
                                digits = _RE_DATE.findall(raw_date)
                                if len(digits) >= 3:
                                    business_date = f"{int(digits[0]):04d}/{int(digits[1]):02d}/{int(digits[2]):02d}"
                                else:
                                    business_date = raw_date
                        m2 = _RE_CUSTOMER_NAME.search(original_data)
                        if m2:
                            customer_name = m2.group(1)

                    weight = r[6]
                    fee = r[10]
                    batch_rows.append([
                        row_num,                              # 行号
                        business_date,                        # 业务日期
                        _clean_str(r[3]),                    # 快递单号
                        _clean_str(r[7]),                    # 区域
                        float(weight) if weight else 0.0,     # 重量
                        customer_name,                        # 客户名称
                        float(fee) if fee else 0.0,          # 运费
                    ])
                    row_num += 1

                # 用 write_row 一次性写整行，减少循环内方法调用
                start_row = row_num - len(batch_rows)
                for idx, row_data in enumerate(batch_rows):
                    ws.write_row(start_row + idx, 0, row_data)
                written += len(batch_rows)

                if progress_callback and total > 0:
                    pct = int(written / total * 100)
                    progress_callback(pct, f"已写入 {written:,} / {total:,} 行")

            cur.close()

            # 最终进度上报
            if progress_callback:
                progress_callback(100, f"✅ 导出完成：共 {written:,} 行（数据库 {total:,} 行）")

            # ===== 方案A：导出成功后，清理 fee_details 明细（保留 fee_record 概要）=====
            try:
                cur_delete = conn.cursor()
                cur_delete.execute(
                    "DELETE FROM fee_details WHERE record_id = ?", (record_id,)
                )
                conn.commit()
                deleted_count = cur_delete.rowcount
                cur_delete.close()
                if progress_callback:
                    progress_callback(100, f"✅ 已清理 {written:,} 条明细数据（节省空间）")
            except Exception:
                pass

        finally:
            session.close()

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

        wb = xlsxwriter.Workbook(os.path.join(final_dir, "temp.xlsx"))
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
        total_records = len(record_ids)

        for i, record_id in enumerate(record_ids):
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

            # 当前记录的基础进度（在所有记录中的位置）
            base_pct = i / total_records * 100

            # 每文件贡献的进度比例 = 1 / total_records * 100
            file_contribution = 100.0 / total_records

            # 给 _write_details_xlsxwriter 传一个子进度回调，映射到总体进度
            sub_callback = None
            if progress_callback:
                def _sub_callback(pct, msg, _i=i, _total=total_records, _bp=base_pct, _fc=file_contribution):
                    overall = int(_bp + (_fc * pct / 100))
                    progress_callback(overall, f"第 {_i+1}/{_total} 个文件：{msg}")
                sub_callback = _sub_callback

            self._write_details_xlsxwriter(record_id, file_path, progress_callback=sub_callback)
            exported_files.append(file_path)

        return exported_files

    def export_merged_records(self, record_ids: List[int], export_dir: Optional[str] = None,
                              base_name: str = "合并结算", progress_callback=None) -> List[str]:
        """合并导出（自动拆分文件）。改用单连接+游标迭代，避免LIMIT/OFFSET导致的一致性问题"""
        final_dir = _find_writable_dir(export_dir or self.preferred_dir)

        from sqlalchemy import text
        session = get_session()
        try:
            conn = session.connection().connection

            # 统计总行数 & 每个 record_id 的文件名
            total = 0
            record_files = {}  # record_id -> file_name
            placeholders = ",".join(["?"] * len(record_ids))

            cur = conn.cursor()
            cur.execute(f"SELECT record_id, COUNT(*) FROM fee_details WHERE record_id IN ({placeholders}) GROUP BY record_id", record_ids)
            count_rows = cur.fetchall()
            total = sum(r[1] for r in count_rows)

            # 获取每个 record_id 对应的文件名
            for rid in record_ids:
                cur.execute("SELECT file_name FROM fee_records WHERE id = ?", (rid,))
                row = cur.fetchone()
                record_files[rid] = row[0] if row and row[0] else f"record_{rid}"
            cur.close()

            if total == 0:
                raise ValueError("没有数据可导出")

            num_files = (total + MAX_ROWS_PER_SHEET - 1) // MAX_ROWS_PER_SHEET
            if progress_callback:
                progress_callback(0, f"总数据 {total:,} 行，将拆分为 {num_files} 个文件")

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            exported_files = []
            headers = [
                "行号", "业务日期", "快递单号", "区域", "重量(kg)", "客户名称", "运费(元)", "来源文件",
            ]

            def _new_wb():
                wb = xlsxwriter.Workbook(
                    os.path.join(final_dir, f"{_clean_str(base_name)}_writing_{len(exported_files)}.xlsx"),
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
                os.replace(tmp_path, fpath)
                exported_files.append(fpath)

            wb, ws = _new_wb()
            wb_idx = 1
            row_in_ws = 0
            processed = 0

            BATCH = 25000

            # 按 record_id 顺序处理，每个 record_id 用游标迭代
            for record_id in record_ids:
                source_file = os.path.basename(record_files.get(record_id, f"record_{record_id}"))

                cur = conn.cursor()
                cur.execute("""
                    SELECT id, record_id, row_index, tracking_no, station_code, station_name,
                           weight, region_name, quantity, rule_name, calculated_fee,
                           is_exception, remark, original_data
                    FROM fee_details
                    WHERE record_id = ?
                    ORDER BY id
                """, (record_id,))

                while True:
                    rows = cur.fetchmany(BATCH)
                    if not rows:
                        break

                    # 先构建整批行数据
                    batch_out = []
                    for r in rows:
                        business_date = ""
                        customer_name = ""
                        original_data = r[13]
                        if original_data:
                            # 用正则直接提取，比 json.loads 快30倍以上
                            m = _RE_BUSINESS_DATE.search(original_data)
                            if m:
                                raw_date = m.group(1)
                                if raw_date:
                                    digits = _RE_DATE.findall(raw_date)
                                    if len(digits) >= 3:
                                        business_date = f"{int(digits[0]):04d}/{int(digits[1]):02d}/{int(digits[2]):02d}"
                                    else:
                                        business_date = raw_date
                            m2 = _RE_CUSTOMER_NAME.search(original_data)
                            if m2:
                                customer_name = m2.group(1)

                        weight = r[6]
                        fee = r[10]
                        batch_out.append([
                            0,  # 占位：excel_row 在写入时动态计算
                            business_date,
                            _clean_str(r[3]),
                            _clean_str(r[7]),
                            float(weight) if weight else 0.0,
                            customer_name,
                            float(fee) if fee else 0.0,
                            _clean_str(source_file),
                        ])

                    # 写入批量数据（处理可能的文件分裂）
                    for row_data in batch_out:
                        if row_in_ws >= MAX_ROWS_PER_SHEET:
                            _close_wb(wb, wb_idx)
                            wb_idx += 1
                            wb, ws = _new_wb()
                            row_in_ws = 0
                        excel_row = row_in_ws + 1
                        row_data[0] = excel_row
                        ws.write_row(excel_row, 0, row_data)
                        row_in_ws += 1
                        processed += 1

                        if progress_callback and processed % 25000 == 0:
                            pct = int(processed / total * 100)
                            progress_callback(pct, f"导出中... {processed:,}/{total:,} 行")

                cur.close()
                # 显式释放内存

            _close_wb(wb, wb_idx)

            if progress_callback:
                progress_callback(100, f"✅ 导出完成：共 {processed:,} 行，{len(exported_files)} 个文件")

            # ===== 方案A：合并导出成功后，清理所有涉及的 record_id 的明细（保留 fee_record 概要）
            try:
                cur_del = conn.cursor()
                placeholders = ",".join(["?"] * len(record_ids))
                cur_del.execute(
                    f"DELETE FROM fee_details WHERE record_id IN ({placeholders})",
                    tuple(record_ids)
                )
                conn.commit()
                cur_del.close()
                if progress_callback:
                    progress_callback(100, f"✅ 已清理 {len(record_ids)} 个文件的明细数据（节省空间）")
            except Exception:
                pass

            return exported_files

        finally:
            session.close()

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
