"""
申通运费Excel解析服务
适配申通快递运费结算Excel的常见列名
支持：pandas (默认)、calamine (高速)、csv
"""
import pandas as pd
import os
from typing import Dict, List
from app.services.column_matcher import ColumnMatcher

# 尝试导入 calamine（Rust实现的快速Excel读取库）
try:
    from calamine import read_excel as calamine_read
    CALAMINE_AVAILABLE = True
except ImportError:
    CALAMINE_AVAILABLE = False


class ExcelParser:
    """Excel解析器 - 支持多种读取引擎"""

    def __init__(self):
        self.matcher = ColumnMatcher()
        self._use_calamine = CALAMINE_AVAILABLE

    def parse(self, file_path: str, sheet_name=None, row_callback=None) -> Dict:
        """
        解析Excel文件（支持分块读取 + 进度回调）
        自动选择最快的读取引擎：calamine > pandas
        :param file_path: 文件路径
        :param sheet_name: 可选，指定Sheet名
        :param row_callback: 可选，进度回调 callback(percent: int)  0-100
        :return: {"columns": 所有列名, "data": 数据列表, "row_count": 总行数, "matched": 列名匹配结果}
        """
        def report(percent):
            if row_callback:
                try:
                    row_callback(int(percent))
                except Exception:
                    pass

        file_ext = os.path.splitext(file_path)[1].lower()

        if file_ext == ".csv":
            return self._parse_csv(file_path, row_callback)
        elif file_ext in (".xlsx", ".xls"):
            # 优先使用 calamine（快3-5倍），不支持xlsb
            if self._use_calamine:
                try:
                    return self._parse_calamine(file_path, sheet_name, row_callback)
                except Exception:
                    # calamine失败时回退到pandas
                    pass
            return self._parse_pandas(file_path, sheet_name, row_callback)
        else:
            # 未知格式，尝试pandas
            return self._parse_pandas(file_path, sheet_name, row_callback)

    def _parse_calamine(self, file_path: str, sheet_name, row_callback) -> Dict:
        """
        使用 calamine 快速解析Excel（Rust实现，比pandas快3-5倍）
        calamine 特点：
        - 读取xlsx/xls/xlsb格式
        - 不支持.xlsx/.xls时，回退到pandas
        """
        def report(percent):
            if row_callback:
                try:
                    row_callback(int(percent))
                except Exception:
                    pass

        report(5)

        # 获取sheet列表
        sheets = calamine_read.sheets(file_path)
        sheet_names = [s.name for s in sheets]

        if not sheet_names:
            raise ValueError("Excel文件中没有找到Sheet")

        sheet_to_read = sheet_name or sheet_names[0]
        report(10)

        # 使用calamine读取
        data = calamine_read(file_path, sheet=sheet_to_read)

        # 获取列名（第一行）
        columns = [str(c) if c is not None else "" for c in data[0]]

        report(40)

        row_count = len(data) - 1  # 减去表头行
        report(50)

        # 转成dict列表
        rows_data = []
        header = data[0]
        for row_idx, row in enumerate(data[1:], start=1):
            row_dict = {}
            for col_idx, cell in enumerate(row):
                if col_idx < len(header):
                    row_dict[header[col_idx]] = str(cell) if cell is not None else ""
                else:
                    row_dict[f"_col_{col_idx}"] = str(cell) if cell is not None else ""
            rows_data.append(row_dict)

            # 进度回调
            if row_callback and row_idx % 10000 == 0:
                progress = 50 + int(row_idx / row_count * 45)
                report(progress)

        report(95)

        # 匹配列名
        matched = self.matcher.match_columns(columns)
        report(100)

        return {
            "columns": columns,
            "data": rows_data,
            "row_count": row_count,
            "matched": matched,
            "file_name": os.path.basename(file_path)
        }

    def _parse_pandas(self, file_path: str, sheet_name, row_callback) -> Dict:
        """使用 pandas 解析Excel（回退方案）"""
        def report(percent):
            if row_callback:
                try:
                    row_callback(int(percent))
                except Exception:
                    pass

        report(5)

        excel_file = pd.ExcelFile(file_path)
        sheet_to_read = sheet_name or excel_file.sheet_names[0]

        report(15)
        df = excel_file.parse(sheet_name=sheet_to_read, dtype=str)
        report(50)

        columns = list(df.columns)
        row_count = len(df)

        # 匹配列名
        matched = self.matcher.match_columns(columns)
        report(70)

        # 转换数据
        data = df.fillna("").to_dict("records")
        report(100)

        return {
            "columns": columns,
            "data": data,
            "row_count": row_count,
            "matched": matched,
            "file_name": os.path.basename(file_path)
        }

    def _parse_csv(self, file_path: str, row_callback) -> Dict:
        """解析CSV文件"""
        def report(percent):
            if row_callback:
                try:
                    row_callback(int(percent))
                except Exception:
                    pass

        report(5)

        # CSV文件一般较小，直接读取
        df = pd.read_csv(file_path, dtype=str)
        report(50)

        columns = list(df.columns)
        row_count = len(df)

        # 匹配列名
        matched = self.matcher.match_columns(columns)
        report(80)

        # 转换数据
        data = df.fillna("").to_dict("records")
        report(100)

        return {
            "columns": columns,
            "data": data,
            "row_count": row_count,
            "matched": matched,
            "file_name": os.path.basename(file_path)
        }

    def preview(self, file_path: str, rows: int = 10) -> Dict:
        """预览Excel前N行"""
        file_ext = os.path.splitext(file_path)[1].lower()

        if file_ext == ".csv":
            df = pd.read_csv(file_path, dtype=str, nrows=rows)
        elif self._use_calamine:
            try:
                sheets = calamine_read.sheets(file_path)
                if sheets:
                    data = calamine_read(file_path, sheet=sheets[0].name, nrows=rows+1)
                    if data and len(data) > 0:
                        header = data[0]
                        rows_data = []
                        for row in data[1:rows+1]:
                            row_dict = {}
                            for col_idx, cell in enumerate(row):
                                if col_idx < len(header):
                                    row_dict[header[col_idx]] = str(cell) if cell is not None else ""
                            rows_data.append(row_dict)
                        return {
                            "columns": [str(c) if c else "" for c in header],
                            "data": rows_data
                        }
            except Exception:
                pass
            # 回退到pandas
            df = pd.read_excel(file_path, dtype=str, nrows=rows)
        else:
            df = pd.read_excel(file_path, dtype=str, nrows=rows)

        return {
            "columns": list(df.columns),
            "data": df.fillna("").to_dict("records")
        }
