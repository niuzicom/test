from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter
from typing import Any


class ExcelHelper:
    """Excel 读写修改工具类(基于 openpyxl,支持 .xlsx)"""

    def __init__(self, file_path: str | None = None,
                 sheet_name: str | None = None,
                 data_only: bool = True,
                 keep_vba: bool = False):
        """
        :param file_path:   Excel 路径;传 None 表示新建空白工作簿
        :param sheet_name:  默认操作的工作表名,None 表示活动表
        :param data_only:   读取时 True 取公式结果,False 取公式原文
        :param keep_vba:    是否保留宏(.xlsm)
        """
        self.file_path = file_path
        self._data_only = data_only
        self._keep_vba = keep_vba

        if file_path:
            self.wb = load_workbook(file_path,
                                    data_only=data_only,
                                    keep_vba=keep_vba)
        else:
            self.wb = Workbook()

        self.ws: Worksheet = (self.wb[sheet_name]
                              if sheet_name else self.wb.active)

    # ---------- 上下文管理 ----------
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def close(self):
        """释放文件句柄"""
        try:
            self.wb.close()
        except Exception:
            pass

    # ---------- 工作表管理 ----------
    def sheet_names(self) -> list[str]:
        """返回所有工作表名"""
        return self.wb.sheetnames

    def use_sheet(self, sheet_name: str) -> "ExcelHelper":
        """切换当前操作的工作表(不存在则创建)"""
        self.ws = (self.wb[sheet_name]
                   if sheet_name in self.wb.sheetnames
                   else self.wb.create_sheet(sheet_name))
        return self

    def add_sheet(self, sheet_name: str) -> Worksheet:
        """新建工作表并返回"""
        return self.wb.create_sheet(sheet_name)

    def delete_sheet(self, sheet_name: str):
        """删除工作表"""
        if sheet_name in self.wb.sheetnames:
            self.wb.remove(self.wb[sheet_name])

    # ---------- 读取 ----------
    def read_rows(self, max_row: int | None = None) -> list[tuple]:
        """读取为二维 tuple 列表(含表头)"""
        rows = list(self.ws.iter_rows(values_only=True))
        return rows[:max_row] if max_row else rows

    def read_dicts(self) -> list[dict]:
        """第一行作为表头,返回 dict 列表"""
        rows = list(self.ws.iter_rows(values_only=True))
        if not rows:
            return []
        header = rows[0]
        return [dict(zip(header, r)) for r in rows[1:]]

    def get_cell(self, row: int, col: int) -> Any:
        """读取单个单元格(行列从 1 开始)"""
        return self.ws.cell(row=row, column=col).value

    def find_header_col(self, header_name: str) -> int | None:
        """在第 1 行查找表头所在列号,找不到返回 None"""
        for idx, cell in enumerate(self.ws[1], start=1):
            if cell.value == header_name:
                return idx
        return None

    # ---------- 修改 ----------
    def update_cell(self, row: int, col: int, value: Any) -> "ExcelHelper":
        """修改单个单元格"""
        self.ws.cell(row=row, column=col, value=value)
        return self

    def update_cells(self, updates: list[tuple[int, int, Any]]) -> "ExcelHelper":
        """批量修改:[(行,列,值), ...]"""
        for row, col, value in updates:
            self.ws.cell(row=row, column=col, value=value)
        return self

    def append_row(self, values: list) -> "ExcelHelper":
        """在末尾追加一行"""
        self.ws.append(values)
        return self

    def append_rows(self, rows: list[list]) -> "ExcelHelper":
        """在末尾追加多行"""
        for r in rows:
            self.ws.append(r)
        return self

    def replace_text(self, old: str, new: str) -> int:
        """全表字符串替换,返回替换次数"""
        count = 0
        for row in self.ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and old in cell.value:
                    cell.value = cell.value.replace(old, new)
                    count += 1
        return count

    def update_by_condition(self, header: str, predicate, new_value) -> int:
        """
        按列条件批量修改
        :param header:    表头名
        :param predicate: 判断函数,如 lambda v: isinstance(v,int) and v>30
        :param new_value: 满足条件时写入的新值(也可以是函数 new_value(v))
        :return: 修改的单元格数
        """
        col = self.find_header_col(header)
        if col is None:
            return 0
        count = 0
        for row in range(2, self.ws.max_row + 1):
            cell = self.ws.cell(row=row, column=col)
            if predicate(cell.value):
                cell.value = (new_value(cell.value)
                              if callable(new_value) else new_value)
                count += 1
        return count

    def delete_row(self, row: int, amount: int = 1):
        """删除行"""
        self.ws.delete_rows(row, amount)

    def insert_row(self, row: int, amount: int = 1):
        """插入行"""
        self.ws.insert_rows(row, amount)

    def delete_col(self, col: int, amount: int = 1):
        self.ws.delete_cols(col, amount)

    def insert_col(self, col: int, amount: int = 1):
        self.ws.insert_cols(col, amount)

    # ---------- 保存 ----------
    def save(self, path: str | None = None):
        """保存,path 为 None 则覆盖原文件"""
        target = path or self.file_path
        if not target:
            raise ValueError("未指定保存路径")
        self.wb.save(target)
        self.file_path = target
        return target

    def save_as(self, path: str):
        """另存为新文件"""
        return self.save(path)
