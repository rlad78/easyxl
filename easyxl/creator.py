from typing import Any
from pathlib import Path

from openpyxl import Workbook

from easyxl.base import ExcelRange, ExcelTable, WorkbookCreatorBase
from easyxl.editor import ExcelEditor


class WorkbookCreator(ExcelEditor):
    def __init__(self) -> None:
        self.wb: Workbook = Workbook()

    def save(self, save_path: Path) -> None:
        self.wb.save(save_path)
