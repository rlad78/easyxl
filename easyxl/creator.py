from typing import Any
from pathlib import Path

from openpyxl import Workbook

from easyxl.editor import WorkbookEditor


class WorkbookCreator(WorkbookEditor):
    def __init__(self) -> None:
        self.wb: Workbook = Workbook()

    def save(self, save_path: Path) -> None:
        self.wb.save(save_path)
