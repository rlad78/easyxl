from typing import Optional, Collection
from pathlib import Path

from openpyxl.worksheet.worksheet import Worksheet

from easyxl.base import (
    NewExcelTable,
    SupportsRange,
    TableData,
    WorkbookEditorBase,
)
from easyxl.reader import WorkbookReader


class WorkbookEditor(WorkbookReader):
    def add_table(
        self,
        ws: Worksheet,
        range: SupportsRange,
        categories: Collection[str],
        name: Optional[str] = None,
        initial_data: Optional[TableData] = None,
    ) -> NewExcelTable:
        return NewExcelTable(ws, range, categories, name, initial_data)

    def convert_ws_to_table(
        self,
        ws: Worksheet,
        categories: Optional[Collection[str]] = None,
        name: Optional[str] = None,
    ) -> NewExcelTable:
        ws_data_range = self.get_worksheet_data_range(ws)
        return NewExcelTable(ws, ws_data_range, categories=categories, name=name)

    def save(self, file_path: Optional[Path] = None) -> None:
        if file_path is None:
            self.wb.save(self.file_path)
        else:
            self.wb.save(file_path)
