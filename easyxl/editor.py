from typing import Optional, Collection
from pathlib import Path

from openpyxl.worksheet.worksheet import Worksheet

from easyxl.base import (
    ExcelRangeWritable,
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
        categories: Collection[str],
        range: Optional[SupportsRange] = None,
        name: Optional[str] = None,
        initial_data: Optional[TableData] = None,
    ) -> NewExcelTable:
        if not range:
            if not categories:
                raise Exception("Cannot create a table with no header categories!")

            width = len(categories)
            length = len(initial_data) if initial_data is not None else 1

            # don't adjust for header row, NewExcelTable does it for us
            range = ExcelRangeWritable(ws, "A1:A1")
            range.expand(right=(width - 1), down=(length - 1))

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
