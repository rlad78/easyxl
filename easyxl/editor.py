from typing import Union, Optional, Any, Sequence
from pathlib import Path

from easyxl.exceptions import InvalidFile, InvalidSheet, InvalidRangeFormat

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.worksheet.table import Table
from openpyxl.cell.cell import Cell
from openpyxl.worksheet.cell_range import CellRange


class ExcelRange:
    def __init__(
        self,
        worksheet: Worksheet,
        range_expression: Optional[str] = None,
        coordinates: Optional[Sequence[tuple[int, int]]] = None,
    ) -> None:
        self.ws = worksheet
        if range_expression is not None:
            try:
                self._range = CellRange(range_expression)
            except (ValueError, TypeError):
                raise InvalidRangeFormat(
                    f"{range_expression} is not a valid range expression"
                )
        elif coordinates is not None:
            try:
                self._range = CellRange(
                    min_row=coordinates[0][0],
                    min_col=coordinates[0][1],
                    max_row=coordinates[1][0],
                    max_col=coordinates[1][1],
                )
            except ValueError | TypeError:
                raise InvalidRangeFormat(
                    f"{coordinates[0]}, {coordinates[1]} are invalid coordinate pairs"
                )
        else:
            raise InvalidRangeFormat("No expression or coordinates given")

    def __str__(self) -> str:
        return self._range.__str__()

    @property
    def rows(self) -> list["ExcelRange"]:
        range_rows: list[ExcelRange] = []
        for row in self._range.rows:
            range_rows.append(ExcelRange(self.ws, coordinates=(row[0], row[-1])))
        return range_rows

    @property
    def next_row(self) -> "ExcelRange":
        bottom_row = ExcelRange(
            self.ws, coordinates=(self._range.bottom[0], self._range.bottom[1])
        )
        bottom_row._range.shift(row_shift=1)
        return bottom_row

    @property
    def cells(self) -> list[Cell]:
        return [self.ws.cell(*cell) for cell in self._range.cells]  # type: ignore

    def is_empty(self) -> bool:
        for cell in self.cells:
            if type(cell.value) == str and cell.value.strip():  # type: ignore
                return False
            elif cell.value:
                return False
        else:
            return True

    def expand(self, rows: int = 0, columns: int = 0) -> str:
        self._range.expand(down=rows, right=columns)
        return self.__str__()

    def issubset(self, other: "ExcelRange") -> bool:
        return self._range.issubset(other._range)


class ExcelTable:
    def __init__(self, ws: Worksheet, table_object: Table) -> None:
        self._parent_ws = ws
        self._table = table_object
        self.categories: list[str] = [col.name for col in self._table.tableColumns]

    @property
    def range(self) -> ExcelRange:
        return ExcelRange(self._parent_ws, self._table.ref)

    @property
    def first_free_row(self) -> ExcelRange:
        for row in self.range.rows[1:]:
            if row.is_empty():
                return row
        else:
            return self.range.next_row

    def append(self, data: Union[list[list[Any]], dict[str, list[Any]]]):
        if type(data) == list:
            for data_row in data:
                working_row = self.first_free_row
                for i, cell in enumerate(working_row.cells):
                    cell.value = data_row[i]
                if not working_row.issubset(self.range):
                    new_range = self.range.expand(rows=1)
                    self._table.ref = new_range


class NewExcelTable(
    ExcelTable
): ...  # todo: create class that allows for table creation


class ExcelEditor:
    def __init__(self, file_path: Union[Path, str]) -> None:
        if isinstance(file_path, Path):
            self.file_path = file_path
        else:
            self.file_path = Path(file_path)

        if not self.file_path.is_file or not self.file_path.suffix.startswith(".xl"):
            raise InvalidFile(self.file_path)

        self.wb: Workbook = load_workbook(str(self.file_path))
        self.current_worksheet = self.wb.worksheets[0]

    def change_worksheet(
        self, title: Optional[str] = None, index: Optional[int] = None
    ) -> None:
        if title is not None:
            if title not in self.wb.sheetnames:
                raise InvalidSheet(self.wb, name=title)
            self.current_worksheet = self.wb[title]
        elif index is not None:
            if index >= len(self.wb.worksheets):
                raise InvalidSheet(self.wb, index=index)
            self.current_worksheet = self.wb.worksheets[index]

    @property
    def tables(self) -> list[ExcelTable]:
        return [
            ExcelTable(self.current_worksheet, table)  # type: ignore
            for table in self.current_worksheet.tables.values()  # type: ignore
        ]

    # todo: allow for new table creation

    def save(self) -> None:
        self.wb.save(self.file_path)


class ExcelCreator(ExcelEditor):
    def __init__(self, file_path: Path | str) -> None:
        self.file_path = file_path
        self.wb: Workbook = Workbook()
        self.current_worksheet = self.wb.worksheets[0]
