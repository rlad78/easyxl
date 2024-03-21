from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from easyxl.base import ExcelRange, ExcelTable, WorkbookOpenBase, Coordinate


class WorkbookReader(WorkbookOpenBase):
    def get_range_by_expression(
        self, worksheet: Worksheet, range_expression: str
    ) -> ExcelRange:
        return ExcelRange(worksheet, range_expression=range_expression)

    def get_range_by_coordinates(
        self, worksheet: Worksheet, coordinates: tuple[Coordinate, Coordinate]
    ) -> ExcelRange:
        return ExcelRange(worksheet, coordinates=coordinates)

    def get_worksheet_data_range(self, worksheet: Worksheet) -> ExcelRange:
        ws_rows = list(worksheet.rows)
        top_left_coordinate = ws_rows[0][0].coordinate
        bottom_right_coordinate = ws_rows[-1][-1].coordinate

        return ExcelRange(
            worksheet,
            range_expression=f"{top_left_coordinate}:{bottom_right_coordinate}",
        )

    def all_worksheet_tables(self, worksheet: Worksheet) -> dict[str, ExcelTable]:
        worksheet_tables = worksheet.tables
        return {
            name: ExcelTable(worksheet, table)
            for name, table in worksheet_tables.items()
        }

    def all_tables(self) -> dict[str, ExcelTable]:
        workbook_tables = {}
        for worksheet in self.wb.worksheets:
            workbook_tables.update(self.all_worksheet_tables(worksheet))
        return workbook_tables

    def get_table_by_range(
        self, worksheet: Worksheet, range: ExcelRange
    ) -> ExcelTable | None:
        ws_tables = list(self.all_worksheet_tables(worksheet).values())
        matching_tables = [t for t in ws_tables if t.range.issuperset(range)]

        if not matching_tables:
            return None

        if len(matching_tables) > 1:
            raise Exception(
                f'Range {range} matches multiple tables: {", ".join(str(t) for t in matching_tables)}'
            )

        return matching_tables[0]
