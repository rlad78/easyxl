from typing import Optional
from pathlib import Path
from openpyxl import Workbook


class InvalidFile(Exception):
    def __init__(self, file_path: Path, *args: object) -> None:
        super().__init__(*args)
        self.file_uri = str(file_path)

        if len(str(file_path)) > 60:
            self.file_uri = "..." + str(file_path)[-60:]

    def __str__(self) -> str:
        return f"{self.file_uri} is not a valid Excel file"


class InvalidSheet(Exception):
    def __init__(
        self,
        wb: Workbook,
        name: Optional[str] = None,
        index: Optional[int] = None,
        *args: object,
    ) -> None:
        super().__init__(*args)
        self.wb = wb

        if name is not None:
            self._response = f"{name} is not a valid worksheet name in this workbook"
            self._response += "\nWorksheets: " + ", ".join(self.wb.sheetnames)
        elif index is not None:
            self._response = f"{index} is an out of bounds worksheet index (count: {len(self.wb.worksheets)})"
        else:
            self._response = "Could not access invalid worksheet"

    def __str__(self) -> str:
        return self._response


class InvalidRangeFormat(Exception):
    pass
