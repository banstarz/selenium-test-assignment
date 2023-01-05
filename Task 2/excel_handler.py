import win32com.client
from dataclasses import astuple
import os

from scraper import CourtCase


class ExcelHandler:

    def __init__(self, filename: str = 'book.xlsx'):
        self.filename = filename
        self.data_sheet = None

    def __enter__(self):
        self.xl_app = win32com.client.Dispatch("Excel.Application")
        self.xl_app.Visible = True

        self.wb = self.xl_app.Workbooks.Add()

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.xl_app.DisplayAlerts = False
        self.wb.Close(True, os.path.join(os.getcwd(), self.filename))
        self.xl_app.quit()

    def write_data_to_excel(self,
                            court_cases: list[CourtCase],
                            sheet_name: str = 'Data') -> None:
        self.data_sheet = self.wb.Sheets(1)
        self.data_sheet.name = sheet_name

        headers = court_cases[0].headers

        self._write_headers(headers)
        self._write_data(court_cases)

        self.customize_display_settings()

    def _write_headers(self, headers: tuple[str]) -> None:
        sheet_range = self.data_sheet.Range(self.data_sheet.Cells(1, 1),
                                            self.data_sheet.Cells(1, len(headers)))
        sheet_range.Value = headers

    def _write_data(self, court_cases: list[CourtCase]) -> None:
        headers = court_cases[0].headers
        line_number_shift = 2
        for num, court_case in enumerate(court_cases):
            sheet_range = self.data_sheet.Range(self.data_sheet.Cells(num + line_number_shift, 1),
                                                self.data_sheet.Cells(num + line_number_shift, len(headers)))
            sheet_range.Value = astuple(court_case)

    def customize_display_settings(self):
        self.data_sheet.Columns(2).ColumnWidth = 70

        self.data_sheet.Columns.AutoFit()
        self.data_sheet.Rows.AutoFit()

        self.data_sheet.Rows(1).RowHeight = 30
        self.data_sheet.Range("A2").Select()
        self.xl_app.ActiveWorkbook.Windows(1).FreezePanes = True
