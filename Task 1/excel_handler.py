import win32com.client
from win32com.client import CDispatch
from dataclasses import astuple
import os

from scraper import Currency


class ExcelHandler:

    def __init__(self, filename: str = 'book.xlsx'):
        self.filename = filename
        self.data_sheet = None
        self.pivot_sheet = None

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
                            currencies: list[Currency],
                            sheet_name: str = 'Data') -> None:
        self.data_sheet = self.wb.Sheets(1)
        self.data_sheet.name = sheet_name

        headers = currencies[0].headers

        self._write_headers(headers)
        self._write_data(currencies)

    def _write_headers(self, headers: tuple[str]) -> None:
        sheet_range = self.data_sheet.Range(self.data_sheet.Cells(1, 1),
                                            self.data_sheet.Cells(1, len(headers)))
        sheet_range.Value = headers

    def _write_data(self, currencies: list[Currency]) -> None:
        headers = currencies[0].headers
        line_number_shift = 2
        for num, currency in enumerate(currencies):
            sheet_range = self.data_sheet.Range(self.data_sheet.Cells(num + line_number_shift, 1),
                                                self.data_sheet.Cells(num + line_number_shift, len(headers)))
            sheet_range.Value = astuple(currency)

    def create_pivot(self) -> None:
        self.pivot_sheet = self.wb.Sheets.Add(Before=None, After=self.wb.Sheets(self.wb.Sheets.count))
        self.pivot_sheet.name = 'Report'

        pt_cache = self.wb.PivotCaches().Create(1, self.data_sheet.Range('A1').CurrentRegion)

        pt = pt_cache.CreatePivotTable(self.pivot_sheet.Range('B3'), 'My Report')
        pt.RowGrand = False
        pt.ColumnGrand = False

        self._insert_fields(pt)

    @staticmethod
    def _insert_fields(pt: CDispatch) -> None:
        pt.PivotFields('Дата').Orientation = 3
        pt.PivotFields('Дата').Position = 1

        pt.PivotFields('Валюта').Orientation = 1
        pt.PivotFields('Валюта').Position = 1

        pt.AddDataField(pt.PivotFields('Курс'), 'Курс Min', -4139)
        pt.AddDataField(pt.PivotFields('Курс'), 'Курс Max', -4136)
        pt.AddDataField(pt.PivotFields('Курс'), 'Курс Avg', -4106)