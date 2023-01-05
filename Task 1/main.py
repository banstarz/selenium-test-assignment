from scraper import CurrencyScraper
from excel_handler import ExcelHandler


with CurrencyScraper(date_range=1) as scraper:
    currency_data = scraper.get_currency_data()

with ExcelHandler(filename='book.xlsx') as handler:
    handler.write_data_to_excel(currency_data)
    handler.create_pivot()
