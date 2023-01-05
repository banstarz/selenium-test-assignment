from scraper import CourtCaseScraper
from excel_handler import ExcelHandler

with CourtCaseScraper() as scraper:
    court_cases = scraper.get_data_from_pages()

with ExcelHandler('book.xlsx') as handler:
    handler.write_data_to_excel(court_cases, 'Data')



