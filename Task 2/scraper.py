import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.remote.webelement import WebElement
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from fake_useragent import UserAgent
from dataclasses import dataclass
import re


@dataclass
class CourtCase:
    number: str
    parties: str
    state: str
    judge: str
    article: str
    category: str
    cases: str

    @property
    def headers(self):
        return (
            'Номер дела ~ материала',
            'Стороны',
            'Текущее состояние',
            'Судья',
            'Статья',
            'Категория дела',
            'Список дел',
        )



class CourtCaseScraper:

    def __init__(self):
        self.url = 'https://mos-gorsud.ru/search'

    def __enter__(self):
        user_agent = UserAgent().chrome
        options = webdriver.ChromeOptions()
        options.add_argument(f'user-agent={user_agent}')
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                                       options=options)

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.driver.quit()

    def get_data_from_pages(self, max_page_num: int = 0) -> list[CourtCase]:
        self.driver.get(self.url)
        self.driver.maximize_window()
        self._set_search_parameters()
        filter_string = '123 - Об установлении рыночной стоимости земельных участков и отдельных объектов недвижимости'

        final_result = []
        page_num = 1
        while page_num <= (max_page_num or self._get_max_page_number()):
            self._go_to_page(page_num)
            page_data = self._get_data_from_page()
            filtered_data = filter(lambda row: filter_string in row.category, page_data)
            final_result.extend(filtered_data)
            print(len(final_result))
            page_num += 1
            time.sleep(0.1)

        return final_result

    def _set_search_parameters(self) -> None:
        element_ids = [
            'custom-select-court-button',
            'ui-id-9',
            'instance-button',
            'ui-id-46',
            'process-type-select-button',
            'ui-id-51',
            'case-index-search-form-btn',
        ]

        self._click_list_elements(element_ids)

    def _click_list_elements(self, element_ids: list[str]) -> None:
        for elem_id in element_ids:
            time.sleep(0.1)
            element = self.driver.find_element(By.ID, elem_id)
            webdriver.ActionChains(self.driver). \
                scroll_by_amount(15, 15). \
                move_to_element(element). \
                click(element). \
                perform()

    def _get_max_page_number(self) -> int:
        pagination_form = self.driver.find_element(By.ID, 'paginationForm')
        max_page_num = re.search(r'\d+', pagination_form.text).group()
        return int(max_page_num)

    def _get_data_from_page(self) -> list[CourtCase]:
        table = self.driver.find_element(By.CLASS_NAME, 'custom_table')
        table_rows = table.find_elements(By.TAG_NAME, 'tr')[1:]
        page_data = [self._get_data_from_row(row) for row in table_rows]
        return page_data

    @staticmethod
    def _get_data_from_row(elem: WebElement) -> CourtCase:
        record_cells = elem.find_elements(By.TAG_NAME, 'td')
        return CourtCase(
            record_cells[0].text,
            record_cells[1].text,
            record_cells[2].text,
            record_cells[3].text,
            record_cells[4].text,
            record_cells[5].text,
            record_cells[6].find_element(By.TAG_NAME, 'a').get_attribute("textContent"),
        )

    def _go_to_page(self, page_num: int) -> None:
        pagination_form_input = self.driver.find_element(By.ID, 'paginationFormInput')
        pagination_form_input.send_keys(Keys.CONTROL + 'a' + Keys.DELETE)
        pagination_form_input.send_keys(f'{page_num}\n')
