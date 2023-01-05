import datetime
import random
import time

from selenium import webdriver
from fake_useragent import UserAgent
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from dataclasses import dataclass


@dataclass
class Currency:
    date: str
    code: str
    name: str
    exchange_rate: float
    denomination: str
    difference: float

    @property
    def headers(self):
        return (
            'Дата',
            'Код',
            'Валюта',
            'Курс',
            'Номинал',
            'Изменение',
        )


class CurrencyScraper:

    def __init__(self, start_date: str = '', date_range: int = 31):
        self.date = datetime.datetime.strptime(start_date, '%Y-%m-%d').date() if start_date \
            else datetime.date.today()
        self.date_range = date_range
        self.url = 'https://ratestats.com/'
        self.scraped_result = []

    def __enter__(self):
        user_agent = UserAgent().chrome
        options = webdriver.ChromeOptions()
        options.add_argument(f'user-agent={user_agent}')
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),
                                       options=options)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.driver.quit()

    def get_currency_data(self) -> list[Currency]:
        self.driver.get(self.url)
        one_day_shift = datetime.timedelta(days=1)
        self.scraped_result = []
        for _ in range(self.date_range):
            self._get_date_currency_page(self.date)
            self.scraped_result.extend(self._scrap_currency())
            self.date -= one_day_shift

        return self.scraped_result

    def _get_date_currency_page(self, search_date: datetime.date) -> None:
        formatted_date = datetime.datetime.strftime(search_date, '%d%m%Y')
        time.sleep(random.randrange(0, 1))
        calendar_element = self.driver.find_element(By.ID, 'frm-calendar-date')
        calendar_element.send_keys(formatted_date + '\n')

    def _scrap_currency(self) -> list[Currency]:
        currency_table = self.driver.find_element(By.CLASS_NAME, 'b-rates_middle')
        currency_elements = currency_table.find_elements(By.CLASS_NAME, 'b-rates__item')
        collected_page_data = map(self._collect_currency_data, currency_elements)

        return list(collected_page_data)

    def _collect_currency_data(self, elem: WebElement) -> Currency:
        names = elem.find_elements(By.CLASS_NAME, 'b-rates__name')
        amounts = elem.find_elements(By.CLASS_NAME, 'b-rates__amount')
        currency = Currency(
            date=datetime.datetime.strftime(self.date, '%Y-%m-%d'),
            code=names[0].text,
            name=names[1].find_element(By.TAG_NAME, 'a').get_attribute('title'),
            exchange_rate=float(amounts[0].text.replace(',', '.')),
            denomination=amounts[1].text,
            difference=float(amounts[2].text.replace(',', '.')),
        )

        return currency
