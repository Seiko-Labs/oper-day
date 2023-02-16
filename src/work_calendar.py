import calendar
import csv
import json
import locale
import os
from datetime import date
from typing import List, Dict

import requests
from bs4 import BeautifulSoup, SoupStrainer
from minify_html import minify

from data_structures import DateInfo, WorkStatus


class Serializer:
    def __init__(self, file_name: str, data: List[Dict[str, str]] = None) -> None:
        self.csv_file_path: str = fr'C:\Users\robot.ad\Desktop\oper_day\resourses\{file_name}.csv'
        self.json_file_path: str = fr'C:\Users\robot.ad\Desktop\oper_day\resourses\{file_name}.json'
        self.data: List[Dict[str, str]] = data

    def save(self, _format: str) -> None:
        if _format == 'json':
            self._save_json()
        elif _format == 'csv':
            self._save_csv()
        else:
            raise NotImplementedError(f'Format {_format} is not supported')

    def load(self, _format: str) -> List[Dict[str, str]]:
        if _format == 'json':
            return self._load_json()
        elif _format == 'csv':
            return self._load_csv()
        else:
            raise NotImplementedError(f'Format {_format} is not supported')

    def _save_json(self) -> None:
        with open(file=self.json_file_path, mode='w', encoding='utf-8') as json_file:
            json.dump(self.data, json_file, indent=4, ensure_ascii=False, default=str)

    def _save_csv(self) -> None:
        with open(file=self.csv_file_path, mode='w', encoding='utf-8', newline='') as csv_file:
            csv_writer = csv.DictWriter(csv_file, self.data[0].keys())
            csv_writer.writeheader()
            csv_writer.writerows(self.data)

    def _load_json(self) -> List[Dict[str, str]]:
        with open(file=self.json_file_path, mode='r', encoding='utf-8') as json_file:
            return json.load(json_file)

    def _load_csv(self) -> List[Dict[str, str]]:
        with open(file=self.csv_file_path, mode='r', encoding='utf-8') as csv_file:
            csv_reader = csv.DictReader(csv_file)
            return [dict(row) for row in csv_reader]


class LocaleManager:
    def __init__(self, locale_name: str) -> None:
        self.locale_name: str = locale_name

    def __enter__(self) -> None:
        try:
            locale.setlocale(category=locale.LC_ALL, locale=self.locale_name)
        except locale.Error:
            locale.setlocale(category=locale.LC_ALL, locale='ru')

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        locale.setlocale(category=locale.LC_ALL, locale='')


class CalendarScraper:
    def __init__(self, year: int, backup_file: str, session: requests.Session) -> None:
        self.year: int = year
        self.calendar_url: str = f'https://online.zakon.kz/accountant/Calendars/Holidays/{self.year}'
        self.backup_file: str = backup_file
        if not os.path.exists(self.backup_file):
            self._save_backup()
        self.html: str = self.get_html(session=session)
        to_strain = SoupStrainer(name='div', attrs={'class': 'app-wrapper'})
        self.soup = BeautifulSoup(self.html, 'html.parser', parse_only=to_strain)
        self.date_infos: List[DateInfo] = self.get_date_infos()
        boundary_dates = self.get_boundary_dates(year=year, session=session)
        self.date_infos.insert(0, boundary_dates[0])
        self.date_infos.append(boundary_dates[1])

    def _save_backup(self, html: str = '') -> None:
        with open(file=self.backup_file, mode='w', encoding='utf-8') as html_file:
            html_file.write(html)

    def _load_backup(self) -> str:
        with open(file=self.backup_file, mode='r', encoding='utf-8') as html_file:
            return html_file.read()

    def get_html(self, session: requests.Session) -> str:
        try:
            response: requests.Response = session.get(url=self.calendar_url)
            response.raise_for_status()

            soup = BeautifulSoup(response.text, 'html.parser', parse_only=SoupStrainer(name='div', attrs={'class': 'app-wrapper'}))

            html: str = minify(str(soup), minify_js=True, minify_css=True)
            file_html: str = self._load_backup()
            if html == file_html:
                return html

            self._save_backup(html=html)
            return html
        except requests.exceptions.HTTPError:
            return self._load_backup()

    def get_date_infos(self) -> List[DateInfo]:
        date_infos: List[DateInfo] = []
        for element in self.soup.find_all('div', {'class': 'calendar-day'}):
            if 'masked' in element.attrs['class']:
                continue
            day = int(element.text)
            with LocaleManager(locale_name='ru_RU.UTF-8'):
                month = list(calendar.month_name).index(element.parent.parent.previous)
            is_holiday = 'holiday' in element.attrs['class']
            date_infos.append(DateInfo(date=date(year=self.year, month=month, day=day), is_work_day=not is_holiday))
        return date_infos

    @staticmethod
    def get_date_info(session: requests.Session, day: int, month: int, year: int) -> DateInfo or None:
        url = f'https://isdayoff.ru/{year}{month:02}{day:02}?cc=kz'
        try:
            response = session.get(url=url)
            response.raise_for_status()
            return DateInfo(date=date(year=year, month=month, day=day), is_work_day=not bool(response.json()))
        except requests.exceptions.HTTPError:
            return None

    @staticmethod
    def get_boundary_dates(year: int, session: requests.Session) -> DateInfo or None:
        return [
            CalendarScraper.get_date_info(session=session, day=31, month=12, year=year - 1),
            CalendarScraper.get_date_info(session=session, day=1, month=1, year=year + 1)
        ]

    @staticmethod
    def get_work_status(today: date, dates: List[DateInfo]) -> int:
        i: int = next(i for i, date_info in enumerate(dates) if date_info.date == today)
        is_today_work_day: bool = dates[i].is_work_day
        is_yesteday_work_day: bool = dates[i - 1].is_work_day
        is_tomorrow_work_day: bool = dates[i + 1].is_work_day

        if is_today_work_day and not is_tomorrow_work_day:
            return WorkStatus.HOLIDAY
        elif not is_today_work_day and is_yesteday_work_day:
            return WorkStatus.LONG
        else:
            return WorkStatus.WORK if is_today_work_day else WorkStatus.HOLIDAY
