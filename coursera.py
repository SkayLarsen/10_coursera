import argparse
import json
import requests
import random
from bs4 import BeautifulSoup
from lxml import etree
from openpyxl import Workbook


def get_course_title(soup):
    return soup.find("div", "course-title").text


def get_course_lang(soup):
    return soup.find("div", "language-info").text.split(',')[0]


def get_course_start(soup):
    try:
        json_data = soup.select('script[type="application/ld+json"]')[0].text
        return json.loads(json_data)['hasCourseInstance'][0]['startDate']
    except (KeyError, IndexError):
        return None


def get_course_duration(soup):
    duration = len(soup.find_all("div", "week"))
    return duration if duration > 0 else None


def get_course_rate(soup):
    try:
        return soup.find('div', "ratings-text").text
    except AttributeError:
        return None


def get_course_info(course_url):
    course_page = requests.get(course_url).content.decode("utf-8", "ignore")
    soup = BeautifulSoup(course_page, "html.parser")
    return [get_course_title(soup), get_course_lang(soup), get_course_start(soup),
            get_course_duration(soup), get_course_rate(soup)]


def get_courses_list(number_of_courses=20):
    url = 'https://www.coursera.org/sitemap~www~courses.xml'
    courses_xml = requests.get(url).content
    tree = etree.fromstring(courses_xml)
    random_curses = random.sample(list(tree), number_of_courses)
    return [course[0].text for course in random_curses]


def make_courses_workbook(courses_info):
    book = Workbook()
    sheet = book.active
    sheet.title = "Курсы"
    sheet.append(['Название', 'Язык', 'Дата начала', 'Количество недель', 'Средняя оценка'])
    for course in courses_info:
        sheet.append([info if info is not None
                      else "Нет данных"
                      for info in course])
    return book


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Сохранение информации о курсах с Coursera в xlsx-файл")
    parser.add_argument('filepath', help='путь к xlsx-файлу')
    args = parser.parse_args()
    courses = [get_course_info(url) for url in get_courses_list()]
    courses_workbook = make_courses_workbook(courses)
    courses_workbook.save(args.filepath)
