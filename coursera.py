import requests
import argparse
import openpyxl
from openpyxl import load_workbook
from bs4 import BeautifulSoup
import xlsxwriter


COUNT_OF_COURSE = 20


def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('save_file_as', help='A file name where will be putted all information ')
    arg = parser.parse_args()
    return arg.save_file_as


def get_course_list():
    urls_store = []
    inf_from_url = requests.get('https://www.coursera.org/sitemap~www~courses.xml')
    content_of_site = BeautifulSoup(inf_from_url.content, 'lxml')
    list_of_course = content_of_site.find_all('loc')[:COUNT_OF_COURSE]
    for course in list_of_course:
        course_url = requests.get(course.text)
        page_inf = BeautifulSoup(course_url.content, 'lxml')
        urls_store.extend([page_inf])
    return urls_store       


def get_course_name_lang_date():
    lang_store = []
    name_store = []
    date_store = []
    course_name_lng_date = get_course_list()
    for course in course_name_lng_date:
        nm = course.find_all('div',{'class':'title display-3-text'})
        dt = course.find_all('div',{'class':'startdate rc-StartDateString caption-text'})
        lng = course.find_all('div',{'class':'language-info'})
        name_store.extend([nm[0].text])
        date_store.extend([dt[0].text])
        lang_store.extend([lng[0].text])
    return name_store, date_store, lang_store,


def get_course_rating():
    rating_store = []
    course_rtg = get_course_list()
    for course in course_rtg:
        if course.find_all('div',{'class':'ratings-text bt3-visible-xs'}):
            rating = course.find_all('div',{'class':'ratings-text bt3-visible-xs'})
            rating_store.extend([rating[0].text])
        else:
            rating_store.extend(['This course does not have a rating'])
    return rating_store


def get_course_duration():
    duration_store = []
    course_dtn = get_course_list()
    for course in course_dtn:
        if course.find('i',{'cif-clock'}):
            implement = course.find('i',{'cif-clock'}).findNext('td')
            duration_store.extend([implement.text])
        else :
            duration_store.extend(["There is no information about duration"])
    return duration_store


def save_cource_inf_as_excell(file_name, course_name, course_date, course_lang, course_rating, course_duration):
    workbook = xlsxwriter.Workbook(file_name)
    cell_format = workbook.add_format({'bold': True, 'italic': True, 'fg_color': '#FFFF00' })
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Course name', cell_format)
    worksheet.write('B1', 'Course date', cell_format)
    worksheet.write('C1', 'Course language', cell_format)
    worksheet.write('D1', 'Course rating', cell_format)
    worksheet.write('E1', 'Course duration', cell_format)
    worksheet.write_column('A2', course_name)
    worksheet.write_column('B2', course_date)
    worksheet.write_column('C2', course_lang)
    worksheet.write_column('D2', course_rating)
    worksheet.write_column('E2', course_duration)
    workbook.close()


if __name__ == '__main__':
    file_name = get_args()
    course_name,course_date,course_lang = get_course_name_lang_date()
    course_rating = get_course_rating()
    course_duration = get_course_duration()
    save_cource_inf_as_excell(file_name, course_name, course_date, course_lang, course_rating, course_duration)
