import requests
import argparse
import openpyxl
from openpyxl import load_workbook
from bs4 import BeautifulSoup
import xlsxwriter


def get_args():
    parser = argparse.ArgumentParser()
    parser.add_argument('save_file_as', help='A file name where will be putted all information ')
    arg = parser.parse_args()
    return arg.save_file_as


def get_course_url():
    _course_count = 20
    inf_from_url = requests.get('https://www.coursera.org/sitemap~www~courses.xml')
    content_of_site = BeautifulSoup(inf_from_url.content, 'lxml')
    list_of_course = content_of_site.find_all('loc')[:_course_count]
    return list_of_course
  
        
def collect_course_information():
    course_information = []
    course_urls = get_course_url()
    for course in course_urls:
        course_url = requests.get(course.text)
        page_inf = BeautifulSoup(course_url.content, 'lxml')
        course_store = {
        'course_name':get_course_name(page_inf),
        'start_date': get_course_start_date(page_inf),
        'language': get_course_language(page_inf),
        'course_rating': get_course_rating(page_inf),
        'course_duration': get_course_duration(page_inf)
        }
        course_information.append(course_store)
    return course_information


def get_course_name(course):
    course_name = course.find('h1',{'class':'title display-3-text'}).text
    return course_name


def get_course_start_date(course):
    start_date = course.find_all('div',{'class':'startdate rc-StartDateString caption-text'})
    return start_date[0].text


def get_course_language(course):
    language = course.find_all('div',{'class':'language-info'})
    return language[0].text


def get_course_rating(course):
    rating = course.find('div',{'class':'ratings-text bt3-visible-xs'})
    if rating is None:
        return'This course does not have a rating'
    else:
        return rating.text


def get_course_duration(course):
    duration = course.find('i',{'cif-clock'})
    if duration is None:
        return 'There is no information about duration'
    else :
        return duration.findNext('td').text


def save_cource_inf_as_excell(file_name, course_name, start_date, course_lang, course_rating, course_duration):
    workbook = xlsxwriter.Workbook(file_name)
    cell_format = workbook.add_format({'bold': True, 'italic': True, 'fg_color': '#FFFF00' })
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', 'Course name', cell_format)
    worksheet.write('B1', 'Course start date', cell_format)
    worksheet.write('C1', 'Course language', cell_format)
    worksheet.write('D1', 'Course rating', cell_format)
    worksheet.write('E1', 'Course duration', cell_format)
    worksheet.write_column('A2', course_name)
    worksheet.write_column('B2', start_date)
    worksheet.write_column('C2', course_lang)
    worksheet.write_column('D2', course_rating)
    worksheet.write_column('E2', course_duration)
    workbook.close()


if __name__ == '__main__':
    file_name = get_args()
    courses = collect_course_information()
    save_cource_inf_as_excell(file_name, (course_name['course_name'] for course_name in courses),
    (start_date['start_date'] for start_date in courses),
    (language['language'] for language in courses),
    (course_rating['course_rating'] for course_rating in courses),
    (course_duration['course_duration'] for course_duration in courses))