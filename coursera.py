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
  
        
def get_course_information():
    urls_store = []
    course_name = []
    course_date = []
    course_lng = []
    course_rating = []
    course_duration = []
    course_urls = get_course_url()
    for course in course_urls:
        course_url = requests.get(course.text)
        page_inf = BeautifulSoup(course_url.content, 'lxml')
        course_name.append(get_course_name(page_inf))
        course_date.append(get_course_date(page_inf))
        course_lng.append(get_course_lng(page_inf))
        course_rating.append(get_course_rating(page_inf))
        course_duration.append(get_course_duration(page_inf))
        store={'name':course_name,
                'date':course_date,
                'lng':course_lng,
                'course_rating':course_rating,
                'course_duration':course_duration}
    return store


def get_course_name(course):
    nm = course.find('h1',{'class':'title display-3-text'}).text
    return nm


def get_course_date(course):
    dt = course.find_all('div',{'class':'startdate rc-StartDateString caption-text'})
    return dt[0].text


def get_course_lng(course):
    lng = course.find_all('div',{'class':'language-info'})
    return lng[0].text


def get_course_rating(course):
    if course.find_all('div',{'class':'ratings-text bt3-visible-xs'}):
        rating = course.find_all('div',{'class':'ratings-text bt3-visible-xs'})
        return rating[0].text
    else:
        return 'This course does not have a rating'


def get_course_duration(course):
    if course.find('i',{'cif-clock'}):
        implement = course.find('i',{'cif-clock'}).findNext('td')
        return implement.text
    else :
        return 'There is no information about duration'


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
    coureses = get_course_information()
    save_cource_inf_as_excell(file_name,coureses['name'], coureses['date'], coureses['lng'], coureses['course_rating'], coureses['course_duration'])

      

    