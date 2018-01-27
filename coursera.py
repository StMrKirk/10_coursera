import requests
import bs4 as bs
from openpyxl import Workbook
import sys


def get_courses_list():
    courses_list = []
    response = requests.get('https://www.coursera.org/sitemap~www~courses.xml')
    soup = bs.BeautifulSoup(response.text, "lxml")
    locs = soup.find_all('loc')
    for loc in locs[:2]:
        link = loc.text
        courses_list.append(link)
    return courses_list


def get_course_info(course_slug):
    courses_info = []
    for course in course_slug:
        response = requests.get(course)
        soup = bs.BeautifulSoup(response.text, "lxml")
        if soup.find('div', class_="ratings-text bt3-hidden-xs") is None:
            stars = 'Not Given Yet'
        else:
            stars = soup.find('div', class_="ratings-text bt3-hidden-xs").text
        course_info = (soup.find('h1', class_="title display-3-text").text,
                       soup.find('p', class_="body-1-text course-description").text,
                       soup.find('div', class_="rc-Language").text,
                       soup.find('div', class_="startdate rc-StartDateString caption-text").text,
                       len(soup.find_all('div', class_="week-heading body-2-text")),
                       stars)
        courses_info.append(course_info)
    return courses_info


def output_courses_info_to_xlsx(filepath, courses_info):
    wb = Workbook()
    ws = wb.active
    for row in ws.iter_rows(min_row=1, max_col=6, max_row=len(courses_info)):
        for cell in row:
            cell.value = courses_info[cell.row-1][cell.col_idx-1]
    wb.save('{}.xlsx'.format(filepath))


if __name__ == '__main__':
    courses_list = get_courses_list()
    courses_info = get_course_info(courses_list)
    output_courses_info_to_xlsx('{}'.format(sys.argv[1]), courses_info)
    
