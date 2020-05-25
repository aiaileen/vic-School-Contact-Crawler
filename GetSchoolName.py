import urllib
import requests
from pyquery import PyQuery
import xlrd
import xlwt


def get_school_name(xl_file):
    worksheet = xlrd.open_workbook(xl_file).sheet_by_index(0)
    rows_count = worksheet.nrows
    schools_name_list = []
    for i in range(0, rows_count):
        schools_name_list.append(worksheet.cell(i, 0).value)
    return schools_name_list


def search(sch_names):
    for school in sch_names:
        query = school.replace(' ', '+').replace('\'', '%27')
        url = f"https://www.google.com/search?q={query}&aqs=chrome.0.0l2.250j0j9&sourceid=chrome&ie=UTF-8"
        respond = requests.get(url)
        if respond.status_code == 200:
            doc = PyQuery('<html></html>')


xl_file = 'schoolName.xlsx'
school_names = get_school_name(xl_file)