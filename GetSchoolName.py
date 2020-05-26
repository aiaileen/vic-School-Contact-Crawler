import urllib
import requests
from requests import RequestException
from pyquery import PyQuery
import xlrd
import xlwt
import re


def get_school_name(xl_file):
    worksheet = xlrd.open_workbook(xl_file).sheet_by_index(0)
    rows_count = worksheet.nrows
    schools_name_list = []
    for i in range(0, rows_count):
        schools_name_list.append(worksheet.cell(i, 0).value)
    return schools_name_list


def get_web_page(url):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            return response.text
        print("Request Failed:", url)
        return None
    except RequestException:
        print("Request Failed:", url)
        return None


def get_school_url(school_name):
    query = school_name.replace(' ', '+').replace('\'', '%27')
    google_url = f"https://www.google.com/search?q={query}&aqs=chrome.0.0l2.250j0j9&sourceid=chrome&ie=UTF-8"
    html = get_web_page(google_url)
    if html is not None:
        doc = PyQuery(html)
        search_results = doc('body').find('.VGHMXd')  # find <div> css class=VGHMXd
        for result in search_results:
            href = result.attrib['href']  # check whether href match with vic.edu.au
            check = re.search(r'(http).+(.vic.edu.au)', href)
            if check:
                sch_url = check.group()
                return sch_url
        return None
    else:
        return


def search_contact_us(keyword, page):
    doc = PyQuery(page)  # html file

    doc('main').filter(doc('h').text()== mat)# re.match keyword

# def write_excel(detail, file_address):
#     wbk = xlwt.Workbook()
#     sheet = wbk.add_sheet(cell_overwrite_ok=True)
#
#     wbk.save(file_address)


# input_xl_file = 'schoolName.xlsx'
input_xl_file = 'Book1.xlsx'
output_xl_file = 'result.xlsx'
school_list = get_school_name(input_xl_file)
num = 0
no_url_school_ist = []  # 最后存入另一个xl
for school in school_list:
    num += 1
    school_url = get_school_url(school)
    if school_url is None:
        no_url_school_ist.append(school)
    else:
        contact_url = school_url + '/contact'
        contact_page = get_web_page(contact_url)
        if contact_page is not None:   # 有contact Us 页面
            suburb = re.findall(r'[A-Z]{2,} ?[A-Z]{2,} ?[A-Z]{2,}', school)  # 提取大写单词作为搜索key
            search_contact_us(suburb[0], contact_page)
            # print(num, ':', contact_url)

            # doc = respond.text.


        # else:  # 在主页搜索

    # get_school_detail(school_url)



