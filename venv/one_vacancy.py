import xlsxwriter
import requests
from bs4 import BeautifulSoup as bs

headers = {'accept': '*/*',
           'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'}
vacancy = 'Программист'
base_url = 'https://ekaterinburg.hh.ru/vacancy/37018880'
jobs = []


def hh_parse(base_url, headers):
    session = requests.Session()
    request = session.get(base_url, headers=headers)
    if request.status_code == 200:
        soup = bs(request.content, 'html.parser')
        description = soup.find('div', attrs={'data-qa': 'vacancy-description'}).text
        print(description)
        #all_txt = [title, compensation, company, content, href]
        all_txt = [description]
        jobs.append(all_txt)

    # Запись в Excel файл
    workbook = xlsxwriter.Workbook('Vacancy_2.xlsx')
    worksheet = workbook.add_worksheet()
    # Добавим стили форматирования
    bold = workbook.add_format({'bold': 1})
    bold.set_align('center')
    center_H_V = workbook.add_format()
    center_H_V.set_align('center')
    center_H_V.set_align('vcenter')
    center_V = workbook.add_format()
    center_V.set_align('vcenter')
    cell_wrap = workbook.add_format()
    cell_wrap.set_text_wrap()

    # Настройка ширины колонок
    worksheet.set_column(0, 0, 35)  # A  https://xlsxwriter.readthedocs.io/worksheet.html#set_column
    worksheet.set_column(1, 1, 20)  # B
    worksheet.set_column(2, 2, 40)  # C
    worksheet.set_column(3, 3, 135)  # D
    worksheet.set_column(4, 4, 45)  # E

    worksheet.write('A1', 'Наименование', bold)
    worksheet.write('B1', 'Зарплата', bold)
    worksheet.write('C1', 'Компания', bold)
    worksheet.write('D1', 'Описание', bold)
    worksheet.write('E1', 'Ссылка', bold)

    row = 1
    col = 0
    for i in jobs:
        worksheet.write_string(row, col, i[0], center_V)
        #worksheet.write_string(row, col + 1, i[1], center_H_V)
       # worksheet.write_string(row, col + 2, i[2], center_H_V)
        #worksheet.write_string(row, col + 3, i[3], cell_wrap)
        # worksheet.write_url (row, col + 4, i[4], center_H_V)
       # worksheet.write_url(row, col + 4, i[4])
        row += 1

        print('OK')
    workbook.close()


hh_parse(base_url, headers)
