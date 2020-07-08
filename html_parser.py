import xlsxwriter
import requests
from bs4 import BeautifulSoup as bs

headers = {'accept': '*/*',
           'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'}
vacancy = 'Программист'
base_url = f'https://ekaterinburg.hh.ru/search/vacancy?area=1261&clusters=true&enable_snippets=true&search_field=name&specialization=1.221&text=%D0%9F%D1%80%D0%BE%D0%B3%D1%80%D0%B0%D0%BC%D0%BC%D0%B8%D1%81%D1%82&industry=7.540&from=cluster_subIndustry&showClusters=true'  # area=1 - Москва, search_period=3 - За 30 последних дня
pages = 2
jobs = []


def hh_parse(base_url, headers):
    zero = 0
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
    worksheet.set_column(1, 1, 135)  # B
    worksheet.set_column(2, 2, 20)  # C
    worksheet.set_column(3, 3, 40)  # D
    worksheet.set_column(4, 4, 135)  # E
    worksheet.set_column(5, 5, 45)  # F

    worksheet.write('A1', 'Наименование', bold)
    worksheet.write('B1', 'Полное описание', bold)
    worksheet.write('C1', 'Зарплата', bold)
    worksheet.write('D1', 'Компания', bold)
    worksheet.write('E1', 'Описание', bold)
    worksheet.write('F1', 'Ссылка', bold)

    while pages > zero:
        zero = str(zero)
        session = requests.Session()
        request = session.get(base_url + zero, headers=headers)
        if request.status_code == 200:
            soup = bs(request.content, 'html.parser')
            divs = soup.find_all('div', attrs={'data-qa': 'vacancy-serp__vacancy'})
            for div in divs:
                title = div.find('a', attrs={'data-qa': 'vacancy-serp__vacancy-title'}).text
                compensation = div.find('div', attrs={'data-qa': 'vacancy-serp__vacancy-compensation'})
                if compensation == None:  # Если зарплата не указана
                    compensation = 'None'
                else:
                    compensation = div.find('div', attrs={'data-qa': 'vacancy-serp__vacancy-compensation'}).text
                href = div.find('a', attrs={'data-qa': 'vacancy-serp__vacancy-title'})['href']
                try:
                    company = div.find('a', attrs={'data-qa': 'vacancy-serp__vacancy-employer'}).text
                except:
                    company = 'None'
                text1 = div.find('div', attrs={'data-qa': 'vacancy-serp__vacancy_snippet_responsibility'}).text
                text2 = div.find('div', attrs={'data-qa': 'vacancy-serp__vacancy_snippet_requirement'}).text
                content = text1 + '  ' + text2
                request2 = session.get(href, headers=headers)
                description = 'None'
                if request2.status_code == 200:
                    soup2 = bs(request2.content, 'html.parser')
                    description = soup2.find('div', attrs={'data-qa': 'vacancy-description'}).text
                all_txt = [title, description, compensation, company, content, href]
                jobs.append(all_txt)
            zero = int(zero)
            zero += 1

        else:
            print('error')

        # Запись в Excel файл

        row = 1
        col = 0
        for i in jobs:
            worksheet.write_string(row, col, i[0], center_V)
            worksheet.write_string(row, col + 1, i[1], cell_wrap)
            worksheet.write_string(row, col + 2, i[2], center_H_V)
            worksheet.write_string(row, col + 3, i[3], center_H_V)
            worksheet.write_string(row, col + 4, i[4], cell_wrap)
            # worksheet.write_url (row, col + 4, i[4], center_H_V)
            worksheet.write_url(row, col + 5, i[5])
            row += 1

        print('OK')
    workbook.close()


hh_parse(base_url, headers)
