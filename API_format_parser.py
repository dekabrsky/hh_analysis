import requests as res
import xlsxwriter
import html2text
import regex as re


def get_vacancies():
    p = dict(text="Программист", area="1261", per_page="100", page="0")
    pages = list()
    count_pages = 9

    for i in range(count_pages):
        p['page'] = i
        get = res.get("https://api.hh.ru/vacancies", params=p)
        pages.append(get.json())
    full_vacancies = list()

    for page in pages:
        for item in page['items']:
            # print(item['employer']['id'])
            full_vacancies.append(res.get("https://api.hh.ru/vacancies/" + item['id']).json())

    return full_vacancies


def make_xls(full_vacancies, skills):
    workbook = xlsxwriter.Workbook('xls/Vacancies.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})
    bold.set_align('center')
    center_H_V = workbook.add_format()
    center_H_V.set_align('center')
    center_H_V.set_align('vcenter')
    center_V = workbook.add_format()
    center_V.set_align('vcenter')
    cell_wrap = workbook.add_format()
    cell_wrap.set_text_wrap()

    worksheet.set_column(0, 0, 35)  # A  https://xlsxwriter.readthedocs.io/worksheet.html#set_column
    worksheet.set_column(1, 1, 20)  # B
    worksheet.set_column(2, 2, 135)  # C
    worksheet.set_column(3, 3, 40)  # D
    worksheet.set_column(4, 4, 50)  # E
    worksheet.set_column(5, 5, 40)

    worksheet.write('A1', 'Наименование', bold)
    worksheet.write('B1', 'Компания', bold)
    worksheet.write('C1', 'Описание', bold)
    worksheet.write('D1', 'Ключевые навыки', bold)
    worksheet.write('E1', 'Автодополненные навыки', bold)
    worksheet.write('F1', 'Ссылка', bold)

    row = 1
    col = 0
    for vacancy in full_vacancies:
        description = html2text.html2text(vacancy['description'])
        worksheet.write_string(row, col, vacancy['name'], center_V)
        worksheet.write_string(row, col + 1, vacancy['employer']['name'], center_H_V)
        worksheet.write_string(row, col + 2, description, cell_wrap)

        skills_str = ""
        for i in range(len(vacancy['key_skills'])):
            skills_str += vacancy['key_skills'][i]['name'] + ', '
        worksheet.write_string(row, col + 3, skills_str, center_H_V)

        description = re.sub(r'[^\w\s]', ' ', description)
        for word in description.split(' '):
            if str.lower(word) in skills.keys():
                skills_str += word + ', '
        worksheet.write_string(row, col + 4, skills_str, center_H_V)

        worksheet.write_string(row, col + 5, vacancy['alternate_url'], cell_wrap)
        row += 1
    workbook.close()
    print('OK')


def get_skills_dict(full_vacancies):
    skills = dict()
    for vacancy in full_vacancies:
        for i in range(len(vacancy['key_skills'])):
            skill = str.lower(vacancy['key_skills'][i]['name'])
            if skill not in skills.keys():
                skills[skill] = 1
            else:
                skills[skill] += 1
    skills = {k: v for k, v in reversed(sorted(skills.items(), key=lambda item: item[1]))}
    print(type(skills))
    return skills


def fill_comps(skills):
    workbook = xlsxwriter.Workbook('xls/Skills.xlsx')
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': 1})
    bold.set_align('center')

    worksheet.set_column(0, 0, 35)
    worksheet.set_column(1, 1, 20)

    worksheet.write('A1', 'Навык', bold)
    worksheet.write('B1', 'Количество', bold)

    row = 1
    col = 0
    for skill in skills.items():
        worksheet.write_string(row, col,  skill[0])
        worksheet.write_string(row, col + 1, str(skill[1]))
        row += 1
    workbook.close()
    print('OK')


def main():
    full_vacancies = get_vacancies()
    skills = get_skills_dict(full_vacancies)
    make_xls(full_vacancies, skills)
    fill_comps(skills)


if __name__ == "__main__":
    main()
