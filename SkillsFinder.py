import xlrd
import xlsxwriter
import regex as re
import pprint


class Record:
    def __init__(self, n, d, hh, link, sbr, hs_list):
        self.name = n
        self.description = d
        self.hh_skills = hh
        self.link = link
        self.skills_by_root = sbr
        self.hs_list = hs_list
        self.hs = self.get_hs()
        self.ss = self.get_ss()

    def get_hs(self):
        soft_skills = sbr.keys()
        hard_skills = self.hs_list
        result = []
        for skill in self.hh_skills.split(', '):
            if str.lower(skill) not in soft_skills:
                result.append(skill)
        for word in re.sub(r'[^\w\s]', ' ', self.description).split():
            if str.lower(word) in hard_skills:
                result.append(word)
        result = list(set(result))
        return result

    def get_ss(self):
        soft_skills = sbr.keys()
        hard_skills = self.hs_list
        result = []
        for skill in self.hh_skills.split(', '):
            if str.lower(skill) in soft_skills and str.lower(skill) not in hard_skills:
                result.append(skill)
        for word in re.sub(r'[^\w\s]', ' ', self.description).split():
            for sskills in sbr.items():
                if str.lower(word) in sskills[1]:
                     result.append(sskills[0])
        for skill in soft_skills:
            if len(re.findall(skill, str.lower(self.description))) != 0:
                result.append(skill)
                print(skill)
        result = list(set(result))
        return result


sbr = {}
workbook = xlrd.open_workbook('xls/SkillsByRoot.xlsx')
worksheet = workbook.sheet_by_index(0)
row = 0
while True:
    try:
        key = worksheet.cell(row, 0).value
        values = worksheet.cell(row, 1).value.split(',')
        sbr[key] = values
        row += 1
    except:
        break

hs_list = []
workbook = xlrd.open_workbook('xls/Skills_Обр.xlsx')
worksheet = workbook.sheet_by_index(0)
row = 0
while True:
    try:
        hs = worksheet.cell(row, 0).value
        hs_list.append(hs)
        row += 1
    except:
        break

records = []
workbook = xlrd.open_workbook('xls/Vacancies.xlsx')
worksheet = workbook.sheet_by_index(0)
row = 1
while True:
    try:
        name = worksheet.cell(row, 0).value
        descr = worksheet.cell(row, 2).value
        skills = worksheet.cell(row, 3).value
        link = worksheet.cell(row, 5).value
        record = Record(name, descr, skills, link, sbr, hs_list)
        records.append(record)
        pprint.pprint(record)
        row += 1
    except:
        break

workbook = xlsxwriter.Workbook('xls/Vacancies_3.xlsx')
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
worksheet.set_column(1, 1, 135)  # B
worksheet.set_column(2, 2, 40)  # C
worksheet.set_column(3, 3, 40)  # D
worksheet.set_column(4, 4, 40)  # E
worksheet.set_column(5, 5, 40)

worksheet.write('A1', 'Наименование', bold)
worksheet.write('B1', 'Описание', bold)
worksheet.write('C1', 'Исконные ключевые навыки', bold)
worksheet.write('D1', 'Hard Skills', bold)
worksheet.write('E1', 'Soft Skills', bold)
worksheet.write('F1', 'Ссылка', bold)

row = 1
col = 0
for record in records:
    worksheet.write_string(row, col, record.name, center_V)
    worksheet.write_string(row, col + 1, record.description, cell_wrap)
    worksheet.write_string(row, col + 2, record.hh_skills,  cell_wrap)

    skills_str = ""
    for skill in record.hs:
        skills_str += '\n' + skill
    worksheet.write_string(row, col + 3, skills_str, cell_wrap)

    skills_str = ""
    for skill in record.ss:
        skills_str += '\n' + skill
    worksheet.write_string(row, col + 4, skills_str, cell_wrap)

    worksheet.write_string(row, col + 5, record.link, cell_wrap)
    row += 1
workbook.close()
print('OK')
