import xlrd
import xlsxwriter

hard_skills = {}
workbook = xlrd.open_workbook('xls/input2.xlsx')
worksheet = workbook.sheet_by_index(0)
row = 1
while True:
    try:
        listik = worksheet.cell(row, 3).value
        listik = listik.replace(';', ',')
        listik = listik.replace(', ', ',')
        print(listik)
        for skill in listik.split(','):
            skill = str.lower(skill)
            print(skill)
            if skill != '':
                if hard_skills.get(skill) is None:
                    hard_skills[skill] = 1
                else:
                    hard_skills[skill] += 1
        row += 1
    except:
        break

hard_skills = {k: v for k, v in reversed(sorted(hard_skills.items(), key=lambda item: item[1]))}

"""soft_skills = {}
workbook = xlrd.open_workbook('Vacancies_3.xlsx')
worksheet = workbook.sheet_by_index(0)
row = 1
while True:
    try:
        for skill in worksheet.cell(row, 4).value.split('\n'):
            skill = str.lower(skill)
            print(skill)
            if skill != '':
                if soft_skills.get(skill) is None:
                    soft_skills[skill] = 1
                else:
                    soft_skills[skill] += 1
        row += 1
    except:
        break

soft_skills = {k: v for k, v in reversed(sorted(soft_skills.items(), key=lambda item: item[1]))}"""

workbook = xlsxwriter.Workbook('xls/skills_by_count_2.xlsx')
worksheet = workbook.add_worksheet()
bold = workbook.add_format({'bold': 1})
bold.set_align('center')

worksheet.set_column(0, 0, 20)
worksheet.set_column(1, 1, 10)

worksheet.write('A1', 'Hard Skill', bold)
worksheet.write('B1', 'Count', bold)
"""worksheet.write('D1', 'Soft Skill', bold)
worksheet.write('E1', 'Count', bold)"""

row = 1
col = 0
for skill in hard_skills.items():
    worksheet.write_string(row, col,  skill[0])
    worksheet.write_string(row, col + 1, str(skill[1]))
    row += 1

"""row = 1
col = 3
for skill in soft_skills.items():
    print(skill)
    worksheet.write_string(row, col,  skill[0])
    worksheet.write_string(row, col + 1, str(skill[1]))
    row += 1"""
workbook.close()
print('OK')
