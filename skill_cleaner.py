import xlrd
import regex as re
import xlsxwriter


def read_xls():
    skills = {}
    workbook = xlrd.open_workbook('xls/Skills.xlsx')
    worksheet = workbook.sheet_by_index(0)
    row = 1
    while True:
        try:
            skills[worksheet.cell(row, 0).value] = int(worksheet.cell(row, 1).value)
            row += 1
        except:
            break
    return skills


def separate_kirillic(skills):
    k_skills = {}
    for skill in skills.items():
        if len(re.findall(r'[а-яА-ЯёЁ]', skill[0])) != 0:
            k_skills[skill[0]] = skills[skill[0]]
    for skill in k_skills.items():
        skills.pop(skill[0])
    return k_skills


def compress_skills(skills):
    c_skills = {}
    for skill in skills.items():
        if len(re.findall('framework', skill[0])) != 0 and skills.get(skill[0][:-10]) is not None:
            c_skills[skill[0][:-10]] = skills[skill[0][:-10]] + skills[skill[0]]
    for skill in c_skills.items():
        skills[skill[0]] = skill[1]
        skills.pop(skill[0])
        skills.pop(skill[0] + ' framework')
    skills = {k: v for k, v in reversed(sorted(skills.items(), key=lambda item: item[1]))}
    return skills


def compress_1c(k_skills):
    skills_1c = {}
    for skill in k_skills.items():
        if (len(re.findall('1с', skill[0])) != 0 or len(re.findall('1c', skill[0])) != 0) and skill[0] != '1с':
            skills_1c[skill[0]] = skill[1]
    k_skills['1с'] = 0
    for skill in skills_1c.items():
        k_skills['1с'] += skill[1]
        k_skills.pop(skill[0])
    k_skills = {k: v for k, v in reversed(sorted(k_skills.items(), key=lambda item: item[1]))}
    return k_skills


def make_xls(skills, k_skills):
    workbook = xlsxwriter.Workbook('xls/Skills_Обр.xlsx')
    sheet = workbook.add_worksheet('Обработанное')
    sheet.set_column(0, 0, 40)  # A
    sheet.set_column(1, 1, 20)  # B
    sheet.set_column(2, 2, 50)  # C
    sheet.set_column(3, 3, 40)  # D
    sheet.set_column(4, 4, 20)  # E
    sheet.write(0, 0, 'Навык ENG')
    sheet.write(0, 1, 'Кол-во')
    sheet.write(0, 3, 'Навык RUS')
    sheet.write(0, 4, 'Кол-во')
    row = 1
    col = 0
    for skill in skills.items():
        sheet.write(row, col, skill[0])
        sheet.write(row, col + 1, skill[1])
        row += 1
    row = 1
    col = 3
    for skill in k_skills.items():
        sheet.write(row, col, skill[0])
        sheet.write(row, col + 1, skill[1])
        row += 1
    workbook.close()


def main():
    skills = read_xls()
    print('Изначально:', skills)
    k_skills = separate_kirillic(skills)
    print('На английском:', skills)
    print('На русском: ', k_skills)
    skills = compress_skills(skills)
    print('Объединил технология и технология+фреймворк: ', skills)
    k_skills = compress_1c(k_skills)
    print('Сжал 1С:', k_skills)
    make_xls(skills, k_skills)


if __name__ == '__main__':
    main()
