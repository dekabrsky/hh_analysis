import xlrd
import regex as re
import xlsxwriter


def get_soft_skills():
    skills = ['опыт работы в команде', 'аналитический склад ума', 'системное мышление','Желание развиваться',
              'умение искать и находить решения самостоятельно', 'обязательность',
             'внимательность к деталям', 'пунктуальность',
              'желание развивать свой продукт', 'умение разбивать задачи на этапы',
              'умение четко выполнять поставленные задачи']
    workbook = xlrd.open_workbook('SoftSkills_withoutOriginalText.xlsx')
    worksheet = workbook.sheet_by_index(0)
    row = 0
    while True:
        try:
            skills.append(worksheet.cell(row, 0).value)
            row += 1
        except:
            break
    return skills


def get_original_text():
    records = {}
    workbook = xlrd.open_workbook('xls/Vacancies.xlsx')
    worksheet = workbook.sheet_by_index(0)
    row = 1
    i = 0
    while True:
        try:
            records[worksheet.cell(row, 2).value] = worksheet.cell(row, 4).value
            row += 1
            i += 1
            if i > 590:
                break
        except:
            row += 1
            i += 1
            if i > 590:
                break
    return records


def make_result(skills, records):
    result = {}
    for record in records.items():
        str_record = str.lower(record[0] + ' ' + record[1])
        for skill in skills:
            if len(re.findall(skill, str_record)) != 0:
                if result.get(record[0]) is None:
                    result[record[0]] = skill
                else:
                    result[record[0]] += '\n' + skill
    return result


def make_xls(result):
    workbook = xlsxwriter.Workbook('xls/SoftSkillsDataset.xlsx')
    sheet = workbook.add_worksheet('Обработанное')
    sheet.set_column(0, 0, 100)  # A
    sheet.set_column(1, 1, 40)  # B
    sheet.write(0, 0, 'inputs')
    sheet.write(0, 1, 'outputs')
    cell_wrap = workbook.add_format()
    cell_wrap.set_text_wrap()
    row = 1
    col = 0
    for skill in result.items():
        sheet.write_string(row, col, skill[0], cell_wrap)
        sheet.write_string(row, col + 1, skill[1], cell_wrap)
        row += 1
    workbook.close()


def main():
    skills = get_soft_skills()
    print(skills)
    records = get_original_text()
    print(records)
    result = make_result(skills, records)
    print(result)
    make_xls(result)


if __name__ == '__main__':
    main()
