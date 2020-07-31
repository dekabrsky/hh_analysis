import xlrd
import regex as re
import xlsxwriter


def get_soft_skills():
    skills = {}
    workbook = xlrd.open_workbook('xls/SoftSkillsDataset.xlsx')
    worksheet = workbook.sheet_by_index(0)
    row = 1
    while True:
        try:
            skills[worksheet.cell(row, 0).value] = worksheet.cell(row, 1).value
            row += 1
        except:
            break
    return skills


def make_xls(result):
    workbook = xlsxwriter.Workbook('xls/SoftSkillsDatasetSmall.xlsx')
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
    result = {}
    for skill in skills.items():
        if len(skill[1].split('\n')) > 2:
            result[skill[0]] = skill[1]
    make_xls(result)


if __name__ == '__main__':
    main()
