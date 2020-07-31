import xlrd
import xlsxwriter
import regex as re
import requests
from bs4 import BeautifulSoup as bs

headers = {'accept': '*/*',
           'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_2) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36'}

skills = []
workbook = xlrd.open_workbook('SoftSkills_withoutOriginalText.xlsx')
worksheet = workbook.sheet_by_index(0)
row = 1
while True:
    try:
        key = str.lower(re.sub(r'[^\w\s]', '', worksheet.cell(row, 0).value))
        skills.append(key)
        row += 1
    except:
        break

skills_by_root = {}
session = requests.Session()
for key in skills:
    skills_by_root[key] = []
    if len(key.split()) == 1:
        request = session.get('https://wordroot.ru/' + key, headers=headers)
        if request.status_code == 200:
            soup = bs(request.content, 'html.parser')
            ols = soup.find_all('ol', attrs={'class': 'words-list'})
            for ol in ols:
                print(ol)
                lis = ol.find_all('li')
                for li in lis:
                    skills_by_root[key].append(li.text)

workbook = xlsxwriter.Workbook('SkillsByRoot.xlsx')
worksheet = workbook.add_worksheet()
x = 0
for skill in skills_by_root.items():
    words = skill[0]
    for word in skill[1]:
        words += ',' + word
    worksheet.write_string(x, 0, skill[0])
    worksheet.write_string(x, 1, words)
    x += 1
workbook.close()


