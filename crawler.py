import requests
import bs4
import os
import openpyxl


def get_vtu_page(sheet, url, usn, row_no):
    form_data = {'usn': usn}
    req = requests.post(url, data=form_data)
    soup = bs4.BeautifulSoup(req.text, "html.parser")
    tags = soup.select('td')
    n = len(tags)
    if tags[4].getText() != "15MAT41":
        return None
    try:
        sheet['A1'] = 'USN'
        sheet['B1'] = 'Name'
        # print(tags[4].getText())
        sheet['C1'] = tags[4].getText()
        sheet['D1'] = tags[10].getText()
        sheet['E1'] = tags[16].getText()
        sheet['F1'] = tags[22].getText()
        sheet['G1'] = tags[28].getText()
        sheet['H1'] = tags[34].getText()
        sheet['I1'] = tags[40].getText()
        sheet['J1'] = tags[46].getText()
    except IndexError:
        return None
    cell = str(row_no)
    usn = tags[1].getText()
    name = tags[3].getText()
    sheet['A'+cell] = usn[3:]
    sheet['B'+cell] = name[1:]
    col = 'C'
    for i in range(8, n-6, 6):
        sheet[col+cell] = tags[i].getText()
        col = get_next_cell(col)
        if col == 'K':
            break


def get_next_cell(current_cell):
    return chr(ord(current_cell)+1)


url = "http://results.vtu.ac.in/cbcs_17/result_page.php"
start_seq = "1BI15CS"
pwd = os.getcwd()
path = os.path.join(pwd, 'tests')
if not os.path.exists(path):
    os.makedirs(path)
wb = openpyxl.Workbook()
sheet = wb.get_active_sheet()
sheet.title = "5th Sem"
print("working", end='', flush=True)
for i in range(1, 200):
    print('.', end='', flush=True)
    if len(str(i)) == 1:
        roll = "00"+str(i)
    elif len(str(i)) == 2:
        roll = "0"+str(i)
    else:
        roll = str(i)
    usn = start_seq + roll
    get_vtu_page(sheet, url, usn, i+1)
print("\ndone.")
wb.save("output.xlsx")
