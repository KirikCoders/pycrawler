import requests
import bs4
import os
import openpyxl


def get_vtu_page(sheet, url, usn, row_no, filter_sub):
    form_data = {'usn': usn}
    req = requests.post(url, data=form_data)
    soup = bs4.BeautifulSoup(req.text, "html.parser")
    tags = soup.select('td')
    n = len(tags)
    if tags[4].getText() != filter_sub:
        return None
    try:
        # print(tags[4].getText())
        sheet['C1'] = tags[4].getText()
        sheet['F1'] = tags[10].getText()
        sheet['I1'] = tags[16].getText()
        sheet['L1'] = tags[22].getText()
        sheet['O1'] = tags[28].getText()
        sheet['R1'] = tags[34].getText()
        sheet['U1'] = tags[40].getText()
        sheet['X1'] = tags[46].getText()
    except IndexError:
        return None
    cell = str(row_no)
    usn = tags[1].getText()
    name = tags[3].getText()
    sheet['A'+cell] = usn[3:]
    sheet['B'+cell] = name[1:]
    col = 'C'
    for i in range(6, n, 6):
        for j in range(3):
            sheet[col+cell] = tags[i+j].getText()
            col = get_next_cell(col)
        if col == "AA":
            break
        # sheet[col+cell] = tags[i+1].getText()


def get_next_cell(current_cell):
    if current_cell == 'Z':
        return "AA"
    return chr(ord(current_cell)+1)


def crawl(url, start_seq, sheet, title, filter_sub):
    sheet.title = title
    sheet.merge_cells('C1:E1')
    sheet.merge_cells('F1:H1')
    sheet.merge_cells('I1:K1')
    sheet.merge_cells('L1:N1')
    sheet.merge_cells('O1:Q1')
    sheet.merge_cells('R1:T1')
    sheet.merge_cells('U1:W1')
    sheet.merge_cells('X1:Z1')
    sheet['A1'] = 'USN'
    sheet['B1'] = 'Name'
    init_cols = 'B'
    cell_value = 2
    for i in range(0, 8):
        init_cols = get_next_cell(init_cols)
        cell = init_cols + str(cell_value)
        sheet[cell] = "IA"
        init_cols = get_next_cell(init_cols)
        cell = init_cols + str(cell_value)
        sheet[cell] = "EA"
        init_cols = get_next_cell(init_cols)
        cell = init_cols + str(cell_value)
        sheet[cell] = "Total"
    for i in range(1, 500):
        print('.', end='', flush=True)
        if len(str(i)) == 1:
            roll = "00" + str(i)
        elif len(str(i)) == 2:
            roll = "0" + str(i)
        else:
            roll = str(i)
        usn = start_seq + roll
        get_vtu_page(sheet, url, usn, i + 2, filter_sub)


wb = openpyxl.Workbook()
sheet1 = wb.get_active_sheet()
print("working", end='', flush=True)
crawl("http://results.vtu.ac.in/cbcs_17/result_page.php", "1BI16CS", sheet1, "2nd sem", "15MAT21")
print("\n")
sheet1 = wb.create_sheet(index=1)
crawl("http://results.vtu.ac.in/cbcs_17/result_page.php", "1BI15CS", sheet1, "4th sem", "15MAT41")
sheet1 = wb.create_sheet(index=2)
crawl("http://results.vtu.ac.in/results17/index.php", "1BI14CS", sheet1, "6th sem", "10AL61")
print("\ndone.")
wb.save("output.xlsx")
