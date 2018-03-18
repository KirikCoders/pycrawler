import openpyxl


def get_grade(item):

    if item < 40:
        return 0
    elif item < 45 and item >= 40:
        return 4
    elif item < 50 and item >= 45:
        return 5
    elif item < 60 and item >= 50:
        return 6
    elif item < 70 and item >= 60:
        return 7
    elif item < 80 and item >= 70:
        return 8
    elif item < 90 and  item >= 80:
        return 9
    elif item >= 90:
        return 10


def get_grade_letter(item):
    if item < 45:
        return 'F'
    elif item < 50 and item >= 45:
        return 'D'
    elif item < 60 and item >= 50:
        return 'C'
    elif item < 70 and item >= 60:
        return 'B'
    elif item < 80 and item >= 70:
        return 'A'
    elif item < 90 and  item >= 80:
        return 'S'
    elif item >= 90:
        return 'S+'


def calc_gpa(sheet):
        sheet["AA"+str(1)] = "SGPA"
        if sheet.title == "3rd sem" or sheet.title == "5th sem":
            for i in range(3, sheet.max_row+1):
                try:
                    sub1 = get_grade(int(sheet["E"+str(i)].value))
                    sub2 = get_grade(int(sheet["H"+str(i)].value))
                    sub3 = get_grade(int(sheet["K"+str(i)].value))
                    sub4 = get_grade(int(sheet["N"+str(i)].value))
                    sub5 = get_grade(int(sheet["Q"+str(i)].value))
                    sub6 = get_grade(int(sheet["T"+str(i)].value))
                    lab1 = get_grade(int(sheet["W"+str(i)].value))
                    lab2 = get_grade(int(sheet["Z"+str(i)].value))
                    total_core = sub1 + sub2 + sub3 + sub4 + sub5 + sub6
                    total_lab = lab1 + lab2
                    total = total_core*4 + total_lab * 2
                    sheet["C" + str(i)] = 4
                    sheet["F" + str(i)] = 4
                    sheet["I" + str(i)] = 4
                    sheet["L" + str(i)] = 4
                    sheet["O" + str(i)] = 4
                    sheet["R" + str(i)] = 4
                    sheet["U" + str(i)] = 2
                    sheet["X" + str(i)] = 2
                    sheet["D" + str(i)] = sub1
                    sheet["G" + str(i)] = sub2
                    sheet["J" + str(i)] = sub3
                    sheet["M" + str(i)] = sub4
                    sheet["P" + str(i)] = sub5
                    sheet["S" + str(i)] = sub6
                    sheet["V" + str(i)] = lab1
                    sheet["Y" + str(i)] = lab2
                    sheet["E"+str(i)] = get_grade_letter(int(sheet["E"+str(i)].value))
                    sheet["H"+str(i)] = get_grade_letter(int(sheet["H"+str(i)].value))
                    sheet["K"+str(i)] = get_grade_letter(int(sheet["K"+str(i)].value))
                    sheet["N"+str(i)] = get_grade_letter(int(sheet["N"+str(i)].value))
                    sheet["Q"+str(i)] = get_grade_letter(int(sheet["Q"+str(i)].value))
                    sheet["T"+str(i)] = get_grade_letter(int(sheet["T"+str(i)].value))
                    sheet["W"+str(i)] = get_grade_letter(int(sheet["W"+str(i)].value))
                    sheet["Z"+str(i)] = get_grade_letter(int(sheet["Z"+str(i)].value))
                    sheet["AA"+str(i)] = round(total / 28, 2)
                    # print("total"+str(i), total/28)
                    if total/28 > 8:
                        print(str(sheet["B"+str(i)].value), "Their score="+str(sheet["AA"+str(i)].value), sep=" ")
                except TypeError as e:
                    print("")
        # do nothing


wb = openpyxl.load_workbook(filename="output.xlsx")
sheet = wb.get_sheet_by_name("3rd sem")
calc_gpa(sheet)
sheet = wb.get_sheet_by_name("5th sem")
calc_gpa(sheet)
wb.save("new_out.xlsx")
