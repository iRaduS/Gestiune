def return_date_language(month):
    french = {"ianuarie": "JANVIER", "februarie": "FÉVRIER", "martie": "MARS", "aprilie": "AVRIL", "mai": "MAI",
              "iunie": "JUIN", "iulie": "JUILLET", "august": "AOÛT", "septembrie": "SEPTEMBRE", "octombrie": "OCTOBRE",
              "noiembrie": "NOVEMBRE", "decembrie": "DÉCEMBRE"}
    return french[month] if month in french else None


def format_worksheet(employee, worksheet, data):
    try:
        worksheet.title = f"{employee[1]} {employee[2]}"
        worksheet['B9'] = f"{return_date_language(data[0])} {data[1]}"
        worksheet['B16'] = employee[1]
        worksheet['B17'] = employee[2]
        worksheet['C23'] = employee[3]
        worksheet['C24'] = employee[6]
        worksheet['C26'] = employee[7]
        worksheet['C28'] = employee[8]
        worksheet['C30'] = float(employee[14]) / float(data[2])
        money_total = (float(employee[14]) - float(employee[12]) - float(employee[11]) - float(employee[10])) / float(data[2]) + float(employee[25])
        money_per_hour = money_total / (float(worksheet['C24'].value) + 1.25 * float(worksheet['C26'].value) + 1.5 * float(worksheet['C28'].value))
        worksheet['C25'] = money_per_hour
        worksheet['C27'] = 1.25 * money_per_hour
        worksheet['C29'] = 1.5 * money_per_hour
        worksheet['C32'] = float(employee[15]) / float(data[2])
        worksheet['C33'] = float(employee[17]) / float(data[2])
        worksheet['C35'] = float(employee[20]) / float(data[2])
        worksheet['C36'] = (float(employee[23]) + float(employee[22]) + float(employee[21])) / float(data[2])
        worksheet['C37'] = employee[25]
        worksheet['C38'] = float(worksheet['C37'].value) + float(worksheet['C36'].value)
    except TypeError:
       print(f"Angajatul {employee[1]} {employee[2]} are eroare la tipuri, verifica campurile lui din fisierul {data[0]} {data[1]} franta.xlsx!")
