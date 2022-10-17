import gspread
from docxtpl import DocxTemplate
from docx2pdf import convert
import os

login = gspread.service_account(filename="Service\service_account.json")
sheet_name = login.open("HOA")
worksheet = sheet_name.worksheet("SEARCH_TOOL")
values = worksheet.get_values("B1:F13")

sunrise_id = values[1][0]
first_name = values[1][1]
second_name = values[1][2]
email = values[1][3]
phone = values[1][4]
street = values[3][0]
city = values[3][1]
state = values[3][2]
zip_code = values[3][3]
hoa_name = values[6][0]
hoa_email = values[6][1]
name = values[6][2]
quantity = values [9][0]
type = values[9][1]
pw = values [9][2]
address = values[12][0]
license_number = values[12][1]
date = values[12][2]
initials = values[12][4]

user = os.getlogin()

def acc():
    print(f"SUNRISE ID: {sunrise_id}")
    forms = ['NABR', 'Gray Hawk', 'Lexington', 'SRID']
    print(f"ATX Forms: \n \n{forms} \n")
    answer = input(f"Please choose the form you want: ")

    if answer == 'NABR':
        nabr()
    elif answer == 'Gray Hawk':
        gray_hawk()
    elif answer == 'Lexington':
        lexington()
    elif answer == 'SRID':
        changeid()
    else:
        choose_again()

def nabr():

    doc = DocxTemplate(r"Forms\\nabr.docx")
    context = {'date': date, 'name': name, 'hoa_name': hoa_name,
               'phone': phone, 'email': email, 'address': address}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"{name} NABR.docx")

    convert(f"{name} NABR.docx", f"{name} NABR.pdf")

    os.remove(f"{name} NABR.docx")
    print("ACC Finished...")

def gray_hawk():

    doc = DocxTemplate(r"Forms\\nabr.docx")
    context = {'date': date, 'name': name, 'hoa_name': hoa_name,
               'phone': phone, 'email': email, 'address': address}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"{name} NABR.docx")

    convert(f"{name} NABR.docx", f"{name} NABR.pdf")

    os.remove(f"{name} NABR.docx")
    print("ACC Finished...")

def lexington():
    doc = DocxTemplate(r"Forms\\lexington.docx")
    context = {'date': date, 'name': name, 'hoa_name': hoa_name,
            'phone': phone, 'email': email, 'address': address}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"{name} Lexington.docx")

    convert(f"{name} Lexington.docx", f"{name} Lexington.pdf")

    os.remove(f"{name} Lexington.docx")
    print("ACC Finished...")

def changeid():
    update_id = input(f"What is the Sunrise ID: ")
    worksheet.update_acell("H2", update_id)
    acc()

def choose_again():
    print(f" \n\n Invalid input, please copy and paste or enter exactly \n\n")
    acc()

def main():
    acc()

if __name__ == '__main__':
    main()