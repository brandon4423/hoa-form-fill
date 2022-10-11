import gspread
from docxtpl import DocxTemplate
from docx2pdf import convert
import os

login = gspread.service_account(filename="Service\service_account.json")
sheet_name = login.open("HOA")

tab_lookup = sheet_name.worksheet("SEARCH_TOOL")

user = os.getlogin()

def acc():
    sunrise_id = str(tab_lookup.acell("H2").value)

    print(f"SUNRISE ID: {sunrise_id}")
    forms = ['NABR', 'SRID']
    print(f"ATX Forms: \n \n{forms} \n")
    answer = input(f"Please choose the form you want: ")

    if answer == 'NABR':
        nabr()
    elif answer == 'SRID':
        changeid()
    else:
        choose_again()

def nabr():

    hoa_name = str(tab_lookup.acell("B7").value)
    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    phone = str(tab_lookup.acell("F2").value)
    email = str(tab_lookup.acell("E2").value)

    doc = DocxTemplate(r"Forms\\nabr.docx")
    context = {'date': date, 'name': name, 'hoa_name': hoa_name,
               'phone': phone, 'email': email, 'address': address}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"{name} NABR.docx")

    convert(f"{name} NABR.docx", f"{name} NABR.pdf")

    os.remove(f"{name} NABR.docx")
    print("ACC Finished...")

def changeid():
    update_id = input(f"What is the Sunrise ID: ")
    tab_lookup.update_acell("H2", update_id)
    acc()

def choose_again():
    print(f" \n\n Invalid input, please copy and paste or enter exactly \n\n")
    acc()

def main():
    acc()

if __name__ == '__main__':
    main()