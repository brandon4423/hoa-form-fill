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
    forms = ['Black Hawk', 'PAMCO', 'Steiner', 'Sun City', 'SRID']
    print(f"ATX Forms: \n \n{forms} \n")
    answer = input(f"Please choose the form you want: ")

    if answer == 'Black Hawk':
        black_hawk()
    elif answer == 'PAMCO':
        pamco()
    elif answer == 'Steiner':
        steiner()
    elif answer == 'Sun City':
        sun_city()
    elif answer == 'SRID':
        changeid()
    else:
        choose_again()

def black_hawk():

    doc = DocxTemplate(r"Forms\\black_hawk.docx")
    context = {'date': date, 'name': name,
               'phone': phone, 'email': email, 'address': address,
               'zip_code': zip_code}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"{name} Black Hawk.docx")

    convert(f"{name} Black Hawk.docx", f"{name} Black Hawk.pdf")

    os.remove(f"{name} Black Hawk.docx")
    print("ACC Finished...")

def pamco():
    os.getcwd()

    doc = DocxTemplate(r"Forms\\pamco.docx")
    context = {'date': date, 'name': name,
               'email': email, 'address': address}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"{name} Pamco.docx")

    convert(f"{name} Pamco.docx", f"{name} Pamco.pdf")

    os.remove(f"{name} Pamco.docx")
    print("ACC Finished...")

def steiner():
    os.getcwd()

    doc = DocxTemplate(r"Forms\\steiner.docx")
    context = {'date': date, 'name': name,
               'phone': phone, 'email': email, 'address': address}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"Steiner.docx")

    convert(f"{name} Steiner.docx", f"{name} Steiner.pdf")
    
    os.remove(f"{name} Steiner.docx")
    print("ACC Finished...")

def sun_city():
    os.getcwd()

    doc = DocxTemplate(r"Forms\\sun_city.docx")
    context = {'date': date, 'name': name,
               'phone': phone, 'email': email, 'address': address,
               'initial': initials}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"{name} Sun City.docx")

    convert(f"{name} Sun City.docx", f"{name} Sun City.pdf")

    os.remove(f"{name} Sun City.docx")
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