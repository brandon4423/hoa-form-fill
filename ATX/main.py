import gspread
from docxtpl import DocxTemplate
from docx2pdf import convert
import os

login = gspread.service_account(filename="Service\service_account.json")
sheet_name = login.open("HOA")

tab_lookup = sheet_name.worksheet("Maya")

user = os.getlogin()

def acc():
    sunrise_id = str(tab_lookup.acell("H2").value)

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

    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    zip_code = str(tab_lookup.acell("E4").value)
    phone = str(tab_lookup.acell("F2").value)
    email = str(tab_lookup.acell("E2").value)

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

    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)

    email = str(tab_lookup.acell("E2").value)

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

    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    phone = str(tab_lookup.acell("F2").value)
    email = str(tab_lookup.acell("E2").value)

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

    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    phone = str(tab_lookup.acell("F2").value)
    email = str(tab_lookup.acell("E2").value)
    initial = str(tab_lookup.acell("F13").value)

    doc = DocxTemplate(r"Forms\\sun_city.docx")
    context = {'date': date, 'name': name,
               'phone': phone, 'email': email, 'address': address,
               'initial': initial}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"{name} Sun City.docx")

    convert(f"{name} Sun City.docx", f"{name} Sun City.pdf")

    os.remove(f"{name} Sun City.docx")
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