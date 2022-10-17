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
    forms = ['ACMI', 'Chaparral', 'Cibolo', 'First Colony', 'First Service', 'HOA Management',
             'King Management', 'PAMCO', 'Prestige', 'SG 2', 'Stillwater', 'Woodlands', 'SRID']
    print(f"SATX & HOU Forms: \n \n{forms} \n")
    answer = input(f"Please choose the form you want: ")

    if answer == 'ACMI':
        acmi()
    elif answer == 'Chaparral':
        chaparral()
    elif answer == 'Cibolo':
        cibolo()
    elif answer == 'First Colony':
        first_colony()
    elif answer == 'First Service':
        fsr()
    elif answer == 'HOA Management':
        hoa_management()
    elif answer == 'King Management':
        king_management()
    elif answer == 'PAMCO':
        pamco()
    elif answer == 'Prestige':
        prestige()
    elif answer == 'SG 2':
        sg()
    elif answer == 'Stillwater':
        stillwater()
    elif answer == 'Woodlands':
        woodlands()
    elif answer == 'SRID':
        changeid()
    elif answer == 'User':
        print(os.getlogin())
    else:
        choose_again()

def acmi():
    os.getcwd()

    doc = DocxTemplate(r"Forms\\acmi.docx")
    context = {'hoa_name': hoa_name, 'date': date, 'name': name,
               'phone': phone, 'email': email, 'address': address}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"{name} Acmi.docx")

    convert(f"{name} Acmi.docx", f"{name} Acmi.pdf")

    os.remove(f"{name} Acmi.docx")
    print("ACC Finished...")


def chaparral():
    os.getcwd()

    doc = DocxTemplate(r"Forms\\\chaparral.docx")
    context = {'date': date, 'name': name, 'phone': phone,
               'email': email, 'address': address, 'quantity': quantity}

    doc.render(context)
    os.chdir(r"C:\\\Users" + "\\" + user + "\\\Downloads")
    doc.save(f"{name} Chaparral.docx")

    convert(f"{name} Chaparral.docx", f"{name} Chaparral.pdf")

    os.remove(f"{name} Chaparral.docx")
    print("ACC Finished...")

def cibolo():
    os.getcwd()

    doc = DocxTemplate(r"Forms\\\cibolo.docx")
    context = {'date': date, 'name': name,
               'phone': phone, 'email': email, 'address': address}

    doc.render(context)
    os.chdir(r"C:\\\Users" + "\\" + user + "\\\Downloads")
    doc.save(f"{name} Cibolo.docx")

    convert(f"{name} Cibolo.docx", f"{name} Cibolo.pdf")

    os.remove(f"{name} Cibolo.docx")
    print("ACC Finished...")

def first_colony():
    os.getcwd()

    doc = DocxTemplate(r"Forms\\\first_colony.docx")
    context = {'date': date, 'name': name, 'city': city,
               'phone': phone, 'email': email, 'address': address,
               'zip_code': zip_code}

    doc.render(context)
    os.chdir(r"C:\\\Users" + "\\" + user + "\\\Downloads")
    doc.save(f"{name} First Colony.docx")

    convert(f"{name} First Colony.docx", f"{name} First Colony.pdf")

    os.remove(f"{name} First Colony.docx")
    print("ACC Finished...")

def fsr():
    os.getcwd()

    doc = DocxTemplate(r"Forms\\\firstservice_acc.docx")
    context = {'hoa_name': hoa_name, 'date': date, 'name': name,
               'phone': phone, 'email': email, 'address': address,
               'city': city, 'state': state, 'zip_code': zip_code}

    doc.render(context)
    os.chdir(r"C:\\\Users" + "\\" + user + "\\\Downloads")
    doc.save(f"{name} FSR.docx")

    convert(f"{name} FSR.docx", f"{name} FSR.pdf")

    os.remove(f"{name} FSR.docx")
    print("ACC Finished...")

def hoa_management():
    os.getcwd()

    doc = DocxTemplate(r"Forms\\\hoa_management.docx")
    context = {'date': date, 'name': name,
               'phone': phone, 'email': email, 'address': address,
               'initial': initials}

    doc.render(context)
    os.chdir(r"C:\\\Users" + "\\" + user + "\\\Downloads")
    doc.save(f"{name} Hoa Management.docx")

    convert(f"{name} Hoa Management.docx", f"{name} Hoa Management.pdf")

    os.remove(f"{name} Hoa Management.docx")
    print("ACC Finished...")

def king_management():
    os.getcwd()

    doc = DocxTemplate(r"Forms\\\king_management.docx")
    context = {'date': date, 'name': name,
               'phone': phone, 'email': email, 'address': address}

    doc.render(context)
    os.chdir(r"C:\\\Users" + "\\" + user + "\\\Downloads")
    doc.save(f"{name} King Management.docx")

    convert(f"{name} King Management.docx", f"{name} King Management.pdf")

    os.remove(f"{name} King Management.docx")
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

def prestige():
    os.getcwd()

    doc = DocxTemplate(r"Forms\\prestige.docx")
    context = {'hoa_name': hoa_name, 'date': date, 'name': name,
               'phone': phone, 'address': address}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"{name} Prestige.docx")

    convert(f"{name} Prestige.docx", f"{name} Prestige.pdf")

    os.remove(f"{name} Prestige.docx")
    print("ACC Finished...")

def sg():
    os.getcwd()

    doc = DocxTemplate(r"Forms\\sg.docx")
    context = {'name': name,
               'phone': phone, 'email': email, 'address': address}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"{name} SG-2.docx")

    convert(f"{name} SG-2.docx", f"{name} SG-2.pdf")

    os.remove(f"{name} SG-2.docx")
    print("ACC Finished...")

def stillwater():
    os.getcwd()

    doc = DocxTemplate(r"Forms\\stillwater.docx")
    context = {'date': date, 'name': name,
               'phone': phone, 'email': email, 'address': address}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"{name} Stillwater.docx")

    convert(f"{name} Stillwater.docx", f"{name} Stillwater.pdf")

    os.remove(f"{name} Stillwater.docx")
    print("ACC Finished...") 

def woodlands():
    os.getcwd()

    doc = DocxTemplate(r"Forms\\woodlands.docx")
    context = {'date': date, 'name': name,
               'phone': phone, 'email': email, 'address': address}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"{name} Woodlands.docx")

    convert(f"{name} Woodlands.docx", f"{name} Woodlands.pdf")

    os.remove(f"{name} Woodlands.docx")
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