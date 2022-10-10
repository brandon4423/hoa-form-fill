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
    forms = ['ACMI', 'Chaparral', 'Cibolo', 'First Colony', 'First Service', 'HOA Management',
             'King Management', 'PAMCO', 'Prestige', 'SG 2', 'Stillwater', 'Woodlands']
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
    elif answer == 'User':
        print(os.getlogin())
    else:
        choose_again()

def acmi():
    os.getcwd()

    hoa_name = str(tab_lookup.acell("B7").value)
    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    phone = str(tab_lookup.acell("F2").value)
    email = str(tab_lookup.acell("E2").value)

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

    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    phone = str(tab_lookup.acell("F2").value)
    email = str(tab_lookup.acell("E2").value)
    quantity = str(tab_lookup.acell("B10").value)

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

    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    phone = str(tab_lookup.acell("F2").value)
    email = str(tab_lookup.acell("E2").value)

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

    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    city = str(tab_lookup.acell("C4").value)
    zip_code = str(tab_lookup.acell("E4").value)
    phone = str(tab_lookup.acell("F2").value)
    email = str(tab_lookup.acell("E2").value)

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

    hoa_name = str(tab_lookup.acell("B7").value)
    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    city = str(tab_lookup.acell("C4").value)
    state = str(tab_lookup.acell("D4").value)
    zip_code = str(tab_lookup.acell("E4").value)
    phone = str(tab_lookup.acell("F2").value)
    email = str(tab_lookup.acell("E2").value)

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

    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    phone = str(tab_lookup.acell("F2").value)
    email = str(tab_lookup.acell("E2").value)
    initial = str(tab_lookup.acell("F13").value)

    doc = DocxTemplate(r"Forms\\\hoa_management.docx")
    context = {'date': date, 'name': name,
               'phone': phone, 'email': email, 'address': address,
               'initial': initial}

    doc.render(context)
    os.chdir(r"C:\\\Users" + "\\" + user + "\\\Downloads")
    doc.save(f"{name} Hoa Management.docx")

    convert(f"{name} Hoa Management.docx", f"{name} Hoa Management.pdf")

    os.remove(f"{name} Hoa Management.docx")
    print("ACC Finished...")

def king_management():
    os.getcwd()

    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    phone = str(tab_lookup.acell("F2").value)
    email = str(tab_lookup.acell("E2").value)

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

def prestige():
    os.getcwd()

    hoa_name = str(tab_lookup.acell("B7").value)
    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    phone = str(tab_lookup.acell("F2").value)

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

    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    phone = str(tab_lookup.acell("F2").value)
    email = str(tab_lookup.acell("E2").value)

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

    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    phone = str(tab_lookup.acell("F2").value)
    email = str(tab_lookup.acell("E2").value)

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

    date = str(tab_lookup.acell("H7").value)
    name = str(tab_lookup.acell("D7").value)
    address = str(tab_lookup.acell("B13").value)
    phone = str(tab_lookup.acell("F2").value)
    email = str(tab_lookup.acell("E2").value)

    doc = DocxTemplate(r"Forms\\woodlands.docx")
    context = {'date': date, 'name': name,
               'phone': phone, 'email': email, 'address': address}

    doc.render(context)
    os.chdir(r"C:\\Users" + "\\" + user + "\\Downloads")
    doc.save(f"{name} Woodlands.docx")

    convert(f"{name} Woodlands.docx", f"{name} Woodlands.pdf")

    os.remove(f"{name} Woodlands.docx")
    print("ACC Finished...")

def choose_again():
    print(f" \n\n Invalid input, please copy and paste or enter exactly \n\n")
    acc()

def main():
    acc()

if __name__ == '__main__':
    main()