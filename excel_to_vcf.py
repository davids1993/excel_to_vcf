#get excel file location
from os import mkdir, rmdir


def get_excel_file():
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename()


#get the first row of an excel spreadsheet as a dictionary key with column letter as value
def get_headers(sheet):
    headers = {}
    for column in range(1, sheet.max_column + 1):
        headers[sheet.cell(row=1, column=column).value] = sheet.cell(row=1, column=column).column_letter
    return headers


#import excel contact list and map columns to headers
def import_contacts(file_location):
    import openpyxl
    wb = openpyxl.load_workbook(file_location)
    sheet = wb['Sheet1']
    headers = get_headers(sheet)
    contacts = []
    for row in range(2, sheet.max_row + 1):
        contact = {}
        contact['name'] = sheet[headers['name'] + str(row)].value
        contact['email'] = sheet[headers['email'] + str(row)].value
        contact['phone'] = sheet[headers['phone'] + str(row)].value
        contact['title'] = sheet[headers['title'] + str(row)].value
        contacts.append(contact)
    return contacts

#convert contacts to .vcf
def contacts_to_vcf(contacts):
    import vobject
    vcfs = []
    for contact in contacts:
        vcard = vobject.vCard()
        vcard.add('n')
        vcard.n.value = vobject.vcard.Name(family=contact['name'])
        vcard.add('fn')
        vcard.fn.value = contact['name']
        vcard.add('email')
        vcard.email.value = contact['email']
        vcard.email.type_param = 'INTERNET'
        vcard.add('tel')
        vcard.tel.value = contact['phone']
        vcard.tel.type_param = 'CELL'
        vcard.add('title')
        vcard.title.value = contact['title']
        vcfs.append(vcard)
    return vcfs

#create new folder named 'contacts'
def create_folder(file_location):
    import os
    import shutil
    file_location = get_path(file_location)
    os.chdir(file_location)
    try:
        os.mkdir('contacts')
        os.mkdir('merged_contacts')
    except FileExistsError:
        #delete old folder and create new one
        shutil.rmtree('contacts')
        shutil.rmtree('merged_contacts')
        os.mkdir('contacts')
        os.mkdir('merged_contacts')


#go to file location and save .vcf files to all contacts folder
def save_vcfs(file_location, vcfs):
    import os
    file_location = get_path(file_location)
    os.chdir(file_location)
    os.chdir('contacts')
    for vcf in vcfs:
        with open(f'{vcf.fn.value}.vcf', 'w') as f:
            f.write(vcf.serialize())
            
#get path without the file name
def get_path(file_location):
    import os
    return os.path.dirname(file_location)

#merge all text files into one
def merge_files(file_location):
    file_location = get_path(file_location)
    import os
    import glob
    os.chdir(file_location)
    os.chdir('contacts')
    files = glob.glob('*.vcf')
    save_path = os.path.join(file_location, 'merged_contacts')
    with open(os.path.join(save_path, 'contacts.vcf'), 'w') as outfile:
        for fname in files:
            with open(fname, 'r') as infile:
                outfile.write(infile.read())

print("\n\nThe excel file should have the following columns in the order below:")
print("Name, Email, Phone, Title")
input("\nPress enter to continue.")



file_location = get_excel_file()
create_folder(file_location)
contacts = import_contacts(file_location)
vcfs = contacts_to_vcf(contacts)
save_vcfs(file_location,vcfs)
merge_files(file_location)

