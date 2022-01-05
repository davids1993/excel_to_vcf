from excel_to_vcf import *

#show user a message box telling them to select a file
def select_file_message():
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo('Select File', 'Please select an excel file to convert to vcf')


def message_contacts_converted():
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo('Success', 'Your contacts have been converted!')
     
                
#give user a link to the merged contacts folder
def show_link(file_location):
    import webbrowser
    file_location = get_path(file_location)
    webbrowser.open(f'file://{file_location}')
    
    
    
if '__main__' == __name__:
    select_file_message()
    file_location = get_excel_file()
    create_folder(file_location)
    contacts = import_contacts(file_location)
    vcfs = contacts_to_vcf(contacts)
    save_vcfs(file_location,vcfs)
    merge_files(file_location)
    message_contacts_converted()
    show_link(file_location)