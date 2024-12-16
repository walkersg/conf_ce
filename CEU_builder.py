from pathlib import Path
from docxtpl import DocxTemplate
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import date
from docx2pdf import convert


### open file path

def select_file():
    global input_file
    input_file = filedialog.askopenfilename()
    if input_file:
        print("Selected file:", input_file)

def select_save_directory():
    global out_folder_path
    out_folder_path = filedialog.askdirectory()
    if out_folder_path:
        print("Selected directory:", out_folder_path)
    root.destroy()

root = tk.Tk()
root.title("File and Directory Selection")
root.geometry('600x300')

select_file_button = tk.Button(root, text="Select File", command=select_file)
select_file_button.pack(pady=20)

select_directory_button = tk.Button(root, text="Select Save Directory", command=select_save_directory)
select_directory_button.pack()

root.mainloop()

try:
    recipient_data = pd.read_csv(input_file, header=0)
    print('csv read')
except UnicodeDecodeError:
    recipient_data = pd.read_excel(input_file, header=0)

#cast all to lower
recipient_data.columns = recipient_data.columns.str.lower()


# right-strip each header
# replace any white space with an '_'
for i, v in enumerate(recipient_data.columns):
    recipient_data.columns.values[i] = str(recipient_data.columns.values[i]).rstrip()
    recipient_data.columns.values[i] = v.replace(' ','_')

#add columns to df
add_cols_list = ['ace_name','ace_number','date']
recipient_data['ace_name'] = 'Annie Chen'
recipient_data['provider_number'] = 'IP-24-11250'
recipient_data['event_date'] = pd.to_datetime(recipient_data['event_date'], errors='coerce')
recipient_data['event_date'] = recipient_data['event_date'].dt.strftime('%m-%d-%Y').astype(str)


#today's date
date_today = str(date.today().strftime('%m-%d-%Y'))
recipient_data['date'] = date_today
recipient_data = recipient_data.astype(str)

# print(recipient_data['event_date'])
for k in recipient_data.to_dict(orient='records'):
    doc = DocxTemplate('./ce_template.docx')
    doc.render(k)
    k['name'] = k['name'].replace(' ','_')
    output_path = './output/'+k['name']+'-'+k['event_date']+'.docx'
    doc.save(output_path)
#convert all to pdf
convert(out_folder_path)


    
    
