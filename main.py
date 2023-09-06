#/bin/usr/python3
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd

from pdf2docx import Converter

app=tk.Tk()
app.title('PDF Convert to Word')
app.geometry('600x350')

# Create a textfield for putting the
# text extracted from file
text = tk.Text(app, height=14)

# Specify the location of textfield
text.grid(column=0, row=0, sticky='nsew')

def open_text_file():
    # Specify the file types
    filetypes = (('text files', '*.pdf'),
                 ('All files', '*.*'))

    # Show the open file dialog by specifying path
    f = fd.askopenfile(filetypes=filetypes,
                       initialdir="D:/Downloads")

# Specifying the pdf & docx files
    pdf_file = f.name
    docx_file = f.name + '.docx'

    try:
        # Converting PDF to Docx
        cv_obj = Converter(pdf_file)
        cv_obj.convert(docx_file)
        cv_obj.close()

    except:
        print('Conversion Failed')
        text.insert('1.0', 'Conversion Failed')

    else:
        print('File Converted Successfully')
        text.insert('1.0', 'File Converted Successfully')

open_button = ttk.Button(app, text='Open a File',
                         command=open_text_file)

# Specify the button position on the app
open_button.grid(sticky='w', padx=250, pady=60)

# Make infinite loop for displaying app on the screen
app.mainloop()