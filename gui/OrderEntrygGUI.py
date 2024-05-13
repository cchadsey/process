import tkinter as tk
from tkinter import filedialog as fd

#warnmessage = 'Placeholder'
#donejob = 'Jobs Done!'

def choose_file():
    file_path = fd.askopenfilename(title='Select a File')
    if file_path:
        order.config(text=f"Selected File: {file_path}")

def choose_folder():
    folder_path = fd.askdirectory(title = 'Select an Output Folder')
    if folder_path:
        folder.config(text=f'Selected Folder: {folder_path}')


m = tk.Tk()
m.title('VIP Supplier Order Entry')

folder = tk.Button(m, text =f'Select Folder', width= 25, command = choose_folder)
order = tk.Button(m, text=f'Select File', width = 25, command = choose_file)

colLabel = tk.Label(m, text = 'Case count Column')
col = tk.Entry(m)



ok = tk.Button(m, text='Begin', width=10, command = m.destroy)
cancel = tk.Button(m, text='Cancel', width=10, command = m.destroy)


#arrange main window in grid
folder.grid(row = 1, column = 0)
order.grid(row = 2, column = 0)
colLabel.grid(row=3, column = 0)
col.grid(row=3, column = 0)
ok.grid(pady = 50,row = 5, column = 0)
cancel.grid(pady = 50, row = 5, column = 1, padx = 20)


m.mainloop()
