import tkinter as tk
from tkinter import filedialog as fd
from macrofunction import process_order
from PIL import Image, ImageTk
import os


file = ''
folder = ''
col= 'g'


def getCol():

    global col
    col = colentry.get()
    colbutton.config(text=f'Column {col} Set!')
    return


def choose_file():
    global file
    file_path = fd.askopenfilename(title='Select a File')
    if file_path:
        orderbutton.config(text=f"Click to choose another file")
        fileLabel = tk.Label(m, text = f'{os.path.normpath(os.path.basename(file_path))}')
        fileLabel.grid(row = 2, column = 0, columnspan= 1)
        orderbutton.grid(row = 2, column = 1, columnspan= 1)
        file = file_path
    return 


def choose_folder():
    global folder
    folder_path = fd.askdirectory(initialdir ='/Users/mcalesterpepsi/Desktop/Orders', title = 'Select an Output Folder')
    if folder_path:
        folderbutton.config(text=f'Folder Selected')
        folderlabel = tk.Label(m, text = f'./{os.path.normpath(os.path.basename(folder_path))}')
        folderlabel.grid(row = 1, column = 0, columnspan= 1)
        folderbutton.grid(row = 1, column = 1, columnspan= 1)
        folder = folder_path
    return 

def action_popup():


    popup = tk.Toplevel()

    popup.title('prepare')

    spacer = tk.Label(popup, text = '')
    spacer.pack()

    label = tk.Label(popup, text = 'Prepare for process. Get PO ready for entry.')
    label.pack()

    label2 = tk.Label(popup, padx = 15, text= 'Once you press READY you have 5s to select the entry box for the item code.')
    label2.pack()

    label3 = tk.Label(popup, text = 'Once the process begins, do not interact with computer until done.')
    label3.pack()

    label4 = tk.Label(popup, text= 'To kill application in an emergency, move mouse to top corner of monitor.')
    label4.pack()

    button = tk.Button(popup, text = 'READY', command = working_popup)
    button2 = tk.Button(popup, text = 'Return', command = popup.destroy)

    
    button.pack()
    button2.pack()

    popup.mainloop()


def working_popup():

    process_order(col,file,folder)
    
    
    fin = tk.Toplevel()
    fin.title('Done!')
    spacer = tk.Label(fin, text='')
    spacer.pack
    label = tk.Label(fin, text = 'Job done! choose another file, or cancel to quit!')
    label.pack()
    button = tk.Button(fin, text = 'Done', command = fin.destroy)
    button.pack()



m = tk.Tk()
m.title('VIP Supplier Order Entry')

folderbutton = tk.Button(m, text =f'Select Finished Folder', width= 25, command = choose_folder)
orderbutton = tk.Button(m, text=f'Select File', width = 25, command = choose_file)


colLabel = tk.Label(m, text = 'Case count Column')
colentry = tk.Entry(m, textvariable = '', width= 5)
colentry.insert(0, 'g')
colbutton = tk.Button(m, text= 'set', width= 10, command= getCol)


ok =  tk.Button(m, text='Begin', width = 10, command = action_popup)
cancel = tk.Button(m, text='Cancel', width=10, command = m.destroy)

img = Image.open('/Users/mcalesterpepsi/Desktop/Code/process/gui/pepsi.jpg')
rimg = img.resize(size = (150,150))
pythImg = ImageTk.PhotoImage(image = rimg)
image = tk.Label(m, image = pythImg, padx = 25, pady = 25)

#arrange main window in grid
folderbutton.grid(row = 1, column = 0, columnspan= 3)
orderbutton.grid(row = 2, column = 0, columnspan= 3)
colLabel.grid(row=3, column = 0)
colentry.grid(row=3, column = 1)
colbutton.grid(pady = 10, row = 3, column = 2)
ok.grid(pady = 20, row = 5, column = 1)
cancel.grid(row = 5, column = 2, padx = 20)
image.grid(row = 1, column = 3,rowspan = 3, columnspan = 3,pady = 15, padx = 15)


m.mainloop()
