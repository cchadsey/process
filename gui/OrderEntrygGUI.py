import tkinter as tk
from tkinter.filedialog import askopenfile
from tkinter.filedialog import askdirectory

warnmessage = 'Placeholder'
donejob = 'Jobs Done!'



m = tk.Tk()



title = tk.Label(m, text = 'VIP Supplier Order Entry')

folder = tk.Button(m, text='Folder', width= 25, command= askdirectory())
file = tk.Button(m, text='File', width = 25, command = askopenfile())

tk.Label(m, text = 'Case count Column').grid(row = 3)
col = tk.Entry(m)



ok = tk.Button(m, text='Begin', width=10, command = m.destroy)
cancel = tk.Button(m, text='Cancel', width=10, command = m.destroy)


#arrange main window in grid
title.grid(row= 0 , column = 0, columnspan = 2)
folder.grid(row = 1, column = 0)
file.grid(row = 2, column = 0)
col.grid(row=3, column = 0)
ok.grid(pady = 50,row = 5, column = 0)
cancel.grid(pady = 50, row = 5, column = 1, padx = 20)


m.mainloop()
