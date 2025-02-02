import tkinter as tk
from tkinter import filedialog as fd
from PIL import Image, ImageTk
import os
import xlrd as xl
import pyautogui
import shutil as shu
from pynput.keyboard import Key, Controller
import datetime

def conopDay(daysfromnow):
    today = datetime.date.today()
    futuredate = today +datetime.timedelta(days= daysfromnow)
    dayofweek = futuredate.weekday()
    
    daystillmonday = 7 - dayofweek if dayofweek !=0 else 0

    if daystillmonday >= 4:
        answ = futuredate + datetime.timedelta(days=daystillmonday)
    else :
        answ = futuredate
    
    return answ.strftime('%m%d%y')


kbd = Controller()


file = ''
folder = ''
col= 'g'

def stroke(key):
    kbd.press(key)
    kbd.release(key)


#def copy():
   # with kbd.pressed(Key.cmd):
   #     kbd.press('a')
   #     kbd.release('a')
   # with kbd.pressed(Key.cmd):
   #     kbd.press('c')
   #     kbd.release('c')
   # with kbd.pressed(Key.cmd):
   #     kbd.press('v')
   #     kbd.release('v')



def process_order(case, file, folder, popup, supplier):

    popup.destroy()
    wp = tk.Toplevel(m, bd=75)
    wp.geometry(f"+{2*height}+{2//width}")

    l1 = tk.Label(wp, text= "you have 5s to select the entry field")
    l1.pack()

    boxlabel = tk.Label(wp, text='Status output')
    boxlabel.pack()

    box1 = tk.Text(wp, width=50, height=10, background='white', foreground="black")
    box1.pack()

    box1.insert(tk.END, f"Begining process")
    wp.update_idletasks()
    wp.update()





    
    alpha = "abcdefghijklmnopqrstuvwxyz"

    pyautogui.PAUSE = 0.5
    pyautogui.FAILSAFE = True


    wp.after(1, box1.insert(tk.END, f"\nGet Ready"))
    wp.update_idletasks()
    wp.update()
    
    i = 5 

    wp.after(1, box1.insert(tk.END, f"\nCountdown"))
    while i > 0:
        wp.after(1, box1.insert(tk.END, f"{i}."))
        wp.update_idletasks()
        wp.update()
        pyautogui.countdown(1)
        i -= 1

    wp.after(1, box1.insert(tk.END, f"\nProcessing"))
    wp.update_idletasks()
    wp.update()

    bk = xl.open_workbook(file)

    sh= bk.sheet_by_index(0)

    ccount = alpha.index(case.lower())

    ddate = alpha.index('j')

    #selecting open window
    pyautogui.click(1000,400)
    pyautogui.countdown(2)

    #making new PO CHANGE TO F6 AFTER TESTING
    pyautogui.press('f6')

    #copying po number to variable for later refrence.
    #copy()
    
    
    #po = pyperclip.paste()
    #print(po)

    #forward screen
    pyautogui.press('enter')
    pyautogui.countdown(2)


    for i in range(sh.nrows):
        if i!= 0 and sh.row_values(i)[ddate] != '':
            row = sh.row_values(i)

            rawdate = row[ddate]
            dstring = []
            dstring = rawdate.split('.')
            month = dstring[0]
            day = dstring[1]
            year = dstring[2]
            if len(month) <2:
                month= '0'+month
            if len(day) <2:
                day= '0'+day
            
            shipdate = month+day+year
            

    pyautogui.countdown(1)

    #set up po


    
    if supplier != '':
        shipdate = conopDay(10)
        stroke('a')
        stroke('a')
        pyautogui.press('tab')
        stroke('A')
        stroke('a')
        stroke('1')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.write(shipdate)
        pyautogui.press('tab')
        pyautogui.write(shipdate)
    else:
        stroke('b')
        stroke('a')
        pyautogui.press('tab')
        stroke('b')
        stroke('a')
        stroke('1')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.write(shipdate)
        pyautogui.press('tab')
        pyautogui.write(shipdate)
        
    pyautogui.press('enter')
    pyautogui.countdown(1)



    for i in range(sh.nrows):
        if i != 0:
            try:
                row = sh.row_values(i)

                c1 = str(row[0])
                c2 = int(row[ccount])

                if c1 != '90502.0' and c1 != '90503.0' and c1 != 'TOTALS':
                    c1 = c1.replace('.0', '')

                    while len(c1) <5:
                        c1 = '0'+ c1
                    
                    pyautogui.write(f'{c1[:5]}'); pyautogui.press('tab'); 

                    pyautogui.write(f'{c2}'); pyautogui.press('enter')

                    if i == 1:
                        
                        wp.after(1, box1.insert(tk.END, f"\nPausing to let the computer think"))
                        wp.update_idletasks()
                        wp.update()
                        pyautogui.sleep(5)
                        wp.after(1, box1.insert(tk.END, f"\nContinuing process"))
                        wp.update_idletasks()
                        wp.update()

                        

                    else:
                        pass
                elif c1 == 'TOTALS':
                    box1.insert(tk.END, f"\nEntry Complete. \nTotal Cases {c2}")
                    endbutton = tk.Button(wp, text="done", command = wp.destroy)
                    endbutton.pack()
                    wp.update_idletasks()
                    wp.update()
                    shu.move(file, folder)


                   

            except:
                pass


def getCol():

    global col
    col = colentry.get()
    colbutton.config(text=f'Column {col} Set!')
    return


def choose_file():
    global file
    fileLabel.configure(text = f'             ')
    file_path = fd.askopenfilename(title='Select a File')
    if file_path:
        orderbutton.config(text=f"Click to choose another file")
        fileLabel.config(text=f'{os.path.normpath(os.path.basename(file_path))}')
        file = file_path
    return 


def choose_folder():
    global folder
    folder_path = fd.askdirectory(initialdir ='/Users/mcalesterpepsi/Desktop/Orders/OrderBatches', title = 'Select an Output Folder')
    if folder_path:
        folderbutton.config(text=f'Folder Selected')
        folderlabel.config(text = f'./{os.path.normpath(os.path.basename(folder_path))}')
        folder = folder_path
    return 

def action_popup():


    popup = tk.Toplevel(m)
    popup.geometry(f"+{2*height}+{2//width}")

    popup.title('prepare')

    spacer = tk.Label(popup, text = '')
    spacer.pack()

    label = tk.Label(popup, text = 'Prepare for process. Get KEYINVEN open.')
    label.pack()

    #label2 = tk.Label(popup, padx = 15, text= 'Once you press BEGIN you have 5s to select the entry box for the item code.')
    #label2.pack()

    label3 = tk.Label(popup, text = 'Once the process begins, do not interact with computer until done.')
    label3.pack()

    label5 = tk.Label(popup, text='')
    label5.pack()

    label4 = tk.Label(popup, text= 'To kill application in an emergency, move mouse to top corner of monitor.')
    label4.pack()

    label6 = tk.Label(popup, text='')
    label6.pack()


    

    button = tk.Button(popup, text = 'BEGIN', command = lambda : process_order(col, file, folder, popup, supplier))
    button2 = tk.Button(popup, text = 'Close Popup', command = popup.destroy)

    
    button.pack()
    button2.pack()

def setvar(strng):
    
    global supplier
    supplier = strng
    if supplier == '':
        suplabel.config(text='Wispak Selected')
    elif supplier =='c':
        suplabel.config(text="Conops Selected")
    m.update_idletasks()
    m.update()


m = tk.Tk()
size = m.winfo_screenwidth()
width = size//2
height = m.winfo_screenheight()
m.geometry(f"+{2*height}+{2//width}")
m.winfo_toplevel()
m.title('VIP Supplier Order Entry')

folderbutton = tk.Button(m, text =f'Select Finished Folder', width= 25, command = choose_folder)
folderlabel = tk.Label(m, text = f'')
orderbutton = tk.Button(m, text=f'Select File', width = 25, command = choose_file)
fileLabel = tk.Label(m, text = f'')



colLabel = tk.Label(m, text = 'Case count Column')
colentry = tk.Entry(m, textvariable = '', width= 5)
colentry.insert(0, 'g')
colbutton = tk.Button(m, text= 'set', width= 10, command= getCol)

supplier = ''
suplabel = tk.Label(m, text= "Choose Supplier")
WpButton = tk.Button(m , text = "Wispak", width = 10, command = lambda: setvar(''))
conButton = tk.Button(m, text = 'Conops', width= 10, command = lambda : setvar('c'))


ok =  tk.Button(m, text='Ready', width = 10, command = action_popup)
cancel = tk.Button(m, text='Cancel', width=10, command = m.destroy)

img = Image.open('/Users/mcalesterpepsi/Desktop/Code/process/gui/pepsi.jpg')
rimg = img.resize(size = (150,150))
pythImg = ImageTk.PhotoImage(image = rimg)
image = tk.Label(m, image = pythImg, padx = 25, pady = 25)

#arrange main window in grid
folderbutton.grid(row = 1, column =1, columnspan= 2)
folderlabel.grid(row = 1, column=0, columnspan=1)
orderbutton.grid(row = 2, column =1, columnspan= 2)
fileLabel.grid(row=2, column=0 ,columnspan=  1)
colLabel.grid(row=3, column = 0)
colentry.grid(row=3, column = 1)
colbutton.grid(pady = 10, row = 3, column = 2)
suplabel.grid(row=4, column=0)
WpButton.grid(row=4, column=1)
conButton.grid(row=4, column = 2)
ok.grid(pady = 20, row = 5, column = 1)
cancel.grid(row = 5, column = 2, padx = 20)
image.grid(row = 1, column = 3,rowspan = 3, columnspan = 3,pady = 15, padx = 15)


m.mainloop()
