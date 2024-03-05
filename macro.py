import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pyautogui
import xlrd as xl
import shutil as shu

#set up pyautogui failsafe and pause duration. tweak pause to give more time between input actions. 
pyautogui.PAUSE = 1.0
pyautogui.FAILSAFE = True

#Tk stuff. 
root = tk.Tk()
root.withdraw()

#Filechooser stuff.
rawfile = filedialog.askopenfilename()
#print (rawfile)

#display start message.
messagebox.showinfo(title = 'get ready', message= 'Doing the thing. Get your po ready, Click "OK", and bring up the page quick! (5s)')


#pause to give user time to select the first box for input.
pyautogui.countdown(5)

#get the workbook ready.
bk = xl.open_workbook(rawfile)

#get the spreadsheet from the workbook.
sh= bk.sheet_by_index(0)


#time to iterate through the spreadsheet
for i in range(sh.nrows):
#    print(i)

    #skip first row. has the column titles.
    if i != 0 :
        try:
            #set the list given from xlrd that contains the data from the row to "row" variable. 
            row = sh.row_values(i)
            #set the two cells we need to variables.
            c1 = str(row[0])
            c2 = int(row[6])
            #print (i, c1[:5], c1)
            
            #detecting the end of the needed data. 
            if c1 !=  '90502.0' and c1 != '90503.0' and c1 != 'TOTALS':
                #emulating the keypresses using the cell data acquired earlier.
                pyautogui.write(f'{c1[:5]}'); pyautogui.press('tab'); pyautogui.write(f'{c2}'); pyautogui.press('enter')
                #print ('hi', c1, c2)            
                
                #addressing some instability with the input page.
                if i == 1:
                    pyautogui.countdown(5)
            #needed for the if statement.        
            else:
                pass
        #needed for the try statement.    
        except:
            pass
#            print(False)

#move the file to "done" folder when completed.
shu.move(rawfile, '/Users/mcalesterpepsi/Desktop/Code/process/done')
#announce completion.
pyautogui.alert('Jobs Done!')        


