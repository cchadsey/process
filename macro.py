import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pyautogui
import xlrd as xl
import shutil as shu

#for handling what cell the case count is in
alpha =  "abcdeghijklmnopqrstuvwxyz"

#set up pyautogui failsafe and pause duration. tweak pause to give more time between input actions. 
pyautogui.PAUSE = 0.5
pyautogui.FAILSAFE = True

#Tk stuff. 
root = tk.Tk()
root.withdraw()

#Filechooser stuff.
donedir = filedialog.askdirectory()
rawfile = filedialog.askopenfilename()
#print (rawfile)

#need to get what cells to deal with
case = pyautogui.prompt(text = 'What column letter is the Case count in?', default = 'g')
ccount = alpha.index(case.lower()) + 1
#print(ccount)

#display start message.
messagebox.showinfo(title = 'Get Ready', message= 'Doing the thing. Get your po ready, Click "OK", and bring up the page quick! (5s) To stop in an emergency, put mouse in top corner of screen.')


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
            c2 = int(row[ccount])
            #print (i, c1[:5], c1)
            
            #detecting the end of the needed data. 
            if c1 !=  '90502.0' and c1 != '90503.0' and c1 != 'TOTALS':
                #removing the .0 at the end without converting to an int so that the leading 0s wont be removed on proper sheets.
                c1 = c1.replace('.0', '')
                
                #adding leading 0s to to short codes as excel removes them in some instances.
                while len(c1) <5:
                    c1 = '0'+ c1

                #emulating the keypresses using the cell data acquired earlier.
                pyautogui.write(f'{c1[:5]}'); pyautogui.press('tab'); 

                pyautogui.write(f'{c2}'); pyautogui.press('enter')
                
                #fixing an issue with loading on the first entry            
                if i == 1:
                    pyautogui.countdown(5)

                else:
                    pass

                
                
                #print ('hi', c1, c2)
                
                #addressing some instability with the input page.
                
            #needed for the if statement.        
            else:
                pass
        #needed for the try statement.    
        except:
            pass
#            print(False)

#move the file to "done" folder when completed.
shu.move(rawfile, donedir)
#announce completion.
pyautogui.alert('Jobs Done!')        


