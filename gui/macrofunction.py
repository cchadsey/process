import xlrd as xl
import pyautogui
import shutil as shu

def process_order(case, file, folder):

    
    alpha = "abcdefghijklmnopqrstuvwxyz"

    pyautogui.PAUSE = 0.5
    pyautogui.FAILSAFE = True

    pyautogui.countdown(5)

    bk = xl.open_workbook(file)

    sh= bk.sheet_by_index(0)

    ccount = alpha.index(case.lower())

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
                        pyautogui.countdown(5)

                    else:
                        pass

                   

            except:
                pass
            
    shu.move(file, folder) 

    