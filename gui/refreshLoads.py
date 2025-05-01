import pyautogui as pag

pag.PAUSE = 0.25

def top():
    pag.moveTo(416,234)
    pag.click()

def refresh():
    pag.moveTo(1745,270)
    pag.click()

def next(x,y):
    
    pag.moveTo(x,y)
    pag.click()
    refresh()

top()
pag.click()
next(530,233)
next(640,233)
next(750,233)
next(850,233)
next(950,233)
next(1050,233)
next(1160,233)
next(1250,233)

