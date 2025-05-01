import pyautogui as pag

pag.PAUSE = 0.25

def top():
    pag.moveTo(416,234)
    pag.click()


def next(x,y):
    
    pag.moveTo(x,y)
    pag.click()
    top()

top()
pag.click()
next(390,420)
next(390,445)
next(390,470)
next(390,495)
next(390,520)
next(390,545)
next(390,570)
next(390,595)

