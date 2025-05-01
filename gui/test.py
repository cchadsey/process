import pyautogui
import easyocr
import numpy as np

reader = easyocr.Reader(['en'])

casecountimg = pyautogui.screenshot(region = (830,687,80,50))
frame = np.array(casecountimg)
text = reader.readtext(frame)
subtext = text[0]
print(subtext[1])
for detection in text:
    print(detection[1])
    print(type(detection[1]))
    if detection[1] == '2060':
        print ('Fuck')
    else:
        print ('fuckyea')
