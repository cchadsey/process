import tkinter as tk
from tkinter import filedialog as fd
from tkinter import messagebox as mb
from PIL import Image, ImageTk
import os
import xlrd as xl
import pyautogui
import shutil as shu
from pynput.keyboard import Key, Controller
import pyperclip
import datetime
import easyocr
import numpy as np

reader = easyocr.Reader(['en'])

kbd = Controller()

def stroke(key):
    kbd.press(key)
    kbd.release(key)


class OrderSheet:


    def __init__(self, file, case):
        
        pass