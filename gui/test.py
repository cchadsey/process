import tkinter as tk
import sys
import pyautogui

pyautogui.PAUSE = 1

def redirOutput(text_widget):
    class TextRedirector(object):
        def __init__(self, widget):
            self.widget = widget

        def write(self, string):
            self.widget.insert(tk.END, string)
            self.widget.see(tk.END)

    sys.stdout = TextRedirector(text_widget)

def perform_pyautogui_task():
    # Replace this with your actual PyAutoGUI code
    print("i'm being dumb")
    pyautogui.sleep(1)
    pyautogui.sleep(4)
    print("PyAutoGUI task completed")

window = tk.Tk()
window.title("PyAutoGUI Console Output")

text_widget = tk.Text(window, wrap=tk.WORD)
text_widget.pack(expand=True, fill="both")

redirOutput(text_widget)

button = tk.Button(window, text="Run PyAutoGUI Task", command=perform_pyautogui_task)
button.pack()

window.mainloop()