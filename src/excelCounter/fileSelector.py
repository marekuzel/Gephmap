from tkinter import Tk
from tkinter.filedialog import askopenfilename

def returnFile ():
    Tk().withdraw() 
    filename = askopenfilename()
    if filename.endswith('.xlsx'):
        return filename
    else:
        print ("Please select a .xlsx file")
        exit()