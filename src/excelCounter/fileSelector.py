from tkinter import Tk
from tkinter.filedialog import askopenfilename

def returnFile () -> str:
    Tk().withdraw() 
    filename = askopenfilename()
    if filename.endswith('.xlsx'):
        return filename, True
    elif filename.endswith('.csv'):
        return filename, False
    else:
        print ("Please select a .xlsx file")
        exit()