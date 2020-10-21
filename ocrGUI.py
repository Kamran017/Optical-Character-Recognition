import sys
from tkinter import filedialog#browse a file or folder
import tabula
import pandas as pd
import os
import glob
import csv
from xlsxwriter.workbook import Workbook



# Directories
HOME_PATH = os.getcwd()
HOME_PATH = HOME_PATH.replace("\\", "/")
HOME_PATH = HOME_PATH[2:]
HOME_PATH = HOME_PATH + "/"
print(HOME_PATH)



try:
    import Tkinter as tk
except ImportError:
    import tkinter as tk

try:
    import ttk
    py3 = False
except ImportError:
    import tkinter.ttk as ttk
    py3 = True

import ocrGUI_support

def vp_start_gui():
    '''Starting point when module is the main routine.'''
    global val, w, root
    root = tk.Tk()
    top = Toplevel1 (root)
    ocrGUI_support.init(root, top)
    root.mainloop()

w = None
def create_Toplevel1(rt, *args, **kwargs):
    '''Starting point when module is imported by another module.
       Correct form of call: 'create_Toplevel1(root, *args, **kwargs)' .'''
    global w, w_win, root
    #rt = root
    root = rt
    w = tk.Toplevel (root)
    top = Toplevel1 (w)
    ocrGUI_support.init(w, top, *args, **kwargs)
    return (w, top)

def destroy_Toplevel1():
    global w
    w.destroy()
    w = None

class Toplevel1:
    def __init__(self, top=None):
        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''
        _bgcolor = '#d9d9d9'  # X11 color: 'gray85'
        _fgcolor = '#000000'  # X11 color: 'black'
        _compcolor = '#d9d9d9' # X11 color: 'gray85'
        _ana1color = '#d9d9d9' # X11 color: 'gray85'
        _ana2color = '#ececec' # Closest X11 color: 'gray92'
        font9 = "-family {Segoe UI} -size 18"

        top.geometry("600x450+650+150")
        top.minsize(120, 1)
        top.maxsize(1924, 1061)
        top.resizable(0, 0)
        top.title("Convert PDF file to Excel")
        top.configure(background="#d9d9d9")

        self.Label1 = tk.Label(top)
        self.Label1.place(relx=0.35, rely=0.022, height=41, width=184)
        self.Label1.configure(background="#d9d9d9")
        self.Label1.configure(disabledforeground="#a3a3a3")
        self.Label1.configure(font=font9)
        self.Label1.configure(foreground="#000000")
        self.Label1.configure(text='''Select a PDF file''')

        self.Button1 = tk.Button(top, command=self.browseFunc)
        self.Button1.place(relx=0.35, rely=0.133, height=24, width=67)
        self.Button1.configure(activebackground="#ececec")
        self.Button1.configure(activeforeground="#000000")
        self.Button1.configure(background="#d9d9d9")
        self.Button1.configure(disabledforeground="#a3a3a3")
        self.Button1.configure(foreground="#000000")
        self.Button1.configure(highlightbackground="#d9d9d9")
        self.Button1.configure(highlightcolor="black")
        self.Button1.configure(pady="0")
        self.Button1.configure(text='''Browse''')

        self.Button2 = tk.Button(top, command=self.convert)
        self.Button2.place(relx=0.533, rely=0.133, height=24, width=67)
        self.Button2.configure(activebackground="#ececec")
        self.Button2.configure(activeforeground="#000000")
        self.Button2.configure(background="#d9d9d9")
        self.Button2.configure(disabledforeground="#a3a3a3")
        self.Button2.configure(foreground="#000000")
        self.Button2.configure(highlightbackground="#d9d9d9")
        self.Button2.configure(highlightcolor="black")
        self.Button2.configure(pady="0")
        self.Button2.configure(text='''Save File''')
        
    #browse a pdf file
    def browseFunc(self):
        #tk.messagebox.showinfo('Information','Please select model with PDF extension!')
        root.filename = filedialog.askopenfilename(initialdir="/", title="Select A File", filetype= (("all files","*.*"), ("pdf files","*.pdf")))
        global dataset
        dataset= root.filename
        
    global dataset
    def convert(self):
        #tk.messagebox.showinfo('Information','Please wait for next information box !')
        global dataset
        pdf = tabula.read_pdf(dataset, 
                     pages='all')
        print(HOME_PATH)
        tabula.convert_into(dataset, r"result.csv" , 
                    output_format="csv",pages='all', stream=True)
        
        #convert csv file to excel            
        for csvfile in glob.glob(os.path.join('.', '*.csv')):
            workbook = Workbook(csvfile[:-4] + '.xlsx')
            worksheet = workbook.add_worksheet()
            with open(csvfile, 'rt', encoding='cp1252') as f:
                reader = csv.reader(f)
                for r, row in enumerate(reader):
                    for c, col in enumerate(row):
                        worksheet.write(r, c, col)  
            workbook.close()
        #tk.messagebox.showinfo('Information','File is converted and saved to this directory:  ' + HOME_PATH)

    @staticmethod
    def popup1(event, *args, **kwargs):
        Popupmenu1 = tk.Menu(root, tearoff=0)
        Popupmenu1.configure(activebackground="#f9f9f9")
        Popupmenu1.configure(activeborderwidth="1")
        Popupmenu1.configure(activeforeground="black")
        Popupmenu1.configure(background="#d9d9d9")
        Popupmenu1.configure(borderwidth="1")
        Popupmenu1.configure(disabledforeground="#a3a3a3")
        Popupmenu1.configure(font="-family {Segoe UI} -size 9")
        Popupmenu1.configure(foreground="black")
        Popupmenu1.post(event.x_root, event.y_root)

    @staticmethod
    def popup2(event, *args, **kwargs):
        Popupmenu2 = tk.Menu(root, tearoff=0)
        Popupmenu2.configure(activebackground="#f9f9f9")
        Popupmenu2.configure(activeborderwidth="1")
        Popupmenu2.configure(activeforeground="black")
        Popupmenu2.configure(background="#d9d9d9")
        Popupmenu2.configure(borderwidth="1")
        Popupmenu2.configure(disabledforeground="#a3a3a3")
        Popupmenu2.configure(font="-family {Segoe UI} -size 9")
        Popupmenu2.configure(foreground="black")
        Popupmenu2.post(event.x_root, event.y_root)

    @staticmethod
    def popup3(event, *args, **kwargs):
        Popupmenu3 = tk.Menu(root, tearoff=0)
        Popupmenu3.configure(activebackground="#f9f9f9")
        Popupmenu3.configure(activeborderwidth="1")
        Popupmenu3.configure(activeforeground="black")
        Popupmenu3.configure(background="#d9d9d9")
        Popupmenu3.configure(borderwidth="1")
        Popupmenu3.configure(disabledforeground="#a3a3a3")
        Popupmenu3.configure(font="-family {Segoe UI} -size 9")
        Popupmenu3.configure(foreground="black")
        Popupmenu3.post(event.x_root, event.y_root)

    @staticmethod
    def popup4(event, *args, **kwargs):
        Popupmenu4 = tk.Menu(root, tearoff=0)
        Popupmenu4.configure(activebackground="#f9f9f9")
        Popupmenu4.configure(activeborderwidth="1")
        Popupmenu4.configure(activeforeground="black")
        Popupmenu4.configure(background="#d9d9d9")
        Popupmenu4.configure(borderwidth="1")
        Popupmenu4.configure(disabledforeground="#a3a3a3")
        Popupmenu4.configure(font="-family {Segoe UI} -size 9")
        Popupmenu4.configure(foreground="black")
        Popupmenu4.post(event.x_root, event.y_root)

    @staticmethod
    def popup5(event, *args, **kwargs):
        Popupmenu5 = tk.Menu(root, tearoff=0)
        Popupmenu5.configure(activebackground="#f9f9f9")
        Popupmenu5.configure(activeborderwidth="1")
        Popupmenu5.configure(activeforeground="black")
        Popupmenu5.configure(background="#d9d9d9")
        Popupmenu5.configure(borderwidth="1")
        Popupmenu5.configure(disabledforeground="#a3a3a3")
        Popupmenu5.configure(font="-family {Segoe UI} -size 9")
        Popupmenu5.configure(foreground="black")
        Popupmenu5.post(event.x_root, event.y_root)

    @staticmethod
    def popup6(event, *args, **kwargs):
        Popupmenu6 = tk.Menu(root, tearoff=0)
        Popupmenu6.configure(activebackground="#f9f9f9")
        Popupmenu6.configure(activeborderwidth="1")
        Popupmenu6.configure(activeforeground="black")
        Popupmenu6.configure(background="#d9d9d9")
        Popupmenu6.configure(borderwidth="1")
        Popupmenu6.configure(disabledforeground="#a3a3a3")
        Popupmenu6.configure(font="-family {Segoe UI} -size 9")
        Popupmenu6.configure(foreground="black")
        Popupmenu6.post(event.x_root, event.y_root)

if __name__ == '__main__':
    vp_start_gui()





