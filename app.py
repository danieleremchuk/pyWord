# app.py - program entry point
import tkinter as tk
from tkinter import filedialog, Menu, messagebox, ttk, font as tkfont

from time import sleep
from PIL import Image, ImageTk

import file
import os

class DocToCSV(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        self.title_font = tkfont.Font(family='Tahoma', size=18, weight="bold")
        self.footer_font = tkfont.Font(family='Tahoma', size=12)

        window = tk.Frame(self)
        window.pack(side="top", fill="both", expand=True)
        window.grid_rowconfigure(0, weight=1)
        window.grid_columnconfigure(0, weight=1)

        menuBar = tk.Menu(self)
        fileMenu = tk.Menu(menuBar, tearoff = False)
        fileMenu.add_command(label = 'New File')
        fileMenu.add_command(label = 'Open File...')
        fileMenu.add('separator')
        fileMenu.add_command(label = 'Quit', command = self.destroy)
        helpMenu = tk.Menu(menuBar, tearoff = False)
        helpMenu.add_command(label = 'Help')
        helpMenu.add_command(label = 'About')
        menuBar.add_cascade(label = 'File', menu = fileMenu)
        menuBar.add_cascade(label = 'Help', menu = helpMenu)

        self.config(menu = menuBar)

        self.frames = {}
        for F in (MainPage, FindFile, BreweryName, ConvertFile):
            page_name = F.__name__
            frame = F(parent=window, controller=self)
            self.frames[page_name] = frame

            # put all of the pages in the same location;
            # the one on the top of the stacking order
            # will be the one that is visible.
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("MainPage")

    def show_frame(self, page_name):
        global doc
        global entry
        frame = self.frames[page_name]

        if "BreweryName" in page_name:
            file_name = doc.get()
            if not file.valid_file(file_name):
                tk.messagebox.showwarning(title="Please select a file", message="You must select a file before proceeding.")
            else:
                frame.tkraise()
        if "ConvertFile" in page_name:
            brew_name = entry.get()
            if not file.valid_name(brew_name):
                tk.messagebox.showwarning(title="Please enter a name", message="You must select a brewery name before proceeding.")
            else:
                frame.tkraise()
        elif "MainPage" in page_name or "FindFile" in page_name:
            frame.tkraise()

class MainPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        topFrame = tk.Frame(self)
        midFrame = tk.Frame(self)
        botFrame = tk.Frame(self)

        topFrame.pack(side="top", fill="both", expand=True)
        botFrame.pack(side="bottom", fill="both", expand=True)
        midFrame.pack(side="bottom", fill="both", expand=True)

        self.logoPath = 'gfx/logo.png'
        self.logoImg = ImageTk.PhotoImage(Image.open(self.logoPath))
        self.label_main = tk.Label(topFrame, text="Main Menu", font=controller.title_font, image=self.logoImg)
        self.label_main.pack(side="top", fill="x", pady=1)

        button_next = tk.Button(midFrame, text="Start Converting", command=lambda: controller.show_frame("FindFile"), padx=10, pady=5)
        button_next.pack(side="top", padx=10, pady=10)

        label_cr = tk.Label(botFrame, text="The Brewery Pioneers, Inc \u00a9", font=controller.footer_font)
        label_cr.pack(side="bottom", padx=10, pady=10)

class FindFile(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        global doc
        global path

        topFrame = tk.Frame(self)
        midFrame = tk.Frame(self)
        botFrame = tk.Frame(self)

        topFrame.pack(side="top", fill="both", expand=True)
        botFrame.pack(side="bottom", fill="both", expand=True)
        midFrame.pack(side="bottom", fill="both", expand=True)

        doc = tk.StringVar()
        doc.set("")
        path = ""

        label = tk.Label(topFrame, text="Get Document File", font=controller.title_font)
        label.pack(side="top", fill="x", pady=10)

        button_find = tk.Button(topFrame, text="Browse", command=self.find_doc, padx=10, pady=5)
        button_find.pack(side="left", padx=20, pady=10)

        entry_find = tk.Entry(topFrame, textvariable=doc, width=60)
        entry_find.pack(side="left", padx=10, pady=10)

        button_main = tk.Button(midFrame, text="Main Menu", command=self.main_func, padx=10, pady=5)
        button_next = tk.Button(midFrame, text="Next", command=lambda: controller.show_frame("BreweryName"), padx=10, pady=5)
        button_main.pack(side="bottom", padx=10, pady=10)
        button_next.pack(side="bottom", padx=10, pady=10)

        label_cr = tk.Label(botFrame, text="The Brewery Pioneers, Inc \u00a9", font=controller.footer_font)
        label_cr.pack(side="bottom", padx=10, pady=10)

    def find_doc(self):
        global doc
        global path

        this_doc = tk.filedialog.askopenfilename(initialdir="/", title="Select a file", filetypes=(("Microsoft Word (2007 - 365)", "*.docx"),("Legacy Microsoft Word (97-2000)", "*.doc")))
        my_file = os.path.basename(this_doc)
        doc.set(my_file)
        path = os.path.abspath(this_doc)

    def main_func(self):
        global doc

        self.controller.show_frame("MainPage")
        doc.set("")

class BreweryName(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        global entry_txt
        global entry

        topFrame = tk.Frame(self)
        midFrame = tk.Frame(self)
        botFrame = tk.Frame(self)

        topFrame.pack(side="top", anchor="center")
        botFrame.pack(side="bottom", fill="both", expand=True)
        midFrame.pack(side="bottom", fill="both", expand=True)

        label = tk.Label(topFrame, text="Brewery Name", font=controller.title_font)
        label.pack(side="top", fill="x", pady=10)

        label_entry = tk.Label(topFrame, text="Enter the Brewery Name: ")
        entry = tk.Entry(topFrame)
        label_entry.pack(side="left",fill="none", expand=True,padx = 10)
        entry.pack(side="left", fill="none", expand=True, padx = 10)

        button_main = tk.Button(midFrame, text="Main Menu", command=self.main_func, padx=10, pady=5)
        button_prev = tk.Button(midFrame, text="Previous", command=lambda: controller.show_frame("FindFile"), padx=10, pady=5)
        button_next = tk.Button(midFrame, text="Next", command=self.next_func, padx=10, pady=5)
        button_main.pack(side="bottom", padx=10, pady=10)
        button_prev.pack(side="bottom", padx=10, pady=10)
        button_next.pack(side="bottom", padx=10, pady=10)

        label_cr = tk.Label(botFrame, text="The Brewery Pioneers, Inc \u00a9", font=controller.footer_font)
        label_cr.pack(side="bottom", padx=10, pady=10)

    def next_func(self):
        global entry_txt
        global entry

        self.controller.show_frame("ConvertFile")
        entry_txt = entry.get()

    def main_func(self):
        global entry
        global doc

        self.controller.show_frame("MainPage")
        entry.delete(0, tk.END)
        doc.set("")

class ConvertFile(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        global progress
        global progress_txt

        progress = tk.DoubleVar()
        progress_txt = tk.StringVar()

        topFrame = tk.Frame(self)
        midFrame = tk.Frame(self)
        botFrame = tk.Frame(self)

        topFrame.pack(side="top", fill="both", expand=True)
        botFrame.pack(side="bottom", fill="both", expand=True)
        midFrame.pack(side="bottom", fill="both", expand=True)

        label = tk.Label(topFrame, text="Convert File", font=controller.title_font)
        label.pack(side="top", fill="x", pady=10)

        button_find = tk.Button(midFrame, text="Convert File", command=self.start_convert, padx=10, pady=5)
        button_find.pack(side="top", padx=10, pady=10)

        progress_pb = ttk.Progressbar(midFrame, orient="horizontal", length=300, mode='determinate', maximum=100, variable=progress)
        progress_text = tk.Label(midFrame, textvariable=progress_txt)
        button_main = tk.Button(midFrame, text="Main Menu", command=lambda: controller.show_frame("MainPage"), padx=10, pady=5)
        button_prev = tk.Button(midFrame, text="Previous", command=lambda: controller.show_frame("BreweryName"), padx=10, pady=5)
        button_main.pack(side="bottom", padx=10, pady=10)
        button_prev.pack(side="bottom", padx=10, pady=10)
        progress_text.pack(side="bottom", padx=10, pady=10)
        progress_pb.pack(side="bottom", padx=10, pady=10)

        label_cr = tk.Label(botFrame, text="The Brewery Pioneers, Inc \u00a9", font=controller.footer_font)
        label_cr.pack(side="bottom", padx=10, pady=10)

    def start_convert(self):
        global path
        global progress
        global progress_txt
        global entry_txt

        my_doc = path
        my_par = file.do_Doc(my_doc)

        i = 0
        while i < my_par+1:
            my_prog = round((i/my_par)*100)
            str_prog = str(my_prog)+ " %"
            progress.set(my_prog)
            if i == my_par:
                progress_txt.set("Conversion is complete")
            else:
                progress_txt.set(str_prog)
            i = i+1
            self.update()
            self.after(1000)

        my_brew = entry_txt
        file.do_CSV(my_par, my_doc, my_brew)

if __name__ == "__main__":
    app = DocToCSV()
    app.title("The Brewery Pioneers: CSV Import")

    w = 400
    h = 350

    # get screen width and height
    ws = app.winfo_screenwidth()
    hs = app.winfo_screenheight()

    # calculate x and y coordinates for the Tk root window
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)

    # set the dimensions of the screen and where it is placed
    app.geometry('%dx%d+%d+%d' % (w, h, x, y))
    app.mainloop()
