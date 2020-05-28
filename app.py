import tkinter as tk
from tkinter import filedialog
from tkinter import font as tkfont 
from tkinter import ttk
from time import sleep
import files
import validate
import os

global doc
global path
global progress
global progress_txt

class DocToCSV(tk.Tk):
    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        self.title_font = tkfont.Font(family='Tahoma', size=18, weight="bold")
        self.footer_font = tkfont.Font(family='Tahoma', size=12)

        window = tk.Frame(self)
        window.pack(side="top", fill="both", expand=True)
        window.grid_rowconfigure(0, weight=1)
        window.grid_columnconfigure(0, weight=1)

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
            if not validate.valid_file(file_name):
                file_msg = validate.valid_file_msg(file_name)
                file_title = "Please select a file"
                self.popup_err(file_msg, file_title)
            else: 
                frame.tkraise()
        if "ConvertFile" in page_name:
            brew_name = entry.get()
            if not validate.valid_name(brew_name):
                name_msg = validate.valid_name_msg(brew_name)
                name_title = "Please enter a name"
                self.popup_err(name_msg, name_title)
            else: 
                frame.tkraise()
        elif "MainPage" in page_name or "FindFile" in page_name:
            frame.tkraise()

    def popup_err(self, msg, title):
        wind = tk.Toplevel(self)
        wind.title(title)
        wind_msg = tk.Label(wind, text=msg, fg="red", font="Tahoma 12 bold", padx=10, pady=10)
        wind_msg.pack(side="top")
        btn_ok = tk.Button(wind, text="OK", command=wind.destroy, padx=10, pady=5)
        btn_ok.pack(side="bottom", padx=10, pady=10)

        wind_w = 400 
        wind_h = 100 

        # get screen width and height
        wind_ws = wind.winfo_screenwidth() 
        wind_hs = wind.winfo_screenheight() 

        # calculate x and y coordinates for the Tk root window
        wind_x = (wind_ws/2) - (wind_w/2)
        wind_y = (wind_hs/2) - (wind_h/2)

        # set the dimensions of the screen and where it is placed
        wind.geometry('%dx%d+%d+%d' % (wind_w, wind_h, wind_x, wind_y))

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

        label_main = tk.Label(topFrame, text="Main Menu", font=controller.title_font)
        label_main.pack(side="top", fill="x", pady=10)

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

        button_find = tk.Button(midFrame, text="Find File", command=self.find_doc, padx=10, pady=5)
        button_find.pack(side="top", padx=10, pady=10)

        label_find = tk.Label(midFrame, textvariable=doc)
        label_find.pack(side="top", padx=10, pady=10)

        button_main = tk.Button(midFrame, text="Main Menu", command=self.main_func, padx=10, pady=5)
        button_next = tk.Button(midFrame, text="Next", command=lambda: controller.show_frame("BreweryName"), padx=10, pady=5)
        button_main.pack(side="bottom", padx=10, pady=10)
        button_next.pack(side="bottom", padx=10, pady=10)

        label_cr = tk.Label(botFrame, text="The Brewery Pioneers, Inc \u00a9", font=controller.footer_font)
        label_cr.pack(side="bottom", padx=10, pady=10)

    def find_doc(self):
        global doc
        global path

        this_doc = tk.filedialog.askopenfilename(initialdir="/", title="Select a file", filetypes=(("docx files", "*.docx"),("doc files", "*.doc")))
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

        topFrame.pack(side="top", fill="both", expand=True)
        botFrame.pack(side="bottom", fill="both", expand=True)
        midFrame.pack(side="bottom", fill="both", expand=True)

        label = tk.Label(topFrame, text="Brewery Name", font=controller.title_font)
        label.pack(side="top", fill="x", pady=10)    
        
        label_entry = tk.Label(midFrame, text="Enter the Brewery")
        entry = tk.Entry(midFrame) 
        label_entry.pack(side="top")
        entry.pack(side="top")

        button_main = tk.Button(midFrame, text="Main Menu", command=self.main_func, padx=10, pady=5)
        button_next = tk.Button(midFrame, text="Next", command=self.next_func, padx=10, pady=5)
        button_prev = tk.Button(midFrame, text="Previous", command=lambda: controller.show_frame("FindFile"), padx=10, pady=5)
        button_main.pack(side="bottom", padx=10, pady=10)
        button_next.pack(side="bottom", padx=10, pady=10)
        button_prev.pack(side="bottom", padx=10, pady=10)

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
        button_prev = tk.Button(midFrame, text="Previous", command=lambda: controller.show_frame("FindFile"), padx=10, pady=5)
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
        my_par = files.do_Doc(my_doc)

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
        files.do_CSV(my_par, my_doc, my_brew)

if __name__ == "__main__":
    app = DocToCSV()
    app.title("The Brewery Pioneers: CSV Import")

    w = 600 
    h = 600 

    # get screen width and height
    ws = app.winfo_screenwidth() 
    hs = app.winfo_screenheight() 

    # calculate x and y coordinates for the Tk root window
    x = (ws/2) - (w/2)
    y = (hs/2) - (h/2)

    # set the dimensions of the screen and where it is placed
    app.geometry('%dx%d+%d+%d' % (w, h, x, y))
    app.mainloop()