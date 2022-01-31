from tkinter import *
from tkinter import filedialog
from os import path
from Converter import Converter
from threading import *

supported = True

def browse_files():
    file_name = filedialog.askopenfilename(initialdir="/", title="Select a File",
                                           filetypes=(("all files", "*.*"), ("KML files", "*.kml*"), ("KMZ files", "*.kmz*"), ("ZIP files", "*.zip*")))
    if len(file_name) > 0:
        if file_name.split(".")[-1] in ["kml", "kmz", "zip"]:
            supported = True
            converter.set_file_name(file_name)
            # Change label contents
            label_file_explorer.configure(text="File: " + file_name)
            if not converter.chosen:
                for i in range(len(file_name) - 1, 0, -1):
                    if file_name[i] == '/':
                        label_folder_explorer.configure(text="Destination: " + file_name[:i])
                        break
        else:
            supported = False
            label_status.configure(text="File Not Supported", fg="red")



def browse_folders():
    # Allow user to select a directory and store it in global var
    # called folder_path
    folder_name = filedialog.askdirectory()
    if len(folder_name) > 0:
        converter.set_folder_name(folder_name)
        converter.chosen = True
        label_folder_explorer.configure(text="Destination: " + folder_name)

def reset_folder():
    if converter.get_file_name() == "":
        converter.set_folder_name("")
        converter.chosen = None
        label_folder_explorer.configure(text="Destination")
    else:
        converter.set_folder_name(converter.get_file_name().rsplit("/")[0])
        converter.chosen = None
        label_folder_explorer.configure(text= "Destination: " + label_file_explorer.cget("text").rsplit(" ", 1)[1].rsplit("/", 1)[0])

def xlsx_path():
    #converter.xlsx_name = converter.final_file_name()
    if converter.get_file_name().rsplit(".", 1)[1] != "zip":
        return path.exists(converter.final_file_name())
    return False

def threading():
    t1 = Thread(target=convert)
    t1.start()

def convert():
    if not supported:
        label_status.configure(text="FILE NOT SUPPORTED", fg="red")
        return
    elif converter.get_file_name() == "":
        if converter.get_folder_name() == None:
            label_status.configure(text="Choose a File or Folder", fg="orange")
        else:
            try:
                label_status.configure(text="Status: Converting...", fg="blue")
                converter.ConvertFolder()
                converter.clear()
                label_status.configure(text="Conversion Successful, file location: \n" + converter.xlsx_name,
                                       fg="green")
            except:
                label_status.configure(text="Error in conversion", fg="red")
                # label_file_explorer.configure(text="Error occurred in conversion, email me")
    elif xlsx_path():
        label_status.configure(text="File already exists in chosen directory", fg="red")
    elif converter.get_file_name().split(".")[-1] in ["kml", "kmz"]:
        label_status.configure(text="Status: Converting...", fg="blue")
        window.update_idletasks()
        try:
            converter.ConvertFile()
            converter.clear()
            label_status.configure(text="Conversion Successful, file location: \n" + converter.xlsx_name,
                                    fg="green")
        except:
            label_status.configure(text="Error in conversion", fg="red")
            # label_file_explorer.configure(text="Error occurred in conversion, email me")
    elif converter.get_file_name().split(".")[-1] == "zip":
        label_status.configure(text="Status: Converting...", fg="blue")
        window.update_idletasks()
        try:
            converter.ConvertZip()
            converter.clear()
            label_status.configure(text="Conversion Successful, file location: \n" + converter.xlsx_name,
                                    fg="green")
        except:
            label_status.configure(text="Error in conversion", fg="red")
            # label_file_explorer.configure(text="Error occurred in conversion, email me")
    else:
        label_status.configure(text="File not Compatable", fg="red")


#Create the root window
converter = Converter()

window = Tk()

var = IntVar()

def sel():
    converter.check_var = var.get()

def quit():
    window.destroy()

# Set window title
window.title('.kml to .csv converter')

# Set window size
window.geometry("700x500")

# Set window background color
window.config(background="white")

browse_frame = Frame(window, width=700, height=25)
browse_frame.pack(side="top", fill="x")

buttons_frame = Frame(window, width=700, height=25, pady="25")
buttons_frame.pack(side="top", fill="x")

label_frame = Frame(window, width=700, height=100, pady=25)
label_frame.pack(side="top", fill="x")


#clear = Button(btns_frame, text = "C", fg = "black", width = 32, height = 3, bd = 0, bg = "#eee", cursor = "hand2", command = lambda: btn_clear()).grid(row = 0, column = 0, columnspan = 3, padx = 1, pady = 1)

# Create a File Explorer label
label_file_explorer = Label(browse_frame, text="No File Chosen", width=52, height=2, fg="blue")#.grid(row=0, column=1, columnspan=3, padx=1, pady=1)

button_explore = Button(browse_frame, text="Browse Files", width=13, command=browse_files)#.grid(row=0, column=0, padx=15, pady=1)

label_folder_explorer = Label(browse_frame, text="Destination", width=52, height=2, fg="blue")#.grid(row=1, column=1, columnspan=3, padx=1, pady=1)

button_folder_explore = Button(browse_frame, text="Browse Folders", width=13, command=browse_folders)#.grid(row=1, column=0, padx=15, pady=1)

button_folder_reset = Button(browse_frame, text="Reset", command=reset_folder)#.grid(row=1, column=3, padx=1, pady=1)

check_xlsxAndkmz = Checkbutton(buttons_frame, text="xlsx and kmz", width=9, variable=var, onvalue=1, offvalue=0, command=sel)     #.grid(row=0, column=0)

button_run = Button(buttons_frame, text="Convert", width=13, command=threading, justify=LEFT)                #.grid(row=0, column=1)

button_exit = Button(buttons_frame, text="Exit", width=13, command=quit, justify=RIGHT)                     #.grid(row=0, column=2)

label_status = Label(label_frame, text="Status: waiting for file", width=20, height=4, fg="blue", bg="white")         #.grid(row=1, column=0)

# -----------------------

label_file_explorer.grid(row=0, column=5, columnspan=3, padx=1, pady=1)

button_explore.grid(row=0, column=0, padx=15, pady=1)

label_folder_explorer.grid(row=1, column=5, columnspan=3, padx=1, pady=1)

button_folder_explore.grid(row=1, column=0, padx=15, pady=1)

button_folder_reset.grid(row=1, column=500, padx=1, pady=1)

check_xlsxAndkmz.grid(row=0, column=0, padx=15, pady=1)

button_run.grid(row=0, column=1, padx=180, pady=1)

button_exit.grid(row=0, column=2, padx=1, pady=1)

# check_xlsxAndkmz.pack(side="left")
# button_run.pack(side="left")
# button_exit.pack()
label_status.grid(row=0, column=0)
window.mainloop()
