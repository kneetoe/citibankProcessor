import openpyxl, os
from tkinter.filedialog import askopenfile
import tkinter as tk
from tkinter.ttk import *
import ntpath
from pathlib import Path

# Project dependecies
from read import storeData
from listHandler import fileLoop
from excelMake import populate

# # open_file() # # 
# This function will be used to open
# file in read mode and only CVS files
# will be opened
def open_file(entry, but, go, check):
    but.config(bg='SystemButtonFace')
    
    ftypes = [('Comma seperated', '*.csv')]
    path = tk.filedialog.askopenfilename(filetypes = ftypes)
    entry.delete(0, tk.END)
    entry.insert(0, path)
    go.config(bg='yellow')
    check.config(bg='yellow')


# # invalidDir() # #
# Helper function for search and ensures path that is given
# creates error message and returns false if path is not valid 
def invalidDir(path, error):
    if(not os.path.exists(path)):
        error.config(text="Could not find File '" + path + "'")
    return os.path.exists(path)
    

# # search() # # 
# Ensures all files are present, and input is acceptable
# Then continues to call other functions from other python files
def search(ent_dir, error, root, var1):
    entry = ent_dir.get()
    main_exl = str(Path(os.getcwd()).parent) + '\currentYear.xlsx'
    auto = var1.get()
    
    if(entry == ""):
        error.config(text="Not a vaild file!")
        error.config(bg = 'red')
        error.config(fg = 'white')
        
    else:
        valid = True
        valid = valid and invalidDir('diction.txt', error)
        valid = valid and invalidDir('list.txt', error)
        valid = valid and invalidDir(str(entry), error)
        
        if(valid):
            try:
                book = openpyxl.load_workbook(main_exl)
                try:
                    book.save(main_exl)
                    root.destroy()
                    populate(fileLoop(storeData(entry, auto)), str(Path(os.getcwd()).parent), entry)
                    goodbyeWindow()
                except:
                    error.config(text="Please close the excel!")
                    error.config(bg = 'red')
                    error.config(fg = 'white')
            except:
                error.config(text="No file excel under the name 'currentYear'")
                error.config(bg = 'red')
                error.config(fg = 'white')

# # close_window() # # 
# Closes the tkinter window root that is given
def close_window(root): 
    root.destroy()

# # goodbyeWindow() # # 
# Window to notify user of excel use success
def goodbyeWindow():
    root = tk.Tk()
    root.geometry('+1000+300')
    root.title('CitiBank Helper')
    error = tk.Label(master=root,
            text='Completed, it is now safe to open excel.').pack(pady = 5, padx =5)
    
    button_quit = Button (root, text = "OK",
                     command = lambda:close_window(root))
    button_quit.pack(pady = 5, padx =5)
            
    
# # createMainWindow() # # 
# Creates the main window using tkinter
# initial function used to call the rest of the project
def createMainWindow():
    #Tk window
    root = tk.Tk()
    root.geometry('+1000+300')
    root.title('CitiBank Helper')


    # # FRAMES # #
    #Holds all other frames
    frame_folder = tk.Frame(master=root, relief=tk.SUNKEN, borderwidth=5)
    frame_dir = tk.Frame(master=frame_folder, relief=tk.GROOVE, borderwidth=5) 
    frame_btn = tk.Frame(master=frame_folder, relief=tk.FLAT, borderwidth=5)
    frame_error = tk.Frame(master=frame_folder, relief=tk.FLAT, borderwidth=5)


    # Title
    lbl_title = tk.Label(master=root,
            text="Citbank Helper")

    # DIR
    ent_dir = tk.Entry(fg="black",
                      bg="white",
                      width=20,
                      master = frame_dir
                      )
    lbl_file = tk.Label(master=frame_dir,
            text="Current File")
    # Error handling
    ent_error = tk.Label(master=frame_error,
            text="",
            )
    # Buttons
    var_auto = tk.IntVar()
    chk_auto = tk.Checkbutton(frame_dir, text="Only Auto", variable=var_auto)
    btn_go = tk.Button(
        frame_dir,
        text ='Go',
        command = lambda:search(ent_dir, ent_error, root, var_auto)
        )
    btn_open = tk.Button(frame_dir,
                         text ='Open',
                         command = lambda:open_file(ent_dir,
                                                    btn_open,
                                                    btn_go,
                                                    chk_auto),
                         bg="yellow")

    # Pack elements
    lbl_title.pack(side=tk.TOP, pady = 10, padx = 10)
    ent_dir.pack(side=tk.LEFT, padx = 10)
    lbl_file.pack(side=tk.LEFT)
    ent_error.pack()
    chk_auto.pack(side = tk.LEFT)
    btn_go.pack(side = tk.RIGHT)
    btn_open.pack(side = tk.RIGHT)
    
    # Pack frames
    frame_dir.pack(side = tk.TOP, pady = 10, padx = 10)
    frame_error.pack(side = tk.TOP)
    frame_btn.pack(side = tk.BOTTOM, pady = 10, padx = 10)
    frame_folder.pack(side = tk.RIGHT, pady = 10, padx = 10)

    root.mainloop()

if __name__ == '__main__':
    createMainWindow()
