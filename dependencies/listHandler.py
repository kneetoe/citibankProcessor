import os, openpyxl, time
import tkinter as tk
from tkinter.ttk import *
from read import storeData
from pathlib import Path

categories = {}
categoriesNames = {}
curCat = "NONE"
curNote = "NONE"
totalItems = 0
curIndexItems = 1

###TODO: deny enter if no text
test = True

##Repeated code
def close_window(root):
    root.destroy()

def changeDesc(new, error):
    global curCat
    curCat = new
    error.config(text=new)

def addListHelp(ent_catg, error, root):
    new = ent_catg.get()
    print(new)
    changeDesc(new, error)

    if(new.upper() != ''):
        categoriesNames[new] = 0
        f = open("list.txt", "a")
        f.write(('\n'))
        f.write(new)
        f.close()
    close_window(root)

desc=0
def retFun2(event):
    global ent_catg, root, desc
    addListHelp(ent_catg, desc, root)

def addList(description):
    global ent_catg, root, desc
    desc = description

    root = tk.Tk()
    root.geometry('+1000+300')
    root.after(1, lambda: root.focus_force())
    frm_entry = tk.Frame(master=root, relief=tk.FLAT, borderwidth=5)
    error = tk.Label(master=frm_entry,
            text='New Category: ').pack()
    ent_catg = tk.Entry(fg="black",
        bg="white",
        width=10,
        master = frm_entry
        )
    ent_catg.pack()
    error = tk.Label(master=root,
            text="Enter a new category name below")


    button_auto = Button (root, text = "Add",
                     command = lambda:addListHelp(ent_catg, desc, root))
    button_quit = Button (root, text = "Quit",
                     command = lambda:close_window(root))
    ent_catg.focus()
    ent_catg.bind('<Return>', retFun2)

    error.pack()
    frm_entry.pack()
    button_auto.pack()
    button_quit.pack()



def addDictList(ent_search, ent_catg, error, root):
    new1 = ent_search.get()
    new2 = ent_catg.get()

    if(new1.upper() != '' and new2.upper() != ''):
        categories[new1] = new2
        f = open("diction.txt", "a")
        f.write(('\n'))
        f.write(new1.upper())
        f.write(":")
        f.write(new2.upper())
        f.close()
        error.config(text="ADDED ENTRY")
        ent_search.delete(0, tk.END)
        ent_catg.delete(0, tk.END)
        close_window(root)
    else:
        error.config(text="Not Valid")

ent_search = 0
ent_catg = 0
error = 0
root = 0

def retFun1(event):
    global ent_search, ent_catg, error, root
    addDictList(ent_search, ent_catg, error, root)

def addDictCat():
    global ent_search, ent_catg, error, root
    global curCat
    root = tk.Tk()
    root.geometry('+1000+300')
    root.after(1, lambda: root.focus_force())

    frm_entry = tk.Frame(master=root, relief=tk.FLAT, borderwidth=5)
    frm_entry2 = tk.Frame(master=root, relief=tk.FLAT, borderwidth=5)
    error = tk.Label(master=frm_entry,
            text='Search for:').pack()
    ent_search = tk.Entry(fg="black",
        bg="white",
        width=10,
        master = frm_entry,
        takefocus = 0
        )
    ent_search.focus()

    error = tk.Label(master=frm_entry2,
            text='Category:').pack()

    ent_catg = tk.Entry(fg="black",
        bg="white",
        width=10,
        master = frm_entry2
                        )
    error = tk.Label(master=root,
            text="")
    button_auto = Button (root, text = "Add",
                     command = lambda:addDictList(ent_search, ent_catg, error, root))
    ent_search.bind('<Return>', retFun1)
    ent_catg.bind('<Return>', retFun1)


    ent_catg.pack(side = tk.TOP)
    ent_search.pack(side = tk.BOTTOM)
    frm_entry.pack(side = tk.TOP)
    frm_entry2.pack(side = tk.TOP)
    error.pack()
    button_auto.pack(side = tk.BOTTOM, pady=5)


def changeComm(ent_comm, error, root):
    global curNote
    new1 = ent_comm.get()
    if(new1 != ''):
        curNote = new1
        close_window(root)
    else:
        error.config(text="Could not add Comment")

ent_comm = 0
error = 0
root = 0

def retFun(event):
    global ent_comm, error, root
    changeComm(ent_comm, error, root)

def addComm():
    global ent_comm, error, root

    root = tk.Tk()
    root.geometry('+1000+300')
    root.after(1, lambda: root.focus_force())
    frm_entry = tk.Frame(master=root, relief=tk.FLAT, borderwidth=5)
    error = tk.Label(master=frm_entry,
            text='Comment: ').pack()
    ent_comm = tk.Entry(fg="black",
        bg="white",
        width=20,
        master = frm_entry
        )
    ent_comm.pack()
    error = tk.Label(master=root,
            text="Enter you're comment below")


    button_auto = Button (root, text = "Add",
                     command = lambda:changeComm(ent_comm, error, root))
    button_quit = Button (root, text = "Quit",
                     command = lambda:close_window(root))
    ent_comm.focus()
    ent_comm.bind('<Return>', retFun)

    error.pack()
    frm_entry.pack()
    button_auto.pack()
    button_quit.pack()

def helpexit():
    exit()

## createButton() ##
# Helper Function for create Window
    # creates buttons from dictionary list
    # max 5 buttons on each row
    # creates temp frame
def createButton(buttonFrame, dspl_catg):
    i = 1
    for x, y in categoriesNames.items():
        if(i%5 == 0):
            frm_temp.pack()
            frm_temp = tk.Frame(master=buttonFrame, borderwidth=5)
        elif (i == 1):
            frm_temp = tk.Frame(master=buttonFrame, borderwidth=5)
        #create button for category that will change current desc
        btn_save = tk.Button(
            frm_temp,
            text = x,
            command = lambda x=x:changeDesc(x, dspl_catg)
            ).pack(side = tk.RIGHT)
        i = i+1
    frm_temp.pack()


def createWindow(date, desc, debit, person, catg):
    global totalItems
    global curIndexItems
    global curCat

    # Establish window and Vars
    root = tk.Tk()
    root.title('CitiBank Helper')
    root.geometry('+1000+300')

    curCat = catg
    curIndexItems = curIndexItems + 1

    # Frames
    frm_folder = tk.Frame(master=root, relief=tk.SUNKEN, borderwidth=5)
    frm_btn = tk.Frame(master=frm_folder, relief=tk.GROOVE, borderwidth=5, width = 200)
    frm_dspl_catg = tk.Frame(master=frm_folder, relief=tk.GROOVE, borderwidth=5)
    frm_info = tk.Frame(master=frm_folder, relief=tk.FLAT, borderwidth=5)
    
    frm_desc = tk.Frame(master=frm_info, relief=tk.FLAT, borderwidth=5)
    frm_person = tk.Frame(master=frm_info, relief=tk.FLAT, borderwidth=5)
    frm_date = tk.Frame(master=frm_info, relief=tk.FLAT, borderwidth=5)
    frm_debit = tk.Frame(master=frm_info, relief=tk.FLAT, borderwidth=5)
    
    frm_add = tk.Frame(master=frm_btn, borderwidth=5)

    # Direction
    lbl_direct = tk.Label(master=root,
            text="Please select category ("
             + str(curIndexItems) + "\\"
             + str(totalItems) + ')'
             )

    # Desc labels
    lbl_desc_a = tk.Label(master=frm_desc,
            text="Description:")
    lbl_desc_b = tk.Label(master=frm_desc,
            text=desc)

    # Person labels
    lbl_pers_a = tk.Label(master=frm_person,
            text="Person:")
    lbl_pers_b = tk.Label(master=frm_person,
            text=person)

    # Date labels
    lbl_date_a = tk.Label(master=frm_date,
            text="Date:")
    lbl_date_b = tk.Label(master=frm_date,
            text=date)

    # Debit labels
    lbl_debit_a = tk.Label(master=frm_debit,
            text="Cost:")
    lbl_debit_b = tk.Label(master=frm_debit,
            text=('$' + debit))

    #dspl_catg handling
    dspl_catg = tk.Label(master=frm_dspl_catg,
            text=catg)
    

    # Buttons [catgs]
    createButton(frm_btn, dspl_catg)
    

    # New button [catgs]
    btn_new = tk.Button(
        frm_add,
        text = 'Add new',
        command = lambda:addList(dspl_catg)
        )

    # None button [catgs]
    btn_save = tk.Button(
        frm_add,
        text = 'NONE',
        command = lambda:changeDesc('NONE', dspl_catg)
        )
    # Buttons [GUI]
    btn_auto = Button (root, text = "Add Auto",
                     command = addDictCat)
    btn_quit = Button (root, text = "Quit",
                     command = helpexit )
    btn_comm = Button (frm_btn, text = "Add Comment",
                     command = addComm )
    btn_nxt = Button (frm_dspl_catg, text = "Next",
                     command = lambda:close_window(root))

    #Packing
    lbl_direct.pack(side=tk.TOP, pady = 10, padx = 10)

    lbl_desc_a.pack(side=tk.LEFT)
    lbl_desc_b.pack(side=tk.LEFT)
    frm_desc.pack()

    lbl_pers_a.pack(side=tk.LEFT)
    lbl_pers_b.pack(side=tk.LEFT)
    frm_desc.pack(side=tk.BOTTOM)

    lbl_date_a.pack(side=tk.LEFT)
    lbl_date_b.pack(side=tk.LEFT)
    frm_date.pack(side=tk.BOTTOM)

    lbl_debit_a.pack(side=tk.LEFT)
    lbl_debit_b.pack(side=tk.LEFT)
    frm_debit.pack(side=tk.BOTTOM)

    dspl_catg.pack()

    btn_new.pack(side = tk.RIGHT)
    btn_save.pack(side = tk.RIGHT)
    frm_add.pack()

    btn_auto.pack()
    btn_quit.pack()
    btn_comm.pack()
    btn_nxt.pack(side =tk.RIGHT)

    frm_info.pack(side = tk.TOP, pady = 10, padx = 10)
    frm_dspl_catg.pack(side = tk.TOP)
    frm_btn.pack(side = tk.BOTTOM, pady = 10, padx = 10,)
    frm_folder.pack(side = tk.RIGHT, pady = 10, padx = 10)

    root.mainloop()


def description(desc):
    count = 0
    catg = 'NONE'
    for x, y in categories.items():
        if(x in desc.upper()):
            if(catg == 'NONE'):
                catg = y
                count = count+1
            elif(catg != y):
                catg = catg + ' ' + y
                count = count+1

    if(count > 1):
        catg = 'MULTI [' + catg + ']'

    return catg

def credit(credit, catg):
    if(credit == '0'):
        catg = 'CREDIT'
    return catg



def fileLoop(data):
    global curCat
    global curNote
    global totalItems
    values = len(data[0])
    i = 1

    if(data[7][0] == True):
        skip = True
        data[7][0] = 'NONE'
    else:
        skip = False
        data[7][0] = 'NONE'
    # Load dictionary
    f = open("diction.txt", "r")
    for line in f:
        categories[line.split(":")[0]] = line.split(":")[1]
    f.close()

    f = open("list.txt", "r")
    for line in f:
        categoriesNames[line.split("\n")[0]] = 0
    f.close()

    if(not skip):
        totalItems = items(data, values) - 1

    while(i < values):

        data[6][i] = description(data[2][i])
        data[6][i] = credit(data[3][i], data[6][i])


        if(data[6][i] == 'NONE'):
            curNote = "NONE"
            if(not skip):
                data[6][i] = description(data[6][i])
                curCat = 'NONE'
                createWindow(data[1][i], data[2][i], data[3][i], data[5][i], data[6][i])
                data[6][i] = curCat
            data[8][i] = curNote
        i = i+1

    cont = checkIfWrite()
    while(not cont):
        cont = checkIfWrite()


    return data

def checkIfWrite():
    pass_bool = False
    print((str(Path(os.getcwd()).parent)))
    try:
        book = openpyxl.load_workbook(str(Path(os.getcwd()).parent) + '\currentYear.xlsx')
        try:
            book.save(str(Path(os.getcwd()).parent) + '\currentYear.xlsx')
            pass_bool = True
        except:
            warningWindow("Please close the excel!")
    except:
        warningWindow("No file excel under the name 'currentYear'")

    return pass_bool

def warningWindow(warn):
    root = tk.Tk()
    error = tk.Label(master=root,
        text=warn,
        )
    error.pack(pady=5)

    btn_retry = tk.Button(
        root,
        text ='retry',
        command = lambda:close_window(root)
        )
    btn_retry.pack(pady=5)
    root.mainloop()


def items(data, values):
    j = 1
    i = 1
    while(i < values):
        data[6][i] = description(data[2][i])
        data[6][i] = credit(data[3][i], data[6][i])

        if(data[6][i] == 'NONE'):
            j = j + 1
        i = i+1
    return j


if __name__ == '__main__':
    if(test):
        fileLoop(storeData('Statement closed Jun 24, 2020.csv', False))
    else:
        print('Please open fileWindow.py first please.')
        print('Closing window in:')
        i = 5
        while(i > 0):
            print(str(i) + ' seconds')
            time.sleep(1)
            i = i-1
