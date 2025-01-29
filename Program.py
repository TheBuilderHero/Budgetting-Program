# make sure to pip3 install openpyxl
import openpyxl
# File GUI:
import tkinter as tk
import tkinter.filedialog as filedialog
from tkinter import ttk
from tkinter.filedialog import asksaveasfile 
# CSV to Excel formatting:
import csv
import os # file extension removal
# Menu bar:
from tkinter import Menu
from tkinter import OptionMenu
#TreeView data to Excel:
import pandas as pd

excel_save_file_extension = '.xlsx'



r = tk.Tk()
# Adjust size
# I want it resizable so that the scrolling and the data boxes are not messed up by it.
r.minsize(width=600,height=10)
r.resizable(width=False, height=False)
r.title('Budgetting App')


# Using treeview widget    
treev = ttk.Treeview(r, height=20, selectmode ='extended')

'''
    Setting the Select Mode:
        browse: (Default) Only one item can be selected at a time. Clicking on an item selects it and deselects any previously selected item.
        extended: Multiple items can be selected by holding down the Shift key or Ctrl key (Command key on macOS).
        none: Disables selection entirely.
'''

# Constructing vertical scrollbar
# with treeview
verscrlbar = ttk.Scrollbar(r, 
                           orient ="vertical", 
                           command = treev.yview)
# Constructing horizontal scrollbar
# with treeview
horzscrlbar = ttk.Scrollbar(r, 
                           orient ="horizontal", 
                           command = treev.xview)

# Calling pack method w.r.to vertical 
# scrollbar
verscrlbar.pack(side ='right', fill ='y', anchor='ne')
horzscrlbar.pack(side='bottom', fill='x', anchor='sw')

# Calling pack method w.r.to treeview
treev.pack(side ='left', fill='x', anchor='nw')
 
# Configuring treeview
treev.configure(yscrollcommand = verscrlbar.set, xscrollcommand=horzscrlbar.set)

# The above packing was done in this order to help with the window layout. Other packing orders have not yeilded this good of a result.

def hide_object(ob):
    if type(ob) is not list:
        ob.pack_forget()
        return
    print(ob)
    for ite in ob:
        ite.pack_forget()
def show_object(ob, appen, fill):
    if type(ob) is not list:
        ob.pack(expand = True, fill=fill, side=appen)
        return
    for ite in ob:
        ite.pack(expand = True, fill=fill, side=appen)
def open_file():
    file_path = filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx'),('Excel', '*.xlsm'),('Excel', '*.xlsb'),('Excel', '*.xltx'),('Excel','*.xls'), ('CSV files', '*.csv')])
    if file_path:
        print("Selected file:", file_path)
        return file_path
    else:
        print("No file selected.")
        return False

def is_decimal_string(input):
    import re

    pattern = r"-?\d+\.\d+"

    match = re.search(pattern, input)

    if match:
        return True
    else:
        return False

def convert_csv_to_excel(fileName): #returns the new file name
    wb = openpyxl.Workbook()
    ws = wb.active
    with open(file=fileName, mode='r') as f:
        reader = csv.reader(f, delimiter=',')
        for row in reader:
            # Cell to modify
            for i in range(len(row)):
                # Convert the value to a number and remove leading zeros
                if is_decimal_string(row[i]): 
                    try:
                        row[i] = float(row[i])
                    except ValueError:
                        pass  # Handle the case where the cell doesn't contain a number
            ws.append(row)

    filename_without_extension = os.path.splitext(fileName)[0]

    newFileName = filename_without_extension + excel_save_file_extension
    wb.save(newFileName)

    return newFileName

    
    #Note for future if I need more control over the data going into each row:
    """ 
    with open('classics.csv') as f:
        reader = csv.reader(f, delimiter=',')

        for row_index, row in enumerate(reader, start=1):
            for column_index, cell_value in enumerate(row, start=1):
            ws.cell(row=row_index, column=column_index).value=cell_value
    """



def load_excel_file():
    file_path = open_file()
    if file_path:
        
        #check if files are CSV then convert to excel:
        file_extension = os.path.splitext(file_path)[1]
        if file_extension == '.CSV':
            file_path = convert_csv_to_excel(file_path)
        
        # Load the workbook
        wb = openpyxl.load_workbook(file_path)

        # Select the active sheet
        sheet = wb.active
        
        firstTimeOver = True
        columnsInTree = ()

        # Read and print the data
        for row in sheet.iter_rows(min_row=1, values_only=True):
            # After the first run through row 1 we will no longer be adding columns
            if firstTimeOver is True:
                    # Defining number of columns
                    treev['columns'] = (row)
            else:
                treev.insert("", 'end', text ="L1", values =(row))
            for i,col in enumerate(row):
                if firstTimeOver is True:
                    # Assigning the heading names to the respective columns
                    treev.column(str(i))
                    treev.heading(column=str(i), text=col)
            
            firstTimeOver = False # After the first run through row 1 we will no longer be adding columns

        # Defining heading
        treev['show'] = 'headings'


#Tree Saving to file:
def write_to_excel(tree, filename):
    """Writes data from the Treeview to an Excel file."""

    # Get the data from the Treeview
    data = []
    for item in tree.get_children():
        values = tree.item(item, "values")
        data.append(values)

    # Create a Pandas DataFrame from the data
    df = pd.DataFrame(data, columns=[tree.heading(col, "text") for col in tree["columns"]])

    # Save the DataFrame to an Excel file
    df.to_excel(filename.name, index=False)

def file_save_as():
    files = [('Excel File', '*.xlsx')]   
            # Not gonna allow saving as CSV since we want it in the fomat of excel anyway.
            #,('CSV File', '*.csv')
    file = asksaveasfile(filetypes = files, defaultextension = files) 
    #with open(file, "wb") as f:
    write_to_excel(treev, file)


#Adding a menu Bar:

def donothing():
    return

menubar = Menu(r)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="New", command=donothing)
filemenu.add_command(label="Open", command=load_excel_file)
filemenu.add_command(label="Save", command=file_save_as)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=r.quit)
menubar.add_cascade(label="File", menu=filemenu)

helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Help Index", command=donothing)
helpmenu.add_command(label="About...", command=donothing)
menubar.add_cascade(label="Help", menu=helpmenu)

r.config(menu=menubar)


'''
Widgets are added here
'''

r.mainloop()