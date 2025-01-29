# make sure to pip3 install openpyxl
import openpyxl
# File GUI:
import tkinter as tk
import tkinter.filedialog as filedialog
from tkinter import ttk
from tkinter.filedialog import asksaveasfile 
from tkinter import messagebox
from tkinter import LabelFrame
from tkinter import Checkbutton
from tkinter import IntVar
# CSV to Excel formatting:
import csv
import os # file extension removal
# Menu bar:
from tkinter import Menu
from tkinter import OptionMenu
#TreeView data to Excel:
import pandas as pd

excel_save_file_extension = '.xlsx'
hasLoadedFileData = False # this if used to tell if we have data in the TreeView or not for saving it when opening a new file.
col_storage_pd = pd.DataFrame()
wb = None



r = tk.Tk()
# Adjust size
# I want it resizable so that the scrolling and the data boxes are not messed up by it.
r.minsize(width=600,height=10)
r.resizable(width=False, height=False)
r.title('Budgetting App')

# This will create a LabelFrame
label_frame = LabelFrame(r, text='Columns Included and useful', font = "50")

ButtonsN = []

# List to hold the checkboxes' variable references
check_vars = []




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


label_frame.pack(side='right', anchor='center', padx=20)
 
# Configuring treeview
treev.configure(yscrollcommand = verscrlbar.set, xscrollcommand=horzscrlbar.set)

# The above packing was done in this order to help with the window layout. Other packing orders have not yeilded this good of a result.

def hide_object(ob):
    if type(ob) is not list:
        ob.pack_forget()
        return
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
    global excel_save_file_extension
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

def show_option_dialog(title, question):
    result = messagebox.askquestion(title, question)
    if result == 'yes':
        return True
    else:
        return False

def file_save_as():
    files = [('Excel File', '*.xlsx')]   
            # Not gonna allow saving as CSV since we want it in the fomat of excel anyway.
            #,('CSV File', '*.csv')
    file = asksaveasfile(filetypes = files, defaultextension = files) 
    #with open(file, "wb") as f:
    if file:
        write_to_excel(treev, file)
        return True
    else:
        return False
    

def remove_column(tree, index):

    # Remove the second column
    tree.column(treev['columns'][index], width=0, stretch=False)
    tree.heading(treev['columns'][index], text="")
    print(treev['columns'])
    #for item in treev[]

    '''# Get the column names
    columns = tree["columns"]

    # Remove the specified column from the list
    columns = [col for i,col in enumerate(columns) if i != column_id]

    # Create a new Treeview with the updated columns
    new_tree = ttk.Treeview(tree.master, columns=columns, show="headings")

    # Copy the headings
    for col in columns:
        new_tree.heading(col, text=tree.heading(col)["text"])

    # Copy the data
    for item in tree.get_children():
        values = tree.item(item, "values")
        new_values = [values[i] for i, col in enumerate(columns) if col != column_id]
        new_tree.insert("", "end", values=new_values)
        treev.delete(item)

    # Replace the old Treeview with the new one
    tree.destroy()
    return new_tre'''

def copy_column(column_index):
    global treev
    """Copy selected column from Treeview to a pandas DataFrame column."""

    # Get the column name 
    column_name = treev.heading(column_index, "text")

    # Create a list of values in the selected column
    column_values = []
    for child in treev.get_children():
        column_values.append(treev.item(child, "values")[column_index])

    # Create a pandas DataFrame from the column values
    df = pd.DataFrame({column_name: column_values})

    return df
    
# Callback function to handle the checkbox state change
def checkbox_state_changed(index):
    global col_storage_pd
    global treev
    """Callback function that gets triggered when a checkbox state changes."""
    state = check_vars[index].get()  # Get the state of the checkbox
    #print(f"Checkbox {index + 1} is {'Checked' if state else 'Unchecked'}")
    if state:
        #this means that the column should be removed
        #pd.concat([col_storage_pd, copy_column(index)], ignore_index=True)
        remove_column(treev, index)
    else:
        #this means the column should be included
        #move back to treeview
        return




# Create checkboxes and associate each with an IntVar to track state

'''for i in range(5):
    var = tk.IntVar()
    check_vars.append(var)  # Store the IntVar for each checkbox
    
    # Create the checkbox and pack it into the window
    checkbox = tk.Checkbutton(r, text=f"Checkbox {i+1}", variable=var, command=lambda index=i: checkbox_state_changed(index))
    checkbox.pack()
    '''



def load_file():
    global hasLoadedFileData
    global check_vars
    global ButtonsN
    global wb
    file_path = open_file()
    if file_path:
        if hasLoadedFileData:
            if show_option_dialog("Save current Data!", "Would you like to first save the data you are working with before you overwrite it with new data?"):
                if file_save_as():
                    firstTimeOver = False
                else:
                    return
            # Remove all data in the Treeview
            for item in treev.get_children():
                treev.delete(item)
        hasLoadedFileData = True #this prevents us from loading two different data sets with different columns.
        
        #check if files are CSV then convert to excel:
        file_extension = os.path.splitext(file_path)[1]
        if file_extension == '.CSV':
            file_path = convert_csv_to_excel(file_path)
        
        # Load the workbook
        print(file_path)
        wb = openpyxl.load_workbook(file_path)

        # Select the active sheet
        sheet = wb.active
        
        firstTimeOver = True
        #columnsInTree = ()

        # Read and print the data
        for row in sheet.iter_rows(min_row=1, values_only=True):
            # After the first run through row 1 we will no longer be adding columns
            if firstTimeOver is True:
                # Defining number of columns
                treev['columns'] = (row)
                ButtonsN = []
                check_vars = []
                for i,header in enumerate(row):
                    var = IntVar()
                    
                    check_vars.append(var)
                    tempButt = Checkbutton(label_frame, 
                        text = header, 
                        variable = var, 
                        onvalue = 1, 
                        offvalue = 0, 
                        height = 2, 
                        width = 10,
                        command=lambda index=i: checkbox_state_changed(index))
                    ButtonsN.append(tempButt)
                    # Set the trace to call checkbox_state_changed whenever the variable changes
                    #var.trace_add("write", lambda var, index=i, *args: checkbox_state_changed(index))
                for butt in ButtonsN:
                    butt.pack(side='top', anchor='center')
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




#Adding a menu Bar:

def donothing():
    return

menubar = Menu(r)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="New", command=donothing)
filemenu.add_command(label="Open", command=load_file)
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