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
from tkinter import StringVar
# CSV to Excel formatting:
import csv
import os # file extension removal
# Menu bar:
from tkinter import Menu
from tkinter import OptionMenu
#TreeView data to Excel:
import pandas as pd

################################
#Program Version:
VERSION_MAJOR = 1
VERSION_MINOR = 0
VERSION_PATCH = 0
################################

excel_save_file_extension = '.xlsx'
hasLoadedFileData = False # this if used to tell if we have data in the TreeView or not for saving it when opening a new file.
col_storage_pd = pd.DataFrame()
wb = None
removed_columns_str = "(REMOVED COLUMN)"
allow_multiple_file_uploads_bool = False
custom_category_options = []

r = tk.Tk()

# Using treeview widget    
treev = ttk.Treeview(r, height=20, selectmode ='extended')

'''
    Setting the Select Mode:
        browse: (Default) Only one item can be selected at a time. Clicking on an item selects it and deselects any previously selected item.
        extended: Multiple items can be selected by holding down the Shift key or Ctrl key (Command key on macOS).
        none: Disables selection entirely.
'''



def open_export_window():

    def clear_entry():
        ent_new.delete(0, tk.END)

    def edit():
        # Get selected item to Edit
        selected_item = categories.selection()[0]
        text = categories.item(selected_item)['text']
        value = categories.item(selected_item)['values'][0]
        new_text = text.replace(str(value), ent_new.get())
        categories.item(selected_item, text=new_text, values=(ent_new.get()))
        clear_entry()

    export_screen = tk.Toplevel(r)
    export_screen.title("Export")
    export_screen.minsize(width=400,height=300)
    # Set window position (x, y)
    export_screen.geometry("+150+150") # window position

    # Add widgets to the new window here
    tk.Label(export_screen, text="These are the export options").pack()
    tk.Button(export_screen, text="Export").pack(side="bottom", anchor="se", padx=10, pady=10)

    categories = ttk.Treeview(export_screen, height=20, selectmode ='browse')
    categories.pack(side="left")

    unique_values = []

    for row in treev.get_children():
        custom_val = treev.item(row)['values'][0]
        print(custom_val)
        if custom_val not in unique_values:
            unique_values.append(custom_val)
            str_1_head = "Category Name: " + custom_val
            head_1 = categories.insert("", 'end', text =str_1_head, values=(custom_val))
            child_l1 = categories.insert(head_1, 'end', text ="Budget Min: $0", values=("0"))
            child_l2 = categories.insert(head_1, 'end', text ="Budget Max: $0", values=("0"))


    # This will create a LabelFrame
    label_export_options_1 = LabelFrame(export_screen, text='Export Options 1', font = "50")
    label_export_options_1.pack(side='top', anchor='n')
    # This will create a LabelFrame
    label_export_sub_1 = LabelFrame(label_export_options_1, text='Export sub 1', font = "20")
    label_export_sub_1.pack(side="left")
    # This will create a LabelFrame
    label_export_sub_2 = LabelFrame(label_export_options_1, text='Export sub 2', font = "20")
    label_export_sub_2.pack(side="right")
    check_but1 = tk.Checkbutton(label_export_sub_1, text="test1")
    check_but2 = tk.Checkbutton(label_export_sub_2, text="test2")
    check_but1.pack()
    check_but2.pack()
    # This will create a LabelFrame
    label_export_options_2 = LabelFrame(export_screen, text='Export Options 2', font = "50")
    label_export_options_2.pack(side='top', anchor='n')
    # This will create a LabelFrame
    label_export_sub_3 = LabelFrame(label_export_options_2, text='Export sub 1', font = "20")
    label_export_sub_3.pack(side="left")
    # This will create a LabelFrame
    label_export_sub_4 = LabelFrame(label_export_options_2, text='Export sub 2', font = "20")
    label_export_sub_4.pack(side="right")
    check_but3 = tk.Checkbutton(label_export_sub_3, text="test1")
    check_but4 = tk.Checkbutton(label_export_sub_4, text="test2")
    check_but3.pack()
    check_but4.pack()

    # This will create a LabelFrame
    label_export_options_3 = LabelFrame(export_screen, text='Export Options 3', font = "50")
    label_export_options_3.pack(side='top', anchor='n')
    # This will create a LabelFrame
    label_export_sub_5 = LabelFrame(label_export_options_3, text='Export sub 1', font = "20")
    label_export_sub_5.pack(side="left")
    # This will create a LabelFrame
    label_export_sub_6 = LabelFrame(label_export_options_3, text='Export sub 2', font = "20")
    label_export_sub_6.pack(side="right")
    ent_new = tk.Entry(label_export_options_3)
    but1 = tk.Button(label_export_options_3, text="Change", command=edit)
    ent_new.pack()
    but1.pack()

# Adjust size
# I want it resizable so that the scrolling and the data boxes are not messed up by it.
r.minsize(width=600,height=10)
# Set window position (x, y)
r.geometry("+100+100") # window position
r.resizable(width=False, height=False)
r.title('Budgetting App')

# This will create a LabelFrame
label_frame = LabelFrame(r, text='Columns To Remove When Saved', font = "50")

# This will create a LabelFrame
label_category_selection = LabelFrame(r, text='Custom Categories', font = "50")
# This will create a LabelFrame
label_category_selection_add = LabelFrame(label_category_selection, text='Add Custom Category', font = "20")
# This will create a LabelFrame
label_category_selection_select = LabelFrame(label_category_selection, text='Set Row Category', font = "20")

ButtonsN = []

# List to hold the checkboxes' variable references
check_vars = []



def get_selected_tree_row_add_category():
    selected_items = treev.selection()

    for item in selected_items:
        values_in_row = treev.item(item).get("values")
        '''new_values = []
        item_1 = True
        for values in values_in_row:
            if item_1:
                new_values.append(dropdown.current())
            else:
                new_values.append(values)'''
        
        treev.set(item=item,column=0,value=dropdown["values"][dropdown.current()])
        print(treev.item(item).get("values")[0])



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


label_frame.pack(side='top', anchor='n', padx=20)
label_category_selection.pack(side='bottom', anchor='s', padx=20)
label_category_selection_add.pack(side='left')
label_category_selection_select.pack(side='right')

#adding stuff to the custom categories
entry = ttk.Entry(label_category_selection_add, width = 20)
entry.pack()


def clear_new_category_entry():
        entry.delete(0, tk.END)

def add_new_category():
    message_str = "Would you like to add the new Category \"" + entry.get() + "\" to your column category options?"
    if messagebox.askokcancel("Add New Category", message_str):
        values = list(dropdown["values"])
        dropdown["values"] = values + [entry.get()]
        clear_new_category_entry()


entry_b = ttk.Button(label_category_selection_add, text='Create New Label', width = 20, command=add_new_category)
entry_b.pack()
dropdown = ttk.Combobox(label_category_selection_select, width = 20, state="readonly")
dropdown.pack()
dropdown_add = ttk.Button(label_category_selection_select, text='Label Selected Row', width = 20, command=get_selected_tree_row_add_category)
dropdown_add.pack()
 
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
    

#function to hide and reveal columns
def add_back_column(tree, index):
    global removed_columns_str # this was a string but now is just a heading
    # Remove the second column
    tree.column(treev['columns'][index], width=275, stretch=True)
    tree.heading(treev['columns'][index], text=treev['columns'][index])

#function to hide and reveal columns
def remove_column(tree, index):
    global removed_columns_str
    # Remove the second column
    tree.column(treev['columns'][index], width=0, stretch=False)
    tree.heading(treev['columns'][index], text=removed_columns_str)
    
# Callback function to handle the checkbox state change
def checkbox_state_changed(index):
    #global col_storage_pd
    global treev
    """Callback function that gets triggered when a checkbox state changes."""
    state = check_vars[index].get()  # Get the state of the checkbox
    #print(f"Checkbox {index + 1} is {'Checked' if state else 'Unchecked'}")
    index = index + 1 # this is becuase we are ignoring the custom column
    if state:
        #this means that the column should be removed
        #pd.concat([col_storage_pd, copy_column(index)], ignore_index=True)
        remove_column(treev, index)
    else:
        #this means the column should be included
        #move back to treeview
        add_back_column(treev, index)

def allow_multiple_file_uploads():
    global allow_multiple_file_uploads_bool
    if allow_multiple_file_uploads_bool:
        allow_multiple_file_uploads_bool = False
    else:
        if messagebox.showwarning("DANGEROUS OPTION!", "Are you sure you want to allow for opening muliple file? This can cause issues with columns and currupt data. Only do this if you are sure that the data which is being opened has all the same columns.", type=messagebox.OKCANCEL):
            allow_multiple_file_uploads_bool = True



def load_file():
    global hasLoadedFileData
    global check_vars
    global ButtonsN
    global wb
    file_path = open_file()
    if file_path:
        if not allow_multiple_file_uploads_bool:
            if hasLoadedFileData:
                if show_option_dialog("Save current Data!", "Would you like to first save the data you are working with before you overwrite it with new data?"):
                    if file_save_as():
                        firstTimeOver = False
                    else:
                        return
                # Remove all data in the Treeview
                for item in treev.get_children():
                    treev.delete(item)
        
        #remove all checkboxes:
        print(ButtonsN)
        print(len(ButtonsN))
        for i in range(len(ButtonsN)): # we dont want to use i since it is not the front item thwn we want to 
            #delete first item as we go through the list.
            ButtonsN[0].destroy()
            ButtonsN.pop(0)
                
        hasLoadedFileData = True #this prevents us from loading two different data sets with different columns.
        
        #check if files are CSV then convert to excel:
        file_extension = os.path.splitext(file_path)[1]
        if file_extension == '.CSV':
            file_path = convert_csv_to_excel(file_path)
        
        # Load the workbook
        #print(file_path)
        wb = openpyxl.load_workbook(file_path)

        # Select the active sheet
        sheet = wb.active
        
        firstTimeOver = True
        skip_add_col = -1
        #columnsInTree = ()

        # Read and print the data
        for row in sheet.iter_rows(min_row=1, values_only=True):
            # After the first run through row 1 we will no longer be adding columns
            if firstTimeOver is True:
                # Defining number of columns
                # Find the index of 'Custom Category'
                try:
                    skip_add_col = row.index('Custom Category')
                except ValueError:
                    pass
                if skip_add_col < 0:
                    row = ("Custom Category",) + row
                treev['columns'] = (row)
                check_vars = []
                for i,header in enumerate(row):
                    print(header)
                    print(i , " vs ", True if i == 0 else False)
                    if i == 0: # skip adding the removeal checkbox for custom column
                        continue
                    var = IntVar()
                    
                    check_vars.append(var)
                    tempButt = Checkbutton(label_frame, 
                        text = header, 
                        variable = var, 
                        onvalue = 1, 
                        offvalue = 0, 
                        height = 2, 
                        width = 20,
                        command=lambda index=i-1: checkbox_state_changed(index)) # -1 because we are ignoring the custom column
                    ButtonsN.append(tempButt)
                    # Set the trace to call checkbox_state_changed whenever the variable changes
                    #var.trace_add("write", lambda var, index=i, *args: checkbox_state_changed(index))
                for butt in ButtonsN:
                    butt.pack(side='top', anchor='center')
            else:
                row = ("",) + row
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

    if len(removed_columns_str):
        try:
            df.drop(columns=removed_columns_str, axis=1, inplace=True)
        except KeyError:
            print("NO COLUMNS WERE REMOVED")

    # Save the DataFrame to an Excel file
    df.to_excel(filename.name, index=False)



def attempt_end_program():
    if show_option_dialog("Save and Quit?", "Would you like to save the file before you Exit?"):
        if file_save_as():
            messagebox.showinfo("File Saved","File Saved Sucessfully!")
        else:
            messagebox.showerror("FAILED SAVE","File Did Not Save Sucessfully!")
    if messagebox.askokcancel("Quit?", "Would you like to Quit the Program?"):
        r.destroy()
        

r.protocol("WM_DELETE_WINDOW", attempt_end_program)

def about():
    string_ver = "Current Version: " + str(VERSION_MAJOR) + "." + str(VERSION_MINOR) + "." + str(VERSION_PATCH) + "v"
    messagebox.showinfo("About", string_ver)

def Proper_program_usage():
    string_ex = "The Program Intent\n\n" + "The intentbeing the way that this program functions is to allow the user to load in creditcard data and bank data and easily turn them into a graphical representation that can be viewed and explored.\n\n" + "First thing you should do when using this program is upload your bank and creditcard csv or excel files.\n" + "Then remove whatever columns are useless.\n" + "Then add a custom category to the data rows which you want to classify them all under when exported.\n" + "Finally, export the data and select the type of output you want to see when the export is complete."
    messagebox.showinfo("How To...", string_ex)

#Adding a menu Bar:

def donothing():
    return

menubar = Menu(r)
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="New", command=donothing)
filemenu.add_command(label="Open", command=load_file)
filemenu.add_command(label="Save", command=file_save_as)
filemenu.add_command(label="Export", command=open_export_window)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=attempt_end_program)
menubar.add_cascade(label="File", menu=filemenu)

overrideMenu = Menu(menubar, tearoff=0)
overrideMenu.add_checkbutton(label="Allow Multiple File", command=allow_multiple_file_uploads)
menubar.add_cascade(label="OverRide", menu=overrideMenu)

helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Help Index", command=donothing)
helpmenu.add_command(label="About...", command=about)
menubar.add_cascade(label="Help", menu=helpmenu)

r.config(menu=menubar)


'''
Widgets are added here
'''

r.mainloop()