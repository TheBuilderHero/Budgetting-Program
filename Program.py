# make sure to pip3 install openpyxl
import openpyxl
# GUI:
import tkinter as tk
import tkinter.filedialog as dia
from tkinter import ttk
# CSV to Excel formatting:
import csv



r = tk.Tk()# Adjust size
#r.geometry("1600x800")
#r.minsize(width=1600,height=800)
r.resizable(width=False, height=False)
r.title('Budgetting App')

# Using treeview widget    
treev = ttk.Treeview(r, selectmode ='browse')

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

listOfBoxes = []

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
def button_click():
    fil = dia.askopenfilename(filetypes=[('Excel', '*.xlsx'),('Excel', '*.xlsm'),('Excel', '*.xlsb'),('Excel', '*.xltx'), ('CSV files', '*.csv')])
    #print("Button clicked!")
    #print("ran test")
    if fil is not None:
        #Lb = tk.Listbox(r)

        # Load the workbook
        wb = openpyxl.load_workbook(fil)

        # Select the active sheet
        sheet = wb.active
        
        firstTimeOver = True
        columnsInTree = ()

        # Read and print the data
        for row in sheet.iter_rows(min_row=1, values_only=True):
            if firstTimeOver is True:
                    # Defining number of columns
                    treev['columns'] = (row)
            for i,col in enumerate(row):
                if firstTimeOver is True:
                    # Defining number of columns
                    #treev['columns'] = tuple(treev['columns']) + tuple(str(i),)
                    listOfBoxes.append(tk.Listbox(r, width=25))
                    # Assigning the width and anchor to  the
                    # respective columns
                    treev.column(str(i))
                    # Assigning the heading names to the 
                    # respective columns
                    treev.heading(column=str(i), text=col)
                    print(col)

                else:
                    listOfBoxes[i].insert(tk.END, col)
            
            treev.insert("", 'end', text ="L1", values =(row))
                #print(col)
            # After the first run through row 1 we will no longer be adding columns
            firstTimeOver = False
        #print(cell_obj.value)
        print(treev["columns"])
        # Defining heading
        treev['show'] = 'headings'
        print(treev['show'])
        print(listOfBoxes)
        #show_object(listOfBoxes,"right","both")
#hide_object(listOfBoxes)

button_click()
#button = tk.Button(r, text='Open File', width=25, command=button_click)
#show_object(button,"bottom","none")


'''
Widgets are added here
'''
r.mainloop()