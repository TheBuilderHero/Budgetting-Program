import tkinter as tk
from tkinter import ttk

root = tk.Tk()

tree = ttk.Treeview(root, columns=("col1", "col2"))
tree.heading("#0", text="Main")  # Changes the heading of the default column
tree.heading("col1", text="Column 1")
tree.heading("col2", text="Column 2")

tree.insert("", tk.END, text="Item 1", values=("value 1", "value 2"))

tree.pack()
root.mainloop()