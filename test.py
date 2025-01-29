import tkinter as tk

# Callback function to handle the checkbox state change
def checkbox_state_changed(index):
    """Callback function that gets triggered when a checkbox state changes."""
    state = check_vars[index].get()  # Get the state of the checkbox
    print(f"Checkbox {index + 1} is {'Checked' if state else 'Unchecked'}")

# Create the main window
root = tk.Tk()

# List to hold the checkboxes' variable references
check_vars = []

# Create checkboxes and associate each with an IntVar to track state
checkboxes = []
for i in range(5):
    var = tk.IntVar()
    check_vars.append(var)  # Store the IntVar for each checkbox
    
    # Create the checkbox and pack it into the window
    checkbox = tk.Checkbutton(root, text=f"Checkbox {i+1}", variable=var, command=lambda index=i: checkbox_state_changed(index))
    checkbox.pack()

# Start the Tkinter main event loop
root.mainloop()
