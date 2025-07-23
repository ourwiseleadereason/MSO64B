import tkinter as tk

def on_selection_change(*args):
    print("User selected:", selected_option.get())

root = tk.Tk()

# Create a StringVar and trace it
selected_option = tk.StringVar()
selected_option.trace_add('write', on_selection_change)

# Initialize the dropdown
options = ['Option 1', 'Option 2', 'Option 3']
selected_option.set(options[0])  # default selection

dropdown = tk.OptionMenu(root, selected_option, *options)
dropdown.pack()

root.mainloop()