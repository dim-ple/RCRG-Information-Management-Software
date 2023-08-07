
import tkinter as tk
from tkinter import *

from tkinter import ttk

root = tk.Tk()
root.title('Auto Fill Test')
root.geometry("500x300")

select = tk.StringVar()

def update(data):
    new_list = []
    
    my_combo.delete()

    for item in data: 
        new_list.append(item)
    
    my_combo['values'] = new_list

def fillout(event):
    my_combo.delete()
    my_combo.insert(my_combo.get(ACTIVE))

def check(event):
    # Grab what was typed
    typed = my_combo.get()

    # When nothing is typed, the entry list is populated with all toppings from the toppings list
    if typed == '':
        data = my_combo['values']
    else:
        data = []
        for item in data:
            if typed.lower() in item.lower():
                data.append(item)
    
    # update our listbox with selected items
    update(data)

my_label = Label(root, text="Start Typing...", font=("Helvetica", 14), fg="grey")
my_label.pack(pady=20)

my_combo = ttk.Combobox(root, textvariable = select)
my_combo.pack(pady=40)

my_combo["values"] = ["Test1", "Nest2", "Lest3", "Rest3", "Hello", "Gamer"]



update(my_combo["values"])

my_combo.bind("<<ComboboxSelected>>", fillout)

my_combo.bind("<KeyRelease>", check)

root.mainloop()