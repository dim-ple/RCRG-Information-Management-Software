
import tkinter as tk
from tkinter import *

from tkinter import ttk

root = tk.Tk()
root.title('Auto Fill Test')
root.geometry("500x300")

select = tk.StringVar()

my_label = Label(root, text="Start Typing...", font=("Helvetica", 14), fg="grey")
my_label.pack(pady=20)

my_combo = ttk.Combobox(root, textvariable = select)
my_combo.pack(pady=40)

my_combo['values'] = ['Test1', 'Nest2', 'Lest3', 'Rest3', 'Hello', 'Gamer']

update(my_combo['values'])

root.mainloop()