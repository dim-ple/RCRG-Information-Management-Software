from tkinter import *

from tkinter import ttk


root = Tk()
root.title('Auto Fill Test')
root.geometry("500x300")

def update(data):
    my_list.delete(0, END)

    for item in data:
        my_list.insert(END, item)

def fillout(event):
    my_entry.delete(0, END)
    my_entry.insert(0, my_list.get(ACTIVE))

def check(event):
    # Grab what was typed
    typed = my_entry.get()

    # When nothing is typed, the entry list is populated with all toppings from the toppings list
    if typed == '':
        data = toppings
    else:
        data = []
        for item in toppings:
            if typed.lower() in item.lower():
                data.append(item)
    
    # update our listbox with selected items
    update(data)

    
select = StringVar()

my_label = Label(root, text="Start Typing...", font=("Helvetica", 14), fg="grey")
my_label.pack(pady=20)

my_entry = Entry(root, font=("Helvetica", 20))
my_entry.pack()

my_list = Listbox(root, width=50)
my_list.pack(pady=40)

my_combo = ttk.Combox(root, textvariable = select)

toppings = ["Pepperoni", "Onions", "Cheese", "Ham", "Mushrooms", "Peppers", "Spinach"]

update(toppings)

my_list.bind("<<ListboxSelect>>", fillout)

my_entry.bind("<KeyRelease>", check)

root.mainloop()