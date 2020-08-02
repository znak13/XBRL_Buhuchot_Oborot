from tkinter import *
from tkinter.ttk import Combobox
from tkinter.ttk import Radiobutton
from tkinter.ttk import Style

# Коды цветов в шестнадцатеричной кодировке:
# #ff0000 – красный,
# #ff7d00 – оранжевый,
# #ffff00 – желтый,
# #00ff00 – зеленый,
# #007dff – голубой,
# #0000ff – синий,
# #7d00ff – фиолетовый.

def sel():

    global period
    selection0 = "Выбран    -  "
    label0.config(text=selection0, fg='black', font=("Comic Sans MS", 12, "bold"))
    selection1 = periods[var.get()]
    label1.config(text=selection1, fg='red', font=("Comic Sans MS", 12, "bold"))
    label_year_2.config(text=combo.get(), fg='red', font=("Comic Sans MS", 12, "bold"))
    period = var.get()

def combo_fun(event):
    global year
    year = combo.get()
    label_year_2.config(text=combo.get(), fg='red', font=("Comic Sans MS", 12, "bold"))

root = Tk()
root.title("Выбор параметров")
root_w = 600
root_h = 300
root.geometry(f'{root_w}x{root_h}')

row_begin = 0
label_year_1_row = row_begin
label_periods_row = label_year_1_row + 2
R1_row = label_periods_row + 1
label0_row = R1_row + 6
button_OK_row = label0_row + 2

# Меню выбора Года
label_year_1 = Label(root)
label_year_1.config(text="Год: ", fg='black')
label_year_1.grid(column=1, row=label_year_1_row, sticky=W)

combo = Combobox(root, width = 5, textvariable = 3)
combo['values'] = (2020, 2021, 2022, 2023, 2024)
combo.current(0)  # вариант по умолчанию
# combo.grid(column=1, row=label_year_1_row + 1, sticky=W)
combo.place(x=30, y=1)
# combo_N = combo.get()
# combo.pack()
year = combo.get()
combo.bind("<<ComboboxSelected>>", combo_fun)

# Меню выбора Периода
periods = {0: 'Год',
           1: '1-ый квартал',
           2: '2-ой квартал',
           3: '3-ий квартал'}

label_periods = Label(root)
label_periods.config(text="Отчетные периоды: ", fg='black')
label_periods.grid(column=1, row=label_periods_row, sticky=W)

# выбора периода
var = IntVar()
var.set(1)  # значение по умолчанию
period = var.get()
R1 = Radiobutton(root, text=periods[1], variable=var, value=1, command=sel)
R1.grid(column=1, row=R1_row, sticky=W)
R2 = Radiobutton(root, text=periods[2], variable=var, value=2, command=sel)
R2.grid(column=1, row=R1_row+1, sticky=W)
R3 = Radiobutton(root, text=periods[3], variable=var, value=3, command=sel)
R3.grid(column=1, row=R1_row+2, sticky=W)
R4 = Radiobutton(root, text=periods[0], variable=var, value=0, command=sel)
R4.grid(column=1, row=R1_row+3, sticky=W)

# Метка выбора периода
label0 = Label(root)
label0.grid(column=1, row=label0_row, sticky=W)
# Метка выбранного периода
label1 = Label(root)
label1.grid(column=2, row=label0_row, sticky=W)
# Метка выбранного Года
label_year_2 = Label(root)
label_year_2.grid(column=3, row=label0_row, sticky=W)

# Печатаем значения, установленные по умолчанию
combo_fun('event')
sel()

# Кнопка "ОК"
button_OK = Button(root, text="OK", padx="15", pady="2", font="15", command=root.destroy)
button_OK.grid(column=1, row=button_OK_row, sticky=W)

"""
# Находим размер окна
root.update_idletasks()
root_size = root.geometry()
root_size = root_size.split('+')
root_size = root_size[0].split('x')
root_w = int(root_size[0])
root_h = int(root_size[1])

w = root.winfo_screenwidth() # ширина экрана
h = root.winfo_screenheight() # высота экрана
w = w//2 # середина экрана
h = h//2
w = w - root_w//2 # смещение от середины экрана
h = h - root_h//2
# Координаты окна
# root.geometry(f'{root_w}x{root_h}+{w}+{h}')
root.geometry(f'+{w}+{h}')
"""

root.mainloop()

print(year, period)
