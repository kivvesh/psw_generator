from tkinter import ttk, Tk
from tkinter import *
from tkinter import messagebox
import random
import xlwt


def generate_pswd():
    alp = 'abcdefghijklmnopqrstuvwxyz'
    numbers = '1234567890'
    spec_symb = '._!#$?' 
    try:
        psw = [' ' for i in range(int(var1.get()))]
        mas = [i for i in range(int(var1.get()))]
        dic = {
        'num':int(var2.get()),
        'spec': int(var3.get()),
        'up_reg': int(var4.get()),
        'other': int(var1.get()) - int(var2.get()) - int(var3.get()) -int(var4.get())
        }
    
        mas = [i for i in range(len(psw))]
        for key, value in dic.items():
            while dic[key]>0:
                if key =='num' and value>0:
                    zn = random.choice(mas)
                    psw[zn] = random.choice(numbers)
                    mas.remove(zn)
                    dic[key]-=1
                elif key =='spec' and value>0:
                    zn = random.choice(mas)
                    psw[zn] = random.choice(spec_symb)
                    mas.remove(zn)
                    dic[key]-=1
                elif key =='up_reg' and value>0:
                    zn = random.choice(mas)
                    psw[zn] = random.choice(alp).upper()
                    mas.remove(zn)
                    dic[key]-=1
                elif key =='other' and value>0:
                    zn = random.choice(mas)
                    psw[zn] = random.choice(alp)
                    mas.remove(zn)
                    dic[key]-=1
        return (''.join(psw))
    except:
        messagebox.showinfo('ERROR','Данные введены некорректно')
    

def but_1():
    result = generate_pswd()
    var5.set(result)

def but_2():
    book = xlwt.Workbook(encoding="utf-8")


    sheet1 = book.add_sheet("Пароли")
    sheet1.write(0, 1, "Пароли от п/я")
    for i in range(int(var6.get())):
        sheet1.write(i+1, 1, generate_pswd())
    book.save("psw.xls")
                
            
 
root = Tk()
root.title("генератор паролей")
root.geometry("450x250")

var1 = IntVar()
entry1 = ttk.Entry(textvariable=var1)

var2 = IntVar()
entry2 = ttk.Entry(textvariable=var2)

var3 = IntVar()
entry3 = ttk.Entry(textvariable=var3)

var4 = IntVar()
entry4 = ttk.Entry(textvariable=var4)

var5 = StringVar()
entry5 = ttk.Entry(textvariable=var5)

var6 = IntVar()
entry6 = ttk.Entry(textvariable=var6)

label_1 = ttk.Label(text = 'Количество символов')
label_2 = ttk.Label(text = 'Количество цифр')
label_3 = ttk.Label(text = 'Количество СпецСимволов')
label_4 = ttk.Label(text = 'Количество Вверх_регистр')
label_5 = ttk.Label(text = 'Ваш пароль')
label_6 = ttk.Label(text = 'Количество паролей')

entry1.grid(row=0, column=0, sticky='w')
label_1.grid(row=0, column=1, sticky='w')

entry2.grid(row=1, column=0, sticky='w')
label_2.grid(row=1, column=1, sticky='w')

entry3.grid(row=2, column=0, sticky='w')
label_3.grid(row=2, column=1, sticky='w')

entry3.grid(row=3, column=0, sticky='w')
label_3.grid(row=3, column=1, sticky='w')

entry4.grid(row=4, column=0, sticky='w')
label_4.grid(row=4, column=1, sticky='w')

btn = ttk.Button(text="Сгенерировать пароль",command = but_1 )
btn.grid(row = 5, column=0)

entry5.grid(row=6, column=0, sticky='w')
label_5.grid(row=6, column=1, sticky='w')

entry6.grid(row=7, column=0, sticky='w')
label_6.grid(row=7, column=1, sticky='w')

btn1 = ttk.Button(text="Сгенерировать пароли",command = but_2)
btn1.grid(row = 7, column=2)

if __name__ == '__main__':
    root.mainloop()

