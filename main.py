from tkinter import *
from tkinter import ttk
import functions

window = Tk()

window.title('MS Access tool')
window.geometry('500x300+300+100')
wdnlbl = Label(window, text='Microsoft Access copy access tool', font=('Arial',16))
wdnlbl.pack()

namevar = StringVar()
name1 = Entry(window, textvariable=namevar)
sapidvar = IntVar()
sapid1 = Entry(window, textvariable=sapidvar)
copy_idvar = IntVar()
copy_id1 = Entry(window, textvariable=copy_idvar)
namelbl = Label(window, text='Insert user name')
sapidlbl = Label(window, text='Insert user ID')
copylbl = Label(window,text='Insert mirror ID')

tables = ('Usuario', 'UsuarioEquipamentos',
    'UsuarioFerramenta', 'UsuarioGarantia', 'UsuarioPreventiva', 'UsuarioSistemaDeControle')

tablevar = StringVar()
cb = ttk.Combobox(window, values = tables, textvariable=tablevar)
cb.place(x=190, y=250)
cblbl = Label(window, text='Select the base')
cblbl.place(x=190, y=225)

def valuesget():
    name = namevar.get().title()
    sapid = sapidvar.get()
    copy_id = copy_idvar.get()
    table = tablevar.get()
    functions.new_access(sapid,name,copy_id,table)
    quit()

name1.place(x=190, y=100)
sapid1.place(x=190, y=150)
copy_id1.place(x=190, y=200)
namelbl.place(x=190,y=75)
sapidlbl.place(x=190, y=125)
copylbl.place(x=190, y=175)



btn = Button(window, text='Copy access', command=valuesget)
btn.place(x=350, y=200)
window.mainloop()
