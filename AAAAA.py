import tkinter
import os


def account():
    kit.destroy()
    import accwindow


def report():
    kit.destroy()
    import repwindow


kit = tkinter.Tk()
kit.minsize(400, 400)
kit.iconbitmap(os.path.join('files', 'kit.ico'))
kit.title('HOME ACCOUNTING KIT')
kit.resizable(False, False)
kit.config(bg='white smoke')

acc_image = tkinter.PhotoImage(file=os.path.join('files', 'accpng.png'))
rep_image = tkinter.PhotoImage(file=os.path.join('files', 'reppng.png'))

labelacc = tkinter.Label(kit, text='Expenditure Accounting', bg='white smoke', font='TkDefaultFont 10 italic underline',
                         fg='DodgerBlue2')
labelacc.place(relx=0.09, rely=0.55)
labelacc.place_forget()

labelrep = tkinter.Label(kit, text='Report Generator', bg='white smoke', font='TkDefaultFont 10 italic underline',
                         fg='DodgerBlue2')
labelrep.place(relx=0.64, rely=0.55)
labelrep.place_forget()

buttonacc = tkinter.Button(kit, image=acc_image, command=account, borderwidth=2, bg='white smoke', relief='groove')
buttonacc.place(relx=0.1, rely=0.2)
buttonacc.bind("<Enter>", lambda event=None: labelacc.place(relx=0.09, rely=0.55))
buttonacc.bind("<Leave>", lambda event=None: labelacc.place_forget())

buttonrep = tkinter.Button(kit, image=rep_image, command=report, borderwidth=2, bg='white smoke', relief='groove')
buttonrep.place(relx=0.6, rely=0.2)
buttonrep.bind("<Enter>", lambda event=None: labelrep.place(relx=0.64, rely=0.55))
buttonrep.bind("<Leave>", lambda event=None: labelrep.place_forget())


kit.mainloop()