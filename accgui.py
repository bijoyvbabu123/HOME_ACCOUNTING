import tkinter
import tkcalendar
from tkinter import ttk
from ttkwidgets import autocomplete

accroot = tkinter.Tk()
accroot.minsize(width=1150, height=660)
accroot.title('Home Accounting')

datepicker = tkcalendar.DateEntry(accroot, date_pattern='dd/mm/yyyy')
datepicker.place(relx=0.02, rely=0.0256, relwidth=0.15)

labelitem = tkinter.Label(accroot, text='Item')
labelitem.place(relx=0.02, rely=0.0256 * 4)
itementry = autocomplete.AutocompleteEntry(accroot, completevalues=[])
itementry.place(relx=0.02, rely=0.0256 * 5, relwidth=0.15)

labelcategory = tkinter.Label(accroot, text='Category')
labelcategory.place(relx=0.02, rely=0.0256 * 8)
categoryentry = autocomplete.AutocompleteEntry(accroot, completevalues=[])
categoryentry.place(relx=0.02, rely=0.0256 * 9, relwidth=0.15)

labelmode = tkinter.Label(accroot, text='Mode')
labelmode.place(relx=0.02, rely=0.0256 * 12)
modeentry = autocomplete.AutocompleteCombobox(accroot, completevalues=[])
modeentry.place(relx=0.02, rely=0.0256 * 13, relwidth=0.15)

labelunit = tkinter.Label(accroot, text='Unit')
labelunit.place(relx=0.02, rely=0.0256 * 16)
unitentry = autocomplete.AutocompleteCombobox(accroot, completevalues=[])
unitentry.place(relx=0.02, rely=0.0256 * 17, relwidth=0.15)

labelquantity = tkinter.Label(accroot, text='Quantity')
labelquantity.place(relx=0.02, rely=0.0256 * 20)
quantityentry = tkinter.Entry(accroot)
quantityentry.place(relx=0.02, rely=0.0256 * 21, relwidth=0.15)

labelrate = tkinter.Label(accroot, text='Rate')
labelrate.place(relx=0.02, rely=0.0256 * 24)
rateentry = tkinter.Entry(accroot)
rateentry.place(relx=0.02, rely=0.0256 * 25, relwidth=0.15)

labeldisall = tkinter.Label(accroot, text='Dis. Allowded')
labeldisall.place(relx=0.02, rely=0.0256 * 28)
disallentry = tkinter.Entry(accroot)
disallentry.place(relx=0.02, rely=0.0256 * 29, relwidth=0.07)

labeldisrec = tkinter.Label(accroot, text='Dis. Received')
labeldisrec.place(relx=0.02 + 0.07 + 0.01, rely=0.0256 * 28)
disrecentry = tkinter.Entry(accroot)
disrecentry.place(relx=0.02 + 0.07 + 0.01, rely=0.0256 * 29, relwidth=0.07)

labelnet = tkinter.Label(accroot, text='Net Price')
labelnet.place(relx=0.02, rely=0.0256 * 32)
netentry = tkinter.Entry(accroot)
netentry.place(relx=0.02, rely=0.0256 * 33, relwidth=0.15)

buttonaddordone = tkinter.Button(accroot, text='ADD')
buttonaddordone.place(relx=0.045, rely=0.0256 * 36, relheight=0.0256 * 2, relwidth=0.10)

buttonsaveexit = tkinter.Button(accroot, text='SAVE and EXIT')
buttonsaveexit.place(relwidth=0.15, relheight=0.0256 * 3, relx=0.4, rely=0.859)

buttonnosaveexit = tkinter.Button(accroot, text='EXIT without SAVING')
buttonnosaveexit.place(relwidth=0.15, relheight=0.0256 * 3, relx=0.59, rely=0.859)

buttonedit = tkinter.Button(accroot, text='Edit', state=tkinter.DISABLED)
buttonedit.place(relwidth=0.07, relheight=0.0256 * 2, rely=0.75, relx=0.43)

buttondelete = tkinter.Button(accroot, text='Delete', state=tkinter.DISABLED)
buttondelete.place(relwidth=0.07, relheight=0.0256 * 2, rely=0.75, relx=0.53)

buttonclearselection = tkinter.Button(accroot, text='Clear Selection', state=tkinter.DISABLED)
buttonclearselection.place(relwidth=0.07, relheight=0.0256 * 2, rely=0.75, relx=0.63)

buttonitemlist = tkinter.Button(accroot, text='..')
buttonitemlist.place(relx=0.18, rely=0.0256 * 5, height=21, relwidth=0.02)

buttoncategorylist = tkinter.Button(accroot, text='..')
buttoncategorylist.place(relx=0.18, rely=0.0256 * 9, height=21, relwidth=0.02)

treescroll = tkinter.Scrollbar(accroot, orient=tkinter.VERTICAL)
treescroll.place(relheight=0.7, relx=0.97, rely=0.03)

tree = ttk.Treeview(accroot, show=['headings'], yscrollcommand=treescroll.set)
tree.place(relwidth=0.75, relheight=0.7, relx=0.22, rely=0.03)

treescroll.config(command=tree.yview)

tree['columns'] = ('date', 'item', 'cat', 'mode', 'unit', 'qty', 'rate', 'da', 'dr', 'net')

accroot.update()
tree.update()
tree.column('date', anchor=tkinter.W, width=int(tree.winfo_width()/10))
tree.column('item', anchor=tkinter.W, width=int(tree.winfo_width()/8))
tree.column('cat', anchor=tkinter.W, width=int(tree.winfo_width()/7.5))
tree.column('mode', anchor=tkinter.W, width=int(tree.winfo_width()/9))
tree.column('unit', anchor=tkinter.W, width=int(tree.winfo_width()/12))
tree.column('qty', anchor=tkinter.W, width=int(tree.winfo_width()/11))
tree.column('rate', anchor=tkinter.W, width=int(tree.winfo_width()/11))
tree.column('da', anchor=tkinter.W, width=int(tree.winfo_width()/13))
tree.column('dr', anchor=tkinter.W, width=int(tree.winfo_width()/13))
tree.column('net', anchor=tkinter.W, width=int(tree.winfo_width()/8.5))
tree.heading('date', text='Date')  # , anchor=tkinter.W
tree.heading('item', text='Item')
tree.heading('cat', text='Category')
tree.heading('mode', text='Mode')
tree.heading('unit', text='Unit')
tree.heading('qty', text='Quantity')
tree.heading('rate', text='Rate')
tree.heading('da', text='Dis. All')
tree.heading('dr', text='Dis. Rec')
tree.heading('net', text='Net Price')

print(accroot.cget('bg'))

accroot.mainloop()
