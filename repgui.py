import tkinter
from tkinter import ttk
import tkcalendar
from ttkwidgets import autocomplete

reproot = tkinter.Tk()
reproot.minsize(width=1150, height=650)
reproot.title('Report Generator')
reproot.focus_set()

labelfromdate = tkinter.Label(reproot, text='From :')
labelfromdate.place(relx=0.02, rely=0.02)
fromdatepicker = tkcalendar.DateEntry(reproot, date_pattern='dd/mm/yyyy')
fromdatepicker.place(relx=0.02, rely=0.05, relwidth=0.133, relheight=0.034)

labeltodate = tkinter.Label(reproot, text='To :')
labeltodate.place(relx=0.173, rely=0.02)
todatepicker = tkcalendar.DateEntry(reproot, date_pattern='dd/mm/yyyy')
todatepicker.place(relx=0.173, rely=0.05, relwidth=0.133, relheight=0.034)

labelitems = tkinter.Label(reproot, text='Item :')
labelitems.place(relx=0.326, rely=0.02)
itementry = autocomplete.AutocompleteEntry(reproot, completevalues=[])
itementry.place(relx=0.326, rely=0.05, relwidth=0.133, relheight=0.034)

labelcategory = tkinter.Label(reproot, text='Category :')
labelcategory.place(relx=0.509, rely=0.02)
categoryentry = autocomplete.AutocompleteEntry(reproot, completevalues=[])
categoryentry.place(relx=0.509, rely=0.05, relwidth=0.133, relheight=0.034)

labelmode = tkinter.Label(reproot, text='Mode of Payment :')
labelmode.place(relx=0.692, rely=0.02)
modeentry = autocomplete.AutocompleteCombobox(reproot, completevalues=[])
modeentry.place(relx=0.692, rely=0.05, relwidth=0.133, relheight=0.034)

buttongenerate = tkinter.Button(reproot, text='Generate', relief='groove')
buttongenerate.place(relx=0.86, rely=0.03, relwidth=0.07, relheight=0.055)

buttonopen = tkinter.Button(reproot, text='open', relief='groove')
buttonopen.place(relx=0.87, rely=0.25, relwidth=0.07, relheight=0.055)

buttonreset = tkinter.Button(reproot, text='reset', relief='groove')
buttonreset.place(relx=0.87, rely=0.40, relwidth=0.07, relheight=0.055)

buttonexit = tkinter.Button(reproot, text='exit', relief='groove')
buttonexit.place(relx=0.87, rely=0.55, relwidth=0.07, relheight=0.055)

buttonitemlist = tkinter.Button(reproot, text='..')
buttonitemlist.place(relx=0.469, rely=0.05, relwidth=0.02, relheight=0.034)

buttoncategorylist = tkinter.Button(reproot, text='..')
buttoncategorylist.place(relx=0.652, rely=0.05, relwidth=0.02, relheight=0.034)

scroll = tkinter.Scrollbar(reproot, orient=tkinter.VERTICAL)
scroll.place(relx=0.825, rely=0.13, relheight=0.79)

tree = ttk.Treeview(reproot, show=['headings'], selectmode='browse', yscrollcommand=scroll.set)
tree.place(relx=0.02, rely=0.13, relwidth=0.805, relheight=0.79)

scroll.configure(command=tree.yview)

tree['columns'] = ('date', 'item', 'cat', 'mode', 'unit', 'qty', 'rate', 'da', 'dr', 'net')
reproot.update()
tree.update()
tree.column('date', anchor=tkinter.CENTER, width=int(tree.winfo_width() / 10))
tree.column('item', anchor=tkinter.W, width=int(tree.winfo_width() / 8))
tree.column('cat', anchor=tkinter.W, width=int(tree.winfo_width() / 7.5))
tree.column('mode', anchor=tkinter.W, width=int(tree.winfo_width() / 9))
tree.column('unit', anchor=tkinter.W, width=int(tree.winfo_width() / 12))
tree.column('qty', anchor=tkinter.E, width=int(tree.winfo_width() / 11))
tree.column('rate', anchor=tkinter.E, width=int(tree.winfo_width() / 11))
tree.column('da', anchor=tkinter.E, width=int(tree.winfo_width() / 13))
tree.column('dr', anchor=tkinter.E, width=int(tree.winfo_width() / 13))
tree.column('net', anchor=tkinter.E, width=int(tree.winfo_width() / 8.5))
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

reproot.mainloop()
