from tkinter import Tk, StringVar, Button, Text, Entry, LabelFrame, Radiobutton, IntVar, Toplevel
import xlrd

wb = xlrd.open_workbook('ynm3000.xls')
ws0 = wb.sheet_by_index(0)

win = Tk()
win.title('exam')
sw = win.winfo_screenwidth()
sh = win.winfo_screenheight()
ww = 250
wh = 180
x = (sw - ww) / 8
y = (sh - wh) / 2
win.geometry("%dx%d+%d+%d" % (ww, wh, x, y))

flag0 = 0
r = range(1, 3037)


def Copy(event=0):
    global num
    global numlasat
    numlasat = num
    word0 = t.selection_get()
    t.delete('1.0', 'end')
    t1.delete('1.0', 'end')
    for i in range(2, ws0.nrows + 1):
        if ws0.cell(i - 1, 4).value == word0:
            num = i
            show()
        if word0 in ws0.cell(i - 1, 7).value:
            num = i
            var.set(word0)
            show()
    t.insert('end', '\n')
    t1.insert('end', '\n\n')
    for i in range(2, ws0.nrows + 1):
        if word0 in ws0.cell(i - 1, 8).value:
            num = i
            var.set(word0)
            show()
    # win.clipboard_clear()
    # win.clipboard_append(t.selection_get())


frame1 = LabelFrame(win)
frame1.pack(side='left')
# b2 = Button(frame1, text='Copy', font=(
#     'Times New Roman', 12), command=Copy)
# b2.pack()
t = Text(win, font=('Times New Roman', 18))
t.pack(side='left')
t.bind('<Triple-Button-1>', Copy)
t.bind('<Shift_L>', Copy)


win2 = Toplevel(win)
win2.title('word')
ww = 500
wh = 400
x = (sw - ww) / 2
y = (sh - wh) / 2
win2.geometry("%dx%d+%d+%d" % (ww, wh, x, y))

var = StringVar()
num = 0
flag = 0
numlasat = 0
nnum = 0


def show():
    t.insert('end', ws0.cell(num - 1, 4).value)
    t.insert('end', '\n')
    t1.insert('end', ws0.cell(num - 1, 4).value)
    t1.insert('end', '\n')
    t1.insert('end', ws0.cell(num - 1, 5).value)
    if ws0.cell(num - 1, 3).value != '':
        t1.insert('end', '\n')
        t1.insert('end', '派生词：\n')
        t1.insert('end', ws0.cell(num - 1, 3).value)
    if ws0.cell(num - 1, 7).value != '':
        t1.insert('end', '\n')
        t1.insert('end', '近义词：\n')
        t1.insert('end', ws0.cell(num - 1, 7).value)
    if ws0.cell(num - 1, 8).value != '':
        t1.insert('end', '\n')
        t1.insert('end', '反义词：\n')
        t1.insert('end', ws0.cell(num - 1, 8).value)
    t1.insert('end', '\n\n')


def Hide(event=0):
    global flag
    if flag == 0:
        win2.geometry('230x80')
        flag = 1
    else:
        win2.geometry('500x400')
        flag = 0


def Find2(event):
    global num
    global numlasat
    numlasat = num
    word0 = e.get()
    # find_word(word0, ws0)
    t.delete('1.0', 'end')
    t1.delete('1.0', 'end')
    for i in range(2, ws0.nrows + 1):
        if ws0.cell(i - 1, 4).value == word0:
            num = i
            show()
        if word0 in ws0.cell(i - 1, 7).value:
            num = i
            var.set(word0)
            show()
    t.insert('end', '\n')
    t1.insert('end', '\n\n')
    for i in range(2, ws0.nrows + 1):
        if word0 in ws0.cell(i - 1, 8).value:
            num = i
            var.set(word0)
            show()


def Copy2(event=0):
    global num
    global numlasat
    numlasat = num
    word0 = t1.selection_get()
    t.delete('1.0', 'end')
    t1.delete('1.0', 'end')
    for i in range(2, ws0.nrows + 1):
        if ws0.cell(i - 1, 4).value == word0:
            num = i
            show()
        if word0 in ws0.cell(i - 1, 7).value:
            num = i
            var.set(word0)
            show()
    t.insert('end', '\n')
    t1.insert('end', '\n\n')
    for i in range(2, ws0.nrows + 1):
        if word0 in ws0.cell(i - 1, 8).value:
            num = i
            var.set(word0)
            show()


group0 = LabelFrame(win2)
group0.pack(expand='yes', fill='both')

frame1 = LabelFrame(group0)
frame1.pack(side='left')
frame2 = LabelFrame(group0)
frame2.pack(side='right')

b = Button(frame1, text='Hide', font=(
    'Times New Roman', 12), command=Hide)
# b.pack(expand='yes', fill='both')
b.pack(side='left')
span = 2
# b.grid(row=1, column=1)
# b.place(x=30, y=0)

e = Entry(win2, width=10, textvariable=var, bd=2, font=('Times New Roman', 20))
e.pack(expand='yes', fill='both')

t1 = Text(win2, font=('Times New Roman', 16))
t1.pack(side='left')
t1.bind('<Left>', Hide)
t1.bind('<Shift_L>', Copy2)
e.bind('<Return>', Find2)


win.mainloop()
