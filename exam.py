from tkinter import Tk, StringVar, Button, Text, Entry, LabelFrame, Radiobutton, IntVar, Toplevel
import xlrd
import random

wb = xlrd.open_workbook('ynm3000.xls')
ws0 = wb.sheet_by_index(0)

win = Tk()
win.title('exam')
sw = win.winfo_screenwidth()
sh = win.winfo_screenheight()
ww = 250
wh = 180
x = (sw - ww) / 4
y = (sh - wh) / 2
win.geometry("%dx%d+%d+%d" % (ww, wh, x, y))

var2 = StringVar()
flag0 = 0
r = range(1, 3037)
var2.set(5)


def Reset(event=0):
    global r
    global flag0
    r = range(1, 3037)
    flag0 = 0


def Copy(event=0):
    global num
    global numlasat
    numlasat = num
    word0 = t.selection_get()
    for i in range(2, ws0.nrows + 1):
        if ws0.cell(i - 1, 4).value == word0:
            num = i
            show()
    win.clipboard_clear()
    win.clipboard_append(t.selection_get())
    # print(t.selection_get())


def Exam(event=0):
    global r
    global flag0
    flag2 = 0
    n = var2.get()
    n = int(n)
    if len(r) < n + 1:
        random.shuffle(r)
        flag2 = 1
    else:
        l = random.sample(r, n)
        r = [x for x in r if x not in set(l)]
    t.delete('1.0', 'end')
    if flag2:
        if flag0:
            t.insert('end', 'The End')
        else:
            flag0 = 1
            for i in r:
                t.insert('end', ws0.cell(i, 4).value)
                t.insert('end', '\n')
    else:
        for i in l:
            t.insert('end', ws0.cell(i, 4).value)
            t.insert('end', '\n')


frame1 = LabelFrame(win)
frame1.pack(side='left')
b = Button(frame1, text='Exam', font=(
    'Times New Roman', 12), command=Exam)
# b.pack(side='left')
b.pack()
b1 = Button(frame1, text='Reset', font=(
    'Times New Roman', 12), command=Reset)
# b1.pack(side='left')
b1.pack()
b2 = Button(frame1, text='Copy', font=(
    'Times New Roman', 12), command=Copy)
b2.pack()
e1 = Entry(frame1, width=5, textvariable=var2,
           justify='center', font=('Times New Roman', 12))
# e.pack(side='left')
e1.pack()
t = Text(win, font=('Times New Roman', 18))
t.pack(side='left')
t.bind('<Triple-Button-1>', Copy)
t.bind('<Shift_L>', Copy)
win.bind('<space>', Exam)


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
    var.set(ws0.cell(num - 1, 4).value)
    t1.delete('1.0', 'end')
    t1.insert('end', '难度：')
    t1.insert('end', ws0.cell(num - 1, 0).value)
    t1.insert('end', '\n\n')
    t1.insert('end', '考法义：\n')
    t1.insert('end', ws0.cell(num - 1, 5).value)
    t1.insert('end', '\n\n')
    t1.insert('end', '例句：\n')
    t1.insert('end', ws0.cell(num - 1, 6).value)
    if ws0.cell(num - 1, 2).value != '':
        t1.insert('end', '\n\n')
        t1.insert('end', '记忆法：\n')
        t1.insert('end', ws0.cell(num - 1, 2).value)
    if ws0.cell(num - 1, 3).value != '':
        t1.insert('end', '\n\n')
        t1.insert('end', '派生词：\n')
        t1.insert('end', ws0.cell(num - 1, 3).value)
    if ws0.cell(num - 1, 7).value != '':
        t1.insert('end', '\n\n')
        t1.insert('end', '近义词：\n')
        t1.insert('end', ws0.cell(num - 1, 7).value)
    if ws0.cell(num - 1, 8).value != '':
        t1.insert('end', '\n\n')
        t1.insert('end', '反义词：\n')
        t1.insert('end', ws0.cell(num - 1, 8).value)
    if ws0.cell(num - 1, 9).value != '':
        t1.insert('end', '\n\n')
        t1.insert('end', '填空：\n')
        t1.insert('end', ws0.cell(num - 1, 9).value)


def Next(event=0):
    global num
    global nnum
    global numlasat
    numlasat = num
    if v2.get() == 1:
        nnum = ws0.cell(num - 1, 12).value
        while True:
            nnum += 1
            for i in range(1, ws0.nrows):
                if ws0.cell(i - 1, 12).value == nnum:
                    num = i
            if ws0.cell(num - 1, 0).value == v.get() or v.get() == 0 \
                    or (v.get() == 4 and (ws0.cell(num - 1, 0).value == 2 or ws0.cell(num - 1, 0).value == 3)):
                break
    else:
        while True:
            num += 1
            if ws0.cell(num - 1, 0).value == v.get() or v.get() == 0 \
                    or (v.get() == 4 and (ws0.cell(num - 1, 0).value == 2 or ws0.cell(num - 1, 0).value == 3)):
                break
    show()


def Prev(event=0):
    global num
    global numlasat
    numlasat = num
    if v2.get() == 1:
        nnum = ws0.cell(num - 1, 12).value
        while True:
            nnum -= 1
            for i in range(1, ws0.nrows):
                if ws0.cell(i - 1, 12).value == nnum:
                    num = i
            if ws0.cell(num - 1, 0).value == v.get() or v.get() == 0 \
                    or (v.get() == 4 and (ws0.cell(num - 1, 0).value == 2 or ws0.cell(num - 1, 0).value == 3)):
                break
    else:
        while True:
            num -= 1
            if ws0.cell(num - 1, 0).value == v.get() or v.get() == 0 \
                    or (v.get() == 4 and (ws0.cell(num - 1, 0).value == 2 or ws0.cell(num - 1, 0).value == 3)):
                break
    show()


def Rand(event=0):
    global num
    global numlasat
    numlasat = num
    while True:
        num = random.randint(1, ws0.nrows)
        if ws0.cell(num - 1, 0).value == v.get() or v.get() == 0 \
                or (v.get() == 4 and (ws0.cell(num - 1, 0).value == 2 or ws0.cell(num - 1, 0).value == 3)):
            break
    show()


def Last(event=0):
    global num
    global numlasat
    num, numlasat = numlasat, num
    show()


def Hide(event=0):
    global flag
    if flag == 0:
        win2.geometry('230x80')
        flag = 1
    else:
        win2.geometry('500x400')
        flag = 0


def Find(event):
    global num
    global numlasat
    numlasat = num
    word0 = e.get()
    # find_word(word0, ws0)
    for i in range(2, ws0.nrows + 1):
        if ws0.cell(i - 1, 4).value == word0:
            num = i
            show()


def Copy2(event=0):
    global num
    global numlasat
    numlasat = num
    word0 = t1.selection_get()
    for i in range(2, ws0.nrows + 1):
        if ws0.cell(i - 1, 4).value == word0:
            num = i
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

b2 = Button(frame1, text='Rand', font=(
    'Times New Roman', 12), command=Rand)
b2.pack(padx=0, side='right')

b3 = Button(frame2, text='Next', font=(
    'Times New Roman', 12), command=Next)
b3.pack(padx=0, side='left')

b4 = Button(frame2, text='Prev', font=(
    'Times New Roman', 12), command=Prev)
b4.pack(side='right')

b5 = Button(frame1, text='Last', font=(
    'Times New Roman', 12), command=Last)
b5.pack(side='right')


e = Entry(win2, width=10, textvariable=var, bd=2, font=('Times New Roman', 20))
e.pack(expand='yes', fill='both')

t1 = Text(win2, width=40, font=('Times New Roman', 16))
t1.pack(side='left')
t1.bind('<Up>', Prev)
t1.bind('<Down>', Next)
t1.bind('<Right>', Rand)
t1.bind('<Left>', Hide)
t1.bind('<space>', Last)
t1.bind('<Shift_L>', Copy2)
e.bind('<Return>', Find)

group = LabelFrame(win2, text='难度')
group.pack(side='top', expand='yes', fill='both')

DIFF = [('全部', 0), ('简单', 1), ('可记', 2), ('难记', 3), ('要记', 4)]
v = IntVar()
v.set(0)
for texts, nums in DIFF:
    r1 = Radiobutton(group, text=texts, variable=v, value=nums)
    r1.pack()

group2 = LabelFrame(win2, text='排序')
group2.pack(side='bottom', expand='yes', fill='both')

DIFF2 = [('顺序', 0), ('逆序', 1)]
v2 = IntVar()
v2.set(0)
for texts, nums in DIFF2:
    r2 = Radiobutton(group2, text=texts, variable=v2, value=nums)
    r2.pack()


win.mainloop()
