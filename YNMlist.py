# -*- coding: utf-8 -*-

"""
Module implementing MainWindow.
"""

from PyQt5.QtCore import pyqtSlot, Qt, QEvent
from PyQt5.QtGui import QKeySequence, QTextCursor
from PyQt5.QtWidgets import QMainWindow, QApplication, QShortcut

from Ui_YNMlist import Ui_MainWindow


class MainWindow(QMainWindow, Ui_MainWindow):
    """
    Class documentation goes here.
    """

    def __init__(self, parent=None):
        """
        Constructor

        @param parent reference to the parent widget
        @type QWidget
        """
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)
        self.winsize = self.size()
        self.flag0 = False
        self.flagf = 0
        self.flagf2 = 0
        self.r0 = range(1, 3037)
        self.r1 = [x for x in self.r0 if ws0.cell(x, 0).value == 1]
        self.r2 = [x for x in self.r0 if ws0.cell(x, 0).value == 2]
        self.r3 = [x for x in self.r0 if ws0.cell(x, 0).value == 3]
        self.r13 = list(set(self.r1 + self.r3))
        self.r23 = list(set(self.r2 + self.r3))
        self.r12 = list(set(self.r1 + self.r2))
        self.rexam = self.r0
        self.num = 0
        self.flag = 0
        self.numlasat = 0
        self.nnum = 0
        self.easy = self.checkBox.isChecked()
        self.medium = self.checkBox_2.isChecked()
        self.hard = self.checkBox_3.isChecked()
        self.sequ = self.radioButton.isChecked()
        # self.setWindowFlags(Qt.FramelessWindowHint)
        # self.lineEdit.installEventFilter(self)
        # self.justDoubleClicked = False
        self.pushButton_6.clicked.connect(self.Exam)
        self.pushButton_3.clicked.connect(self.Rand)
        self.pushButton_5.clicked.connect(self.Next)
        self.pushButton_4.clicked.connect(self.Prev)
        self.pushButton_7.clicked.connect(self.Reset)
        self.pushButton_2.clicked.connect(self.Last)
        self.pushButton_8.clicked.connect(self.Copy)
        self.pushButton.clicked.connect(self.Hide)

        self.listWidget.itemClicked.connect(self.Copylist)

        QShortcut(QKeySequence("Right"), self).activated.connect(self.Rand)
        QShortcut(QKeySequence("Up"), self).activated.connect(self.Prev)
        QShortcut(QKeySequence("Down"), self).activated.connect(self.Next)
        QShortcut(QKeySequence("space"), self).activated.connect(self.Last)
        QShortcut(QKeySequence("Return"), self).activated.connect(self.Find)
        QShortcut(QKeySequence("E"), self).activated.connect(self.Exam)
        QShortcut(QKeySequence("Z"), self).activated.connect(self.Copy)
        QShortcut(QKeySequence("Left"), self).activated.connect(self.Hide)

    #@pyqtSlot(bool)
    # def on_checkBox_clicked(self, checked):
        """
        Slot documentation goes here.
        
        @param checked DESCRIPTION
        @type bool
        """
        # TODO: not implemented yet
    #    self.lineEdit.setEnabled(checked)

    # def mouseDoubleClickEvent(self, event):
    #     self.justDoubleClicked = True
    #     self.Copy()

    # def eventFilter(self, watched, event):
    #     if watched == self.textBrowser_2: # 只对label1的点击事件进行过滤，重写其行为，其他的事件会被忽略
    #         if event.type() == QEvent.MouseButtonPress: # 这里对鼠标按下事件进行过滤，重写其行为
    #             print('open_workbook')
    #     return QMainWindow.eventFilter(self, watched, event)

    def Hide(self):
        if self.flag == 0:
            self.resize(590, 250)
            self.flag = 1
        else:
            self.resize(self.winsize)
            self.flag = 0

    def Find(self):
        self.numlasat = self.num
        word0 = self.lineEdit.text()
        for i in range(2, ws0.nrows + 1):
            if ws0.cell(i - 1, 4).value == word0:
                self.num = i
                self.showtext()

    def Copy(self):
        self.numlasat = self.num
        #word0 = self.textBrowser_2.textCursor().selectedText()
        if len(word0) == 0:
            word0 = self.textBrowser_3.textCursor().selectedText()
        if len(word0) == 0:
            word0 = self.textBrowser.textCursor().selectedText()
        for i in range(2, ws0.nrows + 1):
            if ws0.cell(i - 1, 4).value == word0:
                self.num = i
                self.showtext()

    def Copylist(self):
        self.numlasat = self.num
        word0 = self.listWidget.currentItem().text()
        for i in range(2, ws0.nrows + 1):
            if ws0.cell(i - 1, 4).value == word0:
                self.num = i
                self.showtext()

    def Exam(self):
        self.easy = self.checkBox.isChecked()
        self.medium = self.checkBox_2.isChecked()
        self.hard = self.checkBox_3.isChecked()
        flag2 = 0
        flagfilter = self.flagf
        n = self.spinBox.value()
        if self.easy and self.medium and (not self.hard):
            self.flagf = 12
            if flagfilter != self.flagf:
                exec('self.r%d = self.rexam' % flagfilter)
                self.rexam = self.r12
                self.flag0 = False
        elif self.easy and self.hard and (not self.medium):
            self.flagf = 13
            if flagfilter != self.flagf:
                exec('self.r%d = self.rexam' % flagfilter)
                self.rexam = self.r13
                self.flag0 = False
        elif self.medium and self.hard and (not self.easy):
            self.flagf = 23
            if flagfilter != self.flagf:
                exec('self.r%d = self.rexam' % flagfilter)
                self.rexam = self.r23
                self.flag0 = False
        elif self.easy and (not self.medium) and (not self.hard):
            self.flagf = 1
            if flagfilter != self.flagf:
                exec('self.r%d = self.rexam' % flagfilter)
                self.rexam = self.r1
                self.flag0 = False
        elif (not self.easy) and self.medium and (not self.hard):
            self.flagf = 2
            if flagfilter != self.flagf:
                exec('self.r%d = self.rexam' % flagfilter)
                self.rexam = self.r2
                self.flag0 = False
        elif (not self.easy) and (not self.medium) and self.hard:
            self.flagf = 3
            if flagfilter != self.flagf:
                exec('self.r%d = self.rexam' % flagfilter)
                self.rexam = self.r3
                self.flag0 = False
        else:
            self.flagf = 0
            if flagfilter != self.flagf:
                exec('self.r%d = self.rexam' % flagfilter)
                self.rexam = self.r0
                self.flag0 = False
        self.flagf2 == 1
        # print('flagfilter',flagfilter)
        # print('flagf',self.flagf)
        if len(self.rexam) < n + 1:
            random.shuffle(self.rexam)
            flag2 = 1
        else:
            l = random.sample(self.rexam, n)
            self.rexam = [x for x in self.rexam if x not in set(l)]

        self.listWidget.clear()
        if flag2:
            if self.flag0:
                self.listWidget.addItem('The End')
            else:
                self.flag0 = True
                if len(self.rexam) > 0:
                    for i in self.rexam:
                        self.listWidget.addItem(ws0.cell(i, 4).value)
                else:
                    self.listWidget.addItem('The End')
                self.rexam = []
        else:
            for i in l:
                self.listWidget.addItem(ws0.cell(i, 4).value)
        # print('flag2',flag2)
        # print('flag0',self.flag0)

    def Rand(self):
        self.numlasat = self.num
        self.easy = self.checkBox.isChecked()
        self.medium = self.checkBox_2.isChecked()
        self.hard = self.checkBox_3.isChecked()
        while True:
            self.num = random.randint(1, ws0.nrows)
            if self.easy and ws0.cell(self.num - 1, 0).value == 1 \
                    or self.medium and ws0.cell(self.num - 1, 0).value == 2 \
                    or self.hard and ws0.cell(self.num - 1, 0).value == 3:
                break
            elif not (self.easy or self.medium or self.hard):
                break
        self.showtext()

    def Next(self):
        self.numlasat = self.num
        self.easy = self.checkBox.isChecked()
        self.medium = self.checkBox_2.isChecked()
        self.hard = self.checkBox_3.isChecked()
        self.sequ = self.radioButton.isChecked()
        if self.sequ:
            self.nnum = ws0.cell(self.num - 1, 12).value
            while True:
                self.nnum += 1
                for i in range(1, ws0.nrows):
                    if ws0.cell(i - 1, 12).value == self.nnum:
                        self.num = i
                if self.easy and ws0.cell(self.num - 1, 0).value == 1 \
                        or self.medium and ws0.cell(self.num - 1, 0).value == 2 \
                        or self.hard and ws0.cell(self.num - 1, 0).value == 3:
                    break
                elif not (self.easy or self.medium or self.hard):
                    break
        else:
            while True:
                self.num += 1
                if self.easy and ws0.cell(self.num - 1, 0).value == 1 \
                        or self.medium and ws0.cell(self.num - 1, 0).value == 2 \
                        or self.hard and ws0.cell(self.num - 1, 0).value == 3:
                    break
                elif not (self.easy or self.medium or self.hard):
                    break
        self.showtext()

    def Prev(self):
        self.numlasat = self.num
        self.easy = self.checkBox.isChecked()
        self.medium = self.checkBox_2.isChecked()
        self.hard = self.checkBox_3.isChecked()
        self.sequ = self.radioButton.isChecked()
        if self.sequ:
            self.nnum = ws0.cell(self.num - 1, 12).value
            while True:
                self.nnum -= 1
                for i in range(1, ws0.nrows):
                    if ws0.cell(i - 1, 12).value == self.nnum:
                        self.num = i
                if self.easy and ws0.cell(self.num - 1, 0).value == 1 \
                        or self.medium and ws0.cell(self.num - 1, 0).value == 2 \
                        or self.hard and ws0.cell(self.num - 1, 0).value == 3:
                    break
                elif not (self.easy or self.medium or self.hard):
                    break
        else:
            while True:
                self.num += 1
                if self.easy and ws0.cell(self.num - 1, 0).value == 1 \
                        or self.medium and ws0.cell(self.num - 1, 0).value == 2 \
                        or self.hard and ws0.cell(self.num - 1, 0).value == 3:
                    break
                elif not (self.easy or self.medium or self.hard):
                    break
        self.showtext()

    def Last(self):
        self.num, self.numlasat = self.numlasat, self.num
        self.showtext()

    def Reset(self):
        self.r0 = range(1, 3037)
        self.r1 = [x for x in self.r0 if ws0.cell(x, 0).value == 1]
        self.r2 = [x for x in self.r0 if ws0.cell(x, 0).value == 2]
        self.r3 = [x for x in self.r0 if ws0.cell(x, 0).value == 3]
        self.r13 = list(set(self.r1 + self.r3))
        self.r23 = list(set(self.r2 + self.r3))
        self.r12 = list(set(self.r1 + self.r2))
        self.flagf = 0
        self.rexam = range(1, 3037)
        self.flag0 = False
        self.flagf2 = 0

    def showtext(self):
        self.lineEdit.setText(ws0.cell(self.num - 1, 4).value)
        self.textBrowser.clear()
        strr = '难度： ' + str(ws0.cell(self.num - 1, 0).value)
        self.textBrowser.append(strr)
        self.textBrowser.append('')
        self.textBrowser.append('考法义：')
        self.textBrowser.append(ws0.cell(self.num - 1, 5).value)
        self.textBrowser.append('')
        self.textBrowser.append('例句：')
        self.textBrowser.append(ws0.cell(self.num - 1, 6).value)
        if ws0.cell(self.num - 1, 2).value != '':
            self.textBrowser.append('')
            self.textBrowser.append('记忆法：')
            self.textBrowser.append(ws0.cell(self.num - 1, 2).value)
        if ws0.cell(self.num - 1, 3).value != '':
            self.textBrowser.append('')
            self.textBrowser.append('派生词：')
            self.textBrowser.append(ws0.cell(self.num - 1, 3).value)
        if ws0.cell(self.num - 1, 7).value != '':
            self.textBrowser_3.setText(ws0.cell(self.num - 1, 7).value)
        if ws0.cell(self.num - 1, 8).value != '':
            self.textBrowser.append('')
            self.textBrowser.append('反义词：')
            self.textBrowser.append(ws0.cell(self.num - 1, 8).value)
        if ws0.cell(self.num - 1, 9).value != '':
            self.textBrowser.append('')
            self.textBrowser.append('填空：')
            self.textBrowser.append(ws0.cell(self.num - 1, 9).value)


if __name__ == "__main__":
    import sys
    import xlrd
    import random
    wb = xlrd.open_workbook('ynm3000.xls')
    ws0 = wb.sheet_by_index(0)

    app = QApplication(sys.argv)
    myWin = MainWindow()
    myWin.show()
    sys.exit(app.exec_())
