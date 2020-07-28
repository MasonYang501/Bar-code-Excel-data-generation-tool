# -*- coding: utf-8 -*-


import xlrd
import xlwt 

from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(611, 555)
        self.OpTips_Bpx = QtWidgets.QGroupBox(Form)
        self.OpTips_Bpx.setGeometry(QtCore.QRect(30, 210, 551, 271))
        self.OpTips_Bpx.setObjectName("OpTips_Bpx")
        self.textEdit = QtWidgets.QTextEdit(self.OpTips_Bpx)
        self.textEdit.setGeometry(QtCore.QRect(20, 30, 511, 221))
        self.textEdit.setObjectName("textEdit")
        self.LabelBtn_4 = QtWidgets.QPushButton(Form)
        self.LabelBtn_4.setGeometry(QtCore.QRect(320, 90, 91, 28))
        self.LabelBtn_4.setObjectName("LabelBtn_4")
        self.textNum_1 = QtWidgets.QTextEdit(Form)
        self.textNum_1.setGeometry(QtCore.QRect(430, 90, 151, 31))
        self.textNum_1.setObjectName("textNum_1")
        self.LabelBtn_6 = QtWidgets.QPushButton(Form)
        self.LabelBtn_6.setGeometry(QtCore.QRect(320, 150, 91, 28))
        self.LabelBtn_6.setObjectName("LabelBtn_6")
        self.textNum_2 = QtWidgets.QTextEdit(Form)
        self.textNum_2.setGeometry(QtCore.QRect(430, 150, 151, 31))
        self.textNum_2.setObjectName("textNum_2")
        self.LabelBtn_2 = QtWidgets.QPushButton(Form)
        self.LabelBtn_2.setGeometry(QtCore.QRect(320, 30, 91, 28))
        self.LabelBtn_2.setObjectName("LabelBtn_2")
        self.textEdit_3 = QtWidgets.QTextEdit(Form)
        self.textEdit_3.setGeometry(QtCore.QRect(430, 30, 151, 31))
        self.textEdit_3.setObjectName("textEdit_3")
        self.LabelBtn_3 = QtWidgets.QPushButton(Form)
        self.LabelBtn_3.setGeometry(QtCore.QRect(30, 90, 91, 28))
        self.LabelBtn_3.setObjectName("LabelBtn_3")
        self.textEdit_1 = QtWidgets.QTextEdit(Form)
        self.textEdit_1.setGeometry(QtCore.QRect(140, 90, 151, 31))
        self.textEdit_1.setObjectName("textEdit_1")
        self.textEdit_2 = QtWidgets.QTextEdit(Form)
        self.textEdit_2.setGeometry(QtCore.QRect(140, 150, 151, 31))
        self.textEdit_2.setObjectName("textEdit_2")
        self.LabelBtn_5 = QtWidgets.QPushButton(Form)
        self.LabelBtn_5.setGeometry(QtCore.QRect(30, 150, 91, 28))
        self.LabelBtn_5.setObjectName("LabelBtn_5")
        self.textEdit_4 = QtWidgets.QTextEdit(Form)
        self.textEdit_4.setGeometry(QtCore.QRect(140, 30, 151, 31))
        self.textEdit_4.setObjectName("textEdit_4")
        self.LabelBtn_1 = QtWidgets.QPushButton(Form)
        self.LabelBtn_1.setGeometry(QtCore.QRect(30, 30, 91, 28))
        self.LabelBtn_1.setObjectName("LabelBtn_1")
        self.GeneBtn = QtWidgets.QPushButton(Form)
        self.GeneBtn.setGeometry(QtCore.QRect(260, 490, 91, 28))
        font = QtGui.QFont()
        font.setPointSize(9)
        self.GeneBtn.setFont(font)
        self.GeneBtn.setObjectName("GeneBtn")
        self.action = QtWidgets.QAction(Form)
        self.action.setCheckable(True)
        self.action.setObjectName("action")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "条码编号生成工具"))
        self.OpTips_Bpx.setTitle(_translate("Form", "操作提示"))
        self.LabelBtn_4.setText(_translate("Form", "开始编号"))
        self.LabelBtn_6.setText(_translate("Form", "结束编号"))
        self.LabelBtn_2.setText(_translate("Form", "可变位数"))
        self.LabelBtn_3.setText(_translate("Form", "编号前缀"))
        self.LabelBtn_5.setText(_translate("Form", "编号后缀"))
        self.LabelBtn_1.setText(_translate("Form", "文件名称"))
        self.GeneBtn.setText(_translate("Form", "生成"))
        self.action.setText(_translate("Form", "打开"))
        
        self.GeneBtn.clicked.connect(self.btnPress1_clicked)
        
    def btnPress1_clicked(self):


        name = self.textEdit_4.toPlainText()         #获取excel名 
        
        prefix = self.textEdit_1.toPlainText()       # 前缀编号
        postfix = self.textEdit_2.toPlainText()      # 后缀编号
        
        varnum = int(self.textEdit_3.toPlainText())  # 可变位数
        
        startnum = int(self.textNum_1.toPlainText()) # 开始编号
       
        endnum = int(self.textNum_2.toPlainText())   # 结束编号
        
        num = endnum - startnum                      # 编号个数-1
       
        # 创建一个工作表
        sheet = xlwt.Workbook()  

        # 创建工作表
        new_sheet = sheet.add_sheet('sheet1', cell_overwrite_ok=True)  
        
        new_sheet.write(0,0,"barcode")
        
        for i in range(1,num+2):
            
            num  = str(int(startnum) + i - 1).zfill(varnum)
            
            strdata = prefix + num + postfix
            
            print("%s\n" % strdata)
            
            new_sheet.write(i,0,strdata)
            
        sheet.save(name+'.xls')
        
        self.textEdit.setPlainText(name+'.xls'+'创建完成!')
        
            
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    ui.setupUi(Form)
    Form.show()
    sys.exit(app.exec_())
        
        
