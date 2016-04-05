import sys
from PyQt4 import QtGui
from PyQt4 import QtCore
from PyQt4.QtGui import QMessageBox
import zipfile
import os.path
import re
import tempfile
import shutil


class AutoContract(QtGui.QWidget):
    def __init__(self, parent=None):
        QtGui.QWidget.__init__(self, parent)

        self.setWindowTitle('AutoContract')

        self.wtf = []
        self.lines = []
        self.dirpath = tempfile.mkdtemp()

        chooseLbl = QtGui.QLabel('Выберите файл (*.docx):')
        self.chooseBtn = QtGui.QPushButton('Выбрать', self)

        self.grid = QtGui.QGridLayout()
        self.grid.setSpacing(10)

        self.grid.addWidget(chooseLbl, 0, 1)
        self.grid.addWidget(self.chooseBtn, 0, 2)

        self.connect(self.chooseBtn, QtCore.SIGNAL('clicked()'), self.get_fname)

        self.setLayout(self.grid)
        self.resize(500, 80)


    def get_fname(self):
        openfile = QtGui.QFileDialog.getOpenFileName(self, 'Выберите файл', '', '	Doc (*.docx)')
        
        if os.path.splitext(openfile)[1] == '.docx':

            
            z = zipfile.ZipFile(openfile)
            z.extractall(self.dirpath)
            
            doc = open(self.dirpath + '/word/document.xml')
            self.wtf = re.findall(r'##([^##]*)##', doc.read())
            self.lines = list(self.wtf)

            for i in range(0, len(self.wtf)):
                text = re.sub(r'\<[^>]*\>', '', self.wtf[i])
                lbl = QtGui.QLabel(text)

                self.grid.addWidget(lbl, i+1, 1)
                self.wtf[i] = QtGui.QLineEdit()
                self.grid.addWidget(self.wtf[i], i+1, 2)

            saveBtn = QtGui.QPushButton('Сохранить', self)
            self.grid.addWidget(saveBtn, len(self.wtf)+1, 1)
            self.connect(saveBtn, QtCore.SIGNAL('clicked()'), self.check_lines)

            z.close()

        else:
            pass


    def check_lines(self):
        msgBox = QMessageBox()
        err = 1

        #Проверяем, все ли поля заполнены
        for i in range(0, len(self.wtf)):
            if self.wtf[i].text().strip() == '':
                msgBox.setIcon(QMessageBox.Warning)
                msgBox.setText("Пожалуйста, заполните все поля!")
                
                msgBox.exec_()
                err = 0
                break
        
        #Если все поля заполнены, выводим окно с вопросом
        if err:
            reply = QtGui.QMessageBox.question(self, 'Message', "Вы уверены, что хотите сохранить изменения", 
                QtGui.QMessageBox.Yes, QtGui.QMessageBox.No)
            if reply == QtGui.QMessageBox.Yes:
            	self.save()
            else:
            	print('Ни хуя')

        #если есть не заполненые поля, выделяем их красным
        for i in range(0, len(self.wtf)):
            if self.wtf[i].text().strip() == '':
                self.wtf[i].setStyleSheet("border: 1px solid red;")
            else:
                self.wtf[i].setStyleSheet("border: 1px solid green;")


    def save(self):
        doc = open(self.dirpath + '/word/document.xml', 'r+')
        text = doc.read()
        doc.close()

        for i in range(0, len(self.lines)):
            s = '##' + self.lines[i] + '##'
            text = text.replace(s, self.wtf[i].text())

        doc = open(self.dirpath + '/word/document.xml', 'w')
        doc.write(text)
        doc.close()

        self.go_zip(self.dirpath, 'new3')
        #shutil.make_archive("archive.docx",'zip',self.dirpath)


    def go_zip(self,src, dst):
        zf = zipfile.ZipFile("/home/xymfrx/Documents/%s.docx" % (dst), "w", zipfile.ZIP_DEFLATED)
        abs_src = os.path.abspath(src)
        
        for dirname, subdirs, files in os.walk(src):
            for filename in files:
                absname = os.path.abspath(os.path.join(dirname, filename))
                arcname = absname[len(abs_src) + 1:]

                zf.write(absname, arcname)
        
        zf.close()




if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    qb = AutoContract()
    qb.show()
    sys.exit(app.exec_())
