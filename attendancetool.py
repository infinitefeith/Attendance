import sys
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from attendance_ui import *
from attendance import process_attendance

class Main_ui(QMainWindow, Ui_mainwinow):
    def __init__(self, parent=None):
        super(Main_ui, self).__init__(parent)
        self.setupUi(self)
        self.flag = True

    def openfile(self):
        self.flag = False
        try:
            filename = QFileDialog.getOpenFileName(self, "open file", "./")
            self.filepath.setText(filename[0])
        except Exception(e):
            QMessageBox.information(self, u"警告！", u"文件路径不能有中文！")
            return False
        return True

    def startprocess(self):
        print("processing")
        srcfile = self.filepath.text()
        if srcfile == '':
            QMessageBox.information(self, u"警告！", u"请先打开原始文件数据！")
        else:
            process_attendance(srcfile)
            QMessageBox.information(self, u"恭喜！", u"已处理完成!!")
        return

if __name__ == '__main__':
    app = QApplication(sys.argv)
    m = Main_ui()
    m.show()
    sys.exit(app.exec_())

