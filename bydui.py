import sys
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from 报运单.by import *
import qdarkstyle

from 报运单.批量生成报运单最终程序 import gen_byd


class QmyWindow(QMainWindow):

    def __init__(self,parent=None):
        super().__init__(parent)
        self.ui=Ui_MainWindow()
        self.ui.setupUi(self)


    @pyqtSlot()
    def on_pushButton_clicked(self):

        edit1 = self.ui.lineEdit.text()
        print(edit1)
        edit2 = self.ui.lineEdit_2.text()
        print(edit2)
        edit3 = self.ui.lineEdit_3.text()
        print(edit3)
        edit4 = self.ui.lineEdit_4.text()
        print(edit4)

        excel_data_name = edit1
        sheet_name = edit3

        excel_data_template_name = edit2
        doc_template_name = edit4

        gen_byd(excel_data_name=excel_data_name, sheet_name=sheet_name, doc_template_name=doc_template_name,
                excel_data_template_name=excel_data_template_name)


if __name__ == '__main__':
    app=QApplication(sys.argv)
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())
    mywin=QmyWindow()
    mywin.show()
    sys.exit(app.exec_())