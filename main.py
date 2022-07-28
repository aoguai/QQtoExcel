import os
import sys

from GUI import *
from PyQt5.QtWidgets import QWidget, QApplication, QFileDialog
from PyQt5.QtCore import pyqtSlot


class UI(QWidget, Ui_Form):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.SourceFilePath = None  # 源文件地址
        self.TargetPath = './out'  # 保存文件夹地址
        self.TargetPathInput.setText(self.TargetPath)
        self.TableName = 'Test'  # 表名
        self.TableNameInput.setText(self.TableName)
        self.isOutTime = True
        self.isOutVID = True
        self.isOutName = True
        self.isOutContent = True

    # 下面为槽函数
    @pyqtSlot()
    def on_SourceFileSelect_clicked(self):
        self.SourceFilePath = QFileDialog.getOpenFileName(self, '打开文件', './', filter='*.txt')[0]
        self.SourceFilePathInput.setText(self.SourceFilePath)

    @pyqtSlot()
    def on_TargetSelect_clicked(self):
        self.TargetPath = QFileDialog.getExistingDirectory()
        self.TargetPathInput.setText(self.TargetPath)

    @pyqtSlot()
    def on_SourceFilePathInput_textEdited(self):
        self.SourceFilePath = self.SourceFilePathInput.text()

    @pyqtSlot()
    def on_TargetPathInput_textEdited(self):
        self.TargetPath = self.TargetPathInput.text()

    @pyqtSlot()
    def on_TableNameInput_textEdited(self):
        self.TableName = self.TableNameInput.text()

    @pyqtSlot(bool)
    def on_isTime_clicked(self, flag: bool):
        self.isOutTime = flag

    @pyqtSlot(bool)
    def on_isName_clicked(self, flag: bool):
        self.isOutName = flag

    @pyqtSlot(bool)
    def on_isVID_clicked(self, flag: bool):
        self.isOutVID = flag

    @pyqtSlot(bool)
    def on_isContent_clicked(self, flag: bool):
        self.isOutContent = flag

    def start_check(self):
        if os.path.isfile(self.SourceFilePath):
            if os.path.isdir(self.TargetPath):
                if self.TableName == '':
                    self.TableName = 'Test'
                pass
    @pyqtSlot()
    def on_StartButton_clicked(self):
        self.StartButton.setDisabled(True)
        self.textBrowser.append("开始...")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    m = UI()
    m.show()
    sys.exit(app.exec_())
