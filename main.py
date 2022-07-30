import os
import sys

from GUI import *
from PyQt5.QtWidgets import QWidget, QApplication, QFileDialog, QMessageBox
from PyQt5.QtCore import pyqtSlot, pyqtSignal

from Thread import WorkThread


class UI(QWidget, Ui_Form):
    ThreadSignal = pyqtSignal

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
        self.WorkThread = None
        self.TimeTitle = '时间'
        self.NameTitle = '昵称'
        self.VIDTitle = 'QQ（邮箱）'
        self.ContentTitle = '内容'

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
    def on_SourceFilePathInput_editingFinished(self):
        self.SourceFilePath = self.SourceFilePathInput.text()

    @pyqtSlot()
    def on_TargetPathInput_editingFinished(self):
        self.TargetPath = self.TargetPathInput.text()

    @pyqtSlot()
    def on_TableNameInput_editingFinished(self):
        self.TableName = self.TableNameInput.text()

    @pyqtSlot()
    def on_customTime_editingFinished(self):
        if self.customTime.text() != '':
            self.TimeTitle = self.customTime.text()
        else:
            self.TimeTitle = '时间'

    @pyqtSlot()
    def on_customName_editingFinished(self):
        if self.customName != '':
            self.NameTitle = self.customName.text()
            pass
        else:
            self.NameTitle = '昵称'

    @pyqtSlot()
    def on_customVID_editingFinished(self):
        if self.customVID != '':
            self.VIDTitle = self.customVID.text()
        else:
            self.VIDTitle = 'QQ（邮箱）'

    @pyqtSlot()
    def on_customContent_editingFinished(self):
        if self.customContent != '':
            self.ContentTitle = self.customContent.text()
        else:
            self.ContentTitle = '内容'

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

    @pyqtSlot(bool)
    def on_customtitlecheckBox_clicked(self, flag: bool):
        self.label_4.setEnabled(flag)
        self.label_5.setEnabled(flag)
        self.label_6.setEnabled(flag)
        self.label.setEnabled(flag)
        self.customVID.setEnabled(flag)
        self.customName.setEnabled(flag)
        self.customContent.setEnabled(flag)
        self.customTime.setEnabled(flag)
        if not flag:
            self.TimeTitle = '时间'
            self.NameTitle = '昵称'
            self.VIDTitle = 'QQ（邮箱）'
            self.ContentTitle = '内容'

    def change_progress(self, progress: int):
        # print(progress)
        self.progressBar.setValue(progress)
        if progress == 100:
            self.StartButton.setEnabled(True)
            QMessageBox.information(self, '已完成！', '任务已完成！！', QMessageBox.Ok)
            # os.startfile(self.TargetPath)

    def start_check(self) -> bool:
        if self.SourceFilePath == '' or os.path.isfile(self.SourceFilePath):
            # if True:
            if os.path.isdir(self.TargetPath):
                if self.TableName == '':
                    self.TableName = 'Test'
                return True
            elif self.TargetPath == './out':
                os.mkdir('./out')
                if self.TableName == '':
                    self.TableName = 'Test'
                return True
            else:
                QMessageBox.warning(self, '文件夹不存在', '文件夹不存在，请重新输入！！！')
                return False
        else:
            QMessageBox.warning(self, "文件不存在", "文件不存在，请重新输入！！！", QMessageBox.Ok)
            return False

    @pyqtSlot()
    def on_StartButton_clicked(self):
        self.StartButton.setDisabled(True)
        if self.start_check():
            self.progressBar.setValue(0)
            self.WorkThread = WorkThread(self.SourceFilePath, self.TargetPath, self.TableName,
                                         (self.isOutTime, self.isOutName, self.isOutVID, self.isOutContent),
                                         (self.TimeTitle, self.NameTitle, self.VIDTitle, self.ContentTitle))
            self.WorkThread.ProgressRateSignal.connect(self.change_progress)
            self.WorkThread.start()
        else:
            self.StartButton.setEnabled(True)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    m = UI()
    m.show()
    sys.exit(app.exec_())
