import os
import sys

from GUI import *
from PyQt5.QtWidgets import QWidget, QApplication, QFileDialog, QMessageBox
from PyQt5.QtCore import pyqtSlot, pyqtSignal

from Thread import WorkThread


class UI(QWidget, Ui_Form):
    """
    界面类
    """

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
        self.SourceFilePathInput.installEventFilter(QEventHandler(self))  # 安装事件过滤器，实现拖放
        self.TargetPathInput.installEventFilter(QEventHandler(self))  # 安装事件过滤器，实现拖放
        self.customtitlecheckBox.setToolTip("无需更改可留空")
        self.TargetPathInput.setToolTip("默认或不填为”out“")
        self.TableNameInput.setToolTip("默认或不填为”Test“")
        self.SourceFilePathInput.setToolTip("必填")
        self.SourceFileSelect.setToolTip("文本框支持拖放")
        self.TargetSelect.setToolTip("文本框支持拖放")
        self.progressBar.hide()

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
        if self.TargetPathInput.text() == '':
            self.TargetPath = './out'
        else:
            self.TargetPath = self.TargetPathInput.text()

    @pyqtSlot()
    def on_TableNameInput_editingFinished(self):
        if self.TableNameInput.text() == '':
            self.TableName = 'Test'
        else:
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
        """
        自定义按钮槽函数
        :param flag:
        :return: None
        """
        self.label_4.setEnabled(flag)
        self.label_5.setEnabled(flag)
        self.label_6.setEnabled(flag)
        self.label.setEnabled(flag)
        self.customVID.setEnabled(flag)
        self.customName.setEnabled(flag)
        self.customContent.setEnabled(flag)
        self.customTime.setEnabled(flag)
        self.customVID.setToolTip("无需更改可留空")
        self.customTime.setToolTip("无需更改可留空")
        self.customName.setToolTip("无需更改可留空")
        self.customContent.setToolTip("无需更改可留空")
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
            self.progressBar.hide()
            self.StartButton.show()
        # os.startfile(self.TargetPath)

    def start_check(self) -> bool:
        """
        开始检查
        检查输入是否符合要求
        :return: 符合返回True，不符合返回False
        """
        if self.SourceFilePath == '':
            QMessageBox.warning(self, '请输入文件地址', '请输入文件地址', QMessageBox.Ok)
            return False
        elif not os.path.isfile(self.SourceFilePath):
            QMessageBox.warning(self, "文件不存在", "文件不存在，请重新输入！！！", QMessageBox.Ok)
            return False
        elif self.SourceFilePath[-3:].lower() != 'txt':
            QMessageBox.warning(self, '文件类型错误', '文件类型错误，程序只接受.TXT格式文件')
            return False
        elif not os.path.isdir(self.TargetPath):  # 检查保存文件夹是否已存在
            if self.TargetPath == './out':
                os.mkdir('./out')
                return True
            else:
                QMessageBox.warning(self, '文件夹不存在', '文件夹不存在，请重新输入！！！')
                return False
        else:
            return True


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
            return
        self.StartButton.hide()
        self.progressBar.show()


class QEventHandler(QtCore.QObject):
    def eventFilter(self, obj, event):
        """
		处理窗体内出现的事件，如果有需要则自行添加if判断语句；
		目前已经实现将拖到控件上文件的路径设置为控件的显示文本；
		"""
        if event.type() == QtCore.QEvent.DragEnter:
            event.accept()
        if event.type() == QtCore.QEvent.Drop:
            md = event.mimeData()
            if md.hasUrls():
                # 此处md.urls()的返回值为拖入文件的file路径列表，即支持多文件同时拖入；
                # 此处默认读取第一个文件的路径进行处理，可按照个人需求进行相应的修改
                url = md.urls()[0]
                obj.setText(url.toLocalFile())
                return True
        return super().eventFilter(obj, event)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    m = UI()
    m.show()
    sys.exit(app.exec_())
