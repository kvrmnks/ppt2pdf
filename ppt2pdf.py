from PyQt5 import QtWidgets, QtGui, QtCore
import sys
import os
import mytools
import math


class MyProgressBar(QtWidgets.QDialog):

    def update_value(self, x: int) -> None:
        self.progressbar.setValue(x)

    def __init__(self, title: str, label: str):
        super(MyProgressBar, self).__init__()
        self.progressbar = QtWidgets.QProgressBar()
        self.setWindowTitle(title)
        self.label = QtWidgets.QLabel(label)
        self.layout = QtWidgets.QVBoxLayout()
        self.setLayout(self.layout)
        self.layout.addWidget(self.label)
        self.layout.addWidget(self.progressbar)
        self.resize(500, 100)


class ConvertWork(QtCore.QThread):
    trigger = QtCore.pyqtSignal(int, str)
    during_trigger = QtCore.pyqtSignal(int)

    def __init__(self, fileName: list, path: list, index: list):
        super(ConvertWork, self).__init__()
        self.fileName = fileName
        self.path = path
        self.index = index
        print(fileName, path, index)
        # self.result = ''

    def run(self):
        cnt = 0
        for i in range(len(self.fileName)):
            # try:
            result = mytools.convert_ppt_pdf(self.fileName[i], self.path[i])
            self.trigger.emit(self.index[i], result)
            cnt = cnt + 1
            # except Exception as e:
            #     print(e)
            # finally:
            self.during_trigger.emit(math.ceil(100 * cnt / len(self.fileName)))


class MergeWork(QtCore.QThread):
    during_trigger = QtCore.pyqtSignal(int)

    def __init__(self, fileName: list, labelList: list, out: str):
        super(MergeWork, self).__init__()
        self.fileName = fileName
        self.labelList = labelList
        self.out = out

    def run(self) -> None:
        for i in mytools.merge_pdf(self.fileName, self.labelList, self.out):
            self.during_trigger.emit(math.ceil(100 * i / len(self.fileName)))


class MyModel(QtGui.QStandardItemModel):
    def __init__(self):
        super(MyModel, self).__init__()
        self.convert_thread = None
        self.merge_thread = None

    def change_status(self, index: int, path: str):
        print('in change staus', index, path)
        pdfpath = self.item(index, 2)
        status = self.item(index, 5)
        pdfpath.setText(path)
        status.setText('pdf已存在')

    def change_finished(self):
        QtWidgets.QMessageBox.information(None, 'ppt2pdf', 'pdf生成完毕',
                                          QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)

    def merge_finished(self):
        QtWidgets.QMessageBox.information(None, 'ppt2pdf', 'pdf合并完毕',
                                          QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)

    def dropMimeData(self, data, action, row, col, parent):
        """
        Always move the entire row, and don't allow column "shifting"
        """
        return super().dropMimeData(data, action, row, 0, parent)

    def convert_pdf(self):
        fileNameList = []
        pathList = []
        indexList = []
        for i in range(self.rowCount()):
            fileName = self.item(i, 0)
            path = self.item(i, 1)
            pdfpath = self.item(i, 2)
            pptpath = self.item(i, 3)
            label = self.item(i, 4)
            status = self.item(i, 5)
            if status.text() == 'pdf不存在':
                fileNameList.append(fileName.text())
                pathList.append(path.text())
                indexList.append(i)
        self.convert_thread = ConvertWork(fileNameList, pathList, indexList)
        self.convert_thread.trigger.connect(self.change_status)
        self.convert_thread.finished.connect(self.change_finished)
        self.convert_thread.start()

    def merge_pdf(self):
        fileList = []
        labelList = []
        out = os.getcwd() + '/main.pdf'
        print('enter merge')
        for i in range(self.rowCount()):
            fileName = self.item(i, 0)
            path = self.item(i, 1)
            pdfpath = self.item(i, 2)
            pptpath = self.item(i, 3)
            label = self.item(i, 4)
            status = self.item(i, 5)
            # main = PyPDF2.PdfFileMerger()
            if pdfpath != '':
                fileList.append(pdfpath.text())
                labelList.append(label.text())
        self.merge_thread = MergeWork(fileList, labelList, out)
        self.merge_thread.finished.connect(self.merge_finished)
        self.merge_thread.start()


class MyStyle(QtWidgets.QProxyStyle):

    def drawPrimitive(self, element, option, painter, widget=None):
        """
        Draw a line across the entire row rather than just the column
        we're hovering over.  This may not always work depending on global
        style - for instance I think it won't work on OSX.
        """
        if element == self.PE_IndicatorItemViewItemDrop and not option.rect.isNull():
            option_new = QtWidgets.QStyleOption(option)
            option_new.rect.setLeft(0)
            if widget:
                option_new.rect.setRight(widget.width())
            option = option_new
        super().drawPrimitive(element, option, painter, widget)


class MyTableView(QtWidgets.QTableView):
    def __init__(self, parent=None):
        super(MyTableView, self).__init__(parent)
        self.model = MyModel()
        self.setStyle(MyStyle())
        self.model.setHorizontalHeaderLabels(['文件名', '文件路径', 'pdf路径', 'ppt/pptx路径', '标签', '状态'])
        self.setModel(self.model)

        self.setAcceptDrops(True)  # 开启drop
        self.setDragEnabled(True)  # 开启拖出功能
        # self.s
        self.setSelectionBehavior(QtWidgets.QTableView.SelectRows)  # 一次选一行
        # self.setSelectionMode(QtWidgets.QTableView.Select)
        self.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
        self.setDragDropOverwriteMode(False)
        QtWidgets.QShortcut(QtGui.QKeySequence(self.tr("Del")), self, self.deleteRows)  # 删除选中行
        self.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeToContents)
        # print('abab')

        self.pb = MyProgressBar('ppt2pdf', '')
        self.pb.setWindowModality(QtCore.Qt.ApplicationModal)
        self.pb.progressbar.setValue(0)

    def clearRows(self):
        for i in range(self.model.rowCount()):
            self.model.removeRow(0)

    def deleteRows(self):
        selectedList = self.selectionModel().selectedRows()
        for x in selectedList:
            self.model.removeRow(x.row().real)

    def dragEnterEvent(self, e: QtGui.QDragEnterEvent) -> None:
        if e.mimeData().hasText() or e.mimeData().hasFormat('application/x-qabstractitemmodeldatalist'):
            e.acceptProposedAction()

    def dragMoveEvent(self, e: QtGui.QDragMoveEvent) -> None:
        pass

    def dragLeaveEvent(self, e: QtGui.QDragLeaveEvent) -> None:
        pass

    def addPDF(self, url):
        """ append a pdf row when dropping """
        basename = QtGui.QStandardItem(os.path.basename(url))
        abspath = QtGui.QStandardItem(os.path.abspath(url))
        pdfpath = QtGui.QStandardItem(os.path.abspath(url))
        pptpath = QtGui.QStandardItem('')
        label = QtGui.QStandardItem(os.path.basename(url)[:-4])
        status = QtGui.QStandardItem('pdf已存在')
        itemList = [basename, abspath, pdfpath, pptpath, label, status]
        for x in itemList:
            x.setDropEnabled(False)
            x.setEditable(False)
        label.setEditable(True)
        self.model.appendRow(itemList)

    def addPPT(self, url):
        """ append a ppt row when dropping """
        basename = QtGui.QStandardItem(os.path.basename(url))
        abspath = QtGui.QStandardItem(os.path.abspath(url))
        pdfpath = QtGui.QStandardItem('')
        pptpath = QtGui.QStandardItem(os.path.abspath(url))
        name = os.path.basename(url)
        if name.endswith('.ppt'):
            name = name[:-4]
        else:
            name = name[:-5]
        label = QtGui.QStandardItem(name)
        status = QtGui.QStandardItem('pdf不存在')
        itemList = [basename, abspath, pdfpath, pptpath, label, status]
        for x in itemList:
            x.setDropEnabled(False)
            x.setEditable(False)
        label.setEditable(True)
        self.model.appendRow(itemList)

    def dropEvent(self, e: QtGui.QDropEvent) -> None:
        if e.mimeData().hasText():
            urls = e.mimeData().text().replace('file:///', '').split('\n')  # 多个文件时这样分割的
            for x in urls:
                x = x.strip()
                if x.endswith('.pdf'):
                    self.addPDF(x)
                elif x.endswith('.ppt') or x.endswith('.pptx'):
                    self.addPPT(x)
        else:
            super().dropEvent(e)


class Table(QtWidgets.QWidget):
    def generate_pdf(self):
        self.tableView.model.convert_pdf()
        self.tableView.pb.progressbar.setValue(0)
        self.tableView.model.convert_thread.during_trigger.connect(self.tableView.pb.update_value)
        self.tableView.pb.label.setText('生成pdf')
        self.tableView.pb.exec_()

    def merge_pdf(self):
        self.tableView.model.merge_pdf()
        self.tableView.pb.progressbar.setValue(0)
        self.tableView.model.merge_thread.during_trigger.connect(self.tableView.pb.update_value)
        self.tableView.pb.label.setText('合并pdf')
        self.tableView.pb.exec_()
        # QtWidgets.QMessageBox.information(self, 'ppt2pdf', 'pdf合并完毕',
        #                                   QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)

    def __init__(self, arg=None):
        super(Table, self).__init__(arg)
        self.setWindowTitle("ppt2pdf")

        self.tableView = MyTableView()

        self.myLayout = QtWidgets.QVBoxLayout()
        self.myLayout.addWidget(self.tableView)
        self.setLayout(self.myLayout)

        self.widget2 = QtWidgets.QWidget()
        self.myLayout2 = QtWidgets.QHBoxLayout()
        self.widget2.setLayout(self.myLayout2)
        self.myLayout.addWidget(self.widget2)

        self.clearPushButton = QtWidgets.QPushButton('清空')
        self.myLayout2.addWidget(self.clearPushButton)
        self.clearPushButton.clicked.connect(self.tableView.clearRows)
        self.generatePDFButton = QtWidgets.QPushButton('生成pdf')
        self.myLayout2.addWidget(self.generatePDFButton)
        self.generatePDFButton.clicked.connect(self.generate_pdf)
        self.mergePDFButton = QtWidgets.QPushButton('合并pdf')
        self.myLayout2.addWidget(self.mergePDFButton)
        self.mergePDFButton.clicked.connect(self.merge_pdf)


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    table = Table()
    table.resize(700, 400)
    table.show()
    sys.exit(app.exec_())
