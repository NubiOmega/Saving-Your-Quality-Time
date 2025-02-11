# Form implementation generated from reading ui file 'UI\ui_batchxlstoxlsx.ui'
#
# Created by: PyQt6 UI code generator 6.4.2
#
# WARNING: Any manual changes made to this file will be lost when pyuic6 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt6 import QtCore, QtGui, QtWidgets


class Ui_ConvertWindow(object):
    def setupUi(self, ConvertWindow):
        ConvertWindow.setObjectName("ConvertWindow")
        ConvertWindow.resize(392, 620)
        ConvertWindow.setMaximumSize(QtCore.QSize(580, 620))
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/IMAGE/bg/logo.ico"), QtGui.QIcon.Mode.Normal, QtGui.QIcon.State.Off)
        ConvertWindow.setWindowIcon(icon)
        self.main_vLayout = QtWidgets.QWidget(parent=ConvertWindow)
        self.main_vLayout.setObjectName("main_vLayout")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.main_vLayout)
        self.verticalLayout.setObjectName("verticalLayout")
        self.label_judul = QtWidgets.QLabel(parent=self.main_vLayout)
        font = QtGui.QFont()
        font.setFamily("Roboto")
        font.setPointSize(14)
        font.setBold(True)
        self.label_judul.setFont(font)
        self.label_judul.setObjectName("label_judul")
        self.verticalLayout.addWidget(self.label_judul)
        self.textEditDragDropFiles = QtWidgets.QTextEdit(parent=self.main_vLayout)
        self.textEditDragDropFiles.setObjectName("textEditDragDropFiles")
        self.verticalLayout.addWidget(self.textEditDragDropFiles)
        self.sourceBtn_hLayout = QtWidgets.QHBoxLayout()
        self.sourceBtn_hLayout.setSpacing(6)
        self.sourceBtn_hLayout.setObjectName("sourceBtn_hLayout")
        self.LokasiSumberFile_btn = QtWidgets.QPushButton(parent=self.main_vLayout)
        self.LokasiSumberFile_btn.setMinimumSize(QtCore.QSize(140, 32))
        self.LokasiSumberFile_btn.setMaximumSize(QtCore.QSize(182, 32))
        self.LokasiSumberFile_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.LokasiSumberFile_btn.setObjectName("LokasiSumberFile_btn")
        self.sourceBtn_hLayout.addWidget(self.LokasiSumberFile_btn)
        spacerItem = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Minimum)
        self.sourceBtn_hLayout.addItem(spacerItem)
        self.verticalLayout.addLayout(self.sourceBtn_hLayout)
        self.listFileItems_lokasiSumber = QtWidgets.QListWidget(parent=self.main_vLayout)
        self.listFileItems_lokasiSumber.setMinimumSize(QtCore.QSize(0, 160))
        self.listFileItems_lokasiSumber.setObjectName("listFileItems_lokasiSumber")
        self.verticalLayout.addWidget(self.listFileItems_lokasiSumber)
        self.outputBtn_hLayout = QtWidgets.QHBoxLayout()
        self.outputBtn_hLayout.setObjectName("outputBtn_hLayout")
        self.LokasiOutputFolder_btn = QtWidgets.QPushButton(parent=self.main_vLayout)
        self.LokasiOutputFolder_btn.setMinimumSize(QtCore.QSize(0, 32))
        self.LokasiOutputFolder_btn.setMaximumSize(QtCore.QSize(200, 32))
        self.LokasiOutputFolder_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.LokasiOutputFolder_btn.setObjectName("LokasiOutputFolder_btn")
        self.outputBtn_hLayout.addWidget(self.LokasiOutputFolder_btn)
        spacerItem1 = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Minimum)
        self.outputBtn_hLayout.addItem(spacerItem1)
        self.verticalLayout.addLayout(self.outputBtn_hLayout)
        self.listFileItems_lokasiTujuan = QtWidgets.QListWidget(parent=self.main_vLayout)
        self.listFileItems_lokasiTujuan.setMinimumSize(QtCore.QSize(0, 160))
        self.listFileItems_lokasiTujuan.setObjectName("listFileItems_lokasiTujuan")
        self.verticalLayout.addWidget(self.listFileItems_lokasiTujuan)
        self.progressBar = QtWidgets.QProgressBar(parent=self.main_vLayout)
        self.progressBar.setMinimumSize(QtCore.QSize(320, 20))
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar")
        self.verticalLayout.addWidget(self.progressBar, 0, QtCore.Qt.AlignmentFlag.AlignHCenter)
        self.submitBtn_hlLayout = QtWidgets.QHBoxLayout()
        self.submitBtn_hlLayout.setObjectName("submitBtn_hlLayout")
        self.konversi_Btn = QtWidgets.QPushButton(parent=self.main_vLayout)
        self.konversi_Btn.setMinimumSize(QtCore.QSize(180, 32))
        self.konversi_Btn.setMaximumSize(QtCore.QSize(180, 32))
        self.konversi_Btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.konversi_Btn.setObjectName("konversi_Btn")
        self.submitBtn_hlLayout.addWidget(self.konversi_Btn)
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Minimum)
        self.submitBtn_hlLayout.addItem(spacerItem2)
        self.openFolderOutputXLSX_btn = QtWidgets.QPushButton(parent=self.main_vLayout)
        self.openFolderOutputXLSX_btn.setMinimumSize(QtCore.QSize(180, 32))
        self.openFolderOutputXLSX_btn.setMaximumSize(QtCore.QSize(180, 32))
        self.openFolderOutputXLSX_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.openFolderOutputXLSX_btn.setObjectName("openFolderOutputXLSX_btn")
        self.submitBtn_hlLayout.addWidget(self.openFolderOutputXLSX_btn)
        self.verticalLayout.addLayout(self.submitBtn_hlLayout)
        ConvertWindow.setCentralWidget(self.main_vLayout)

        self.retranslateUi(ConvertWindow)
        QtCore.QMetaObject.connectSlotsByName(ConvertWindow)

    def retranslateUi(self, ConvertWindow):
        _translate = QtCore.QCoreApplication.translate
        ConvertWindow.setWindowTitle(_translate("ConvertWindow", "Batch Convert XLS to XLSX"))
        self.label_judul.setText(_translate("ConvertWindow", "Batch Konversi xls ke xlsx"))
        self.textEditDragDropFiles.setPlaceholderText(_translate("ConvertWindow", "Drag file atau folder .xls di sini"))
        self.LokasiSumberFile_btn.setText(_translate("ConvertWindow", "Pilih Lokasi File .xls"))
        self.LokasiOutputFolder_btn.setText(_translate("ConvertWindow", "Pilih Lokasi Folder Output .xlsx"))
        self.konversi_Btn.setText(_translate("ConvertWindow", "Mulai Konversi"))
        self.openFolderOutputXLSX_btn.setText(_translate("ConvertWindow", "Buka Folder .xlsx"))


# if __name__ == "__main__":
#     import sys
#     app = QtWidgets.QApplication(sys.argv)
#     ConvertWindow = QtWidgets.QMainWindow()
#     ui = Ui_ConvertWindow()
#     ui.setupUi(ConvertWindow)
#     ConvertWindow.show()
#     sys.exit(app.exec())
