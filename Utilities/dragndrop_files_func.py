# Utilities\dragndrop_files_func.py
from PyQt6 import QtCore, QtWidgets
import os

def dragEnterEvent(self, event):
    try:
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan pada dragEnterEvent: {e}")

def dragMoveEvent(self, event):
    try:
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan pada dragMoveEvent: {e}")

def dropEvent(self, event):
    try:
        if event.mimeData().hasUrls():
            event.setDropAction(QtCore.Qt.DropAction.CopyAction)
            event.accept()
            urls = event.mimeData().urls()
            for url in urls:
                if url.isLocalFile() and url.toLocalFile().endswith('.xlsx'):
                    file_path = url.toLocalFile()
                    if not self.fileExists(file_path):
                        # Tambahkan item ke listItems_filesSource
                        item = QtWidgets.QListWidgetItem(file_path)
                        item.setToolTip(file_path)
                        self.ui.listItems_filesSource.addItem(item)
                    else:
                        QtWidgets.QMessageBox.warning(self, "File Duplikat", f"File '{file_path}' sudah ada dalam daftar.")
                else:
                    QtWidgets.QMessageBox.warning(self, "Format File Salah", f"File '{url.toLocalFile()}' bukan file Excel (.xlsx).")
        else:
            event.ignore()
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan pada dropEvent: {e}")

def fileExists(self, file_path):
    try:
        for row in range(self.ui.listItems_filesSource.count()):
            item = self.ui.listItems_filesSource.item(row)
            if item.toolTip() == file_path:
                return True
        return False
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat memeriksa keberadaan file: {e}")

def dropEvent_convert_xls_xlsx(self, event):
    try:
        if event.mimeData().hasUrls():
            event.setDropAction(QtCore.Qt.DropAction.CopyAction)
            event.accept()
            urls = event.mimeData().urls()
            for url in urls:
                if url.isLocalFile() and url.toLocalFile().endswith('.xls'):
                    file_path = url.toLocalFile()
                    if not self.fileExists_convert_xls_xlsx(file_path):
                        item = QtWidgets.QListWidgetItem(os.path.basename(file_path))
                        item.setToolTip(file_path)
                        self.listFileItems_lokasiSumber.addItem(item)
                        self.source_files.append(file_path)  # Pastikan file ditambahkan ke source_files
                    else:
                        QtWidgets.QMessageBox.warning(self, "File Duplikat", f"File '{file_path}' sudah ada dalam daftar.")
                else:
                    QtWidgets.QMessageBox.warning(self, "Format File Salah", f"File '{url.toLocalFile()}' bukan file Excel (.xls).")
        else:
            event.ignore()
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan pada dropEvent: {e}")

def fileExists_convert_xls_xlsx(self, file_path):
    try:
        for row in range(self.listFileItems_lokasiSumber.count()):
            item = self.listFileItems_lokasiSumber.item(row)
            if item.toolTip() == file_path:
                return True
        return False
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat memeriksa keberadaan file: {e}")

