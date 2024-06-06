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
                        # Ambil informasi file
                        file_info = QtCore.QFileInfo(file_path)
                        file_name = file_info.fileName()
                        file_modified_date = file_info.lastModified().toString(QtCore.Qt.DateFormat.ISODate)
                        file_type = file_info.suffix()
                        file_size = f"{file_info.size() / 1024:.2f} KB"

                        # Tambahkan item ke daftarInputFiles_treeWidget dengan informasi file
                        item = QtWidgets.QTreeWidgetItem([file_name, file_modified_date, file_type, file_size])
                        item.setToolTip(0, file_path)  # Mengatur tooltip untuk item
                        self.ui.daftarInputFiles_treeWidget.addTopLevelItem(item)  # Menambahkan item ke daftarInputFiles_treeWidget
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
        for index in range(self.ui.daftarInputFiles_treeWidget.topLevelItemCount()):
            item = self.ui.daftarInputFiles_treeWidget.topLevelItem(index)
            if item.toolTip(0) == file_path:
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
                        self.ui.listFileItems_lokasiSumber.addItem(item)
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
        for row in range(self.ui.listFileItems_lokasiSumber.count()):
            item = self.ui.listFileItems_lokasiSumber.item(row)
            if item.toolTip() == file_path:
                return True
        return False
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat memeriksa keberadaan file: {e}")

