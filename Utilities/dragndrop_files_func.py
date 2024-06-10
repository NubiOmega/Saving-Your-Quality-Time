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
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                    if not file_path.endswith(('.xls', '.xlsx')):
                        QtWidgets.QMessageBox.warning(self, "Format File Salah", "File yang dimasukkan bukan file Excel .xls atau .xlsx.")
                        continue

                    file_info = QtCore.QFileInfo(file_path)
                    
                    # Mengambil informasi file
                    file_name = file_info.fileName()
                    file_modified_date = file_info.lastModified().toString(QtCore.Qt.DateFormat.ISODate)
                    file_type = file_info.suffix()
                    file_size = f"{file_info.size() / 1024:.2f} KB"
                    
                    # Cek apakah file sudah ada dalam daftar
                    if self.fileExists(file_path):
                        QtWidgets.QMessageBox.warning(self, "File Duplikat", f"File '{file_path}' sudah ada dalam daftar.")
                        continue
                    
                    # Membuat item baru dengan informasi file
                    item = QtWidgets.QTreeWidgetItem([file_name, file_modified_date, file_type, file_size])
                    item.setToolTip(0, file_path)  # Mengatur tooltip untuk item
                    self.ui.daftarInputFiles_treeWidget.addTopLevelItem(item)  # Menambahkan item ke daftarInputFiles_treeWidget
                
            # Mengatur ukuran kolom sesuai dengan isi konten setelah semua item ditambahkan
            for i in range(self.ui.daftarInputFiles_treeWidget.columnCount()):
                self.ui.daftarInputFiles_treeWidget.resizeColumnToContents(i)
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat memproses file: {e}")

def fileExists(self, file_path):
    try:
        # Mendapatkan nama file tanpa ekstensi
        file_name_without_ext = QtCore.QFileInfo(file_path).completeBaseName()
        for index in range(self.ui.daftarInputFiles_treeWidget.topLevelItemCount()):
            item = self.ui.daftarInputFiles_treeWidget.topLevelItem(index)
            # Mendapatkan nama file dari daftar tanpa ekstensi
            item_file_name_without_ext = QtCore.QFileInfo(item.toolTip(0)).completeBaseName()
            if item_file_name_without_ext == file_name_without_ext:
                return True
        return False
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat memeriksa keberadaan file: {e}")
        return False

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
                        # Tambahkan file ke self.source_files
                        self.source_files.append(file_path)

                        # Ambil informasi file
                        file_info = QtCore.QFileInfo(file_path)
                        file_name = file_info.fileName()
                        file_modified_date = file_info.lastModified().toString(QtCore.Qt.DateFormat.ISODate)
                        file_type = file_info.suffix()
                        file_size = f"{file_info.size() / 1024:.2f} KB"

                        # Tambahkan item ke lokasiSumber_treeWidget dengan informasi file
                        item = QtWidgets.QTreeWidgetItem([file_name, file_modified_date, file_type, file_size])
                        item.setToolTip(0, file_path)
                        self.ui.lokasiSumber_treeWidget.addTopLevelItem(item)
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
        file_name = QtCore.QFileInfo(file_path).fileName()
        for index in range(self.ui.lokasiSumber_treeWidget.topLevelItemCount()):
            item = self.ui.lokasiSumber_treeWidget.topLevelItem(index)
            if QtCore.QFileInfo(item.toolTip(0)).fileName() == file_name:
                return True
        return False
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat memeriksa keberadaan file: {e}")
