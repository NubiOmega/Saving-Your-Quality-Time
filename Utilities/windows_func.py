import os
import subprocess
from PyQt6 import QtWidgets, QtCore
from Utilities.dragndrop_files_func import *

def browse_files(self):
    try:
        file_dialog = QtWidgets.QFileDialog(self)
        file_dialog.setWindowTitle("Pilih File Excel atau Folder")
        file_dialog.setFileMode(QtWidgets.QFileDialog.FileMode.ExistingFiles)
        file_dialog.setViewMode(QtWidgets.QFileDialog.ViewMode.List)
        file_dialog.setNameFilter("Excel Files (*xls *.xlsx)")
        if file_dialog.exec():
            file_paths = file_dialog.selectedFiles()
            for file_path in file_paths:
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
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat memilih file: {e}")

def validate_and_start_import(self):
    try:
        if self.ui.daftarInputFiles_treeWidget.topLevelItemCount() == 0:
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Tidak ada file yang diunggah. Silakan unggah file Excel .xlsx terlebih dahulu!")
        else:
            self.start_import_excel()
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat memulai impor: {e}")

def open_selected_output_folder(self):
    try:
        folder_path = os.path.abspath("hasil data")
        if os.path.exists(folder_path):
            if os.name == 'nt':
                subprocess.Popen(f'explorer "{folder_path}"', shell=True)
            else:
                subprocess.Popen(['xdg-open', folder_path])
        else:
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Folder output masih belum tersedia,\nsilahkan lakukan proses impor dan isi data Excel .xlsx terlebih dahulu!")
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat membuka folder output: {e}")

def open_selected_source_file(self):
    try:
        current_item = self.ui.daftarInputFiles_treeWidget.currentItem()
        if current_item:
            file_path = current_item.toolTip(0)
            if os.path.exists(file_path):
                if os.name == 'nt':
                    os.startfile(file_path)
                else:
                    subprocess.call(['xdg-open', file_path])
            else:
                QtWidgets.QMessageBox.warning(self, "Peringatan", "File tidak ditemukan!")
        else:
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Tidak ada file yang dipilih!")
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat membuka file: {e}")

def open_selected_source_folder(self):
    try:
        current_item = self.ui.daftarInputFiles_treeWidget.currentItem()
        if current_item:
            folder_path = os.path.dirname(current_item.toolTip(0))
            if os.path.exists(folder_path):
                if os.name == 'nt':
                    os.startfile(folder_path)
                else:
                    subprocess.call(['xdg-open', folder_path])
            else:
                QtWidgets.QMessageBox.warning(self, "Peringatan", "Folder tidak ditemukan!\nSilahkan pilih file Excel .xlsx terlebih dahulu")
        else:
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Tidak ada folder yang dipilih!")
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat membuka folder: {e}")

def open_output_context_menu(self, position):
    try:
        menu = QtWidgets.QMenu()
        open_file_action = menu.addAction("Buka File")
        open_folder_action = menu.addAction("Buka Folder")
        delete_selected_action = menu.addAction("Hapus Item Terpilih")
        delete_all_action = menu.addAction("Hapus Semua Item")

        action = menu.exec(self.ui.daftarOutputFiles_treeWidget.mapToGlobal(position))
        if action == open_file_action:
            self.open_selected_output_file()
        elif action == open_folder_action:
            self.open_selected_output_folder()
        elif action == delete_selected_action:
            self.delete_selected_output_items()
        elif action == delete_all_action:
            self.delete_all_output_items()
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat membuka menu konteks: {e}")

def open_source_context_menu(self, position):
    try:
        menu = QtWidgets.QMenu()
        open_file_action = menu.addAction("Buka File")
        open_folder_action = menu.addAction("Buka Folder")
        delete_selected_action = menu.addAction("Hapus Item Terpilih")
        delete_all_action = menu.addAction("Hapus Semua Item")

        action = menu.exec(self.ui.daftarInputFiles_treeWidget.mapToGlobal(position))
        if action == open_file_action:
            self.open_selected_source_file()
        elif action == open_folder_action:
            self.open_selected_source_folder()
        elif action == delete_selected_action:
            self.delete_selected_input_items()
        elif action == delete_all_action:
            self.delete_all_input_items()
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat membuka menu konteks: {e}")

def open_selected_output_file(self):
    try:
        current_item = self.ui.daftarOutputFiles_treeWidget.currentItem()
        if current_item:
            file_path = os.path.join("hasil data", current_item.text(0))
            if os.path.exists(file_path):
                if os.name == 'nt':
                    os.startfile(file_path)
                else:
                    subprocess.call(['xdg-open', file_path])
            else:
                QtWidgets.QMessageBox.warning(self, "Peringatan", "Belum ada File Excel yang dipilih,\nSilahkan lakukan proses impor dan isi data Excel .xlsx terlebih dahulu!!")
        else:
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Tidak ada item yang dipilih!")
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat membuka file output: {e}")

def open_output_folder(self):
    try:
        folder_path = os.path.abspath("hasil data")
        if os.path.exists(folder_path):
            if os.name == 'nt':
                subprocess.Popen(f'explorer "{folder_path}"', shell=True)
            else:
                subprocess.Popen(['xdg-open', folder_path])
        else:
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Belum ada file Excel yang diproses,\nSilahkan lakukan proses impor dan isi data Excel .xlsx terlebih dahulu!!")
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat membuka folder output: {e}")

def delete_selected_input_items(self):
    try:
        selected_items = self.ui.daftarInputFiles_treeWidget.selectedItems()
        if selected_items:
            for item in selected_items:
                index = self.ui.daftarInputFiles_treeWidget.indexOfTopLevelItem(item)
                self.ui.daftarInputFiles_treeWidget.takeTopLevelItem(index)
        else:
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Tidak ada item yang dipilih!")
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat menghapus item terpilih: {e}")

def delete_all_input_items(self):
    try:
        self.ui.daftarInputFiles_treeWidget.clear()
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat menghapus semua item: {e}")

def delete_selected_output_items(self):
    try:
        selected_items = self.ui.daftarOutputFiles_treeWidget.selectedItems()
        if selected_items:
            for item in selected_items:
                index = self.ui.daftarOutputFiles_treeWidget.indexOfTopLevelItem(item)
                self.ui.daftarOutputFiles_treeWidget.takeTopLevelItem(index)
        else:
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Tidak ada item yang dipilih!")
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat menghapus item terpilih: {e}")

def delete_all_output_items(self):
    try:
        self.ui.daftarOutputFiles_treeWidget.clear()
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat menghapus semua item: {e}")
