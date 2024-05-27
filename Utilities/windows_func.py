import os
import subprocess
from PyQt6 import QtWidgets

def browse_files(self):
    try:
        file_dialog = QtWidgets.QFileDialog(self)
        file_dialog.setWindowTitle("Pilih File Excel atau Folder")
        file_dialog.setFileMode(QtWidgets.QFileDialog.FileMode.ExistingFiles)
        file_dialog.setViewMode(QtWidgets.QFileDialog.ViewMode.List)
        file_dialog.setNameFilter("Excel Files (*.xlsx)")
        if file_dialog.exec():
            file_paths = file_dialog.selectedFiles()
            self.ui.listItems_filesSource.clear()  # Mereset konten list item
            for file_path in file_paths:
                self.ui.listItems_filesSource.addItem(file_path)  # Menggunakan path lengkap untuk source files
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat memilih file: {e}")

def validate_and_start_import(self):
    try:
        if self.ui.listItems_filesSource.count() == 0:
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
        current_item = self.ui.listItems_filesSource.currentItem()
        if current_item:
            file_path = current_item.text()
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
        current_item = self.ui.listItems_filesSource.currentItem()
        if current_item:
            folder_path = os.path.dirname(current_item.text())
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

        action = menu.exec(self.ui.listItems_outputFilesXLSX.mapToGlobal(position))
        if action == open_file_action:
            self.open_selected_output_file()
        elif action == open_folder_action:
            self.open_selected_output_folder()
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat membuka menu konteks: {e}")

def open_source_context_menu(self, position):
    try:
        menu = QtWidgets.QMenu()
        open_file_action = menu.addAction("Buka File")
        open_folder_action = menu.addAction("Buka Folder")

        action = menu.exec(self.ui.listItems_filesSource.mapToGlobal(position))
        if action == open_file_action:
            self.open_selected_source_file()
        elif action == open_folder_action:
            self.open_selected_source_folder()
    except Exception as e:
        QtWidgets.QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat membuka menu konteks: {e}")

def open_selected_output_file(self):
    try:
        current_item = self.ui.listItems_outputFilesXLSX.currentItem()
        if current_item:
            file_path = os.path.join("hasil data", current_item.text())
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
