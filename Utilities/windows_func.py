import os
import subprocess
from PyQt6 import QtWidgets
from batchconvert import BatchConvertApp
from compare_xls_xlsx import CompareXlsXlsxWindow

def browse_files(self):
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

def validate_and_start_import(self):
    if self.ui.listItems_filesSource.count() == 0:
        QtWidgets.QMessageBox.warning(self, "Peringatan", "Tidak ada file yang diunggah. Silakan unggah file Excel .xlsx terlebih dahulu!")
    else:
        self.start_import_excel()

def open_selected_output_folder(self):
    folder_path = os.path.abspath("hasil data")
    if os.path.exists(folder_path):
        os.startfile(folder_path) if os.name == 'nt' else subprocess.call(['xdg-open', folder_path])
    else:
        QtWidgets.QMessageBox.warning(self, "Peringatan", "Folder output masih belum tersedia,\nsilahkan lakukan proses impor dan isi data Excel .xlsx terlebih dahulu!")

def open_selected_source_file(self):
    current_item = self.ui.listItems_filesSource.currentItem()
    if current_item:
        file_path = current_item.text()
        if os.path.exists(file_path):
            os.startfile(file_path) if os.name == 'nt' else subprocess.call(['xdg-open', file_path])
        else:
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Tidak ada file yang dipilih !")

def open_selected_source_folder(self):
    current_item = self.ui.listItems_filesSource.currentItem()
    if current_item:
        folder_path = os.path.dirname(current_item.text())
        if os.path.exists(folder_path):
            os.startfile(folder_path) if os.name == 'nt' else subprocess.call(['xdg-open', folder_path])
        else:
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Tidak ada folder yang dipilih!\nSilahkan pilih folder file Excel .xlsx terlebih dahulu")

def open_output_context_menu(self, position):
    menu = QtWidgets.QMenu()
    open_file_action = menu.addAction("Buka File")
    open_folder_action = menu.addAction("Buka Folder")

    action = menu.exec(self.ui.listItems_outputFilesXLSX.mapToGlobal(position))
    if action == open_file_action:
        self.open_selected_output_file()
    elif action == open_folder_action:
        self.open_selected_output_folder()

def open_source_context_menu(self, position):
    menu = QtWidgets.QMenu()
    open_file_action = menu.addAction("Buka File")
    open_folder_action = menu.addAction("Buka Folder")

    action = menu.exec(self.ui.listItems_filesSource.mapToGlobal(position))
    if action == open_file_action:
        self.open_selected_source_file()
    elif action == open_folder_action:
        self.open_selected_source_folder()

def open_selected_output_file(self):
    current_item = self.ui.listItems_outputFilesXLSX.currentItem()
    if current_item:
        file_path = os.path.join("hasil data", current_item.text())
        if os.path.exists(file_path):
            os.startfile(file_path) if os.name == 'nt' else subprocess.call(['xdg-open', file_path])
        else:
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Belum ada File Excel yang dipilih,\nSilahkan lakukan proses impor dan isi data Excel .xlsx terlebih dahulu!!")

def open_output_folder(self):
    if os.path.exists("hasil data"):
        folder_path = os.path.abspath("hasil data")
        if os.name == 'nt':  
            subprocess.Popen(f'explorer "{folder_path}"', shell=True)
        else:  
            subprocess.Popen(['xdg-open', folder_path])
    else:
        QtWidgets.QMessageBox.warning(self, "Peringatan", "Belum ada file Excel yang diproses,\nSilahkan lakukan proses impor dan isi data Excel .xlsx terlebih dahulu!!")

        