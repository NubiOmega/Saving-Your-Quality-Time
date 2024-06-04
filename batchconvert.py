import os
import sys
import win32com.client as win32
from PyQt6 import QtCore, QtGui, QtWidgets
from UI.ui_batchxlstoxlsx import Ui_ConvertWindow
from Utilities.dragndrop_files_func import *
import logging

logging.basicConfig(filename='konversi_xls_ke_xlsx_errors_log.txt', level=logging.ERROR)
def log_error(message):
    logging.error(message)

class BatchConvertApp(QtWidgets.QMainWindow, Ui_ConvertWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        
        self.LokasiSumberFile_btn.clicked.connect(self.select_source_files)
        self.LokasiOutputFolder_btn.clicked.connect(self.select_output_folder)
        self.konversi_Btn.clicked.connect(self.start_conversion)
        self.openFolderOutputXLSX_btn.clicked.connect(self.open_output_folder)

        # Menghubungkan QTextEdit yang sudah ada dengan fungsi drag and drop
        self.textEditDragDropFiles.setAcceptDrops(True)
        self.textEditDragDropFiles.dragEnterEvent = self.dragEnterEvent
        self.textEditDragDropFiles.dragMoveEvent = self.dragMoveEvent
        self.textEditDragDropFiles.dropEvent = self.dropEvent_convert_xls_xlsx

        self.source_files = []
        self.output_folder = ""
        self.progressBar.setValue(0)

        # Mengatur mode pemilihan QListWidget ke MultiSelection
        self.listFileItems_lokasiSumber.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.MultiSelection)
        self.listFileItems_lokasiTujuan.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.MultiSelection)

        # Menambahkan event handler untuk klik kanan
        self.listFileItems_lokasiSumber.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.listFileItems_lokasiSumber.customContextMenuRequested.connect(self.show_source_context_menu)
        
        self.listFileItems_lokasiTujuan.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.listFileItems_lokasiTujuan.customContextMenuRequested.connect(self.show_output_context_menu)


    def select_source_files(self):
        try:
            files, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "Pilih File Sumber", "", "Excel Files (*.xls)")
            if files:
                for filename in files:
                    if not self.fileExists_convert_xls_xlsx(filename):
                        self.source_files.append(filename)
                        item = QtWidgets.QListWidgetItem(os.path.basename(filename))
                        item.setToolTip(filename)
                        self.listFileItems_lokasiSumber.addItem(item)
            else:
                raise Exception("Tidak ada file yang dipilih.")
        except Exception as e:
            self.show_message_box("Error", f"Terjadi kesalahan: {str(e)}", QtWidgets.QMessageBox.Icon.Critical)

    def select_output_folder(self):
        try:
            folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Pilih Folder Output")
            if folder:
                self.output_folder = folder
            else:
                raise Exception("Tidak ada folder yang dipilih.")
        except Exception as e:
            self.show_message_box("Error", f"Terjadi kesalahan: {str(e)}", QtWidgets.QMessageBox.Icon.Critical)

    def open_output_folder(self):
        try:
            if self.output_folder:
                os.startfile(self.output_folder)
            else:
                raise Exception("Folder output belum dipilih.")
        except Exception as e:
            self.show_message_box("Peringatan", f"Terjadi kesalahan: {str(e)}", QtWidgets.QMessageBox.Icon.Warning)

    def convert_xls_to_xlsx(self, source_file, dest_file):
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        try:
            wb = excel.Workbooks.Open(source_file)
            wb.SaveAs(dest_file, FileFormat=51)  # 51 adalah format untuk xlsx
            wb.Close()
            return True, None
        except Exception as e:
            return False, f"Gagal mengonversi {source_file}. Error: {str(e)}"
        finally:
            excel.Application.Quit()

    def start_conversion(self):
        try:
            if not self.source_files or not self.output_folder:
                raise Exception("Silakan pilih file sumber dan folder output terlebih dahulu.")

            total_files = len(self.source_files)
            if total_files == 0:
                raise Exception("Tidak ada file .xls yang dipilih.")
            
            if not os.listdir(self.output_folder):  # Check if output folder is empty
                overwrite_all = True  # Automatically overwrite all files if output folder is empty
            else:
                reply = QtWidgets.QMessageBox.question(
                    self, 'Konfirmasi Menimpa File',
                    'Apakah Anda ingin menimpa semua file yang sudah ada di folder output?',
                    QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
                    QtWidgets.QMessageBox.StandardButton.No
                )
                overwrite_all = reply == QtWidgets.QMessageBox.StandardButton.Yes

            self.progressBar.setValue(0)
            successful_conversions = 0

            self.listFileItems_lokasiTujuan.clear()  # Menghapus daftar item pada listFileItems_lokasiTujuan sebelum konversi baru

            for index, source_file in enumerate(self.source_files):
                filename = os.path.basename(source_file)
                dest_file = os.path.abspath(os.path.join(self.output_folder, filename.replace('.xls', '.xlsx')))

                if os.path.exists(dest_file):
                    if not overwrite_all:
                        continue
                    else:
                        os.remove(dest_file)

                success, error = self.convert_xls_to_xlsx(source_file, dest_file)
                if success:
                    item = QtWidgets.QListWidgetItem(filename.replace('.xls', '.xlsx'))
                    item.setToolTip(dest_file)
                    self.listFileItems_lokasiTujuan.addItem(item)
                    successful_conversions += 1
                else:
                    self.show_message_box('Gagal', f'Gagal mengonversi {source_file}. Error: {error}', QtWidgets.QMessageBox.Icon.Critical)
                self.progressBar.setValue(int((index + 1) / total_files * 100))

            if successful_conversions == total_files:
                self.show_message_box('Selesai', 'Semua file berhasil dikonversi.', QtWidgets.QMessageBox.Icon.Information)
            else:
                self.show_message_box('Selesai', f'Proses konversi selesai. {successful_conversions} dari {total_files} file berhasil dikonversi.', QtWidgets.QMessageBox.Icon.Information)
        except Exception as e:
            self.show_message_box("Error", f"Terjadi kesalahan: {str(e)}", QtWidgets.QMessageBox.Icon.Critical)


    def show_message_box(self, title, message, icon):
        log_error(f"{title}: {message}")
        msg = QtWidgets.QMessageBox(self)
        msg.setIcon(icon)
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.exec()

    def show_source_context_menu(self, position):
        try:
            menu = QtWidgets.QMenu()
            open_file_action = menu.addAction("Buka File")
            open_folder_action = menu.addAction("Buka Folder")
            delete_all_files_action = menu.addAction("Hapus Semua File")
            delete_selected_files_action = menu.addAction("Hapus File yang Dipilih")
            action = menu.exec(self.listFileItems_lokasiSumber.mapToGlobal(position))

            if action == open_file_action:
                self.open_selected_file(self.listFileItems_lokasiSumber)
            elif action == open_folder_action:
                if self.source_files:
                    self.open_selected_folder(self.listFileItems_lokasiSumber, os.path.dirname(self.source_files[0]))
                else:
                    raise Exception("Tidak ada folder sumber yang dipilih.")
            elif action == delete_all_files_action:
                self.delete_all_files(self.listFileItems_lokasiSumber, self.source_files)
            elif action == delete_selected_files_action:
                self.delete_selected_files(self.listFileItems_lokasiSumber, self.source_files)
        except Exception as e:
            self.show_message_box("Error", f"Terjadi kesalahan: {str(e)}", QtWidgets.QMessageBox.Icon.Critical)

    def show_output_context_menu(self, position):
        try:
            menu = QtWidgets.QMenu()
            open_file_action = menu.addAction("Buka File")
            open_folder_action = menu.addAction("Buka Folder")
            delete_selected_files_action = menu.addAction("Hapus File yang Dipilih")
            action = menu.exec(self.listFileItems_lokasiTujuan.mapToGlobal(position))

            if action == open_file_action:
                self.open_selected_file(self.listFileItems_lokasiTujuan)
            elif action == open_folder_action:
                if self.output_folder:
                    self.open_selected_folder(self.listFileItems_lokasiTujuan, self.output_folder)
                else:
                    raise Exception("Tidak ada folder output yang dipilih.")
            elif action == delete_selected_files_action:
                self.delete_selected_files(self.listFileItems_lokasiTujuan)
        except Exception as e:
            self.show_message_box("Error", f"Terjadi kesalahan: {str(e)}", QtWidgets.QMessageBox.Icon.Critical)

    def delete_all_files(self, list_widget, source_list=None):
        try:
            list_widget.clear()
            if source_list:
                source_list.clear()
        except Exception as e:
            self.show_message_box("Error", f"Terjadi kesalahan: {str(e)}", QtWidgets.QMessageBox.Icon.Critical)

    def delete_selected_files(self, list_widget, source_list=None):
        try:
            selected_items = list_widget.selectedItems()
            for item in selected_items:
                file_path = item.toolTip()
                if source_list:
                    source_list.remove(file_path)
                list_widget.takeItem(list_widget.row(item))
        except Exception as e:
            self.show_message_box("Error", f"Terjadi kesalahan: {str(e)}", QtWidgets.QMessageBox.Icon.Critical)

    def open_selected_file(self, list_widget):
        try:
            selected_item = list_widget.currentItem()
            if selected_item:
                file_path = selected_item.toolTip()
                if os.path.exists(file_path):
                    os.startfile(file_path)
                else:
                    raise Exception("File tidak ditemukan.")
        except Exception as e:
            self.show_message_box("Peringatan", f"Terjadi kesalahan: {str(e)}", QtWidgets.QMessageBox.Icon.Warning)

    def open_selected_folder(self, list_widget, folder):
        try:
            if os.path.exists(folder):
                os.startfile(folder)
            else:
                raise Exception("Folder tidak ditemukan.")
        except Exception as e:
            self.show_message_box("Peringatan", f"Terjadi kesalahan: {str(e)}", QtWidgets.QMessageBox.Icon.Warning)
    
    # Metode drag drop file dari Utilities\dragndrop_files_func.py
    dragEnterEvent = dragEnterEvent
    dragMoveEvent = dragMoveEvent
    dropEvent_convert_xls_xlsx = dropEvent_convert_xls_xlsx
    fileExists_convert_xls_xlsx = fileExists_convert_xls_xlsx


# if __name__ == "__main__":
#     app = QtWidgets.QApplication(sys.argv)
#     window = BatchConvertApp()
#     window.show()
#     sys.exit(app.exec())
