import os
import sys
import win32com.client as win32
from PyQt6 import QtCore, QtGui, QtWidgets
from UI.batchxlstoxlsx_ui import Ui_ConvertWindow
from Utilities.dragndrop_files_func import *
import logging
from datetime import datetime

log_folder = "LOG"
log_filename = ""

def setup_logging():
    global log_filename
    if not log_filename:  # Only set up logging if it hasn't been done yet
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_filename = f'{log_folder}/konversi_log_{timestamp}.txt'

        # Memastikan folder LOG sudah ada atau membuatnya jika belum ada
        if not os.path.exists(log_folder):
            os.makedirs(log_folder)

        logging.basicConfig(
            filename=log_filename,
            level=logging.ERROR,
            format='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )

def log_error(message):
    setup_logging()
    logging.error(message)

class BatchConvertApp(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.ui = Ui_ConvertWindow()
        self.ui.setupUi(self)
        
        self.ui.LokasiSumberFile_btn.clicked.connect(self.select_source_files)
        self.ui.LokasiOutputFolder_btn.clicked.connect(self.select_output_folder)
        self.ui.konversi_Btn.clicked.connect(self.start_conversion)
        self.ui.openFolderOutputXLSX_btn.clicked.connect(self.open_output_folder)

        # Menghubungkan QTextEdit yang sudah ada dengan fungsi drag and drop
        self.ui.lokasiSumber_treeWidget.setAcceptDrops(True)
        self.ui.lokasiSumber_treeWidget.dragEnterEvent = self.dragEnterEvent
        self.ui.lokasiSumber_treeWidget.dragMoveEvent = self.dragMoveEvent
        self.ui.lokasiSumber_treeWidget.dropEvent = self.dropEvent_convert_xls_xlsx

        self.source_files = []
        self.output_folder = ""
        self.ui.progressBar.setValue(0)

        # Mengatur mode pemilihan QListWidget ke MultiSelection
        self.ui.lokasiSumber_treeWidget.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.MultiSelection)
        self.ui.lokasiTujuan_treeWidget.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.MultiSelection)

        # Menambahkan event handler untuk klik kanan
        self.ui.lokasiSumber_treeWidget.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.ui.lokasiSumber_treeWidget.customContextMenuRequested.connect(self.show_source_context_menu)
        self.ui.lokasiSumber_treeWidget.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection)
        
        self.ui.lokasiTujuan_treeWidget.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.ui.lokasiTujuan_treeWidget.customContextMenuRequested.connect(self.show_output_context_menu)
        self.ui.lokasiTujuan_treeWidget.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection)

    def select_source_files(self):
        try:
            files, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "Pilih File Sumber", "", "Excel Files (*.xls)")
            if files:
                for filename in files:
                    if not self.fileExists_convert_xls_xlsx(filename):
                        self.source_files.append(filename)
                        file_info = QtCore.QFileInfo(filename)
                        file_name = file_info.fileName()
                        file_modified_date = file_info.lastModified().toString(QtCore.Qt.DateFormat.ISODate)
                        file_type = file_info.suffix()
                        file_size = f"{file_info.size() / 1024:.2f} KB"
                        item = QtWidgets.QTreeWidgetItem([file_name, file_modified_date, file_type, file_size])
                        item.setToolTip(0, filename)
                        self.ui.lokasiSumber_treeWidget.addTopLevelItem(item)
            else:
                raise Exception("Tidak ada file yang dipilih.")
        except Exception as e:
            self.show_message_box("Peringatan", f"{str(e)}", QtWidgets.QMessageBox.Icon.Warning)

    def select_output_folder(self):
        try:
            folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Pilih Folder Output")
            if folder:
                self.output_folder = folder
            else:
                raise Exception("Tidak ada folder yang dipilih.")
        except Exception as e:
            self.show_message_box("Peringatan", f"{str(e)}", QtWidgets.QMessageBox.Icon.Warning)

    def open_output_folder(self):
        try:
            if self.output_folder:
                os.startfile(self.output_folder)
            else:
                raise Exception("Folder output belum dipilih.")
        except Exception as e:
            self.show_message_box("Peringatan", f"{str(e)}", QtWidgets.QMessageBox.Icon.Warning)

    def convert_xls_to_xlsx(self, source_file, dest_file):
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        try:
            wb = excel.Workbooks.Open(source_file)
            wb.SaveAs(dest_file, FileFormat=51)  # 51 adalah format untuk xlsx
            wb.Close()
            return True, None
        except Exception as e:
            log_error(f"Gagal mengonversi {source_file} ke {dest_file}. Error: {str(e)}")
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

            self.ui.progressBar.setValue(0)
            successful_conversions = 0

            self.ui.lokasiTujuan_treeWidget.clear()  # Menghapus daftar item pada lokasiTujuan_treeWidget sebelum konversi baru

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
                    file_info = QtCore.QFileInfo(dest_file)
                    file_name = file_info.fileName()
                    file_modified_date = file_info.lastModified().toString(QtCore.Qt.DateFormat.ISODate)
                    file_type = file_info.suffix()
                    file_size = f"{file_info.size() / 1024:.2f} KB"
                    item = QtWidgets.QTreeWidgetItem([file_name, file_modified_date, file_type, file_size])
                    item.setToolTip(0, dest_file)
                    self.ui.lokasiTujuan_treeWidget.addTopLevelItem(item)
                    successful_conversions += 1
                else:
                    self.show_message_box('Gagal', f'Gagal mengonversi {source_file}. Error: {error}', QtWidgets.QMessageBox.Icon.Warning)
                self.ui.progressBar.setValue(int((index + 1) / total_files * 100))

            if successful_conversions == total_files:
                self.show_message_box('Selesai', 'Semua file berhasil dikonversi.', QtWidgets.QMessageBox.Icon.Information)
            else:
                self.show_message_box('Selesai', f'Proses konversi selesai. {successful_conversions} dari {total_files} file berhasil dikonversi.', QtWidgets.QMessageBox.Icon.Information)
        except Exception as e:
            log_error(f"Proses konversi gagal. Error: {str(e)}")
            self.show_message_box("Peringatan", f"{str(e)}", QtWidgets.QMessageBox.Icon.Warning)

    def show_message_box(self, title, message, icon):
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
            action = menu.exec(self.ui.lokasiSumber_treeWidget.mapToGlobal(position))

            if action == open_file_action:
                self.open_selected_file(self.ui.lokasiSumber_treeWidget)
            elif action == open_folder_action:
                if self.source_files:
                    self.open_selected_folder(self.ui.lokasiSumber_treeWidget, os.path.dirname(self.source_files[0]))
                else:
                    raise Exception("Tidak ada folder sumber yang dipilih.")
            elif action == delete_all_files_action:
                self.delete_all_files(self.ui.lokasiSumber_treeWidget, self.source_files)
            elif action == delete_selected_files_action:
                self.delete_selected_files(self.ui.lokasiSumber_treeWidget, self.source_files)
        except Exception as e:
            self.show_message_box("Peringatan", f"{str(e)}", QtWidgets.QMessageBox.Icon.Warning)

    def show_output_context_menu(self, position):
        try:
            menu = QtWidgets.QMenu()
            open_file_action = menu.addAction("Buka File")
            open_folder_action = menu.addAction("Buka Folder")
            delete_all_files_action = menu.addAction("Hapus Semua File")
            delete_selected_files_action = menu.addAction("Hapus File yang Dipilih")
            action = menu.exec(self.ui.lokasiTujuan_treeWidget.mapToGlobal(position))

            if action == open_file_action:
                self.open_selected_file(self.ui.lokasiTujuan_treeWidget)
            elif action == open_folder_action:
                if self.output_folder:
                    self.open_selected_folder(self.ui.lokasiTujuan_treeWidget, self.output_folder)
                else:
                    raise Exception("Tidak ada folder tujuan yang dipilih.")
            elif action == delete_all_files_action:
                self.delete_all_files(self.ui.lokasiTujuan_treeWidget, [])
            elif action == delete_selected_files_action:
                self.delete_selected_files(self.ui.lokasiTujuan_treeWidget, [])
        except Exception as e:
            self.show_message_box("Peringatan", f"{str(e)}", QtWidgets.QMessageBox.Icon.Warning)

    def open_selected_file(self, treeWidget):
        try:
            selected_items = treeWidget.selectedItems()
            if selected_items:
                for item in selected_items:
                    file_path = item.toolTip(0)
                    if os.path.exists(file_path):
                        os.startfile(file_path)
                    else:
                        raise Exception(f"File tidak ditemukan: {file_path}")
            else:
                raise Exception("Tidak ada file yang dipilih.")
        except Exception as e:
            self.show_message_box("Peringatan", f"{str(e)}", QtWidgets.QMessageBox.Icon.Warning)

    def open_selected_folder(self, treeWidget, folder_path):
        try:
            if os.path.exists(folder_path):
                os.startfile(folder_path)
            else:
                raise Exception(f"Folder tidak ditemukan: {folder_path}")
        except Exception as e:
            self.show_message_box("Peringatan", f"{str(e)}", QtWidgets.QMessageBox.Icon.Warning)

    def delete_all_files(self, treeWidget, files_list):
        try:
            reply = QtWidgets.QMessageBox.question(
                self, 'Konfirmasi Hapus Semua',
                'Apakah Anda yakin ingin menghapus semua file?',
                QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
                QtWidgets.QMessageBox.StandardButton.No
            )
            if reply == QtWidgets.QMessageBox.StandardButton.Yes:
                treeWidget.clear()
                files_list.clear()
        except Exception as e:
            self.show_message_box("Peringatan", f"{str(e)}", QtWidgets.QMessageBox.Icon.Warning)

    def delete_selected_files(self, treeWidget, files_list):
        try:
            selected_items = treeWidget.selectedItems()
            if selected_items:
                reply = QtWidgets.QMessageBox.question(
                    self, 'Konfirmasi Hapus',
                    'Apakah Anda yakin ingin menghapus file yang dipilih?',
                    QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
                    QtWidgets.QMessageBox.StandardButton.No
                )
                if reply == QtWidgets.QMessageBox.StandardButton.Yes:
                    for item in selected_items:
                        file_path = item.toolTip(0)
                        index = next((i for i, f in enumerate(files_list) if f == file_path), None)
                        if index is not None:
                            files_list.pop(index)
                        treeWidget.takeTopLevelItem(treeWidget.indexOfTopLevelItem(item))
            else:
                raise Exception("Tidak ada file yang dipilih.")
        except Exception as e:
            self.show_message_box("Peringatan", f"{str(e)}", QtWidgets.QMessageBox.Icon.Warning)

    # Metode drag drop file dari Utilities\dragndrop_files_func.py
    dragEnterEvent = dragEnterEvent
    dragMoveEvent = dragMoveEvent
    dropEvent_convert_xls_xlsx = dropEvent_convert_xls_xlsx
    fileExists_convert_xls_xlsx = fileExists_convert_xls_xlsx

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = BatchConvertApp()
    window.show()
    sys.exit(app.exec())
