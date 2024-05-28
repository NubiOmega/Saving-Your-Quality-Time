import os
import sys
import win32com.client as win32
from PyQt6 import QtCore, QtGui, QtWidgets
from UI.ui_batchxlstoxlsx import Ui_MainWindow

class BatchConvertApp(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        
        self.LokasiSumberFolder_btn.clicked.connect(self.select_source_folder)
        self.LokasiOutputFolder_btn.clicked.connect(self.select_output_folder)
        self.konversi_Btn.clicked.connect(self.start_conversion)
        self.openFolderOutputXLSX_btn.clicked.connect(self.open_output_folder)

        self.source_folder = ""
        self.output_folder = ""
        self.progressBar.setValue(0)

        # Menambahkan event handler untuk klik kanan
        self.listFileItems_lokasiSumber.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.listFileItems_lokasiSumber.customContextMenuRequested.connect(self.show_source_context_menu)
        
        self.listFileItems_lokasiTujuan.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.listFileItems_lokasiTujuan.customContextMenuRequested.connect(self.show_output_context_menu)

    def select_source_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Pilih Folder Sumber")
        if folder:
            self.source_folder = folder
            self.textBrowser_lokasiSumber.setText(folder)
            self.listFileItems_lokasiSumber.clear()
            for filename in os.listdir(folder):
                if filename.endswith('.xls'):
                    self.listFileItems_lokasiSumber.addItem(filename)

    def select_output_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Pilih Folder Output")
        if folder:
            self.output_folder = folder
            self.textBrowser_lokasiTujuan.setText(folder)

    def open_output_folder(self):
        if self.output_folder:
            os.startfile(self.output_folder)
        else:
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Folder output belum dipilih.")

    def convert_xls_to_xlsx(self, source_file, dest_file):
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        try:
            wb = excel.Workbooks.Open(source_file)
            wb.SaveAs(dest_file, FileFormat=51)  # 51 adalah format untuk xlsx
            wb.Close()
            return True, None
        except Exception as e:
            return False, str(e)
        finally:
            excel.Application.Quit()

    def start_conversion(self):
        if not self.source_folder or not self.output_folder:
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Silakan pilih folder sumber dan output terlebih dahulu.")
            return

        total_files = self.listFileItems_lokasiSumber.count()
        if total_files == 0:
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Tidak ada file .xls di folder sumber.")
            return
        
        reply = QtWidgets.QMessageBox.question(
            self, 'Konfirmasi Menimpa File',
            'Apakah Anda ingin menimpa semua file yang sudah ada di folder output?',
            QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
            QtWidgets.QMessageBox.StandardButton.No
        )
        overwrite_all = reply == QtWidgets.QMessageBox.StandardButton.Yes

        self.progressBar.setValue(0)
        successful_conversions = 0

        for index in range(total_files):
            filename = self.listFileItems_lokasiSumber.item(index).text()
            source_file = os.path.abspath(os.path.join(self.source_folder, filename))
            dest_file = os.path.abspath(os.path.join(self.output_folder, filename.replace('.xls', '.xlsx')))

            if os.path.exists(dest_file):
                if not overwrite_all:
                    continue
                else:
                    os.remove(dest_file)

            success, error = self.convert_xls_to_xlsx(source_file, dest_file)
            if success:
                self.listFileItems_lokasiTujuan.addItem(filename.replace('.xls', '.xlsx'))
                successful_conversions += 1
            else:
                self.show_message_box('Gagal', f'Gagal mengonversi {source_file}. Error: {error}', QtWidgets.QMessageBox.Icon.Critical)
            self.progressBar.setValue(int((index + 1) / total_files * 100))

        if successful_conversions == total_files:
            self.show_message_box('Selesai', 'Semua file berhasil dikonversi.', QtWidgets.QMessageBox.Icon.Information)
        else:
            self.show_message_box('Selesai', f'Proses konversi selesai. {successful_conversions} dari {total_files} file berhasil dikonversi.', QtWidgets.QMessageBox.Icon.Information)

    def show_message_box(self, title, message, icon):
        msg_box = QtWidgets.QMessageBox()
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setIcon(icon)
        msg_box.exec()

    def show_source_context_menu(self, position):
        menu = QtWidgets.QMenu()
        open_file_action = menu.addAction("Open File")
        open_folder_action = menu.addAction("Open Folder")
        action = menu.exec(self.listFileItems_lokasiSumber.mapToGlobal(position))
        
        if action == open_file_action:
            self.open_selected_file(self.listFileItems_lokasiSumber)
        elif action == open_folder_action:
            self.open_selected_folder(self.listFileItems_lokasiSumber, self.source_folder)

    def show_output_context_menu(self, position):
        menu = QtWidgets.QMenu()
        open_file_action = menu.addAction("Open File")
        open_folder_action = menu.addAction("Open Folder")
        action = menu.exec(self.listFileItems_lokasiTujuan.mapToGlobal(position))
        
        if action == open_file_action:
            self.open_selected_file(self.listFileItems_lokasiTujuan)
        elif action == open_folder_action:
            self.open_selected_folder(self.listFileItems_lokasiTujuan, self.output_folder)

    def open_selected_file(self, list_widget):
        selected_item = list_widget.currentItem()
        if selected_item:
            file_path = os.path.join(self.source_folder if list_widget == self.listFileItems_lokasiSumber else self.output_folder, selected_item.text())
            if os.path.exists(file_path):
                os.startfile(file_path)
            else:
                QtWidgets.QMessageBox.warning(self, "Peringatan", "File tidak ditemukan.")

    def open_selected_folder(self, list_widget, base_folder):
        selected_item = list_widget.currentItem()
        if selected_item:
            file_path = os.path.join(base_folder, selected_item.text())
            if os.path.exists(file_path):
                os.startfile(os.path.dirname(file_path))
            else:
                QtWidgets.QMessageBox.warning(self, "Peringatan", "Folder tidak ditemukan.")

# if __name__ == "__main__":
#     app = QtWidgets.QApplication(sys.argv)
#     window = BatchConvertApp()
#     window.show()
#     sys.exit(app.exec())
