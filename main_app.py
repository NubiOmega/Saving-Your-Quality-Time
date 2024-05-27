import sys
from PyQt6 import QtCore, QtWidgets, QtGui
from UI.ui_main_app import Ui_MainWindow  # Mengimpor kelas UI yang telah dihasilkan
from datetime import datetime, time
from openpyxl.styles import PatternFill
from openpyxl.utils.cell import get_column_letter, column_index_from_string
from batchconvert import BatchConvertApp
from compare_xls_xlsx import CompareXlsXlsxWindow
from Utilities.windows_func import *
from App.impor_proses_data import *
from App.salin_data_waktu_suhu import *
from App.validasi_data_waktu import *
from Utilities.dragndrop_files_func import *
from Utilities.pengaturan_func import *
from configparser import ConfigParser
from PyQt6.QtWidgets import QColorDialog  # Mengimpor QColorDialog untuk memilih warna


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.BrowseFile_btn.clicked.connect(self.browse_files)
        self.ui.progressBar.setValue(0)  # Mengatur nilai default progress bar menjadi 0
        self.ui.openFileFolder_btn.clicked.connect(self.open_output_folder)
        self.ui.isidansalinData_btn.clicked.connect(self.validate_and_start_import)
        self.ui.actionBatch_Convert_XLS_to_XLSX.triggered.connect(self.buka_window_batch_convert)
        self.ui.actionValidation_XLS_XLSX.triggered.connect(self.buka_window_compare_xls_xlsx)

        self.ui.pilihWarnaHoldingTime_btn.clicked.connect(self.pilih_warna_holding_time)  # Menghubungkan tombol pilih warna
        self.ui.simpanPengaturan_btn.clicked.connect(self.simpan_pengaturan_conf)  # Menghubungkan tombol simpan

        # Context menu untuk listItems_filesSource
        self.ui.listItems_filesSource.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.ui.listItems_filesSource.customContextMenuRequested.connect(self.open_source_context_menu)
        
        # Context menu untuk listItems_outputFilesXLSX
        self.ui.listItems_outputFilesXLSX.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.ui.listItems_filesSource.customContextMenuRequested.connect(self.open_output_context_menu)

        # Variabel untuk melacak pilihan menimpa otomatis
        self.auto_overwrite = False
        # Membuat objek untuk memberikan warna background kuning dan merah
        self.fill_kuning = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        self.fill_merah = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        # Menghubungkan QTextEdit yang sudah ada dengan fungsi drag and drop
        self.ui.textEditDragDropFiles.setAcceptDrops(True)
        self.ui.textEditDragDropFiles.dragEnterEvent = self.dragEnterEvent
        self.ui.textEditDragDropFiles.dragMoveEvent = self.dragMoveEvent
        self.ui.textEditDragDropFiles.dropEvent = self.dropEvent

        # Membaca pengaturan konfigurasi dari file saat aplikasi dibuka
        self.baca_pengaturan_conf()

    def buka_window_batch_convert(self):
        self.window_batch_convert = BatchConvertApp()
        self.window_batch_convert.show()

    def buka_window_compare_xls_xlsx(self):
        self.window_compare_xls_xlsx = CompareXlsXlsxWindow()
        self.window_compare_xls_xlsx.show()

    def simpan_pengaturan_conf(self):
        # Ambil nilai dari QDoubleSpinBox
        batas_suhu_1 = self.ui.batasSuhu1DoubleSpinBox.value()
        batas_suhu_2 = self.ui.batasRangeSuhu2EG710DerajatDoubleSpinBox.value()
        
        # Ambil warna holding time dari QLabel
        warna_holding_time = self.ui.fillWarnaHoldingTime_label.styleSheet().split("background-color:")[1].strip().replace(";", "")

        # Simpan nilai ke file konfigurasi
        config = ConfigParser()
        config['Pengaturan Suhu'] = {
            'Suhu1': batas_suhu_1,
            'Suhu2': batas_suhu_2
        }
        config['Pengaturan Fill Warna'] = {
            'FillSuhu1': '#FFFF00',
            'FillSuhu2': '#FF0000',
            'FillHoldingTime': warna_holding_time
        }
        with open('pengaturan_app.conf', 'w') as configfile:
            config.write(configfile)
        
        # Tampilkan pesan konfirmasi
        QtWidgets.QMessageBox.information(self, "Berhasil Disimpan", "Konfigurasi Pengaturan telah disimpan.")

    def baca_pengaturan_conf(self):
        config = ConfigParser()
        try:
            config.read('pengaturan_app.conf')
            if 'Pengaturan Suhu' in config:
                batas_suhu_1 = config.getfloat('Pengaturan Suhu', 'Suhu1')
                batas_suhu_2 = config.getfloat('Pengaturan Suhu', 'Suhu2')
                self.ui.batasSuhu1DoubleSpinBox.setValue(batas_suhu_1)
                self.ui.batasRangeSuhu2EG710DerajatDoubleSpinBox.setValue(batas_suhu_2)
            if 'Pengaturan Fill Warna' in config:
                fill_suhu1 = config.get('Pengaturan Fill Warna', 'FillSuhu1')
                fill_suhu2 = config.get('Pengaturan Fill Warna', 'FillSuhu2')
                fill_holding_time = config.get('Pengaturan Fill Warna', 'FillHoldingTime', fallback='#FFFFFF')
                # Tampilkan warna holding time pada label
                self.ui.fillWarnaHoldingTime_label.setStyleSheet(f"background-color: {fill_holding_time};")
        except Exception as e:
            print(f'Error membaca file pengaturan: {e}')

    def pilih_warna_holding_time(self):
        color = QColorDialog.getColor()
        if color.isValid():
            hex_color = color.name()  # Mendapatkan warna dalam format hex
            # Set warna yang dipilih ke label
            self.ui.fillWarnaHoldingTime_label.setStyleSheet(f"background-color: {hex_color};")
        
    # Metode dari Utilities/windows_func.py
    def browse_files(self):
        browse_files(self)

    def validate_and_start_import(self):
        validate_and_start_import(self)

    def open_selected_output_folder(self):
        open_selected_output_folder(self)

    def open_selected_source_file(self):
        open_selected_source_file(self)

    def open_selected_source_folder(self):
        open_selected_source_folder(self)

    def open_output_context_menu(self):
        open_output_context_menu(self)

    def open_source_context_menu(self):
        open_source_context_menu(self)

    def open_selected_output_file(self):
        open_selected_output_file(self)

    def open_output_folder(self):
        open_output_folder(self)

    # Metode dari App/impor_proses_data.py
    def start_import_excel(self):
        start_import_excel(self)

    def confirm_continue(self):
        return confirm_continue(self)

    def import_excel_with_progress(self, file_path, failed_files):
        import_excel_with_progress(self, file_path, failed_files)

    def salin_data_waktu_suhu_71_kesheet_DATA(self, ws, ws_data):
        salin_data_waktu_suhu_71_kesheet_DATA(self, ws, ws_data)

    def cek_waktu(self, ws_data, total_rows):
        cek_waktu(self, ws_data, total_rows)

    # Metode dari Utilities/dragndrop_files_func.py
    def dragEnterEvent(self, event):
        dragEnterEvent(event)

    def dragMoveEvent(self, event):
        dragMoveEvent(event)

    def dropEvent(self, event):
        dropEvent(event)

    def fileExists(self, file_path):
        return fileExists(file_path)

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)

    # Muat dan terapkan stylesheet jika ada
    # with open("UI/styles.qss", "r") as style_file:
    #     app.setStyleSheet(style_file.read())

    window = MainWindow()
    window.show()
    sys.exit(app.exec())