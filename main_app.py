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

        self.ui.pilihWarnaSuhu1_btn.clicked.connect(self.pilih_warna_suhu1)  # Menghubungkan tombol pilih warna suhu 1
        self.ui.pilihWarnaSuhu2_btn.clicked.connect(self.pilih_warna_suhu2)  # Menghubungkan tombol pilih warna suhu 2
        self.ui.pilihWarnaHoldingTime_btn.clicked.connect(self.pilih_warna_holding_time)  # Menghubungkan tombol pilih warna holding time
        self.ui.simpanPengaturan_btn.clicked.connect(self.simpan_pengaturan_conf)  # Menghubungkan tombol simpan konfigurasi

        # Context menu untuk listItems_filesSource
        self.ui.listItems_filesSource.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.ui.listItems_filesSource.customContextMenuRequested.connect(self.open_source_context_menu)
        
        # Context menu untuk listItems_outputFilesXLSX
        self.ui.listItems_outputFilesXLSX.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.ui.listItems_outputFilesXLSX.customContextMenuRequested.connect(self.open_output_context_menu)

        # Variabel untuk melacak pilihan menimpa otomatis
        self.auto_overwrite = False

        # Membuat objek untuk memberikan warna background kuning dan merah
        self.fill_kuning = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        # self.fill_merah = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

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

    # Metode dari Utilities/windows_func.py
    browse_files = browse_files
    validate_and_start_import = validate_and_start_import
    open_selected_output_folder = open_selected_output_folder
    open_selected_source_file = open_selected_source_file
    open_selected_source_folder = open_selected_source_folder
    open_output_context_menu = open_output_context_menu
    open_source_context_menu = open_source_context_menu
    open_selected_output_file = open_selected_output_file
    open_output_folder = open_output_folder

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
    dragEnterEvent = dragEnterEvent
    dragMoveEvent = dragMoveEvent
    dropEvent = dropEvent
    fileExists = fileExists
    
    # Metode dari Utilities/pengaturan_func.py
    simpan_pengaturan_conf = simpan_pengaturan_conf
    baca_pengaturan_conf = baca_pengaturan_conf
    get_batas_suhu_1 = get_batas_suhu_1
    get_batas_suhu_2 = get_batas_suhu_2
    pilih_warna_suhu1 = pilih_warna_suhu1
    pilih_warna_suhu2 = pilih_warna_suhu2
    pilih_warna_holding_time = pilih_warna_holding_time

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)

    # Muat dan terapkan stylesheet jika ada
    # with open("UI/styles.qss", "r") as style_file:
    #     app.setStyleSheet(style_file.read())

    window = MainWindow()
    window.show()
    sys.exit(app.exec())
