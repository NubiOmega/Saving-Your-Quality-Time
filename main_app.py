import sys
from PyQt6 import QtCore, QtWidgets
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

        # Context menu untuk listItems_outputFilesXLSX
        self.ui.listItems_outputFilesXLSX.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.ui.listItems_outputFilesXLSX.customContextMenuRequested.connect(self.open_output_context_menu)

        # Context menu untuk listItems_filesSource
        self.ui.listItems_filesSource.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.ui.listItems_filesSource.customContextMenuRequested.connect(self.open_source_context_menu)

        # Variabel untuk melacak pilihan menimpa otomatis
        self.auto_overwrite = False
        #membuat objek untuk memberikan warna background kuning dan merah
        self.fill_kuning = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        self.fill_merah = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

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

    # Metode dari App/imporprosesdata.py
    start_import_excel = start_import_excel
    confirm_continue = confirm_continue
    import_excel_with_progress = import_excel_with_progress
    salin_data_waktu_suhu_71_kesheet_DATA = salin_data_waktu_suhu_71_kesheet_DATA
    cek_waktu = cek_waktu

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

