import sys
import os
from PyQt6 import QtCore, QtGui, QtWidgets
from UI.ui_validasixlsdanxlsx import Ui_MainWindow
import pandas as pd
from PyQt6.QtCore import QStringListModel

class ValidationWorker(QtCore.QObject):
    progressChanged = QtCore.pyqtSignal(int)
    validationFinished = QtCore.pyqtSignal(pd.DataFrame, bool, str)

    def __init__(self, file_xls, file_xlsx, nama_sheet_xls, nama_sheet_xlsx):
        super().__init__()
        self.file_xls = file_xls
        self.file_xlsx = file_xlsx
        self.nama_sheet_xls = nama_sheet_xls
        self.nama_sheet_xlsx = nama_sheet_xlsx

    def run(self):
        try:
            self.progressChanged.emit(10)
            df_xls = pd.read_excel(self.file_xls, sheet_name=self.nama_sheet_xls)
            self.progressChanged.emit(40)
            df_xlsx = pd.read_excel(self.file_xlsx, sheet_name=self.nama_sheet_xlsx)
            self.progressChanged.emit(70)

            df_xls, df_xlsx = df_xls.align(df_xlsx, join='outer', axis=0)
            df_xls, df_xlsx = df_xls.align(df_xlsx, join='outer', axis=1)

            df_xls = df_xls.fillna('N/A')
            df_xlsx = df_xlsx.fillna('N/A')

            perbedaan = df_xls.compare(df_xlsx)
            self.progressChanged.emit(100)
            self.validationFinished.emit(perbedaan, True, "")
        except Exception as e:
            self.validationFinished.emit(pd.DataFrame(), False, str(e))

class CompareXlsXlsxWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        
        self.list_model = QtGui.QStandardItemModel(self)
        self.ui.listView_itemsOutputs.setModel(self.list_model)
        
        self.ui.cekValidasi_btn.clicked.connect(self.cek_validasi)
        self.ui.lokasiFolderSumberXLS.clicked.connect(self.pilih_folder_xls)
        self.ui.lokasiFolderSumberXLSX.clicked.connect(self.pilih_folder_xlsx)
        
        self.ui.listView_itemsOutputs.setContextMenuPolicy(QtCore.Qt.ContextMenuPolicy.CustomContextMenu)
        self.ui.listView_itemsOutputs.customContextMenuRequested.connect(self.tampilkan_menu_konteks)
        
        self.ui.detailDifferendData_btn.clicked.connect(self.tampilkan_rincian_perbedaan)
        
        # Inisialisasi progress bar ke 0 saat pertama kali aplikasi dijalankan
        self.ui.progressBar.setValue(0)
        
        self.show()

    def pilih_folder(self, jenis):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, f"Pilih Folder {jenis.upper()}")
        if folder:
            if jenis == 'xls':
                self.ui.text_lokasisumberfilesxls.setText(folder)
                self.tampilkan_file(folder, jenis)
                self.tampilkan_datasheet(jenis)
            else:
                self.ui.text_lokasitujuanfilesxlsx.setText(folder)
                self.tampilkan_file(folder, jenis)
                self.tampilkan_datasheet(jenis)
        else:
            QtWidgets.QMessageBox.warning(self, "Peringatan", f"Tidak ada folder yang dipilih. Silakan pilih folder yang berisi file {jenis.upper()}.")

    def pilih_folder_xls(self):
        self.pilih_folder('xls')

    def pilih_folder_xlsx(self):
        self.pilih_folder('xlsx')

    def tampilkan_file(self, folder, jenis):
        list_widget = self.ui.listItems_FilesSumber_xls if jenis == 'xls' else self.ui.listItems_FilesSumber_xlsx
        list_widget.clear()
        files = [file for file in os.listdir(folder) if file.endswith(f".{jenis}")]
        list_widget.addItems(files)

    def tampilkan_datasheet(self, jenis):
        folder = self.ui.text_lokasisumberfilesxls.toPlainText() if jenis == 'xls' else self.ui.text_lokasitujuanfilesxlsx.toPlainText()
        files = [file for file in os.listdir(folder) if file.endswith(f".{jenis}")]
        if files:
            nama_datasheet = self.baca_datasheet(os.path.join(folder, files[0]))
            combo_box = self.ui.comboBox_dtsheet_xls if jenis == 'xls' else self.ui.comboBox_dtsheet_xlsx
            combo_box.clear()
            combo_box.addItems(nama_datasheet)

    def baca_datasheet(self, file):
        try:
            return pd.ExcelFile(file).sheet_names
        except Exception as e:
            error_message = f"Terjadi kesalahan saat membaca file: {e}"
            self.list_model.appendRow(QtGui.QStandardItem(error_message))
            QtWidgets.QMessageBox.critical(self, "Error", error_message)
            return []

    def cek_validasi(self):
        file_xls_item = self.ui.listItems_FilesSumber_xls.currentItem()
        file_xlsx_item = self.ui.listItems_FilesSumber_xlsx.currentItem()
        nama_sheet_xls = self.ui.comboBox_dtsheet_xls.currentText()
        nama_sheet_xlsx = self.ui.comboBox_dtsheet_xlsx.currentText()

        if not (file_xls_item and file_xlsx_item and nama_sheet_xls and nama_sheet_xlsx):
            QtWidgets.QMessageBox.warning(self, "Peringatan", "Harap pilih file .xls dan .xlsx serta data sheet terlebih dahulu.")
            return

        file_xls = os.path.join(self.ui.text_lokasisumberfilesxls.toPlainText(), file_xls_item.text())
        file_xlsx = os.path.join(self.ui.text_lokasitujuanfilesxlsx.toPlainText(), file_xlsx_item.text())

        # Mengatur nilai awal progress bar
        self.ui.progressBar.setValue(0)

        self.progress = QtWidgets.QProgressDialog("Memvalidasi file...", "Batal", 0, 100, self)
        self.progress.setWindowModality(QtCore.Qt.WindowModality.WindowModal)
        self.progress.setMinimumDuration(0)

        self.worker = ValidationWorker(file_xls, file_xlsx, nama_sheet_xls, nama_sheet_xlsx)
        self.thread = QtCore.QThread()
        self.worker.moveToThread(self.thread)

        self.worker.progressChanged.connect(self.progress.setValue)
        self.worker.progressChanged.connect(self.ui.progressBar.setValue)
        self.worker.validationFinished.connect(self.on_validation_finished)
        self.worker.validationFinished.connect(self.thread.quit)
        self.worker.validationFinished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.started.connect(self.worker.run)
        self.thread.start()

    def on_validation_finished(self, perbedaan, success, error_message):
        if success:
            if perbedaan.empty:
                model = QStringListModel(["Kedua file memiliki konten data yang sama."])
                self.ui.listView_itemsOutputs.setModel(model)
                self.list_model.clear()
            else:
                model = QStringListModel(["Kedua file xls dan xlsx memiliki nilai data yang berbeda."])
                self.ui.listView_itemsOutputs.setModel(model)
            self.perbedaan = perbedaan
        else:
            self.list_model.appendRow(QtGui.QStandardItem(error_message))
            QtWidgets.QMessageBox.critical(self, "Error", error_message)

    def tampilkan_menu_konteks(self, pos):
        menu = QtWidgets.QMenu()
        detail_action = menu.addAction("Lihat Detail Perbedaan")
        copy_action = menu.addAction("Salin Teks")
        copy_action.triggered.connect(self.salinteks_item_listView)
        menu.exec(self.ui.listView_itemsOutputs.viewport().mapToGlobal(pos))

    def salinteks_item_listView(self):
        item_terpilih = self.ui.listView_itemsOutputs.currentIndex()
        teks_terpilih = self.list_model.itemData(item_terpilih)
        if teks_terpilih:
            QtWidgets.QApplication.clipboard().setText(teks_terpilih[0])

    def tampilkan_rincian_perbedaan(self):
        try:
            if hasattr(self, 'perbedaan') and not self.perbedaan.empty:
                self.ui.tableView_outputDiffData.setModel(None)
                model = PandasModel(self.perbedaan)
                self.ui.tableView_outputDiffData.setModel(model)
            else:
                QtWidgets.QMessageBox.information(self, "Informasi", "Tidak ada perbedaan yang ditemukan atau belum dilakukan validasi.")
        except Exception as e:
            error_message = (f"Terjadi kesalahan saat menampilkan rincian perbedaan: {e}")
            self.list_model.appendRow(QtGui.QStandardItem(error_message))
            QtWidgets.QMessageBox.critical(self, "Error", error_message)

class PandasModel(QtCore.QAbstractTableModel):
    def __init__(self, data):
        super(PandasModel, self).__init__()
        self._data = data

    def rowCount(self, parent=None):
        return len(self._data.values)

    def columnCount(self, parent=None):
        return self._data.columns.size

    def data(self, index, role=QtCore.Qt.ItemDataRole.DisplayRole):
        if index.isValid():
            if role == QtCore.Qt.ItemDataRole.DisplayRole or role == QtCore.Qt.ItemDataRole.EditRole:
                return str(self._data.values[index.row()][index.column()])
        return None

    def headerData(self, section, orientation, role):
        if orientation == QtCore.Qt.Orientation.Horizontal and role == QtCore.Qt.ItemDataRole.DisplayRole:
            return str(self._data.columns[section])
        elif orientation == QtCore.Qt.Orientation.Vertical and role == QtCore.Qt.ItemDataRole.DisplayRole:
            return str(self._data.index[section])
        return None

# if __name__ == "__main__":
#     app = QtWidgets.QApplication(sys.argv)
#     window = CompareXlsXlsxWindow()
#     sys.exit(app.exec())
