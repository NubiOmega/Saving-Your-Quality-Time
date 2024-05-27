
from PyQt6 import QtCore, QtWidgets
from PyQt6.QtWidgets import QColorDialog
from openpyxl.styles import PatternFill
from configparser import ConfigParser

def simpan_pengaturan_conf(self):
    # Ambil nilai dari QDoubleSpinBox
    batas_suhu_1 = self.ui.batasSuhu1_DoubleSpinBox.value()
    batas_suhu_2 = self.ui.batasRangeSuhu2_DoubleSpinBox.value()
    
    # Ambil warna suhu dari QLabel
    warna_suhu1 = self.ui.pilihWarnaSuhu1_btn.styleSheet().split("background-color:")[1].strip().replace(";", "")
    warna_suhu2 = self.ui.pilihWarnaSuhu2_btn.styleSheet().split("background-color:")[1].strip().replace(";", "")
    # Ambil warna holding time dari QLabel
    warna_holding_time = self.ui.pilihWarnaHoldingTime_btn.styleSheet().split("background-color:")[1].strip().replace(";", "")
    
    # Ambil nilai holding time dari QLineEdit
    holding_time = self.ui.holdingTime_lineEdit.text()

    # Simpan nilai ke file konfigurasi
    config = ConfigParser()
    config['Pengaturan Suhu'] = {
        'Suhu1': batas_suhu_1,
        'Suhu2': batas_suhu_2
    }
    config['Pengaturan Fill Warna'] = {
        'FillSuhu1': warna_suhu1,
        'FillSuhu2': warna_suhu2,
        'FillHoldingTime': warna_holding_time
    }
    config['Pengaturan Holding Time'] = {
        'HoldingTime': holding_time
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
            self.ui.batasSuhu1_DoubleSpinBox.setValue(batas_suhu_1)
            self.ui.batasRangeSuhu2_DoubleSpinBox.setValue(batas_suhu_2)
        if 'Pengaturan Fill Warna' in config:
            fill_suhu1 = config.get('Pengaturan Fill Warna', 'Fillsuhu1', fallback='#FFFFFF')
            fill_suhu2 = config.get('Pengaturan Fill Warna', 'Fillsuhu2', fallback='#FFFFFF')
            fill_holding_time = config.get('Pengaturan Fill Warna', 'FillHoldingTime', fallback='#FFFFFF')
            # Tampilkan warna suhu 1 pada label
            self.ui.pilihWarnaSuhu1_btn.setStyleSheet(f"background-color: {fill_suhu1};")
            self.fill_suhu1 = PatternFill(start_color=fill_suhu1.lstrip('#'),
                                                    end_color=fill_suhu1.lstrip('#'),
                                                    fill_type="solid")
            # Tampilkan warna suhu 2 pada label
            self.ui.pilihWarnaSuhu2_btn.setStyleSheet(f"background-color: {fill_suhu2};")
            self.fill_suhu2 = PatternFill(start_color=fill_suhu2.lstrip('#'),
                                                    end_color=fill_suhu2.lstrip('#'),
                                                    fill_type="solid")
            # Tampilkan warna holding time pada label
            self.ui.pilihWarnaHoldingTime_btn.setStyleSheet(f"background-color: {fill_holding_time};")
            self.fill_holding_time = PatternFill(start_color=fill_holding_time.lstrip('#'),
                                                    end_color=fill_holding_time.lstrip('#'),
                                                    fill_type="solid")
        if 'Pengaturan Holding Time' in config:
            holding_time = config.get('Pengaturan Holding Time', 'HoldingTime', fallback='9:00')
            self.ui.holdingTime_lineEdit.setText(holding_time)
    except Exception as e:
        print(f'Error membaca file pengaturan: {e}')
        
def get_batas_suhu_1(self):
    return self.ui.batasSuhu1_DoubleSpinBox.value()

def get_batas_suhu_2(self):
    return self.ui.batasRangeSuhu2_DoubleSpinBox.value()

def pilih_warna_suhu1(self):
    color = QColorDialog.getColor()
    if color.isValid():
        hex_color = color.name()  # Mendapatkan warna dalam format hex
        # Set warna yang dipilih ke label
        self.ui.pilihWarnaSuhu1_btn.setStyleSheet(f"background-color: {hex_color};")

def pilih_warna_suhu2(self):
    color = QColorDialog.getColor()
    if color.isValid():
        hex_color = color.name()  # Mendapatkan warna dalam format hex
        # Set warna yang dipilih ke label
        self.ui.pilihWarnaSuhu2_btn.setStyleSheet(f"background-color: {hex_color};")

def pilih_warna_holding_time(self):
    color = QColorDialog.getColor()
    if color.isValid():
        hex_color = color.name()  # Mendapatkan warna dalam format hex
        # Set warna yang dipilih ke label
        self.ui.pilihWarnaHoldingTime_btn.setStyleSheet(f"background-color: {hex_color};")

