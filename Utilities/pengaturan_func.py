
from PyQt6.QtWidgets import QColorDialog

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

