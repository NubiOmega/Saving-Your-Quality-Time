
from PyQt6.QtWidgets import QColorDialog

def get_batas_suhu_1(self):
    return self.ui.batasSuhu1DoubleSpinBox.value()

def get_batas_suhu_2(self):
    return self.ui.batasRangeSuhu2EG710DerajatDoubleSpinBox.value()

def pilih_warna(self):
    color = QColorDialog.getColor()
    if color.isValid():
        return color.name()  # Mengembalikan warna dalam format hex
    return None

