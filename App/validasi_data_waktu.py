from PyQt6 import QtCore, QtWidgets
from datetime import datetime, time

# pengecekan dan pemberian warna background pada total waktu yang dibawah dari 9 menit

def cek_waktu(self, ws_data, total_data_rows):

    for col in range(2, 70, 2):  # Loop untuk kolom B, D, F, H, J, ..., BR
        waktu_min = None
        waktu_max = None

        for row in range(15, total_data_rows + 1):
            waktu = ws_data.cell(row=row, column=col).value

            if waktu is None:
                continue

            try:
                if isinstance(waktu, str):
                    try:
                        waktu = datetime.strptime(waktu, "%H:%M:%S").time()
                    except ValueError:
                        waktu = datetime.strptime(waktu, "%I:%M:%S %p").time()
                elif isinstance(waktu, datetime):
                    waktu = waktu.time()
                elif not isinstance(waktu, time):
                    continue
            except ValueError as ve:
                QtWidgets.QMessageBox.critical(self, "Error", f"Kesalahan parsing waktu pada baris {row}, kolom {col}: {ve}")
                continue

            if waktu_min is None or waktu < waktu_min:
                waktu_min = waktu

            if waktu_max is None or waktu > waktu_max:
                waktu_max = waktu

        if waktu_min is not None and waktu_max is not None:
            selisih_waktu = (datetime.combine(datetime.today(), waktu_max) - datetime.combine(datetime.today(), waktu_min)).total_seconds()

            for row in range(15, total_data_rows + 1):
                waktu = ws_data.cell(row=row, column=col).value

                if waktu is None:
                    continue

                if selisih_waktu < 540:  # Kurang dari 9 menit
                    ws_data.cell(row=row, column=col).fill = self.fill_merah
                    ws_data.cell(row=row, column=col + 1).fill = self.fill_merah
                else:
                    ws_data.cell(row=row, column=col).fill = self.fill_kuning
                    ws_data.cell(row=row, column=col + 1).fill = self.fill_kuning

# ===================================================================================