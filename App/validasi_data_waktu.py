from PyQt6 import QtCore, QtWidgets
from datetime import datetime, time
from configparser import ConfigParser

# pengecekan dan pemberian warna background pada total waktu yang dibawah dari batas holding time yang diatur

def cek_waktu(self, ws_data, total_data_rows):
    # print("Metode cek_waktu dipanggil")  # Debug print untuk memastikan metode dipanggil
    # # Membaca pengaturan holding time dari file konfigurasi
    config = ConfigParser()
    config.read('pengaturan_app.conf')
    holding_time_str = config.get('Pengaturan Holding Time', 'HoldingTime', fallback='9:00')

    # Konversi holding_time_str dari format menit:detik ke detik
    holding_time_parts = holding_time_str.split(':')
    holding_time_limit = int(holding_time_parts[0]) * 60 + int(holding_time_parts[1])

    # # Cetak holding_time_limit ke console untuk pengecekan
    # print(f"Holding Time Limit (detik): {holding_time_limit}")

    # # Cetak fill_holding_time ke console untuk pengecekan
    # print(f"Fill Holding Time: {self.fill_holding_time}")

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

                if selisih_waktu < holding_time_limit:  # Kurang dari holding time yang ditentukan
                    ws_data.cell(row=row, column=col).fill = self.fill_holding_time
                    ws_data.cell(row=row, column=col + 1).fill = self.fill_holding_time
                else:
                    ws_data.cell(row=row, column=col).fill = self.fill_suhu2
                    ws_data.cell(row=row, column=col + 1).fill = self.fill_suhu2
                

# ===================================================================================