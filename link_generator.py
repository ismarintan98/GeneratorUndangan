import os
import time

import pandas as pd


def link_gen(nama_tamu, Nomor):

    with open("template _text.txt", "r") as file:
        isi = file.read()
        # print(isi)

        nama_tamu_temp = nama_tamu
        link_main = "https://satumomen.com/khalisa-dan-marin?to="

        # check if nama tamu contains space replace with %20
        if " " in nama_tamu:
            nama_tamu = nama_tamu.replace(" ", "%20")

        link_plus_name = link_main + nama_tamu

        # print(link_plus_name)
        # Ganti placeholder dengan data
        isi_undangan = isi.format(
            nama_tamu=nama_tamu_temp, link_undangan=link_plus_name
        )

        # Tampilkan atau simpan hasil
        # print(isi_undangan)

        # save path to Hasil/Nomor_Nama.txt
        save_path = f"Hasil/{Nomor}_{nama_tamu_temp}.txt"
        with open(save_path, "w") as file:
            file.write(isi_undangan)
            print(f"File {save_path} berhasil dibuat.")


print("~~~ Halo Ola, The Moon is beautiful, isn't it? ~~~")
time.sleep(3)
print("-----Program Link Generator Undangan Pernikahan Khalisa dan Marin-----")
print("-----Created by : Moh Ismarintan Zazuli -----")
print("-----ismarintan@its.ac.id-----")
# delay 1 detik
time.sleep(1)
print("Loading Database...")
# delay 1 detik
time.sleep(1)
# # load data from excel
try:
    df = pd.read_excel("Database/db.xlsx")
    print("-> Database berhasil dimuat.")
except FileNotFoundError:
    print("-> File Database/db.xlsx tidak ditemukan.")
    exit()

try:
    nama = df["Nama"].tolist()
    nomor = df["No"].tolist()
except KeyError:
    print("-> Kolom 'Nama' atau 'No' tidak ditemukan di database.")
    exit()
print("-> Kolom 'Nama' dan 'No' ditemukan di database.")
print("-> Jumlah data yang ditemukan : ", len(nama))

print("Apakah Anda ingin melanjutkan? (y/n)")
jawaban = input("-> ")
if jawaban.lower() != "y":
    print("-> Proses dibatalkan.")
    exit()
print("-> Proses dilanjutkan.")

# cek apakah folder Hasil ada
if not os.path.exists("Hasil"):
    os.makedirs("Hasil")
    print("-> Folder Hasil berhasil dibuat.")
else:
    print("-> Folder Hasil sudah ada.")
    print("-> Semua file di folder Hasil akan ditimpa.")
    time.sleep(1)
print("-> Proses pembuatan link undangan dimulai...")

# cek apakah file template_text.txt ada
if not os.path.exists("template _text.txt"):
    print("-> File template_text.txt tidak ditemukan.")
    exit()
else:
    print("-> File template_text.txt ditemukan.")
    time.sleep(1)
print("-> Proses pembuatan file undangan dimulai...")
# delay 1 detik
time.sleep(1)

print()
print()

for i in range(len(nama)):
    link_gen(nama[i], nomor[i])

print()
print()

print("-> Proses pembuatan file undangan selesai.")
print("-> Semua file undangan sudah disimpan di folder Hasil.")
print()
print("~~~ Ola Bon Bon Mentul Mentul ~~~")
print()
print("Tekan q untuk keluar.")
while True:
    if input() == "q":
        break
