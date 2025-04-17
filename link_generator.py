import os
import time
import urllib.parse

import pandas as pd
from fpdf import FPDF

name_list = []
link_list = []


def link_gen(nama_tamu, Nomor, no_wa):
    with open(
        "template_text.txt", "r", encoding="utf-8"
    ) as file:  # Ensure we read with UTF-8 encoding
        isi = file.read()

        nama_tamu_temp = nama_tamu
        link_main = "https://satumomen.com/khalisa-dan-marin?to="

        # Check if nama tamu contains space and replace with %20
        if " " in nama_tamu:
            nama_tamu = nama_tamu.replace(" ", "%20")

        link_plus_name = link_main + nama_tamu

        link_wa = "https://wa.me/" + str(no_wa)

        # Replace placeholders with actual data from the template
        isi_undangan = isi.format(
            nama_tamu=nama_tamu_temp, link_undangan=link_plus_name
        )

        # URL encode the rest of the message (but NOT the emojis)
        pesaan_wa = urllib.parse.quote(
            isi_undangan, safe=":/?&="
        )  # Keep the special URL characters safe
        isi_undangan = f"{link_wa}?text={pesaan_wa}"

        print(isi_undangan)  # This should now show the proper emojis

        name_list.append(nama_tamu_temp)
        link_list.append(isi_undangan)


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

# Load data from excel
try:
    df = pd.read_excel("Database/db.xlsx")
    print("-> Database berhasil dimuat.")
except FileNotFoundError:
    print("-> File Database/db.xlsx tidak ditemukan.")
    exit()

try:
    nama = df["Nama"].tolist()
    nomor = df["No"].tolist()
    telepon = df["Nomor Telepon"].tolist()
except KeyError:
    print("-> Kolom 'Nama' atau 'No' atau Nomor Telepon tidak ditemukan di database.")
    exit()
print("-> Kolom 'Nama' dan 'No' dan Nomor Telepon ditemukan di database.")
print("-> Jumlah data yang ditemukan : ", len(nama))

# print(telepon)
# exit()


print("Apakah Anda ingin melanjutkan? (y/n)")
jawaban = input("-> ")
if jawaban.lower() != "y":
    print("-> Proses dibatalkan.")
    exit()
print("-> Proses dilanjutkan.")

# Check if folder Hasil exists
if not os.path.exists("Hasil"):
    os.makedirs("Hasil")
    print("-> Folder Hasil berhasil dibuat.")
else:
    print("-> Folder Hasil sudah ada.")
    print("-> Semua file di folder Hasil akan ditimpa.")
    time.sleep(1)
print("-> Proses pembuatan link undangan dimulai...")

# Check if template_text.txt exists
if not os.path.exists("template_text.txt"):
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
    link_gen(nama[i], nomor[i], telepon[i])

# Generate name and hyperlink in each name in the PDF
pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", size=12)
pdf.cell(200, 10, txt="Link Undangan Pernikahan Khalisa dan Marin", ln=1, align="C")
pdf.cell(200, 10, txt="Daftar Nama dan Link Undangan", ln=1, align="C")

for i in range(len(name_list)):
    pdf.cell(
        200,
        10,
        txt=f"{nomor[i]}. {name_list[i]}",
        ln=True,
        link=link_list[i],
    )

pdf.output("Hasil/Daftar_Nama_dan_Link_Undangan.pdf")
print("-> File Daftar_Nama_dan_Link_Undangan.pdf berhasil dibuat.")

exit()


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
