# Registrasi Tree

Registrasi Tree adalah sebuah aplikasi sederhana yang memungkinkan pengguna untuk melakukan registrasi menggunakan data dari file Excel ke situs web registrasi. Aplikasi ini menggunakan Selenium untuk mengotomatisasi proses registrasi menggunakan browser web.

## Struktur Data
```
├── Data
│   └── daftar.xlxs
├── driver
│   └── chromedriver.exe
├── lib
│   ├── __init__.py
│   ├── gui.py
│   └── register.py
├── main.py
└── requirements.txt
```
Struktur data aplikasi ini terdiri dari beberapa direktori dan file:

- **Data**: Direktori untuk menyimpan file Excel yang berisi data untuk proses registrasi.
- **Driver**: Direktori untuk menyimpan file `chromedriver.exe` yang diperlukan oleh Selenium WebDriver.
- **Lib**: Direktori untuk menyimpan modul-modul Python yang digunakan dalam aplikasi.
- **main.py**: File utama yang digunakan untuk menjalankan aplikasi.
- **requirements.txt**: File yang berisi daftar paket Python yang diperlukan untuk menjalankan aplikasi.

## Cara Penggunaan

1. Pastikan Anda memiliki Python dan pip terinstal di komputer Anda.

2. Instal paket-paket yang diperlukan dengan menjalankan perintah berikut di terminal atau command prompt:
```
pip install -r requirements.txt
```
3. Letakkan file Excel yang berisi data untuk registrasi di dalam direktori `Data`. Pastikan file tersebut memiliki nama `daftar.xlsx`.

4. Letakkan file `chromedriver.exe` di dalam direktori `driver`.

5. Jalankan aplikasi dengan menjalankan perintah berikut di terminal atau command prompt:
```
python main.py
```
6. Aplikasi akan membuka jendela GUI yang memungkinkan Anda untuk memasukkan nomor KTP, nomor KK, dan ICCID.

7. Masukkan nomor KTP, nomor KK, dan ICCID yang sesuai ke dalam input fields.

8. Klik tombol "Submit" untuk memulai proses registrasi.

9. Tunggu sampai aplikasi menampilkan pesan apakah registrasi berhasil atau gagal.

10. Jika registrasi berhasil, pesan "Registrasi Berhasil" akan ditampilkan. Jika registrasi gagal, pesan "Registrasi Gagal" akan ditampilkan bersama dengan pesan kesalahan yang terkait.

## Catatan

- Pastikan Anda memiliki koneksi internet yang stabil saat menjalankan aplikasi ini, karena aplikasi akan melakukan registrasi secara online menggunakan browser web.
- Pastikan Anda memiliki akun dengan hak akses yang sesuai untuk melakukan registrasi di situs web yang dituju.
