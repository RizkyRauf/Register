import tkinter as tk
from tkinter import messagebox
import os
from lib.register import Tree

class RegisterApp:
    """Class untuk aplikasi pendaftaran."""

    def __init__(self, master):
        """Inisialisasi aplikasi pendaftaran."""
        self.master = master
        master.title("Registrasi Tree")

        # Label dan input field untuk Nomor KTP
        self.label_ktp = tk.Label(master, text="Masukkan Nomor KTP:")
        self.label_ktp.grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        self.entry_ktp = tk.Entry(master)
        self.entry_ktp.grid(row=0, column=1, padx=10, pady=5)

        # Label dan input field untuk Nomor KK
        self.label_kk = tk.Label(master, text="Masukkan Nomor KK:")
        self.label_kk.grid(row=1, column=0, padx=10, pady=5, sticky=tk.W)
        self.entry_kk = tk.Entry(master)
        self.entry_kk.grid(row=1, column=1, padx=10, pady=5)
        
        # Label dan input field untuk Nomor ICCID
        self.label_iccid = tk.Label(master, text="Masukkan ICCID:")
        self.label_iccid.grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
        self.entry_iccid = tk.Entry(master)
        self.entry_iccid.grid(row=2, column=1, padx=10, pady=5)

        # Tombol submit
        self.button_submit = tk.Button(master, text="Submit", command=self.submit)
        self.button_submit.grid(row=3, column=0, columnspan=2, pady=10)

    def submit(self):
        """Fungsi yang dipanggil saat tombol submit ditekan."""
        # Ambil nilai KTP dan KK dari input fields
        ktp = self.entry_ktp.get()
        kk = self.entry_kk.get()
        iccid = self.entry_iccid.get()
        
        # Pastikan nilai KTP dan KK tidak kosong
        if ktp and kk:
            # Tentukan path ke Chrome driver
            chrome_driver_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "../driver/chromedriver.exe"))

            # Jalankan proses registrasi
            result = self.register(chrome_driver_path, ktp, kk, iccid)
            if result:
                messagebox.showinfo("Registrasi Berhasil", "Registrasi berhasil!")
            else:
                messagebox.showerror("Registrasi Gagal", "Registrasi gagal. Terjadi kesalahan.")
        else:
            messagebox.showerror("Kesalahan", "Harap masukkan Nomor KTP, Nomor KK dan ICCID. Tidak boleh kosong.")

    def register(self, chrome_driver_path, ktp, kk, iccid):
        """Fungsi untuk melakukan registrasi."""
        # Inisialisasi objek Tree
        my_tree = Tree()

        # Baca nomor telepon dari file Excel
        nomor_telepon = my_tree.read_excel()

        if nomor_telepon:
            # Jalankan registrasi dengan nomor telepon yang dibaca
            result = my_tree.run_register(chrome_driver_path, ktp, kk, iccid)
            
            # Cek apakah registrasi sukses dan tidak ada pesan kesalahan
            if result:
                # Tambahkan nomor telepon ke file Excel
                my_tree.update_excel(nomor_telepon, ktp, kk, iccid)
                return True
            else:
                messagebox.showerror("Registrasi Gagal", "Registrasi gagal. Terjadi kesalahan.")
                return False
        else:
            messagebox.showerror("Kesalahan", "Tidak ada nomor telepon yang tersedia dalam file Excel.")
            return False