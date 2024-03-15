import logging
import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class Tree:
    
    def __init__(self, cookies_file=None):
        self.cookies_file = cookies_file
        
    def initialize_driver(self, chrome_driver_path):
        try:
            options = webdriver.ChromeOptions()
            options.add_argument('--disable-gpu')
            options.add_argument('--headless')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            options.add_argument('--disable-extensions')
            options.add_argument('--disable-blink-features=AutomationControlled')
            options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36')

            service = ChromeService(executable_path=chrome_driver_path)
            driver = webdriver.Chrome(service=service, options=options)
            return driver
        except Exception as e:
            logging.error(f"Error initializing driver: {e}")
            raise
        
    def read_excel(self):
        file_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "../data/daftar.xlsx"))
        try:
            # Gunakan pandas untuk membaca file Excel
            df = pd.read_excel(file_path)
            ktp_column = df['KTP']
            
            # Cek apakah kolom KTP sudah memiliki nilai
            if ktp_column.isnull().values.any():
                # Jika kolom KTP belum memiliki nilai, ambil nomor telepon yang sesuai
                nomor_telepon_column = df['Nomor_Telpon']
                nomor_telepon = str(nomor_telepon_column[ktp_column.isnull()].iloc[0])
                
                # Ubah format nomor telepon dari notasi ilmiah menjadi string biasa
                nomor_telepon = "{:.0f}".format(float(nomor_telepon.replace(',', '')))  # Menghilangkan koma dan mengubah ke format string biasa
                
                print(f"Nomor telepon: {nomor_telepon}")
                return nomor_telepon
            else:
                print("Semua entri dalam kolom KTP sudah terisi.")
                return None
        except Exception as e:
            # Tangani jika terjadi kesalahan saat membaca file
            print(f"Error reading Excel file: {e}")
            return None
           
    def update_excel(self, nomor_telepon, ktp, kk, iccid):
        try:
            # Perbarui jalur file Excel dengan menggunakan jalur relatif ke direktori 'data'
            file_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "../data/daftar.xlsx"))

            # Baca file Excel
            df = pd.read_excel(file_path)
            
            # Perbarui nilai KTP dan KK sesuai nomor telepon yang telah dimasukkan
            df.loc[df['Nomor_Telpon'] == (nomor_telepon), 'KTP'] = f"{ktp}"
            df.loc[df['Nomor_Telpon'] == (nomor_telepon), 'KK'] = f"{kk}"
            df.loc[df['Nomor_Telpon'] == (nomor_telepon), 'ICCID'] = f"{iccid}"
            
            # Simpan perubahan kembali ke file Excel
            df.to_excel(file_path, index=False)
            print("Data Excel berhasil diperbarui.")
            return True
        except Exception as e:
            # Tangani jika terjadi kesalahan saat memperbarui file Excel
            print(f"Error updating Excel file: {e}")
            return False

    def header(self, driver, chrome_driver_path):
        date = time.strftime("%a, %d %b %Y %H:%M:%S GMT", time.gmtime())
        
        # Membuat objek ChromeOptions baru
        options = webdriver.ChromeOptions()
        
        # Mendapatkan nilai existing options dari driver
        existing_options = driver.capabilities['goog:chromeOptions']
        
        # Menambahkan nilai yang ada ke options baru
        for key, value in existing_options.items():
            if isinstance(value, list):
                for val in value:
                    options.add_argument(f"--{key}={val}")
            elif isinstance(value, bool):
                if value:
                    options.add_argument(f"--{key}")
            else:
                options.add_argument(f"--{key}={value}")
        
        headers = {
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
            'Accept-Encoding': 'gzip, deflate, br, zstd',
            'Accept-Language': 'id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7',
            'Cache-Control': 'max-age=0',
            'Connection': 'keep-alive',
            'Host': 'registrasi.tri.co.id',
            'If-Modified-Since': date,
            'If-None-Match': 'W/"e7f-18d186630e4"',
            'Sec-Ch-Ua': '"Chromium";v="122", "Not(A:Brand";v="24", "Google Chrome";v="122"',
            'Sec-Ch-Ua-Mobile': '?0',
            'Sec-Ch-Ua-Platform': '"Windows"',
            'Sec-Fetch-Dest': 'document',
            'Sec-Fetch-Mode': 'navigate',
            'Sec-Fetch-Site': 'same-origin',
            'Sec-Fetch-User': '?1',
            'Upgrade-Insecure-Requests': '1',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
        }
        
        # Tambahkan header ke options
        for key, value in headers.items():
            options.add_argument(f"--{key}={value}")
        
        # Buat ulang driver dengan opsi yang diperbarui
        service = ChromeService(executable_path=chrome_driver_path)
        driver.quit()
        driver = webdriver.Chrome(service=service, options=options)
        
        return driver

    
    def form_fields(self, driver, nomor_telepon, ktp, kk, iccid):
        try:
            xpath_nomor = "//input[@formcontrolname='msisdn']"
            xpath_ktp = "//div[@class='trinumber'][2]/input[@class='formgistrasi ng-untouched ng-pristine ng-invalid']"
            xpath_kk = "//div[@class='trinumber'][3]/input[@class='formgistrasi ng-untouched ng-pristine ng-invalid']"
            xpath_iicd = "//mat-radio-button[@id='mat-radio-3']/label"
            xpath_iicd_field = "//div[@class='secondradiotext']/input"
            xpath_submit = "//button[@class='mat-focus-indicator buttonsyle mat-flat-button mat-button-base mat-warn']/span[contains(text(), 'Kirim')]"
            xpath_konfirmasi = "//a[@class='btn btn-yes']"  # Ubah XPath untuk konfirmasi
            detail_status_xpath = "//span[@class='detailstatus']"
            detail_padding_xpath = "//div[@class='detailpadding']"

            # Wait for the element to be visible
            input_nomor = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, xpath_nomor)))

            # Mengisi nomor telepon
            input_nomor.clear()
            input_nomor.send_keys(f"0{nomor_telepon}")

            # Mengisi KTP
            input_ktp = driver.find_element(By.XPATH, xpath_ktp)
            input_ktp.clear()
            input_ktp.send_keys(ktp)

            # Mengisi KK
            input_kk = driver.find_element(By.XPATH, xpath_kk)
            input_kk.clear()
            input_kk.send_keys(kk)

            iicd = driver.find_element(By.XPATH, xpath_iicd)
            iicd.click()

            iidc_field = driver.find_element(By.XPATH, xpath_iicd_field)
            iidc_field.clear()
            iidc_field.send_keys(iccid)

            submit = driver.find_element(By.XPATH, xpath_submit)
            submit.click()

            # Wait for the confirmation button to be clickable
            konfirmasi_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath_konfirmasi)))
            konfirmasi_button.click()
            
            time.sleep(3)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

            # Cari pesan kesalahan
            detail_status = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, detail_status_xpath)))
            detail_padding = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.XPATH, detail_padding_xpath)))
            
            # Inisialisasi variabel untuk menyimpan pesan
            success_message = "BERHASIL di registrasi."
            failure_message = "GAGAL di registrasi."
            success = False

            # Cek apakah pesan registrasi berhasil atau gagal
            for status, padding in zip(detail_status, detail_padding):
                if success_message in status.text:
                    print(status.text)
                    print(padding.text)
                    success = True
                elif failure_message in status.text:
                    print(status.text)
                    print(padding.text)
                    return False

            # Kembalikan True jika registrasi berhasil, False jika gagal
            return success
  
        except Exception as e:
            logging.error(f"Error filling fields: {e}")
            return False

    def url(self, driver):
        url = "https://registrasi.tri.co.id/daftar"
        driver.get(url)
        return driver
    
    def run_register(self, chrome_driver_path, ktp, kk, iccid):
        try:
            # Baca nomor telepon dari file Excel
            nomor_telepon = self.read_excel()
            if nomor_telepon is None:
                print("Tidak dapat menemukan nomor telepon yang belum terdaftar di file Excel.")
                return None
            
            # Inisialisasi driver
            driver = self.initialize_driver(chrome_driver_path)
            
            # Tambahkan header
            driver = self.header(driver, chrome_driver_path)
            
            # Buka URL registrasi
            driver = self.url(driver)
            
            # Isi formulir dengan nomor telepon, KTP, dan KK
            result = self.form_fields(driver, nomor_telepon, ktp, kk, iccid)
        
            # Update file Excel dengan data yang baru jika tidak ada pesan kesalahan
            if result:
                # Cek apakah pesan kesalahan ditemukan
                if not result:
                    # Tidak ada pesan kesalahan, maka lanjutkan untuk memperbarui file Excel
                    self.update_excel(nomor_telepon, ktp, kk, iccid)
                
                return True
            else:
                # Registrasi gagal karena pesan kesalahan ditemukan
                return False
        except Exception as e:
            logging.error(f"Error running registration: {e}")
        return None