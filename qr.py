import qrcode
import os

# Fungsi untuk menghasilkan QR Code
def generate_qr_code(name, kelas, jabatan):
    # Format data untuk QR code: Nama,Kelas,Jabatan
    qr_data = f"{name},{kelas},{jabatan}"
    
    # Generate QR code
    qr_img = qrcode.make(qr_data)
    
    # Pastikan folder qr_codes ada
    folder = "qr_codes"
    if not os.path.exists(folder):
        os.makedirs(folder)  # Buat folder jika belum ada
    
    # Menyimpan QR code sebagai file PNG dalam folder qr_codes
    file_path = os.path.join(folder, f"qrcode_{name.replace(' ', '_')}.png")
    qr_img.save(file_path)
    print(f"✅ QR Code untuk {name} berhasil dibuat dan disimpan di '{file_path}'!")

# Fungsi utama untuk memasukkan data berulang kali
def input_data():
    while True:
        name = input("Masukkan Nama (atau ketik 'selesai' untuk selesai): ")
        if name.lower() == 'selesai':
            print("Proses input selesai.")
            break
        kelas = input("Masukkan Kelas: ")
        jabatan = input("Masukkan Jabatan: ")
        
        # Panggil fungsi untuk membuat QR Code
        generate_qr_code(name, kelas, jabatan)

# Jalankan fungsi input data
if __name__ == "__main__":
    input_data()
