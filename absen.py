import cv2
from pyzbar.pyzbar import decode
import pandas as pd
from datetime import datetime
import numpy as np

# File untuk menyimpan data absensi
attendance_file = "attendance.xlsx"

# Fungsi untuk mencatat data absensi ke sheet tertentu
def record_attendance(data, sheet_name="Sheet1"):
    # Pisahkan data berdasarkan format QR code
    try:
        name, kelas, jabatan = data.split(",")  # Format: Nama,Kelas,Jabatan
    except ValueError:
        print("⚠️ Format data QR code tidak valid! Gunakan format: Nama,Kelas,Jabatan")
        return
    
    # Cek apakah file sudah ada
    try:
        # Load file Excel dengan engine openpyxl
        with pd.ExcelWriter(attendance_file, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            try:
                # Baca sheet yang sudah ada
                df = pd.read_excel(attendance_file, sheet_name=sheet_name)
            except ValueError:
                # Jika sheet belum ada, buat DataFrame kosong dengan kolom yang benar
                print(f"Sheet '{sheet_name}' tidak ditemukan, membuat sheet baru...")
                df = pd.DataFrame(columns=["Nama", "Kelas", "Jabatan", "Timestamp"])
            
            # Tambahkan data baru
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            new_entry = {"Nama": name.strip(), "Kelas": kelas.strip(), "Jabatan": jabatan.strip(), "Timestamp": timestamp}

            # Hindari duplikasi absensi dalam waktu singkat
            if not ((df["Nama"] == name.strip()) & (df["Kelas"] == kelas.strip()) & 
                    (df["Jabatan"] == jabatan.strip()) & 
                    (df["Timestamp"].str[:16] == timestamp[:16])).any():
                df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
                # Tulis kembali ke file di sheet yang dimaksud
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                print(f"✅ {name.strip()} berhasil diabsen pada {timestamp} di sheet '{sheet_name}'")
            else:
                print(f"⚠️ {name.strip()} dari kelas '{kelas.strip()}' dan jabatan '{jabatan.strip()}' sudah tercatat sebelumnya di sheet '{sheet_name}'!")
    except FileNotFoundError:
        # Jika file tidak ada, buat file baru dengan sheet default
        print(f"File '{attendance_file}' tidak ditemukan. Membuat file baru...")
        df = pd.DataFrame(columns=["Nama", "Kelas", "Jabatan", "Timestamp"])
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        new_entry = {"Nama": name.strip(), "Kelas": kelas.strip(), "Jabatan": jabatan.strip(), "Timestamp": timestamp}
        df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
        df.to_excel(attendance_file, sheet_name=sheet_name, index=False)
        print(f"✅ {name.strip()} berhasil diabsen pada {timestamp} di sheet '{sheet_name}'")

# Fungsi utama untuk mendeteksi QR code
def qr_code_scanner():
    sheet_name = input("Masukkan nama sheet (atau tekan Enter untuk default 'Sheet1'): ").strip()
    if not sheet_name:
        sheet_name = "Sheet1"
    
    print("📸 Memulai kamera... Tekan 'q' untuk keluar.")
    cap = cv2.VideoCapture(0)  # Gunakan kamera bawaan
    
    while True:
        ret, frame = cap.read()
        if not ret:
            print("❌ Tidak dapat mengakses kamera!")
            break

        # Decode QR code dari frame
        for qrcode in decode(frame):
            qr_data = qrcode.data.decode('utf-8')
            qr_type = qrcode.type
            print(f"📋 Deteksi {qr_type}: {qr_data}")
            
            # Rekam data absensi
            record_attendance(qr_data, sheet_name)
            
            # Tampilkan QR code di layar
            pts = qrcode.polygon
            if len(pts) > 0:
                pts = [(pt.x, pt.y) for pt in pts]
                pts = np.array(pts, np.int32).reshape((-1, 1, 2))
                cv2.polylines(frame, [pts], True, (0, 255, 0), 2)
            cv2.putText(frame, qr_data, (qrcode.rect.left, qrcode.rect.top - 10),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)

        # Tampilkan video di jendela
        cv2.imshow("QR Code Scanner", frame)

        # Tekan 'q' untuk keluar
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

# Jalankan scanner
if __name__ == "__main__":
    qr_code_scanner()
