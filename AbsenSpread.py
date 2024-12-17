import cv2
from pyzbar.pyzbar import decode
import gspread
from google.oauth2 import service_account
from datetime import datetime
import numpy as np

# Kredensial Google Sheets
credentials = service_account.Credentials.from_service_account_file(
    r'D:\Kumpulan Kammil\web\optimum-time-444901-a2-6ef1becc746b.json',  # Gunakan raw string
    scopes=["https://www.googleapis.com/auth/spreadsheets"]
)

# Autentikasi dan buka spreadsheet
gc = gspread.authorize(credentials)
spreadsheet = gc.open_by_key('1d-CKcHEM5eTnSg6uIAdjX4UyXMwvYYNkARhvY84p5kw')
worksheet = spreadsheet.sheet1

# Fungsi untuk mencatat data absensi ke Google Sheets
def record_attendance_google(data):
    try:
        name, kelas, jabatan = data.split(",")
    except ValueError:
        print("⚠️ Format data QR code tidak valid! Gunakan format: Nama,Kelas,Jabatan")
        return

    existing_data = worksheet.get_all_records()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Ambil waktu saat ini dalam format datetime untuk perbandingan
    current_time = datetime.now()

    # Periksa apakah data QR code sudah ada dalam 30 detik terakhir
    for record in existing_data:
        # Pastikan ada key 'Timestamp' dalam record
        if 'Timestamp' not in record:
            continue

        try:
            record_time = datetime.strptime(record['Timestamp'], "%Y-%m-%d %H:%M:%S")
        except ValueError:
            print(f"⚠️ Format timestamp tidak valid untuk record: {record}")
            continue
        
        # Jika pencatatan lebih dari 30 detik yang lalu, anggap itu sebagai entri baru
        if (record['Nama'] == name.strip() and
            record['Kelas'] == kelas.strip() and
            record['Jabatan'] == jabatan.strip() and
            (current_time - record_time).seconds <= 30):  # Periksa waktu
            print(f"⚠️ {name.strip()} sudah tercatat sebelumnya dalam 30 detik terakhir!")
            return

    # Entri baru jika tidak ada duplikasi
    new_entry = [name.strip(), kelas.strip(), jabatan.strip(), timestamp]
    worksheet.append_row(new_entry)
    print(f"✅ {name.strip()} berhasil diabsen pada {timestamp}")

# Fungsi utama untuk mendeteksi QR code
def qr_code_scanner():
    print("📸 Memulai kamera... Tekan 'q' untuk keluar.")
    cap = cv2.VideoCapture(0)

    while True:
        ret, frame = cap.read()
        if not ret:
            print("❌ Tidak dapat mengakses kamera!")
            break

        for qrcode in decode(frame):
            qr_data = qrcode.data.decode('utf-8')
            print(f"📋 Deteksi QR Code: {qr_data}")
            record_attendance_google(qr_data)

            pts = qrcode.polygon
            if pts:
                pts = [(pt.x, pt.y) for pt in pts]
                pts = np.array(pts, np.int32).reshape((-1, 1, 2))
                cv2.polylines(frame, [pts], True, (0, 255, 0), 2)
            cv2.putText(frame, qr_data, (qrcode.rect.left, qrcode.rect.top - 10),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)

        cv2.imshow("QR Code Scanner", frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()

# Jalankan scanner
if __name__ == "__main__":
    qr_code_scanner()
