Sebuah tool monitoring dan pelaporan jaringan komprehensif yang dirancang untuk manajemen infrastruktur jaringan enterprise. Tool ini mengotomatisasi pengumpulan data utilisasi, metrik performa sistem, dan inventori hardware dari peralatan jaringan, menghasilkan laporan Excel profesional dengan analitik yang detail.

## 🚀 Fitur

### Kemampuan Inti
- **Multi-Device Monitoring**: Monitoring bersamaan untuk beberapa perangkat jaringan
- **Analisis Utilisasi FPC**: Analisis detail utilisasi port dan traffic
- **Inventori Hardware**: Pelacakan komponen hardware yang komprehensif
- **Monitoring Performa Sistem**: Metrik CPU, memori, storage, dan temperatur
- **Manajemen Alarm**: Pengumpulan alarm real-time dan monitoring status
- **Pelaporan Profesional**: Laporan Excel otomatis dengan beberapa worksheet

### Fitur Lanjutan
- **Deteksi Modul Intelligent**: Identifikasi otomatis modul FPC dari hardware chassis
- **Inferensi Tipe SFP**: Deteksi pintar tipe optical transceiver
- **Analisis Pola Traffic**: Analitik utilisasi traffic yang canggih
- **Timestamping Zone-aware**: Deteksi otomatis zona waktu Indonesia (WIB/WITA/WIT)
- **Sequential Processing**: Pemrosesan satu-perangkat-per-waktu yang reliable untuk tingkat keberhasilan maksimum
- **Debug Logging**: Logging komprehensif dengan output debug yang terorganisir

### Worksheet Laporan
1. **Utilisasi FPC** - Data utilisasi FPC utama
2. **Utilisasi Port** - Utilisasi detail level port
3. **Status Alarm** - Kondisi alarm saat ini
4. **Inventori Hardware** - Daftar lengkap komponen hardware  
5. **Performa Sistem** - Metrik kesehatan sistem
6. **Ringkasan Dashboard** - Ringkasan eksekutif dan metrik kunci

## 📋 Persyaratan

### Kebutuhan Sistem
- **Sistem Operasi**: Windows 10/11, Linux, macOS
- **Versi Python**: 3.7 atau lebih tinggi
- **Memori**: Minimum 4GB RAM yang direkomendasikan
- **Storage**: 500MB ruang kosong untuk laporan dan log

### Kebutuhan Jaringan
- Konektivitas jaringan ke perangkat target
- Akses TACACS+ server untuk autentikasi
- Akses SSH ke peralatan jaringan (biasanya port 22)

### Dependensi Python
```bash
pip install paramiko openpyxl
```

## 🛠️ Instalasi

### 1. Persiapan Folder
Pastikan Anda memiliki folder dengan struktur berikut:
```
Lab/
├── lab.py          # Script utama
├── access_lab.xml        # File konfigurasi akses
├── list_lab.txt            # Daftar perangkat yang akan dimonitor
└── README.md                # Dokumentasi ini
```

### 2. Install Dependensi Python
```bash
pip install paramiko openpyxl
```

### 3. Verifikasi Script
Buka Command Prompt atau PowerShell, navigasi ke folder script:
```bash
cd "C:\Users\User\OneDrive\Desktop\lab"
python lab.py
```

## ⚙️ Konfigurasi

### 1. File Konfigurasi Akses
Edit file `lab_access.xml` yang sudah ada di folder script untuk mengatur autentikasi:

```xml
<access>
  <tacacs-user>xxx</tacacs-user>
  <tacacs-pass>yyyyy</tacacs-pass>
  <!-- Optional hop: jika perlu ssh dari TACACS ke router, isi ini -->
  <router-user>aaa</router-user>
  <router-pass>bbb</router-pass>
  <tacacs-server>ccc</tacacs-server>
</access>
```

**Catatan**: Ganti username dan password sesuai dengan kredensial yang Anda miliki.

### 2. Konfigurasi Daftar Perangkat
File `list_lab.txt` sudah tersedia dengan daftar perangkat default. Edit sesuai kebutuhan:

```
dut-1
```

**Catatan**: Tambahkan atau hapus perangkat sesuai dengan yang ingin Anda monitor.

### 3. Pengaturan Autentikasi

| Parameter | Deskripsi | Contoh |
|-----------|-----------|--------|
| `tacacs-server` | Alamat IP server TACACS+ | ccc |
| `tacacs-user` | Username untuk autentikasi TACACS+ | network-admin |
| `tacacs-pass` | Password untuk autentikasi TACACS+ | SecurePass123 |
| `router-user` | Username level perangkat | device-admin |
| `router-pass` | Password level perangkat | DevicePass456 |

## 🚀 Penggunaan

### Cara Menjalankan Script

#### Metode 1: Menggunakan Command Prompt
1. Buka Command Prompt (cmd) atau PowerShell
2. Navigasi ke folder script:
   ```powershell
   cd "C:\Users\User\OneDrive\Desktop\lab"
   ```
3. Jalankan script:
   ```powershell
   python lab.py
   ```

#### Metode 2: Langsung dari File Explorer
1. Buka folder `C:\Users\User\OneDrive\Desktop\lab`
2. Klik kanan pada area kosong, pilih "Open PowerShell window here"
3. Ketik: `python lab.py`

#### Metode 3: Double-click (jika Python sudah di-associate)
1. Double-click file `lab.py` langsung dari File Explorer
2. Script akan berjalan di command prompt yang terbuka otomatis

### Eksekusi Dasar
```powershell
cd "C:\Users\User\OneDrive\Desktop\lab"
python lab.py
```

### Opsi Command Line
Script mendukung berbagai mode eksekusi dan parameter (konfigurasi dalam script):

- **Sequential Processing**: Mode default untuk reliabilitas maksimum
- **Debug Mode**: Logging yang ditingkatkan untuk troubleshooting
- **Custom Timeout**: Timeout koneksi yang dapat disesuaikan
- **Retry Logic**: Percobaan ulang yang dapat dikonfigurasi untuk koneksi yang gagal

### Alur Eksekusi
1. **Inisialisasi**: Muat konfigurasi dan validasi dependensi
2. **Autentikasi**: Koneksi ke server TACACS+
3. **Device Discovery**: Baca daftar perangkat dari konfigurasi
4. **Pengumpulan Data**: Monitoring perangkat secara berurutan
5. **Pemrosesan Data**: Parse dan analisis data yang dikumpulkan
6. **Pembuatan Laporan**: Buat laporan Excel dengan beberapa worksheet
7. **Cleanup**: Organisir file debug dan log

## 📊 File Output

### Lokasi Script dan File
**Lokasi Script**:
```
C:\Users\User\OneDrive\Desktop\lab\
├── lab.py          # Script utama yang dijalankan
├── access_lab.xml        # Konfigurasi akses dan kredensial
├── list_lab.txt            # Daftar perangkat target
└── README.md                # Dokumentasi
```

### Lokasi Output
Setelah script dijalankan, file hasil akan dibuat di Desktop:
```
C:\Users\User\Desktop\
├── LAB-Occupancy\     # Folder utama laporan
|   ├── LAB_Occupancy_Report_29Dec2025_0920.xlsx
│   ├── Capture_FPC-Occupancy20251229\
│   │   └── All Debug\
│           ├── Debug Logs\
│           ├── Debug XML\
│           └── Temp Files
│           ├── dut1_alarms.xml
```

**Catatan**: Script akan otomatis membuat folder-folder ini jika belum ada.

### Isi Laporan

#### Sheet Utilisasi Utama
- Nama Node dan Divisi Regional
- Nama Interface dan Deskripsi  
- Tipe Modul dan Kapasitas Port
- Traffic Saat Ini (GB) dan Utilisasi (%)
- Alert Traffic dan Indikator Status

#### Sheet Utilisasi Port
- Statistik detail level port
- Analisis Last Flapped
- Deteksi keberadaan SFP
- Status konfigurasi
- Alert flapping

#### Sheet Status Alarm
- Kondisi alarm real-time
- Klasifikasi dan tingkat keparahan alarm
- Timestamp dan deskripsi
- Pelacakan status

#### Sheet Inventori Hardware
- Daftar lengkap komponen
- Nomor part dan serial number
- Deskripsi model dan versi
- Status dan kesehatan komponen

#### Sheet Performa Sistem
- Utilisasi CPU dan load average
- Penggunaan dan ketersediaan memori
- Kapasitas dan utilisasi storage
- Monitoring temperatur
- Informasi versi software

#### Sheet Ringkasan Dashboard
- Metrik ringkasan eksekutif
- Key performance indicator
- Analisis tren
- Ringkasan alert
