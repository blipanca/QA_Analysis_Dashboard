# QA Product Dashboard · Streamlit Python Version

Versi Python dari dashboard QA Telkomsel, dibangun dengan Streamlit. Membaca langsung dari folder `data/*.xlsx` (tidak embed data di kode).

## Struktur folder

```
python_dashboard/
├── app.py                  # Main Streamlit app
├── requirements.txt        # Python dependencies
├── README.md               # File ini
└── data/                   # Source data (XLSX cleaned)
    ├── KIP_April_2026_CLEANED.xlsx
    ├── cohort_oct25_mar26_lis_CLEANED.xlsx
    └── Dashboard_QA_Product_Q1_2026_CLEANED.xlsx
```

## Setup (sekali saja)

### 1. Pastikan Python 3.9+ terpasang

```bash
python3 --version
```

Kalau belum ada, download dari https://www.python.org/downloads/ atau pakai package manager OS anda.

### 2. Buat virtual environment (recommended)

```bash
cd python_dashboard
python3 -m venv venv
source venv/bin/activate          # macOS/Linux
# atau:
venv\Scripts\activate              # Windows
```

### 3. Install dependencies

```bash
pip install -r requirements.txt
```

## Menjalankan dashboard

```bash
streamlit run app.py
```

Browser akan otomatis terbuka di `http://localhost:8501`. Kalau tidak, buka URL itu manual.

Untuk akses dari HP di jaringan lokal yang sama (misal kantor):
```bash
streamlit run app.py --server.address 0.0.0.0
```
Lalu di HP buka `http://<IP-komputer>:8501`.

## Fitur

### 📊 5 view utama (sidebar kiri)

1. **★ Executive Summary** — KPI, bullet points (Performance/Risk/CX), glossary 9 parameter, tren chart, latency status, 6 recommended actions
2. **KIP April 2026** — Detail tiket bulanan, channel mix, daily trend, top issues
3. **Cohort LIS** — Tren bulanan, breakdown area/region/ARPU, kategori tiket
4. **QA Product Q1** — Multi-line trend, top issues, latency benchmark
5. **(Imported files)** — File yang anda upload muncul di sini

### 📁 Add File (sidebar bawah)

Tap **"Upload XLSX/XLS"** untuk tambah file:
- **Smart header detection**: dashboard auto-deteksi baris mana yang berisi header (skip merged title, baris kosong, dst)
- **Pilih sheet manual**: kalau file punya banyak sheet (seperti GA4 workbook), anda pilih satu per import
- **Override header row**: kalau auto-detect salah, anda bisa input baris header manual
- **Quality badge**: 🟢 Great / 🟡 OK / 🔴 Poor — indikasi apakah sheet layak di-chart
- **Preview**: lihat 5 baris pertama sebelum import
- File yang di-import muncul sebagai view baru di sidebar, bisa dihapus dengan tombol ×

## Cache & performa

Dashboard pakai `@st.cache_data` untuk file data, jadi load cuma sekali. Kalau anda update file XLSX di folder `data/`, restart Streamlit atau tap tombol "Clear cache" di menu Streamlit (3 titik kanan atas).

## Update data

Cara paling gampang refresh data dashboard:
1. Ganti file XLSX di folder `data/` dengan nama yang sama
2. Di browser, klik 3-titik kanan atas → "Clear cache" → "Rerun"

Atau restart Streamlit (Ctrl+C lalu `streamlit run app.py` lagi).

## Deploy ke cloud (optional)

Untuk akses dari mana saja tanpa nyalakan komputer:

### Opsi gratis: Streamlit Community Cloud
1. Push folder ini ke GitHub (private repo OK)
2. Login di https://share.streamlit.io
3. New app → pilih repo → klik Deploy
4. Selesai, dapat URL public

### Opsi internal Telkomsel
Deploy ke server internal pakai Docker:

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
EXPOSE 8501
CMD ["streamlit", "run", "app.py", "--server.address=0.0.0.0"]
```

## Troubleshooting

**Q: `FileNotFoundError: data/...xlsx`**
A: Pastikan menjalankan `streamlit run app.py` dari folder `python_dashboard/`, atau folder `data/` ada di lokasi yang sama dengan `app.py`.

**Q: Chart tidak muncul / lambat**
A: Refresh browser. Streamlit dapat agak lambat di first render karena loading library.

**Q: Error saat upload XLSX besar (>50MB)**
A: Default Streamlit batasi upload 200MB. Untuk file lebih besar:
```bash
streamlit run app.py --server.maxUploadSize 500
```

**Q: Mau pakai dark mode di Streamlit**
A: Sudah dark by default (custom CSS di dalam `app.py`). Kalau ada konflik dengan theme Streamlit default, edit `.streamlit/config.toml`:
```toml
[theme]
base = "dark"
primaryColor = "#ED1C24"
backgroundColor = "#0a0a0a"
secondaryBackgroundColor = "#141414"
textColor = "#f5f5f5"
```

## Perbandingan dengan versi HTML

| Aspek | HTML (dashboard.html) | Python (app.py) |
|---|---|---|
| Setup | Tinggal buka di browser | Perlu install Python + deps |
| Data | Embedded (~16KB summary) | Baca langsung dari XLSX (full data) |
| Refresh data | Rebuild HTML | Replace XLSX di folder `data/` |
| Upload file | Drag-drop di browser | Sidebar uploader |
| Akses mobile | Buka file HTML langsung | Perlu server / cloud deploy |
| Filter & drill-down | Terbatas | Streamlit native widgets |
| Deploy ke production | File static | Streamlit Cloud / Docker |
| Best for | Quick share, offline | Analisis berulang, multi-user |

Pakai HTML kalau cuma butuh quick view atau share ke orang yang tidak punya Python.
Pakai Python kalau anda akan refresh data berkala, atau akan deploy ke beberapa user.
