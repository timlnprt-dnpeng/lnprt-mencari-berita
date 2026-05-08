# Scraper Berita LNPRT (Streamlit)

Aplikasi web scraper ini digunakan untuk mengumpulkan berita terkait Lembaga Non-Profit yang Melayani Rumah Tangga (LNPRT) di berbagai wilayah Indonesia. Sumber berita meliputi Google News dan Bing News.

## 🚀 Fitur
- **Filter Berdasarkan Wilayah**: Pilih Provinsi dan Kabupaten/Kota.
- **Filter Kategori**: Scraping otomatis menggunakan daftar kata kunci yang ditentukan dari `Kata Kunci.xlsx`.
- **Dukungan Multi-Sumber**: Mengambil artikel dari Google News dan Bing News (100% Gratis, Tanpa API Key).
- **Exclusion List**: Mengecualikan hasil dari sumber atau URL tertentu yang didaftarkan di `Exclusion-list.txt`.
- **Parsing Tanggal Relatif**: Mampu membaca format waktu relatif (misal: "12j", "1h", "2thn") dan mengonversinya ke format tanggal standar.
- **Mode Gelap / Terang**: UI sepenuhnya responsif dan mendukung Streamlit Dark Mode.

## 📦 Menjalankan Secara Lokal

1. Pastikan Python sudah terinstall.
2. Install dependency:
   ```bash
   pip install -r requirements.txt
   ```
3. Jalankan aplikasi Streamlit:
   ```bash
   streamlit run lnprt_app.py
   ```

## 🌐 Deploy ke Streamlit Cloud

1. Upload repo ini ke GitHub (private/public).
2. Login ke [share.streamlit.io](https://share.streamlit.io).
3. Deploy aplikasi → pilih `lnprt_app.py` sebagai entrypoint.
(Tidak perlu API Key atau konfigurasi Secrets tambahan apapun)

## 📁 File Penting
- `lnprt_app.py`: Main entrypoint aplikasi.
- `Kata Kunci.xlsx`: Daftar kategori dan kata kunci scraping.
- `Daftar Wilayah.xlsx`: Database wilayah provinsi dan kabupaten/kota.
- `Exclusion-list.txt`: Daftar domain/URL yang tidak dimasukkan dalam hasil scraping.
- `Logo.png`: Logo aplikasi.
