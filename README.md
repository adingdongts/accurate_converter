# Accurate SO Converter

Konversi otomatis data Excel transaksi ke format import **Sales Order Accurate**.

## Struktur file

```
accurate_converter/
├── app.py                  # Streamlit web app
├── accurate_convert.py     # CLI script (tanpa Streamlit)
└── requirements.txt
```

---

## Cara pakai

### Install dependensi

```bash
pip install -r requirements.txt
```

---

### A · Streamlit (web app, direkomendasikan)

```bash
streamlit run app.py
```

Buka browser ke `http://localhost:8501`, lalu:
1. Upload file template Accurate (`.xlsx`)
2. Upload file data transaksi (`.xlsx`)
3. Klik **Konversi sekarang**
4. Download hasilnya

---

### B · CLI (terminal)

```bash
# Cara dasar
python accurate_convert.py --template template_accurate.xlsx --input data_so.xlsx

# Tentukan nama file output sendiri
python accurate_convert.py -t template_accurate.xlsx -i data_so.xlsx -o hasil_import.xlsx
```

---

## Format data input

File Excel dengan **7 kolom** (urutan harus sesuai):

| Kolom | Keterangan |
|---|---|
| Tgl | Tanggal transaksi |
| No Transaksi | Nomor SO (misal: SO-0626-001) |
| ID Pelanggan | Nama / kode pelanggan |
| Kode Produk | Kode / nama barang |
| Qty | Jumlah |
| Satuan | Satuan (PCS, DUS, dll) |
| Harga Satuan | Harga per satuan (angka) |

Satu `No Transaksi` boleh punya banyak baris item.

---

## Format output

Mengikuti struktur template Accurate:
- **3 baris pertama** = header kolom dari template (tidak diubah)
- Per order: **1 baris HEADER** + **N baris ITEM**
- **Tanpa baris EXPENSE**
