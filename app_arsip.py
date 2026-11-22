"""Pemindai Arsip -> Excel

Cara pakai:
1. Instal Tesseract OCR (sertakan paket bahasa Indonesia "ind")
2. pip install streamlit pytesseract pillow pandas openpyxl pymupdf opencv-python-headless
3. streamlit run app_arsip.py
"""
import io
import re
import platform
import shutil

import pandas as pd
import pytesseract
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font
from PIL import Image, ImageOps

st.set_page_config(page_title="Pemindai Arsip ke Excel", layout="wide")
st.title("Pemindai Arsip → Excel")

with st.sidebar:
    st.header("Pengaturan")
    default_tesseract = (
        r"C:\Program Files\Tesseract-OCR\tesseract.exe"
        if platform.system() == "Windows"
        else (shutil.which("tesseract") or "tesseract")
    )
    path_tesseract = st.text_input(
        "Lokasi tesseract (kosongkan jika sudah di PATH)",
        default_tesseract,
    ).strip()
    bahasa = st.text_input("Bahasa OCR", "ind+eng")
    pola_nomor = st.text_input(
        "Pola nomor arsip (regex)",
        r"(?=\S*[A-Za-z])\b\d{1,6}\s?/\s?[A-Za-z0-9.\-]+(?:\s?/\s?[A-Za-z0-9.\-]+)+",
        help="Dipakai jika label 'Nomor :' tidak ditemukan. Contoh cocok: 00580/LS/KEBTKE.2/VI/2017.",
    )
    excel_lama = st.file_uploader("Excel tujuan (opsional, data ditambahkan di bawahnya)", type=["xlsx"])

pytesseract.pytesseract.tesseract_cmd = path_tesseract or "tesseract"

st.session_state.setdefault("baris", [])
st.session_state.setdefault("sudah", set())


KOLOM = [
    "No", "Kode Klasifikasi Arsip", "Nomor Arsip/Berkas", "Uraian Informasi Arsip",
    "Kurun Waktu", "Tingkat Perkembangan", "Jumlah", "Keterangan",
    "Nomor Boks", "Lokasi Simpan", "Jangka Simpan dan Nasib Akhir", "Kategori Arsip",
]


def nomor_berikut(ws):
    """Nomor urut berikutnya = angka terbesar di kolom A + 1."""
    n = 0
    for (v,) in ws.iter_rows(min_col=1, max_col=1, values_only=True):
        if isinstance(v, (int, float)):
            n = max(n, int(v))
    return n + 1


def hilangkan_bayangan(img):
    try:
        import cv2
        import numpy as np
    except ImportError:
        return ImageOps.autocontrast(img, cutoff=2)
    a = np.array(img)
    bg = cv2.medianBlur(cv2.dilate(a, np.ones((25, 25), np.uint8)), 51)
    norm = cv2.normalize(255 - cv2.absdiff(a, bg), None, 0, 255, cv2.NORM_MINMAX)
    return Image.fromarray(norm)


def luruskan(img):
    """Putar otomatis jika foto miring 90/180/270 derajat."""
    try:
        osd = pytesseract.image_to_osd(img, output_type=pytesseract.Output.DICT)
        sudut = osd.get("rotate", 0)
        return img.rotate(-sudut, expand=True) if sudut else img
    except Exception:
        pass
    terbaik, skor = img, -1.0
    for sudut in (0, 90, 180, 270):
        k = img.rotate(sudut, expand=True)
        d = pytesseract.image_to_data(k, lang=bahasa, output_type=pytesseract.Output.DICT)
        nilai = sum(float(c) for c, t in zip(d["conf"], d["text"]) if t.strip() and float(c) >= 0)
        if nilai > skor:
            terbaik, skor = k, nilai
    return terbaik


def siapkan(img):
    img = ImageOps.exif_transpose(img).convert("L")
    if max(img.size) < 2000:
        r = 2000 / max(img.size)
        img = img.resize((int(img.width * r), int(img.height * r)))
    img = hilangkan_bayangan(img)
    return luruskan(img)


def ekstrak(teks):
    nomor = ""
    m = re.search(r"\bno(?:mor)?\.?\s*(?:arsip|surat|berkas|spm)?\s*[:\-]\s*(\S+)", teks, re.I)
    if m:
        nomor = m.group(1).strip()
    else:
        m = re.search(pola_nomor, teks)
        if m:
            nomor = m.group(0).strip()

    uraian = ""
    cocok_uraian = list(
        re.finditer(
            r"(?:^|\n)[^\n]*\bur[aei]i?[aeil]n\b[^\n]*(?:\n[^\n]*){0,5}",
            teks,
            re.I,
        )
    )
    if cocok_uraian:
        potongan = cocok_uraian[-1].group(0)
        # Hapus bagian label sebelum/pada kata uraian sampai tanda pemisah (: atau -) atau spasi panjang
        potongan = re.sub(r"^.*?ur[aei]i?[aeil]n\b[^\n:;\-=]*[:;\-=]?\s*", "", potongan, flags=re.I | re.S)
        # Hentikan jika sudah masuk catatan kaki / tanda tangan / NIP di bagian bawah dokumen SPM
        potongan = re.split(
            r"\n\s*(?:[•\*\-]\s*)?(?:semua\s+bukti|kebenaran\s+perhitungan|jakarta[,\s]|a\.n\.|kuasa\s+pengguna|pejabat\s+penandatangan|nip\.?|233\s+\d+)\b",
            potongan,
            maxsplit=1,
            flags=re.I,
        )[0]
        potongan = re.sub(r"^[\s:=\-–—.]+", "", potongan)
        uraian = " ".join(potongan.split())
    return nomor, uraian

def halaman(f):
    if f.name.lower().endswith(".pdf"):
        import fitz  # pymupdf

        dok = fitz.open(stream=f.read(), filetype="pdf")
        for i, p in enumerate(dok):
            png = p.get_pixmap(dpi=200).tobytes("png")
            yield f"{f.name} hal.{i + 1}", Image.open(io.BytesIO(png))
    else:
        yield f.name, Image.open(f)


def proses(nama, img):
    if nama in st.session_state.sudah:
        return
    teks = pytesseract.image_to_string(siapkan(img), lang=bahasa)
    nomor, uraian = ekstrak(teks)
    st.session_state.setdefault("teks_ocr", {})
    st.session_state.teks_ocr[nama] = teks
    st.session_state.baris.append([nomor, uraian, nama])
    st.session_state.sudah.add(nama)


tab1, tab2 = st.tabs(["Unggah foto/scan", "Kamera"])
with tab1:
    berkas = st.file_uploader(
        "Pilih banyak file sekaligus", type=["jpg", "jpeg", "png", "pdf"], accept_multiple_files=True
    )
    if berkas:
        bar = st.progress(0.0)
        for i, f in enumerate(berkas):
            for nama, img in halaman(f):
                proses(nama, img)
            bar.progress((i + 1) / len(berkas))
with tab2:
    foto = st.camera_input("Foto dokumen (lurus, cahaya cukup)")
    if foto:
        proses(f"kamera_{foto.file_id}", Image.open(foto))

tampilkan_ocr = st.checkbox("Tampilkan hasil OCR mentah (untuk cek)")
if tampilkan_ocr and st.session_state.baris:
    st.session_state.setdefault("teks_ocr", {})
    for sumber, t_ocr in st.session_state.teks_ocr.items():
        with st.expander(f"Teks OCR: {sumber}"):
            st.text(t_ocr)

st.subheader("Hasil (periksa dan koreksi sebelum diekspor)")
df = pd.DataFrame(st.session_state.baris, columns=["Nomor Arsip", "Uraian Arsip", "Sumber"])
hasil = st.data_editor(df, num_rows="dynamic", use_container_width=True)

c1, c2 = st.columns(2)
with c1:
    if hasil.empty:
        st.info("Belum ada data.")
    else:
        if excel_lama:
            wb = load_workbook(excel_lama)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = "Daftar Arsip"
            ws.append(KOLOM)
            for c in ws[1]:
                c.font = Font(bold=True)
                c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
            for huruf, lebar in zip("ABCDEFGHIJKL", (5, 16, 30, 60, 14, 16, 10, 14, 12, 14, 22, 14)):
                ws.column_dimensions[huruf].width = lebar
        no = nomor_berikut(ws)
        for _, r in hasil.iterrows():
            baris = [""] * len(KOLOM)
            baris[0] = no
            baris[2] = r["Nomor Arsip"]
            baris[3] = r["Uraian Arsip"]
            ws.append(baris)
            ws.cell(ws.max_row, 4).alignment = Alignment(wrap_text=True, vertical="top")
            no += 1
        buf = io.BytesIO()
        wb.save(buf)
        st.download_button(
            "Unduh Excel",
            buf.getvalue(),
            "arsip.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
with c2:
    if st.button("Kosongkan semua"):
        st.session_state.baris = []
        st.session_state.sudah = set()
        st.rerun()
