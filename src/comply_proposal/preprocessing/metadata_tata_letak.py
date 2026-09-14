"""
Bagian C — Ekstraksi metadata tata letak (jalur JM), terpisah dari Bagian A & B.

PRINSIP UTAMA (WAJIB dijaga di seluruh modul ini): jalur JM tidak boleh
membaca margin/font dari teks hasil ekstraksi Bagian A. Semua fungsi di sini
membuka file asli sendiri (python-docx untuk .docx, PyMuPDF untuk .pdf) dan
sama sekali tidak menerima/menggunakan output `ekstrak_docx`/`ekstrak_pdf`
sebagai input. Satu-satunya helper yang diimpor dari Bagian A adalah
`_petakan_indeks_section_docx`, karena fungsi itu murni menelusuri struktur
XML dokumen (bukan teks hasil ekstraksi) untuk memetakan section — bug-prone
kalau diduplikasi, jadi sengaja dipakai bersama.

Temuan penting dari template resmi (WAJIB diperhatikan saat memakai output
modul ini, lihat juga konteks percakapan):
- Template resmi punya 3 section dengan margin BERBEDA (Sampul / badan
  proposal Bab I-III / Lampiran I-II). Karena itu setiap entri di sini
  SELALU menyertakan section/halaman mana yang diukur — JANGAN mengambil
  satu margin lalu menganggapnya berlaku untuk seluruh dokumen.
- Style bawaan "Title" pada file template resmi = 22pt, sedangkan Pedoman
  v0.2 Pasal 22 ayat (6) huruf a menyebut 16pt untuk identitas sampul.
  Modul ini TIDAK memutuskan mana yang "benar" — ia hanya melaporkan apa
  adanya nilai yang berhasil diresolusi (lihat `sumber_ukuran_font`, akan
  bernilai "gaya:Title" untuk kasus ini). Keputusan metodologis (pakai
  16pt sesuai ketentuan tertulis, atau 22pt sesuai file template) ada di
  tangan peneliti, bukan kode ini.

Tingkat kepercayaan:
- DOCX: `tingkat_kepercayaan = "tinggi"` — margin & font dibaca langsung
  dari properti dokumen (section.*_margin, run.font.*), bukan estimasi.
- PDF: `tingkat_kepercayaan = "rendah"` — PDF tidak menyimpan properti
  margin section. Margin PDF di sini adalah ESTIMASI dari bounding box
  konten teks per halaman (jarak terdekat teks ke tepi kertas), BUKAN
  nilai presisi dari pengaturan dokumen aslinya.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

from .ekstraksi_teks import EkstraksiError, _petakan_indeks_section_docx

PathLike = Union[str, Path]

# ---------------------------------------------------------------------------
# Util konversi & deteksi penomoran bab (independen dari Bagian A/B)
# ---------------------------------------------------------------------------


def _pt_ke_cm(pt: float) -> float:
    return pt / 72.0 * 2.54


def _length_ke_cm(nilai: Any) -> Optional[float]:
    """Mengonversi objek `Length` python-docx (EMU) ke cm. None-safe."""
    if nilai is None:
        return None
    return round(nilai.cm, 3)


_POLA_BAB = re.compile(r"^\s*bab\s+([ivxlcdm]+|\d+)\b", re.IGNORECASE)
_POLA_ROMAWI_VALID = re.compile(
    r"^M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$", re.IGNORECASE
)


def _klasifikasi_gaya_nomor(token: str) -> str:
    """Mengklasifikasi token nomor bab ("I", "1", dst.) sebagai romawi/angka."""
    if token.isdigit():
        return "angka"
    if token and _POLA_ROMAWI_VALID.match(token):
        return "romawi"
    return "tidak_dikenali"


# ---------------------------------------------------------------------------
# Bagian DOCX
# ---------------------------------------------------------------------------


def _resolusi_font_run(run: Any, gaya_paragraf: Any) -> Dict[str, Any]:
    """Meresolusi nama & ukuran font EFEKTIF satu run docx.

    python-docx hanya mengembalikan nilai eksplisit yang di-override di
    level run (`run.font.name`/`run.font.size`); kalau None, nilainya
    diwariskan dari gaya paragraf, dan kalau gaya itu sendiri tidak
    menimpanya, diwariskan lagi dari `gaya.base_style`, dst. Fungsi ini
    menelusuri rantai itu SECARA TERPISAH untuk nama dan ukuran (karena
    keduanya bisa diresolusi dari level berbeda), dan mencatat level mana
    yang akhirnya menentukan nilai (`sumber_nama_font`/`sumber_ukuran_font`)
    supaya transparan untuk audit (mis. kasus style "Title" 22pt di atas).

    Return: {"nama_font", "sumber_nama_font", "ukuran_font_pt", "sumber_ukuran_font"}
    """
    nama_font = run.font.name
    sumber_nama = "run" if nama_font is not None else None
    ukuran_font = run.font.size.pt if run.font.size is not None else None
    sumber_ukuran = "run" if ukuran_font is not None else None

    gaya = gaya_paragraf
    tingkat = 0
    while (nama_font is None or ukuran_font is None) and gaya is not None:
        label = f"gaya:{gaya.name}" if tingkat == 0 else f"gaya_dasar:{gaya.name}"
        if nama_font is None and gaya.font.name is not None:
            nama_font = gaya.font.name
            sumber_nama = label
        if ukuran_font is None and gaya.font.size is not None:
            ukuran_font = gaya.font.size.pt
            sumber_ukuran = label
        gaya = gaya.base_style
        tingkat += 1

    return {
        "nama_font": nama_font,
        "sumber_nama_font": sumber_nama or "tidak_diketahui",
        "ukuran_font_pt": ukuran_font,
        "sumber_ukuran_font": sumber_ukuran or "tidak_diketahui",
    }


def ekstrak_metadata_docx(path: PathLike) -> Dict[str, Any]:
    """Mengekstrak metadata tata letak dari file .docx (jalur JM).

    Input:
        path: path ke file .docx.

    Output (dict):
        {
            "path_asal": str,
            "format_asli": "docx",
            "tingkat_kepercayaan": "tinggi",
            "section": [
                {
                    "indeks_section": int,
                    "margin_atas_cm": float | None, "margin_bawah_cm": float | None,
                    "margin_kiri_cm": float | None, "margin_kanan_cm": float | None,
                    "lebar_halaman_cm": float | None, "tinggi_halaman_cm": float | None,
                    "orientasi": "portrait" | "landscape" | None,
                },
                ...
            ],
            "font_per_paragraf": [   # HANYA paragraf yang punya run berteks (bukan kosong)
                {
                    "indeks_paragraf": int,       # sama dengan indexing `document.paragraphs` di Bagian A
                    "indeks_section": int | None,
                    "gaya": str | None,
                    "cuplikan_teks": str,          # 60 char pertama, MENTAH (belum dinormalisasi A)
                    "run": [
                        {
                            "indeks_run": int,
                            "nama_font": str | None, "sumber_nama_font": str,
                            "ukuran_font_pt": float | None, "sumber_ukuran_font": str,
                            "tebal": bool | None,   # nilai run.font.bold MENTAH (TIDAK diresolusi lewat rantai gaya)
                            "miring": bool | None,  # idem run.font.italic
                        },
                        ...
                    ],
                },
                ...
            ],
            "penomoran_bab_terdeteksi": [
                {"indeks_paragraf": int, "teks_tertangkap": str, "gaya_penomoran": "angka"|"romawi"|"tidak_dikenali"},
                ...
            ],
            "peringatan": [str, ...],
        }

    Catatan: `tebal`/`miring` sengaja TIDAK diresolusi lewat rantai gaya
    seperti nama/ukuran font (python-docx tidak expose ini semudah
    font.name/size), jadi bisa None meski secara visual tampak
    tebal/miring karena diwariskan dari gaya. Kalau butuh nilai efektif,
    itu perlu diresolusi terpisah oleh pemanggil lewat `gaya`/style chain.

    Keterbatasan lain: `_resolusi_font_run` hanya menelusuri rantai
    run -> style paragraf -> base_style. Kalau font TIDAK di-override di
    level manapun dalam rantai itu (diwariskan dari `docDefaults` pada
    styles.xml, mis. gaya "Normal" yang benar-benar polos), hasilnya
    `nama_font`/`ukuran_font_pt` = None dengan sumber "tidak_diketahui" —
    ini BUKAN berarti dokumen tidak punya font efektif (Word tetap
    merender sesuatu, mis. default Word Calibri 11), hanya berarti
    python-docx tidak expose docDefaults semudah style.font. Kalau kasus
    ini sering muncul di data riil, perlu penanganan tambahan (parsing
    manual `styles.xml` -> `w:docDefaults`) di luar cakupan modul ini.

    Raises:
        EkstraksiError: jika file tidak bisa dibuka.
    """
    try:
        from docx import Document
    except ImportError as exc:  # pragma: no cover
        raise EkstraksiError(
            "Paket 'python-docx' belum terpasang. Jalankan: pip install python-docx"
        ) from exc

    path = Path(path)
    peringatan: List[str] = []

    try:
        dokumen = Document(str(path))
    except Exception as exc:
        raise EkstraksiError(
            f"Gagal membuka file DOCX '{path.name}' untuk ekstraksi metadata JM: {exc}"
        ) from exc

    section_info: List[Dict[str, Any]] = []
    for i, section in enumerate(dokumen.sections):
        try:
            orientasi = section.orientation.name.lower()
        except Exception:
            orientasi = None
        section_info.append(
            {
                "indeks_section": i,
                "margin_atas_cm": _length_ke_cm(section.top_margin),
                "margin_bawah_cm": _length_ke_cm(section.bottom_margin),
                "margin_kiri_cm": _length_ke_cm(section.left_margin),
                "margin_kanan_cm": _length_ke_cm(section.right_margin),
                "lebar_halaman_cm": _length_ke_cm(section.page_width),
                "tinggi_halaman_cm": _length_ke_cm(section.page_height),
                "orientasi": orientasi,
            }
        )

    indeks_section_per_paragraf = _petakan_indeks_section_docx(dokumen)

    font_per_paragraf: List[Dict[str, Any]] = []
    penomoran_bab: List[Dict[str, Any]] = []

    for i, paragraf in enumerate(dokumen.paragraphs):
        teks_baris = paragraf.text.strip()

        cocok_bab = _POLA_BAB.match(teks_baris)
        if cocok_bab:
            penomoran_bab.append(
                {
                    "indeks_paragraf": i,
                    "teks_tertangkap": teks_baris,
                    "gaya_penomoran": _klasifikasi_gaya_nomor(cocok_bab.group(1)),
                }
            )

        run_info: List[Dict[str, Any]] = []
        for j, run in enumerate(paragraf.runs):
            if not run.text.strip():
                continue
            resolusi = _resolusi_font_run(run, paragraf.style)
            run_info.append(
                {
                    "indeks_run": j,
                    "tebal": run.font.bold,
                    "miring": run.font.italic,
                    **resolusi,
                }
            )

        if not run_info:
            continue

        gaya_nama = None
        try:
            if paragraf.style is not None:
                gaya_nama = paragraf.style.name
        except Exception:
            gaya_nama = None

        font_per_paragraf.append(
            {
                "indeks_paragraf": i,
                "indeks_section": (
                    indeks_section_per_paragraf[i] if i < len(indeks_section_per_paragraf) else None
                ),
                "gaya": gaya_nama,
                "cuplikan_teks": teks_baris[:60],
                "run": run_info,
            }
        )

    if len(section_info) > 1:
        peringatan.append(
            f"Dokumen memiliki {len(section_info)} section dengan kemungkinan margin "
            "berbeda-beda antar section (lihat temuan template resmi); JANGAN asumsikan "
            "satu margin berlaku untuk seluruh dokumen. Gunakan field 'indeks_section' "
            "pada tiap entri untuk mencocokkan margin yang relevan."
        )

    return {
        "path_asal": str(path),
        "format_asli": "docx",
        "tingkat_kepercayaan": "tinggi",
        "section": section_info,
        "font_per_paragraf": font_per_paragraf,
        "penomoran_bab_terdeteksi": penomoran_bab,
        "peringatan": peringatan,
    }


# ---------------------------------------------------------------------------
# Bagian PDF
# ---------------------------------------------------------------------------


def _ambil_baris_dan_span_halaman(halaman: Any) -> List[Dict[str, Any]]:
    """Mengambil baris teks PDF beserta span font-nya LANGSUNG dari file asli.

    Sengaja terpisah dari `_ambil_baris_halaman` di Bagian A (yang hanya
    butuh teks per baris untuk pembersihan header/footer) karena di sini
    granularitas SPAN (bukan baris) yang dibutuhkan — satu baris bisa
    berisi beberapa span dengan font berbeda.

    Return: list of {"teks": str, "span": [{"nama_font", "ukuran_font_pt",
        "tebal", "miring", "x0", "y0", "x1", "y1"}, ...]}
    """
    data = halaman.get_text("dict")
    hasil: List[Dict[str, Any]] = []
    for blok in data.get("blocks", []):
        for baris in blok.get("lines", []):
            span_list: List[Dict[str, Any]] = []
            potongan_teks: List[str] = []
            for span in baris.get("spans", []):
                teks_span = span.get("text", "")
                if not teks_span.strip():
                    continue
                potongan_teks.append(teks_span)
                flags = span.get("flags", 0)
                nama_font = span.get("font")
                x0, y0, x1, y1 = span.get("bbox", (0.0, 0.0, 0.0, 0.0))
                span_list.append(
                    {
                        "nama_font": nama_font,
                        "ukuran_font_pt": round(float(span.get("size", 0.0)), 2),
                        # flags bit 4 (0x10) = bold, bit 1 (0x2) = italic (dok. PyMuPDF);
                        # ditambah cek nama font sbg jaring pengaman krn flags kadang tidak akurat
                        # tergantung aplikasi pembuat PDF.
                        "tebal": bool(flags & 0x10) or "bold" in (nama_font or "").lower(),
                        "miring": bool(flags & 0x2)
                        or any(k in (nama_font or "").lower() for k in ("italic", "oblique")),
                        "x0": x0,
                        "y0": y0,
                        "x1": x1,
                        "y1": y1,
                    }
                )
            if span_list:
                hasil.append({"teks": "".join(potongan_teks).strip(), "span": span_list})
    return hasil


def _perkirakan_margin_halaman(
    span_list: List[Dict[str, Any]], lebar_pt: float, tinggi_pt: float
) -> Dict[str, Optional[float]]:
    """Mengestimasi margin dari bounding box konten teks per halaman.

    Ini BUKAN margin presisi (PDF tidak menyimpan properti margin section
    seperti .docx) — hanya jarak terdekat teks yang terdeteksi ke tepi
    kertas. Bisa meleset kalau halaman punya sedikit konten (margin
    ter-estimasi lebih besar dari aslinya) atau ada elemen non-teks
    (gambar/tabel garis) yang sebenarnya lebih dekat ke tepi.
    """
    if not span_list:
        return {
            "margin_atas_cm": None,
            "margin_bawah_cm": None,
            "margin_kiri_cm": None,
            "margin_kanan_cm": None,
        }
    margin_kiri_pt = min(s["x0"] for s in span_list)
    margin_kanan_pt = lebar_pt - max(s["x1"] for s in span_list)
    margin_atas_pt = min(s["y0"] for s in span_list)
    margin_bawah_pt = tinggi_pt - max(s["y1"] for s in span_list)
    return {
        "margin_atas_cm": round(_pt_ke_cm(margin_atas_pt), 3),
        "margin_bawah_cm": round(_pt_ke_cm(margin_bawah_pt), 3),
        "margin_kiri_cm": round(_pt_ke_cm(margin_kiri_pt), 3),
        "margin_kanan_cm": round(_pt_ke_cm(margin_kanan_pt), 3),
    }


def ekstrak_metadata_pdf(path: PathLike) -> Dict[str, Any]:
    """Mengekstrak metadata tata letak dari file .pdf (jalur JM), berupa ESTIMASI.

    Input:
        path: path ke file .pdf.

    Output (dict):
        {
            "path_asal": str,
            "format_asli": "pdf",
            "tingkat_kepercayaan": "rendah",
            "halaman": [
                {
                    "indeks_halaman": int,
                    "lebar_halaman_cm": float, "tinggi_halaman_cm": float,
                    "margin_atas_cm": float | None, "margin_bawah_cm": float | None,
                    "margin_kiri_cm": float | None, "margin_kanan_cm": float | None,
                },
                ...
            ],
            "font_per_halaman": [
                {"indeks_halaman": int, "span": [{"nama_font", "ukuran_font_pt", "tebal", "miring", "x0","y0","x1","y1"}, ...]},
                ...
            ],
            "penomoran_bab_terdeteksi": [
                {"indeks_halaman": int, "teks_tertangkap": str, "gaya_penomoran": str},
                ...
            ],
            "peringatan": [str, ...],   # SELALU berisi disclaimer estimasi margin
        }

    Raises:
        EkstraksiError: jika file tidak bisa dibuka atau tidak punya halaman.
    """
    try:
        import pymupdf
    except ImportError as exc:  # pragma: no cover
        raise EkstraksiError(
            "Paket 'PyMuPDF' belum terpasang. Jalankan: pip install PyMuPDF"
        ) from exc

    path = Path(path)
    peringatan: List[str] = [
        "Estimasi margin PDF dihitung dari bounding box konten teks per halaman, "
        "BUKAN nilai margin presisi dari pengaturan dokumen aslinya (PDF tidak "
        "menyimpan properti margin section seperti .docx)."
    ]

    try:
        dokumen = pymupdf.open(str(path))
    except Exception as exc:
        raise EkstraksiError(
            f"Gagal membuka file PDF '{path.name}' untuk ekstraksi metadata JM: {exc}"
        ) from exc

    try:
        jumlah_halaman = dokumen.page_count
        if jumlah_halaman == 0:
            raise EkstraksiError(f"File PDF '{path.name}' tidak memiliki halaman.")

        halaman_info: List[Dict[str, Any]] = []
        font_per_halaman: List[Dict[str, Any]] = []
        penomoran_bab: List[Dict[str, Any]] = []
        ada_span_sama_sekali = False

        for idx_halaman, halaman in enumerate(dokumen):
            lebar = float(halaman.rect.width)
            tinggi = float(halaman.rect.height)
            baris_span = _ambil_baris_dan_span_halaman(halaman)
            semua_span = [s for baris in baris_span for s in baris["span"]]
            if semua_span:
                ada_span_sama_sekali = True

            margin_estimasi = _perkirakan_margin_halaman(semua_span, lebar, tinggi)
            halaman_info.append(
                {
                    "indeks_halaman": idx_halaman,
                    "lebar_halaman_cm": round(_pt_ke_cm(lebar), 3),
                    "tinggi_halaman_cm": round(_pt_ke_cm(tinggi), 3),
                    **margin_estimasi,
                }
            )
            font_per_halaman.append({"indeks_halaman": idx_halaman, "span": semua_span})

            for baris in baris_span:
                cocok = _POLA_BAB.match(baris["teks"])
                if cocok:
                    penomoran_bab.append(
                        {
                            "indeks_halaman": idx_halaman,
                            "teks_tertangkap": baris["teks"],
                            "gaya_penomoran": _klasifikasi_gaya_nomor(cocok.group(1)),
                        }
                    )

        if not ada_span_sama_sekali:
            peringatan.append(
                "Tidak ada span teks yang berhasil diambil dari file ini (kemungkinan "
                "hasil pindai/scan); metadata font & estimasi margin tidak tersedia "
                "untuk halaman manapun."
            )

        return {
            "path_asal": str(path),
            "format_asli": "pdf",
            "tingkat_kepercayaan": "rendah",
            "halaman": halaman_info,
            "font_per_halaman": font_per_halaman,
            "penomoran_bab_terdeteksi": penomoran_bab,
            "peringatan": peringatan,
        }
    finally:
        dokumen.close()


# ---------------------------------------------------------------------------
# Fungsi pemilih format
# ---------------------------------------------------------------------------


def ekstrak_metadata_dokumen(path: PathLike) -> Dict[str, Any]:
    """Mendeteksi format dari ekstensi file lalu memanggil ekstraktor metadata JM yang sesuai."""
    path = Path(path)
    ekstensi = path.suffix.lower()
    if ekstensi == ".docx":
        return ekstrak_metadata_docx(path)
    if ekstensi == ".pdf":
        return ekstrak_metadata_pdf(path)
    raise EkstraksiError(
        f"Format file tidak didukung untuk '{path.name}' (ekstensi '{ekstensi}'). "
        "Hanya .docx dan .pdf yang didukung."
    )


if __name__ == "__main__":
    import json
    import sys

    if len(sys.argv) != 2:
        print("Pemakaian: python -m comply_proposal.preprocessing.metadata_tata_letak <path_file>")
        raise SystemExit(1)

    try:
        hasil = ekstrak_metadata_dokumen(sys.argv[1])
    except EkstraksiError as exc:
        print(f"Gagal ekstraksi metadata JM: {exc}")
        raise SystemExit(1)

    if hasil["format_asli"] == "docx":
        ringkasan = {
            "format_asli": hasil["format_asli"],
            "tingkat_kepercayaan": hasil["tingkat_kepercayaan"],
            "section": hasil["section"],
            "jumlah_paragraf_dengan_font": len(hasil["font_per_paragraf"]),
            "penomoran_bab_terdeteksi": hasil["penomoran_bab_terdeteksi"],
            "peringatan": hasil["peringatan"],
        }
    else:
        ringkasan = {
            "format_asli": hasil["format_asli"],
            "tingkat_kepercayaan": hasil["tingkat_kepercayaan"],
            "halaman": hasil["halaman"],
            "penomoran_bab_terdeteksi": hasil["penomoran_bab_terdeteksi"],
            "peringatan": hasil["peringatan"],
        }
    print(json.dumps(ringkasan, indent=2, ensure_ascii=False))
