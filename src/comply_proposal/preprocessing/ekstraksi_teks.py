"""
Bagian A — Ekstraksi teks mentah dari proposal (jalur JP/JC).

Modul ini bertanggung jawab HANYA untuk mengubah file .docx / .pdf menjadi
teks yang bersih dari artefak ekstraksi (header/footer berulang, nomor
halaman berulang, spasi ganda), TANPA menghilangkan informasi struktural
(nama gaya/heading, nomor bab/pasal, batas section/halaman) — informasi
struktural itu dibutuhkan oleh modul segmentasi unit (Bagian B).

Asumsi:
- PEMBARUAN: tahap redaksi PII terpisah TIDAK dipakai di implementasi ini
  (keputusan peneliti). Modul ini TIDAK melakukan deteksi/redaksi PII apa
  pun, dan data yang diproses TIDAK dianonimkan — nama, NIM, tanda tangan,
  dan identitas lain yang ada di dokumen asli akan ikut masuk APA ADANYA
  ke `teks_unit` pada output JSONL (Bagian G). Kalau perlindungan privasi
  tetap dibutuhkan (mis. sebelum data dibagikan ke pihak lain), itu harus
  ditangani terpisah di luar modul ini — JANGAN anggap pipeline ini sudah
  menanganinya.
- Deteksi placeholder redaksi (mis. "[DIREDAKSI]") di bawah tetap
  dipertahankan sebagai jaring pengaman murni (berjaga-jaga kalau ada sisa
  penanda manual di sebagian dokumen), TAPI karena tahap redaksi tidak
  lagi dipakai secara sistematis, jangan mengandalkan mekanisme ini
  sebagai bentuk perlindungan PII yang sebenarnya.
- Ekstraksi metadata tata letak (font, margin) BUKAN tanggung jawab modul
  ini — itu ada di modul terpisah untuk jalur JM (Bagian C), karena JM wajib
  membaca langsung dari file asli, bukan dari hasil ekstraksi teks di sini.
- Deteksi "kemungkinan hasil pindai/scan" pada modul ini hanya menghasilkan
  sinyal mentah (`kemungkinan_hasil_pindai`); keputusan akhir menandai
  proposal sebagai NA-01 diambil pada tahap penggabungan data (Bagian D),
  supaya satu proposal yang bermasalah tidak menghentikan seluruh batch.
"""

from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List, Set, Union

PathLike = Union[str, Path]


class EkstraksiError(Exception):
    """Dilempar untuk kegagalan ekstraksi yang terduga (file rusak/tidak bisa dibuka/kosong).

    Sengaja dipisahkan dari exception generik supaya pemanggil batch (lihat
    Bagian G) bisa menangkap ini secara spesifik, mencatatnya ke log, dan
    melanjutkan ke proposal berikutnya tanpa menghentikan seluruh proses.
    """


# ---------------------------------------------------------------------------
# Normalisasi teks & deteksi placeholder redaksi (dipakai bersama DOCX & PDF)
# ---------------------------------------------------------------------------

_POLA_SPASI_UNICODE = re.compile(
    "[\u00A0\u2000-\u200A\u202F\u3000\u200B]"
)
_POLA_SPASI_GANDA = re.compile(r"[ \t]{2,}")
_POLA_BARIS_KOSONG_BERLEBIH = re.compile(r"\n{3,}")
_TABEL_KUTIP_PINTAR = str.maketrans(
    {
        "‘": "'",
        "’": "'",
        "“": '"',
        "”": '"',
    }
)

# Token placeholder redaksi yang mungkin masih tersisa di teks. Daftar ini
# sengaja eksplisit (bukan pola "[...]" umum) supaya tidak salah menandai
# kurung siku yang memang bagian isi proposal (mis. kutipan referensi).
_POLA_PLACEHOLDER_REDAKSI = re.compile(
    r"\[\s*(DIREDAKSI|REDACTED|NAMA\s+DIREDAKSI|NIM\s+DIREDAKSI|TTD\s+DIREDAKSI)\s*\]",
    re.IGNORECASE,
)


def _normalisasi_teks(teks: str) -> str:
    """Merapikan spasi & tanda kutip tanpa mengubah struktur baris/paragraf.

    Langkah (sengaja terbatas & didokumentasikan, karena teks ini juga
    dipakai untuk pemeriksaan BHS/FMT — normalisasi berlebihan bisa
    mengaburkan pelanggaran yang sebenarnya):
    1. Samakan berbagai karakter spasi unicode (nbsp, dll.) jadi spasi biasa,
       TERMASUK zero-width space (U+200B) -- ditemukan empiris (2026-09-18)
       pada PDF hasil ekspor tool tertentu yang menyisipkan U+200B sebagai
       SATU-SATUNYA pemisah antar-kata (tanpa spasi biasa sama sekali, mis.
       "1.1​​Latar​​Belakang"). Karena U+200B bukan
       whitespace menurut `\s` regex Python, ini bikin SEMUA pola heading
       Bagian B/C gagal total pada dokumen yang terkena (bukan cuma fallback
       ke pola lebih lemah -- headingnya sungguh tidak ketemu sama sekali).
    2. Ratakan spasi/tab ganda dalam satu baris jadi satu spasi.
    3. Samakan tanda kutip pintar jadi tanda kutip lurus.
    4. Batasi baris kosong berturut-turut maksimal satu (antar paragraf).

    Input kosong/None-safe: mengembalikan string kosong.
    """
    if not teks:
        return ""

    teks = _POLA_SPASI_UNICODE.sub(" ", teks)
    teks = teks.translate(_TABEL_KUTIP_PINTAR)
    baris_dirapikan = [_POLA_SPASI_GANDA.sub(" ", b).strip() for b in teks.split("\n")]
    teks = "\n".join(baris_dirapikan)
    teks = _POLA_BARIS_KOSONG_BERLEBIH.sub("\n\n", teks)
    return teks.strip()


def _mengandung_placeholder_redaksi(teks: str) -> bool:
    """True jika teks masih mengandung sisa placeholder redaksi PII."""
    return bool(_POLA_PLACEHOLDER_REDAKSI.search(teks or ""))


# ---------------------------------------------------------------------------
# Bagian DOCX
# ---------------------------------------------------------------------------


def _petakan_indeks_section_docx(dokumen: Any) -> List[int]:
    """Memetakan setiap paragraf level-dokumen ke indeks section (0-based).

    python-docx tidak menyediakan pemetaan ini secara langsung. Section baru
    ditandai oleh elemen <w:sectPr> yang bersarang di <w:pPr> paragraf
    TERAKHIR pada section tsb (section terakhir dalam dokumen justru
    sectPr-nya anak langsung <w:body>, bukan di dalam paragraf, sehingga
    tidak menambah indeks lagi setelah paragraf terakhir).

    Pemetaan ini penting karena template resmi punya margin berbeda per
    section (Sampul vs badan proposal vs lampiran) — dibutuhkan oleh jalur
    JM (Bagian C), meski modul ekstraksi teks ini sendiri tidak membaca
    margin.

    Urutan hasil mengikuti urutan `dokumen.paragraphs` (paragraf di dalam
    tabel tidak termasuk, sama seperti behavior `document.paragraphs`).
    """
    from docx.oxml.ns import qn

    indeks_section = 0
    hasil: List[int] = []
    body = dokumen.element.body
    for anak in body.iterchildren():
        if anak.tag == qn("w:p"):
            hasil.append(indeks_section)
            ppr = anak.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
                indeks_section += 1
    return hasil


def ekstrak_docx(path: PathLike) -> Dict[str, Any]:
    """Mengekstrak teks terstruktur dari file .docx.

    Input:
        path: path ke file .docx. PII di dalam dokumen TIDAK diredaksi oleh
            modul ini maupun tahap sebelumnya (lihat catatan asumsi di
            docstring modul) — akan ikut apa adanya di `teks_penuh`/`unit_teks`.

    Output (dict):
        {
            "path_asal": str,
            "format_asli": "docx",
            "teks_penuh": str,              # gabungan semua paragraf non-kosong
            "unit_teks": [                  # satu entri per paragraf level-dokumen
                {
                    "indeks_paragraf": int,
                    "indeks_section": int | None,   # lihat _petakan_indeks_section_docx
                    "gaya": str | None,              # nama style, mis. "Heading 1"
                    "level_outline": int | None,
                    "teks": str,                     # sudah dinormalisasi
                    "mengandung_placeholder_redaksi": bool,
                },
                ...
            ],
            "jumlah_section": int,
            "peringatan": [str, ...],
        }

    Header/footer TIDAK perlu dibersihkan secara heuristik di sini seperti
    pada PDF: pada format .docx, header/footer tersimpan sebagai objek
    terpisah per section (section.header / section.footer), bukan
    tercampur ke dalam alur paragraf body, sehingga `dokumen.paragraphs`
    sudah otomatis tidak menyertakannya.

    Raises:
        EkstraksiError: jika file tidak bisa dibuka/dibaca sebagai .docx.
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
        raise EkstraksiError(f"Gagal membuka file DOCX '{path.name}': {exc}") from exc

    indeks_section_per_paragraf = _petakan_indeks_section_docx(dokumen)

    unit_teks: List[Dict[str, Any]] = []
    potongan_teks_penuh: List[str] = []

    for i, paragraf in enumerate(dokumen.paragraphs):
        teks_mentah = paragraf.text
        teks_bersih = _normalisasi_teks(teks_mentah)

        gaya = None
        try:
            if paragraf.style is not None:
                gaya = paragraf.style.name
        except Exception:
            gaya = None

        level_outline = None
        try:
            level_outline = paragraf.paragraph_format.outline_level
        except Exception:
            level_outline = None

        unit_teks.append(
            {
                "indeks_paragraf": i,
                "indeks_section": (
                    indeks_section_per_paragraf[i]
                    if i < len(indeks_section_per_paragraf)
                    else None
                ),
                "gaya": gaya,
                "level_outline": level_outline,
                "teks": teks_bersih,
                "mengandung_placeholder_redaksi": _mengandung_placeholder_redaksi(teks_mentah),
            }
        )
        if teks_bersih:
            potongan_teks_penuh.append(teks_bersih)

    if not potongan_teks_penuh:
        peringatan.append(
            "Tidak ada teks yang berhasil diekstrak dari file DOCX ini (dokumen mungkin kosong)."
        )

    return {
        "path_asal": str(path),
        "format_asli": "docx",
        "teks_penuh": "\n".join(potongan_teks_penuh),
        "unit_teks": unit_teks,
        "jumlah_section": len(dokumen.sections),
        "peringatan": peringatan,
    }


# ---------------------------------------------------------------------------
# Bagian PDF
# ---------------------------------------------------------------------------

_RASIO_ZONA_HEADER_FOOTER = 0.12
_AMBANG_PROPORSI_HALAMAN_UNTUK_ARTEFAK = 0.5
_MINIMUM_HALAMAN_UNTUK_DETEKSI_ARTEFAK = 3
_AMBANG_KARAKTER_PER_HALAMAN_UNTUK_PINDAI = 20.0

_POLA_NOMOR_HALAMAN = re.compile(
    r"^(halaman|page|hal\.?)?\s*\.?\s*\d{1,4}(\s*(dari|of|/)\s*\d{1,4})?\s*\.?\s*$",
    re.IGNORECASE,
)
# Validasi angka romawi kanonik (bukan sekadar cek karakter), untuk mengurangi
# false positive terhadap kata singkat berbahasa Indonesia (mis. "di").
_POLA_ANGKA_ROMAWI = re.compile(
    r"^M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$",
    re.IGNORECASE,
)


def _adalah_kandidat_nomor_halaman(teks: str) -> bool:
    """True jika satu baris teks terlihat seperti nomor halaman berdiri sendiri.

    Dipakai HANYA pada baris yang sudah dipastikan berada di zona
    header/footer (lihat pemanggil) dan HANYA dianggap artefak jika pola ini
    konsisten muncul di posisi yang sama pada banyak halaman (lihat
    `_deteksi_pola_nomor_halaman_berulang`) — supaya kata pendek yang
    kebetulan mirip angka romawi (mis. "di") tidak salah terhapus hanya
    karena muncul sekali di zona footer suatu halaman.
    """
    t = teks.strip()
    if not t:
        return False
    if _POLA_NOMOR_HALAMAN.match(t):
        return True
    if 1 < len(t) <= 6 and _POLA_ANGKA_ROMAWI.match(t):
        return True
    return False


def _ambil_baris_halaman(halaman: Any) -> List[Dict[str, Any]]:
    """Mengambil baris teks + posisi (bbox) dari satu halaman PDF via PyMuPDF.

    Diurutkan dari atas ke bawah lalu kiri ke kanan, dengan asumsi tata letak
    satu kolom (sesuai template proposal resmi). Baris kosong dibuang.
    """
    data = halaman.get_text("dict")
    baris_list: List[Dict[str, Any]] = []
    for blok in data.get("blocks", []):
        for baris in blok.get("lines", []):
            teks_baris = "".join(span.get("text", "") for span in baris.get("spans", [])).strip()
            if not teks_baris:
                continue
            x0, y0, x1, y1 = baris.get("bbox", (0.0, 0.0, 0.0, 0.0))
            baris_list.append({"teks": teks_baris, "x0": x0, "y0": y0, "x1": x1, "y1": y1})

    baris_list.sort(key=lambda b: (round(b["y0"], 1), b["x0"]))
    return baris_list


def _deteksi_artefak_header_footer(
    baris_per_halaman: List[List[Dict[str, Any]]],
    tinggi_halaman: List[float],
) -> Set[str]:
    """Mengidentifikasi teks yang tampil PERSIS SAMA di zona header/footer
    (12% teratas / 12% terbawah halaman) pada banyak halaman -> dianggap
    artefak header/footer berulang (mis. judul dokumen, nama institusi),
    bukan bagian isi proposal.

    Perbandingan dilakukan setelah normalisasi & lowercasing agar variasi
    spasi tidak mengelabui deteksi. Butuh minimal
    `_MINIMUM_HALAMAN_UNTUK_DETEKSI_ARTEFAK` halaman supaya "berulang" punya
    arti statistik; dokumen yang lebih pendek dari itu tidak melalui langkah
    ini (dianggap terlalu berisiko salah tandai).

    Return: himpunan teks (sudah dinormalisasi, huruf kecil) yang dianggap
    artefak header/footer.
    """
    jumlah_halaman = len(baris_per_halaman)
    if jumlah_halaman < _MINIMUM_HALAMAN_UNTUK_DETEKSI_ARTEFAK:
        return set()

    kemunculan: Dict[str, Set[int]] = defaultdict(set)
    for idx_halaman, (baris_list, tinggi) in enumerate(zip(baris_per_halaman, tinggi_halaman)):
        zona = tinggi * _RASIO_ZONA_HEADER_FOOTER
        for baris in baris_list:
            di_header = baris["y1"] <= zona
            di_footer = baris["y0"] >= (tinggi - zona)
            if not (di_header or di_footer):
                continue
            kunci = _normalisasi_teks(baris["teks"]).lower()
            if kunci:
                kemunculan[kunci].add(idx_halaman)

    ambang = max(
        _MINIMUM_HALAMAN_UNTUK_DETEKSI_ARTEFAK,
        int(jumlah_halaman * _AMBANG_PROPORSI_HALAMAN_UNTUK_ARTEFAK),
    )
    return {kunci for kunci, halaman_muncul in kemunculan.items() if len(halaman_muncul) >= ambang}


def _deteksi_pola_nomor_halaman_berulang(
    baris_per_halaman: List[List[Dict[str, Any]]],
    tinggi_halaman: List[float],
) -> Dict[str, bool]:
    """Mendeteksi apakah zona header dan/atau footer SECARA KONSISTEN berisi
    nomor halaman (nilainya berubah tiap halaman, sehingga tidak tertangkap
    oleh `_deteksi_artefak_header_footer` yang membandingkan teks persis
    sama). Yang dibandingkan di sini adalah POLA (mis. "12", "Halaman 12
    dari 45"), bukan nilainya.

    Return: {"header": bool, "footer": bool} — True jika zona tsb dianggap
    berisi penomoran halaman berulang pada mayoritas halaman.
    """
    jumlah_halaman = len(baris_per_halaman)
    if jumlah_halaman < _MINIMUM_HALAMAN_UNTUK_DETEKSI_ARTEFAK:
        return {"header": False, "footer": False}

    halaman_dengan_nomor_di_header: Set[int] = set()
    halaman_dengan_nomor_di_footer: Set[int] = set()

    for idx_halaman, (baris_list, tinggi) in enumerate(zip(baris_per_halaman, tinggi_halaman)):
        zona = tinggi * _RASIO_ZONA_HEADER_FOOTER
        for baris in baris_list:
            if not _adalah_kandidat_nomor_halaman(baris["teks"]):
                continue
            if baris["y1"] <= zona:
                halaman_dengan_nomor_di_header.add(idx_halaman)
            elif baris["y0"] >= (tinggi - zona):
                halaman_dengan_nomor_di_footer.add(idx_halaman)

    ambang = max(
        _MINIMUM_HALAMAN_UNTUK_DETEKSI_ARTEFAK,
        int(jumlah_halaman * _AMBANG_PROPORSI_HALAMAN_UNTUK_ARTEFAK),
    )
    return {
        "header": len(halaman_dengan_nomor_di_header) >= ambang,
        "footer": len(halaman_dengan_nomor_di_footer) >= ambang,
    }


def ekstrak_pdf(path: PathLike) -> Dict[str, Any]:
    """Mengekstrak teks terstruktur dari file .pdf.

    Input:
        path: path ke file .pdf. PII di dalam dokumen TIDAK diredaksi oleh
            modul ini maupun tahap sebelumnya (lihat catatan asumsi di
            docstring modul) — akan ikut apa adanya di `teks_penuh`/`unit_teks`.

    Output (dict):
        {
            "path_asal": str,
            "format_asli": "pdf",
            "teks_penuh": str,              # gabungan semua halaman non-kosong
            "unit_teks": [                  # satu entri per halaman
                {
                    "indeks_halaman": int,
                    "teks": str,             # setelah dibersihkan header/footer & dinormalisasi
                    "jumlah_baris_terhapus_sebagai_artefak": int,
                    "mengandung_placeholder_redaksi": bool,
                },
                ...
            ],
            "jumlah_halaman": int,
            "kemungkinan_hasil_pindai": bool,   # sinyal mentah utk kandidat NA-01
            "header_footer_terdeteksi": [str, ...],  # utk transparansi/audit
            "peringatan": [str, ...],
        }

    Catatan: fungsi ini TIDAK memutuskan status NA-01 secara final — itu
    tanggung jawab modul penggabungan data (Bagian D), yang bisa
    mempertimbangkan konteks lain (mis. ukuran file, jumlah gambar).
    Fungsi ini hanya melaporkan sinyal `kemungkinan_hasil_pindai`.

    Raises:
        EkstraksiError: jika file tidak bisa dibuka atau tidak punya halaman.
    """
    try:
        import pymupdf  # PyMuPDF (nama modul 'pymupdf'; 'fitz' adalah alias lama/deprecated)
    except ImportError as exc:  # pragma: no cover
        raise EkstraksiError(
            "Paket 'PyMuPDF' belum terpasang. Jalankan: pip install PyMuPDF"
        ) from exc

    path = Path(path)
    peringatan: List[str] = []

    try:
        dokumen = pymupdf.open(str(path))
    except Exception as exc:
        raise EkstraksiError(f"Gagal membuka file PDF '{path.name}': {exc}") from exc

    try:
        jumlah_halaman = dokumen.page_count
        if jumlah_halaman == 0:
            raise EkstraksiError(f"File PDF '{path.name}' tidak memiliki halaman.")

        baris_per_halaman: List[List[Dict[str, Any]]] = []
        tinggi_halaman: List[float] = []
        for halaman in dokumen:
            tinggi_halaman.append(float(halaman.rect.height))
            baris_per_halaman.append(_ambil_baris_halaman(halaman))

        artefak_teks_persis = _deteksi_artefak_header_footer(baris_per_halaman, tinggi_halaman)
        pola_nomor_berulang = _deteksi_pola_nomor_halaman_berulang(
            baris_per_halaman, tinggi_halaman
        )

        unit_teks: List[Dict[str, Any]] = []
        potongan_teks_penuh: List[str] = []
        total_karakter_mentah = 0

        for idx_halaman, (baris_list, tinggi) in enumerate(zip(baris_per_halaman, tinggi_halaman)):
            zona = tinggi * _RASIO_ZONA_HEADER_FOOTER
            baris_dipertahankan: List[str] = []
            jumlah_terhapus = 0

            for baris in baris_list:
                total_karakter_mentah += len(baris["teks"])
                di_header = baris["y1"] <= zona
                di_footer = baris["y0"] >= (tinggi - zona)
                kunci = _normalisasi_teks(baris["teks"]).lower()

                if kunci in artefak_teks_persis:
                    jumlah_terhapus += 1
                    continue
                if di_header and pola_nomor_berulang["header"] and _adalah_kandidat_nomor_halaman(
                    baris["teks"]
                ):
                    jumlah_terhapus += 1
                    continue
                if di_footer and pola_nomor_berulang["footer"] and _adalah_kandidat_nomor_halaman(
                    baris["teks"]
                ):
                    jumlah_terhapus += 1
                    continue

                baris_dipertahankan.append(baris["teks"])

            teks_halaman = _normalisasi_teks("\n".join(baris_dipertahankan))
            mengandung_placeholder = any(
                _mengandung_placeholder_redaksi(b) for b in baris_dipertahankan
            )

            unit_teks.append(
                {
                    "indeks_halaman": idx_halaman,
                    "teks": teks_halaman,
                    "jumlah_baris_terhapus_sebagai_artefak": jumlah_terhapus,
                    "mengandung_placeholder_redaksi": mengandung_placeholder,
                }
            )
            if teks_halaman:
                potongan_teks_penuh.append(teks_halaman)

        rata_rata_karakter_per_halaman = total_karakter_mentah / jumlah_halaman
        kemungkinan_hasil_pindai = (
            rata_rata_karakter_per_halaman < _AMBANG_KARAKTER_PER_HALAMAN_UNTUK_PINDAI
        )

        if kemungkinan_hasil_pindai:
            peringatan.append(
                "Rata-rata karakter per halaman sangat rendah "
                f"({rata_rata_karakter_per_halaman:.1f}); kemungkinan file ini hasil "
                "pindai/scan (kandidat NA-01). Keputusan akhir NA-01 diambil pada "
                "tahap penggabungan data (Bagian D)."
            )
        if not potongan_teks_penuh:
            peringatan.append("Tidak ada teks yang berhasil diekstrak dari file PDF ini.")

        return {
            "path_asal": str(path),
            "format_asli": "pdf",
            "teks_penuh": "\n\n".join(potongan_teks_penuh),
            "unit_teks": unit_teks,
            "jumlah_halaman": jumlah_halaman,
            "kemungkinan_hasil_pindai": kemungkinan_hasil_pindai,
            "header_footer_terdeteksi": sorted(artefak_teks_persis),
            "peringatan": peringatan,
        }
    finally:
        dokumen.close()


# ---------------------------------------------------------------------------
# Fungsi pemilih format (dipakai modul-modul selanjutnya, mis. Bagian D/G)
# ---------------------------------------------------------------------------


def ekstrak_dokumen(path: PathLike) -> Dict[str, Any]:
    """Mendeteksi format dari ekstensi file lalu memanggil ekstraktor yang sesuai.

    Raises:
        EkstraksiError: jika ekstensi file bukan .docx atau .pdf, atau jika
            ekstraksi gagal (lihat `ekstrak_docx`/`ekstrak_pdf`).
    """
    path = Path(path)
    ekstensi = path.suffix.lower()
    if ekstensi == ".docx":
        return ekstrak_docx(path)
    if ekstensi == ".pdf":
        return ekstrak_pdf(path)
    raise EkstraksiError(
        f"Format file tidak didukung untuk '{path.name}' (ekstensi '{ekstensi}'). "
        "Hanya .docx dan .pdf yang didukung."
    )


if __name__ == "__main__":
    # Demo pemanggilan tunggal untuk verifikasi cepat Bagian A saja.
    # Pemrosesan batch + gabungan ground truth ada di Bagian G.
    import json
    import sys

    if len(sys.argv) != 2:
        print("Pemakaian: python ekstraksi_teks.py <path_ke_file.docx_atau_.pdf>")
        raise SystemExit(1)

    try:
        hasil = ekstrak_dokumen(sys.argv[1])
    except EkstraksiError as exc:
        print(f"Gagal ekstraksi: {exc}")
        raise SystemExit(1)

    ringkasan = {
        "path_asal": hasil["path_asal"],
        "format_asli": hasil["format_asli"],
        "jumlah_unit_teks": len(hasil["unit_teks"]),
        "panjang_teks_penuh": len(hasil["teks_penuh"]),
        "peringatan": hasil["peringatan"],
    }
    print(json.dumps(ringkasan, indent=2, ensure_ascii=False))
