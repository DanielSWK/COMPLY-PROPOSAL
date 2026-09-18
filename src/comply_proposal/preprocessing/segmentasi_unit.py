"""
Bagian B — Segmentasi teks proposal ke 16 unit anotasi (U00-U15).

Modul ini mengambil hasil ekstraksi teks dari Bagian A (`ekstraksi_teks.py`)
dan memecahnya menjadi 16 unit sesuai Panduan Anotasi dan Taksonomi
Pelanggaran v0.2, berdasarkan deteksi heading/subheading di dalam teks.

PENTING — batasan sumber kebenaran:
Saya (asisten) TIDAK punya akses ke teks literal Pedoman Penulisan Proposal
Kegiatan Mahasiswa v0.2, hanya ke daftar nama 16 unit dan dua contoh heading
("1.1 Latar Belakang" -> U03, "2.5 Struktur Kepanitiaan" -> U10) yang
diberikan di percakapan. Tabel kata kunci `_KATA_KUNCI_UNIT` di bawah adalah
tebakan terbaik berdasarkan nama unit + contoh tsb, BUKAN dikutip langsung
dari Pedoman v0.2. Peneliti WAJIB mengkalibrasi ulang tabel ini terhadap
teks asli Pedoman v0.2 dan sampel proposal riil sebelum dipakai untuk hasil
skripsi — jangan anggap tabel ini sudah final.

Asumsi:
- Input `teks_terstruktur` adalah dict keluaran `ekstrak_docx`/`ekstrak_pdf`
  (Bagian A), BUKAN path file mentah.
- U00 (Dokumen) tidak punya heading tekstual — propertinya (bahasa, kertas,
  margin, huruf, penomoran bab) diisi lewat metadata JM (Bagian C), sehingga
  di modul ini U00 selalu dicatat dengan teks kosong + metode "tidak_berlaku".
- U01 (Sampul) biasanya tidak punya heading eksplisit di badan dokumen;
  secara default seluruh konten SEBELUM heading unit pertama yang terdeteksi
  dianggap sebagai U01, kecuali ditemukan heading eksplisit untuknya.
- Unit yang sama sekali tidak terdeteksi TETAP dicatat (teks kosong,
  ditemukan=False) — bukan dihilangkan dari output — karena dibutuhkan oleh
  jalur JC (checklist keberadaan).
- Deteksi heading yang tidak lolos pola baku (nomor + kata kunci) dicatat
  sebagai fallback dengan `metode_deteksi` & `catatan` yang eksplisit,
  supaya tidak gagal diam-diam.
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Set, Tuple

from .ekstraksi_teks import _normalisasi_teks

# ---------------------------------------------------------------------------
# Daftar kanonik 16 unit anotasi
# ---------------------------------------------------------------------------

DAFTAR_UNIT: List[Tuple[str, str]] = [
    ("U00", "Dokumen"),
    ("U01", "Sampul"),
    ("U02", "Daftar Isi"),
    ("U03", "Latar Belakang"),
    ("U04", "Maksud/Tujuan/Sasaran"),
    ("U05", "Indikator Keberhasilan"),
    ("U06", "Nama Kegiatan"),
    ("U07", "Bentuk Kegiatan"),
    ("U08", "Waktu dan Tempat"),
    ("U09", "Peserta"),
    ("U10", "Struktur Kepanitiaan"),
    ("U11", "Susunan Acara"),
    ("U12", "Perencanaan Keuangan"),
    ("U13", "Penutup dan Pengesahan"),
    ("U14", "Lampiran I (Formulir Manajemen Risiko)"),
    ("U15", "Lampiran II (Timeline)"),
]

ID_UNIT_VALID: Set[str] = {uid for uid, _ in DAFTAR_UNIT}


# ---------------------------------------------------------------------------
# Pola heading per unit (lihat catatan kalibrasi di docstring modul)
# ---------------------------------------------------------------------------

# Awalan penomoran umum: "1.1 ", "2.5. ", "BAB III ", "iv. ", dll.
_POLA_AWALAN_NOMOR = r"^\s*(bab\s+)?(\d+(\.\d+){0,3}|[ivxlcdm]+)[.\)]?\s+"

_KATA_KUNCI_UNIT: Dict[str, List[str]] = {
    "U01": [r"sampul"],
    "U02": [r"daftar\s+isi"],
    "U03": [r"latar\s+belakang"],
    "U04": [
        r"maksud.{0,20}tujuan",
        r"tujuan.{0,20}sasaran",
        r"maksud\s+dan\s+tujuan",
    ],
    "U05": [r"indikator\s+keberhasilan"],
    "U06": [r"nama\s+kegiatan"],
    "U07": [r"bentuk\s+(dan\s+jenis\s+)?kegiatan"],
    "U08": [r"waktu\s+dan\s+tempat", r"waktu.{0,20}tempat\s+pelaksanaan"],
    "U09": [r"peserta\s+kegiatan", r"peserta"],
    "U10": [r"struktur\s+kepanitiaan", r"susunan\s+kepanitiaan", r"susunan\s+panitia"],
    "U11": [r"susunan\s+acara", r"rundown\s+acara", r"jadwal\s+acara"],
    "U12": [
        r"perencanaan\s+keuangan",
        r"rencana\s+anggaran",
        r"anggaran\s+(biaya|dana)",
        r"rincian\s+anggaran",
    ],
    "U13": [r"penutup", r"pengesahan"],
    "U14": [
        r"lampiran\s+(i|1)(?!\w)",
        r"formulir\s+manajemen\s+risiko",
        r"manajemen\s+risiko",
    ],
    "U15": [
        r"lampiran\s+(ii|2)(?!\w)",
        r"timeline",
        r"jadwal\s+kegiatan\s*\(?\s*timeline\s*\)?",
    ],
}

_POLA_KETAT: Dict[str, re.Pattern] = {
    uid: re.compile(_POLA_AWALAN_NOMOR + "(" + "|".join(kk) + ")", re.IGNORECASE)
    for uid, kk in _KATA_KUNCI_UNIT.items()
}
_POLA_LONGGAR: Dict[str, re.Pattern] = {
    uid: re.compile("(" + "|".join(kk) + ")", re.IGNORECASE) for uid, kk in _KATA_KUNCI_UNIT.items()
}

_PANJANG_MAKS_BARIS_FALLBACK = 120
_AKHIRAN_BUKAN_HEADING = (".", ",", ";")
_AWALAN_GAYA_HEADING = ("heading", "judul")  # "Judul" = lokalisasi Indonesia utk style Word

# Baris entri Daftar Isi (mis. "1.1 Latar Belakang ................... 3") secara
# tekstual cocok PERSIS dengan pola heading unit yang sebenarnya -- tanpa
# pengecualian ini, entri Daftar Isi akan "membajak" anchor unit tsb (karena
# anchor pertama yang menang) sebelum heading asli di badan dokumen tercapai,
# membuat teks unit itu berisi baris Daftar Isi, BUKAN paragraf sungguhan.
# Ditemukan empiris (2026-09-19) pada ~29% proposal riil yang diuji peneliti.
# Titik-titik penuntun (dot leader) sepanjang ini praktis tidak pernah muncul
# di prosa biasa, jadi dipakai sebagai penanda "ini entri Daftar Isi, bukan
# heading" -- baris yang cocok langsung dianggap BUKAN kandidat heading sama
# sekali, apa pun isinya.
_POLA_ENTRI_DAFTAR_ISI = re.compile(r"\.{4,}")


def _cari_unit_untuk_baris(teks_baris: str, gaya: Optional[str]) -> Optional[Tuple[str, str]]:
    """Mencoba mencocokkan satu baris/paragraf dengan pola heading salah satu unit.

    Prioritas pencocokan (berhenti di percobaan pertama yang berhasil):
    0. Baris berpola entri Daftar Isi (titik penuntun panjang) -> SELALU None,
       tidak pernah dianggap heading (lihat catatan `_POLA_ENTRI_DAFTAR_ISI`).
    1. Pola ketat: awalan nomor/bab + kata kunci -> metode "heading_bernomor".
    2. Kalau gaya paragraf menandakan heading Word (Heading */Judul *) ->
       coba pola longgar (tanpa syarat nomor) -> metode "heading_gaya".
    3. Pola longgar dengan batas panjang baris & tidak diakhiri tanda baca
       kalimat -> metode "heading_kata_kunci_fallback".

    Return None jika tidak ada yang cocok. Baris kosong selalu None.
    """
    baris = teks_baris.strip()
    if not baris:
        return None
    if _POLA_ENTRI_DAFTAR_ISI.search(baris):
        return None

    for unit_id in _KATA_KUNCI_UNIT:
        if _POLA_KETAT[unit_id].match(baris):
            return unit_id, "heading_bernomor"

    gaya_menandakan_heading = bool(gaya) and gaya.lower().startswith(_AWALAN_GAYA_HEADING)
    if gaya_menandakan_heading:
        for unit_id in _KATA_KUNCI_UNIT:
            if _POLA_LONGGAR[unit_id].match(baris):
                return unit_id, "heading_gaya"

    if len(baris) <= _PANJANG_MAKS_BARIS_FALLBACK and not baris.endswith(_AKHIRAN_BUKAN_HEADING):
        for unit_id in _KATA_KUNCI_UNIT:
            if _POLA_LONGGAR[unit_id].match(baris):
                return unit_id, "heading_kata_kunci_fallback"

    return None


# ---------------------------------------------------------------------------
# Adaptasi struktur hasil Bagian A (beda antara docx & pdf) ke representasi seragam
# ---------------------------------------------------------------------------


def _ambil_blok_berurutan(teks_terstruktur: Dict[str, Any], format_asli: str) -> List[Dict[str, Any]]:
    """Menyeragamkan keluaran Bagian A menjadi daftar blok teks berurutan.

    - format_asli == "docx": satu blok = satu paragraf (`gaya` = nama style
      paragraf tsb, dipakai sebagai sinyal tambahan deteksi heading).
    - format_asli == "pdf": satu blok = satu baris teks dalam satu halaman
      (`gaya` selalu None; info style tidak tersedia dari ekstraksi PDF —
      itu ranah modul JM/Bagian C, bukan modul ini).

    Urutan blok mengikuti urutan alami dokumen sehingga potongan teks antar
    dua anchor heading bisa diambil dengan slicing sederhana.
    """
    format_asli = (format_asli or "").lower()
    blok_list: List[Dict[str, Any]] = []

    if format_asli == "docx":
        for unit in teks_terstruktur.get("unit_teks", []):
            blok_list.append({"teks": unit.get("teks", ""), "gaya": unit.get("gaya")})
    elif format_asli == "pdf":
        for halaman in teks_terstruktur.get("unit_teks", []):
            for baris in halaman.get("teks", "").split("\n"):
                blok_list.append({"teks": baris, "gaya": None})
    else:
        raise ValueError(f"format_asli tidak dikenal: '{format_asli}'. Gunakan 'docx' atau 'pdf'.")

    return blok_list


def _normalisasi_gabungan(potongan: List[str]) -> str:
    """Menggabungkan beberapa potongan teks (baris/paragraf) jadi satu teks unit."""
    gabungan = "\n".join(p for p in potongan if p)
    return _normalisasi_teks(gabungan)


# ---------------------------------------------------------------------------
# Fungsi utama
# ---------------------------------------------------------------------------


def segmentasi_unit(teks_terstruktur: Dict[str, Any], format_asli: str) -> Dict[str, Any]:
    """Memecah teks hasil ekstraksi (Bagian A) menjadi 16 unit anotasi.

    Input:
        teks_terstruktur: dict keluaran `ekstrak_docx`/`ekstrak_pdf`/`ekstrak_dokumen`.
        format_asli: "docx" atau "pdf" (biasanya sama dengan
            `teks_terstruktur["format_asli"]`, dipisah sebagai parameter agar
            fungsi ini tidak diam-diam bergantung pada field internal Bagian A).

    Output (dict):
        {
            "unit": {
                "U00": {"teks": "", "ditemukan": False, "metode_deteksi": "tidak_berlaku", "catatan": "..."},
                "U01": {"teks": str, "ditemukan": bool, "metode_deteksi": str, "catatan": str | None},
                ...
                "U15": {...},
            },
            "peringatan": [str, ...],   # peringatan level dokumen (unit hilang, gagal total, dst.)
        }

    `metode_deteksi` salah satu dari:
        "tidak_berlaku"                  -> khusus U00
        "tidak_ditemukan"                -> unit tidak ada sama sekali di dokumen
        "posisi_awal_dokumen"            -> khusus U01 tanpa heading eksplisit
        "heading_bernomor"               -> pola ketat (nomor + kata kunci), paling andal
        "heading_gaya"                   -> gaya Word Heading/Judul + kata kunci, tanpa nomor
        "heading_kata_kunci_fallback"    -> kata kunci saja, tanpa nomor/gaya (paling lemah)
        "gagal_segmentasi_seluruh_dokumen" -> tidak ada heading apa pun terdeteksi

    Raises:
        ValueError: jika `format_asli` bukan "docx"/"pdf".
    """
    blok_list = _ambil_blok_berurutan(teks_terstruktur, format_asli)

    anchor: List[Tuple[int, str, str]] = []
    unit_sudah_anchor: Set[str] = set()
    for i, blok in enumerate(blok_list):
        hasil = _cari_unit_untuk_baris(blok["teks"], blok.get("gaya"))
        if hasil is None:
            continue
        unit_id, metode = hasil
        if unit_id in unit_sudah_anchor:
            continue
        anchor.append((i, unit_id, metode))
        unit_sudah_anchor.add(unit_id)
    anchor.sort(key=lambda a: a[0])

    unit_hasil: Dict[str, Dict[str, Any]] = {
        uid: {"teks": "", "ditemukan": False, "metode_deteksi": "tidak_ditemukan", "catatan": None}
        for uid, _ in DAFTAR_UNIT
    }
    unit_hasil["U00"] = {
        "teks": "",
        "ditemukan": False,
        "metode_deteksi": "tidak_berlaku",
        "catatan": (
            "U00 adalah properti tata letak keseluruhan dokumen (bahasa, kertas, "
            "margin, huruf, penomoran bab); diisi lewat metadata JM (Bagian C), "
            "bukan lewat segmentasi teks ini."
        ),
    }

    peringatan: List[str] = []

    if not anchor:
        peringatan.append(
            "Tidak ada satu pun heading unit yang terdeteksi di seluruh dokumen "
            "(kemungkinan format penomoran sangat menyimpang dari pola yang dikenali "
            "atau tabel kata kunci perlu dikalibrasi ulang). Seluruh isi disimpan "
            "sebagai U01 secara default; WAJIB ditinjau manual."
        )
        teks_gabungan = _normalisasi_gabungan([b["teks"] for b in blok_list])
        unit_hasil["U01"] = {
            "teks": teks_gabungan,
            "ditemukan": bool(teks_gabungan),
            "metode_deteksi": "gagal_segmentasi_seluruh_dokumen",
            "catatan": "Tidak ada heading terdeteksi; seluruh teks dokumen dianggap satu blok.",
        }
    else:
        indeks_anchor_pertama, unit_id_pertama, _ = anchor[0]
        if unit_id_pertama != "U01":
            teks_sampul = _normalisasi_gabungan([b["teks"] for b in blok_list[:indeks_anchor_pertama]])
            if teks_sampul:
                unit_hasil["U01"] = {
                    "teks": teks_sampul,
                    "ditemukan": True,
                    "metode_deteksi": "posisi_awal_dokumen",
                    "catatan": (
                        "U01 tidak punya heading eksplisit; diasumsikan semua konten "
                        f"sebelum heading unit pertama yang terdeteksi ({unit_id_pertama})."
                    ),
                }

        for idx_anchor, (indeks_blok, unit_id, metode) in enumerate(anchor):
            akhir = anchor[idx_anchor + 1][0] if idx_anchor + 1 < len(anchor) else len(blok_list)
            teks_unit = _normalisasi_gabungan([b["teks"] for b in blok_list[indeks_blok:akhir]])

            catatan = None
            if metode == "heading_kata_kunci_fallback":
                catatan = (
                    f"Heading '{unit_id}' terdeteksi lewat pencocokan kata kunci tanpa pola "
                    "penomoran baku (mis. mahasiswa mengubah format penomoran); perlu verifikasi manual."
                )
            elif metode == "heading_gaya":
                catatan = (
                    f"Heading '{unit_id}' terdeteksi lewat gaya paragraf Word (Heading/Judul) "
                    "tanpa pola kata kunci+nomor standar; perlu verifikasi manual."
                )

            unit_hasil[unit_id] = {
                "teks": teks_unit,
                "ditemukan": bool(teks_unit),
                "metode_deteksi": metode,
                "catatan": catatan,
            }
            if catatan:
                peringatan.append(f"[{unit_id}] {catatan}")

    unit_tidak_ditemukan = [
        uid for uid, info in unit_hasil.items() if uid != "U00" and not info["ditemukan"]
    ]
    if unit_tidak_ditemukan:
        peringatan.append(
            "Unit tidak ditemukan di dokumen ini (dicatat dengan teks kosong untuk "
            "jalur JC/checklist): " + ", ".join(sorted(unit_tidak_ditemukan))
        )

    return {"unit": unit_hasil, "peringatan": peringatan}


if __name__ == "__main__":
    # Demo pemanggilan tunggal (Bagian A + B) untuk verifikasi cepat.
    import json
    import sys

    from .ekstraksi_teks import EkstraksiError, ekstrak_dokumen

    if len(sys.argv) != 2:
        print("Pemakaian: python -m comply_proposal.preprocessing.segmentasi_unit <path_file>")
        raise SystemExit(1)

    try:
        hasil_ekstraksi = ekstrak_dokumen(sys.argv[1])
    except EkstraksiError as exc:
        print(f"Gagal ekstraksi: {exc}")
        raise SystemExit(1)

    hasil_segmentasi = segmentasi_unit(hasil_ekstraksi, hasil_ekstraksi["format_asli"])

    ringkasan = {
        uid: {
            "ditemukan": info["ditemukan"],
            "metode_deteksi": info["metode_deteksi"],
            "panjang_teks": len(info["teks"]),
            "cuplikan": info["teks"][:60],
        }
        for uid, info in hasil_segmentasi["unit"].items()
    }
    print(json.dumps({"unit": ringkasan, "peringatan": hasil_segmentasi["peringatan"]}, indent=2, ensure_ascii=False))
