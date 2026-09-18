"""
Bagian D — Pemetaan ground truth.

Dua tanggung jawab:
1. `muat_ground_truth`: membaca & memvalidasi sheet "Anotasi" dari workbook
   ground truth (proposal_id, unit_id, kode_pelanggaran, dst.).
2. `gabungkan_teks_dan_ground_truth`: menggabungkan hasil segmentasi teks
   Bagian B (+ metadata JM Bagian C bila relevan) dengan ground truth,
   menghasilkan satu record per (proposal_id, unit_id).

PEMBARUAN (2026-09-18) — sumber kebenaran terkonfirmasi:
Setelah memeriksa langsung workbook ground truth asli peneliti (sheet
"Referensi" dan "Anotasi_Gabungan"), dua hal di bawah ini SUDAH TIDAK
lagi tebakan, melainkan dikonfirmasi dari data asli:
- `KODE_TAKSONOMI_AKTIF_DEFAULT`: 21 kode taksonomi yang benar-benar aktif
  (dikutip dari sheet "Referensi", kolom "DAFTAR KODE PELANGGARAN", status
  "AKTIF" — mengecualikan FMT-03 & BHS-02 yang eksplisit ditandai
  dinonaktifkan/"gunakan NA-03"). Ini sekarang jadi DEFAULT `muat_ground_truth`,
  bukan lagi opsional.
- `PETA_KOLOM_ANOTASI_DEFAULT` & `nama_sheet="Anotasi_Gabungan"`: nama
  kolom dan nama sheet di bawah sekarang mencerminkan struktur ASLI
  workbook peneliti (proposal_id, unit_id, kode_pelanggaran, bukti,
  dasar_pedoman, keyakinan, anotator, catatan, status_adjudikasi).

Catatan: ini dikonfirmasi untuk WORKBOOK SPESIFIK peneliti ini (skripsi
ini), bukan klaim skema 11-kolom resmi Panduan Anotasi v0.2 berlaku
universal. Kalau workbook lain/versi lebih baru punya struktur berbeda,
tetap gunakan parameter `peta_kolom`/`nama_sheet`/`kode_taksonomi_valid`
untuk menyesuaikan — jangan asumsikan default ini akan selalu cocok.

Asumsi lain:
- Satu baris di sheet "Anotasi" = satu instans anotasi untuk satu
  (proposal_id, unit_id); kolom `kode_pelanggaran` boleh berisi lebih dari
  satu kode dalam satu sel (dipisah koma/titik-koma).
- Kalau tidak ada baris ground truth sama sekali untuk suatu
  (proposal_id, unit_id), itu DICATAT sebagai status "TIDAK_ADA_ANOTASI"
  (bukan diam-diam dilewati), karena bisa berarti unit itu belum
  dianotasi ATAU memang tidak relevan — perlu ditinjau manusia.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import AbstractSet, Any, Dict, FrozenSet, List, Optional, Set, Tuple, Union

import pandas as pd

from .ekstraksi_teks import EkstraksiError
from .segmentasi_unit import DAFTAR_UNIT, ID_UNIT_VALID

PathLike = Union[str, Path]


class GroundTruthError(EkstraksiError):
    """Kesalahan saat memuat/memvalidasi berkas ground truth.

    Subclass dari `EkstraksiError` (Bagian A) supaya pemanggil batch
    (Bagian G) bisa menangani kegagalan di modul A-D lewat satu jenis
    exception yang sama tanpa menghentikan seluruh proses.
    """


# ---------------------------------------------------------------------------
# Skema kolom & kode (lihat catatan kalibrasi di docstring modul)
# ---------------------------------------------------------------------------

PETA_KOLOM_ANOTASI_DEFAULT: Dict[str, str] = {
    "proposal_id": "proposal_id",
    "unit_id": "unit_id",
    "kode_pelanggaran": "kode_pelanggaran",
    "jalur_deteksi": "jalur_deteksi",
    "dasar_pedoman": "dasar_pedoman",
    "tanggal_anotasi": "tanggal_anotasi",
    "catatan": "catatan",
    # Ditambahkan setelah memeriksa workbook asli (lihat "PEMBARUAN" di
    # docstring modul) — semua opsional, dilewati kalau kolomnya tidak ada.
    "bukti": "bukti",
    "keyakinan": "keyakinan",
    "anotator": "anotator",
    "status_adjudikasi": "status_adjudikasi",
}

_KOLOM_WAJIB = ("proposal_id", "unit_id", "kode_pelanggaran")

# Dikutip langsung dari sheet "Referensi" workbook ground truth peneliti
# (kolom "DAFTAR KODE PELANGGARAN", status "AKTIF"). FMT-03 dan BHS-02
# SENGAJA tidak disertakan -- keduanya ditandai eksplisit "dinonaktifkan,
# gunakan NA-03" di sheet yang sama. frozenset supaya aman dipakai sbg
# nilai default parameter (immutable, tidak kena masalah mutable default).
KODE_TAKSONOMI_AKTIF_DEFAULT: FrozenSet[str] = frozenset(
    {
        "KEL-01", "KEL-02", "KEL-03", "KEL-04", "KEL-05", "KEL-06",
        "ISI-01", "ISI-02", "ISI-03",
        "ANG-01", "ANG-02",
        "KON-01", "KON-02", "KON-03", "KON-04",
        "ADM-01", "ADM-02", "ADM-04",
        "FMT-01", "FMT-02", "FMT-04",
        "BHS-01",
    }
)

_STATUS_KHUSUS: Set[str] = {"SESUAI", "NA-01", "NA-02", "NA-03"}
_POLA_KODE_TAKSONOMI = re.compile(r"^(KEL|ISI|ANG|KON|ADM|FMT|BHS)-\d{2}$")
_POLA_JALUR_DETEKSI = re.compile(r"^(JP|JC|JD|JM)$", re.IGNORECASE)
_POLA_DELIMITER_MULTI_NILAI = re.compile(r"[,;]\s*")


def _adalah_kode_valid(kode: str, kode_taksonomi_valid: Optional[AbstractSet[str]]) -> bool:
    """True jika `kode` adalah status khusus, ATAU kode taksonomi yang dikenali.

    Kalau `kode_taksonomi_valid` diberikan, dipakai sebagai keanggotaan set
    yang ketat (sumber kebenaran dari peneliti). Kalau tidak, jatuh ke
    validasi pola generik (lihat catatan kalibrasi di docstring modul).
    """
    if kode in _STATUS_KHUSUS:
        return True
    if kode_taksonomi_valid is not None:
        return kode in kode_taksonomi_valid
    return bool(_POLA_KODE_TAKSONOMI.match(kode))


def _pecah_nilai_multi(nilai: Any) -> List[str]:
    """Memecah satu sel yang mungkin berisi beberapa nilai dipisah koma/titik-koma."""
    if nilai is None or (isinstance(nilai, float) and pd.isna(nilai)) or pd.isna(nilai):
        return []
    teks = str(nilai).strip()
    if not teks:
        return []
    return [bagian.strip() for bagian in _POLA_DELIMITER_MULTI_NILAI.split(teks) if bagian.strip()]


def _ambil_nilai_tunggal(nilai: Any) -> Optional[str]:
    """Mengambil satu sel sebagai string tunggal APA ADANYA (TIDAK dipecah
    koma/titik-koma seperti `_pecah_nilai_multi`) -- dipakai untuk kolom
    teks bebas/kategorikal (mis. `bukti`, `keyakinan`) yang bisa secara sah
    mengandung koma sbg tanda baca biasa, bukan pemisah multi-nilai.
    """
    if nilai is None or pd.isna(nilai):
        return None
    teks = str(nilai).strip()
    return teks or None


# ---------------------------------------------------------------------------
# D1: muat_ground_truth
# ---------------------------------------------------------------------------


def muat_ground_truth(
    path_excel: PathLike,
    peta_kolom: Optional[Dict[str, str]] = None,
    kode_taksonomi_valid: Optional[AbstractSet[str]] = KODE_TAKSONOMI_AKTIF_DEFAULT,
    nama_sheet: str = "Anotasi_Gabungan",
) -> pd.DataFrame:
    """Memuat & memvalidasi sheet ground truth dari workbook Excel.

    Input:
        path_excel: path ke workbook ground truth (.xlsx).
        peta_kolom: pemetaan nama kolom logis -> nama kolom aktual di file
            Anda (mengganti sebagian/seluruh `PETA_KOLOM_ANOTASI_DEFAULT`).
            Default-nya sudah cocok utk workbook peneliti (lihat "PEMBARUAN"
            di docstring modul); sesuaikan kalau workbook Anda berbeda.
        kode_taksonomi_valid: set kode taksonomi aktif yang dipakai untuk
            validasi keanggotaan ketat. Default `KODE_TAKSONOMI_AKTIF_DEFAULT`
            (21 kode aktif asli, lihat docstring modul). Pass `None` secara
            eksplisit untuk kembali ke validasi pola generik yang lebih
            longgar (kalau memang perlu memproses workbook dgn daftar kode
            berbeda yang belum diketahui).
        nama_sheet: nama sheet ground truth di workbook (default
            "Anotasi_Gabungan", sesuai workbook peneliti).

    Output:
        pandas.DataFrame — semua baris asli TETAP disertakan apa adanya
        (baris bermasalah TIDAK dibuang diam-diam), ditambah 2 kolom:
        - "_baris_excel": nomor baris asli di file Excel (1-based, sudah
          memperhitungkan baris header), untuk audit manual.
        - "_masalah_validasi": list[str] berisi pesan masalah validasi
          baris tsb (list kosong kalau baris valid).
        Ringkasan agregat validasi tersimpan di `DataFrame.attrs["ringkasan_validasi"]`
        (dict: sheet, jumlah_baris, jumlah_baris_bermasalah, kolom_dipakai).

    Validasi per baris:
        - unit_id harus salah satu dari 16 unit (U00-U15), tidak
          peka-huruf-besar/kecil untuk pengecekan (nilai asli tidak diubah).
        - kode_pelanggaran (boleh multi-nilai dipisah koma/titik-koma) harus
          SESUAI/NA-01/NA-02/NA-03 atau lolos `_adalah_kode_valid`.
        - jalur_deteksi (kalau kolomnya ada) harus salah satu JP/JC/JD/JM.
        - tanggal_anotasi (kalau kolomnya ada) harus bisa diparse sebagai
          tanggal (dayfirst=True, sesuai konvensi Indonesia).

    Raises:
        GroundTruthError: file/sheet tidak bisa dibaca, atau kolom WAJIB
            (proposal_id, unit_id, kode_pelanggaran — setelah `peta_kolom`
            diterapkan) tidak ditemukan di sheet.
    """
    peta = {**PETA_KOLOM_ANOTASI_DEFAULT, **(peta_kolom or {})}
    path_excel = Path(path_excel)

    try:
        # dtype=str: semua kolom dibaca sbg teks apa adanya supaya regex
        # validasi & pemecahan multi-nilai konsisten (tanggal tetap bisa
        # diparse ulang lewat pd.to_datetime di bawah).
        df = pd.read_excel(path_excel, sheet_name=nama_sheet, engine="openpyxl", dtype=str)
    except Exception as exc:
        raise GroundTruthError(
            f"Gagal membaca sheet '{nama_sheet}' dari '{path_excel.name}': {exc}"
        ) from exc

    kolom_wajib_hilang = [peta[k] for k in _KOLOM_WAJIB if peta[k] not in df.columns]
    if kolom_wajib_hilang:
        raise GroundTruthError(
            f"Kolom wajib tidak ditemukan di sheet '{nama_sheet}': {kolom_wajib_hilang}. "
            "Kalau nama kolom di file Excel Anda berbeda, gunakan parameter `peta_kolom` "
            "untuk memetakan nama kolom aktual (lihat PETA_KOLOM_ANOTASI_DEFAULT)."
        )

    df = df.reset_index(drop=True)
    df["_baris_excel"] = df.index + 2  # +2: 1-based, baris 1 = header

    masalah_per_baris: List[List[str]] = []
    for _, baris in df.iterrows():
        masalah: List[str] = []

        unit_id_mentah = baris.get(peta["unit_id"])
        unit_id_norm = str(unit_id_mentah).strip().upper() if pd.notna(unit_id_mentah) else ""
        if unit_id_norm not in ID_UNIT_VALID:
            masalah.append(f"unit_id tidak valid: '{unit_id_mentah}'")

        daftar_kode = _pecah_nilai_multi(baris.get(peta["kode_pelanggaran"]))
        if not daftar_kode:
            masalah.append("kode_pelanggaran kosong")
        else:
            for kode in daftar_kode:
                if not _adalah_kode_valid(kode.upper(), kode_taksonomi_valid):
                    masalah.append(f"kode_pelanggaran tidak dikenali: '{kode}'")

        kolom_jalur = peta.get("jalur_deteksi")
        if kolom_jalur and kolom_jalur in df.columns:
            nilai_jalur = baris.get(kolom_jalur)
            if pd.notna(nilai_jalur) and str(nilai_jalur).strip():
                if not _POLA_JALUR_DETEKSI.match(str(nilai_jalur).strip()):
                    masalah.append(f"jalur_deteksi tidak dikenali: '{nilai_jalur}' (harus JP/JC/JD/JM)")

        kolom_tanggal = peta.get("tanggal_anotasi")
        if kolom_tanggal and kolom_tanggal in df.columns:
            nilai_tanggal = baris.get(kolom_tanggal)
            if pd.notna(nilai_tanggal) and str(nilai_tanggal).strip():
                tanggal_terparse = pd.to_datetime(nilai_tanggal, errors="coerce", dayfirst=True)
                if pd.isna(tanggal_terparse):
                    masalah.append(f"tanggal_anotasi tidak valid: '{nilai_tanggal}'")

        masalah_per_baris.append(masalah)

    df["_masalah_validasi"] = masalah_per_baris

    jumlah_bermasalah = sum(1 for m in masalah_per_baris if m)
    df.attrs["ringkasan_validasi"] = {
        "sheet": nama_sheet,
        "jumlah_baris": len(df),
        "jumlah_baris_bermasalah": jumlah_bermasalah,
        "kolom_dipakai": peta,
    }

    return df


# ---------------------------------------------------------------------------
# D2: gabungkan_teks_dan_ground_truth
# ---------------------------------------------------------------------------


def _ringkas_status_dan_kode(semua_nilai: List[str]) -> Tuple[str, List[str]]:
    """Meringkas semua nilai kode_pelanggaran (gabungan seluruh baris ground
    truth) untuk satu (proposal_id, unit_id) menjadi satu status ringkas +
    daftar kode pelanggaran nyata (SESUAI/NA-0X TIDAK termasuk di daftar
    ini karena bukan kode pelanggaran, hanya status).

    Prioritas: ada kode pelanggaran nyata (pola taksonomi) > NA-01/02/03 >
    SESUAI > tidak ada anotasi sama sekali.
    """
    semua_nilai_upper = [v.upper() for v in semua_nilai]

    kode_nyata: List[str] = []
    terlihat: Set[str] = set()
    for v in semua_nilai_upper:
        if _POLA_KODE_TAKSONOMI.match(v) and v not in terlihat:
            kode_nyata.append(v)
            terlihat.add(v)
    if kode_nyata:
        return "PELANGGARAN", kode_nyata

    for status_na in ("NA-01", "NA-02", "NA-03"):
        if status_na in semua_nilai_upper:
            return status_na, []

    if "SESUAI" in semua_nilai_upper:
        return "SESUAI", []

    return "TIDAK_ADA_ANOTASI", []


def gabungkan_teks_dan_ground_truth(
    hasil_segmentasi: Dict[str, Any],
    gt_dataframe: pd.DataFrame,
    proposal_id: str,
    metadata_jm: Optional[Dict[str, Any]] = None,
    bahasa_asli: Optional[str] = None,
    format_asli: Optional[str] = None,
    peta_kolom: Optional[Dict[str, str]] = None,
) -> List[Dict[str, Any]]:
    """Menggabungkan teks unit (Bagian B) + ground truth (D1) jadi 1 record/unit.

    Input:
        hasil_segmentasi: keluaran `segmentasi_unit()` (Bagian B) — dict
            `{"unit": {unit_id: {...}}, "peringatan": [...]}`. Dict polos
            `{unit_id: {"teks": ...}}` juga diterima untuk fleksibilitas.
        gt_dataframe: keluaran `muat_ground_truth()` (D1).
        proposal_id: ID proposal yang sedang diproses (dicocokkan ke kolom
            proposal_id di `gt_dataframe`, dibandingkan sbg string setelah di-strip).
        metadata_jm: dict keluaran `ekstrak_metadata_docx`/`ekstrak_metadata_pdf`
            (Bagian C). Hanya dilekatkan ke record U00 (unit lain -> None),
            karena hanya U00 yang merepresentasikan properti tata letak
            keseluruhan dokumen.
        bahasa_asli: "ID"/"EN" dari metadata proposal (sheet Daftar_Proposal).
        format_asli: "docx"/"pdf" (biasanya `hasil_ekstraksi["format_asli"]` dari Bagian A).
        peta_kolom: sama seperti di `muat_ground_truth` — HARUS konsisten
            dengan peta_kolom yang dipakai saat memanggil `muat_ground_truth`,
            supaya nama kolom yang dicocokkan sama.

    Output:
        list[dict], SELALU 16 record (satu per unit U00-U15, urutan
        `DAFTAR_UNIT`), masing-masing:
        {
            "proposal_id": str,
            "unit_id": str,
            "bahasa_asli": str | None,
            "format_asli": str | None,
            "teks_unit": str,
            "status_ground_truth": "PELANGGARAN"|"SESUAI"|"NA-01"|"NA-02"|"NA-03"|"TIDAK_ADA_ANOTASI",
            "kode_pelanggaran": list[str],      # kosong kalau bukan PELANGGARAN
            "dasar_pedoman": list[str],         # rujukan Bab/Pasal/ayat, gabungan semua baris terkait
            "bukti": list[str],                  # kutipan/evidence per baris anotasi terkait unit ini
            "keyakinan": list[str],              # mis. "Tinggi"/"Sedang"/"Rendah", satu per baris
            "anotator": list[str],               # mis. "A1"/"A2"/"A1 + A2", satu per baris
            "status_adjudikasi": list[str],      # mis. "Disepakati"/"Diubah", satu per baris (kalau kolomnya ada & terisi)
                                                  # ^ keempat list di atas (dasar_pedoman s.d. status_adjudikasi)
                                                  # dikumpulkan APA ADANYA per baris ground truth yang cocok,
                                                  # TIDAK dijamin berpasangan 1:1 dgn tiap kode di 'kode_pelanggaran'
                                                  # kalau satu baris berisi banyak kode sekaligus dalam satu sel.
            "metadata_jm": dict | None,          # hanya diisi utk unit_id == "U00"
            "catatan_ekstraksi": str | None,     # gabungan warning Bagian B + masalah validasi GT
        }

    Raises:
        GroundTruthError: kalau `gt_dataframe` tidak punya kolom proposal_id/unit_id
            yang dipetakan (mis. dipanggil dgn DataFrame yang bukan dari `muat_ground_truth`).
    """
    peta = {**PETA_KOLOM_ANOTASI_DEFAULT, **(peta_kolom or {})}
    kolom_proposal_id = peta["proposal_id"]
    kolom_unit_id = peta["unit_id"]
    kolom_kode = peta["kode_pelanggaran"]
    kolom_dasar = peta.get("dasar_pedoman")
    kolom_bukti = peta.get("bukti")
    kolom_keyakinan = peta.get("keyakinan")
    kolom_anotator = peta.get("anotator")
    kolom_status_adjudikasi = peta.get("status_adjudikasi")

    for kolom in (kolom_proposal_id, kolom_unit_id, kolom_kode):
        if kolom not in gt_dataframe.columns:
            raise GroundTruthError(
                f"Kolom '{kolom}' tidak ditemukan di gt_dataframe. Pastikan `peta_kolom` "
                "yang dipakai di sini sama dengan yang dipakai saat `muat_ground_truth()`."
            )

    unit_map: Dict[str, Any] = hasil_segmentasi.get("unit", hasil_segmentasi)

    baris_proposal = gt_dataframe[
        gt_dataframe[kolom_proposal_id].astype(str).str.strip() == str(proposal_id).strip()
    ]

    hasil: List[Dict[str, Any]] = []
    for unit_id, _nama_unit in DAFTAR_UNIT:
        info_unit = unit_map.get(unit_id, {}) or {}
        teks_unit = info_unit.get("teks", "")
        catatan_segmentasi = info_unit.get("catatan")

        baris_unit = baris_proposal[
            baris_proposal[kolom_unit_id].astype(str).str.strip().str.upper() == unit_id
        ]

        semua_nilai_kode: List[str] = []
        daftar_dasar: List[str] = []
        daftar_bukti: List[str] = []
        daftar_keyakinan: List[str] = []
        daftar_anotator: List[str] = []
        daftar_status_adjudikasi: List[str] = []
        masalah_gt_unit: List[str] = []
        for _, baris in baris_unit.iterrows():
            semua_nilai_kode.extend(_pecah_nilai_multi(baris.get(kolom_kode)))
            if kolom_dasar and kolom_dasar in gt_dataframe.columns:
                daftar_dasar.extend(_pecah_nilai_multi(baris.get(kolom_dasar)))
            if kolom_bukti and kolom_bukti in gt_dataframe.columns:
                nilai = _ambil_nilai_tunggal(baris.get(kolom_bukti))
                if nilai:
                    daftar_bukti.append(nilai)
            if kolom_keyakinan and kolom_keyakinan in gt_dataframe.columns:
                nilai = _ambil_nilai_tunggal(baris.get(kolom_keyakinan))
                if nilai:
                    daftar_keyakinan.append(nilai)
            if kolom_anotator and kolom_anotator in gt_dataframe.columns:
                nilai = _ambil_nilai_tunggal(baris.get(kolom_anotator))
                if nilai:
                    daftar_anotator.append(nilai)
            if kolom_status_adjudikasi and kolom_status_adjudikasi in gt_dataframe.columns:
                nilai = _ambil_nilai_tunggal(baris.get(kolom_status_adjudikasi))
                if nilai:
                    daftar_status_adjudikasi.append(nilai)
            if "_masalah_validasi" in gt_dataframe.columns:
                masalah_gt_unit.extend(baris.get("_masalah_validasi") or [])

        status_gt, kode_pelanggaran = _ringkas_status_dan_kode(semua_nilai_kode)
        daftar_dasar_unik = list(dict.fromkeys(daftar_dasar))
        daftar_bukti_unik = list(dict.fromkeys(daftar_bukti))
        daftar_keyakinan_unik = list(dict.fromkeys(daftar_keyakinan))
        daftar_anotator_unik = list(dict.fromkeys(daftar_anotator))
        daftar_status_adjudikasi_unik = list(dict.fromkeys(daftar_status_adjudikasi))

        catatan_bagian: List[str] = []
        if catatan_segmentasi:
            catatan_bagian.append(catatan_segmentasi)
        if baris_unit.empty:
            catatan_bagian.append(
                f"Tidak ada baris anotasi ground truth untuk (proposal_id={proposal_id}, unit_id={unit_id})."
            )
        if masalah_gt_unit:
            catatan_bagian.append("Masalah validasi ground truth: " + "; ".join(masalah_gt_unit))

        hasil.append(
            {
                "proposal_id": proposal_id,
                "unit_id": unit_id,
                "bahasa_asli": bahasa_asli,
                "format_asli": format_asli,
                "teks_unit": teks_unit,
                "status_ground_truth": status_gt,
                "kode_pelanggaran": kode_pelanggaran,
                "dasar_pedoman": daftar_dasar_unik,
                "bukti": daftar_bukti_unik,
                "keyakinan": daftar_keyakinan_unik,
                "anotator": daftar_anotator_unik,
                "status_adjudikasi": daftar_status_adjudikasi_unik,
                "metadata_jm": metadata_jm if unit_id == "U00" else None,
                "catatan_ekstraksi": "; ".join(catatan_bagian) if catatan_bagian else None,
            }
        )

    return hasil


if __name__ == "__main__":
    # Demo pemanggilan tunggal (Bagian A + B + C jika docx + D) untuk verifikasi cepat.
    import json
    import sys

    from .ekstraksi_teks import ekstrak_dokumen
    from .metadata_tata_letak import ekstrak_metadata_dokumen
    from .segmentasi_unit import segmentasi_unit

    if len(sys.argv) != 4:
        print(
            "Pemakaian: python -m comply_proposal.preprocessing.ground_truth "
            "<path_file_proposal> <path_ground_truth.xlsx> <proposal_id>"
        )
        raise SystemExit(1)

    path_proposal, path_gt, proposal_id_cli = sys.argv[1], sys.argv[2], sys.argv[3]

    try:
        hasil_ekstraksi = ekstrak_dokumen(path_proposal)
        hasil_metadata = ekstrak_metadata_dokumen(path_proposal)
    except EkstraksiError as exc:
        print(f"Gagal ekstraksi: {exc}")
        raise SystemExit(1)

    hasil_segmentasi = segmentasi_unit(hasil_ekstraksi, hasil_ekstraksi["format_asli"])

    try:
        gt_df = muat_ground_truth(path_gt)
    except GroundTruthError as exc:
        print(f"Gagal memuat ground truth: {exc}")
        raise SystemExit(1)

    print("Ringkasan validasi ground truth:", gt_df.attrs.get("ringkasan_validasi"))

    record_list = gabungkan_teks_dan_ground_truth(
        hasil_segmentasi,
        gt_df,
        proposal_id_cli,
        metadata_jm=hasil_metadata,
        format_asli=hasil_ekstraksi["format_asli"],
    )
    print(json.dumps(record_list, indent=2, ensure_ascii=False, default=str))
