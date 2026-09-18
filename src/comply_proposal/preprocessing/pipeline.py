"""
Bagian G — Orkestrasi Bagian A-F + format output akhir.

Modul ini merangkai seluruh tahap sebelumnya (A: ekstraksi teks, B:
segmentasi unit, C: metadata JM, D: gabung ground truth, E: statistik,
F: split dev/test) jadi satu pipeline yang bisa dijalankan atas satu
folder proposal + satu workbook ground truth, menghasilkan:
- satu file JSONL (satu baris per (proposal_id, unit_id)) untuk dikonsumsi
  pipeline RAG, dan
- satu file log audit TEKS BIASA (bukan JSONL, bukan data) berisi seluruh
  warning/error per proposal, untuk ditinjau manusia.

Prinsip "satu proposal gagal, batch tetap lanjut" diterapkan di
`proses_folder_proposal`: kegagalan total satu file (mis. file corrupt)
DITANGKAP di situ, dicatat ke log, dan proses lanjut ke file berikutnya —
lihat komentar di dalam fungsi tsb untuk penjelasan kenapa `except
Exception` yang luas SENGAJA dipakai di titik itu saja.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Union

import pandas as pd

from .ekstraksi_teks import EkstraksiError, ekstrak_dokumen
from .ground_truth import (
    GroundTruthError,
    gabungkan_teks_dan_ground_truth,
    muat_ground_truth,
)
from .metadata_tata_letak import ekstrak_metadata_dokumen
from .segmentasi_unit import segmentasi_unit
from .split_dataset import split_dataset
from .statistik import hitung_statistik

PathLike = Union[str, Path]


# ---------------------------------------------------------------------------
# Metadata proposal (sheet "Daftar_Proposal")
# ---------------------------------------------------------------------------


def muat_metadata_proposal(
    path_excel: PathLike,
    nama_sheet: str = "Daftar_Proposal",
    kolom_proposal_id: str = "proposal_id",
) -> Dict[str, Dict[str, Any]]:
    """Memuat metadata per proposal (bahasa_asli, dst.) dari sheet "Daftar_Proposal".

    Setiap baris sheet menjadi satu entri `{proposal_id: {nama_kolom: nilai, ...}}`
    — SELURUH kolom disertakan apa adanya (bukan cuma bahasa_asli), supaya
    fungsi ini tetap berguna kalau sheet Anda punya kolom metadata lain.

    PENTING (lihat juga catatan kalibrasi serupa di Bagian D): nama sheet &
    nama kolom proposal_id di sini adalah TEBAKAN berdasar deskripsi umum
    ("field bahasa_asli sudah ada di metadata proposal, sheet Daftar_Proposal")
    — saya tidak punya skema kolom literal sheet ini. Sesuaikan
    `nama_sheet`/`kolom_proposal_id` kalau berbeda di file Anda.

    Raises:
        GroundTruthError: sheet tidak bisa dibaca atau kolom proposal_id tidak ada.
    """
    path_excel = Path(path_excel)
    try:
        df = pd.read_excel(path_excel, sheet_name=nama_sheet, engine="openpyxl", dtype=str)
    except Exception as exc:
        raise GroundTruthError(
            f"Gagal membaca sheet '{nama_sheet}' dari '{path_excel.name}': {exc}"
        ) from exc

    if kolom_proposal_id not in df.columns:
        raise GroundTruthError(f"Kolom '{kolom_proposal_id}' tidak ditemukan di sheet '{nama_sheet}'.")

    hasil: Dict[str, Dict[str, Any]] = {}
    for _, baris in df.iterrows():
        pid = str(baris[kolom_proposal_id]).strip()
        if not pid or pid.lower() == "nan":
            continue
        hasil[pid] = baris.to_dict()
    return hasil


# ---------------------------------------------------------------------------
# Override status NA-01 (lihat catatan di Bagian A: keputusan akhir NA-01
# diambil di tahap ini, bukan di modul ekstraksi)
# ---------------------------------------------------------------------------


def _terapkan_override_na01(record_list: List[Dict[str, Any]], proposal_id: str, log: List[str]) -> None:
    """Meng-override `status_ground_truth` -> "NA-01" untuk SEMUA unit satu
    proposal, dipanggil ketika Bagian A mendeteksi `kemungkinan_hasil_pindai=True`
    pada file PDF-nya.

    Ini SENGAJA mengubah `record_list` in-place (dipanggil tepat sebelum
    `proses_satu_proposal` mengembalikan hasilnya). Rasional: kalau teks
    tidak berhasil diekstrak dengan andal dari file hasil pindai, evaluasi
    JP/JC/JD terhadap teks itu tidak bermakna, sehingga proposal ini lebih
    tepat ditandai NA-01 daripada memakai status dari ground truth apa
    adanya. Status ASLI (sebelum override) tetap dicatat ke log supaya
    keputusan ini bisa diaudit/dibatalkan manual kalau ternyata keliru
    (mis. PDF yang genuinely sedikit teksnya tapi bukan hasil scan).
    """
    for rec in record_list:
        status_asli = rec["status_ground_truth"]
        if status_asli != "NA-01":
            log.append(
                f"[{proposal_id}][{rec['unit_id']}] status_ground_truth di-override dari "
                f"'{status_asli}' menjadi 'NA-01' karena ekstraksi teks (Bagian A) mengindikasikan "
                "file ini kemungkinan hasil pindai/scan."
            )
        rec["status_ground_truth"] = "NA-01"
        rec["kode_pelanggaran"] = []
        catatan_tambahan = "Status di-override jadi NA-01: proposal terdeteksi kemungkinan hasil pindai/scan (Bagian A)."
        rec["catatan_ekstraksi"] = (
            f"{rec['catatan_ekstraksi']}; {catatan_tambahan}" if rec["catatan_ekstraksi"] else catatan_tambahan
        )


# ---------------------------------------------------------------------------
# Satu proposal end-to-end (A -> B -> C -> D)
# ---------------------------------------------------------------------------


def proses_satu_proposal(
    path_file: PathLike,
    proposal_id: str,
    gt_dataframe: pd.DataFrame,
    bahasa_asli: Optional[str] = None,
    peta_kolom: Optional[Dict[str, str]] = None,
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """Memproses SATU file proposal dari Bagian A sampai D, hasilkan 16 record + log.

    Input:
        path_file: path ke file proposal (.docx/.pdf).
        proposal_id: ID proposal (biasanya nama file tanpa ekstensi).
        gt_dataframe: keluaran `muat_ground_truth()` (Bagian D).
        bahasa_asli: "ID"/"EN" dari metadata proposal, kalau ada.
        peta_kolom: diteruskan ke `gabungkan_teks_dan_ground_truth` (Bagian D).

    Output:
        (record_list, log_proposal) — 16 record (satu per unit, format
        Bagian D) dan list pesan log KHUSUS proposal ini (sudah diberi
        prefix "[proposal_id]"/"[proposal_id][unit_id]").

    Penanganan kegagalan (lihat juga docstring modul):
        - Kegagalan Bagian A (ekstraksi teks) atau Bagian B (segmentasi)
          DIBIARKAN merambat (raise) ke pemanggil, karena tanpa teks tidak
          ada apa pun yang bisa diproses secara bermakna untuk proposal
          ini — pemanggil batch (`proses_folder_proposal`) yang
          bertanggung jawab menangkapnya supaya batch tetap lanjut.
        - Kegagalan Bagian C (metadata JM) DITANGKAP DI SINI dan
          didegradasi jadi `metadata_jm = None` + catatan log, karena
          jalur JP/JC/JD masih bisa jalan tanpa metadata tata letak —
          kehilangan JM tidak seharusnya menggagalkan seluruh proposal.

    Raises:
        EkstraksiError: kegagalan Bagian A.
        (Exception lain dari Bagian B/D diteruskan apa adanya.)
    """
    log: List[str] = []

    hasil_ekstraksi = ekstrak_dokumen(path_file)
    log.extend(f"[{proposal_id}] {p}" for p in hasil_ekstraksi.get("peringatan", []))

    hasil_segmentasi = segmentasi_unit(hasil_ekstraksi, hasil_ekstraksi["format_asli"])
    log.extend(f"[{proposal_id}] {p}" for p in hasil_segmentasi.get("peringatan", []))

    try:
        hasil_metadata = ekstrak_metadata_dokumen(path_file)
        log.extend(f"[{proposal_id}] {p}" for p in hasil_metadata.get("peringatan", []))
    except EkstraksiError as exc:
        log.append(
            f"[{proposal_id}] Gagal ekstraksi metadata JM (Bagian C): {exc}. "
            "metadata_jm akan bernilai None untuk U00 proposal ini."
        )
        hasil_metadata = None

    record_list = gabungkan_teks_dan_ground_truth(
        hasil_segmentasi,
        gt_dataframe,
        proposal_id,
        metadata_jm=hasil_metadata,
        bahasa_asli=bahasa_asli,
        format_asli=hasil_ekstraksi["format_asli"],
        peta_kolom=peta_kolom,
    )

    for rec in record_list:
        if rec["catatan_ekstraksi"]:
            log.append(f"[{proposal_id}][{rec['unit_id']}] {rec['catatan_ekstraksi']}")

    if hasil_ekstraksi["format_asli"] == "pdf" and hasil_ekstraksi.get("kemungkinan_hasil_pindai"):
        _terapkan_override_na01(record_list, proposal_id, log)

    return record_list, log


# ---------------------------------------------------------------------------
# Batch: seluruh folder
# ---------------------------------------------------------------------------


def proses_folder_proposal(
    folder_proposal: PathLike,
    path_ground_truth_excel: PathLike,
    metadata_proposal: Optional[Dict[str, Dict[str, Any]]] = None,
    peta_kolom: Optional[Dict[str, str]] = None,
    pola_glob: str = "*.docx,*.pdf",
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """Memproses SEMUA file proposal (.docx/.pdf campur) di satu folder.

    Input:
        folder_proposal: folder berisi file proposal. `proposal_id` diambil
            dari nama file TANPA ekstensi (mis. "P001.docx" -> proposal_id
            "P001") — pastikan skema penamaan file ini konsisten dengan
            proposal_id di ground truth. CATATAN: tahap redaksi PII terpisah
            TIDAK dipakai di implementasi ini (lihat catatan asumsi di
            docstring Bagian A/`ekstraksi_teks.py`) — isi file proposal
            TIDAK dianonimkan oleh pipeline ini.
        path_ground_truth_excel: workbook ground truth (dibaca sekali via
            `muat_ground_truth`, dipakai untuk semua proposal).
        metadata_proposal: dict proposal_id -> metadata (mis. dari
            `muat_metadata_proposal`), dipakai mengisi `bahasa_asli`. Kalau
            None, `bahasa_asli` semua proposal akan None.
        peta_kolom: diteruskan ke `muat_ground_truth`/`gabungkan_teks_dan_ground_truth`.
        pola_glob: pola glob file yang diproses, dipisah koma (default
            "*.docx,*.pdf").

    Output:
        (semua_record, semua_log) — gabungan record SELURUH proposal (siap
        ditulis lewat `tulis_output_jsonl`) dan gabungan log SELURUH
        proposal + pesan tingkat-batch (siap ditulis lewat `tulis_log_audit`).

    Raises:
        GroundTruthError: kalau workbook ground truth SENDIRI gagal dimuat
            (ini kegagalan tingkat BATCH, bukan tingkat proposal — tanpa
            ground truth, tidak ada satu proposal pun yang bisa diproses
            secara bermakna, jadi wajar dihentikan di sini, bukan di-skip
            per proposal).

    Kegagalan MEMPROSES SATU FILE PROPOSAL (docx/pdf corrupt, format tak
    dikenal, dll.) TIDAK menghentikan batch: ditangkap, dicatat ke log
    dengan detail error, lanjut ke file berikutnya. Exception generik
    (`Exception`, bukan cuma `EkstraksiError`) SENGAJA ditangkap di titik
    ini SATU-SATUNYA supaya bug tak terduga pada satu file pun tidak
    menggagalkan seluruh batch pemrosesan puluhan/ratusan proposal lain —
    di luar titik ini (Bagian A-F, fungsi single-proposal di atas), error
    dibiarkan merambat normal supaya mudah dites/didebug.
    """
    folder_proposal = Path(folder_proposal)
    metadata_proposal = metadata_proposal or {}

    try:
        gt_dataframe = muat_ground_truth(path_ground_truth_excel, peta_kolom=peta_kolom)
    except GroundTruthError as exc:
        raise GroundTruthError(
            f"Gagal memuat ground truth dari '{path_ground_truth_excel}', seluruh batch "
            f"dibatalkan (tanpa ground truth tidak ada yang bisa diproses bermakna): {exc}"
        ) from exc

    semua_record: List[Dict[str, Any]] = []
    semua_log: List[str] = [
        f"Ringkasan validasi ground truth: {gt_dataframe.attrs.get('ringkasan_validasi')}",
    ]

    daftar_file = sorted(
        {p for pola in pola_glob.split(",") for p in folder_proposal.glob(pola.strip())}
    )
    if not daftar_file:
        semua_log.append(
            f"Tidak ada file proposal ditemukan di folder '{folder_proposal}' (pola: {pola_glob})."
        )

    for path_file in daftar_file:
        proposal_id = path_file.stem
        info_proposal = metadata_proposal.get(proposal_id, {})
        bahasa_asli = info_proposal.get("bahasa_asli")

        try:
            record_list, log_proposal = proses_satu_proposal(
                path_file, proposal_id, gt_dataframe, bahasa_asli=bahasa_asli, peta_kolom=peta_kolom
            )
        except Exception as exc:  # noqa: BLE001 -- lihat penjelasan di docstring
            semua_log.append(
                f"[{proposal_id}] GAGAL DIPROSES TOTAL ({type(exc).__name__}): {exc}. "
                "Proposal ini TIDAK menghasilkan record apa pun di output JSONL; "
                "perlu ditinjau manual (file corrupt? format tak terduga?)."
            )
            continue

        semua_record.extend(record_list)
        semua_log.extend(log_proposal)

    return semua_record, semua_log


# ---------------------------------------------------------------------------
# Penulisan output akhir
# ---------------------------------------------------------------------------


def tulis_output_jsonl(record_list: List[Dict[str, Any]], path_output: PathLike) -> None:
    """Menulis `record_list` ke file JSONL, satu baris JSON per (proposal_id, unit_id).

    Field per record mengikuti keluaran `gabungkan_teks_dan_ground_truth`
    (Bagian D): proposal_id, unit_id, bahasa_asli, format_asli, teks_unit,
    metadata_jm, status_ground_truth, kode_pelanggaran, dasar_pedoman,
    bukti, keyakinan, anotator, status_adjudikasi, catatan_ekstraksi.
    File ini yang dikonsumsi pipeline RAG.
    """
    path_output = Path(path_output)
    with path_output.open("w", encoding="utf-8") as f:
        for rec in record_list:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")


def tulis_log_audit(daftar_log: List[str], path_log: PathLike) -> None:
    """Menulis seluruh pesan log ke file TEKS BIASA (bukan JSONL, bukan data
    untuk pipeline RAG) — satu pesan per baris, untuk ditinjau manusia
    (mis. proposal mana yang NA-01, unit mana yang gagal terdeteksi
    headingnya, kolom ground truth mana yang bermasalah, dst).
    """
    path_log = Path(path_log)
    with path_log.open("w", encoding="utf-8") as f:
        for baris in daftar_log:
            f.write(str(baris) + "\n")


if __name__ == "__main__":
    # Contoh pemanggilan end-to-end: satu folder proposal + satu workbook
    # ground truth -> JSONL + log audit, lalu tampilkan statistik (Bagian E)
    # dan split dev/test (Bagian F) sebagai demonstrasi pipeline penuh.
    import sys

    if len(sys.argv) != 4:
        print(
            "Pemakaian: python -m comply_proposal.preprocessing.pipeline "
            "<folder_proposal> <ground_truth.xlsx> <folder_output>"
        )
        raise SystemExit(1)

    folder_proposal_cli, path_gt_cli, folder_output_cli = sys.argv[1], sys.argv[2], sys.argv[3]
    folder_output_path = Path(folder_output_cli)
    folder_output_path.mkdir(parents=True, exist_ok=True)

    try:
        metadata_proposal_cli = muat_metadata_proposal(path_gt_cli)
    except GroundTruthError as exc:
        print(
            f"Peringatan: gagal memuat sheet Daftar_Proposal ({exc}); "
            "bahasa_asli akan kosong untuk semua proposal."
        )
        metadata_proposal_cli = {}

    try:
        semua_record, semua_log = proses_folder_proposal(
            folder_proposal_cli, path_gt_cli, metadata_proposal=metadata_proposal_cli
        )
    except GroundTruthError as exc:
        print(f"Batch dibatalkan: {exc}")
        raise SystemExit(1)

    path_jsonl = folder_output_path / "dataset_uji.jsonl"
    path_log = folder_output_path / "log_audit.txt"
    tulis_output_jsonl(semua_record, path_jsonl)
    tulis_log_audit(semua_log, path_log)
    print(f"Selesai: {len(semua_record)} record ditulis ke {path_jsonl}")
    print(f"Log audit ({len(semua_log)} baris) ditulis ke {path_log}")

    statistik = hitung_statistik(semua_record)
    print("\nStatistik dataset (Bagian E):")
    print(json.dumps(statistik, indent=2, ensure_ascii=False))

    daftar_proposal_id = sorted({rec["proposal_id"] for rec in semua_record})
    dev_ids, test_ids = split_dataset(daftar_proposal_id, metadata_proposal_cli)
    print(f"\nSplit dev/test (Bagian F): {len(dev_ids)} dev, {len(test_ids)} test")
    print("dev:", dev_ids)
    print("test:", test_ids)
