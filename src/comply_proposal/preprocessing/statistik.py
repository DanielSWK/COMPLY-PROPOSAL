"""
Bagian E — Statistik dataset (kebutuhan revisi pembimbing, C2).

Modul ini menghitung ringkasan statistik dari kumpulan record hasil Bagian D
(`gabungkan_teks_dan_ground_truth`), DIGABUNG LINTAS SELURUH PROPOSAL dalam
dataset — dipakai untuk melaporkan profil dataset ke pembimbing: jumlah
proposal, distribusi pelanggaran per kode/kategori/unit, dan kode taksonomi
dengan instans terlalu sedikit untuk dianalisis terpisah (kandidat digabung
kategori "lain-lain").

Asumsi:
- `list_record` adalah gabungan list[dict] dari SEMUA proposal (hasil
  `gabungkan_teks_dan_ground_truth` dipanggil berulang per proposal lalu
  di-`extend()` jadi satu list besar), BUKAN hanya satu proposal.
- `kode_pelanggaran` pada tiap record HANYA berisi kode taksonomi asli
  (mis. "ISI-02"), bukan status seperti SESUAI/NA-0X — ini dijamin oleh
  `_ringkas_status_dan_kode` di Bagian D — sehingga kategori (KEL/ISI/ANG/
  KON/ADM/FMT/BHS) bisa diambil langsung dari prefiks sebelum "-".
- Statistik bahasa_asli/format_asli dihitung PER PROPOSAL (bukan per unit),
  supaya tidak "dobel-hitung" hingga 16x untuk proposal yang sama (tiap
  proposal punya 16 record, satu per unit, dengan bahasa_asli/format_asli
  yang identik di semua unit-nya).
"""

from __future__ import annotations

from collections import Counter, defaultdict
from typing import Any, Dict, List, Set


def _kategori_dari_kode(kode: str) -> str:
    """Mengambil kategori taksonomi (KEL/ISI/ANG/KON/ADM/FMT/BHS) dari kode, mis. "ISI-02" -> "ISI"."""
    return kode.split("-", 1)[0] if "-" in kode else "TIDAK_DIKETAHUI"


def hitung_statistik(
    list_record: List[Dict[str, Any]],
    ambang_instans_rendah: int = 8,
) -> Dict[str, Any]:
    """Menghitung ringkasan statistik dataset dari kumpulan record Bagian D.

    Input:
        list_record: gabungan record `gabungkan_teks_dan_ground_truth` lintas
            SEMUA proposal dalam dataset (bukan satu proposal saja).
        ambang_instans_rendah: batas jumlah instans (EKSKLUSIF) — kode
            dengan jumlah instans < ambang ini ditandai sebagai kandidat
            digabung ke kategori "lain-lain" (default 8, sesuai kebutuhan
            revisi C2).

    Output (dict):
        {
            "jumlah_proposal": int,                     # jumlah proposal_id unik
            "jumlah_unit_total": int,                   # = len(list_record)
            "jumlah_pelanggaran_per_kode": {kode: int},  # terurut menurun berdasar jumlah, lalu abjad
            "rata_rata_pelanggaran_per_proposal": float,
            "kode_instans_rendah": [kode, ...],          # jumlah < ambang_instans_rendah, terurut abjad

            # --- statistik tambahan, membantu profil dataset utk revisi C2 ---
            "jumlah_pelanggaran_per_kategori": {kategori: int},
            "jumlah_pelanggaran_per_unit": {unit_id: int},
            "distribusi_status_ground_truth": {status: int},   # dihitung per record (proposal, unit)
            "distribusi_bahasa_asli": {bahasa: int},            # PER PROPOSAL, bukan per unit
            "distribusi_format_asli": {format_asli: int},       # PER PROPOSAL, bukan per unit
            "proposal_tanpa_anotasi_sama_sekali": [proposal_id, ...],  # semua unit-nya TIDAK_ADA_ANOTASI

            "peringatan": [str, ...],   # record cacat (field hilang/tipe salah) dilaporkan di sini, bukan diam-diam dilewati
        }
    """
    peringatan: List[str] = []

    if not list_record:
        peringatan.append("list_record kosong; seluruh statistik dikembalikan sebagai nol/kosong.")
        return {
            "jumlah_proposal": 0,
            "jumlah_unit_total": 0,
            "jumlah_pelanggaran_per_kode": {},
            "rata_rata_pelanggaran_per_proposal": 0.0,
            "kode_instans_rendah": [],
            "jumlah_pelanggaran_per_kategori": {},
            "jumlah_pelanggaran_per_unit": {},
            "distribusi_status_ground_truth": {},
            "distribusi_bahasa_asli": {},
            "distribusi_format_asli": {},
            "proposal_tanpa_anotasi_sama_sekali": [],
            "peringatan": peringatan,
        }

    proposal_ids: Set[str] = set()
    hitung_kode: Counter = Counter()
    hitung_kategori: Counter = Counter()
    hitung_per_unit: Counter = Counter()
    hitung_status: Counter = Counter()
    bahasa_per_proposal: Dict[Any, Any] = {}
    format_per_proposal: Dict[Any, Any] = {}
    status_per_proposal: Dict[Any, List[str]] = defaultdict(list)

    for i, rec in enumerate(list_record):
        proposal_id = rec.get("proposal_id")
        unit_id = rec.get("unit_id")
        if proposal_id is None or unit_id is None:
            peringatan.append(
                f"Record indeks {i} tidak punya proposal_id/unit_id yang valid, dilewati dari statistik."
            )
            continue
        proposal_ids.add(proposal_id)

        status = rec.get("status_ground_truth", "TIDAK_DIKETAHUI")
        hitung_status[status] += 1
        status_per_proposal[proposal_id].append(status)

        kode_list = rec.get("kode_pelanggaran") or []
        if not isinstance(kode_list, list):
            peringatan.append(
                f"Record (proposal_id={proposal_id}, unit_id={unit_id}): 'kode_pelanggaran' "
                f"bertipe {type(kode_list).__name__} (bukan list), dilewati dari hitungan kode."
            )
            kode_list = []
        for kode in kode_list:
            hitung_kode[kode] += 1
            hitung_kategori[_kategori_dari_kode(kode)] += 1
            hitung_per_unit[unit_id] += 1

        if proposal_id not in bahasa_per_proposal:
            bahasa_per_proposal[proposal_id] = rec.get("bahasa_asli")
        if proposal_id not in format_per_proposal:
            format_per_proposal[proposal_id] = rec.get("format_asli")

    jumlah_proposal = len(proposal_ids)
    total_pelanggaran = sum(hitung_kode.values())
    rata_rata = (total_pelanggaran / jumlah_proposal) if jumlah_proposal else 0.0

    jumlah_pelanggaran_per_kode = dict(sorted(hitung_kode.items(), key=lambda kv: (-kv[1], kv[0])))
    kode_instans_rendah = sorted(
        kode for kode, jumlah in hitung_kode.items() if jumlah < ambang_instans_rendah
    )

    proposal_tanpa_anotasi = sorted(
        pid
        for pid, daftar_status in status_per_proposal.items()
        if daftar_status and all(s == "TIDAK_ADA_ANOTASI" for s in daftar_status)
    )
    if proposal_tanpa_anotasi:
        peringatan.append(
            "Proposal berikut sama sekali tidak punya anotasi ground truth (seluruh "
            "unit-nya berstatus TIDAK_ADA_ANOTASI), kemungkinan belum dianotasi: "
            + ", ".join(str(p) for p in proposal_tanpa_anotasi)
        )

    return {
        "jumlah_proposal": jumlah_proposal,
        "jumlah_unit_total": len(list_record),
        "jumlah_pelanggaran_per_kode": jumlah_pelanggaran_per_kode,
        "rata_rata_pelanggaran_per_proposal": round(rata_rata, 3),
        "kode_instans_rendah": kode_instans_rendah,
        "jumlah_pelanggaran_per_kategori": dict(
            sorted(hitung_kategori.items(), key=lambda kv: (-kv[1], kv[0]))
        ),
        "jumlah_pelanggaran_per_unit": dict(sorted(hitung_per_unit.items())),
        "distribusi_status_ground_truth": dict(
            sorted(hitung_status.items(), key=lambda kv: (-kv[1], kv[0]))
        ),
        "distribusi_bahasa_asli": dict(Counter(bahasa_per_proposal.values())),
        "distribusi_format_asli": dict(Counter(format_per_proposal.values())),
        "proposal_tanpa_anotasi_sama_sekali": proposal_tanpa_anotasi,
        "peringatan": peringatan,
    }


if __name__ == "__main__":
    # Demo: baca file JSONL berisi record Bagian D (satu record per baris,
    # lintas banyak proposal -- lihat Bagian G untuk cara menghasilkan file ini)
    # lalu cetak statistiknya.
    import json
    import sys

    if len(sys.argv) != 2:
        print("Pemakaian: python -m comply_proposal.preprocessing.statistik <path_file.jsonl>")
        raise SystemExit(1)

    daftar_record: List[Dict[str, Any]] = []
    with open(sys.argv[1], "r", encoding="utf-8") as f:
        for baris in f:
            baris = baris.strip()
            if baris:
                daftar_record.append(json.loads(baris))

    hasil = hitung_statistik(daftar_record)
    print(json.dumps(hasil, indent=2, ensure_ascii=False))
