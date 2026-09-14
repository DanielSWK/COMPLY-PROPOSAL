"""
Bagian F — Split dataset dev/test di level proposal, stratifikasi bahasa_asli.

Modul ini membagi daftar proposal_id menjadi dua himpunan (dev, test) di
LEVEL PROPOSAL (bukan level unit — satu proposal tidak boleh punya
sebagian unit di dev dan sebagian lain di test, karena unit-unit dalam
satu proposal saling terkait/dibandingkan pada jalur JD). Pembagian
distratifikasi berdasarkan `bahasa_asli` supaya proporsi ID/EN di dev dan
test mirip, DENGAN PENGECUALIAN: proposal berbahasa Inggris di test set
dijamin minimal `minimum_per_bahasa["EN"]` (default 8) — kalau rasio umum
(`test_ratio`) tidak cukup menghasilkan itu, proporsi test untuk bahasa
EN dinaikkan melebihi rasio umum. Kalau jumlah proposal EN yang tersedia
bahkan kurang dari minimum itu sendiri, SEMUA proposal EN dialokasikan ke
test dan hal ini dilaporkan lewat `warnings.warn` (bukan gagal diam-diam),
karena minimum tidak akan tercapai.
"""

from __future__ import annotations

import random
import warnings
from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple


def split_dataset(
    list_proposal_id: List[str],
    metadata_per_proposal: Dict[str, Dict[str, Any]],
    test_ratio: float = 0.2,
    seed: int = 42,
    minimum_per_bahasa: Optional[Dict[str, int]] = None,
    kolom_bahasa: str = "bahasa_asli",
) -> Tuple[List[str], List[str]]:
    """Membagi proposal_id jadi dev/test, stratifikasi bahasa dgn jaminan minimum EN.

    Input:
        list_proposal_id: daftar proposal_id (level proposal, bukan unit).
            Duplikat dibuang otomatis (dgn peringatan) sebelum split.
        metadata_per_proposal: dict proposal_id -> metadata (minimal berisi
            key `kolom_bahasa`, mis. "ID"/"EN"). Proposal yang tidak
            ditemukan di sini atau tidak punya nilai `kolom_bahasa` yang
            valid dikelompokkan sbg bahasa "TIDAK_DIKETAHUI" (dgn peringatan),
            BUKAN menyebabkan proses gagal.
        test_ratio: proporsi standar ke test set per kelompok bahasa
            (default 0.2 = 20%). Proporsi AKTUAL bisa lebih tinggi untuk
            kelompok yang kena aturan `minimum_per_bahasa` (lihat di bawah).
        seed: seed RNG (Python `random.Random`) supaya split deterministik/
            dapat direproduksi untuk seed yang sama.
        minimum_per_bahasa: override manual jumlah minimum proposal per
            bahasa yang WAJIB masuk test set, mis. `{"EN": 8}` (default
            kalau None). Set ke `{}` untuk menonaktifkan jaminan minimum
            sama sekali (murni `test_ratio` stratifikasi biasa utk semua
            bahasa). Bahasa yang tidak disebut di sini tidak punya jaminan
            minimum (murni ikut `test_ratio`).
        kolom_bahasa: nama key bahasa di `metadata_per_proposal[pid]`
            (default "bahasa_asli").

    Output:
        (dev_ids, test_ids) — dua list proposal_id yang saling lepas,
        gabungannya = seluruh `list_proposal_id` (setelah duplikat
        dibuang), urutan mengikuti urutan asli `list_proposal_id`.

    Efek samping: memanggil `warnings.warn()` (BUKAN mengembalikan nilai
    tambahan, supaya kontrak return `(dev_ids, test_ids)` tetap sederhana)
    untuk kondisi yang perlu ditinjau manusia: duplikat proposal_id,
    proposal tanpa metadata bahasa, dan kasus minimum per-bahasa yang
    menaikkan proporsi test di atas `test_ratio` atau bahkan tidak
    tercapai sama sekali.

    Raises:
        ValueError: `test_ratio` di luar rentang (0, 1).
    """
    if not 0 < test_ratio < 1:
        raise ValueError(f"test_ratio harus di antara 0 dan 1 (eksklusif), diberikan: {test_ratio}")

    minimum_per_bahasa = dict(minimum_per_bahasa) if minimum_per_bahasa is not None else {"EN": 8}

    daftar_unik = list(dict.fromkeys(list_proposal_id))
    if len(daftar_unik) != len(list_proposal_id):
        warnings.warn(
            f"list_proposal_id mengandung {len(list_proposal_id) - len(daftar_unik)} "
            "proposal_id duplikat; duplikat dibuang (dipertahankan kemunculan pertama) "
            "sebelum split.",
            stacklevel=2,
        )

    kelompok: Dict[str, List[str]] = defaultdict(list)
    for pid in daftar_unik:
        info = metadata_per_proposal.get(pid)
        bahasa = info.get(kolom_bahasa) if info else None
        if not bahasa:
            bahasa = "TIDAK_DIKETAHUI"
            warnings.warn(
                f"proposal_id '{pid}' tidak punya metadata '{kolom_bahasa}' yang valid "
                "di `metadata_per_proposal`; dikelompokkan sbg bahasa 'TIDAK_DIKETAHUI' "
                "untuk keperluan stratifikasi.",
                stacklevel=2,
            )
        kelompok[bahasa].append(pid)

    rng = random.Random(seed)
    dev_ids: List[str] = []
    test_ids: List[str] = []

    for bahasa, anggota in kelompok.items():
        anggota_acak = anggota[:]
        rng.shuffle(anggota_acak)

        jumlah_test_standar = round(len(anggota_acak) * test_ratio)
        jumlah_minimum = minimum_per_bahasa.get(bahasa, 0)
        jumlah_test = max(jumlah_test_standar, min(jumlah_minimum, len(anggota_acak)))

        if jumlah_minimum > len(anggota_acak):
            warnings.warn(
                f"Bahasa '{bahasa}': hanya ada {len(anggota_acak)} proposal, kurang dari "
                f"minimum yang diminta di test set ({jumlah_minimum}). SEMUA proposal "
                f"bahasa ini ({len(anggota_acak)}) dialokasikan ke test; minimum TIDAK "
                "tercapai — pertimbangkan menambah data atau menurunkan `minimum_per_bahasa`.",
                stacklevel=2,
            )
        elif jumlah_test > jumlah_test_standar:
            warnings.warn(
                f"Bahasa '{bahasa}': rasio standar ({test_ratio:.0%}) hanya menghasilkan "
                f"{jumlah_test_standar} proposal di test, di bawah minimum {jumlah_minimum}. "
                f"Proporsi test untuk bahasa ini dinaikkan jadi {jumlah_test} proposal "
                f"(~{jumlah_test / len(anggota_acak):.0%}) supaya minimum terpenuhi.",
                stacklevel=2,
            )

        test_ids.extend(anggota_acak[:jumlah_test])
        dev_ids.extend(anggota_acak[jumlah_test:])

    # Urutkan kembali sesuai urutan asli list_proposal_id (bukan urutan
    # acak per-kelompok) supaya hasilnya deterministik & enak dibaca;
    # ini TIDAK mengubah proposal mana yang masuk dev/test, hanya urutannya.
    urutan_asli = {pid: i for i, pid in enumerate(daftar_unik)}
    dev_ids.sort(key=lambda pid: urutan_asli[pid])
    test_ids.sort(key=lambda pid: urutan_asli[pid])

    return dev_ids, test_ids


if __name__ == "__main__":
    # Demo: dataset sintetis 40 proposal ID + 10 proposal EN, cek proporsi hasil split.
    import json

    metadata = {}
    daftar_id = []
    for i in range(1, 41):
        pid = f"P{i:03d}"
        metadata[pid] = {"bahasa_asli": "ID"}
        daftar_id.append(pid)
    for i in range(41, 51):
        pid = f"P{i:03d}"
        metadata[pid] = {"bahasa_asli": "EN"}
        daftar_id.append(pid)

    dev_ids, test_ids = split_dataset(daftar_id, metadata)

    def ringkas(ids):
        bahasa_count = defaultdict(int)
        for pid in ids:
            bahasa_count[metadata[pid]["bahasa_asli"]] += 1
        return {"jumlah": len(ids), "per_bahasa": dict(bahasa_count)}

    print(json.dumps({"dev": ringkas(dev_ids), "test": ringkas(test_ids)}, indent=2, ensure_ascii=False))
