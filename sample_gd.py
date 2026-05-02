"""Sinh và ghi thêm giao dịch mẫu vào sheet GiaoDich (dùng từ app hoặc chạy trực tiếp)."""
from __future__ import annotations

import random
import uuid
from datetime import date, timedelta

import pandas as pd

import excel_io as ex

LOAI_GD = [
    "Nhập kho",
    "Xuất kho",
    "Điều chuyển",
    "Kiểm kê",
    "Hủy / loại bỏ",
]


def _ensure_master_nl() -> pd.DataFrame:
    nl = ex.read_sheet(ex.SHEET_NL)
    if nl.empty or "MaNL" not in nl.columns:
        return pd.DataFrame(
            [{"MaNL": "NL001", "TenNL": "Nguyên liệu mẫu", "DonViMacDinh": "kg", "GhiChu": ""}]
        )
    return nl


def _ensure_master_kho() -> pd.DataFrame:
    kho = ex.read_sheet(ex.SHEET_KHO)
    if kho.empty or "MaKho" not in kho.columns:
        return pd.DataFrame([{"MaKho": "K01", "TenKho": "Kho mẫu", "DiaDiem": ""}])
    return kho


def _ensure_master_ncc() -> pd.DataFrame:
    ncc = ex.read_sheet(ex.SHEET_NCC)
    if ncc.empty or "MaNCC" not in ncc.columns:
        return pd.DataFrame([{"MaNCC": "NCC01", "TenNCC": "NCC mẫu", "DienThoai": ""}])
    return ncc


def _normalize_gd(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    for c in ex.GD_COLUMNS:
        if c not in d.columns:
            d[c] = ""
    return d[ex.GD_COLUMNS]


def generate_sample_rows(n: int, *, seed: int | None = 42) -> pd.DataFrame:
    if seed is not None:
        random.seed(seed)
    nl = _ensure_master_nl()
    kho = _ensure_master_kho()
    ncc = _ensure_master_ncc()
    batch = uuid.uuid4().hex[:10].upper()
    today = date.today()
    rows: list[dict] = []
    weights = [45, 35, 8, 7, 5]
    for i in range(n):
        rnl = nl.sample(n=1, random_state=None).iloc[0]
        ma_nl = str(rnl["MaNL"]).strip()
        ten_nl = str(rnl.get("TenNL", "")).strip()
        dv = str(rnl.get("DonViMacDinh", "kg")).strip() or "kg"
        rk = kho.sample(n=1).iloc[0]
        ma_k = str(rk["MaKho"]).strip()
        loai = random.choices(LOAI_GD, weights=weights, k=1)[0]
        ma_ncc = ""
        if loai == "Nhập kho" and random.random() < 0.88:
            ma_ncc = str(ncc.sample(n=1).iloc[0]["MaNCC"]).strip()
        elif loai in ("Xuất kho", "Điều chuyển") and random.random() < 0.4:
            ma_ncc = str(ncc.sample(n=1).iloc[0]["MaNCC"]).strip()
        ngay = today - timedelta(days=random.randint(0, 220))
        hsd = pd.NaT
        if loai == "Nhập kho" and random.random() < 0.72:
            hsd = pd.Timestamp(ngay + timedelta(days=random.randint(120, 600)))
        so_lo = f"L{ngay.strftime('%y%m')}-{random.randint(10000, 99999)}"
        so_luong = round(random.uniform(50.0, 28000.0), random.choice([0, 1, 2, 3]))
        ma_phieu = f"MAU-{batch}-{i + 1:04d}"
        ghi = random.choice(
            [
                "",
                "Theo kế hoạch sản xuất",
                "Ưu tiên lô A",
                "Chờ QC",
                "Đối chiếu NCC",
                "Kiểm kê định kỳ",
                "Điều chuyển nội bộ",
            ]
        )
        if random.random() < 0.35:
            ghi = ""
        rows.append(
            {
                "ID": ex.new_id(),
                "Ngay": pd.Timestamp(ngay),
                "MaPhieu": ma_phieu,
                "LoaiGiaoDich": loai,
                "MaNguyenLieu": ma_nl,
                "TenNguyenLieu": ten_nl,
                "SoLuong": float(so_luong),
                "DonVi": dv,
                "MaKho": ma_k,
                "MaNCC": ma_ncc,
                "SoLo": so_lo,
                "HanSuDung": hsd,
                "GhiChu": ghi,
            }
        )
    return pd.DataFrame(rows, columns=ex.GD_COLUMNS)


def append_sample_transactions(n: int = 200) -> int:
    """Thêm n dòng giao dịch mẫu vào cuối sheet GiaoDich. Trả về số dòng đã thêm."""
    ex.ensure_app_schema()
    old = _normalize_gd(ex.read_sheet(ex.SHEET_GD))
    new = generate_sample_rows(n, seed=None)
    merged = pd.concat([old, new], ignore_index=True)
    ex.save_giao_dich_df(merged)
    return len(new)


if __name__ == "__main__":
    k = append_sample_transactions(200)
    print(f"Da them {k} giao dich mau vao {ex.DATA_FILE}")
