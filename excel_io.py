"""Đọc/ghi file Excel kho thức ăn — nhiều sheet GiaoDich + Master."""
from __future__ import annotations

import os
import uuid
from pathlib import Path
from typing import Any

import pandas as pd

DATA_DIR = Path(__file__).resolve().parent / "data"
DATA_FILE = DATA_DIR / "kho_thuc_an.xlsx"

SHEET_GD = "GiaoDich"
SHEET_NL = "DM_NguyenLieu"
SHEET_KHO = "DM_Kho"
SHEET_NCC = "DM_NhaCungCap"
SHEET_DV = "DM_DonVi"
SHEET_USER = "DM_NguoiDung"

USER_COLUMNS = [
    "ID",
    "MaNV",
    "HoTen",
    "VaiTro",
    "TenDangNhap",
    "MatKhauHash",
    "Salt",
    "KichHoat",
    "GhiChu",
]

GD_COLUMNS = [
    "ID",
    "Ngay",
    "MaPhieu",
    "LoaiGiaoDich",
    "MaNguyenLieu",
    "TenNguyenLieu",
    "SoLuong",
    "DonVi",
    "MaKho",
    "MaNCC",
    "SoLo",
    "HanSuDung",
    "GhiChu",
]


def _default_dm_nguyen_lieu() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"MaNL": "NL001", "TenNL": "Bắp hạt", "DonViMacDinh": "kg", "GhiChu": ""},
            {"MaNL": "NL002", "TenNL": "Khô dầu đậu nành", "DonViMacDinh": "kg", "GhiChu": ""},
            {"MaNL": "NL003", "TenNL": "Cám gạo", "DonViMacDinh": "kg", "GhiChu": ""},
            {"MaNL": "NL004", "TenNL": "Premix vitamin khoáng", "DonViMacDinh": "kg", "GhiChu": ""},
        ]
    )


def _default_dm_kho() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"MaKho": "K01", "TenKho": "Kho nguyên liệu A", "DiaDiem": "Nhà máy 1"},
            {"MaKho": "K02", "TenKho": "Kho thành phẩm", "DiaDiem": "Nhà máy 1"},
            {"MaKho": "K03", "TenKho": "Kho phụ gia", "DiaDiem": "Nhà máy 1"},
        ]
    )


def _default_dm_ncc() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"MaNCC": "NCC01", "TenNCC": "Công ty Thương mại ABC", "DienThoai": ""},
            {"MaNCC": "NCC02", "TenNCC": "Nhà cung cấp XYZ", "DienThoai": ""},
        ]
    )


def _default_dm_donvi() -> pd.DataFrame:
    return pd.DataFrame([{"MaDV": "kg", "TenDV": "Kilogram"}, {"MaDV": "tan", "TenDV": "Tấn"}])


def _empty_giao_dich() -> pd.DataFrame:
    return pd.DataFrame(columns=GD_COLUMNS)


def _default_users_df() -> pd.DataFrame:
    import auth as auth_mod

    salt, pw_hash = auth_mod.hash_password("admin123")
    return pd.DataFrame(
        [
            {
                "ID": str(uuid.uuid4()),
                "MaNV": "NV000",
                "HoTen": "Quản trị hệ thống",
                "VaiTro": "QuanTri",
                "TenDangNhap": "admin",
                "MatKhauHash": pw_hash,
                "Salt": salt,
                "KichHoat": "Co",
                "GhiChu": "Đổi mật khẩu sau khi triển khai (mặc định: admin123)",
            }
        ]
    )


def ensure_workbook() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if DATA_FILE.exists():
        return
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as w:
        _empty_giao_dich().to_excel(w, sheet_name=SHEET_GD, index=False)
        _default_dm_nguyen_lieu().to_excel(w, sheet_name=SHEET_NL, index=False)
        _default_dm_kho().to_excel(w, sheet_name=SHEET_KHO, index=False)
        _default_dm_ncc().to_excel(w, sheet_name=SHEET_NCC, index=False)
        _default_dm_donvi().to_excel(w, sheet_name=SHEET_DV, index=False)
        _default_users_df().to_excel(w, sheet_name=SHEET_USER, index=False)


def _read_all_sheets_raw() -> dict[str, pd.DataFrame]:
    """Đọc toàn bộ sheet hiện có (không gọi migration — tránh đệ quy)."""
    if not DATA_FILE.exists():
        return {}
    xl = pd.ExcelFile(DATA_FILE, engine="openpyxl")
    return {s: pd.read_excel(DATA_FILE, sheet_name=s, engine="openpyxl") for s in xl.sheet_names}


def ensure_app_schema() -> None:
    """Tạo file mới hoặc bổ sung sheet người dùng cho file Excel đã có từ trước."""
    ensure_workbook()
    if not DATA_FILE.exists():
        return
    xl = pd.ExcelFile(DATA_FILE, engine="openpyxl")
    if SHEET_USER in xl.sheet_names:
        return
    all_sheets = _read_all_sheets_raw()
    all_sheets[SHEET_USER] = _default_users_df()
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as w:
        for name, d in all_sheets.items():
            d.to_excel(w, sheet_name=name, index=False)


def read_sheet(sheet: str) -> pd.DataFrame:
    ensure_app_schema()
    return pd.read_excel(DATA_FILE, sheet_name=sheet, engine="openpyxl")


def write_all_sheets(sheets: dict[str, pd.DataFrame]) -> None:
    ensure_app_schema()
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name, index=False)


def read_all_for_save() -> dict[str, pd.DataFrame]:
    ensure_app_schema()
    return _read_all_sheets_raw()


def save_giao_dich_df(df: pd.DataFrame) -> None:
    all_sheets = read_all_for_save()
    all_sheets[SHEET_GD] = df
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as w:
        for name, d in all_sheets.items():
            d.to_excel(w, sheet_name=name, index=False)


def save_master_sheet(sheet_name: str, df: pd.DataFrame) -> None:
    all_sheets = read_all_for_save()
    all_sheets[sheet_name] = df
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as w:
        for name, d in all_sheets.items():
            d.to_excel(w, sheet_name=name, index=False)


def new_id() -> str:
    return str(uuid.uuid4())


def row_from_form(values: dict[str, Any], existing_id: str | None) -> dict[str, Any]:
    row = {k: values.get(k, "") for k in GD_COLUMNS}
    row["ID"] = existing_id or new_id()
    return row


def enrich_ten_nl(row: dict[str, Any], dm_nl: pd.DataFrame) -> dict[str, Any]:
    ma = str(row.get("MaNguyenLieu", "") or "").strip()
    if not ma or dm_nl.empty or "MaNL" not in dm_nl.columns:
        return row
    hit = dm_nl[dm_nl["MaNL"].astype(str).str.strip() == ma]
    if not hit.empty and "TenNL" in hit.columns:
        row["TenNguyenLieu"] = str(hit.iloc[0]["TenNL"])
    return row
