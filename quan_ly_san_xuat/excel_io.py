"""Đọc/ghi file Excel quản lý sản xuất — sheet Master + lệnh sản xuất, tiêu hao, phiếu kho."""
from __future__ import annotations

import uuid
from pathlib import Path
from typing import Any, Callable

import pandas as pd

DATA_DIR = Path(__file__).resolve().parent / "data"
DATA_FILE = DATA_DIR / "san_xuat.xlsx"

SHEET_NL = "DM_NguyenLieu"
SHEET_KHO = "DM_Kho"
SHEET_TP = "DM_ThanhPham"
SHEET_DC = "DM_DayChuyen"
SHEET_CT = "DM_CongThuc"
SHEET_LSX = "LenhSanXuat"
SHEET_THSX = "TieuHaoSX"
SHEET_GD = "PhieuKho"
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

NL_COLUMNS = ["MaNL", "TenNL", "DonViMacDinh", "GhiChu"]
KHO_COLUMNS = ["MaKho", "TenKho", "DiaDiem"]
TP_COLUMNS = ["MaTP", "TenTP", "DonVi", "QuyCach", "GhiChu"]
DC_COLUMNS = ["MaDC", "TenDC", "CongSuat", "GhiChu"]
CT_COLUMNS = ["ID", "MaTP", "MaNL", "TenNL", "DinhMuc", "GhiChu"]
LSX_COLUMNS = [
    "ID",
    "MaLenh",
    "NgayKH",
    "MaTP",
    "TenTP",
    "SanLuongKH",
    "DonVi",
    "MaDC",
    "Ca",
    "MaKhoNL",
    "MaKhoTP",
    "SoLo",
    "TrangThai",
    "SanLuongTT",
    "NgayHoanThanh",
    "NguoiTao",
    "GhiChu",
]
THSX_COLUMNS = ["ID", "MaLenh", "MaNL", "TenNL", "DinhMucKH", "SoLuongTT", "DonVi"]

# Phiếu kho sinh ra từ sản xuất (Xuất NL / Nhập thành phẩm) — cùng cột với GiaoDich của app kho
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


def _default_dm_thanh_pham() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"MaTP": "TP001", "TenTP": "Cám heo thịt 15-30kg", "DonVi": "kg", "QuyCach": "Bao 25kg", "GhiChu": ""},
            {"MaTP": "TP002", "TenTP": "Cám gà đẻ", "DonVi": "kg", "QuyCach": "Bao 40kg", "GhiChu": ""},
        ]
    )


def _default_dm_day_chuyen() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {"MaDC": "DC01", "TenDC": "Dây chuyền viên 1", "CongSuat": "10", "GhiChu": "tấn/giờ"},
            {"MaDC": "DC02", "TenDC": "Dây chuyền bột 2", "CongSuat": "8", "GhiChu": "tấn/giờ"},
        ]
    )


def _default_dm_cong_thuc() -> pd.DataFrame:
    """Định mức: kg nguyên liệu cho 1.000 kg thành phẩm."""
    rows = [
        ("TP001", "NL001", "Bắp hạt", 550),
        ("TP001", "NL002", "Khô dầu đậu nành", 250),
        ("TP001", "NL003", "Cám gạo", 180),
        ("TP001", "NL004", "Premix vitamin khoáng", 20),
        ("TP002", "NL001", "Bắp hạt", 600),
        ("TP002", "NL002", "Khô dầu đậu nành", 220),
        ("TP002", "NL003", "Cám gạo", 150),
        ("TP002", "NL004", "Premix vitamin khoáng", 30),
    ]
    return pd.DataFrame(
        [
            {"ID": str(uuid.uuid4()), "MaTP": tp, "MaNL": nl, "TenNL": ten, "DinhMuc": dm, "GhiChu": ""}
            for tp, nl, ten, dm in rows
        ]
    )


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


def _sheet_defaults() -> dict[str, Callable[[], pd.DataFrame]]:
    """Mọi sheet ứng dụng cần và hàm tạo dữ liệu mặc định (thứ tự = thứ tự sheet trong file)."""
    return {
        SHEET_LSX: lambda: pd.DataFrame(columns=LSX_COLUMNS),
        SHEET_THSX: lambda: pd.DataFrame(columns=THSX_COLUMNS),
        SHEET_GD: lambda: pd.DataFrame(columns=GD_COLUMNS),
        SHEET_TP: _default_dm_thanh_pham,
        SHEET_CT: _default_dm_cong_thuc,
        SHEET_NL: _default_dm_nguyen_lieu,
        SHEET_DC: _default_dm_day_chuyen,
        SHEET_KHO: _default_dm_kho,
        SHEET_USER: _default_users_df,
    }


def _write_sheets(sheets: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as w:
        for name, d in sheets.items():
            d.to_excel(w, sheet_name=name, index=False)


def _read_all_sheets_raw() -> dict[str, pd.DataFrame]:
    if not DATA_FILE.exists():
        return {}
    return pd.read_excel(DATA_FILE, sheet_name=None, engine="openpyxl")


def ensure_app_schema() -> None:
    """Tạo file mới hoặc bổ sung các sheet còn thiếu (giữ nguyên dữ liệu đã có)."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    defaults = _sheet_defaults()
    if not DATA_FILE.exists():
        _write_sheets({name: make() for name, make in defaults.items()})
        return
    existing = pd.ExcelFile(DATA_FILE, engine="openpyxl").sheet_names
    missing = [s for s in defaults if s not in existing]
    if not missing:
        return
    all_sheets = _read_all_sheets_raw()
    for name in missing:
        all_sheets[name] = defaults[name]()
    _write_sheets(all_sheets)


def read_sheet(sheet: str) -> pd.DataFrame:
    ensure_app_schema()
    return pd.read_excel(DATA_FILE, sheet_name=sheet, engine="openpyxl")


def save_sheets(sheets: dict[str, pd.DataFrame]) -> None:
    """Ghi đồng thời nhiều sheet trong một lần ghi file (các sheet khác giữ nguyên)."""
    ensure_app_schema()
    all_sheets = _read_all_sheets_raw()
    all_sheets.update(sheets)
    _write_sheets(all_sheets)


def save_master_sheet(sheet_name: str, df: pd.DataFrame) -> None:
    save_sheets({sheet_name: df})


def new_id() -> str:
    return str(uuid.uuid4())


def row_from_form(values: dict[str, Any], existing_id: str | None) -> dict[str, Any]:
    row = {k: values.get(k, "") for k in GD_COLUMNS}
    row["ID"] = existing_id or new_id()
    return row
