"""Template Excel và kiểm tra dữ liệu nhập giao dịch hàng loạt."""
from __future__ import annotations

from io import BytesIO
from typing import Any

import pandas as pd

import excel_io as ex

# Trùng với app.LOAI_GD_OPTIONS — tránh import streamlit
LOAI_GD_OPTIONS = [
    "Nhập kho",
    "Xuất kho",
    "Điều chuyển",
    "Kiểm kê",
    "Hủy / loại bỏ",
]

SHEET_NHAP = "NhapGiaoDich"
SHEET_HD = "HuongDan"
SHEET_DM_NL = "DM_NguyenLieu"
SHEET_DM_KHO = "DM_Kho"
SHEET_DM_NCC = "DM_NhaCungCap"

# Cột người dùng điền (không có ID — hệ thống sinh)
GD_IMPORT_COLS = [
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

HUONG_DAN_ROWS = [
    {
        "Cot": "Ngay",
        "BatBuoc": "Co",
        "MoTa": "Ngày chứng từ (Excel date hoặc dd/mm/yyyy)",
        "ViDu": "2026-05-01",
    },
    {
        "Cot": "MaPhieu",
        "BatBuoc": "Co",
        "MoTa": "Mã phiếu duy nhất; không trùng trong file và không trùng dữ liệu đã có",
        "ViDu": "PN-2026-001",
    },
    {
        "Cot": "LoaiGiaoDich",
        "BatBuoc": "Co",
        "MoTa": "Chính xác một trong các giá trị cho phép (xem sheet HuongDan hoặc nhập tay đúng chữ)",
        "ViDu": "Nhập kho",
    },
    {
        "Cot": "MaNguyenLieu",
        "BatBuoc": "Co",
        "MoTa": "Phải trùng MaNL trong sheet DM_NguyenLieu",
        "ViDu": "NL001",
    },
    {
        "Cot": "TenNguyenLieu",
        "BatBuoc": "Khong",
        "MoTa": "Để trống sẽ tự điền theo DM; có thể ghi tay để đối chiếu",
        "ViDu": "",
    },
    {
        "Cot": "SoLuong",
        "BatBuoc": "Co",
        "MoTa": "Số > 0",
        "ViDu": "1500.5",
    },
    {
        "Cot": "DonVi",
        "BatBuoc": "Khong",
        "MoTa": "Để trống lấy ĐVT mặc định từ DM nguyên liệu",
        "ViDu": "kg",
    },
    {
        "Cot": "MaKho",
        "BatBuoc": "Co",
        "MoTa": "Phải trùng MaKho trong sheet DM_Kho",
        "ViDu": "K01",
    },
    {
        "Cot": "MaNCC",
        "BatBuoc": "Khong",
        "MoTa": "Nên có với phiếu nhập; phải trùng MaNCC trong DM nếu điền",
        "ViDu": "NCC01",
    },
    {
        "Cot": "SoLo",
        "BatBuoc": "Khong",
        "MoTa": "Số lô hàng",
        "ViDu": "L2505-10001",
    },
    {
        "Cot": "HanSuDung",
        "BatBuoc": "Khong",
        "MoTa": "HSD lô (date hoặc để trống)",
        "ViDu": "",
    },
    {
        "Cot": "GhiChu",
        "BatBuoc": "Khong",
        "MoTa": "Ghi chú thêm",
        "ViDu": "",
    },
    {
        "Cot": "(Dong mau)",
        "BatBuoc": "",
        "MoTa": "Cac dong co MaPhieu bat dau bang ~ la VI DU — xoa truoc khi nhap that",
        "ViDu": "~XOA-DONG-NAY",
    },
]


def _example_rows() -> pd.DataFrame:
    """Hai dòng mẫu; người dùng xóa hoặc thay (mã bắt đầu ~ có thể bị bỏ qua khi import nếu tùy chọn)."""
    return pd.DataFrame(
        [
            {
                "Ngay": pd.Timestamp(2026, 5, 1),
                "MaPhieu": "~VI-DU-1-XOA",
                "LoaiGiaoDich": "Nhập kho",
                "MaNguyenLieu": "NL001",
                "TenNguyenLieu": "",
                "SoLuong": 2500.0,
                "DonVi": "kg",
                "MaKho": "K01",
                "MaNCC": "NCC01",
                "SoLo": "L2505-90001",
                "HanSuDung": pd.Timestamp(2027, 4, 30),
                "GhiChu": "Dòng ví dụ — xóa trước khi import thật",
            },
            {
                "Ngay": pd.Timestamp(2026, 5, 2),
                "MaPhieu": "~VI-DU-2-XOA",
                "LoaiGiaoDich": "Xuất kho",
                "MaNguyenLieu": "NL002",
                "TenNguyenLieu": "",
                "SoLuong": 800.0,
                "DonVi": "kg",
                "MaKho": "K02",
                "MaNCC": "",
                "SoLo": "L2504-80002",
                "HanSuDung": "",
                "GhiChu": "",
            },
        ]
    )


def build_import_template_bytes(
    dm_nl: pd.DataFrame,
    dm_kho: pd.DataFrame,
    dm_ncc: pd.DataFrame,
) -> bytes:
    """File .xlsx: sheet nhập + hướng dẫn + bản sao DM để đối chiếu."""
    buf = BytesIO()
    hd = pd.DataFrame(HUONG_DAN_ROWS)
    hd2 = pd.DataFrame(
        {
            "LoaiGiaoDich_cho_phep": LOAI_GD_OPTIONS,
        }
    )
    empty = pd.DataFrame(columns=GD_IMPORT_COLS)
    demo = _example_rows()
    data_sheet = pd.concat([demo, empty], ignore_index=True)

    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        data_sheet.to_excel(w, sheet_name=SHEET_NHAP, index=False)
        hd.to_excel(w, sheet_name=SHEET_HD, index=False)
        hd2.to_excel(w, sheet_name="LoaiGiaoDich_Mau", index=False)
        if not dm_nl.empty:
            dm_nl.to_excel(w, sheet_name=SHEET_DM_NL, index=False)
        if not dm_kho.empty:
            dm_kho.to_excel(w, sheet_name=SHEET_DM_KHO, index=False)
        if not dm_ncc.empty:
            dm_ncc.to_excel(w, sheet_name=SHEET_DM_NCC, index=False)
    return buf.getvalue()


def read_import_sheet(content: bytes) -> pd.DataFrame:
    bio = BytesIO(content)
    xl = pd.ExcelFile(bio, engine="openpyxl")
    name = SHEET_NHAP if SHEET_NHAP in xl.sheet_names else xl.sheet_names[0]
    return pd.read_excel(bio, sheet_name=name, engine="openpyxl")


def _norm_headers(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    d.columns = [str(c).strip().lstrip("\ufeff") for c in d.columns]
    return d


def _parse_ngay(val: Any) -> pd.Timestamp | None:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, pd.Timestamp):
        return val.normalize()
    ts = pd.to_datetime(val, dayfirst=True, errors="coerce")
    if pd.isna(ts):
        return None
    return ts.normalize()


def _parse_hsd(val: Any) -> Any:
    if val is None or (isinstance(val, float) and pd.isna(val)) or str(val).strip() == "":
        return pd.NaT
    ts = pd.to_datetime(val, dayfirst=True, errors="coerce")
    if pd.isna(ts):
        return pd.NaT
    return ts.normalize()


def validate_and_build_rows(
    df_raw: pd.DataFrame,
    dm_nl: pd.DataFrame,
    dm_kho: pd.DataFrame,
    dm_ncc: pd.DataFrame,
    existing_ma_phieu: set[str],
    *,
    skip_demo_rows: bool = True,
) -> tuple[list[dict[str, Any]], list[str]]:
    """
    Trả về (danh_sach_dong_hop_le, loi).
    Mỗi dict trong danh sách khớp ex.GD_COLUMNS (có ID mới).
    """
    errors: list[str] = []
    df = _norm_headers(df_raw)
    missing = [c for c in GD_IMPORT_COLS if c not in df.columns]
    if missing:
        errors.append(
            f"Thiếu cột bắt buộc trong file: {', '.join(missing)}. Tải lại template chuẩn và không đổi tên cột."
        )
        return [], errors

    nl = dm_nl.copy()
    kho = dm_kho.copy()
    ncc = dm_ncc.copy()
    nl_codes = set(nl["MaNL"].astype(str).str.strip()) if not nl.empty and "MaNL" in nl.columns else set()
    kho_codes = set(kho["MaKho"].astype(str).str.strip()) if not kho.empty and "MaKho" in kho.columns else set()
    ncc_codes = set(ncc["MaNCC"].astype(str).str.strip()) if not ncc.empty and "MaNCC" in ncc.columns else set()
    nl_lookup: dict[str, pd.Series] = {}
    if not nl.empty and "MaNL" in nl.columns:
        for _, r in nl.iterrows():
            m = str(r.get("MaNL", "")).strip()
            if m and m not in nl_lookup:
                nl_lookup[m] = r

    loai_set = set(LOAI_GD_OPTIONS)
    pending: list[dict[str, Any]] = []
    seen_phieu: set[str] = set()

    def _row_empty(r: pd.Series) -> bool:
        mp = str(r.get("MaPhieu", "") or "").strip()
        mn = str(r.get("MaNguyenLieu", "") or "").strip()
        if mp or mn:
            return False
        slv = r.get("SoLuong")
        if slv is None or (isinstance(slv, float) and pd.isna(slv)):
            return True
        try:
            return float(slv) == 0
        except (TypeError, ValueError):
            return True

    for pos, (_, row) in enumerate(df.iterrows()):
        excel_row = pos + 2
        if _row_empty(row):
            continue

        ma_phieu = str(row.get("MaPhieu", "") or "").strip()
        ma_nl = str(row.get("MaNguyenLieu", "") or "").strip()

        if skip_demo_rows and ma_phieu.startswith("~"):
            continue

        re: list[str] = []
        if not ma_phieu:
            re.append(f"Dòng {excel_row}: Mã phiếu trống.")
        elif ma_phieu in seen_phieu:
            re.append(f"Dòng {excel_row}: Mã phiếu '{ma_phieu}' trùng trong file.")
        elif ma_phieu in existing_ma_phieu:
            re.append(f"Dòng {excel_row}: Mã phiếu '{ma_phieu}' đã tồn tại trong hệ thống.")
        if re:
            errors.extend(re)
            continue
        seen_phieu.add(ma_phieu)

        ngay = _parse_ngay(row.get("Ngay"))
        if ngay is None:
            errors.append(f"Dòng {excel_row}: Ngày không hợp lệ (dùng định dạng ngày hoặc dd/mm/yyyy).")
            continue

        loai = str(row.get("LoaiGiaoDich", "") or "").strip()
        if loai not in loai_set:
            errors.append(
                f"Dòng {excel_row}: Loại giao dịch '{loai}' không hợp lệ — phải trùng ký tự với sheet 'LoaiGiaoDich_Mau'."
            )
            continue

        if not ma_nl:
            errors.append(f"Dòng {excel_row}: Mã nguyên liệu trống.")
            continue
        if nl_codes and ma_nl not in nl_codes:
            errors.append(f"Dòng {excel_row}: Mã nguyên liệu '{ma_nl}' không có trong danh mục.")
            continue

        try:
            sl = float(row.get("SoLuong", 0))
        except (TypeError, ValueError):
            sl = float("nan")
        if pd.isna(sl) or sl <= 0:
            errors.append(f"Dòng {excel_row}: Số lượng phải là số lớn hơn 0.")
            continue

        ma_k = str(row.get("MaKho", "") or "").strip()
        if not ma_k:
            errors.append(f"Dòng {excel_row}: Mã kho trống.")
            continue
        if kho_codes and ma_k not in kho_codes:
            errors.append(f"Dòng {excel_row}: Mã kho '{ma_k}' không có trong danh mục.")
            continue

        ma_ncc = str(row.get("MaNCC", "") or "").strip()
        if ma_ncc and ncc_codes and ma_ncc not in ncc_codes:
            errors.append(f"Dòng {excel_row}: Mã NCC '{ma_ncc}' không có trong danh mục.")
            continue

        ten_nl = str(row.get("TenNguyenLieu", "") or "").strip()
        don_vi = str(row.get("DonVi", "") or "").strip()
        rnl = nl_lookup.get(ma_nl)
        if rnl is not None:
            if not ten_nl and "TenNL" in rnl.index:
                ten_nl = str(rnl.get("TenNL", "") or "").strip()
            if not don_vi and "DonViMacDinh" in rnl.index:
                don_vi = str(rnl.get("DonViMacDinh", "") or "").strip() or "kg"
        if not don_vi:
            don_vi = "kg"

        so_lo = str(row.get("SoLo", "") or "").strip()
        hsd = _parse_hsd(row.get("HanSuDung"))
        ghi = str(row.get("GhiChu", "") or "").strip()

        pending.append(
            {
                "ID": ex.new_id(),
                "Ngay": ngay,
                "MaPhieu": ma_phieu,
                "LoaiGiaoDich": loai,
                "MaNguyenLieu": ma_nl,
                "TenNguyenLieu": ten_nl,
                "SoLuong": sl,
                "DonVi": don_vi,
                "MaKho": ma_k,
                "MaNCC": ma_ncc,
                "SoLo": so_lo,
                "HanSuDung": hsd,
                "GhiChu": ghi,
            }
        )

    if errors:
        return [], errors
    return pending, []
