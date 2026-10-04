"""Nghiệp vụ sản xuất: công thức (BOM), lệnh sản xuất, tiêu hao, báo cáo — không phụ thuộc streamlit."""
from __future__ import annotations

from datetime import date
from typing import Any

import pandas as pd

import excel_io as ex

TT_KE_HOACH = "Kế hoạch"
TT_DANG_SX = "Đang sản xuất"
TT_HOAN_THANH = "Hoàn thành"
TT_HUY = "Hủy"
TRANG_THAI_OPTIONS = [TT_KE_HOACH, TT_DANG_SX, TT_HOAN_THANH, TT_HUY]

CA_OPTIONS = ["Ca 1", "Ca 2", "Ca 3"]

# Định mức trong BOM tính cho 1.000 kg thành phẩm
BOM_BASE = 1000.0
BOM_TOLERANCE = 0.5

_TRANSITIONS = {
    TT_KE_HOACH: {TT_DANG_SX, TT_HUY},
    TT_DANG_SX: {TT_HOAN_THANH, TT_HUY},
    TT_HOAN_THANH: set(),
    TT_HUY: set(),
}


def _s(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def to_num(v: Any) -> float:
    x = pd.to_numeric(v, errors="coerce")
    return 0.0 if pd.isna(x) else float(x)


def norm_df(df: pd.DataFrame | None, cols: list[str]) -> pd.DataFrame:
    d = pd.DataFrame(columns=cols) if df is None else df.copy()
    for c in cols:
        if c not in d.columns:
            d[c] = ""
    return d[cols]


def can_transition(old: str, new: str) -> bool:
    return new in _TRANSITIONS.get(_s(old), set())


def next_ma_lenh(lsx: pd.DataFrame, ngay: date) -> str:
    """Sinh mã LSX-YYYYMMDD-NNN tiếp theo trong ngày, không trùng mã đã có."""
    prefix = f"LSX-{ngay:%Y%m%d}-"
    codes = set(lsx["MaLenh"].astype(str).str.strip()) if "MaLenh" in lsx.columns else set()
    nums = [int(c[len(prefix):]) for c in codes if c.startswith(prefix) and c[len(prefix):].isdigit()]
    n = (max(nums) + 1) if nums else 1
    while f"{prefix}{n:03d}" in codes:
        n += 1
    return f"{prefix}{n:03d}"


def bom_for(ma_tp: str, dm_ct: pd.DataFrame) -> pd.DataFrame:
    ct = norm_df(dm_ct, ex.CT_COLUMNS)
    out = ct[ct["MaTP"].astype(str).str.strip() == _s(ma_tp)].copy()
    out["DinhMuc"] = pd.to_numeric(out["DinhMuc"], errors="coerce").fillna(0.0)
    return out.reset_index(drop=True)


def validate_bom(bom: pd.DataFrame) -> list[str]:
    """Trả về danh sách lỗi/cảnh báo của một công thức (đã lọc theo 1 thành phẩm)."""
    errs: list[str] = []
    if bom.empty:
        return ["Công thức chưa có nguyên liệu."]
    ma = bom["MaNL"].astype(str).str.strip()
    if (ma == "").any():
        errs.append("Có dòng chưa chọn mã nguyên liệu.")
    dup = sorted(set(ma[ma.duplicated() & (ma != "")]))
    if dup:
        errs.append("Nguyên liệu bị lặp: " + ", ".join(dup) + ".")
    dm = pd.to_numeric(bom["DinhMuc"], errors="coerce")
    if dm.isna().any() or (dm <= 0).any():
        errs.append("Định mức phải là số > 0.")
    total = float(dm.fillna(0).sum())
    if abs(total - BOM_BASE) > BOM_TOLERANCE:
        errs.append(f"Tổng định mức = {total:,.2f} kg, cần bằng {BOM_BASE:,.0f} kg cho 1 tấn thành phẩm.")
    return errs


def planned_consumption(ma_tp: str, san_luong: float, dm_ct: pd.DataFrame, dm_nl: pd.DataFrame) -> pd.DataFrame:
    """Nhu cầu nguyên liệu theo định mức cho sản lượng (kg) thành phẩm."""
    bom = bom_for(ma_tp, dm_ct)
    cols = ["MaNL", "TenNL", "DinhMuc", "DinhMucKH", "DonVi"]
    if bom.empty:
        return pd.DataFrame(columns=cols)
    nl = norm_df(dm_nl, ["MaNL", "TenNL", "DonViMacDinh"])
    nl["MaNL"] = nl["MaNL"].astype(str).str.strip()
    bom["MaNL"] = bom["MaNL"].astype(str).str.strip()
    m = bom.merge(nl, on="MaNL", how="left", suffixes=("", "_dm"))
    m["TenNL"] = m["TenNL_dm"].where(m["TenNL_dm"].astype(str).str.strip().ne("") & m["TenNL_dm"].notna(), m["TenNL"])
    m["DonVi"] = m["DonViMacDinh"].fillna("").astype(str).replace("", "kg")
    m["DinhMucKH"] = (m["DinhMuc"] * float(san_luong) / BOM_BASE).round(3)
    return m[cols].reset_index(drop=True)


def validate_lenh(row: dict[str, Any], dm_ct: pd.DataFrame) -> list[str]:
    errs: list[str] = []
    if not _s(row.get("MaTP")):
        errs.append("Chưa chọn thành phẩm.")
    elif bom_for(row["MaTP"], dm_ct).empty:
        errs.append("Thành phẩm chưa có công thức (BOM).")
    if to_num(row.get("SanLuongKH")) <= 0:
        errs.append("Sản lượng kế hoạch phải > 0.")
    if not _s(row.get("MaKhoNL")) or not _s(row.get("MaKhoTP")):
        errs.append("Chọn kho xuất nguyên liệu và kho nhập thành phẩm.")
    return errs


def complete_order(
    lenh: dict[str, Any],
    san_luong_tt: float,
    tieu_hao: pd.DataFrame,
    lsx: pd.DataFrame,
    thsx: pd.DataFrame,
    gd: pd.DataFrame,
    ngay_ht: date,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Hoàn thành lệnh: cập nhật LenhSanXuat, ghi TieuHaoSX, sinh phiếu Xuất kho NL / Nhập kho TP.

    tieu_hao: cột MaNL, TenNL, DinhMucKH, SoLuongTT, DonVi.
    Trả về (lsx, thsx, gd) mới. Raise ValueError nếu dữ liệu không hợp lệ.
    """
    ma_lenh = _s(lenh.get("MaLenh"))
    if not can_transition(lenh.get("TrangThai", ""), TT_HOAN_THANH):
        raise ValueError(f"Lệnh {ma_lenh} đang ở trạng thái «{_s(lenh.get('TrangThai'))}», không thể hoàn thành.")
    if san_luong_tt <= 0:
        raise ValueError("Sản lượng thực tế phải > 0.")
    th = norm_df(tieu_hao, ["MaNL", "TenNL", "DinhMucKH", "SoLuongTT", "DonVi"])
    th["SoLuongTT"] = pd.to_numeric(th["SoLuongTT"], errors="coerce")
    if th["SoLuongTT"].isna().any() or (th["SoLuongTT"] < 0).any():
        raise ValueError("Số lượng tiêu hao thực tế phải là số ≥ 0.")
    phieu_x, phieu_n = f"{ma_lenh}-X", f"{ma_lenh}-N"
    gd = norm_df(gd, ex.GD_COLUMNS)
    if gd["MaPhieu"].astype(str).str.strip().isin([phieu_x, phieu_n]).any():
        raise ValueError(f"Phiếu kho của lệnh {ma_lenh} đã tồn tại trong GiaoDich.")

    ngay = pd.Timestamp(ngay_ht)
    so_lo = _s(lenh.get("SoLo")) or ma_lenh
    new_gd: list[dict[str, Any]] = []
    for _, r in th[th["SoLuongTT"] > 0].iterrows():
        new_gd.append(
            ex.row_from_form(
                {
                    "Ngay": ngay,
                    "MaPhieu": phieu_x,
                    "LoaiGiaoDich": "Xuất kho",
                    "MaNguyenLieu": _s(r["MaNL"]),
                    "TenNguyenLieu": _s(r["TenNL"]),
                    "SoLuong": float(r["SoLuongTT"]),
                    "DonVi": _s(r["DonVi"]) or "kg",
                    "MaKho": _s(lenh.get("MaKhoNL")),
                    "MaNCC": "",
                    "SoLo": so_lo,
                    "HanSuDung": pd.NaT,
                    "GhiChu": f"Xuất NL cho lệnh SX {ma_lenh}",
                },
                None,
            )
        )
    new_gd.append(
        ex.row_from_form(
            {
                "Ngay": ngay,
                "MaPhieu": phieu_n,
                "LoaiGiaoDich": "Nhập kho",
                "MaNguyenLieu": _s(lenh.get("MaTP")),
                "TenNguyenLieu": _s(lenh.get("TenTP")),
                "SoLuong": float(san_luong_tt),
                "DonVi": _s(lenh.get("DonVi")) or "kg",
                "MaKho": _s(lenh.get("MaKhoTP")),
                "MaNCC": "",
                "SoLo": so_lo,
                "HanSuDung": pd.NaT,
                "GhiChu": f"Nhập thành phẩm từ lệnh SX {ma_lenh}",
            },
            None,
        )
    )
    gd_out = pd.concat([gd, pd.DataFrame(new_gd, columns=ex.GD_COLUMNS)], ignore_index=True)

    th_rows = th.assign(ID=[ex.new_id() for _ in range(len(th))], MaLenh=ma_lenh)
    thsx = norm_df(thsx, ex.THSX_COLUMNS)
    thsx = thsx[thsx["MaLenh"].astype(str).str.strip() != ma_lenh]
    thsx_out = pd.concat([thsx, th_rows[ex.THSX_COLUMNS]], ignore_index=True)

    lsx_out = norm_df(lsx, ex.LSX_COLUMNS).astype(object)
    m = lsx_out["MaLenh"].astype(str).str.strip() == ma_lenh
    if not m.any():
        raise ValueError(f"Không tìm thấy lệnh {ma_lenh}.")
    lsx_out.loc[m, "TrangThai"] = TT_HOAN_THANH
    lsx_out.loc[m, "SanLuongTT"] = float(san_luong_tt)
    lsx_out.loc[m, "NgayHoanThanh"] = ngay
    return lsx_out, thsx_out, gd_out


# ---------- Báo cáo ----------

GROUP_OPTIONS = {
    "Ngày": "NgayKH",
    "Thành phẩm": "TenTP",
    "Dây chuyền": "MaDC",
    "Ca": "Ca",
}


def filter_lsx(lsx: pd.DataFrame, d0: date, d1: date) -> pd.DataFrame:
    d = norm_df(lsx, ex.LSX_COLUMNS)
    nd = pd.to_datetime(d["NgayKH"], errors="coerce").dt.date
    return d[nd.notna() & (nd >= d0) & (nd <= d1)].copy()


def report_output(lsx: pd.DataFrame, group_col: str) -> pd.DataFrame:
    """Sản lượng kế hoạch vs thực tế (bỏ lệnh Hủy), % hoàn thành theo nhóm."""
    d = norm_df(lsx, ex.LSX_COLUMNS)
    d = d[d["TrangThai"].astype(str) != TT_HUY].copy()
    cols = [group_col, "SoLenh", "SoLenhHoanThanh", "SanLuongKH", "SanLuongTT", "TyLeHoanThanh"]
    if d.empty:
        return pd.DataFrame(columns=cols)
    if group_col == "NgayKH":
        d["NgayKH"] = pd.to_datetime(d["NgayKH"], errors="coerce").dt.date
    d["SanLuongKH"] = pd.to_numeric(d["SanLuongKH"], errors="coerce").fillna(0.0)
    d["SanLuongTT"] = pd.to_numeric(d["SanLuongTT"], errors="coerce").fillna(0.0)
    d["_ht"] = (d["TrangThai"].astype(str) == TT_HOAN_THANH).astype(int)
    g = (
        d.groupby(group_col, dropna=False)
        .agg(SoLenh=("MaLenh", "count"), SoLenhHoanThanh=("_ht", "sum"), SanLuongKH=("SanLuongKH", "sum"), SanLuongTT=("SanLuongTT", "sum"))
        .reset_index()
    )
    g["TyLeHoanThanh"] = (g["SanLuongTT"] / g["SanLuongKH"].where(g["SanLuongKH"] > 0) * 100).round(2)
    return g[cols].sort_values(group_col).reset_index(drop=True)


def report_variance(thsx: pd.DataFrame, ma_lenh: list[str] | None = None) -> pd.DataFrame:
    """Tiêu hao định mức vs thực tế theo nguyên liệu; ChenhLech > 0 là hao hụt vượt định mức."""
    d = norm_df(thsx, ex.THSX_COLUMNS)
    if ma_lenh is not None:
        d = d[d["MaLenh"].astype(str).isin(ma_lenh)]
    cols = ["MaNL", "TenNL", "DinhMucKH", "SoLuongTT", "ChenhLech", "TyLeHaoHut"]
    if d.empty:
        return pd.DataFrame(columns=cols)
    d["DinhMucKH"] = pd.to_numeric(d["DinhMucKH"], errors="coerce").fillna(0.0)
    d["SoLuongTT"] = pd.to_numeric(d["SoLuongTT"], errors="coerce").fillna(0.0)
    g = d.groupby(["MaNL", "TenNL"], dropna=False)[["DinhMucKH", "SoLuongTT"]].sum().reset_index()
    g["ChenhLech"] = (g["SoLuongTT"] - g["DinhMucKH"]).round(3)
    g["TyLeHaoHut"] = (g["ChenhLech"] / g["DinhMucKH"].where(g["DinhMucKH"] > 0) * 100).round(2)
    return g[cols].sort_values("MaNL").reset_index(drop=True)


# ---------- Tồn kho ----------

LOAI_NHAP = "Nhập kho"
LOAI_XUAT = "Xuất kho"
LOAI_DIEU_CHINH = "Điều chỉnh"
LOAI_PHIEU_TAY = [LOAI_NHAP, LOAI_DIEU_CHINH]


def signed_qty(gd: pd.DataFrame) -> pd.Series:
    """Nhập kho: +SoLuong, Xuất kho: -SoLuong, Điều chỉnh: SoLuong giữ dấu (có thể âm)."""
    q = pd.to_numeric(gd["SoLuong"], errors="coerce").fillna(0.0)
    loai = gd["LoaiGiaoDich"].astype(str).str.strip()
    sign = loai.map({LOAI_NHAP: 1.0, LOAI_XUAT: -1.0, LOAI_DIEU_CHINH: 1.0}).fillna(0.0)
    return q * sign


def stock_on_hand(gd: pd.DataFrame) -> pd.DataFrame:
    """Tồn kho hiện tại theo kho và mã hàng (nguyên liệu hoặc thành phẩm)."""
    d = norm_df(gd, ex.GD_COLUMNS)
    cols = ["MaKho", "MaNguyenLieu", "TenNguyenLieu", "TonKho", "DonVi"]
    if d.empty:
        return pd.DataFrame(columns=cols)
    d["_q"] = signed_qty(d)
    d["MaKho"] = d["MaKho"].astype(str).str.strip()
    d["MaNguyenLieu"] = d["MaNguyenLieu"].astype(str).str.strip()
    g = (
        d.groupby(["MaKho", "MaNguyenLieu"])
        .agg(TenNguyenLieu=("TenNguyenLieu", "last"), TonKho=("_q", "sum"), DonVi=("DonVi", "last"))
        .reset_index()
    )
    g["TonKho"] = g["TonKho"].round(3)
    return g[cols].sort_values(["MaKho", "MaNguyenLieu"]).reset_index(drop=True)


def check_material(
    ma_tp: str, san_luong: float, ma_kho: str, dm_ct: pd.DataFrame, dm_nl: pd.DataFrame, gd: pd.DataFrame
) -> pd.DataFrame:
    """Nhu cầu NL theo định mức so với tồn tại kho xuất; ThieuHut > 0 là thiếu."""
    need = planned_consumption(ma_tp, san_luong, dm_ct, dm_nl)
    stock = stock_on_hand(gd)
    stock = stock[stock["MaKho"] == _s(ma_kho)].set_index("MaNguyenLieu")["TonKho"]
    need["TonKho"] = need["MaNL"].astype(str).map(stock).fillna(0.0).astype(float)
    need["ThieuHut"] = (need["DinhMucKH"] - need["TonKho"]).clip(lower=0).round(3)
    return need[["MaNL", "TenNL", "DinhMucKH", "TonKho", "ThieuHut", "DonVi"]]


def next_ma_phieu(gd: pd.DataFrame, ngay: date, prefix: str = "PN") -> str:
    p = f"{prefix}-{ngay:%Y%m%d}-"
    codes = set(gd["MaPhieu"].astype(str).str.strip()) if "MaPhieu" in gd.columns else set()
    nums = [int(c[len(p):]) for c in codes if c.startswith(p) and c[len(p):].isdigit()]
    n = (max(nums) + 1) if nums else 1
    while f"{p}{n:03d}" in codes:
        n += 1
    return f"{p}{n:03d}"


# ---------- In phiếu lệnh ----------


def _html_table(df: pd.DataFrame, headers: dict[str, str]) -> str:
    from html import escape

    head = "".join(f"<th>{escape(h)}</th>" for h in headers.values())
    rows = []
    for _, r in df.iterrows():
        cells = []
        for c in headers:
            v = r.get(c, "")
            if isinstance(v, (int, float)) and not pd.isna(v):
                cells.append(f'<td class="n">{v:,.2f}</td>')
            else:
                cells.append(f"<td>{escape(_s(v))}</td>")
        rows.append("<tr>" + "".join(cells) + "</tr>")
    return f"<table><thead><tr>{head}</tr></thead><tbody>{''.join(rows)}</tbody></table>"


def order_sheet_html(
    lenh: dict[str, Any], consumption: pd.DataFrame, names: dict[str, dict[str, str]] | None = None
) -> str:
    """Phiếu lệnh sản xuất dạng HTML để in (mở trong trình duyệt → Ctrl+P).

    consumption: MaNL, TenNL, DinhMucKH và (nếu đã hoàn thành) SoLuongTT.
    names: {"kho": {...}, "dc": {...}} để hiển thị tên kho / dây chuyền.
    """
    from html import escape

    names = names or {}
    kho, dc = names.get("kho", {}), names.get("dc", {})

    def fmt_date(v: Any) -> str:
        ts = pd.to_datetime(v, errors="coerce")
        return "" if pd.isna(ts) else ts.strftime("%d/%m/%Y")

    def with_name(code: Any, m: dict[str, str]) -> str:
        c = _s(code)
        return f"{c} — {m[c]}" if m.get(c) else c

    info = [
        ("Mã lệnh", _s(lenh.get("MaLenh"))),
        ("Trạng thái", _s(lenh.get("TrangThai"))),
        ("Ngày kế hoạch", fmt_date(lenh.get("NgayKH"))),
        ("Ngày hoàn thành", fmt_date(lenh.get("NgayHoanThanh"))),
        ("Thành phẩm", f"{_s(lenh.get('MaTP'))} — {_s(lenh.get('TenTP'))}"),
        ("Số lô", _s(lenh.get("SoLo")) or _s(lenh.get("MaLenh"))),
        ("Sản lượng kế hoạch", f"{to_num(lenh.get('SanLuongKH')):,.2f} {_s(lenh.get('DonVi')) or 'kg'}"),
        ("Sản lượng thực tế", f"{to_num(lenh.get('SanLuongTT')):,.2f}" if to_num(lenh.get("SanLuongTT")) else ""),
        ("Dây chuyền / Ca", f"{with_name(lenh.get('MaDC'), dc)} · {_s(lenh.get('Ca'))}"),
        ("Kho xuất NL", with_name(lenh.get("MaKhoNL"), kho)),
        ("Kho nhập TP", with_name(lenh.get("MaKhoTP"), kho)),
        ("Ghi chú", _s(lenh.get("GhiChu"))),
    ]
    info_html = "".join(f"<tr><th>{escape(k)}</th><td>{escape(v)}</td></tr>" for k, v in info)
    headers = {"MaNL": "Mã NL", "TenNL": "Tên nguyên liệu", "DinhMucKH": "Định mức (kg)"}
    if "SoLuongTT" in consumption.columns:
        headers["SoLuongTT"] = "Thực tế (kg)"
    else:
        headers["_tt"] = "Thực tế (kg)"
    table = _html_table(consumption.assign(_tt=""), headers)
    title = f"Phiếu lệnh sản xuất {_s(lenh.get('MaLenh'))}"
    return f"""<!doctype html>
<html lang="vi"><head><meta charset="utf-8"><title>{escape(title)}</title>
<style>
body{{font-family:Arial,Helvetica,sans-serif;color:#111;margin:24px;font-size:13px}}
h1{{font-size:20px;margin:0 0 4px}} .sub{{color:#555;margin-bottom:16px}}
table{{border-collapse:collapse;width:100%;margin-bottom:18px}}
th,td{{border:1px solid #999;padding:6px 8px;text-align:left;vertical-align:top}}
.info th{{width:28%;background:#f2f2f2}} thead th{{background:#f2f2f2}} td.n{{text-align:right}}
.sign{{display:flex;justify-content:space-between;margin-top:36px;text-align:center}}
.sign div{{width:30%}} .sign small{{display:block;color:#555;margin-top:56px}}
@media print{{body{{margin:0}}}}
</style></head><body>
<h1>{escape(title)}</h1>
<div class="sub">In ngày {date.today():%d/%m/%Y}</div>
<table class="info">{info_html}</table>
<h2 style="font-size:15px">Nguyên liệu</h2>
{table}
<div class="sign"><div><b>Người lập</b><small>(Ký, ghi rõ họ tên)</small></div>
<div><b>Tổ trưởng sản xuất</b><small>(Ký, ghi rõ họ tên)</small></div>
<div><b>Thủ kho</b><small>(Ký, ghi rõ họ tên)</small></div></div>
</body></html>"""
