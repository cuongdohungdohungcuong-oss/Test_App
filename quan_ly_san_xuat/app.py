"""
Ứng dụng quản lý sản xuất thức ăn chăn nuôi — Streamlit + Excel.
"""
from __future__ import annotations

import hashlib
from datetime import date, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st

import auth
import excel_io as ex
import production as pr
import sample_sx

st.set_page_config(page_title="Quản lý sản xuất", layout="wide", initial_sidebar_state="expanded")

MASTER_NEW = "— Thêm mới —"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

VAITRO_OPTIONS = ("QuanTri", "QuanLySanXuat", "ToTruongSX", "XemBaoCao")
VAITRO_LABELS = {
    "QuanTri": "Quản trị (cấp tài khoản)",
    "QuanLySanXuat": "Quản lý sản xuất (kế hoạch, công thức)",
    "ToTruongSX": "Tổ trưởng sản xuất (ghi nhận thực tế)",
    "XemBaoCao": "Chỉ xem báo cáo",
}
KICH_HOAT_OPTIONS = ("Co", "Khong")
KICH_HOAT_LABELS = {"Co": "Có (được đăng nhập)", "Khong": "Không (khóa tài khoản)"}

# Màu biểu đồ: slot 1 (xanh) = kế hoạch, slot 2 (cam) = thực tế
CHART_COLORS = ["#2a78d6", "#eb6834"]

SPEC_NL = {
    "p": "nl",
    "state_key": "_dm_nl",
    "sheet": ex.SHEET_NL,
    "pk": "MaNL",
    "cols": ex.NL_COLUMNS,
    "empty": {"MaNL": "", "TenNL": "", "DonViMacDinh": "kg", "GhiChu": ""},
    "labels": {
        "MaNL": "Mã nguyên liệu",
        "TenNL": "Tên nguyên liệu",
        "DonViMacDinh": "Đơn vị mặc định",
        "GhiChu": "Ghi chú",
    },
    "required": ("MaNL", "TenNL"),
}

SPEC_TP = {
    "p": "tp",
    "state_key": "_dm_tp",
    "sheet": ex.SHEET_TP,
    "pk": "MaTP",
    "cols": ex.TP_COLUMNS,
    "empty": {"MaTP": "", "TenTP": "", "DonVi": "kg", "QuyCach": "", "GhiChu": ""},
    "labels": {
        "MaTP": "Mã thành phẩm",
        "TenTP": "Tên thành phẩm",
        "DonVi": "Đơn vị",
        "QuyCach": "Quy cách đóng gói",
        "GhiChu": "Ghi chú",
    },
    "required": ("MaTP", "TenTP"),
}

SPEC_DC = {
    "p": "dc",
    "state_key": "_dm_dc",
    "sheet": ex.SHEET_DC,
    "pk": "MaDC",
    "cols": ex.DC_COLUMNS,
    "empty": {"MaDC": "", "TenDC": "", "CongSuat": "", "GhiChu": ""},
    "labels": {
        "MaDC": "Mã dây chuyền",
        "TenDC": "Tên dây chuyền",
        "CongSuat": "Công suất (tấn/giờ)",
        "GhiChu": "Ghi chú",
    },
    "required": ("MaDC", "TenDC"),
}

SPEC_KHO = {
    "p": "kho",
    "state_key": "_dm_kho",
    "sheet": ex.SHEET_KHO,
    "pk": "MaKho",
    "cols": ex.KHO_COLUMNS,
    "empty": {"MaKho": "", "TenKho": "", "DiaDiem": ""},
    "labels": {"MaKho": "Mã kho", "TenKho": "Tên kho", "DiaDiem": "Địa điểm"},
    "required": ("MaKho", "TenKho"),
}


# ---------- Dữ liệu & phân quyền ----------


def _norm_users_df(df: pd.DataFrame) -> pd.DataFrame:
    return pr.norm_df(df, ex.USER_COLUMNS).fillna("")


def reload_session_data() -> None:
    st.session_state["_lsx"] = pr.norm_df(ex.read_sheet(ex.SHEET_LSX), ex.LSX_COLUMNS)
    st.session_state["_thsx"] = pr.norm_df(ex.read_sheet(ex.SHEET_THSX), ex.THSX_COLUMNS)
    st.session_state["_gd"] = pr.norm_df(ex.read_sheet(ex.SHEET_GD), ex.GD_COLUMNS)
    st.session_state["_dm_ct"] = pr.norm_df(ex.read_sheet(ex.SHEET_CT), ex.CT_COLUMNS)
    st.session_state["_dm_nl"] = ex.read_sheet(ex.SHEET_NL)
    st.session_state["_dm_tp"] = ex.read_sheet(ex.SHEET_TP)
    st.session_state["_dm_dc"] = ex.read_sheet(ex.SHEET_DC)
    st.session_state["_dm_kho"] = ex.read_sheet(ex.SHEET_KHO)
    st.session_state["_dm_users"] = _norm_users_df(ex.read_sheet(ex.SHEET_USER))


def _role() -> str:
    return str((st.session_state.get("auth_user") or {}).get("role", "")).strip()


def _is_quantri() -> bool:
    return _role() == "QuanTri"


def can_plan() -> bool:
    """Lập kế hoạch, sửa công thức, danh mục."""
    return _role() in ("QuanTri", "QuanLySanXuat")


def can_record() -> bool:
    """Bắt đầu lệnh và ghi nhận thực tế."""
    return _role() in ("QuanTri", "QuanLySanXuat", "ToTruongSX")


def _kich_hoat_ok(val) -> bool:
    return str(val).strip().lower() in ("co", "1", "yes", "true", "có")


def _next_ma_nhan_vien(df: pd.DataFrame) -> str:
    codes = set(df["MaNV"].astype(str).str.strip()) if "MaNV" in df.columns else set()
    nums = [int(c[2:]) for c in codes if len(c) >= 3 and c[:2].upper() == "NV" and c[2:].isdigit()]
    n = (max(nums) + 1) if nums else 1
    while f"NV{n:04d}" in codes:
        n += 1
    return f"NV{n:04d}"


def authenticate(username: str, password: str) -> dict | None:
    df = _norm_users_df(ex.read_sheet(ex.SHEET_USER))
    hit = df[df["TenDangNhap"].astype(str).str.strip().str.lower() == str(username).strip().lower()]
    if hit.empty:
        return None
    r = hit.iloc[0]
    if not _kich_hoat_ok(r.get("KichHoat", "")):
        return None
    if not auth.verify_password(password, str(r.get("Salt", "")), str(r.get("MatKhauHash", ""))):
        return None
    return {
        "id": str(r.get("ID", "")).strip(),
        "username": str(r.get("TenDangNhap", "")).strip(),
        "hoten": str(r.get("HoTen", "")).strip(),
        "role": str(r.get("VaiTro", "")).strip(),
        "manv": str(r.get("MaNV", "")).strip(),
    }


def logout_user() -> None:
    st.session_state.clear()
    st.rerun()


def render_login() -> None:
    st.title("Đăng nhập — Quản lý sản xuất")
    st.caption("Chỉ tài khoản được cấp **tên đăng nhập** và **mật khẩu** mới sử dụng được phần mềm.")
    st.info(
        "Lần đầu chạy: tài khoản mặc định **admin** / **admin123** (sheet `DM_NguoiDung`). "
        "Nên đổi mật khẩu và tạo tài khoản riêng trong tab Danh mục — Người dùng."
    )
    with st.form("login_form"):
        u = st.text_input("Tên đăng nhập", autocomplete="username")
        p = st.text_input("Mật khẩu", type="password", autocomplete="current-password")
        submitted = st.form_submit_button("Đăng nhập", type="primary", width="stretch")
    if submitted:
        if not str(u).strip() or not str(p):
            st.error("Nhập đủ tên đăng nhập và mật khẩu.")
        else:
            usr = authenticate(u, p)
            if usr:
                st.session_state.auth_user = usr
                st.rerun()
            else:
                st.error("Sai tên đăng nhập hoặc mật khẩu, hoặc tài khoản đã bị khóa.")


# ---------- Tiện ích hiển thị ----------


def _cell(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def _fmt_dates(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    d = df.copy()
    for c in cols:
        if c in d.columns:
            d[c] = pd.to_datetime(d[c], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
    return d


def _code_options(df: pd.DataFrame, pk: str) -> list[str]:
    if df is None or df.empty or pk not in df.columns:
        return []
    s = df[pk].astype(str).str.strip()
    return sorted(set(s[s != ""]))


def _name_map(df: pd.DataFrame, pk: str, name_col: str) -> dict[str, str]:
    if df is None or df.empty or pk not in df.columns or name_col not in df.columns:
        return {}
    return {str(k).strip(): _cell(v) for k, v in zip(df[pk], df[name_col])}


def _labeler(names: dict[str, str]):
    return lambda k: f"{k} — {names[k]}" if names.get(k) else str(k)


def _idx(options: list[str], value: str) -> int:
    return options.index(value) if value in options else 0


def _to_date(v, default: date) -> date:
    ts = pd.to_datetime(v, errors="coerce")
    return default if pd.isna(ts) else ts.date()


# ---------- Danh mục (Master) ----------


def render_one_master(spec: dict, editable: bool) -> None:
    p, state_key, sheet, pk = spec["p"], spec["state_key"], spec["sheet"], spec["pk"]
    cols, empty, labels, required = spec["cols"], spec["empty"], spec["labels"], spec["required"]

    df = pr.norm_df(st.session_state[state_key], cols).fillna("")
    if not editable:
        st.dataframe(df, width="stretch", hide_index=True)
        return

    ver = int(st.session_state.setdefault(f"m_{p}_ver", 0))
    pick = st.selectbox("Chọn dòng để sửa hoặc xóa", [MASTER_NEW] + _code_options(df, pk), key=f"m_sb_{p}_{ver}")
    if pick == MASTER_NEW:
        cur = {c: _cell(empty.get(c, "")) for c in cols}
    else:
        r = df[df[pk].astype(str).str.strip() == pick].iloc[0]
        cur = {c: _cell(r.get(c)) for c in cols}

    def _bump(msg: str) -> None:
        reload_session_data()
        st.session_state[f"m_{p}_ver"] = ver + 1
        st.session_state["_flash"] = msg
        st.rerun()

    c1, c2, _ = st.columns([1, 1, 2])
    if c1.button("Tạo mới", key=f"m_new_{p}_{ver}", width="stretch"):
        st.session_state[f"m_{p}_ver"] = ver + 1
        st.rerun()
    if c2.button("Xóa", key=f"m_del_{p}_{ver}", width="stretch"):
        if pick == MASTER_NEW:
            st.warning("Hãy chọn một dòng cụ thể trong danh sách để xóa.")
        else:
            ex.save_master_sheet(sheet, df[df[pk].astype(str).str.strip() != pick].reset_index(drop=True))
            _bump("Đã xóa.")

    slug = hashlib.md5(pick.encode("utf-8")).hexdigest()[:12]
    with st.form(f"m_form_{p}_{ver}_{slug}"):
        vals = {
            c: st.text_input(labels.get(c, c), value=cur[c], disabled=bool(c == pk and pick != MASTER_NEW))
            for c in cols
        }
        submitted = st.form_submit_button("Lưu", type="primary", width="stretch")
    if submitted:
        row = {c: _cell(vals[c]) for c in cols}
        if pick != MASTER_NEW:
            row[pk] = pick
        missing = [labels.get(c, c) for c in required if not row.get(c)]
        exists = df[pk].astype(str).str.strip() == row[pk]
        if missing:
            st.error("Thiếu thông tin bắt buộc: " + ", ".join(missing) + ".")
        elif pick == MASTER_NEW and exists.any():
            st.error("Mã đã tồn tại. Chọn dòng đó trong danh sách để sửa hoặc nhập mã khác.")
        elif pick == MASTER_NEW:
            ex.save_master_sheet(sheet, pd.concat([df, pd.DataFrame([row])], ignore_index=True))
            _bump("Đã thêm mới.")
        else:
            new_df = df.copy()
            for c in cols:
                new_df.loc[exists, c] = row[c]
            ex.save_master_sheet(sheet, new_df)
            _bump("Đã cập nhật.")

    st.caption("Danh sách hiện tại trong sheet")
    st.dataframe(df, width="stretch", hide_index=True)


def render_user_admin() -> None:
    st.markdown("Cấp tài khoản cho người dùng. **Mật khẩu** được lưu dạng băm (không lưu chữ thường).")
    df = _norm_users_df(ex.read_sheet(ex.SHEET_USER))
    au = st.session_state.auth_user
    show = df.drop(columns=["MatKhauHash", "Salt"]).copy()
    show["VaiTro"] = show["VaiTro"].map(lambda x: VAITRO_LABELS.get(str(x).strip(), x))
    st.dataframe(show, width="stretch", hide_index=True)

    fmt_vt = lambda x: VAITRO_LABELS.get(x, x)  # noqa: E731
    fmt_kh = lambda x: KICH_HOAT_LABELS.get(x, x)  # noqa: E731

    st.subheader("Thêm người dùng")
    st.caption(f"Mã nhân viên gán tự động (bản xem trước: **{_next_ma_nhan_vien(df)}**).")
    with st.form("add_user", clear_on_submit=False):
        c1, c2 = st.columns(2)
        ho_ten = c1.text_input("Họ tên *")
        vaitro = c1.selectbox("Vai trò *", list(VAITRO_OPTIONS), index=1, format_func=fmt_vt)
        kich_hoat = c1.selectbox("Kích hoạt *", list(KICH_HOAT_OPTIONS), format_func=fmt_kh)
        ten_dn = c2.text_input("Tên đăng nhập *")
        mk1 = c2.text_input("Mật khẩu *", type="password")
        mk2 = c2.text_input("Nhập lại mật khẩu *", type="password")
        ghichu = st.text_input("Ghi chú")
        add_sub = st.form_submit_button("Thêm tài khoản", type="primary")
    if add_sub:
        ten_dn = str(ten_dn).strip()
        if not str(ho_ten).strip() or not ten_dn:
            st.error("Họ tên và tên đăng nhập là bắt buộc.")
        elif mk1 != mk2:
            st.error("Hai lần nhập mật khẩu không khớp.")
        elif len(str(mk1)) < 6:
            st.error("Mật khẩu tối thiểu 6 ký tự.")
        elif df["TenDangNhap"].astype(str).str.strip().str.lower().eq(ten_dn.lower()).any():
            st.error("Tên đăng nhập đã tồn tại.")
        else:
            salt, pw_hash = auth.hash_password(str(mk1))
            row = {
                "ID": ex.new_id(),
                "MaNV": _next_ma_nhan_vien(df),
                "HoTen": str(ho_ten).strip(),
                "VaiTro": vaitro,
                "TenDangNhap": ten_dn,
                "MatKhauHash": pw_hash,
                "Salt": salt,
                "KichHoat": kich_hoat,
                "GhiChu": str(ghichu).strip(),
            }
            ex.save_master_sheet(ex.SHEET_USER, pd.concat([df, pd.DataFrame([row])], ignore_index=True))
            reload_session_data()
            st.session_state["_flash"] = "Đã thêm tài khoản."
            st.rerun()

    st.subheader("Sửa / khóa / đặt lại mật khẩu")
    names = _code_options(df, "TenDangNhap")
    if not names:
        return
    pick = st.selectbox("Chọn tài khoản", names, key="adm_pick_user")
    r0 = df[df["TenDangNhap"].astype(str).str.strip() == pick].iloc[0]
    uid = str(r0["ID"]).strip()
    cur_role = str(r0["VaiTro"]).strip()
    cur_kh = "Co" if _kich_hoat_ok(r0["KichHoat"]) else "Khong"
    n_quantri = int((df["VaiTro"].astype(str).str.strip() == "QuanTri").sum())
    with st.form(f"edit_user_{uid}"):
        e_ht = st.text_input("Họ tên *", value=str(r0["HoTen"]))
        e_vt = st.selectbox("Vai trò *", list(VAITRO_OPTIONS), index=_idx(list(VAITRO_OPTIONS), cur_role), format_func=fmt_vt)
        e_kh = st.selectbox("Kích hoạt *", list(KICH_HOAT_OPTIONS), index=_idx(list(KICH_HOAT_OPTIONS), cur_kh), format_func=fmt_kh)
        e_gc = st.text_input("Ghi chú", value=str(r0["GhiChu"]))
        st.caption("Đặt lại mật khẩu: chỉ điền khi cần đổi.")
        e_pw1 = st.text_input("Mật khẩu mới", type="password")
        e_pw2 = st.text_input("Nhập lại mật khẩu mới", type="password")
        b1, b2 = st.columns(2)
        save_u = b1.form_submit_button("Lưu thay đổi", type="primary")
        del_u = b2.form_submit_button("Xóa tài khoản này")
    losing_admin = cur_role == "QuanTri" and n_quantri <= 1
    if save_u:
        if not str(e_ht).strip():
            st.error("Họ tên là bắt buộc.")
        elif e_vt != "QuanTri" and (uid == au.get("id") or losing_admin):
            st.error("Không thể bỏ vai trò quản trị của chính mình / của quản trị viên duy nhất.")
        elif e_pw1 and (e_pw1 != e_pw2 or len(e_pw1) < 6):
            st.error("Mật khẩu mới không khớp hoặc ít hơn 6 ký tự.")
        else:
            new_df = df.copy()
            m = new_df["ID"].astype(str).str.strip() == uid
            new_df.loc[m, ["HoTen", "VaiTro", "KichHoat", "GhiChu"]] = [str(e_ht).strip(), e_vt, e_kh, str(e_gc).strip()]
            if e_pw1:
                salt, pw_hash = auth.hash_password(str(e_pw1))
                new_df.loc[m, ["Salt", "MatKhauHash"]] = [salt, pw_hash]
            ex.save_master_sheet(ex.SHEET_USER, new_df)
            if uid == au.get("id"):
                st.session_state.auth_user = {**au, "hoten": str(e_ht).strip(), "role": e_vt}
            reload_session_data()
            st.session_state["_flash"] = "Đã cập nhật người dùng."
            st.rerun()
    if del_u:
        if uid == au.get("id") or losing_admin:
            st.error("Không thể xóa tài khoản đang đăng nhập hoặc quản trị viên duy nhất.")
        else:
            ex.save_master_sheet(ex.SHEET_USER, df[df["ID"].astype(str).str.strip() != uid].reset_index(drop=True))
            reload_session_data()
            st.session_state["_flash"] = "Đã xóa tài khoản."
            st.rerun()


# ---------- Lệnh sản xuất ----------


def _save_lsx(lsx: pd.DataFrame, msg: str) -> None:
    ex.save_master_sheet(ex.SHEET_LSX, lsx)
    reload_session_data()
    st.session_state["lsx_ver"] = int(st.session_state.get("lsx_ver", 0)) + 1
    st.session_state["_flash"] = msg
    st.rerun()


def render_lenh_sx() -> None:
    lsx: pd.DataFrame = st.session_state._lsx
    dm_ct, dm_nl = st.session_state._dm_ct, st.session_state._dm_nl
    tp_names = _name_map(st.session_state._dm_tp, "MaTP", "TenTP")
    dc_names = _name_map(st.session_state._dm_dc, "MaDC", "TenDC")
    kho_names = _name_map(st.session_state._dm_kho, "MaKho", "TenKho")

    st.subheader("Danh sách lệnh sản xuất")
    f1, f2, f3 = st.columns([1, 1, 2])
    d0 = f1.date_input("Từ ngày", value=date.today() - timedelta(days=30), key="lsx_d0")
    d1 = f2.date_input("Đến ngày", value=date.today() + timedelta(days=30), key="lsx_d1")
    tt_sel = f3.multiselect("Trạng thái", pr.TRANG_THAI_OPTIONS, default=pr.TRANG_THAI_OPTIONS, key="lsx_tt")
    view = pr.filter_lsx(lsx, d0, d1)
    view = view[view["TrangThai"].astype(str).isin(tt_sel)].sort_values(["NgayKH", "MaLenh"], ascending=False)
    first = ["MaLenh", "NgayKH", "TrangThai", "MaTP", "TenTP", "SanLuongKH", "SanLuongTT"]
    view = view[first + [c for c in ex.LSX_COLUMNS if c not in first and c != "ID"]]
    st.dataframe(
        _fmt_dates(view, ["NgayKH", "NgayHoanThanh"]).fillna(""),
        width="stretch",
        hide_index=True,
        height=320,
    )

    ver = int(st.session_state.get("lsx_ver", 0))
    if can_record():
        st.subheader("Chuyển trạng thái")
        active = lsx[lsx["TrangThai"].astype(str).isin([pr.TT_KE_HOACH, pr.TT_DANG_SX])]
        codes = sorted(active["MaLenh"].astype(str).tolist(), reverse=True)
        if not codes:
            st.caption("Không có lệnh ở trạng thái Kế hoạch / Đang sản xuất.")
        else:
            labels = dict(zip(active["MaLenh"].astype(str), active["TenTP"].astype(str) + " · " + active["TrangThai"].astype(str)))
            pick = st.selectbox("Chọn lệnh", codes, format_func=_labeler(labels), key=f"lsx_tt_pick_{ver}")
            cur = str(lsx.loc[lsx["MaLenh"].astype(str) == pick, "TrangThai"].iloc[0])
            m = lsx["MaLenh"].astype(str) == pick
            b1, b2, b3, _ = st.columns([1, 1, 1, 2])
            if b1.button("Bắt đầu sản xuất", disabled=cur != pr.TT_KE_HOACH, width="stretch", key=f"lsx_start_{ver}"):
                new = lsx.astype(object)
                new.loc[m, "TrangThai"] = pr.TT_DANG_SX
                _save_lsx(new, f"Lệnh {pick} chuyển sang «Đang sản xuất».")
            if b2.button("Hủy lệnh", disabled=not can_plan(), width="stretch", key=f"lsx_cancel_{ver}"):
                new = lsx.astype(object)
                new.loc[m, "TrangThai"] = pr.TT_HUY
                _save_lsx(new, f"Đã hủy lệnh {pick}.")
            if b3.button(
                "Xóa lệnh", disabled=not can_plan() or cur != pr.TT_KE_HOACH, width="stretch", key=f"lsx_del_{ver}"
            ):
                _save_lsx(lsx[~m].reset_index(drop=True), f"Đã xóa lệnh {pick}.")
            st.caption("Hoàn thành lệnh tại tab «Ghi nhận thực tế». Chỉ xóa được lệnh đang ở trạng thái Kế hoạch.")

    if not can_plan():
        return

    st.subheader("Tạo / sửa lệnh (trạng thái Kế hoạch)")
    ke_hoach = lsx[lsx["TrangThai"].astype(str) == pr.TT_KE_HOACH]
    pick = st.selectbox(
        "Chọn lệnh để sửa", [MASTER_NEW] + sorted(ke_hoach["MaLenh"].astype(str).tolist(), reverse=True), key=f"lsx_pick_{ver}"
    )
    cur = {c: "" for c in ex.LSX_COLUMNS}
    if pick != MASTER_NEW:
        cur = {c: ke_hoach[ke_hoach["MaLenh"].astype(str) == pick].iloc[0][c] for c in ex.LSX_COLUMNS}
    k = f"lsx_{ver}_{hashlib.md5(pick.encode('utf-8')).hexdigest()[:8]}"

    tp_opts, dc_opts, kho_opts = list(tp_names), list(dc_names), list(kho_names)
    if not tp_opts or not kho_opts:
        st.warning("Cần khai báo thành phẩm và kho trong tab Danh mục trước.")
        return
    c1, c2, c3 = st.columns(3)
    ngay = c1.date_input("Ngày kế hoạch *", value=_to_date(cur["NgayKH"], date.today()), key=f"{k}_ngay")
    ma_tp = c1.selectbox("Thành phẩm *", tp_opts, index=_idx(tp_opts, _cell(cur["MaTP"])), format_func=_labeler(tp_names), key=f"{k}_tp")
    sl_kh = c1.number_input(
        "Sản lượng kế hoạch (kg) *", min_value=0.0, step=500.0, value=pr.to_num(cur["SanLuongKH"]) or 10000.0,
        key=f"{k}_sl",
    )
    ma_dc = c2.selectbox(
        "Dây chuyền", dc_opts or [""], index=_idx(dc_opts, _cell(cur["MaDC"])), format_func=_labeler(dc_names), key=f"{k}_dc"
    )
    ca = c2.selectbox("Ca", pr.CA_OPTIONS, index=_idx(pr.CA_OPTIONS, _cell(cur["Ca"])), key=f"{k}_ca")
    so_lo = c2.text_input("Số lô (để trống = mã lệnh)", value=_cell(cur["SoLo"]), key=f"{k}_lo")
    kho_nl = c3.selectbox(
        "Kho xuất nguyên liệu *", kho_opts, index=_idx(kho_opts, _cell(cur["MaKhoNL"]) or "K01"), format_func=_labeler(kho_names), key=f"{k}_knl"
    )
    kho_tp = c3.selectbox(
        "Kho nhập thành phẩm *", kho_opts, index=_idx(kho_opts, _cell(cur["MaKhoTP"]) or "K02"), format_func=_labeler(kho_names), key=f"{k}_ktp"
    )
    ghi_chu = c3.text_input("Ghi chú", value=_cell(cur["GhiChu"]), key=f"{k}_gc")

    st.markdown("**Nhu cầu nguyên liệu theo định mức**")
    for w in pr.validate_bom(pr.bom_for(ma_tp, dm_ct)):
        st.warning(w)
    st.dataframe(
        pr.planned_consumption(ma_tp, sl_kh, dm_ct, dm_nl).rename(
            columns={"DinhMuc": "DinhMuc (kg/tấn)", "DinhMucKH": "CanXuat"}
        ),
        width="stretch",
        hide_index=True,
    )

    if st.button("Lưu lệnh sản xuất", type="primary", key=f"{k}_save"):
        row = {
            **cur,
            "NgayKH": pd.Timestamp(ngay),
            "MaTP": ma_tp,
            "TenTP": tp_names.get(ma_tp, ""),
            "SanLuongKH": float(sl_kh),
            "DonVi": "kg",
            "MaDC": ma_dc,
            "Ca": ca,
            "MaKhoNL": kho_nl,
            "MaKhoTP": kho_tp,
            "SoLo": so_lo.strip(),
            "TrangThai": pr.TT_KE_HOACH,
            "GhiChu": ghi_chu.strip(),
        }
        errs = pr.validate_lenh(row, dm_ct)
        if errs:
            for e in errs:
                st.error(e)
        elif pick == MASTER_NEW:
            row.update(ID=ex.new_id(), MaLenh=pr.next_ma_lenh(lsx, ngay), NguoiTao=st.session_state.auth_user.get("username", ""))
            _save_lsx(pd.concat([lsx, pd.DataFrame([row])], ignore_index=True), f"Đã tạo lệnh {row['MaLenh']}.")
        else:
            new = lsx.astype(object)
            m = new["MaLenh"].astype(str) == pick
            for c in ex.LSX_COLUMNS:
                new.loc[m, c] = row[c]
            _save_lsx(new, f"Đã cập nhật lệnh {pick}.")


# ---------- Ghi nhận thực tế ----------


def render_ghi_nhan() -> None:
    if not can_record():
        st.info("Tài khoản của bạn chỉ có quyền xem.")
        return
    lsx: pd.DataFrame = st.session_state._lsx
    dang = lsx[lsx["TrangThai"].astype(str) == pr.TT_DANG_SX]
    if dang.empty:
        st.info("Không có lệnh nào đang sản xuất. Chuyển lệnh sang «Đang sản xuất» ở tab Lệnh sản xuất.")
        return
    labels = dict(zip(dang["MaLenh"].astype(str), dang["TenTP"].astype(str)))
    ver = int(st.session_state.get("lsx_ver", 0))
    pick = st.selectbox("Lệnh đang sản xuất", sorted(labels, reverse=True), format_func=_labeler(labels), key=f"gn_pick_{ver}")
    lenh = dang[dang["MaLenh"].astype(str) == pick].iloc[0].to_dict()
    sl_kh = pr.to_num(lenh["SanLuongKH"])

    i1, i2, i3, i4 = st.columns(4)
    i1.metric("Sản lượng kế hoạch (kg)", f"{sl_kh:,.0f}")
    i2.metric("Dây chuyền / Ca", f"{_cell(lenh['MaDC'])} · {_cell(lenh['Ca'])}")
    i3.metric("Kho NL → Kho TP", f"{_cell(lenh['MaKhoNL'])} → {_cell(lenh['MaKhoTP'])}")
    i4.metric("Ngày kế hoạch", _to_date(lenh["NgayKH"], date.today()).strftime("%d/%m/%Y"))

    c1, c2 = st.columns(2)
    sl_tt = c1.number_input("Sản lượng thực tế (kg) *", min_value=0.0, step=100.0, value=sl_kh, key=f"gn_sl_{pick}")
    ngay_ht = c2.date_input("Ngày hoàn thành", value=date.today(), key=f"gn_ngay_{pick}")

    st.markdown("**Tiêu hao nguyên liệu thực tế** — `DinhMucKH` tính theo sản lượng thực tế; sửa cột `SoLuongTT`.")
    base = pr.planned_consumption(lenh["MaTP"], sl_tt, st.session_state._dm_ct, st.session_state._dm_nl)
    base["SoLuongTT"] = base["DinhMucKH"]
    th = st.data_editor(
        base[["MaNL", "TenNL", "DinhMucKH", "SoLuongTT", "DonVi"]],
        disabled=["MaNL", "TenNL", "DinhMucKH", "DonVi"],
        column_config={"SoLuongTT": st.column_config.NumberColumn("SoLuongTT", min_value=0.0, step=1.0, format="%.2f")},
        width="stretch",
        hide_index=True,
        key=f"gn_ed_{pick}_{sl_tt}",
    )
    var = pr.report_variance(th.assign(MaLenh=pick, ID=""))
    if not var.empty:
        tong_dm, tong_tt = var["DinhMucKH"].sum(), var["SoLuongTT"].sum()
        hh = (tong_tt - tong_dm) / tong_dm * 100 if tong_dm else 0.0
        st.caption(f"Tổng định mức {tong_dm:,.2f} kg · thực tế {tong_tt:,.2f} kg · chênh lệch {hh:+.2f}%")

    if st.button("Hoàn thành lệnh & ghi phiếu kho", type="primary", key=f"gn_done_{pick}"):
        try:
            lsx2, thsx2, gd2 = pr.complete_order(
                lenh, float(sl_tt), th, lsx, st.session_state._thsx, st.session_state._gd, ngay_ht
            )
        except ValueError as e:
            st.error(str(e))
        else:
            ex.save_sheets({ex.SHEET_LSX: lsx2, ex.SHEET_THSX: thsx2, ex.SHEET_GD: gd2})
            reload_session_data()
            st.session_state["lsx_ver"] = ver + 1
            st.session_state["_flash"] = f"Đã hoàn thành lệnh {pick}; sinh phiếu {pick}-X (xuất NL) và {pick}-N (nhập TP)."
            st.rerun()


# ---------- Công thức (BOM) ----------


def render_cong_thuc() -> None:
    dm_ct: pd.DataFrame = st.session_state._dm_ct
    dm_nl = st.session_state._dm_nl
    tp_names = _name_map(st.session_state._dm_tp, "MaTP", "TenTP")
    nl_names = _name_map(dm_nl, "MaNL", "TenNL")
    if not tp_names:
        st.info("Chưa có thành phẩm. Thêm ở tab Danh mục.")
        return
    st.caption("Định mức = kg nguyên liệu cho **1.000 kg** thành phẩm; tổng các dòng phải bằng 1.000.")
    ma_tp = st.selectbox("Thành phẩm", list(tp_names), format_func=_labeler(tp_names), key="ct_tp")
    bom = pr.bom_for(ma_tp, dm_ct)[["MaNL", "DinhMuc", "GhiChu"]].fillna("")
    ver = int(st.session_state.get("ct_ver", 0))
    edited = st.data_editor(
        bom,
        num_rows="dynamic" if can_plan() else "fixed",
        disabled=not can_plan(),
        column_config={
            "MaNL": st.column_config.SelectboxColumn("Nguyên liệu", options=list(nl_names), required=True),
            "DinhMuc": st.column_config.NumberColumn("Định mức (kg/tấn)", min_value=0.0, step=1.0, format="%.3f"),
            "GhiChu": st.column_config.TextColumn("Ghi chú"),
        },
        width="stretch",
        hide_index=True,
        key=f"ct_ed_{ma_tp}_{ver}",
    )
    edited = edited.dropna(how="all")
    total = float(pd.to_numeric(edited["DinhMuc"], errors="coerce").fillna(0).sum())
    st.metric("Tổng định mức (kg / 1.000 kg)", f"{total:,.3f}")
    errs = pr.validate_bom(edited)
    for e in errs:
        st.warning(e)
    if can_plan() and st.button("Lưu công thức", type="primary", disabled=bool(errs), key=f"ct_save_{ver}"):
        new_rows = edited.assign(
            ID=[ex.new_id() for _ in range(len(edited))],
            MaTP=ma_tp,
            TenNL=edited["MaNL"].astype(str).map(nl_names).fillna(""),
        )[ex.CT_COLUMNS]
        rest = dm_ct[dm_ct["MaTP"].astype(str).str.strip() != ma_tp]
        ex.save_master_sheet(ex.SHEET_CT, pd.concat([rest, new_rows], ignore_index=True))
        reload_session_data()
        st.session_state["ct_ver"] = ver + 1
        st.session_state["_flash"] = f"Đã lưu công thức {ma_tp}."
        st.rerun()


# ---------- Báo cáo ----------


def render_bao_cao() -> None:
    lsx: pd.DataFrame = st.session_state._lsx
    f1, f2, f3 = st.columns([1, 1, 2])
    d0 = f1.date_input("Từ ngày", value=date.today() - timedelta(days=30), key="bc_d0")
    d1 = f2.date_input("Đến ngày", value=date.today(), key="bc_d1")
    group_label = f3.radio("Nhóm theo", list(pr.GROUP_OPTIONS), horizontal=True, key="bc_group")
    view = pr.filter_lsx(lsx, d0, d1)
    if view.empty:
        st.info("Không có lệnh sản xuất trong khoảng thời gian này.")
        return

    con = view[view["TrangThai"].astype(str) != pr.TT_HUY]
    done = con[con["TrangThai"].astype(str) == pr.TT_HOAN_THANH]
    kh_done = pd.to_numeric(done["SanLuongKH"], errors="coerce").sum()
    tt_done = pd.to_numeric(done["SanLuongTT"], errors="coerce").sum()
    var = pr.report_variance(st.session_state._thsx, done["MaLenh"].astype(str).tolist())
    dm_sum = var["DinhMucKH"].sum() if not var.empty else 0.0
    hh = (var["SoLuongTT"].sum() - dm_sum) / dm_sum * 100 if dm_sum else 0.0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Lệnh hoàn thành / tổng", f"{len(done)} / {len(con)}")
    m2.metric("Sản lượng thực tế (tấn)", f"{tt_done / 1000:,.1f}")
    m3.metric("Đạt kế hoạch (lệnh đã xong)", f"{(tt_done / kh_done * 100) if kh_done else 0:,.1f}%")
    m4.metric("Hao hụt NL so với định mức", f"{hh:+.2f}%")

    group_col = pr.GROUP_OPTIONS[group_label]
    out = pr.report_output(view, group_col)
    st.markdown(f"**Sản lượng kế hoạch và thực tế theo {group_label.lower()} (kg)**")
    chart = out.set_index(group_col)[["SanLuongKH", "SanLuongTT"]].rename(
        columns={"SanLuongKH": "Kế hoạch", "SanLuongTT": "Thực tế"}
    )
    if group_col == "NgayKH":
        chart.index = pd.to_datetime(chart.index).strftime("%Y-%m-%d")
    st.bar_chart(chart, color=CHART_COLORS, stack=False, height=320)
    st.dataframe(out, width="stretch", hide_index=True)
    st.caption("Lệnh Hủy không tính. TyLeHoanThanh = SanLuongTT / SanLuongKH (%), gồm cả lệnh chưa hoàn thành.")

    st.markdown("**Tiêu hao nguyên liệu: định mức và thực tế (lệnh đã hoàn thành)**")
    st.dataframe(var, width="stretch", hide_index=True)
    st.caption("ChenhLech > 0: dùng vượt định mức (hao hụt).")

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        out.to_excel(w, sheet_name="SanLuong", index=False)
        var.to_excel(w, sheet_name="TieuHao", index=False)
        view.to_excel(w, sheet_name="LenhSanXuat", index=False)
    st.download_button(
        "Tải báo cáo Excel",
        data=buf.getvalue(),
        file_name=f"bao_cao_san_xuat_{d0:%Y%m%d}_{d1:%Y%m%d}.xlsx",
        mime=XLSX_MIME,
    )


def render_phieu_kho() -> None:
    gd: pd.DataFrame = st.session_state._gd
    st.caption(
        "Phiếu kho sinh tự động khi hoàn thành lệnh: `<MaLenh>-X` xuất nguyên liệu, `<MaLenh>-N` nhập thành phẩm. "
        "Cùng cột với sheet GiaoDich của app kho để có thể chép sang."
    )
    q = st.text_input("Lọc theo từ khóa", key="pk_q", placeholder="Ví dụ: LSX-2026, NL001, Nhập kho…")
    view = gd.copy()
    if q.strip():
        s = q.strip().lower()
        view = view[view.astype(str).apply(lambda c: c.str.lower().str.contains(s, regex=False)).any(axis=1)]
    st.metric("Số dòng", len(view))
    st.dataframe(_fmt_dates(view, ["Ngay", "HanSuDung"]), width="stretch", hide_index=True, height=480)
    if not view.empty:
        st.download_button(
            "Tải CSV (theo bộ lọc)",
            data=view.to_csv(index=False).encode("utf-8-sig"),
            file_name="phieu_kho_san_xuat.csv",
            mime="text/csv",
        )


# ---------- Main ----------


def main() -> None:
    ex.ensure_app_schema()
    if not st.session_state.get("auth_user"):
        render_login()
        return
    if "_lsx" not in st.session_state:
        reload_session_data()

    st.title("Quản lý sản xuất thức ăn chăn nuôi")
    st.caption(f"Dữ liệu lưu tại: `{ex.DATA_FILE}`")
    flash = st.session_state.pop("_flash", None)
    if flash:
        st.success(flash)

    lsx: pd.DataFrame = st.session_state._lsx
    with st.sidebar:
        au = st.session_state.auth_user
        st.subheader("Tài khoản")
        st.write(f"**{au.get('hoten', '')}**")
        st.caption(f"{VAITRO_LABELS.get(au.get('role', ''), au.get('role', ''))} · `{au.get('username', '')}`")
        if st.button("Đăng xuất", width="stretch"):
            logout_user()
        st.divider()
        st.subheader("Tổng quan")
        tt = lsx["TrangThai"].astype(str)
        st.metric("Lệnh đang sản xuất", int((tt == pr.TT_DANG_SX).sum()))
        st.caption(
            f"Kế hoạch: {int((tt == pr.TT_KE_HOACH).sum())} · Hoàn thành: {int((tt == pr.TT_HOAN_THANH).sum())} · "
            f"Hủy: {int((tt == pr.TT_HUY).sum())}"
        )
        if st.button("Tải lại dữ liệu từ Excel", width="stretch"):
            reload_session_data()
            st.rerun()
        if can_plan():
            with st.expander("Dữ liệu mẫu"):
                st.caption("Thêm 30 lệnh SX mẫu trong 30 ngày gần nhất (kèm tiêu hao, phiếu kho).")
                if st.button("Thêm 30 lệnh mẫu", width="stretch"):
                    n = sample_sx.append_sample_orders(30)
                    reload_session_data()
                    st.session_state["_flash"] = f"Đã thêm {n} lệnh sản xuất mẫu."
                    st.rerun()
        st.divider()
        if ex.DATA_FILE.exists():
            st.download_button(
                "Tải xuống san_xuat.xlsx",
                data=ex.DATA_FILE.read_bytes(),
                file_name="san_xuat.xlsx",
                mime=XLSX_MIME,
                width="stretch",
            )

    tabs = st.tabs(["Lệnh sản xuất", "Ghi nhận thực tế", "Công thức (BOM)", "Báo cáo sản xuất", "Phiếu kho SX", "Danh mục"])
    with tabs[0]:
        render_lenh_sx()
    with tabs[1]:
        render_ghi_nhan()
    with tabs[2]:
        render_cong_thuc()
    with tabs[3]:
        render_bao_cao()
    with tabs[4]:
        render_phieu_kho()
    with tabs[5]:
        labels = ["Thành phẩm", "Nguyên liệu", "Dây chuyền", "Kho"] + (["Người dùng (quản trị)"] if _is_quantri() else [])
        sub = st.tabs(labels)
        for t, spec in zip(sub, (SPEC_TP, SPEC_NL, SPEC_DC, SPEC_KHO)):
            with t:
                render_one_master(spec, can_plan())
        if _is_quantri():
            with sub[4]:
                render_user_admin()


if __name__ == "__main__":
    main()
