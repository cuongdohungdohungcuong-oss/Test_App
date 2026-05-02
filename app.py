"""
Ứng dụng nhập liệu quản lý kho thức ăn chăn nuôi — Streamlit + Excel.
"""
from __future__ import annotations

import hashlib
from datetime import date, datetime, timedelta
from io import BytesIO

import pandas as pd
import streamlit as st

import auth
import excel_io as ex
import gd_import
import sample_gd

st.set_page_config(page_title="Quản lý kho thức ăn", layout="wide", initial_sidebar_state="expanded")

LOAI_GD_OPTIONS = [
    "Nhập kho",
    "Xuất kho",
    "Điều chuyển",
    "Kiểm kê",
    "Hủy / loại bỏ",
]

MASTER_NEW = "— Thêm mới —"

VAITRO_OPTIONS = ("QuanTri", "NhanVienKho", "KeToanKho")
VAITRO_LABELS = {
    "QuanTri": "Quản trị (cấp tài khoản)",
    "NhanVienKho": "Nhân viên kho",
    "KeToanKho": "Kế toán kho",
}
KICH_HOAT_OPTIONS = ("Co", "Khong")
KICH_HOAT_LABELS = {"Co": "Có (được đăng nhập)", "Khong": "Không (khóa tài khoản)"}

SPEC_NL = {
    "p": "nl",
    "state_key": "_dm_nl",
    "sheet": ex.SHEET_NL,
    "pk": "MaNL",
    "cols": ["MaNL", "TenNL", "DonViMacDinh", "GhiChu"],
    "empty": {"MaNL": "", "TenNL": "", "DonViMacDinh": "kg", "GhiChu": ""},
    "labels": {
        "MaNL": "Mã nguyên liệu",
        "TenNL": "Tên nguyên liệu",
        "DonViMacDinh": "Đơn vị mặc định",
        "GhiChu": "Ghi chú",
    },
    "required": ("MaNL", "TenNL"),
}

SPEC_KHO = {
    "p": "kho",
    "state_key": "_dm_kho",
    "sheet": ex.SHEET_KHO,
    "pk": "MaKho",
    "cols": ["MaKho", "TenKho", "DiaDiem"],
    "empty": {"MaKho": "", "TenKho": "", "DiaDiem": ""},
    "labels": {"MaKho": "Mã kho", "TenKho": "Tên kho", "DiaDiem": "Địa điểm"},
    "required": ("MaKho", "TenKho"),
}

SPEC_NCC = {
    "p": "ncc",
    "state_key": "_dm_ncc",
    "sheet": ex.SHEET_NCC,
    "pk": "MaNCC",
    "cols": ["MaNCC", "TenNCC", "DienThoai"],
    "empty": {"MaNCC": "", "TenNCC": "", "DienThoai": ""},
    "labels": {"MaNCC": "Mã NCC", "TenNCC": "Tên nhà cung cấp", "DienThoai": "Điện thoại"},
    "required": ("MaNCC", "TenNCC"),
}

SPEC_DV = {
    "p": "dv",
    "state_key": "_dm_dv",
    "sheet": ex.SHEET_DV,
    "pk": "MaDV",
    "cols": ["MaDV", "TenDV"],
    "empty": {"MaDV": "", "TenDV": ""},
    "labels": {"MaDV": "Mã đơn vị", "TenDV": "Tên đơn vị"},
    "required": ("MaDV", "TenDV"),
}


def load_gd() -> pd.DataFrame:
    df = ex.read_sheet(ex.SHEET_GD)
    for c in ex.GD_COLUMNS:
        if c not in df.columns:
            df[c] = ""
    return df[ex.GD_COLUMNS]


def _norm_users_df(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    for c in ex.USER_COLUMNS:
        if c not in d.columns:
            d[c] = ""
    return d[ex.USER_COLUMNS].fillna("")


def reload_session_data() -> None:
    st.session_state["_gd"] = load_gd()
    st.session_state["_dm_nl"] = ex.read_sheet(ex.SHEET_NL)
    st.session_state["_dm_kho"] = ex.read_sheet(ex.SHEET_KHO)
    st.session_state["_dm_ncc"] = ex.read_sheet(ex.SHEET_NCC)
    st.session_state["_dm_dv"] = ex.read_sheet(ex.SHEET_DV)
    st.session_state["_dm_users"] = _norm_users_df(ex.read_sheet(ex.SHEET_USER))


def _is_quantri(user: dict | None) -> bool:
    return bool(user) and str(user.get("role", "")).strip() == "QuanTri"


def _next_ma_nhan_vien(df: pd.DataFrame) -> str:
    """Sinh mã NVdddd tiếp theo, không trùng cột MaNV hiện có."""
    if df.empty or "MaNV" not in df.columns:
        return "NV0001"
    codes = set(df["MaNV"].astype(str).str.strip())
    nums: list[int] = []
    for v in codes:
        u = str(v).strip()
        if len(u) >= 3 and u[:2].upper() == "NV" and u[2:].isdigit():
            nums.append(int(u[2:], 10))
    n = (max(nums) + 1) if nums else 1
    candidate = f"NV{n:04d}"
    while candidate in codes:
        n += 1
        candidate = f"NV{n:04d}"
    return candidate


def _kich_hoat_ok(val) -> bool:
    s = str(val).strip().lower()
    return s in ("co", "1", "yes", "true", "có")


def authenticate(username: str, password: str) -> dict | None:
    df = _norm_users_df(ex.read_sheet(ex.SHEET_USER))
    if df.empty:
        return None
    u = str(username).strip().lower()
    df = df.copy()
    df["_u"] = df["TenDangNhap"].astype(str).str.strip().str.lower()
    hit = df[df["_u"] == u]
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
    st.title("Đăng nhập hệ thống")
    st.caption(
        "Chỉ nhân viên kho, kế toán kho hoặc quản trị được cấp **tên đăng nhập** và **mật khẩu** mới sử dụng được phần mềm."
    )
    st.info(
        "Lần đầu chạy: tài khoản mặc định **admin** / **admin123** (sheet `DM_NguoiDung`). "
        "Nên đổi mật khẩu và tạo tài khoản riêng trong tab Danh mục — Người dùng (quyền quản trị)."
    )
    with st.form("login_form"):
        u = st.text_input("Tên đăng nhập", autocomplete="username")
        p = st.text_input("Mật khẩu", type="password", autocomplete="current-password")
        submitted = st.form_submit_button("Đăng nhập", type="primary", use_container_width=True)
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


def render_user_admin() -> None:
    st.markdown("Cấp tài khoản cho nhân viên kho / kế toán kho. **Mật khẩu** được lưu dạng băm (không lưu chữ thường).")
    df = _norm_users_df(ex.read_sheet(ex.SHEET_USER))
    au = st.session_state.auth_user

    show = df.drop(columns=["MatKhauHash", "Salt"], errors="ignore").copy()
    show["VaiTro"] = show["VaiTro"].map(lambda x: VAITRO_LABELS.get(str(x).strip(), x))
    st.dataframe(show, use_container_width=True, hide_index=True)

    st.subheader("Thêm người dùng")
    ma_preview = _next_ma_nhan_vien(df)
    st.caption(f"Mã nhân viên sẽ được hệ thống gán tự động (bản xem trước: **{ma_preview}**).")
    with st.form("add_user"):
        c1, c2 = st.columns(2)
        with c1:
            ho_ten = st.text_input("Họ tên *", key="add_hoten")
            vaitro = st.selectbox(
                "Vai trò *",
                list(VAITRO_OPTIONS),
                format_func=lambda x: VAITRO_LABELS.get(x, x),
                key="add_vaitro",
            )
            kich_hoat = st.selectbox(
                "Kích hoạt *",
                list(KICH_HOAT_OPTIONS),
                format_func=lambda x: KICH_HOAT_LABELS.get(x, x),
                index=0,
                key="add_kh",
            )
        with c2:
            ten_dn = st.text_input("Tên đăng nhập *", key="add_user")
            mk1 = st.text_input("Mật khẩu *", type="password", key="add_pw1")
            mk2 = st.text_input("Nhập lại mật khẩu *", type="password", key="add_pw2")
        ghichu = st.text_input("Ghi chú (không bắt buộc)", key="add_ghichu")
        add_sub = st.form_submit_button("Thêm tài khoản", type="primary")
    if add_sub:
        ten_dn = str(ten_dn).strip()
        ho_clean = str(ho_ten).strip()
        if not ho_clean:
            st.error("Họ tên là bắt buộc.")
        elif not ten_dn:
            st.error("Tên đăng nhập là bắt buộc.")
        elif not str(mk1).strip() or not str(mk2).strip():
            st.error("Mật khẩu và nhập lại mật khẩu là bắt buộc.")
        elif mk1 != mk2:
            st.error("Hai lần nhập mật khẩu không khớp.")
        elif len(str(mk1)) < 6:
            st.error("Mật khẩu tối thiểu 6 ký tự.")
        elif df["TenDangNhap"].astype(str).str.strip().str.lower().eq(ten_dn.lower()).any():
            st.error("Tên đăng nhập đã tồn tại.")
        else:
            salt, pw_hash = auth.hash_password(str(mk1))
            ma_nv = _next_ma_nhan_vien(df)
            row = {
                "ID": ex.new_id(),
                "MaNV": ma_nv,
                "HoTen": ho_clean,
                "VaiTro": vaitro,
                "TenDangNhap": ten_dn,
                "MatKhauHash": pw_hash,
                "Salt": salt,
                "KichHoat": kich_hoat,
                "GhiChu": str(ghichu).strip(),
            }
            new_df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
            ex.save_master_sheet(ex.SHEET_USER, new_df)
            reload_session_data()
            st.success("Đã thêm tài khoản.")
            st.rerun()

    st.subheader("Sửa / khóa / đặt lại mật khẩu")
    names = df["TenDangNhap"].astype(str).str.strip()
    names = names[names != ""].tolist()
    if not names:
        st.warning("Chưa có người dùng nào trong sheet.")
        return
    pick = st.selectbox("Chọn tài khoản", sorted(set(names)), key="adm_pick_user")
    hit = df[df["TenDangNhap"].astype(str).str.strip() == pick]
    if hit.empty:
        return
    r0 = hit.iloc[0]
    uid = str(r0.get("ID", "")).strip()

    with st.form("edit_user"):
        st.text_input(
            "Mã nhân viên (hệ thống)",
            value=str(r0.get("MaNV", "")),
            disabled=True,
            key="ed_manv_ro",
        )
        e_ma = str(r0.get("MaNV", "")).strip()
        e_ht = st.text_input("Họ tên *", value=str(r0.get("HoTen", "")), key="ed_hoten")
        cur_role = str(r0.get("VaiTro", "NhanVienKho")).strip()
        if cur_role not in VAITRO_OPTIONS:
            cur_role = "NhanVienKho"
        e_vt = st.selectbox(
            "Vai trò *",
            list(VAITRO_OPTIONS),
            index=list(VAITRO_OPTIONS).index(cur_role),
            format_func=lambda x: VAITRO_LABELS.get(x, x),
            key="ed_vaitro",
        )
        kh = str(r0.get("KichHoat", "Co")).strip()
        if kh not in KICH_HOAT_OPTIONS:
            kh = "Co" if _kich_hoat_ok(kh) else "Khong"
        e_kh = st.selectbox(
            "Kích hoạt *",
            list(KICH_HOAT_OPTIONS),
            index=list(KICH_HOAT_OPTIONS).index(kh),
            format_func=lambda x: KICH_HOAT_LABELS.get(x, x),
            key="ed_kh",
        )
        e_gc = st.text_input("Ghi chú (không bắt buộc)", value=str(r0.get("GhiChu", "")), key="ed_gc")
        st.caption("Đặt lại mật khẩu: chỉ điền khi cần đổi (để trống nếu giữ nguyên).")
        e_pw1 = st.text_input("Mật khẩu mới", type="password", key="ed_pw1")
        e_pw2 = st.text_input("Nhập lại mật khẩu mới", type="password", key="ed_pw2")
        b1, b2 = st.columns(2)
        save_u = b1.form_submit_button("Lưu thay đổi", type="primary")
        del_u = b2.form_submit_button("Xóa tài khoản này")

        if save_u:
            n_quantri = (df["VaiTro"].astype(str).str.strip() == "QuanTri").sum()
            if not str(e_ht).strip():
                st.error("Họ tên là bắt buộc.")
            elif uid == au.get("id") and e_vt != "QuanTri":
                st.error("Bạn không thể tự bỏ vai trò quản trị của chính mình.")
            elif str(r0.get("VaiTro", "")).strip() == "QuanTri" and e_vt != "QuanTri" and n_quantri <= 1:
                st.error("Phải còn ít nhất một tài khoản Quản trị.")
            elif str(e_pw1).strip() and str(e_pw1) != str(e_pw2):
                st.error("Hai lần nhập mật khẩu mới không khớp.")
            elif str(e_pw1).strip() and len(str(e_pw1).strip()) < 6:
                st.error("Mật khẩu mới tối thiểu 6 ký tự.")
            else:
                new_df = df.copy()
                m = new_df["ID"].astype(str).str.strip() == uid
                new_df.loc[m, "MaNV"] = e_ma
                new_df.loc[m, "HoTen"] = str(e_ht).strip()
                new_df.loc[m, "VaiTro"] = e_vt
                new_df.loc[m, "KichHoat"] = e_kh
                new_df.loc[m, "GhiChu"] = str(e_gc).strip()
                if str(e_pw1).strip():
                    salt, pw_hash = auth.hash_password(str(e_pw1))
                    new_df.loc[m, "Salt"] = salt
                    new_df.loc[m, "MatKhauHash"] = pw_hash
                ex.save_master_sheet(ex.SHEET_USER, new_df)
                reload_session_data()
                if uid == au.get("id"):
                    rr = new_df.loc[m].iloc[0]
                    st.session_state.auth_user = {
                        "id": uid,
                        "username": str(rr.get("TenDangNhap", "")).strip(),
                        "hoten": str(rr.get("HoTen", "")).strip(),
                        "role": str(rr.get("VaiTro", "")).strip(),
                        "manv": str(rr.get("MaNV", "")).strip(),
                    }
                st.success("Đã cập nhật người dùng.")
                st.rerun()

        if del_u:
            n_quantri = (df["VaiTro"].astype(str).str.strip() == "QuanTri").sum()
            if uid == au.get("id"):
                st.error("Không thể xóa chính tài khoản đang đăng nhập.")
            elif str(r0.get("VaiTro", "")).strip() == "QuanTri" and n_quantri <= 1:
                st.error("Không thể xóa tài khoản Quản trị duy nhất.")
            else:
                new_df = df[df["ID"].astype(str).str.strip() != uid].reset_index(drop=True)
                ex.save_master_sheet(ex.SHEET_USER, new_df)
                reload_session_data()
                st.success("Đã xóa tài khoản.")
                st.rerun()


def init_state() -> None:
    if "_gd" not in st.session_state:
        reload_session_data()
    if "edit_id" not in st.session_state:
        st.session_state.edit_id = None
    if "search_q" not in st.session_state:
        st.session_state.search_q = ""
    if "last_search" not in st.session_state:
        st.session_state.last_search = ""
    if "form_ma_phieu" not in st.session_state:
        clear_form_to_new()


def clear_form_to_new() -> None:
    st.session_state.edit_id = None
    st.session_state.form_ma_phieu = ""
    st.session_state.form_loai = LOAI_GD_OPTIONS[0]
    st.session_state.form_ma_nl = ""
    st.session_state.form_sl = 0.0
    st.session_state.form_dv = "kg"
    st.session_state.form_ma_kho = ""
    st.session_state.form_ma_ncc = ""
    st.session_state.form_so_lo = ""
    st.session_state.form_ghi_chu = ""
    st.session_state.form_ngay = date.today()
    st.session_state.form_hsd = date.today()
    st.session_state.form_has_hsd = False


def filtered_gd(df: pd.DataFrame, q: str) -> pd.DataFrame:
    if not q or not str(q).strip():
        return df.copy()
    s = str(q).strip().lower()
    mask = pd.Series(False, index=df.index)
    for col in df.columns:
        mask = mask | df[col].astype(str).str.lower().str.contains(s, na=False)
    return df.loc[mask].copy()


def render_gd_viewer(gd: pd.DataFrame) -> None:
    st.subheader("Xem dữ liệu giao dịch trên web")
    st.caption(
        "Lọc theo ngày, loại phiếu, kho, nguyên liệu và từ khóa; bảng hiển thị trực tiếp dữ liệu đang dùng (đồng bộ với Excel sau khi bấm «Tải lại» hoặc sau khi lưu)."
    )

    work = gd.copy()
    if work.empty:
        st.info("Chưa có giao dịch. Có thể thêm mẫu trong mục bên dưới hoặc nhập tại tab «Giao dịch kho».")
        with st.expander("Thêm ~200 giao dịch mẫu vào Excel"):
            if st.button("Thêm 200 giao dịch mẫu", key="seed_empty"):
                n = sample_gd.append_sample_transactions(200)
                reload_session_data()
                st.success(f"Đã thêm {n} dòng vào sheet GiaoDich.")
                st.rerun()
        return

    _ts = pd.to_datetime(work["Ngay"], errors="coerce")
    work["_d"] = _ts.dt.date
    valid_ts = _ts.dropna()
    if valid_ts.empty:
        d_min = date.today() - timedelta(days=365)
        d_max = date.today()
    else:
        d_min = valid_ts.min().date()
        d_max = valid_ts.max().date()

    r1, r2, r3, r4 = st.columns([1, 1, 1, 2])
    with r1:
        d0 = st.date_input("Từ ngày", value=d_min, key="vw_d0")
    with r2:
        d1 = st.date_input("Đến ngày", value=d_max, key="vw_d1")
    with r3:
        loai_opts = sorted(work["LoaiGiaoDich"].dropna().astype(str).unique().tolist())
        if not loai_opts:
            loai_opts = list(LOAI_GD_OPTIONS)
        loai_sel = st.multiselect(
            "Loại giao dịch",
            options=loai_opts if loai_opts else list(LOAI_GD_OPTIONS),
            default=loai_opts if loai_opts else list(LOAI_GD_OPTIONS),
            key="vw_loai",
        )
    with r4:
        qv = st.text_input("Từ khóa (lọc thêm trên các cột)", key="vw_q", placeholder="Ví dụ: MAU-, NL001, Nhập kho…")

    r5, r6 = st.columns(2)
    with r5:
        nl_opts = sorted(work["MaNguyenLieu"].dropna().astype(str).str.strip().unique().tolist())
        nl_sel = st.multiselect("Mã nguyên liệu", options=nl_opts, default=nl_opts, key="vw_nl")
    with r6:
        k_opts = sorted(work["MaKho"].dropna().astype(str).str.strip().unique().tolist())
        k_sel = st.multiselect("Mã kho", options=k_opts, default=k_opts, key="vw_kho")

    m = pd.Series(True, index=work.index)
    m &= work["_d"].notna()
    m &= (work["_d"] >= d0) & (work["_d"] <= d1)
    if loai_sel:
        m &= work["LoaiGiaoDich"].astype(str).isin(loai_sel)
    if nl_sel:
        m &= work["MaNguyenLieu"].astype(str).str.strip().isin(nl_sel)
    if k_sel:
        m &= work["MaKho"].astype(str).str.strip().isin(k_sel)
    filt = work.loc[m].copy()
    if qv and str(qv).strip():
        filt = filtered_gd(filt.drop(columns=["_d"], errors="ignore"), str(qv).strip())
    else:
        filt = filt.drop(columns=["_d"], errors="ignore")

    st.metric("Số dòng sau lọc", len(filt))
    disp = filt.copy()
    if not disp.empty:
        if "Ngay" in disp.columns:
            disp["Ngay"] = pd.to_datetime(disp["Ngay"], errors="coerce").dt.strftime("%d/%m/%Y")
        if "HanSuDung" in disp.columns:
            disp["HanSuDung"] = pd.to_datetime(disp["HanSuDung"], errors="coerce").dt.strftime("%d/%m/%Y")
        if "SoLuong" in disp.columns:
            disp["SoLuong"] = pd.to_numeric(disp["SoLuong"], errors="coerce").round(3)
    st.dataframe(disp, use_container_width=True, hide_index=True, height=520)

    c_dl, c_seed = st.columns([1, 2])
    with c_dl:
        if not disp.empty:
            csv = filt.drop(columns=["_d"], errors="ignore").to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "Tải CSV (đúng bộ lọc hiện tại)",
                data=csv,
                file_name="giao_dich_loc.csv",
                mime="text/csv",
                key="vw_dl",
            )
    with c_seed:
        with st.expander("Thêm ~200 giao dịch mẫu vào Excel (minh họa)"):
            st.caption("Thêm vào **cuối** sheet GiaoDich, không xóa dữ liệu cũ. Mã phiếu dạng MAU-… để dễ nhận biết.")
            if st.button("Thêm 200 giao dịch mẫu", key="vw_seed"):
                n = sample_gd.append_sample_transactions(200)
                reload_session_data()
                st.success(f"Đã thêm {n} dòng.")
                st.rerun()


def form_defaults_from_row(row: pd.Series) -> None:
    st.session_state.edit_id = str(row.get("ID", "") or "")
    st.session_state.form_ma_phieu = str(row.get("MaPhieu", "") or "")
    st.session_state.form_loai = str(row.get("LoaiGiaoDich", "") or LOAI_GD_OPTIONS[0])
    if st.session_state.form_loai not in LOAI_GD_OPTIONS:
        st.session_state.form_loai = LOAI_GD_OPTIONS[0]
    st.session_state.form_ma_nl = str(row.get("MaNguyenLieu", "") or "")
    st.session_state.form_sl = float(row.get("SoLuong", 0) or 0)
    st.session_state.form_dv = str(row.get("DonVi", "") or "kg")
    st.session_state.form_ma_kho = str(row.get("MaKho", "") or "")
    st.session_state.form_ma_ncc = str(row.get("MaNCC", "") or "")
    st.session_state.form_so_lo = str(row.get("SoLo", "") or "")
    st.session_state.form_ghi_chu = str(row.get("GhiChu", "") or "")
    ngay = row.get("Ngay")
    if pd.notna(ngay):
        if isinstance(ngay, datetime):
            st.session_state.form_ngay = ngay.date()
        else:
            try:
                st.session_state.form_ngay = pd.to_datetime(ngay).date()
            except Exception:
                st.session_state.form_ngay = date.today()
    else:
        st.session_state.form_ngay = date.today()
    hsd = row.get("HanSuDung")
    if pd.notna(hsd):
        try:
            st.session_state.form_hsd = pd.to_datetime(hsd).date()
            st.session_state.form_has_hsd = True
        except Exception:
            st.session_state.form_hsd = date.today()
            st.session_state.form_has_hsd = False
    else:
        st.session_state.form_hsd = date.today()
        st.session_state.form_has_hsd = False


def _dm_cell_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v).strip()


def _norm_dm(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    d = df.copy()
    for c in cols:
        if c not in d.columns:
            d[c] = ""
    return d[cols].fillna("")


def _master_pick_options(df: pd.DataFrame, pk: str) -> list[str]:
    s = df[pk].astype(str).str.strip()
    s = s[s != ""]
    keys = sorted(set(s.tolist()))
    return [MASTER_NEW] + keys


def render_one_master(spec: dict) -> None:
    p = spec["p"]
    state_key = spec["state_key"]
    sheet = spec["sheet"]
    pk = spec["pk"]
    cols = spec["cols"]
    empty = spec["empty"]
    labels = spec["labels"]
    required = spec["required"]

    df = _norm_dm(st.session_state[state_key], cols)
    ver = int(st.session_state.setdefault(f"m_{p}_ver", 0))
    prev_k = f"m_{p}_prev_pick"

    options = _master_pick_options(df, pk)
    pick = st.selectbox(
        "Chọn dòng để sửa hoặc xóa",
        options,
        key=f"m_sb_{p}_{ver}",
    )

    if st.session_state.get(prev_k) != pick:
        if pick == MASTER_NEW:
            st.session_state[f"m_buf_{p}"] = {c: _dm_cell_str(empty.get(c, "")) for c in cols}
        else:
            hit = df[df[pk].astype(str).str.strip() == str(pick).strip()]
            if hit.empty:
                st.session_state[f"m_buf_{p}"] = {c: _dm_cell_str(empty.get(c, "")) for c in cols}
            else:
                r = hit.iloc[0]
                st.session_state[f"m_buf_{p}"] = {c: _dm_cell_str(r.get(c)) for c in cols}
        st.session_state[prev_k] = pick

    buf = st.session_state.setdefault(f"m_buf_{p}", {c: _dm_cell_str(empty.get(c, "")) for c in cols})
    for c in cols:
        buf.setdefault(c, _dm_cell_str(empty.get(c, "")))

    c_btn1, c_btn2, c_btn3 = st.columns([1, 1, 2])
    with c_btn1:
        if st.button("Tạo mới", key=f"m_new_{p}_{ver}", use_container_width=True):
            st.session_state[f"m_{p}_ver"] = ver + 1
            st.session_state.pop(prev_k, None)
            st.rerun()
    with c_btn2:
        if st.button("Xóa", key=f"m_del_{p}_{ver}", use_container_width=True):
            if pick == MASTER_NEW:
                st.warning("Hãy chọn một dòng cụ thể trong danh sách để xóa.")
            else:
                pk_val = str(pick).strip()
                new_df = df[df[pk].astype(str).str.strip() != pk_val].reset_index(drop=True)
                ex.save_master_sheet(sheet, new_df)
                reload_session_data()
                st.session_state[f"m_{p}_ver"] = ver + 1
                st.session_state.pop(prev_k, None)
                st.success("Đã xóa.")
                st.rerun()

    slug = hashlib.md5(str(pick).encode("utf-8")).hexdigest()[:12]
    with st.form(f"m_form_{p}_{ver}_{slug}"):
        row_vals: dict[str, str] = {}
        for c in cols:
            lab = labels.get(c, c)
            dis = bool(c == pk and pick != MASTER_NEW)
            row_vals[c] = st.text_input(lab, value=str(buf.get(c, "")), disabled=dis)
        submitted = st.form_submit_button("Lưu", type="primary", use_container_width=True)

    if submitted:
        row = {c: _dm_cell_str(row_vals.get(c, "")) for c in cols}
        missing = [labels.get(c, c) for c in required if not row.get(c, "")]
        if missing:
            st.error("Thiếu thông tin bắt buộc: " + ", ".join(missing) + ".")
        else:
            new_df = df.copy()
            exist_mask = new_df[pk].astype(str).str.strip() == row[pk]
            if pick == MASTER_NEW:
                if exist_mask.any():
                    st.error("Mã đã tồn tại. Chọn dòng đó trong danh sách để sửa hoặc nhập mã khác.")
                else:
                    new_df = pd.concat([new_df, pd.DataFrame([row])], ignore_index=True)
                    ex.save_master_sheet(sheet, new_df)
                    reload_session_data()
                    st.session_state[f"m_{p}_ver"] = ver + 1
                    st.session_state.pop(prev_k, None)
                    st.success("Đã thêm mới.")
                    st.rerun()
            else:
                m = new_df[pk].astype(str).str.strip() == str(pick).strip()
                if not m.any():
                    st.error("Không tìm thấy dòng đang sửa. Bấm \"Tải lại dữ liệu từ Excel\" ở thanh bên.")
                else:
                    for c in cols:
                        new_df.loc[m, c] = row[c]
                    ex.save_master_sheet(sheet, new_df)
                    reload_session_data()
                    st.session_state[f"m_{p}_ver"] = ver + 1
                    st.session_state.pop(prev_k, None)
                    st.success("Đã cập nhật.")
                    st.rerun()

    st.caption("Danh sách hiện tại trong sheet")
    st.dataframe(df, use_container_width=True, hide_index=True)


def main() -> None:
    ex.ensure_app_schema()
    if not st.session_state.get("auth_user"):
        render_login()
        return

    init_state()

    st.title("Quản lý kho thức ăn chăn nuôi")
    st.caption(f"Dữ liệu lưu tại: `{ex.DATA_FILE}` — có sheet giao dịch và các sheet Master.")

    gd: pd.DataFrame = st.session_state._gd
    dm_nl: pd.DataFrame = st.session_state._dm_nl
    dm_kho: pd.DataFrame = st.session_state._dm_kho
    dm_ncc: pd.DataFrame = st.session_state._dm_ncc

    with st.sidebar:
        au = st.session_state.auth_user
        st.subheader("Tài khoản")
        st.write(f"**{au.get('hoten', '')}**")
        st.caption(f"{VAITRO_LABELS.get(au.get('role', ''), au.get('role', ''))} · `{au.get('username', '')}`")
        if st.button("Đăng xuất", use_container_width=True):
            logout_user()
        st.divider()
        st.subheader("Tổng quan")
        st.metric("Số dòng giao dịch", len(gd))
        nhap = (gd["LoaiGiaoDich"].astype(str) == "Nhập kho").sum() if not gd.empty else 0
        xuat = (gd["LoaiGiaoDich"].astype(str) == "Xuất kho").sum() if not gd.empty else 0
        st.caption(f"Phiếu nhập: {nhap}  ·  Phiếu xuất: {xuat}")
        if st.button("Tải lại dữ liệu từ Excel", use_container_width=True):
            reload_session_data()
            st.rerun()
        st.divider()
        st.markdown("**Tải file Excel**")
        if ex.DATA_FILE.exists():
            st.download_button(
                label="Tải xuống kho_thuc_an.xlsx",
                data=ex.DATA_FILE.read_bytes(),
                file_name="kho_thuc_an.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    tab_gd, tab_xem, tab_master, tab_bc = st.tabs(
        ["Giao dịch kho", "Xem dữ liệu giao dịch", "Danh mục (Master)", "Báo cáo nhanh"],
    )

    with tab_gd:
        col_btn, col_search = st.columns([1, 3])
        with col_btn:
            if st.button("Tạo mới", type="primary", use_container_width=True):
                clear_form_to_new()
                st.rerun()
        with col_search:
            c1, c2 = st.columns([4, 1])
            with c1:
                q = st.text_input(
                    "Tìm kiếm (bất kỳ cột nào)",
                    value=st.session_state.search_q,
                    key="inp_search",
                    placeholder="Ví dụ: NL001, Nhập kho, K01...",
                )
            with c2:
                st.write("")
                st.write("")
                if st.button("Tìm kiếm", use_container_width=True):
                    st.session_state.search_q = q
                    st.session_state.last_search = q
                    st.rerun()

        with st.expander("Nhập hàng loạt từ Excel (template + kiểm tra)", expanded=False):
            st.markdown(
                "Tải **template** (đủ sheet hướng dẫn và danh mục tham chiếu), điền đúng tên cột trên sheet "
                "**NhapGiaoDich**, xóa các dòng ví dụ (mã phiếu bắt đầu bằng **~**), rồi tải file lên. "
                "Chỉ khi **không còn lỗi** hệ thống mới ghi vào Excel."
            )
            c_tpl, c_up = st.columns([1, 1])
            with c_tpl:
                tpl_bytes = gd_import.build_import_template_bytes(dm_nl, dm_kho, dm_ncc)
                st.download_button(
                    label="Tải template nhập giao dịch (.xlsx)",
                    data=tpl_bytes,
                    file_name="mau_nhap_giao_dich.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="dl_gd_tpl",
                )
            with c_up:
                up_gd = st.file_uploader(
                    "Chọn file Excel đã điền (sheet NhapGiaoDich)",
                    type=["xlsx"],
                    key="up_gd_import",
                )
            if up_gd is not None:
                try:
                    raw_imp = gd_import.read_import_sheet(up_gd.getvalue())
                except Exception as e:
                    st.error(f"Không đọc được file: {e}")
                else:
                    exist_phieu = set(gd["MaPhieu"].dropna().astype(str).str.strip().tolist())
                    rows_ok, errs = gd_import.validate_and_build_rows(
                        raw_imp,
                        dm_nl,
                        dm_kho,
                        dm_ncc,
                        exist_phieu,
                    )
                    if errs:
                        st.error(f"Có **{len(errs)}** lỗi — không ghi dữ liệu. Sửa file và thử lại.")
                        st.code("\n".join(errs), language=None)
                    elif not rows_ok:
                        st.warning("Không có dòng dữ liệu hợp lệ (có thể chỉ còn dòng trống hoặc dòng ví dụ ~).")
                    else:
                        prv = pd.DataFrame(rows_ok)[ex.GD_COLUMNS]
                        st.success(f"**{len(rows_ok)}** dòng hợp lệ — xem trước bên dưới. Bấm **Ghi vào Excel** để lưu.")
                        st.dataframe(prv.head(50), use_container_width=True, hide_index=True)
                        if len(rows_ok) > 50:
                            st.caption(f"Chỉ hiển thị 50/{len(rows_ok)} dòng trong bản xem trước.")
                        if st.button("Ghi vào Excel", type="primary", key="commit_gd_import"):
                            merged = pd.concat([gd, prv], ignore_index=True)
                            ex.save_giao_dich_df(merged)
                            reload_session_data()
                            st.session_state.last_search = ""
                            st.success(f"Đã thêm {len(rows_ok)} giao dịch vào sheet GiaoDich.")
                            st.rerun()

        disp = filtered_gd(gd, st.session_state.last_search)
        st.caption(f"Hiển thị {len(disp)} / {len(gd)} dòng.")

        if not disp.empty:
            labels = []
            id_by_label: dict[str, str] = {}
            for _, r in disp.iterrows():
                rid = str(r.get("ID", ""))
                if not rid:
                    continue
                lbl = f"{str(r.get('MaPhieu', ''))} | {r.get('Ngay', '')} | {str(r.get('TenNguyenLieu', ''))[:30]} | {rid[:8]}…"
                labels.append(lbl)
                id_by_label[lbl] = rid
            pick = st.selectbox("Chọn dòng để sửa / xóa (tùy chọn)", ["— Không chọn —"] + labels)
            if pick != "— Không chọn —":
                sel_id = id_by_label.get(pick)
                if sel_id and st.session_state.edit_id != sel_id:
                    row = gd.loc[gd["ID"].astype(str) == sel_id].iloc[0]
                    form_defaults_from_row(row)
        else:
            st.info("Chưa có dữ liệu hoặc không khớp tìm kiếm.")

        st.subheader("Phiếu nhập / xuất / kiểm kê")
        with st.form("form_gd", clear_on_submit=False):
            fc1, fc2, fc3 = st.columns(3)
            with fc1:
                ngay = st.date_input("Ngày", value=st.session_state.get("form_ngay", date.today()))
            with fc2:
                ma_phieu = st.text_input("Mã phiếu", value=st.session_state.get("form_ma_phieu", ""))
            with fc3:
                loai = st.selectbox(
                    "Loại giao dịch",
                    LOAI_GD_OPTIONS,
                    index=LOAI_GD_OPTIONS.index(st.session_state.get("form_loai", LOAI_GD_OPTIONS[0]))
                    if st.session_state.get("form_loai", LOAI_GD_OPTIONS[0]) in LOAI_GD_OPTIONS
                    else 0,
                )

            nl_codes = (
                dm_nl["MaNL"].astype(str).tolist()
                if not dm_nl.empty and "MaNL" in dm_nl.columns
                else []
            )
            k_codes = (
                dm_kho["MaKho"].astype(str).tolist()
                if not dm_kho.empty and "MaKho" in dm_kho.columns
                else []
            )
            ncc_codes = (
                dm_ncc["MaNCC"].astype(str).tolist()
                if not dm_ncc.empty and "MaNCC" in dm_ncc.columns
                else [""]
            )
            dv_codes = (
                dm_nl["DonViMacDinh"].dropna().astype(str).unique().tolist()
                if not dm_nl.empty and "DonViMacDinh" in dm_nl.columns
                else ["kg", "tan"]
            )

            fr1, fr2, fr3 = st.columns(3)
            with fr1:
                ma_nl = st.selectbox(
                    "Mã nguyên liệu",
                    [""] + nl_codes,
                    index=(
                        ([""] + nl_codes).index(st.session_state.get("form_ma_nl", ""))
                        if st.session_state.get("form_ma_nl", "") in ([""] + nl_codes)
                        else 0
                    ),
                )
            with fr2:
                so_luong = st.number_input("Số lượng", min_value=0.0, value=float(st.session_state.get("form_sl", 0.0)), step=0.001, format="%.3f")
            with fr3:
                don_vi = st.selectbox(
                    "Đơn vị",
                    dv_codes if dv_codes else ["kg"],
                    index=(
                        dv_codes.index(st.session_state.get("form_dv", "kg"))
                        if st.session_state.get("form_dv", "kg") in dv_codes
                        else 0
                    )
                    if dv_codes
                    else 0,
                )

            fr4, fr5, fr6 = st.columns(3)
            with fr4:
                ma_kho = st.selectbox(
                    "Mã kho",
                    [""] + k_codes,
                    index=(
                        ([""] + k_codes).index(st.session_state.get("form_ma_kho", ""))
                        if st.session_state.get("form_ma_kho", "") in ([""] + k_codes)
                        else 0
                    ),
                )
            with fr5:
                ma_ncc = st.selectbox(
                    "Mã NCC (tuỳ chọn)",
                    [""] + ncc_codes,
                    index=(
                        ([""] + ncc_codes).index(st.session_state.get("form_ma_ncc", ""))
                        if st.session_state.get("form_ma_ncc", "") in ([""] + ncc_codes)
                        else 0
                    ),
                )
            with fr6:
                has_hsd = st.checkbox(
                    "Có HSD (lô)",
                    value=st.session_state.get("form_has_hsd", False),
                )
                hsd = st.date_input(
                    "Hạn sử dụng",
                    value=st.session_state.get("form_hsd", date.today()),
                    format="DD/MM/YYYY",
                    disabled=not has_hsd,
                )

            so_lo = st.text_input("Số lô", value=st.session_state.get("form_so_lo", ""))
            ghi_chu = st.text_area("Ghi chú", value=st.session_state.get("form_ghi_chu", ""), height=68)

            bc1, bc2, bc3 = st.columns(3)
            with bc1:
                save = st.form_submit_button("Lưu", type="primary", use_container_width=True)
            with bc2:
                delete = st.form_submit_button("Xóa phiếu đang chọn", use_container_width=True)
            with bc3:
                reset = st.form_submit_button("Xóa form (không xóa DB)", use_container_width=True)

            if reset:
                clear_form_to_new()
                st.rerun()

            if delete:
                eid = st.session_state.edit_id
                if not eid:
                    st.warning("Chưa chọn dòng để xóa. Chọn một dòng ở danh sách phía trên.")
                else:
                    new_df = gd[gd["ID"].astype(str) != str(eid)].reset_index(drop=True)
                    ex.save_giao_dich_df(new_df)
                    reload_session_data()
                    clear_form_to_new()
                    st.success("Đã xóa và cập nhật file Excel.")
                    st.rerun()

            if save:
                if not str(ma_nl).strip():
                    st.error("Vui lòng chọn mã nguyên liệu.")
                elif not str(ma_kho).strip():
                    st.error("Vui lòng chọn mã kho.")
                elif so_luong <= 0:
                    st.error("Số lượng phải lớn hơn 0.")
                elif not str(ma_phieu).strip():
                    st.error("Vui lòng nhập mã phiếu.")
                else:
                    values = {
                        "Ngay": pd.Timestamp(ngay),
                        "MaPhieu": str(ma_phieu).strip(),
                        "LoaiGiaoDich": loai,
                        "MaNguyenLieu": str(ma_nl).strip(),
                        "TenNguyenLieu": "",
                        "SoLuong": so_luong,
                        "DonVi": don_vi,
                        "MaKho": str(ma_kho).strip(),
                        "MaNCC": str(ma_ncc).strip() if ma_ncc else "",
                        "SoLo": str(so_lo).strip(),
                        "HanSuDung": pd.Timestamp(hsd) if has_hsd else pd.NaT,
                        "GhiChu": str(ghi_chu).strip(),
                    }
                    row = ex.row_from_form(values, st.session_state.edit_id)
                    row = ex.enrich_ten_nl(row, dm_nl)
                    eid = str(row["ID"])
                    if st.session_state.edit_id:
                        new_df = gd.copy()
                        mask = new_df["ID"].astype(str) == eid
                        if mask.any():
                            for c in ex.GD_COLUMNS:
                                new_df.loc[mask, c] = row[c]
                        else:
                            new_df = pd.concat([gd, pd.DataFrame([row])], ignore_index=True)
                    else:
                        new_df = pd.concat([gd, pd.DataFrame([row])], ignore_index=True)
                    ex.save_giao_dich_df(new_df)
                    reload_session_data()
                    st.session_state.edit_id = eid
                    st.success("Đã lưu vào sheet GiaoDich trong file Excel.")
                    st.rerun()

        st.subheader("Bảng dữ liệu (theo bộ lọc tìm kiếm)")
        show = disp.copy()
        if not show.empty and "Ngay" in show.columns:
            show["Ngay"] = pd.to_datetime(show["Ngay"], errors="coerce").dt.strftime("%d/%m/%Y")
        if not show.empty and "HanSuDung" in show.columns:
            show["HanSuDung"] = pd.to_datetime(show["HanSuDung"], errors="coerce").dt.strftime("%d/%m/%Y")
        st.dataframe(show, use_container_width=True, hide_index=True)

    with tab_xem:
        render_gd_viewer(gd)

    with tab_master:
        st.markdown(
            "Mỗi danh mục có **Tạo mới**, **Lưu** (thêm khi đang «Thêm mới» hoặc cập nhật khi đã chọn dòng), "
            "và **Xóa** (theo dòng đang chọn). Dữ liệu ghi vào sheet Master trong file Excel."
        )
        tab_labels = ["Nguyên liệu", "Kho", "Nhà cung cấp", "Đơn vị tính"]
        if _is_quantri(st.session_state.auth_user):
            tab_labels.append("Người dùng (quản trị)")
        tabs = st.tabs(tab_labels)
        with tabs[0]:
            render_one_master(SPEC_NL)
        with tabs[1]:
            render_one_master(SPEC_KHO)
        with tabs[2]:
            render_one_master(SPEC_NCC)
        with tabs[3]:
            render_one_master(SPEC_DV)
        if _is_quantri(st.session_state.auth_user):
            with tabs[4]:
                render_user_admin()

    with tab_bc:
        st.subheader("Tổng hợp theo nguyên liệu (số liệu minh họa — không trừ tồn theo FIFO)")
        if gd.empty:
            st.info("Chưa có giao dịch.")
        else:
            work = gd.copy()
            work["SoLuong"] = pd.to_numeric(work["SoLuong"], errors="coerce").fillna(0)
            work["signed"] = work.apply(
                lambda r: r["SoLuong"]
                if str(r.get("LoaiGiaoDich", "")) == "Nhập kho"
                else (-r["SoLuong"] if str(r.get("LoaiGiaoDich", "")) == "Xuất kho" else 0.0),
                axis=1,
            )
            pivot = (
                work.groupby(["MaNguyenLieu", "TenNguyenLieu"], dropna=False)["signed"]
                .sum()
                .reset_index(name="BienDongLuyKe")
            )
            st.dataframe(pivot, use_container_width=True, hide_index=True)
            st.caption(
                "Cột **BienDongLuyKe**: cộng số lượng phiếu Nhập kho, trừ Xuất kho (các loại khác tạm = 0). "
                "Có thể mở rộng logic tồn kho chính xác sau."
            )
        buf = BytesIO()
        if not gd.empty:
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                gd.to_excel(w, sheet_name="XuatBaoCao", index=False)
            st.download_button(
                "Xuất báo cáo nhanh (chỉ sheet giao dịch hiện tại)",
                data=buf.getvalue(),
                file_name="bao_cao_nhanh.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()
