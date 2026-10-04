"""Sinh lệnh sản xuất mẫu (kèm tiêu hao và phiếu kho) để minh họa báo cáo."""
from __future__ import annotations

import random
from datetime import date, timedelta

import pandas as pd

import excel_io as ex
import production as pr


def append_sample_orders(n: int = 30, days: int = 30, seed: int | None = None) -> int:
    """Thêm n lệnh SX trong `days` ngày gần nhất; phần lớn hoàn thành, còn lại đang SX / kế hoạch / hủy."""
    rnd = random.Random(seed)
    lsx = pr.norm_df(ex.read_sheet(ex.SHEET_LSX), ex.LSX_COLUMNS).astype(object)
    thsx = ex.read_sheet(ex.SHEET_THSX)
    gd = ex.read_sheet(ex.SHEET_GD)
    dm_ct = ex.read_sheet(ex.SHEET_CT)
    dm_nl = ex.read_sheet(ex.SHEET_NL)
    tp = pr.norm_df(ex.read_sheet(ex.SHEET_TP), ex.TP_COLUMNS)
    dc = pr.norm_df(ex.read_sheet(ex.SHEET_DC), ex.DC_COLUMNS)
    tp = tp[tp["MaTP"].astype(str).map(lambda m: not pr.bom_for(m, dm_ct).empty)]
    if tp.empty:
        return 0
    dc_codes = dc["MaDC"].astype(str).tolist() or [""]

    today = date.today()
    added = 0
    for _ in range(n):
        ngay = today - timedelta(days=rnd.randint(0, days))
        t = tp.iloc[rnd.randrange(len(tp))]
        sl_kh = float(rnd.choice([5, 8, 10, 12, 15, 20]) * 1000)
        lenh = {c: "" for c in ex.LSX_COLUMNS}
        lenh.update(
            ID=ex.new_id(),
            MaLenh=pr.next_ma_lenh(lsx, ngay),
            NgayKH=pd.Timestamp(ngay),
            MaTP=str(t["MaTP"]),
            TenTP=str(t["TenTP"]),
            SanLuongKH=sl_kh,
            DonVi=str(t["DonVi"] or "kg"),
            MaDC=rnd.choice(dc_codes),
            Ca=rnd.choice(pr.CA_OPTIONS),
            MaKhoNL="K01",
            MaKhoTP="K02",
            SoLo=f"MAU-{ngay:%y%m%d}-{rnd.randint(100, 999)}",
            TrangThai=pr.TT_DANG_SX,
            NguoiTao="mau",
            GhiChu="Dữ liệu mẫu",
        )
        lsx = pd.concat([lsx, pd.DataFrame([lenh])], ignore_index=True)
        added += 1

        roll = rnd.random()
        if ngay >= today - timedelta(days=2) and roll < 0.5:
            lsx.loc[lsx["MaLenh"] == lenh["MaLenh"], "TrangThai"] = pr.TT_KE_HOACH if roll < 0.25 else pr.TT_DANG_SX
            continue
        if roll < 0.05:
            lsx.loc[lsx["MaLenh"] == lenh["MaLenh"], "TrangThai"] = pr.TT_HUY
            continue
        sl_tt = round(sl_kh * rnd.uniform(0.92, 1.02), 1)
        th = pr.planned_consumption(lenh["MaTP"], sl_tt, dm_ct, dm_nl)
        th["SoLuongTT"] = (th["DinhMucKH"] * [rnd.uniform(0.99, 1.03) for _ in range(len(th))]).round(2)
        lsx, thsx, gd = pr.complete_order(lenh, sl_tt, th, lsx, thsx, gd, ngay)

    ex.save_sheets({ex.SHEET_LSX: lsx, ex.SHEET_THSX: thsx, ex.SHEET_GD: gd})
    return added


if __name__ == "__main__":
    print(f"Đã thêm {append_sample_orders(30)} lệnh sản xuất mẫu.")
