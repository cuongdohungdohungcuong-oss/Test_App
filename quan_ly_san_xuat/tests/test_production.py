from datetime import date
from pathlib import Path
import sys

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import excel_io as ex  # noqa: E402
import production as pr  # noqa: E402


@pytest.fixture
def dm_ct():
    return ex._default_dm_cong_thuc()


@pytest.fixture
def dm_nl():
    return ex._default_dm_nguyen_lieu()


def _lenh(**kw):
    base = {c: "" for c in ex.LSX_COLUMNS}
    base.update(
        ID="1", MaLenh="LSX-20261004-001", NgayKH=pd.Timestamp("2026-10-04"), MaTP="TP001",
        TenTP="Cám heo", SanLuongKH=2000, DonVi="kg", MaDC="DC01", Ca="Ca 1",
        MaKhoNL="K01", MaKhoTP="K02", SoLo="L1", TrangThai=pr.TT_DANG_SX,
    )
    base.update(kw)
    return base


def test_next_ma_lenh():
    d = date(2026, 10, 4)
    assert pr.next_ma_lenh(pd.DataFrame(columns=["MaLenh"]), d) == "LSX-20261004-001"
    df = pd.DataFrame({"MaLenh": ["LSX-20261004-001", "LSX-20261004-007", "LSX-20261003-020"]})
    assert pr.next_ma_lenh(df, d) == "LSX-20261004-008"


def test_default_bom_valid(dm_ct):
    for tp in ("TP001", "TP002"):
        assert pr.validate_bom(pr.bom_for(tp, dm_ct)) == []


def test_validate_bom_errors():
    bom = pd.DataFrame({"MaNL": ["NL001", "NL001"], "DinhMuc": [300, -1]})
    errs = pr.validate_bom(bom)
    assert any("lặp" in e for e in errs)
    assert any("> 0" in e for e in errs)
    assert any("Tổng định mức" in e for e in errs)


def test_planned_consumption(dm_ct, dm_nl):
    pc = pr.planned_consumption("TP001", 2000, dm_ct, dm_nl)
    assert pc["DinhMucKH"].sum() == pytest.approx(2000)
    assert pc.set_index("MaNL").loc["NL001", "DinhMucKH"] == pytest.approx(1100)
    assert pc.set_index("MaNL").loc["NL001", "TenNL"] == "Bắp hạt"


def test_transitions():
    assert pr.can_transition(pr.TT_KE_HOACH, pr.TT_DANG_SX)
    assert pr.can_transition(pr.TT_DANG_SX, pr.TT_HOAN_THANH)
    assert not pr.can_transition(pr.TT_KE_HOACH, pr.TT_HOAN_THANH)
    assert not pr.can_transition(pr.TT_HOAN_THANH, pr.TT_HUY)


def test_complete_order(dm_ct, dm_nl):
    lenh = _lenh()
    lsx = pd.DataFrame([lenh])
    th = pr.planned_consumption("TP001", 2000, dm_ct, dm_nl)
    th["SoLuongTT"] = th["DinhMucKH"] * 1.01
    gd = pd.DataFrame(columns=ex.GD_COLUMNS)
    lsx2, thsx2, gd2 = pr.complete_order(lenh, 1980, th, lsx, pd.DataFrame(), gd, date(2026, 10, 5))

    assert lsx2.loc[0, "TrangThai"] == pr.TT_HOAN_THANH
    assert lsx2.loc[0, "SanLuongTT"] == 1980
    assert len(thsx2) == 4
    x = gd2[gd2["MaPhieu"] == "LSX-20261004-001-X"]
    n = gd2[gd2["MaPhieu"] == "LSX-20261004-001-N"]
    assert len(x) == 4 and set(x["LoaiGiaoDich"]) == {"Xuất kho"} and set(x["MaKho"]) == {"K01"}
    assert len(n) == 1 and n.iloc[0]["SoLuong"] == 1980 and n.iloc[0]["MaKho"] == "K02"

    # Không cho hoàn thành lại / sinh trùng phiếu
    with pytest.raises(ValueError):
        pr.complete_order(lsx2.iloc[0].to_dict(), 1980, th, lsx2, thsx2, gd2, date(2026, 10, 5))
    with pytest.raises(ValueError):
        pr.complete_order(lenh, 1980, th, lsx, thsx2, gd2, date(2026, 10, 5))

    var = pr.report_variance(thsx2)
    assert var["TyLeHaoHut"].round(1).eq(1.0).all()


def test_report_output():
    lsx = pd.DataFrame(
        [
            _lenh(MaLenh="A", TrangThai=pr.TT_HOAN_THANH, SanLuongKH=1000, SanLuongTT=950),
            _lenh(MaLenh="B", TrangThai=pr.TT_DANG_SX, SanLuongKH=1000),
            _lenh(MaLenh="C", TrangThai=pr.TT_HUY, SanLuongKH=5000),
        ]
    )
    r = pr.report_output(lsx, "MaDC")
    assert r.loc[0, "SoLenh"] == 2
    assert r.loc[0, "SoLenhHoanThanh"] == 1
    assert r.loc[0, "SanLuongKH"] == 2000
    assert r.loc[0, "TyLeHoanThanh"] == pytest.approx(47.5)
