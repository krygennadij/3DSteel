# -*- coding: utf-8 -*-
"""Тесты расчётного ядра calc_gamma.py (методика γ_T для изгибаемых элементов).

Запуск:  .venv\\Scripts\\python.exe tests\\test_calc_gamma.py
"""
import os
import sys

import numpy as np

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))

from calc import SteelDatabase  # noqa: E402
from calc_gamma import (  # noqa: E402
    compute_bending_gamma, get_critical_temperature, as_capacity_curve_result,
)


def get_db() -> SteelDatabase:
    return SteelDatabase.from_dir(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))


def test_critical_temperature_bounds():
    db = get_db()
    # γ_T >= 1 -> критическая температура равна 20°C
    assert get_critical_temperature(db, "С345", 1.5)["value"] == 20.0
    # Очень маленький γ_T -> температура не выше последней точки таблицы
    res = get_critical_temperature(db, "С345", 0.01)
    assert res["value"] <= 850.0


def test_reference_case_ibeam():
    """20Б1, С345, момент/поперечная сила в разумных пределах."""
    db = get_db()
    profile_df = db.get_profile_data("ГОСТ 26020-83")
    profile = profile_df.loc["20Б1"].squeeze()
    ry0 = db.get_strength_data()["С345"].iloc[0]

    res = compute_bending_gamma(
        db, "С345", ry0,
        m_load_knm=30.0, q_load_kn=40.0,
        exposure="4_sides", profile=profile,
    )
    assert 0.0 < res.gamma_t < 1.0, res.gamma_t
    assert 20.0 < res.critical_temp < 800.0, res.critical_temp
    assert res.delta_np_mm > 0
    assert res.fire_limit_minute is None or res.fire_limit_minute > 0


def test_overload_gives_zero_limit():
    """Если γ_T >= 1 (перегрузка уже при 20°C), предел огнестойкости ~ 0."""
    db = get_db()
    profile_df = db.get_profile_data("ГОСТ 26020-83")
    profile = profile_df.loc["20Б1"].squeeze()
    ry0 = db.get_strength_data()["С345"].iloc[0]

    res = compute_bending_gamma(
        db, "С345", ry0,
        m_load_knm=100000.0, q_load_kn=1.0,
        exposure="4_sides", profile=profile,
    )
    assert res.gamma_t >= 1.0
    assert res.critical_temp == 20.0
    assert res.fire_limit_minute is not None
    assert res.fire_limit_minute < 1.0


def test_custom_dims():
    db = get_db()
    res = compute_bending_gamma(
        db, "С255", 240.0,
        m_load_knm=15.0, q_load_kn=20.0,
        exposure="3_sides",
        dims={"h": 300.0, "b": 150.0, "tf": 10.2, "tw": 6.5},
    )
    assert res.geometry["Wx"] > 0
    assert res.perimeter_mm > 0


def test_capacity_curve_matches_gamma_t_and_moment():
    """Кривая несущей способности: значение при 20°C = M/γ_T, монотонно убывает,
    а момент пересечения (целая минута) согласован с непрерывным расчётом
    (fire_limit_minute), в пределах шага 1 мин."""
    db = get_db()
    profile = db.get_profile_data("ГОСТ 26020-83").loc["20Б1"].squeeze()
    ry0 = db.get_strength_data()["С345"].iloc[0]

    res = compute_bending_gamma(
        db, "С345", ry0, m_load_knm=15.963, q_load_kn=10.141,
        exposure="4_sides", profile=profile, max_time_min=60,
    )
    cap = res.capacity_curve["Несущая способность, кНм"].to_numpy()

    assert abs(cap[0] - (15.963 / res.gamma_t)) < 1e-6
    assert (np.diff(cap) <= 1e-9).all()
    assert res.capacity_fire_limit_minute is not None
    assert abs(res.capacity_fire_limit_minute - res.fire_limit_minute) <= 1.0

    cc = as_capacity_curve_result(res)
    assert cc.applied_moment_value == 15.963
    assert cc.fire_limit_minute == res.capacity_fire_limit_minute
    assert list(cc.load_capacity.columns) == ["Время, мин", "Несущая способность, кНм"]


if __name__ == "__main__":
    tests = [v for k, v in list(globals().items()) if k.startswith("test_")]
    for t in tests:
        t()
        print(f"OK: {t.__name__}")
    print(f"\n{len(tests)} тестов пройдено.")
