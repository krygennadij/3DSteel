# -*- coding: utf-8 -*-
"""Тесты расчётного ядра calc.py.

Запуск:  .venv\\Scripts\\python.exe tests\\test_calc.py
(или через pytest, если он установлен)
"""
import os
import sys

import numpy as np

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))

from calc import SteelDatabase, compute  # noqa: E402


def get_db() -> SteelDatabase:
    return SteelDatabase.from_dir(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".."))


def test_reference_case():
    """Контрольный пример: ГОСТ 26020-83, 20Б1, С345, 2000 кг, 6 м."""
    res = compute(get_db(), "ГОСТ 26020-83", "20Б1", "С345", 2000.0, 6.0)
    cap = res.load_capacity["Несущая способность, кНм"].to_numpy()
    assert abs(cap[0] - 87.4415) < 1e-3, cap[0]
    assert abs(res.applied_moment_value - 30.4280) < 1e-3, res.applied_moment_value
    assert res.fire_limit_minute == 22, res.fire_limit_minute


def test_capacity_decreases():
    """Несущая способность при нагреве не растёт (на конечных значениях)."""
    res = compute(get_db(), "СТО АСЧМ 20-93", "100Ш1", "С255", 1000.0, 6.0)
    cap = res.load_capacity["Несущая способность, кНм"].to_numpy()
    finite = cap[np.isfinite(cap)]
    assert len(finite) > 5
    assert (np.diff(finite) <= 1e-9).all()


def test_no_load_no_limit():
    """Без внешней нагрузки крупный профиль предела не достигает."""
    res = compute(get_db(), "СТО АСЧМ 20-93", "100Ш1", "С255", 0.0, 0.0)
    assert res.applied_moment_value == 0.0
    # capacity > 0 хотя бы в начале
    cap0 = res.load_capacity["Несущая способность, кНм"].iloc[0]
    assert cap0 > 0


def test_all_profiles_compute():
    """compute() не падает ни на одном профиле сортамента."""
    db = get_db()
    n = 0
    for doc in db.profile_docs:
        for key in db.get_profile_data(doc).index:
            compute(db, doc, key, "С255", 1000.0, 6.0)
            n += 1
    assert n > 400, n


if __name__ == "__main__":
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            fn()
            print(f"OK  {name}")
    print("ALL TESTS PASSED")
