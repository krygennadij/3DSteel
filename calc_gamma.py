# -*- coding: utf-8 -*-
"""Расчётное ядро: альтернативная методика оценки предела огнестойкости
изгибаемого элемента (двутавра) по СП 16.13330 — через коэффициент
использования несущей способности γ_T, критическую температуру стали и
равномерный прогрев сечения по приведённой толщине металла δ_np = A / П.

Реализация по мотивам приложения https://github.com/krygennadij/fireresiscience
(модули src/structural.py, src/thermal.py, src/data.py), адаптированная под
данные и соглашения этого проекта:
  - критическая температура определяется не по грубой 3-групповой таблице
    оригинала, а обратной интерполяцией по точной таблице прочности
    prochnost.json (та же таблица, что использует calc.py для основной
    методики) — это точнее и работает для всех марок стали проекта;
  - геометрия сечения берётся из сортамента (beam_profiles.json) или
    вычисляется по формулам для произвольного двутавра.

В отличие от calc.py (детальный расчёт с раздельным прогревом трёх участков
сечения по готовым кривым температуры), здесь предполагается один и тот же
прогрев по всему сечению, а исходные данные о прогреве не нужны — они
вычисляются из геометрии сечения и стандартного температурного режима пожара.

Модуль не зависит от Streamlit.
"""
from dataclasses import dataclass
from typing import Optional

import math

import numpy as np
import pandas as pd

from calc import SteelDatabase

# Параметры теплофизической модели (как в первоисточнике)
_RHO_STEEL = 7800.0   # кг/м3 — плотность стали
_EMISSIVITY = 0.563   # приведённая степень черноты
_C_STEEL = 310.0      # Дж/(кг·К) — базовая удельная теплоёмкость
_D_STEEL = 0.48       # коэффициент роста теплоёмкости с температурой
_T0_K = 293.0         # начальная температура, К (≈ 20 °C)


def standard_fire_curve_k(t_sec: float) -> float:
    """Стандартный температурный режим пожара, К. t_sec — время, с."""
    return 345.0 * math.log10((8.0 / 60.0) * t_sec + 1.0) + _T0_K


# ── Геометрия сечения ────────────────────────────────────────────────────────

def geometry_from_profile(profile: pd.Series) -> dict:
    """Геометрические характеристики двутавра из строки сортамента
    (все нужные величины уже посчитаны в beam_profiles.json)."""
    h, b, tf, tw = (profile["h, мм"], profile["b, мм"],
                    profile["t, мм"], profile["s, мм"])
    return {
        "h": h, "b": b, "tf": tf, "tw": tw,
        "A": profile["А, см2"] * 100.0,      # см² -> мм²
        "Ix": profile["Ix, см4"] * 1e4,      # см⁴ -> мм⁴
        "Wx": profile["Wx, см3"] * 1e3,      # см³ -> мм³
        "Sx": profile["Sx, см3"] * 1e3,      # см³ -> мм³
        "Af": b * tf,
        "Aw": (h - 2 * tf) * tw,
    }


def geometry_from_dims(h: float, b: float, tf: float, tw: float) -> dict:
    """Геометрия по формулам сопромата — для сечения, заданного вручную."""
    af = b * tf
    h_web = h - 2 * tf
    aw = h_web * tw
    area = 2 * af + aw
    b_inner = (b - tw) / 2.0
    ix = (b * h ** 3) / 12.0 - 2 * (b_inner * h_web ** 3) / 12.0
    wx = ix / (h / 2.0)
    dist_f = h / 2.0 - tf / 2.0
    dist_w = h_web / 4.0
    sx = af * dist_f + (tw * h_web / 2.0) * dist_w
    return {"h": h, "b": b, "tf": tf, "tw": tw,
            "A": area, "Ix": ix, "Wx": wx, "Sx": sx, "Af": af, "Aw": aw}


# ── Прочностная задача: γ_T ──────────────────────────────────────────────────

def calc_c1_coefficient(af: float, aw: float) -> dict:
    """Коэффициент c1 (учёт развития пластических деформаций), Табл. Е.1
    СП 16.13330. Линейная интерполяция между n = Af/Aw = 0.5 (c1 = 1.07)
    и n = 1.0 (c1 = 1.12), с ограничением диапазоном [1.0; 1.2]."""
    if aw <= 0:
        return {"value": 1.0, "n": 0.0, "trace": None}
    n = af / aw
    c1 = 1.07 + (n - 0.5) * 0.1
    c1 = min(max(c1, 1.0), 1.2)
    return {"value": c1, "n": n, "trace": {"low": (0.5, 1.07), "high": (1.0, 1.12)}}


def calc_gamma_bending(m_load_nm: float, wx_m3: float, ry_pa: float,
                        gamma_c: float = 1.0, c1: float = 1.0) -> float:
    """γ_T = M / (c1 · Wx · Ry · γc)."""
    if wx_m3 <= 0:
        return float("inf")
    return m_load_nm / (c1 * wx_m3 * ry_pa * gamma_c)


def calc_gamma_shear(q_load_n: float, sx_m3: float, ix_m4: float, tw_m: float,
                      ry_pa: float, gamma_c: float = 1.0) -> float:
    """γ_T = Q·Sx / (Ix·tw·Rs·γc), Rs = 0.58·Ry."""
    if ix_m4 <= 0 or tw_m <= 0:
        return float("inf")
    rs_pa = 0.58 * ry_pa
    return (q_load_n * sx_m3) / (ix_m4 * tw_m * rs_pa * gamma_c)


# ── Критическая температура (по таблице прочности проекта) ──────────────────

def get_critical_temperature(db: SteelDatabase, grade: str, gamma_target: float) -> dict:
    """Критическая температура T_cr, при которой Ry(T)/Ry(20°C) = gamma_target.
    Обратная линейная интерполяция по prochnost.json (та же таблица, что и в
    основной методике calc.py — компонента k = yield(T)/yield(20°C))."""
    strength_df = db.get_strength_data()
    if grade not in strength_df.columns:
        raise ValueError(f"Марка стали {grade} не найдена в таблице прочности.")

    temps = strength_df["Температура, ℃"].to_numpy(dtype=float)
    ry_vals = strength_df[grade].to_numpy(dtype=float)
    ry0 = ry_vals[0]
    if ry0 <= 0:
        raise ValueError("Нулевой предел текучести при 20°C для марки стали.")
    ratios = ry_vals / ry0

    if gamma_target >= 1.0:
        return {"value": float(temps[0]), "trace": None}
    if gamma_target <= ratios[-1]:
        return {"value": float(temps[-1]), "trace": None}

    for i in range(len(temps) - 1):
        g1, g2 = ratios[i], ratios[i + 1]
        t1, t2 = temps[i], temps[i + 1]
        if g1 >= gamma_target >= g2:
            t_cr = t1 if g1 == g2 else t1 + (gamma_target - g1) * (t2 - t1) / (g2 - g1)
            return {"value": float(t_cr), "trace": {"t1": t1, "g1": g1, "t2": t2, "g2": g2}}

    return {"value": float(temps[-1]), "trace": None}


# ── Теплотехническая задача: прогрев по приведённой толщине ─────────────────

def calc_heated_perimeter(h: float, b: float, tw: float, exposure: str = "4_sides") -> float:
    """Обогреваемый периметр двутавра, мм (Табл. 4.2 методики)."""
    if exposure == "3_sides":
        return 2 * h + 3 * b - 2 * tw
    return 2 * h + 4 * b - 2 * tw


def simulate_heating(delta_np_mm: float, crit_temp_c: float,
                      max_time_min: float = 60.0, dt_sec: float = 1.0) -> dict:
    """Нагрев стали по стандартному температурному режиму пожара при
    равномерном прогреве сечения (приведённая толщина δ_np = A / П)."""
    if delta_np_mm <= 0:
        raise ValueError("Приведённая толщина металла должна быть положительной.")
    delta_np_m = delta_np_mm / 1000.0

    t_steel_k = _T0_K
    crit_temp_k = crit_temp_c + 273.15

    max_time_sec = max_time_min * 60.0
    steps = int(max_time_sec / dt_sec)

    times_min = [0.0]
    gas_c = [_T0_K - 273.15]
    steel_c = [_T0_K - 273.15]
    fire_limit_min: Optional[float] = None
    t_sec = 0.0

    for _ in range(steps):
        t_sec += dt_sec
        t_gas_k = standard_fire_curve_k(t_sec)
        diff = t_gas_k - t_steel_k
        if abs(diff) < 0.1:
            alpha = 29.0
        else:
            alpha = 29.0 + 5.67 * _EMISSIVITY * (
                (t_gas_k / 100.0) ** 4 - (t_steel_k / 100.0) ** 4
            ) / diff
        c_steel = _C_STEEL + _D_STEEL * t_steel_k
        delta_t = (dt_sec * alpha * diff) / (_RHO_STEEL * delta_np_m * c_steel)
        t_steel_k += delta_t

        times_min.append(t_sec / 60.0)
        gas_c.append(t_gas_k - 273.15)
        steel_c.append(t_steel_k - 273.15)

        if fire_limit_min is None and t_steel_k >= crit_temp_k:
            fire_limit_min = t_sec / 60.0

    history = pd.DataFrame({
        "Время, мин": times_min,
        "Газовая среда, °C": gas_c,
        "Сталь (равномерный прогрев), °C": steel_c,
    })
    return {"history": history, "fire_limit_minute": fire_limit_min}


# ── Полный расчёт ─────────────────────────────────────────────────────────────

@dataclass
class BendingGammaResult:
    geometry: dict
    c1: float
    c1_trace: Optional[dict]
    gamma_bending: float
    gamma_shear: float
    gamma_t: float
    critical_temp: float
    critical_temp_trace: Optional[dict]
    perimeter_mm: float
    delta_np_mm: float
    history: pd.DataFrame
    fire_limit_minute: Optional[float]
    m_load_knm: float
    capacity_curve: pd.DataFrame          # "Время, мин" (целые), "Несущая способность, кНм"
    capacity_fire_limit_minute: Optional[int]


@dataclass
class GammaCapacityCurve:
    """Тонкая обёртка над результатом метода γ_T с теми же тремя атрибутами,
    что использует main.make_comparison_chart для FireCalcResult — позволяет
    наложить кривую несущей способности по γ_T на общий график сравнения
    без изменения самой функции построения графика."""
    load_capacity: pd.DataFrame
    applied_moment_value: float
    fire_limit_minute: Optional[int]


def as_capacity_curve_result(result: "BendingGammaResult") -> GammaCapacityCurve:
    return GammaCapacityCurve(
        load_capacity=result.capacity_curve,
        applied_moment_value=result.m_load_knm,
        fire_limit_minute=result.capacity_fire_limit_minute,
    )


def _capacity_curve_per_minute(
    db: SteelDatabase, grade: str, m_load_knm: float, gamma_t: float,
    history: pd.DataFrame, max_time_min: float,
) -> dict:
    """Несущая способность (кНм) на целых минутах — по аналогии с
    load_capacity в calc.FireCalcResult, чтобы кривую можно было сравнивать
    с результатами основной методики на одном графике.

    При 20°C несущая способность равна M / γ_T (запас по худшей из проверок —
    изгиб или сдвиг). При нагреве оба коэффициента γ_T масштабируются с Ry(T)
    одинаково (Rs = 0.58·Ry), поэтому несущая способность снижается
    пропорционально тому же коэффициенту k(T) = Ry(T)/Ry(20°C), что и в
    основной методике (calc.py)."""
    minutes = np.arange(int(max_time_min) + 1)
    t_hist = history["Время, мин"].to_numpy(dtype=float)
    steel_hist = history["Сталь (равномерный прогрев), °C"].to_numpy(dtype=float)
    steel_at_minutes = np.interp(minutes, t_hist, steel_hist)

    yield_at_t = db.interpolate_strength(pd.Series(steel_at_minutes), grade)
    k_t = yield_at_t / yield_at_t[0]

    m_cap0 = m_load_knm / gamma_t if gamma_t > 0 else float("inf")
    capacity = m_cap0 * k_t

    load_capacity = pd.DataFrame({
        "Время, мин": minutes,
        "Несущая способность, кНм": capacity,
    })

    holds = capacity > m_load_knm
    if holds.all():
        fire_limit = None
    elif not holds.any():
        fire_limit = 0
    else:
        fire_limit = int(minutes[holds][-1])

    return {"load_capacity": load_capacity, "fire_limit_minute": fire_limit}


def compute_bending_gamma(
    db: SteelDatabase,
    grade: str,
    ry_mpa: float,
    m_load_knm: float,
    q_load_kn: float,
    exposure: str = "4_sides",
    gamma_c: float = 1.0,
    max_time_min: float = 60.0,
    profile: Optional[pd.Series] = None,
    dims: Optional[dict] = None,
) -> BendingGammaResult:
    """Расчёт предела огнестойкости изгибаемого двутавра методом γ_T.

    Геометрия берётся из `profile` (строка сортамента) либо из `dims`
    (словарь {"h", "b", "tf", "tw"} в мм) — должен быть задан ровно один
    из этих параметров.
    """
    if profile is not None:
        geometry = geometry_from_profile(profile)
    elif dims is not None:
        geometry = geometry_from_dims(dims["h"], dims["b"], dims["tf"], dims["tw"])
    else:
        raise ValueError("Не заданы геометрия сечения (profile или dims).")

    c1_res = calc_c1_coefficient(geometry["Af"], geometry["Aw"])

    ry_pa = ry_mpa * 1e6
    wx_m3 = geometry["Wx"] * 1e-9
    ix_m4 = geometry["Ix"] * 1e-12
    sx_m3 = geometry["Sx"] * 1e-9
    tw_m = geometry["tw"] * 1e-3

    m_nm = abs(m_load_knm) * 1000.0
    q_n = abs(q_load_kn) * 1000.0

    gamma_bend = calc_gamma_bending(m_nm, wx_m3, ry_pa, gamma_c, c1_res["value"])
    gamma_shear = calc_gamma_shear(q_n, sx_m3, ix_m4, tw_m, ry_pa, gamma_c)
    gamma_t = max(gamma_bend, gamma_shear)

    crit_res = get_critical_temperature(db, grade, gamma_t)

    perimeter_mm = calc_heated_perimeter(geometry["h"], geometry["b"], geometry["tw"], exposure)
    delta_np_mm = geometry["A"] / perimeter_mm if perimeter_mm > 0 else 0.0

    heating = simulate_heating(delta_np_mm, crit_res["value"], max_time_min=max_time_min)

    capacity = _capacity_curve_per_minute(
        db, grade, m_load_knm, gamma_t, heating["history"], max_time_min,
    )

    return BendingGammaResult(
        geometry=geometry,
        c1=c1_res["value"], c1_trace=c1_res["trace"],
        gamma_bending=gamma_bend, gamma_shear=gamma_shear, gamma_t=gamma_t,
        critical_temp=crit_res["value"], critical_temp_trace=crit_res["trace"],
        perimeter_mm=perimeter_mm, delta_np_mm=delta_np_mm,
        history=heating["history"], fire_limit_minute=heating["fire_limit_minute"],
        m_load_knm=m_load_knm,
        capacity_curve=capacity["load_capacity"],
        capacity_fire_limit_minute=capacity["fire_limit_minute"],
    )
