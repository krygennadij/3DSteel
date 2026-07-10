# -*- coding: utf-8 -*-
"""Расчётное ядро: предел огнестойкости стальной двутавровой балки
с учётом неравномерного прогрева сечения.

Модуль не зависит от Streamlit — его можно импортировать в тестах
и других программах.
"""
from dataclasses import dataclass
from typing import Optional

import json
import os

import numpy as np
import pandas as pd

# Ускорение свободного падения, м/с² — единое для всего проекта: используется
# и при переводе МПа в кгс/см², и при переводе усилий (кгс) в моменты (кН·м),
# и при переводе нагрузки/массы (кг) в момент от нагрузки (кН·м). Раньше эти
# переводы были рассогласованы (местами использовалось g≈10), из-за чего
# несущая способность при 20°C отличалась от точного пластического момента
# сечения (Ry·Wpl) примерно на 2% — единое g устраняет это расхождение.
G = 9.81

# Перевод МПа в кгс/см²
MPA_TO_KGF_CM2 = 100.0 / G


class SteelDatabase:
    """Доступ к данным: температурный режим, прочность сталей, сортамент."""

    def __init__(self, temp_path: str, strength_path: str, profile_path: str):
        self.temperature_data = self._load_json(temp_path)
        self.strength_data = self._load_json(strength_path)
        self.profile_data = self._load_json(profile_path)
        self.profile_docs = list(self.profile_data.keys())

    @classmethod
    def from_dir(cls, base_dir: Optional[str] = None) -> "SteelDatabase":
        """База из JSON-файлов, лежащих рядом с этим модулем."""
        if base_dir is None:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        return cls(
            os.path.join(base_dir, "temperature.json"),
            os.path.join(base_dir, "prochnost.json"),
            os.path.join(base_dir, "beam_profiles.json"),
        )

    @staticmethod
    def _load_json(path: str) -> dict:
        try:
            with open(path, "r", encoding="utf-8") as file:
                return json.load(file)
        except FileNotFoundError:
            raise FileNotFoundError(f"Файл {path} не найден.")
        except json.JSONDecodeError:
            raise ValueError(f"Ошибка декодирования JSON в файле {path}.")

    def get_temperature_data(self, component: str) -> pd.Series:
        if component not in self.temperature_data:
            raise ValueError(f"Данные о температуре для компонента {component} не найдены.")
        return pd.Series(self.temperature_data[component])

    def get_strength_data(self) -> pd.DataFrame:
        return pd.DataFrame(self.strength_data)

    @property
    def steel_grades(self) -> list:
        grades = list(self.strength_data.keys())
        grades.remove("Температура, ℃")
        return grades

    def get_profile_data(self, doc: str) -> pd.DataFrame:
        return pd.DataFrame(self.profile_data.get(doc, {})).T

    def interpolate_strength(self, temperature_series: pd.Series, steel_grade: str) -> np.ndarray:
        strength_df = self.get_strength_data()
        if steel_grade not in strength_df.columns:
            raise ValueError(f"Марка стали {steel_grade} не найдена в данных.")
        return np.interp(temperature_series, strength_df["Температура, ℃"], strength_df[steel_grade])


@dataclass
class FireCalcResult:
    """Все таблицы расчёта + итоговый предел огнестойкости."""

    strength_table: pd.DataFrame        # температуры, предел текучести, коэффициенты снижения
    compressed_zone: pd.DataFrame       # показатель сжатой зоны
    normative_resistance: pd.DataFrame  # нормативное сопротивление, кгс/см²
    efforts: pd.DataFrame               # усилия растяжения и сжатия
    lever_arms: pd.DataFrame            # плечи равнодействующих сил, мм
    bending_moments: pd.DataFrame       # изгибающие моменты, кНм
    load_capacity: pd.DataFrame         # несущая способность, кНм
    applied_moment: pd.DataFrame        # момент от нагрузки, кНм
    applied_moment_value: float         # тот же момент числом
    fire_limit_minute: Optional[int]    # первая минута, когда несущая способность < момента
                                        # None — предел не достигнут в рассматриваемом диапазоне


def compute(db: SteelDatabase, doc: str, profile_key: str, grade: str,
            load_kg: float = 0.0, length_m: float = 0.0,
            temp_lower_ext: Optional[np.ndarray] = None,
            temp_web_ext: Optional[np.ndarray] = None,
            temp_upper_ext: Optional[np.ndarray] = None,
            custom_profile: Optional[dict] = None) -> FireCalcResult:
    """Полный расчёт по методике для одного профиля и марки стали.

    Если переданы temp_lower_ext / temp_web_ext / temp_upper_ext —
    используются они вместо данных из базы.
    Если передан custom_profile (ключи "b, мм", "t, мм", "s, мм", "h, мм",
    "M, кг") — используется он вместо поиска профиля в сортаменте (doc/profile_key
    в этом случае не используются для расчёта, только для подписи отчёта).
    """
    if custom_profile is not None:
        profile = pd.Series(custom_profile)
    else:
        profile_df = db.get_profile_data(doc)
        if profile_df.empty or profile_key not in profile_df.index:
            raise ValueError(f"Данные для профиля {profile_key} не найдены.")
        profile = profile_df.loc[profile_key].squeeze()
    b = profile.get("b, мм")
    t = profile.get("t, мм")
    s = profile.get("s, мм")
    h = profile.get("h, мм")
    m = profile.get("M, кг")
    if b is None or t is None or s is None or h is None:
        raise ValueError("Ошибка в данных профиля.")

    # 1. Температуры трёх участков сечения по минутам
    if temp_lower_ext is not None:
        temp_lower = pd.Series(temp_lower_ext, dtype=float)
        temp_web   = pd.Series(temp_web_ext,   dtype=float)
        temp_upper = pd.Series(temp_upper_ext, dtype=float)
    else:
        temp_lower = db.get_temperature_data("lower_flange")
        temp_web   = db.get_temperature_data("web")
        temp_upper = db.get_temperature_data("upper_flange")
    minutes = np.arange(len(temp_lower))

    # 2. Предел текучести при нагреве и коэффициенты его снижения
    yield_lower = db.interpolate_strength(temp_lower, grade)
    yield_web = db.interpolate_strength(temp_web, grade)
    yield_upper = db.interpolate_strength(temp_upper, grade)
    k_lower = yield_lower / yield_lower[0]
    k_web = yield_web / yield_web[0]
    k_upper = yield_upper / yield_upper[0]

    strength_table = pd.DataFrame({
        "Время, мин": minutes,
        "Температура нижней полки, ℃": temp_lower,
        "Предел текучести нижней полки": yield_lower,
        "Коэффициент снижения предела текучести нижней полки": k_lower,
        "Температура стенки, ℃": temp_web,
        "Предел текучести стенки": yield_web,
        "Коэффициент снижения предела текучести стенки": k_web,
        "Температура верхней полки, ℃": temp_upper,
        "Предел текучести верхней полки": yield_upper,
        "Коэффициент снижения предела текучести верхней полки": k_upper,
    })

    # 3. Показатель сжатой зоны (положение нейтральной оси a)
    # При снижении прочности стенки до нуля возникает деление на ноль —
    # дальше расчёт теряет смысл, соответствующие значения остаются NaN/inf.
    with np.errstate(divide="ignore", invalid="ignore"):
        a_flange = (
            k_upper * b * t + k_web * s * h - 2 * k_web * s * t + k_lower * b * t
        ) / (2 * k_upper * b)
        a_web = (
            k_web * h * s - k_upper * t * b + k_lower * t * b
        ) / (2 * k_web * s)

        compressed_zone = pd.DataFrame({
            "Время, мин": minutes,
            "Показатель сжатой зоны при x < a": a_flange,
            "Показатель сжатой зоны при x > a": a_web,
        })

        # 4. Нормативное сопротивление, кгс/см²
        rn_lower = k_lower * yield_lower[0] * MPA_TO_KGF_CM2
        rn_web = k_web * yield_web[0] * MPA_TO_KGF_CM2
        rn_upper = k_upper * yield_upper[0] * MPA_TO_KGF_CM2

        normative_resistance = pd.DataFrame({
            "Время, мин": minutes,
            "Нормативное сопротивление нижней полки, кгс/см²": rn_lower,
            "Нормативное сопротивление стенки, кгс/см²": rn_web,
            "Нормативное сопротивление верхней полки, кгс/см²": rn_upper,
        })

        # 5. Усилия растяжения и сжатия.
        # Ветвление: нейтральная ось в стенке (a_flange > t) или в полке.
        in_web = a_flange > t

        tensile_lower = rn_lower * b * t * 0.01
        tensile_web = np.where(
            in_web,
            rn_web * (h * 0.1 - a_web * 0.1 - t * 0.1) * s * 0.1,
            rn_web * s * 0.1 * (h * 0.1 - 2 * t * 0.1),
        )
        # Когда НО в стенке (in_web=True):  сжатие верхней части стенки
        # Когда НО в полке (in_web=False):  растяжение нижней части верхней полки;
        #   материал — верхняя полка (rn_upper), не стенка!
        compression_web = np.where(
            in_web,
            rn_web * (a_web * 0.1 - t * 0.1) * s * 0.1,
            rn_upper * (t * 0.1 - a_flange * 0.1) * b * 0.1,
        )
        compression_upper = np.where(
            in_web,
            rn_upper * t * b * 0.01,
            rn_upper * a_flange * b * 0.01,
        )
        # Сумма нормальных усилий: при НО в полке нижняя часть верхней полки
        # работает на растяжение, поэтому знак compression_web меняется.
        total_efforts = np.where(
            in_web,
            compression_upper + compression_web - tensile_web - tensile_lower,
            compression_upper - compression_web - tensile_web - tensile_lower,
        )

        efforts = pd.DataFrame({
            "Время, мин": minutes,
            "Усилие растяжения в нижней полке": tensile_lower,
            "Усилие растяжения в нижней части стенки": tensile_web,
            "Усилие сжатия в верхней части стенки": compression_web,
            "Усилие сжатия в верхней полке": compression_upper,
            "Сумма всех нормальных усилий в сечении": total_efforts,
        })

        # 6. Плечи равнодействующих сил от нейтральной оси, мм
        # НО в стенке (in_web=True): плечи от a_web
        # НО в полке (in_web=False): плечи от a_flange; для нижней части верхней
        #   полки плечо = (t − a_flange)/2, для стенки — от h/2 до a_flange
        arm_tensile_lower = np.where(in_web,
            h - a_web - t / 2,
            h - t / 2 - a_flange)
        arm_tensile_web = np.where(in_web,
            (h - a_web - t) / 2,
            h / 2 - a_flange)
        arm_compression_web = np.where(in_web,
            (a_web - t) / 2,
            (t - a_flange) / 2)
        arm_compression_upper = np.where(in_web,
            a_web - t / 2,
            a_flange / 2)

        lever_arms = pd.DataFrame({
            "Время, мин": minutes,
            "Плечо равнодействующей силы растяжения в нижней полке, мм": arm_tensile_lower,
            "Плечо равнодействующей силы растяжения в нижней части стенки, мм": arm_tensile_web,
            "Плечо равнодействующей силы сжатия в верхней части стенки, мм": arm_compression_web,
            "Плечо равнодействующей силы сжатия в верхней полке, мм": arm_compression_upper,
        })

        # 7. Изгибающие моменты, кНм. Перевод кгс·мм -> кН·м: 1 кгс·мм = G·10⁻⁶ кН·м.
        _KGF_MM_TO_KNM = G * 1e-6
        moment_lower = tensile_lower * arm_tensile_lower * _KGF_MM_TO_KNM
        moment_web = tensile_web * arm_tensile_web * _KGF_MM_TO_KNM
        moment_upper_web = compression_web * arm_compression_web * _KGF_MM_TO_KNM
        moment_upper = compression_upper * arm_compression_upper * _KGF_MM_TO_KNM

        bending_moments = pd.DataFrame({
            "Время, мин": minutes,
            "Изгибающий момент в нижней полке, кНм": moment_lower,
            "Изгибающий момент в нижней части стенки, кНм": moment_web,
            "Изгибающий момент в верхней части стенки, кНм": moment_upper_web,
            "Изгибающий момент в верхней полке, кНм": moment_upper,
        })

        # 8. Несущая способность — сумма изгибающих моментов
        capacity = moment_lower + moment_web + moment_upper_web + moment_upper

    load_capacity = pd.DataFrame({
        "Время, мин": minutes,
        "Несущая способность, кНм": capacity,
    })

    # 9. Момент от нагрузки в середине пролёта с учётом собственного веса.
    # length_m — пролёт в МЕТРАХ.
    # M_нагр = P·L/4, кг·м → кНм: умножить на g/1000
    # M_своб = q·L²/8, кг·м → кНм: умножить на g/1000
    if m is None:
        raise ValueError("Ошибка: данные о массе 'M, кг' не найдены в данных профиля.")
    _KG_M_TO_KNM = G / 1000.0
    moment_value = (load_kg * length_m / 4) * _KG_M_TO_KNM + (m * length_m ** 2 / 8) * _KG_M_TO_KNM

    applied_moment = pd.DataFrame({
        "Время, мин": minutes,
        "Момент, кНм": moment_value * np.ones(len(minutes)),
    })

    # 10. Предел огнестойкости: последняя целая минута, когда несущая способность
    # ещё превышала момент. Если пересечение произошло между N и N+1 мин,
    # принимается меньшее (N). NaN считается исчерпанием несущей способности.
    holds = capacity > moment_value  # NaN -> False
    if holds.all():
        fire_limit = None
    elif not holds.any():
        fire_limit = 0
    else:
        fire_limit = int(minutes[holds][-1])

    return FireCalcResult(
        strength_table=strength_table,
        compressed_zone=compressed_zone,
        normative_resistance=normative_resistance,
        efforts=efforts,
        lever_arms=lever_arms,
        bending_moments=bending_moments,
        load_capacity=load_capacity,
        applied_moment=applied_moment,
        applied_moment_value=moment_value,
        fire_limit_minute=fire_limit,
    )
