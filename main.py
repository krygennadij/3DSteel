# -*- coding: utf-8 -*-
"""Интерфейс (Streamlit). Вся математика — в calc.py."""
import math
import os
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from calc import SteelDatabase, FireCalcResult, compute
from report import make_word_report, make_excel_report

_OGZ_EXCEL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "температура с ОГЗ.xlsx")

st.set_page_config(page_title="Огнестойкость стальных балок", layout="wide")


@st.cache_resource
def load_db() -> SteelDatabase:
    return SteelDatabase.from_dir()


@st.cache_resource
def load_ogz_data(mtime: float = 0.0) -> dict:
    """Загружает температуры с ОГЗ из Excel-файла.

    Возвращает словарь {имя_листа: {'lower': array, 'web': array, 'upper': array}}
    с поминутными значениями (0..150 мин включительно).
    Если файл не найден — пустой словарь.

    Аргумент mtime участвует в ключе кэша Streamlit: при изменении файла
    на диске кэш автоматически считается устаревшим и данные перечитываются.
    """
    if not os.path.exists(_OGZ_EXCEL):
        return {}
    try:
        xl = pd.ExcelFile(_OGZ_EXCEL)
    except Exception:
        return {}

    minutes = np.arange(0, 151, dtype=float)
    result = {}
    for sheet in xl.sheet_names:
        df = xl.parse(sheet, header=0)
        time_s = df.iloc[:, 0].astype(float).values
        upper  = df.iloc[:, 1].astype(float).values
        web    = df.iloc[:, 2].astype(float).values
        lower  = df.iloc[:, 3].astype(float).values
        time_min = time_s / 60.0
        result[sheet] = {
            "lower": np.interp(minutes, time_min, lower),
            "web":   np.interp(minutes, time_min, web),
            "upper": np.interp(minutes, time_min, upper),
        }
    return result


def _nice_dtick(max_value: float, target_ticks: int = 15):
    """Возвращает (шаг делений, верхняя граница оси), при которых и 0, и
    max_value гарантированно попадают на подписанное деление: шаг подбирается
    близким к target_ticks делений, а граница оси округляется вверх до
    ближайшего кратного шагу (если max_value само не делится нацело)."""
    if max_value <= 0:
        return 1, 1
    raw = max_value / target_ticks
    candidates = [1, 2, 5, 10, 15, 20, 25, 30, 50, 100, 150, 200, 250, 500, 1000]
    divisors = [c for c in candidates if c <= max_value and max_value % c == 0
                and 0.4 * raw <= c <= 2.5 * raw]
    if divisors:
        step = min(divisors, key=lambda c: abs(c - raw))
    else:
        step = next((c for c in candidates if c >= raw), candidates[-1])
    axis_max = math.ceil(max_value / step) * step
    return step, axis_max


def make_chart(res: FireCalcResult) -> go.Figure:
    t_arr = res.load_capacity["Время, мин"].to_numpy(dtype=float)
    cap   = res.load_capacity["Несущая способность, кНм"].to_numpy(dtype=float)
    mom   = res.applied_moment_value
    limit = res.fire_limit_minute

    finite_cap = cap[np.isfinite(cap)]
    y_max = float(finite_cap[0]) * 1.08 if len(finite_cap) > 0 else 100.0
    x_max = float(t_arr[-1]) if len(t_arr) > 0 else 1.0

    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=t_arr, y=cap,
        mode="lines",
        name="Несущая способность, кНм",
        line=dict(color="#e60000", width=3),
        connectgaps=False,
        hovertemplate="%{x:.0f} мин — %{y:.2f} кНм<extra>Несущая способность</extra>",
    ))

    fig.add_trace(go.Scatter(
        x=t_arr, y=np.full_like(t_arr, mom),
        mode="lines",
        name="Момент от нагрузки, кНм",
        line=dict(color="#0055cc", width=3, dash="dash"),
        hovertemplate="%{x:.0f} мин — %{y:.3f} кНм<extra>Момент от нагрузки</extra>",
    ))

    if limit is not None and limit > 0:
        y0 = cap[limit - 1] if np.isfinite(cap[limit - 1]) else mom + 1
        y1 = cap[limit] if limit < len(cap) and np.isfinite(cap[limit]) else mom - 1
        x_cross = float(limit - 1) + (y0 - mom) / (y0 - y1) if y0 != y1 else float(limit)

        # Тонкая вертикальная пунктирная линия — без маркеров
        fig.add_vline(x=x_cross, line=dict(color="black", width=1, dash="dot"))

        # Текстовая подпись у верхнего края линии
        fig.add_annotation(
            x=x_cross, y=y_max,
            xanchor="left", yanchor="top",
            text=f"t<sub>пред</sub> = {limit} мин",
            font=dict(size=11, color="black"),
            showarrow=False,
            bgcolor="white", borderpad=2,
        )

    GRID = dict(showgrid=True, gridwidth=1, gridcolor="rgba(0,0,0,0.12)",
                ticks="inside", ticklen=6,
                showline=True, linewidth=1.5, linecolor="black",
                mirror=True, tickfont=dict(size=13))

    x_dtick, x_axis_max = _nice_dtick(x_max)

    fig.update_layout(
        template="simple_white",
        xaxis=dict(title=dict(text="Время, мин", font=dict(size=14)),
                   dtick=x_dtick, range=[0, x_axis_max], **GRID),
        yaxis=dict(title=dict(text="Момент, кНм", font=dict(size=14)),
                   range=[0, y_max], **GRID),
        hovermode="x unified",
        legend=dict(
            orientation="h", yanchor="top", y=-0.18,
            xanchor="center", x=0.5, font=dict(size=13),
        ),
        font=dict(size=13),
        margin=dict(t=20, b=15, l=70, r=20),
        height=600,
    )
    return fig


def make_comparison_chart(scenarios: list) -> go.Figure:
    """scenarios: список (подпись, FireCalcResult, цвет). Момент от нагрузки один
    для всех (не зависит от температуры), рисуется одной общей линией."""
    fig = go.Figure()
    mom = scenarios[0][1].applied_moment_value if scenarios else 0.0
    x_max = 0.0
    y_max = 0.0

    for label, res, color in scenarios:
        t_arr = res.load_capacity["Время, мин"].to_numpy(dtype=float)
        cap   = res.load_capacity["Несущая способность, кНм"].to_numpy(dtype=float)
        finite_cap = cap[np.isfinite(cap)]
        if len(finite_cap) > 0:
            y_max = max(y_max, float(finite_cap[0]) * 1.08)
        if len(t_arr) > 0:
            x_max = max(x_max, float(t_arr[-1]))

        fig.add_trace(go.Scatter(
            x=t_arr, y=cap, mode="lines", name=label,
            line=dict(color=color, width=3), connectgaps=False,
            hovertemplate="%{x:.0f} мин — %{y:.2f} кНм<extra>" + label + "</extra>",
        ))

        limit = res.fire_limit_minute
        if limit is not None and limit > 0:
            y0 = cap[limit - 1] if np.isfinite(cap[limit - 1]) else mom + 1
            y1 = cap[limit] if limit < len(cap) and np.isfinite(cap[limit]) else mom - 1
            x_cross = float(limit - 1) + (y0 - mom) / (y0 - y1) if y0 != y1 else float(limit)
            fig.add_vline(x=x_cross, line=dict(color=color, width=1, dash="dot"))

    if x_max > 0:
        fig.add_trace(go.Scatter(
            x=[0, x_max], y=[mom, mom], mode="lines",
            name="Момент от нагрузки, кНм",
            line=dict(color="#0055cc", width=2, dash="dash"),
            hovertemplate="%{y:.3f} кНм<extra>Момент от нагрузки</extra>",
        ))

    GRID = dict(showgrid=True, gridwidth=1, gridcolor="rgba(0,0,0,0.12)",
                ticks="inside", ticklen=6,
                showline=True, linewidth=1.5, linecolor="black",
                mirror=True, tickfont=dict(size=13))

    x_dtick, x_axis_max = _nice_dtick(x_max) if x_max > 0 else (1, 1)

    fig.update_layout(
        template="simple_white",
        xaxis=dict(title=dict(text="Время, мин", font=dict(size=14)),
                   dtick=x_dtick, range=[0, x_axis_max], **GRID),
        yaxis=dict(title=dict(text="Момент, кНм", font=dict(size=14)),
                   range=[0, y_max if y_max > 0 else 100], **GRID),
        hovermode="x unified",
        legend=dict(
            orientation="h", yanchor="top", y=-0.18,
            xanchor="center", x=0.5, font=dict(size=13),
        ),
        font=dict(size=13),
        margin=dict(t=20, b=15, l=70, r=20),
        height=600,
    )
    return fig


_SIDE_ORDER = ["bottom", "right", "left", "top"]   # порядок trace'ов 1..4


def _path_pts(coords, n=60):
    """n равномерно распределённых точек вдоль ломаной (для line+markers trace)."""
    segs, total = [], 0.0
    for i in range(len(coords) - 1):
        a, b = coords[i], coords[i + 1]
        L = ((b[0] - a[0]) ** 2 + (b[1] - a[1]) ** 2) ** 0.5
        segs.append((a, b, L))
        total += L
    if total == 0 or n < 2:
        return [coords[0][0]] * n, [coords[0][1]] * n
    xs, ys, si, seg_start = [], [], 0, 0.0
    for k in range(n):
        target = total * k / (n - 1)
        while si < len(segs) - 1 and seg_start + segs[si][2] < target - 1e-9:
            seg_start += segs[si][2]
            si += 1
        a, b, L = segs[si]
        tt = max(0.0, min(1.0, (target - seg_start) / L)) if L > 0 else 0.0
        xs.append(a[0] + tt * (b[0] - a[0]))
        ys.append(a[1] + tt * (b[1] - a[1]))
    return xs, ys


def make_section_figure(b: float, h: float, t: float, s: float, label: str,
                         heated_sides=None) -> go.Figure:
    """Интерактивный Plotly-рисунок поперечного сечения двутавра.

    Trace 0 — заливка; Trace 1=bottom, 2=right, 3=left, 4=top (кликабельные).
    Оранжевые линии идут по реальному контуру сечения. Клик переключает сторону.
    """
    if heated_sides is None:
        heated_sides = {"bottom", "left", "right"}

    hh, hb, hs = h / 2, b / 2, s / 2
    ORANGE = "#ff8800"
    COLD   = "rgba(130,130,130,0.28)"
    N      = 60    # маркеров вдоль каждой стороны (для кликабельности)

    fig = go.Figure()

    # ── Контур двутавра (серая заливка) ──────────────────────────────────────
    px = [-hb,  hb,  hb,  hs,  hs,  hb,  hb, -hb, -hb, -hs, -hs, -hb, -hb]
    py = [ hh,  hh, hh-t, hh-t, -(hh-t), -(hh-t), -hh, -hh, -(hh-t), -(hh-t), hh-t, hh-t, hh]
    fig.add_trace(go.Scatter(
        x=px, y=py, fill="toself",
        fillcolor="rgba(200,200,200,0.9)",
        line=dict(color="#444", width=1.5),
        mode="lines", hoverinfo="skip", showlegend=False, name="_bg",
    ))

    # ── 4 стороны обогреваемого периметра — по реальному контуру сечения ─────
    # Правая сторона: верх → ступенька у стенки → низ (по внешнему контуру)
    right_path = [
        ( hb,  hh),
        ( hb,  hh - t),
        ( hs,  hh - t),
        ( hs, -(hh - t)),
        ( hb, -(hh - t)),
        ( hb, -hh),
    ]
    # Левая сторона — зеркально
    left_path = [
        (-hb, -hh),
        (-hb, -(hh - t)),
        (-hs, -(hh - t)),
        (-hs,  hh - t),
        (-hb,  hh - t),
        (-hb,  hh),
    ]
    SIDES = [
        ("bottom", [( hb, -hh), (-hb, -hh)], "Нижняя сторона"),
        ("right",  right_path,               "Правая сторона"),
        ("left",   left_path,                "Левая сторона"),
        ("top",    [(-hb,  hh), ( hb,  hh)], "Верхняя сторона"),
    ]
    for side_name, coords, side_lbl in SIDES:
        heated = side_name in heated_sides
        color  = ORANGE if heated else COLD
        lw     = 6.0    if heated else 1.5
        xs, ys = _path_pts(coords, N)
        state  = "🔥 Обогревается"       if heated else "❄ Не обогревается"
        action = "Нажмите, чтобы убрать" if heated else "Нажмите, чтобы включить"
        fig.add_trace(go.Scatter(
            x=xs, y=ys,
            mode="lines+markers",
            line=dict(color=color, width=lw),
            marker=dict(size=14, opacity=0.01, color=color, symbol="circle"),
            name=side_lbl,
            showlegend=False,
            hovertemplate=f"{state}<br><i>{action}</i><extra>{side_lbl}</extra>",
        ))

    # ── Размерные линии ───────────────────────────────────────────────────────
    padx = max(hb * 0.42, 18.0)   # компактнее → балка занимает б́ольшую долю фигуры
    pady = max(hh * 0.14, 14.0)
    DC   = "#444"
    LW   = 1.6
    TK   = max(min(hb, hh) * 0.06, 4.0)   # пропорционально сечению

    def seg(xa0, ya0, xa1, ya1, w=LW):
        fig.add_shape(type="line", x0=xa0, y0=ya0, x1=xa1, y1=ya1,
                      line=dict(color=DC, width=w), layer="above")

    def tick45(x, y):
        """Засечка 45° (нижний-левый → верхний-правый)."""
        fig.add_shape(type="line", x0=x - TK, y0=y - TK, x1=x + TK, y1=y + TK,
                      line=dict(color=DC, width=LW), layer="above")

    def ann(x, y, text, angle=0):
        fig.add_annotation(x=x, y=y, text=text, showarrow=False,
                           font=dict(size=12, color="#111"), textangle=angle,
                           bgcolor="rgba(255,255,255,0.92)", borderpad=2.5)

    # b — над верхней полкой
    by = hh + pady
    seg(-hb, hh, -hb, by + TK)
    seg( hb, hh,  hb, by + TK)
    seg(-hb, by,  hb, by)
    tick45(-hb, by)
    tick45( hb, by)
    ann(0, by + pady * 0.65, f"b = {b:.0f}")

    # h — слева
    lx = -(hb + padx * 0.80)
    seg(-hb, -hh, lx - TK, -hh)
    seg(-hb,  hh, lx - TK,  hh)
    seg(lx, -hh, lx, hh)
    tick45(lx, -hh)
    tick45(lx,  hh)
    ann(lx - padx * 0.25, 0, f"h = {h:.0f}", angle=90)

    # t — справа, верхняя полка
    rx = hb + padx * 0.58
    seg(hb, hh,   rx + TK, hh)
    seg(hb, hh-t, rx + TK, hh-t)
    seg(rx, hh-t, rx, hh)
    tick45(rx, hh)
    tick45(rx, hh - t)
    ann(rx + max(padx * 0.55, 14), hh - t / 2, f"t = {t:.1f}")

    # s — толщина стенки: размерная линия в стенке, подпись — правее профиля
    seg(-hs, 0, hs, 0)
    tick45(-hs, 0)
    tick45( hs, 0)
    seg(hs, 0, hb + padx * 0.05, 0)          # выносная линия вправо через пустое пространство
    ann(hb + padx * 0.22, 0, f"s = {s:.0f}")

    # ── Макет ─────────────────────────────────────────────────────────────────
    xr_l = hb + padx * 1.6
    xr_r = hb + padx * 1.25
    yr_t = hh + pady * 2.5
    yr_b = hh + pady * 1.2

    fig.update_layout(
        height=600,
        template="simple_white",
        title=dict(text=label, x=0.5, font=dict(size=11, color="#111")),
        xaxis=dict(scaleanchor="y", scaleratio=1,
                   showgrid=False, zeroline=False, showticklabels=False,
                   showline=False, range=[-(xr_l), xr_r]),
        yaxis=dict(showgrid=False, zeroline=False, showticklabels=False,
                   showline=False, range=[-(yr_b), yr_t]),
        margin=dict(t=35, b=5, l=5, r=5),
        paper_bgcolor="white",
        plot_bgcolor="white",
        dragmode="pan",
        hoverdistance=20,
    )
    return fig


def calc_geometry(b, h, t, s, m_kgm, A_cm2_table=None, R_mm=0.0):
    """Геометрические параметры I-балки для проектирования огнезащиты.

    Площадь — из сортамента (А, см²) или, если не задана, через погонную массу.

    Каждая боковая сторона двутавра имеет 2 внутренних закруглённых угла
    (стык стенки с полкой). Острый угол заменяется дугой πR/2 вместо двух
    прямых отрезков длиной R, поправка на угол: R(π/2 − 2) ≈ −0.43R.
    Для всех схем, включающих боковые стороны, действуют 4 таких угла.
    """
    # Площадь
    if A_cm2_table is not None:
        A_cm2 = float(A_cm2_table)
    else:
        A_cm2 = m_kgm / 0.785      # плотность 7850 кг/м³ ↔ 0.785 кг/(м·см²)
    A_mm2 = A_cm2 * 100

    # Поправка на скругления: 4 угла (по 2 на каждую боковую сторону)
    R = float(R_mm) if R_mm else 0.0
    fillet_4 = 4 * R * (math.pi / 2 - 2)   # отрицательная (~−1.72R)

    p_bot = b
    p_top = b
    # Периметр двух боковых без учёта скруглений (острые углы)
    p_2sides_sharp = 2 * (h + b - s)

    cases = [
        ("1 — нижняя сторона",               p_bot),
        ("2 — обе боковых стороны",           p_2sides_sharp + fillet_4),
        ("3 — нижняя + обе боковых",          p_bot + p_2sides_sharp + fillet_4),
        ("4 — все стороны",                   p_bot + p_top + p_2sides_sharp + fillet_4),
    ]
    rows = []
    for label, u in cases:
        u_m = u / 1000
        rows.append({
            "Схема обогрева":                label,
            "Обогреваемый периметр, мм":     round(u, 1),
            "Площадь пов-сти / 1 м,  м²/м": round(u_m, 4),
            "Площадь пов-сти / т,  м²/т":   round(u_m * 1000 / m_kgm, 2),
            "Привед. толщина δ,  мм":        round(A_mm2 / u, 2),
        })
    return A_mm2, A_cm2, pd.DataFrame(rows)


_TEMP_COLS = ["Время (мин)", "Нижняя полка (°C)", "Стенка (°C)", "Верхняя полка (°C)"]


def _default_temp_df(db: SteelDatabase) -> pd.DataFrame:
    """Таблица температур из встроенных данных — используется как шаблон."""
    t_lower = db.get_temperature_data("lower_flange")
    t_web   = db.get_temperature_data("web")
    t_upper = db.get_temperature_data("upper_flange")
    return pd.DataFrame({
        "Время (мин)":         np.arange(len(t_lower), dtype=int),
        "Нижняя полка (°C)":  t_lower.values.astype(float),
        "Стенка (°C)":        t_web.values.astype(float),
        "Верхняя полка (°C)": t_upper.values.astype(float),
    })


def _parse_temp_csv(uploaded_file) -> "pd.DataFrame | str":
    """Парсит загруженный CSV. Возвращает DataFrame или строку с ошибкой."""
    try:
        df = pd.read_csv(uploaded_file)
    except Exception as exc:
        return f"Не удалось прочитать файл: {exc}"
    missing = [c for c in _TEMP_COLS[1:] if c not in df.columns]
    if missing:
        return f"Отсутствуют столбцы: {', '.join(missing)}"
    if len(df) < 2:
        return "Файл должен содержать минимум 2 строки."
    try:
        for col in _TEMP_COLS[1:]:
            df[col] = df[col].astype(float)
    except ValueError as exc:
        return f"Ошибка в данных: {exc}"
    return df


def calc_step(num: int, title: str, explanation: str, df: pd.DataFrame) -> None:
    with st.expander(f"Шаг {num}. {title}"):
        st.markdown(explanation)
        st.dataframe(df.round(3), hide_index=True, width="stretch")


def main():
    # Session state: стороны обогреваемого периметра
    if "heated_sides" not in st.session_state:
        st.session_state.heated_sides = {"bottom", "left", "right"}

    try:
        db = load_db()
    except (FileNotFoundError, ValueError) as e:
        st.error(str(e))
        return

    # ── Боковая панель ───────────────────────────────────────────────────────
    with st.sidebar:
        st.title("Исходные данные")

        st.markdown("##### 1. Профиль конструкции")
        _DEFAULT_DOC     = "СТО АСЧМ 20-93"
        _DEFAULT_PROFILE = "20Б1"
        _DEFAULT_GRADE   = "С345"

        profile_source = st.radio(
            "Источник сечения",
            ["Из сортамента", "Задать вручную"],
            horizontal=True,
            help="Выбор готового двутавра по ГОСТ/СТО или ввод произвольных размеров сечения",
        )

        custom_profile = None
        if profile_source == "Из сортамента":
            doc_idx = db.profile_docs.index(_DEFAULT_DOC) if _DEFAULT_DOC in db.profile_docs else 0
            selected_doc = st.selectbox(
                "Нормативный документ", db.profile_docs, index=doc_idx,
                help="ГОСТ или СТО, по которому выбирается сортамент двутавров",
            )
            profile_keys = db.get_profile_data(selected_doc).index.tolist()
            prof_idx = profile_keys.index(_DEFAULT_PROFILE) if _DEFAULT_PROFILE in profile_keys else 0
            selected_profile = st.selectbox(
                "Профиль двутавра", profile_keys, index=prof_idx,
                help="Обозначение профиля по выбранному документу",
            )
        else:
            st.caption("Размеры двутаврового сечения")
            h_custom = st.number_input(
                "Высота сечения h, мм", min_value=10.0, value=300.0, step=1.0, format="%.1f",
            )
            b_custom = st.number_input(
                "Ширина полки b, мм", min_value=10.0, value=150.0, step=1.0, format="%.1f",
            )
            t_custom = st.number_input(
                "Толщина полки t, мм", min_value=1.0, value=10.2, step=0.1, format="%.1f",
            )
            s_custom = st.number_input(
                "Толщина стенки s, мм", min_value=1.0, value=6.5, step=0.1, format="%.1f",
            )
            _area_est_mm2 = 2 * b_custom * t_custom + max(h_custom - 2 * t_custom, 0.0) * s_custom
            _m_est = _area_est_mm2 / 100.0 * 0.785  # площадь, см² × плотность стали 7850 кг/м³
            m_custom = st.number_input(
                "Погонная масса M, кг/м", min_value=0.1, value=round(_m_est, 1), step=0.1, format="%.2f",
                help="По умолчанию оценена по площади сечения (сталь, 7850 кг/м³) без учёта скруглений; можно скорректировать.",
            )
            selected_doc = "Пользовательское сечение"
            selected_profile = st.text_input(
                "Обозначение сечения",
                value=f"{h_custom:.0f}x{b_custom:.0f}x{t_custom:.1f}/{s_custom:.1f}",
                help="Произвольное название для отчёта",
            )
            custom_profile = {
                "b, мм": b_custom, "t, мм": t_custom, "s, мм": s_custom,
                "h, мм": h_custom, "M, кг": m_custom,
            }

        st.markdown("##### 2. Материал")
        grade_idx = db.steel_grades.index(_DEFAULT_GRADE) if _DEFAULT_GRADE in db.steel_grades else 0
        selected_grade = st.selectbox(
            "Марка стали", db.steel_grades, index=grade_idx,
            help="Определяет нормативный предел текучести и его снижение при нагреве",
        )

        st.markdown("##### 3. Нагрузка и геометрия")
        load_value = st.number_input(
            "Нагрузка P, кг",
            min_value=0.0, value=86865.0, step=100.0, format="%.1f",
            help="Нормативная сосредоточенная нагрузка в середине пролёта",
        )
        length_value = st.number_input(
            "Длина пролёта L, м",
            min_value=0.0, value=3.2, step=0.5, format="%.2f",
            help="Расчётный пролёт балки",
        )

        inputs_ok = load_value > 0 and length_value > 0
        if not inputs_ok:
            st.info("Задайте нагрузку и длину — иначе учитывается только собственный вес балки.")

        st.markdown("##### 4. Температуры сечения")
        _ogz_mtime = os.path.getmtime(_OGZ_EXCEL) if os.path.exists(_OGZ_EXCEL) else 0.0
        _ogz_data = load_ogz_data(_ogz_mtime)
        _ogz_options = [f"ОГЗ: {s}" for s in _ogz_data.keys()]
        _temp_source_options = ["Встроенные данные (без ОГЗ)"] + _ogz_options + ["Загрузить CSV", "Ввести вручную"]
        temp_source = st.radio(
            "Источник температур",
            _temp_source_options,
            label_visibility="collapsed",
        )

        _temp_lower_ext = _temp_web_ext = _temp_upper_ext = None

        # Загрузка данных из Excel (ОГЗ)
        if temp_source.startswith("ОГЗ: "):
            _sheet = temp_source[len("ОГЗ: "):]
            _ogz = _ogz_data[_sheet]
            _temp_lower_ext = _ogz["lower"]
            _temp_web_ext   = _ogz["web"]
            _temp_upper_ext = _ogz["upper"]
            st.caption(f"Загружено {len(_temp_lower_ext)} мин (0–{len(_temp_lower_ext)-1} мин)")

        if temp_source == "Загрузить CSV":
            tmpl_df = _default_temp_df(db)
            st.download_button(
                "Скачать шаблон CSV",
                tmpl_df.to_csv(index=False),
                file_name="temperature_template.csv",
                mime="text/csv",
                help="Заполните значения температур и загрузите файл обратно",
                use_container_width=True,
            )
            uploaded_temp = st.file_uploader(
                "CSV с температурами", type=["csv"],
                label_visibility="collapsed",
            )
            if uploaded_temp is not None:
                parsed = _parse_temp_csv(uploaded_temp)
                if isinstance(parsed, str):
                    st.error(parsed)
                else:
                    _temp_lower_ext = parsed["Нижняя полка (°C)"].values
                    _temp_web_ext   = parsed["Стенка (°C)"].values
                    _temp_upper_ext = parsed["Верхняя полка (°C)"].values
                    st.success(f"Загружено {len(parsed)} строк (0–{len(parsed)-1} мин)")

    # ── Расчёт ──────────────────────────────────────────────────────────────

    # Редактируемая таблица температур — вне sidebar, чтобы уместить все столбцы
    if temp_source == "Ввести вручную":
        with st.expander("Таблица температур сечения (редактировать)", expanded=True):
            st.caption(
                "Введите температуру каждой зоны сечения для каждой минуты пожара. "
                "При изменении значений расчёт пересчитывается автоматически."
            )
            edited_temp_df = st.data_editor(
                _default_temp_df(db),
                key="manual_temp_table",
                use_container_width=True,
                height=400,
                hide_index=True,
                column_config={
                    "Время (мин)": st.column_config.NumberColumn(
                        "Время (мин)", disabled=True, format="%d",
                    ),
                    "Нижняя полка (°C)": st.column_config.NumberColumn(
                        "Нижняя полка (°C)", min_value=0.0, max_value=1300.0, format="%.2f",
                    ),
                    "Стенка (°C)": st.column_config.NumberColumn(
                        "Стенка (°C)", min_value=0.0, max_value=1300.0, format="%.2f",
                    ),
                    "Верхняя полка (°C)": st.column_config.NumberColumn(
                        "Верхняя полка (°C)", min_value=0.0, max_value=1300.0, format="%.2f",
                    ),
                },
            )
            _temp_lower_ext = edited_temp_df["Нижняя полка (°C)"].values
            _temp_web_ext   = edited_temp_df["Стенка (°C)"].values
            _temp_upper_ext = edited_temp_df["Верхняя полка (°C)"].values

    try:
        res = compute(db, selected_doc, selected_profile, selected_grade,
                      load_value, length_value,
                      temp_lower_ext=_temp_lower_ext,
                      temp_web_ext=_temp_web_ext,
                      temp_upper_ext=_temp_upper_ext,
                      custom_profile=custom_profile)
    except ValueError as e:
        st.error(str(e))
        return

    if custom_profile is not None:
        b_dim = custom_profile["b, мм"]
        h_dim = custom_profile["h, мм"]
        t_dim = custom_profile["t, мм"]
        s_dim = custom_profile["s, мм"]
        m_dim = custom_profile["M, кг"]
        A_dim = None   # точной площади из сортамента нет — оценивается по погонной массе
        R_dim = None   # радиус скругления неизвестен
    else:
        profile_row = db.get_profile_data(selected_doc).loc[selected_profile].squeeze()
        b_dim = profile_row.get("b, мм")
        h_dim = profile_row.get("h, мм")
        t_dim = profile_row.get("t, мм")
        s_dim = profile_row.get("s, мм")
        m_dim = profile_row.get("M, кг")
        A_dim = profile_row.get("А, см2")   # точная площадь из сортамента
        R_dim = profile_row.get("R, мм")    # внутренний радиус скругления
    has_dims = all(v is not None for v in (b_dim, h_dim, t_dim, s_dim))

    limit = res.fire_limit_minute
    cap0  = res.load_capacity["Несущая способность, кНм"].iloc[0]

    # ── Заголовок ────────────────────────────────────────────────────────────
    st.title("Предел огнестойкости стальной балки")
    st.caption(
        "Расчёт с учётом неравномерного прогрева двутаврового сечения. "
        "Параметры — в боковой панели слева."
    )

    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 Несущая способность", "📐 Параметры огнезащиты",
        "📈 Сравнение ОГЗ", "📄 Отчёт",
    ])

    with tab1:
        # ── Баннер с результатом ─────────────────────────────────────────────
        if limit is None:
            color    = "#2ecc71"
            headline = "✅ Предел огнестойкости не достигнут в диапазоне 0–30 мин"
            detail   = "Несущая способность превышает момент от нагрузки на всём расчётном периоде."
        elif limit == 0:
            color    = "#e74c3c"
            headline = "❌ Несущей способности не хватает уже без нагрева"
            detail   = "Увеличьте сечение профиля или уменьшите нагрузку."
        else:
            color    = "#2ecc71" if limit >= 60 else ("#f39c12" if limit >= 30 else "#e74c3c")
            headline = f"🔥 Предел огнестойкости: **{limit} мин**"
            detail   = (
                f"На {limit}-й минуте пожара несущая способность балки снижается до уровня "
                f"момента от нагрузки ({res.applied_moment_value:.3f} кНм)."
            )

        st.markdown(
            f"""<div style="background:{color}22;border-left:5px solid {color};
            padding:14px 20px;border-radius:6px;margin-bottom:16px">
            <div style="font-size:1.2em">{headline}</div>
            <div style="color:#555;margin-top:4px">{detail}</div>
            </div>""",
            unsafe_allow_html=True,
        )

        # ── Метрики ──────────────────────────────────────────────────────────
        c1, c2, c3 = st.columns(3)
        c1.metric(
            "Предел огнестойкости",
            f"{limit} мин" if (limit is not None and limit > 0) else ("> 30 мин" if limit is None else "< 1 мин"),
            help="Минута пожара, когда несущая способность опускается ниже момента от нагрузки",
        )
        c2.metric(
            "Несущая способность при t = 0",
            f"{cap0:.1f} кНм",
            help="До начала пожарного воздействия",
        )
        c3.metric(
            "Момент от нагрузки",
            f"{res.applied_moment_value:.3f} кНм" if inputs_ok else "—",
            help="M = PL/4 + qL²/8, где q — собственный вес балки",
        )

        # ── График и сечение рядом ───────────────────────────────────────────
        col_chart, col_sec = st.columns([2, 1])

        with col_chart:
            st.plotly_chart(make_chart(res), width="stretch")
            st.caption(
                "Сплошная линия — несущая способность балки при нагреве. "
                "Пунктирная линия — момент от нагрузки (постоянен). "
                "Вертикальная линия — предел огнестойкости."
            )

        with col_sec:
            if has_dims:
                fig_sec = make_section_figure(
                    b_dim, h_dim, t_dim, s_dim,
                    selected_profile, st.session_state.heated_sides,
                )
                event = st.plotly_chart(
                    fig_sec,
                    on_select="rerun",
                    selection_mode="points",
                    key="section_fig",
                    use_container_width=True,
                )
                # Обработка клика по стороне
                sel_pts = (event.selection.points
                           if (event and hasattr(event, "selection") and event.selection)
                           else [])
                if sel_pts:
                    pt = sel_pts[0]
                    if isinstance(pt, dict):
                        cn = pt.get("curve_number", -1)
                    else:
                        cn = getattr(pt, "curve_number", -1)
                    if 1 <= cn <= 4:
                        side = _SIDE_ORDER[cn - 1]
                        if side in st.session_state.heated_sides:
                            st.session_state.heated_sides.discard(side)
                        else:
                            st.session_state.heated_sides.add(side)
                    # Сброс состояния виджета — при следующем рендере клик не повторится
                    if "section_fig" in st.session_state:
                        del st.session_state["section_fig"]
                    st.rerun()
                cnt = len(st.session_state.heated_sides)
                st.caption(
                    f"Обогревается {cnt} из 4 сторон  ·  "
                    "Нажмите на контур сечения для переключения"
                )
            else:
                st.info("Размеры сечения недоступны.")

        # ── Методика расчёта ─────────────────────────────────────────────────
        st.divider()
        st.subheader("Как выполнен расчёт")
        st.markdown(
            """
            Расчёт ведётся пошагово для каждой минуты (0–30 мин):
            температуры трёх участков сечения → коэффициенты снижения прочности →
            нормативное сопротивление → положение нейтральной оси →
            усилия → изгибающие моменты → **несущая способность**.
            Разверните нужный шаг, чтобы увидеть промежуточные значения.
            """
        )

        calc_step(1, "Температуры и коэффициенты снижения предела текучести",
            """
            По температурному режиму пожара определяется температура трёх участков сечения —
            **нижняя полка** (красная на схеме, горячее всего), **стенка** (оранжевая) и
            **верхняя полка** (синяя, холоднее всего). По кривой «температура → предел текучести»
            вычисляется коэффициент снижения **k = σ(T) / σ(T₀)**: при нагреве до 800 °C сталь
            теряет почти всю прочность.
            """, res.strength_table)

        calc_step(2, "Положение нейтральной оси (показатель сжатой зоны)",
            """
            Из условия равновесия нормальных сил в сечении находится расстояние **a**
            от верхней грани до нейтральной оси. При нагреве снизу нейтральная ось
            постепенно смещается вниз — нижняя полка слабеет и уже не может нести
            прежнее усилие растяжения.
            """, res.compressed_zone)

        calc_step(3, "Нормативное сопротивление стали, кгс/см²",
            """
            **Rₙ = k · σ(T₀) · 10,197** — предел текучести при данной температуре,
            переведённый в кгс/см². Чем горячее участок, тем меньше Rₙ.
            """, res.normative_resistance)

        calc_step(4, "Усилия растяжения и сжатия",
            """
            Нейтральная ось делит сечение на сжатую (выше) и растянутую (ниже) зоны.
            Усилие в каждой части = Rₙ × площадь. Сумма нормальных усилий должна быть ≈ 0.
            """, res.efforts)

        calc_step(5, "Плечи равнодействующих сил, мм",
            """
            Плечо = расстояние от центра тяжести части сечения до нейтральной оси.
            Изгибающий момент от части = Усилие × Плечо.
            """, res.lever_arms)

        calc_step(6, "Изгибающие моменты от усилий, кНм",
            """
            Каждая часть сечения создаёт свой изгибающий момент. Их сумма — несущая способность.
            """, res.bending_moments)

        calc_step(7, "Несущая способность и момент от нагрузки",
            """
            **Несущая способность** = сумма изгибающих моментов всех частей сечения.
            **Предел огнестойкости** — первая минута, когда она опускается ниже момента от нагрузки.
            """,
            pd.concat([
                res.load_capacity.set_index("Время, мин"),
                res.applied_moment.set_index("Время, мин"),
            ], axis=1).reset_index())

    # ── Вкладка 2: геометрические параметры ─────────────────────────────────
    with tab2:
        if has_dims and m_dim is not None:
            A_mm2, A_cm2, geom_df = calc_geometry(
                b_dim, h_dim, t_dim, s_dim, m_dim,
                A_cm2_table=A_dim, R_mm=R_dim,
            )

            st.markdown(
                f"**{selected_doc} / {selected_profile}** · "
                f"b = {b_dim:.0f} мм,  h = {h_dim:.0f} мм,  "
                f"t = {t_dim:.1f} мм,  s = {s_dim:.1f} мм,  "
                f"M = {m_dim:.1f} кг/м"
            )

            ca, cb = st.columns(2)
            area_src = "Из сортамента (А, см²)" if A_dim is not None else "Через погонную массу: M / 7850 · 10⁶"
            ca.metric("Площадь сечения A", f"{A_cm2:.2f} см²", help=area_src)
            cb.metric("Масса погонная M", f"{m_dim:.1f} кг/м")

            st.divider()
            st.subheader("Обогреваемый периметр и производные величины")
            st.dataframe(geom_df, hide_index=True, use_container_width=True)

            # Подсветка текущей схемы обогрева
            n_h = len(st.session_state.heated_sides)
            if 1 <= n_h <= 4:
                labels = ["сторона", "стороны", "стороны", "стороны"]
                row = geom_df.iloc[n_h - 1]
                st.info(
                    f"**Текущая схема ({n_h} {labels[n_h - 1]}, вкладка «Несущая способность»):**  "
                    f"периметр = **{row['Обогреваемый периметр, мм']:.0f} мм** · "
                    f"пов-сть / 1 м = **{row['Площадь пов-сти / 1 м,  м²/м']:.4f} м²/м** · "
                    f"пов-сть / т = **{row['Площадь пов-сти / т,  м²/т']:.2f} м²/т** · "
                    f"δ = **{row['Привед. толщина δ,  мм']:.2f} мм**"
                )

            r_note = (f"Учтены скругления R = {R_dim:.0f} мм: поправка на 4 угла = "
                      f"{4 * float(R_dim) * (math.pi / 2 - 2):.1f} мм.  "
                      if R_dim else "")
            st.caption(
                "Периметр одной боковой стороны (без скруглений): p = h + b − s, включая "
                "внутренние поверхности полок и стенки.  "
                + r_note +
                "Приведённая толщина металла δ = A / u_обогр."
            )
        else:
            st.warning("Геометрические данные или погонная масса недоступны для этого профиля.")

    # ── Вкладка 3: Сравнение вариантов ОГЗ ───────────────────────────────────
    with tab3:
        st.subheader("Несущая способность: без ОГЗ и с ОГЗ (ГКЛ)")
        st.caption(
            "Три варианта огнезащиты для текущего профиля, стали, нагрузки и пролёта "
            "(см. боковую панель): без защиты (встроенные данные), 1 слой ГКЛ и "
            "2 слоя ГКЛ (по данным листов «температура с ОГЗ.xlsx»)."
        )

        _cmp_specs = [
            ("Без ОГЗ",           None,          "#e60000"),
            ("ОГЗ: 1 слой ГКЛ",   "1 слой ГКЛ",  "#f39c12"),
            ("ОГЗ: 2 слой ГКЛ",   "2 слой ГКЛ",  "#2ecc71"),
        ]
        _cmp_scenarios = []
        _cmp_missing = []
        for _label, _sheet_key, _color in _cmp_specs:
            try:
                if _sheet_key is None:
                    _res_cmp = compute(db, selected_doc, selected_profile, selected_grade,
                                       load_value, length_value, custom_profile=custom_profile)
                elif _sheet_key in _ogz_data:
                    _d = _ogz_data[_sheet_key]
                    _res_cmp = compute(db, selected_doc, selected_profile, selected_grade,
                                       load_value, length_value,
                                       temp_lower_ext=_d["lower"], temp_web_ext=_d["web"],
                                       temp_upper_ext=_d["upper"], custom_profile=custom_profile)
                else:
                    _cmp_missing.append(_label)
                    continue
            except ValueError as e:
                st.error(f"{_label}: {e}")
                continue
            _cmp_scenarios.append((_label, _res_cmp, _color))

        if _cmp_missing:
            st.warning(
                "Не найдены данные ОГЗ для: " + ", ".join(_cmp_missing) +
                " (файл «температура с ОГЗ.xlsx» отсутствует или неполон)."
            )

        if _cmp_scenarios:
            st.plotly_chart(make_comparison_chart(_cmp_scenarios), width="stretch")

            _cmp_rows = []
            for _label, _res_cmp, _ in _cmp_scenarios:
                _lim = _res_cmp.fire_limit_minute
                if _lim is None:
                    _lim_str = f"> {int(_res_cmp.load_capacity['Время, мин'].iloc[-1])} мин"
                elif _lim == 0:
                    _lim_str = "< 1 мин"
                else:
                    _lim_str = f"{_lim} мин"
                _cmp_rows.append({
                    "Вариант":                            _label,
                    "Предел огнестойкости":                _lim_str,
                    "Несущая способность при t = 0, кНм":  round(
                        _res_cmp.load_capacity["Несущая способность, кНм"].iloc[0], 1),
                })
            st.dataframe(pd.DataFrame(_cmp_rows), hide_index=True, use_container_width=True)
        else:
            st.info("Нет доступных данных для сравнения.")

    # ── Вкладка 4: Отчёт ─────────────────────────────────────────────────────
    with tab4:
        st.subheader("Выгрузка отчёта")
        st.markdown(
            "**Word (.docx)** — подробный текстовый отчёт со схемами, графиком "
            "и всеми расчётными таблицами.  \n"
            "**Excel (.xlsx)** — все таблицы расчёта и нативный Excel-график "
            "несущей способности (редактируемый)."
        )

        _geom_df_report = None
        if has_dims and m_dim is not None:
            _, _, _geom_df_report = calc_geometry(
                b_dim, h_dim, t_dim, s_dim, m_dim,
                A_cm2_table=A_dim, R_mm=R_dim,
            )

        col_w, col_x = st.columns(2)

        with col_w:
            st.markdown("#### Word-отчёт")
            if st.button("Сформировать Word", use_container_width=True):
                with st.spinner("Генерация..."):
                    docx_bytes = make_word_report(
                        res=res,
                        doc_name=selected_doc,
                        profile_key=selected_profile,
                        grade=selected_grade,
                        load_kg=load_value,
                        length_m=length_value,
                        temp_source=temp_source,
                        b_mm=b_dim if has_dims else None,
                        h_mm=h_dim if has_dims else None,
                        t_mm=t_dim if has_dims else None,
                        s_mm=s_dim if has_dims else None,
                        m_kgm=m_dim,
                        heated_sides=st.session_state.heated_sides,
                        geom_df=_geom_df_report,
                    )
                st.download_button(
                    label="⬇ Скачать Word (.docx)",
                    data=docx_bytes,
                    file_name=f"огнестойкость_{selected_profile}_{selected_grade}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )

        with col_x:
            st.markdown("#### Excel-отчёт")
            if st.button("Сформировать Excel", use_container_width=True):
                with st.spinner("Генерация..."):
                    xlsx_bytes = make_excel_report(
                        res=res,
                        doc_name=selected_doc,
                        profile_key=selected_profile,
                        grade=selected_grade,
                        load_kg=load_value,
                        length_m=length_value,
                        temp_source=temp_source,
                        b_mm=b_dim if has_dims else None,
                        h_mm=h_dim if has_dims else None,
                        t_mm=t_dim if has_dims else None,
                        s_mm=s_dim if has_dims else None,
                        m_kgm=m_dim,
                        geom_df=_geom_df_report,
                        comparison=_cmp_scenarios if _cmp_scenarios else None,
                    )
                st.download_button(
                    label="⬇ Скачать Excel (.xlsx)",
                    data=xlsx_bytes,
                    file_name=f"огнестойкость_{selected_profile}_{selected_grade}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )


if __name__ == "__main__":
    main()
