---
name: run-steel
description: Run, start, smoke-test or screenshot the Steel fire-resistance Streamlit app (расчет предела огнестойкости стальных конструкций). Use when asked to run the app, verify a change works, or capture the UI.
---

# Run the Steel app

Streamlit web app (`main.py`) that computes the fire-resistance limit of steel
I-beams (огнестойкость стальных изгибаемых конструкций). All data comes from
three JSON files next to the scripts (`temperature.json`, `prochnost.json`,
`beam_profiles.json`). The agent path is the driver at
`.claude/skills/run-steel/driver.py` — it drives the app headlessly via
`streamlit.testing.v1.AppTest` (no browser) and can screenshot the real server
via Playwright + system Edge.

All paths below are relative to the project root (the directory containing
`main.py`). All commands were verified on Windows 10 / PowerShell.

## Which entry point

- **`main.py`** — the Streamlit UI (sidebar inputs, fire-limit metric, chart,
  tables in expanders). Refactored 2026-06-12 from the old monolithic
  `main1.py`.
- **`calc.py`** — the calculation core, no Streamlit dependency. For PRs that
  touch formulas, drive this directly:
  `from calc import SteelDatabase, compute` — see `tests/test_calc.py`.
- `xlm_to_json.py` — one-off Excel→JSON converter; still has hardcoded
  `C:\Users\kryge\Documents\...` paths, edit before running. Don't run it
  casually: it overwrites `beam_profiles.json`.

## Setup (once)

Requires `uv` (present on this machine) and Python 3.14 via the `py` launcher.
The `python` on PATH is an unrelated agent venv **without pip** — don't use it.

```powershell
uv venv .venv --python 3.14
uv pip install --python .venv\Scripts\python.exe streamlit pandas numpy matplotlib openpyxl playwright
```

No `playwright install` needed — the driver uses the system Edge
(`channel="msedge"`).

## Run (agent path) — the driver

Headless smoke: selects document → profile → steel grade, sets load/length,
re-runs, prints the load-capacity and moment tables. Exits 1 on `st.error` or
an uncaught exception.

```powershell
$env:PYTHONIOENCODING='utf-8'
.venv\Scripts\python.exe .claude\skills\run-steel\driver.py smoke main.py --load 2000 --length 6
# optional: --doc "ГОСТ 26020-83" --profile "20Б1" --grade "С345"
```

Expected tail: a `Несущая способность, кНм` table with decreasing values and
`SMOKE PASSED`.

Screenshot (boots `streamlit run` on a free port, captures via headless Edge,
kills the server):

```powershell
.venv\Scripts\python.exe .claude\skills\run-steel\driver.py screenshot main.py --out shot.png
# --height 2400 (default) controls how much of the page is captured
```

It prints the page `<h1>` title and the absolute screenshot path. Verify the
title is the Russian app title, not an error page.

## Run (human path)

```powershell
.venv\Scripts\python.exe -m streamlit run main.py
```

Opens a browser at http://localhost:8501. Ctrl-C to stop. Headless, this just
sits there — use the driver instead.

## Test

```powershell
$env:PYTHONIOENCODING='utf-8'
.venv\Scripts\python.exe tests\test_calc.py
```

Plain-assert tests of `calc.py` (no pytest needed): a golden reference case
(ГОСТ 26020-83 / 20Б1 / С345 / 2000 кг / 6 м → limit 25 min), monotonicity,
and a sweep over all ~500 profiles. Expect `ALL TESTS PASSED`.

## Gotchas

- **Hardcoded data paths (fixed 2026-06-12).** The app used absolute paths to
  non-existent directories; `calc.SteelDatabase.from_dir()` now resolves the
  JSONs relative to `calc.py`. If you see `Файл ... не найден`, someone
  reverted that.
- **NaN tail in capacity tables is expected.** At late time steps the
  interpolated yield strength reaches 0 and the neutral-axis division blows
  up; `calc.py` suppresses the numpy warnings and the fire-limit logic treats
  NaN as exhausted capacity.
- **Section-8 parenthesis bug (fixed 2026-06-12):** the old `main1.py`
  recomputed efforts three times; the second copy dropped parentheses in the
  a ≤ t branch, inflating capacity in the last minutes before the limit. The
  single implementation in `calc.py` uses the correct (section-7) formulas —
  `tests/test_calc.py` golden values reflect the fix.
- **Selectbox order matters to the driver:** 0 = document, 1 = beam profile,
  2 = steel grade (all in the sidebar). Profile options reload only after
  `at.run()` following a document change.
- **`AppTest` doesn't add the project root to `sys.path`** — `import calc`
  fails inside the tested script unless the driver inserts `ROOT` first (it
  does).
- **Tables sit inside collapsed `st.expander`s** — Playwright must wait with
  `state="attached"`, not the default `visible` (the driver does).
- **Console mojibake:** the app's labels are Russian; without
  `PYTHONIOENCODING=utf-8` the cp1251 console prints `???`. The driver also
  reconfigures stdout to UTF-8 itself.
- **Streamlit full-page screenshots don't work** — the app scrolls inside its
  own container, so `full_page=True` captures only the viewport. The driver
  uses a tall viewport (`--height`) instead.
- **Dependencies** are pinned loosely in `requirements.txt` (added
  2026-06-12); the Setup section above installs the same set via uv.

## Troubleshooting

- `No module named pip` → you used the PATH `python` (an agent venv). Use
  `.venv\Scripts\python.exe` or `uv pip install --python ...`.
- `streamlit server did not become healthy in 30s` → port collision or slow
  cold start; re-run, or pass `--port`.
- Playwright `Executable doesn't exist` for msedge → Edge missing (it's
  preinstalled on this Windows 10 box); fall back to
  `.venv\Scripts\python.exe -m playwright install chromium` and change
  `channel="msedge"` to default chromium in `driver.py`.
