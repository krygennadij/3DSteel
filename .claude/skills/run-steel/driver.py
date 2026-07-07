# -*- coding: utf-8 -*-
"""Driver for the Steel fire-resistance Streamlit app.

Run with the project venv interpreter (.venv/Scripts/python.exe):

    python .claude/skills/run-steel/driver.py smoke  [main.py] [options]
    python .claude/skills/run-steel/driver.py screenshot [main.py] [--out shot.png]

`smoke` drives the app headlessly via streamlit.testing.v1.AppTest (no browser,
no server): selects a profile document, a beam profile and a steel grade, fills
load/length, re-runs the script and dumps the resulting tables. Exit code 1 if
the app raised or showed st.error.

`screenshot` boots a real `streamlit run` server on a free port and captures a
full-page PNG through Playwright using the system Edge browser (channel=msedge,
nothing to download).
"""
import argparse
import os
import socket
import subprocess
import sys
import time
import urllib.request

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")

SKILL_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(SKILL_DIR, "..", "..", ".."))


def app_path(name: str) -> str:
    p = os.path.join(ROOT, name)
    if not os.path.exists(p):
        sys.exit(f"App file not found: {p}")
    return p


def report(at, label):
    print(f"--- {label} ---")
    if len(at.exception):
        for e in at.exception:
            print("EXCEPTION:", e.value)
        return False
    errs = [e.value for e in at.error]
    if errs:
        for e in errs:
            print("st.error:", e)
        return False
    print(f"OK: {len(at.dataframe)} dataframes, {len(at.selectbox)} selectboxes")
    return True


def smoke(args):
    from streamlit.testing.v1 import AppTest

    # AppTest runs the script in-process; make project-root imports (calc.py) work
    if ROOT not in sys.path:
        sys.path.insert(0, ROOT)

    at = AppTest.from_file(app_path(args.app), default_timeout=120)
    at.run()
    if not report(at, "initial run (defaults)"):
        sys.exit(1)

    # selectbox order in the app: 0 = document, 1 = beam profile, 2 = steel grade
    doc_options = at.selectbox[0].options
    grade_options = at.selectbox[2].options
    print("documents:", doc_options)
    print("steel grades:", grade_options)

    at.selectbox[0].select(args.doc or doc_options[0])
    at.run()
    profile_options = at.selectbox[1].options
    print(f"profiles in '{at.selectbox[0].value}': {len(profile_options)} "
          f"(first 10: {profile_options[:10]})")

    at.selectbox[1].select(args.profile or profile_options[0])
    at.selectbox[2].select(args.grade or grade_options[0])
    at.number_input[0].set_value(args.load)
    at.number_input[1].set_value(args.length)
    at.run()
    ok = report(at, f"run: doc={at.selectbox[0].value} profile={at.selectbox[1].value} "
                    f"grade={at.selectbox[2].value} load={args.load} length={args.length}")
    if not ok:
        sys.exit(1)

    try:
        for m in at.metric:
            print(f"metric: {m.label} = {m.value}")
    except Exception:
        pass

    # last two tables: load capacity over time, applied moment
    for df_el in at.dataframe[-2:]:
        df = df_el.value
        print()
        print(df.head(12).to_string(index=False))
    print("\nSMOKE PASSED")


def free_port() -> int:
    s = socket.socket()
    s.bind(("127.0.0.1", 0))
    port = s.getsockname()[1]
    s.close()
    return port


def screenshot(args):
    from playwright.sync_api import sync_playwright

    port = args.port or free_port()
    url = f"http://127.0.0.1:{port}"
    proc = subprocess.Popen(
        [sys.executable, "-m", "streamlit", "run", app_path(args.app),
         "--server.headless", "true", "--server.port", str(port),
         "--browser.gatherUsageStats", "false"],
        cwd=ROOT, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    try:
        for _ in range(60):
            try:
                if urllib.request.urlopen(f"{url}/_stcore/health", timeout=1).read() == b"ok":
                    break
            except OSError:
                time.sleep(0.5)
        else:
            sys.exit("streamlit server did not become healthy in 30s")
        print(f"server healthy at {url}")

        with sync_playwright() as p:
            browser = p.chromium.launch(channel="msedge", headless=True)
            # Streamlit scrolls inside its own container, so full_page only
            # captures the viewport — use a tall viewport instead.
            page = browser.new_page(viewport={"width": 1400, "height": args.height})
            page.goto(url, wait_until="networkidle")
            # Streamlit renders progressively; wait until tables exist.
            # state="attached": tables inside collapsed st.expander are
            # present in the DOM but not visible.
            page.wait_for_selector("[data-testid='stDataFrame']",
                                   state="attached", timeout=30000)
            page.wait_for_timeout(2000)
            out = os.path.abspath(args.out)
            page.screenshot(path=out, full_page=True)
            title = page.locator("h1").first.inner_text()
            browser.close()
        print("title:", title)
        print("screenshot:", out)
    finally:
        proc.kill()


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    sub = ap.add_subparsers(dest="cmd", required=True)

    sm = sub.add_parser("smoke", help="headless AppTest flow")
    sm.add_argument("app", nargs="?", default="main.py")
    sm.add_argument("--doc", help="profile document (default: first)")
    sm.add_argument("--profile", help="beam profile (default: first in doc)")
    sm.add_argument("--grade", help="steel grade (default: first)")
    sm.add_argument("--load", type=float, default=1000.0, help="load, kg")
    sm.add_argument("--length", type=float, default=6.0, help="span length, m")
    sm.set_defaults(func=smoke)

    sc = sub.add_parser("screenshot", help="run server + Edge screenshot")
    sc.add_argument("app", nargs="?", default="main.py")
    sc.add_argument("--out", default="app_screenshot.png")
    sc.add_argument("--port", type=int, default=None)
    sc.add_argument("--height", type=int, default=2400, help="viewport height px")
    sc.set_defaults(func=screenshot)

    args = ap.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
