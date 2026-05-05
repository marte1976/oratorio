from __future__ import annotations

import importlib.util
from pathlib import Path
import shutil
import subprocess
import sys


ROOT = Path(__file__).resolve().parent.parent
APP_FILE = ROOT / "app.py"
OUTPUT_DIR = ROOT / "static" / "guide-media" / "images" / "full_course"
LOCAL_BASE_URL = "http://127.0.0.1:8000"
WORK_YEAR = "2026"
TEMP_HTML_DIR = ROOT / "outputs" / "video-tutorial-html"
HEADLESS_PROFILE_DIR = ROOT / "outputs" / "video-tutorial-browser-profile"
CHROME_CANDIDATES = [
    Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
    Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
    Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
    Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
]


SLIDES = [
    {"order": "01", "slug": "dashboard", "function": "dashboard_page"},
    {"order": "02", "slug": "nuovo-associato", "function": "associati_page"},
    {"order": "03", "slug": "anagrafica-storica-associati", "function": "associati_page", "query": {"vista": "dati"}},
    {"order": "04", "slug": "rinnovo-tesseramento", "function": "tesseramenti_page"},
    {"order": "05", "slug": "tesserati", "function": "tesserati_page"},
    {"order": "06", "slug": "dettaglio-tesserato", "function": "associato_report_page", "detail_associato": True},
    {"order": "07", "slug": "consiglio-direttivo", "function": "consiglio_direttivo_page"},
    {"order": "08", "slug": "corsi", "function": "corsi_page"},
    {"order": "09", "slug": "oratorio", "function": "oratorio_page"},
    {"order": "10", "slug": "campo-estivo", "function": "campi_estivi_page"},
    {"order": "11", "slug": "eventi", "function": "eventi_page"},
    {"order": "12", "slug": "pagamenti", "function": "pagamenti_multi_area_page"},
    {"order": "13", "slug": "report-posizione-tesserati", "function": "report_associati"},
    {"order": "14", "slug": "report-incassi", "function": "report_incassi"},
    {"order": "15", "slug": "backup", "function": "backup_page"},
    {"order": "16", "slug": "importa-associati", "function": "importa_associati_page"},
    {"order": "17", "slug": "aggiornamenti", "function": "aggiornamenti_page"},
]


def load_app_module():
    spec = importlib.util.spec_from_file_location("oratorio_app", APP_FILE)
    if spec is None or spec.loader is None:
        raise RuntimeError("Impossibile caricare app.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules["oratorio_app"] = module
    spec.loader.exec_module(module)
    return module


def resolve_browser() -> Path:
    for candidate in CHROME_CANDIDATES:
        if candidate.exists():
            return candidate
    raise RuntimeError("Browser headless non trovato.")


def absolutize_html(html: str) -> str:
    replacements = {
        'href="/': f'href="{LOCAL_BASE_URL}/',
        "href='/": f"href='{LOCAL_BASE_URL}/",
        'src="/': f'src="{LOCAL_BASE_URL}/',
        "src='/": f"src='{LOCAL_BASE_URL}/",
        'action="/': f'action="{LOCAL_BASE_URL}/',
        "action='/": f"action='{LOCAL_BASE_URL}/",
    }
    for source, target in replacements.items():
        html = html.replace(source, target)
    return html


def inject_slide_overrides(html: str, slide: dict[str, object]) -> str:
    css_chunks: list[str] = []
    slug = str(slide.get("slug") or "")

    if slug == "consiglio-direttivo":
        css_chunks.append(
            """
            .direttivo-member-header h3 {
              color: transparent !important;
              text-shadow: none !important;
              filter: blur(14px);
              background: rgba(239, 127, 26, 0.12);
              border-radius: 14px;
              width: fit-content;
              min-width: 260px;
              min-height: 28px;
              user-select: none;
            }
            .direttivo-member-inline-item span:last-child {
              color: transparent !important;
              text-shadow: none !important;
              filter: blur(12px);
              background: rgba(239, 127, 26, 0.14);
              border-radius: 999px;
              padding: 2px 12px;
              user-select: none;
            }
            """
        )

    if slug in {"corsi", "campo-estivo", "eventi"}:
        css_chunks.append(
            """
            .cards-stack > .card:nth-child(n+2) {
              display: none !important;
            }
            """
        )

    if slug in {"corsi", "oratorio", "campo-estivo", "eventi", "pagamenti", "report-posizione-tesserati", "report-incassi", "backup", "importa-associati", "aggiornamenti"}:
        css_chunks.append(
            """
            body {
              zoom: 0.94;
            }
            """
        )

    if not css_chunks:
        return html

    override_block = "<style>" + "\n".join(css_chunks) + "</style>"
    if "</head>" in html:
        return html.replace("</head>", override_block + "</head>", 1)
    return override_block + html


def screenshot_html(browser: Path, html_file: Path, output_file: Path, *, window_size: str = "1800,1450") -> None:
    output_file.parent.mkdir(parents=True, exist_ok=True)
    HEADLESS_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    command = [
        str(browser),
        "--headless=new",
        "--disable-gpu",
        "--hide-scrollbars",
        "--no-first-run",
        "--disable-crash-reporter",
        "--run-all-compositor-stages-before-draw",
        f"--user-data-dir={HEADLESS_PROFILE_DIR}",
        "--force-device-scale-factor=1",
        f"--window-size={window_size}",
        "--virtual-time-budget=5000",
        f"--screenshot={output_file}",
        html_file.resolve().as_uri(),
    ]
    subprocess.run(command, check=True, capture_output=True, text=True)


def main() -> None:
    module = load_app_module()
    browser = resolve_browser()
    base_query = {"anno": WORK_YEAR}
    current_user = {"username": "administrator", "is_admin": True}
    sample_associato_row = module.fetch_one(
        """
        SELECT a.id
        FROM associati a
        LEFT JOIN tesseramenti_annuali t ON t.associato_id = a.id AND t.anno_sociale = ?
        ORDER BY CASE WHEN t.id IS NULL THEN 1 ELSE 0 END, t.id, a.id
        LIMIT 1
        """,
        (int(WORK_YEAR),),
    )
    sample_associato_id = int(sample_associato_row["id"]) if sample_associato_row else None
    if TEMP_HTML_DIR.exists():
        shutil.rmtree(TEMP_HTML_DIR, ignore_errors=True)
    TEMP_HTML_DIR.mkdir(parents=True, exist_ok=True)
    if OUTPUT_DIR.exists():
        shutil.rmtree(OUTPUT_DIR, ignore_errors=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    for slide in SLIDES:
        page_fn = getattr(module, str(slide["function"]))
        slide_query = dict(base_query)
        slide_query.update(slide.get("query", {}))
        if slide.get("detail_associato"):
            if sample_associato_id is None:
                continue
            raw_html = page_fn(sample_associato_id, slide_query, current_user)
        else:
            raw_html = page_fn(slide_query, current_user)
        html = absolutize_html(raw_html.decode("utf-8"))
        html = inject_slide_overrides(html, slide)
        html_path = TEMP_HTML_DIR / f"{slide['order']}-{slide['slug']}.html"
        html_path.write_text(html, encoding="utf-8")
        screenshot_path = OUTPUT_DIR / f"{slide['order']}.png"
        window_size = "1800,1450"
        if str(slide.get("slug") or "") in {
            "corsi",
            "oratorio",
            "campo-estivo",
            "eventi",
            "pagamenti",
            "report-posizione-tesserati",
            "report-incassi",
            "backup",
            "importa-associati",
            "aggiornamenti",
        }:
            window_size = "1800,1820"
        if str(slide.get("slug") or "") == "pagamenti":
            window_size = "1800,1980"
        screenshot_html(browser, html_path, screenshot_path, window_size=window_size)


if __name__ == "__main__":
    main()
