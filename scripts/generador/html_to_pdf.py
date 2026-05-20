"""
generador/html_to_pdf.py — Convierte HTML → PDF usando Playwright headless.

Playwright renderiza el HTML tal como Chrome, respetando:
  - CSS @page (A4, sin márgenes)
  - print-color-adjust: exact (fondos negros)
  - page-break-after: always (2 páginas por folleto)
"""
import subprocess, shutil
from pathlib import Path


def html_a_pdf(html_path: Path, pdf_dir: Path) -> Path:
    """Convierte un HTML a PDF con Playwright. Retorna la ruta del PDF generado."""
    pdf_dir.mkdir(parents=True, exist_ok=True)
    pdf_path = pdf_dir / (html_path.stem + ".pdf")

    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
            page    = browser.new_page()
            page.goto(f"file://{html_path.resolve()}", wait_until="networkidle", timeout=30000)
            page.emulate_media(media="print")
            page.pdf(
                path=str(pdf_path),
                format="A4",
                margin={"top": "0", "bottom": "0", "left": "0", "right": "0"},
                print_background=True,
            )
            browser.close()
        return pdf_path

    except Exception as e:
        raise RuntimeError(f"Playwright falló para {html_path.name}: {e}")


def html_a_pdf_batch(html_paths: list, pdf_dir: Path) -> list:
    """Convierte una lista de HTMLs a PDFs en batch. Más eficiente: abre el browser una sola vez."""
    pdf_dir.mkdir(parents=True, exist_ok=True)
    results = []

    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
            for html_path in html_paths:
                try:
                    pdf_path = pdf_dir / (html_path.stem + ".pdf")
                    page = browser.new_page()
                    page.goto(f"file://{html_path.resolve()}", wait_until="networkidle", timeout=30000)
                    page.emulate_media(media="print")
                    page.pdf(
                        path=str(pdf_path),
                        format="A4",
                        margin={"top":"0","bottom":"0","left":"0","right":"0"},
                        print_background=True,
                    )
                    page.close()
                    results.append(pdf_path)
                    print(f"    [PDF] {pdf_path.name}")
                except Exception as e:
                    print(f"    [ERR] {html_path.name}: {e}")
            browser.close()
    except Exception as e:
        raise RuntimeError(f"Playwright no disponible: {e}")

    return results
