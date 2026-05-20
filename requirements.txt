"""
generador/pdf_exporter.py — Convierte PPTX a PDF usando LibreOffice headless.
"""
import subprocess
from pathlib import Path
from config import LIBREOFFICE_PATH


def pptx_a_pdf(pptx_path: Path, pdf_dir: Path) -> Path:
    """
    Convierte un .pptx a .pdf con LibreOffice headless.
    Retorna la ruta del PDF generado.
    """
    pdf_dir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [
            LIBREOFFICE_PATH,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(pdf_dir),
            str(pptx_path),
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice falló para {pptx_path.name}:\n"
            f"STDOUT: {result.stdout}\nSTDERR: {result.stderr}"
        )
    pdf_path = pdf_dir / pptx_path.with_suffix(".pdf").name
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF no fue creado: {pdf_path}")
    return pdf_path
