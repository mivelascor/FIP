"""
generador/pdf_exporter.py — Convierte PPTX a PDF usando LibreOffice headless.
"""
import subprocess, shutil
from pathlib import Path


def _libreoffice_path() -> str:
    for cmd in ("libreoffice", "soffice", "/usr/bin/libreoffice", "/usr/bin/soffice"):
        if shutil.which(cmd):
            return cmd
    return "libreoffice"


def pptx_a_pdf(pptx_path: Path, pdf_dir: Path) -> Path:
    pdf_dir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [_libreoffice_path(), "--headless", "--convert-to", "pdf",
         "--outdir", str(pdf_dir), str(pptx_path)],
        capture_output=True, text=True, timeout=120,
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
