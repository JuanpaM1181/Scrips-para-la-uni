import subprocess
import shutil


def convert_to_pdf(docx_path: str) -> str:
    pdf_path = docx_path.rsplit(".", 1)[0] + ".pdf"
    out_dir = "/".join(docx_path.split("/")[:-1]) or "."

    if not shutil.which("libreoffice"):
        raise RuntimeError("libreoffice no encontrado. Instalalo con: sudo pacman -S libreoffice-fresh")

    result = subprocess.run(
        ["libreoffice", "--headless", "--convert-to", "pdf", docx_path, "--outdir", out_dir],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"Error al convertir a PDF: {result.stderr}")
    return pdf_path
