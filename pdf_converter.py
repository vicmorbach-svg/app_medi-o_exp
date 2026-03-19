import subprocess
import platform
from pathlib import Path
from config import OUTPUT_DIR


def excel_para_pdf(excel_path: Path, aba: str) -> Path:
    """
    Converte uma aba específica do Excel para PDF.

    Estratégia:
    - Windows: usa xlwings + Excel COM (requer Excel instalado)
    - Linux/Mac (ex: servidor Streamlit Cloud): usa LibreOffice headless

    Retorna o Path do PDF gerado.
    """
    nome_pdf = excel_path.stem + f"_{aba}.pdf"
    pdf_path = OUTPUT_DIR / nome_pdf

    sistema = platform.system()

    if sistema == "Windows":
        _converter_windows(excel_path, aba, pdf_path)
    else:
        _converter_libreoffice(excel_path, pdf_path)

    return pdf_path


def _converter_windows(excel_path: Path, aba: str, pdf_path: Path):
    """Usa xlwings/Excel COM no Windows para exportar aba como PDF."""
    import xlwings as xw

    with xw.App(visible=False) as app:
        wb = app.books.open(str(excel_path.resolve()))
        sheet = wb.sheets[aba]
        sheet.activate()
        wb.api.ActiveSheet.ExportAsFixedFormat(
            Type=0,  # 0 = xlTypePDF
            Filename=str(pdf_path.resolve()),
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
        )
        wb.close()


def _converter_libreoffice(excel_path: Path, pdf_path: Path):
    """
    Usa LibreOffice headless para converter Excel em PDF.
    Converte o arquivo inteiro; para separar abas seria necessário
    pós-processamento com pypdf.
    Requer: sudo apt-get install libreoffice
    """
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "pdf",
        "--outdir", str(OUTPUT_DIR),
        str(excel_path.resolve())
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice erro: {result.stderr}")

    # LibreOffice gera o PDF com o mesmo nome do xlsx
    gerado = OUTPUT_DIR / (excel_path.stem + ".pdf")
    if gerado.exists():
        gerado.rename(pdf_path)


def gerar_pdfs_medicao(excel_path: Path) -> dict:
    """
    Gera PDF para PROTOCOLO e BOLETIM.
    Retorna dict com paths: {"PROTOCOLO": Path, "BOLETIM": Path}
    """
    pdf_protocolo = excel_para_pdf(excel_path, "PROTOCOLO")
    pdf_boletim   = excel_para_pdf(excel_path, "BOLETIM")
    return {
        "PROTOCOLO": pdf_protocolo,
        "BOLETIM":   pdf_boletim
    }
