from pathlib import Path

BASE_DIR  = Path(__file__).parent
DADOS_DIR = BASE_DIR / "dados"
MODELO_DIR = BASE_DIR / "modelo"
OUTPUT_DIR = BASE_DIR / "output"

for d in [DADOS_DIR, MODELO_DIR, OUTPUT_DIR]:
    d.mkdir(exist_ok=True)

CONTRATOS_FILE = DADOS_DIR / "contratos.xlsx"
MEDICOES_FILE  = DADOS_DIR / "medicoes.xlsx"
MODELO_FILE    = MODELO_DIR / "Modelo_medio.xlsx"
