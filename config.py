from pathlib import Path

# Diretórios
BASE_DIR = Path(__file__).parent
DADOS_DIR = BASE_DIR / "dados"
MODELO_DIR = BASE_DIR / "modelo"
OUTPUT_DIR = BASE_DIR / "output"

for d in [DADOS_DIR, MODELO_DIR, OUTPUT_DIR]:
    d.mkdir(exist_ok=True)

# Arquivos locais
CONTRATOS_FILE = DADOS_DIR / "contratos.xlsx"
MEDICOES_FILE  = DADOS_DIR / "medicoes.xlsx"
MODELO_FILE    = MODELO_DIR / "Modelo_medio.xlsx"

# SharePoint / Microsoft Graph
# Preencha com seus dados do Azure AD App Registration
TENANT_ID     = "SEU_TENANT_ID"
CLIENT_ID     = "SEU_CLIENT_ID"
CLIENT_SECRET = "SEU_CLIENT_SECRET"

# Caminho no SharePoint
SHAREPOINT_SITE_ID = "SEU_SITE_ID"   # ex: contoso.sharepoint.com,abc123,...
SHAREPOINT_DRIVE_ID = "SEU_DRIVE_ID" # ID da drive/biblioteca de documentos
# Caminho do arquivo de acompanhamento no SharePoint
SHAREPOINT_PLANILHA_PATH = "/Geral/Acompanhamento_Medicoes.xlsx"
