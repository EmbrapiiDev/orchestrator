import os
from dotenv import load_dotenv
import inspect

# carregar .env
load_dotenv()
ROOT = os.getenv("ROOT_FORMS")

# Definição dos caminhos
STEP1 = os.path.abspath(os.path.join(ROOT, "step_1_data_raw"))
STEP2 = os.path.abspath(os.path.join(ROOT, "step_2_stage_area"))
STEP3 = os.path.abspath(os.path.join(ROOT, "step_3_data_processed"))
BACKUP = os.path.abspath(os.path.join(ROOT, "backup"))

SHAREPOINT_SITE = os.getenv("sharepoint_url_site")
SHAREPOINT_SITE_NAME = os.getenv("sharepoint_site_name")
SHAREPOINT_DOC = os.getenv("sharepoint_doc_library")


# from office365_api.download_files import get_file, get_files
from scripts.connect_sharepoint import SharepointClient


def buscar_arquivos_sharepoint():
    print("🟡 " + inspect.currentframe().f_code.co_name)

    sp = SharepointClient()
    sp.download_file(f"Formulários Sebrae/formularios.zip", os.path.join(STEP1, "formularios_anteriores.zip"))
    sp.download_file(f"Formulários Sebrae/respostas_formularios_sebrae.xlsx", os.path.join(STEP1, "respostas_formularios_sebrae_anterior.xlsx"))

    print("🟢 " + inspect.currentframe().f_code.co_name)

# Executar função
if __name__ == "__main__":
    buscar_arquivos_sharepoint()
