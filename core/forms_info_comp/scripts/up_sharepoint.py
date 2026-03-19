import os
import zipfile
import inspect
from dotenv import load_dotenv
from scripts.connect_sharepoint import SharepointClient 

#carregar .env
load_dotenv()
ROOT = os.getenv('ROOT_FORMS')

#Definição dos caminhos
STEP1 = os.path.abspath(os.path.join(ROOT, 'step_1_data_raw'))
STEP2 = os.path.abspath(os.path.join(ROOT, 'step_2_stage_area'))
STEP3 = os.path.abspath(os.path.join(ROOT, 'step_3_data_processed'))
BACKUP = os.path.abspath(os.path.join(ROOT, 'backup'))

def zipar_arquivos(pasta_origem, zip_saida, zip_extra=None):
    """
    pasta_origem: pasta cujos arquivos serão zipados
    zip_saida: caminho do zip final
    zip_extra: outro zip a ser incorporado
    """

    with zipfile.ZipFile(zip_saida, 'w', zipfile.ZIP_DEFLATED) as zip_out:

        # 1️⃣ Zipa todos os arquivos da pasta
        for root, _, files in os.walk(pasta_origem):
            for file in files:
                caminho_completo = os.path.join(root, file)

                arcname = os.path.relpath(
                    caminho_completo,
                    pasta_origem
                )

                zip_out.write(caminho_completo, arcname)

        # 2️⃣ Incorpora outro zip (sem subpasta)
        if zip_extra and os.path.exists(zip_extra):
            with zipfile.ZipFile(zip_extra, 'r') as zip_in:
                for info in zip_in.infolist():
                    with zip_in.open(info.filename) as f:
                        zip_out.writestr(info.filename, f.read())

def up_sharepoint(nome_zip):
    print("🟡 " + inspect.currentframe().f_code.co_name)

    sp = SharepointClient()
    try:
        zipar_arquivos(pasta_origem=STEP1, zip_saida= os.path.join(STEP3, 'formularios.zip'), zip_extra=os.path.join(STEP1, 'formularios_anteriores.zip'))
        sp.upload_file(os.path.join(STEP3, 'formularios.zip'), f"Formulários Sebrae/formularios.zip")
        sp.upload_file(os.path.join(STEP3, 'respostas_formularios_sebrae.xlsx'), f"Formulários Sebrae/respostas_formularios_sebrae.xlsx")
        sp.upload_file(os.path.join(BACKUP, nome_zip), f"Formulários Sebrae/backup/{nome_zip}")
    except Exception as e:
        print(f"🔴 Erro: {e}")


