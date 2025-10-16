import os
import sys
from dotenv import load_dotenv

#carregar .env
load_dotenv()
ROOT = os.getenv('ROOT_SEBRAE_UFS')

#Definição dos caminhos
STEP1 = os.path.abspath(os.path.join(ROOT, 'step_1_data_raw'))
STEP2 = os.path.abspath(os.path.join(ROOT, 'step_2_stage_area'))
STEP3 = os.path.abspath(os.path.join(ROOT, 'step_3_data_processed'))
PATH_OFFICE = os.path.abspath(os.path.join(ROOT, 'office365_api'))
SCRIPTS = os.path.abspath(os.path.join(ROOT, 'scripts'))

# Verifica e cria a pasta se não existir
if not os.path.exists(STEP1):
    os.makedirs(STEP1)

# Verifica e cria a pasta se não existir
if not os.path.exists(STEP2):
    os.makedirs(STEP2)

# Verifica e cria a pasta se não existir
if not os.path.exists(STEP3):
    os.makedirs(STEP3)

# Adiciona o diretório correto ao sys.path
sys.path.append(SCRIPTS)
sys.path.append(PATH_OFFICE)

# from office365_api.download_files import get_file, get_files
from scripts.apagar_arquivos_pasta import apagar_arquivos_pasta
from scripts.connect_sharepoint import SharepointClient

def buscar_arquivos_sharepoint(gerar_novo = False):
    apagar_arquivos_pasta(STEP1)
    apagar_arquivos_pasta(STEP2)
    apagar_arquivos_pasta(STEP3)

    sp = SharepointClient()
    pasta_srinfo = "DWPII/srinfo"
    arquivos_srinfo = ['portfolio.xlsx', 'macroentregas.xlsx', 'projetos_empresas.xlsx', 'informacoes_empresas.xlsx',
                'empresas_contratantes.xlsx', 'pedidos_pi.xlsx', 'info_unidades_embrapii.xlsx']
    
    for arquivo in arquivos_srinfo:
        sp.download_file(f"{pasta_srinfo}/{arquivo}", os.path.join(STEP1, arquivo))

    pasta_lookup = "DWPII/lookup_tables"
    arquivos_lookup = ['ibge_municipios.xlsx', 'Contas Bancárias - UE.xlsx']
    
    for arquivo in arquivos_lookup:
        sp.download_file(f"{pasta_lookup}/{arquivo}", os.path.join(STEP1, arquivo))

    sp.download_file("dw_pii/oni_companies.xlsx", os.path.join(STEP1, "oni_companies.xlsx"))

    if gerar_novo == False:
        pasta_origem = "DWPII/sebrae_ufs"
        arquivos = sp.list_files(pasta_origem)
        for arquivo in arquivos:
            nome_arquivo = arquivo["name"]
            sp.download_file(f"{pasta_origem}/{nome_arquivo}", os.path.join(STEP1, nome_arquivo))