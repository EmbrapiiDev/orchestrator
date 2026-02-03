from scripts.buscar_arquivos_sharepoint import buscar_arquivos_sharepoint
from scripts.gerar_planilhas import gerar_planilhas, processar_planilhas
from scripts.connect_sharepoint import SharepointClient
from scripts.zipar_arquivos import zipar_arquivos
from scripts.main_companies import main_companies
import os
from dotenv import load_dotenv

load_dotenv()
ROOT = os.getenv('ROOT_BFA')
STEP3 = os.path.abspath(os.path.join(ROOT, 'data', 'step_3_data_processed'))
BACKUP = os.path.abspath(os.path.join(ROOT, 'data', 'backup'))
COPY = os.path.abspath(os.path.join(ROOT, 'data', 'copy'))

def main():
    # Buscando arquivo do BFA no SharePoint
    print("Passo 1/5: Buscando arquivo do BFA no SharePoint...")
    buscar_arquivos_sharepoint()
    zipar_arquivos(COPY, BACKUP)

    # Gerando planilhas a partir do arquivo do BFA
    print("Passo 2/5: Gerando planilhas a partir do arquivo do BFA...")
    gerar_planilhas()

    # Processando as planilhas geradas
    print("Passo 3/5: Processando as planilhas geradas...")
    processar_planilhas()

    # # Pesquisar informações das empresas no databricks do oni
    # print("Passo 4/5: Buscando informações das empresas no Databricks do Observatório Nacional da Indústria...")
    # main_companies()

    # Levando as planilhas processadas para o SharePoint
    print("Passo 5/5: Enviando planilhas processadas para o SharePoint...")
    sp = SharepointClient()

    for nome_arquivo in os.listdir(STEP3):
        caminho_do_arquivo = os.path.join(STEP3, nome_arquivo)
        if os.path.isfile(caminho_do_arquivo):
            sp.upload_file_to_folder(caminho_do_arquivo, 'dw_pii')

    for nome_arquivo in os.listdir(BACKUP):
        caminho_do_arquivo = os.path.join(BACKUP, nome_arquivo)
        if os.path.isfile(caminho_do_arquivo):
            sp.upload_file_to_folder(caminho_do_arquivo, 'DWPII_backup')


if __name__ == "__main__":
    main()