import os
from dotenv import load_dotenv
from datetime import datetime
import pandas as pd
import zipfile
import shutil

# carregar .env
load_dotenv()
ROOT = os.getenv("ROOT_FORMS")

STEP1= os.path.abspath(os.path.join(ROOT, "step_1_data_raw"))
STEP2 = os.path.abspath(os.path.join(ROOT, "step_2_stage_area"))
STEP3 = os.path.abspath(os.path.join(ROOT, "step_3_data_processed"))
BACKUP = os.path.abspath(os.path.join(ROOT, "backup"))

# Importar módulos necessários
from scripts.apagar_arquivos_pasta import apagar_arquivos_pasta
from scripts.buscar_arquivos_databricks import srinfo_partnership_fundsapproval, respostas_formularios_sebrae
from scripts.buscar_arquivos_sharepoint import buscar_arquivos_sharepoint
from scripts.webdriver import configurar_webdriver
from scripts.baixar_dados_srinfo import baixar_dados_srinfo
from scripts.ler_planilhas import consolidar_planilhas, ajustes, gerar_planilhas_ia
from scripts.ler_pdfs import ler_pdfs, ler_pdfs_ia
from scripts.gerar_planilha_final import gerar_planilha_final
from scripts.up_sharepoint import zipar_arquivos, up_sharepoint

def start_clean():
    # apagar arquivos das pastas
    apagar_arquivos_pasta(STEP1)
    apagar_arquivos_pasta(STEP2)
    apagar_arquivos_pasta(STEP3)

def mover_arquivos(pasta_origem, pasta_destino, extensoes = None):
    """
    Move todos os arquivos da pasta_origem para pasta_destino ou apenas os arquivos com as extensoes especificadas.
    Se extensoes for None, move todos os arquivos.
    """
    os.makedirs(pasta_destino, exist_ok=True)

    for item in os.listdir(pasta_origem):
        caminho_origem = os.path.join(pasta_origem, item)

        if extensoes is None:
            if os.path.isfile(caminho_origem):
                caminho_destino = os.path.join(pasta_destino, item)
                shutil.move(caminho_origem, caminho_destino)
        else:
            if item.lower().endswith(tuple(extensoes)):
                shutil.move(
                    os.path.join(pasta_origem, item),
                    os.path.join(pasta_destino, item)
                )

def main():
    # Passo 1: Limpando as pastas
    print(f"Iniciando processo em {datetime.now()}")
    print("Passo 1/8: Limpando pastas e movendo arquivos anteriores para backup...")
    start_clean()
    print("Pastas limpas.")


    # Passo 2: Coletando dados dos tickets e resultados anteriores
    print("Passo 2/8: Coletando dados dos tickets e resultados anteriores ...")
    srinfo_partnership_fundsapproval()
    respostas_formularios_sebrae()
    print("Dados coletados.")


    # Passo 3: Buscando formulários anteriores no SharePoint
    print("Passo 3/8: Buscando formulários anteriores no SharePoint...")
    buscar_arquivos_sharepoint()
    print("Arquivos buscados.")


    # Passo 4: Zipando e movendo arquivos antigos para backup
    print("Passo 4/8: Zipando e movendo arquivos antigos para backup...")
    nome_zip = f"info_comp_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
    zipar_arquivos(STEP1, os.path.join(BACKUP, nome_zip), zip_extra=os.path.join(STEP1, 'formularios_anteriores.zip'))
    print("Arquivos zipados e movidos para backup.")


    # Passo 5: Configurando WebDriver e baixando documentos do SRInfo
    print("Passo 5/8: Configurando WebDriver e baixando documentos do SRInfo...")
    driver = configurar_webdriver(STEP1)
    baixar_dados_srinfo(driver)
    print("Documentos baixados.")


    # Passo 6: Gerando planilhas iniciais
    print("Passo 6/8: Gerando planilhas iniciais...")
    df = consolidar_planilhas(STEP1)
    ajustes(df, STEP2)
    print("Planilhas iniciais geradas.")


    # Passo 6.1: Para as planilhas que tiveram algum erro, fazer a requisição para a IA
    # Verificar quais não possuem todos os dados em info_complementar
    info_complementar = pd.read_excel(os.path.join(STEP2, "info_complementares.xlsx"))
    contagem = (info_complementar.groupby("arquivo").size().reset_index(name="qtd_linhas"))
    arquivos_refazer = contagem.loc[contagem["qtd_linhas"] < 21, "arquivo"].tolist()
    if len(arquivos_refazer) > 0:
        print("Passo 6.1/8: Gerando planilhas com auxílio da IA...")
        print(f"{len(arquivos_refazer)} arquivos a refazer com IA: {arquivos_refazer}")
        gerar_planilhas_ia(STEP1, arquivos_refazer, STEP2)
        print("Planilhas com auxílio da IA geradas.")


    # Passo 6.2: Lendo PDFs (se houver)
    print("Passo 6.2/8: Lendo PDFs (se houver)...")
    arquivos_pdf = [f for f in os.listdir(STEP1) if f.endswith(('.pdf', '.PDF'))]
    if len(arquivos_pdf) != 0:
        print(f"PDFs em STEP1: {len(arquivos_pdf)}")
        textos_problematicos = ler_pdfs()

        # Passo 6.3: Para os PDFs que tiveram algum erro, fazer a requisição para a IA
        if textos_problematicos:
            print("Passo 6.3/8: Gerando planilhas com auxílio da IA...")
            print(f"{len(textos_problematicos)} textos problemáticos encontrados enviados para IA")
            ler_pdfs_ia(textos_problematicos, STEP2)

        print("PDFs lidos.")
    else:
        print("Nenhum PDF encontrado em STEP1.")


    # Passo 7: Juntando planilhas e gerando planilha final
    print("Passo 7/8: Juntando planilhas e gerando planilha final...")
    gerar_planilha_final()
    print("Planilha final gerada.")


    # Passo 8: Enviando planilha final para o SharePoint
    print("Passo 8/8: Enviando planilha final para o SharePoint...")
    up_sharepoint(nome_zip)
    print("Planilha final enviada.")
    print(f"Processo finalizado em {datetime.now()}")


if __name__ == "__main__":
    main()