from scripts.srinfo_sebrae_sourceamount import srinfo_sebrae_sourceamount
from scripts.srinfo_unit import srinfo_unit
from scripts.srinfo_company import srinfo_company_company
from scripts.buscar_arquivos_sharepoint import buscar_arquivos_sharepoint
from scripts.apagar_arquivos_pasta import apagar_arquivos_pasta
from scripts.gerar_planilha_geral import gerar_planilha_geral, gerar_planilha_erros
from scripts.gerar_planilhas_ufs import gerar_planilhas_uf
from scripts.connect_sharepoint import SharepointClient
import os
from dotenv import load_dotenv

load_dotenv()
ROOT = os.getenv('ROOT_SEBRAE_UFS')
STEP3 = os.path.abspath(os.path.join(ROOT, 'step_3_data_processed'))
STEP1 = os.path.abspath(os.path.join(ROOT, 'step_1_data_raw'))


def main():
    # Carregar arquivos do SharePoint
    print("Passo 1/4: Buscando arquivos do SharePoint")
    apagar_arquivos_pasta(STEP1)
    apagar_arquivos_pasta(STEP3)
    buscar_arquivos_sharepoint(gerar_novo=False)

    # Consulta databricks
    print("Passo 2/4: Consultando valores por fonte no Databricks")
    srinfo_sebrae_sourceamount()
    srinfo_unit()
    srinfo_company_company()
    sp = SharepointClient()
    # arquivos = ['srinfo_sebrae_sourceamount.xlsx', 'srinfo_unit.xlsx', 'srinfo_company_company.xlsx']
    # for arquivo in arquivos:
    #     sp.upload_file_to_folder(os.path.join(STEP3, arquivo), 'dw_pii')
    #     shutil.move(os.path.join(STEP3, arquivo), os.path.join(STEP1, arquivo))

    # # Gerando planilhas
    print("Passo 3/4: Gerando planilhas")
    planilha_geral, combinado, municipios, port_ue, proj_emp, port_emp, port_me = gerar_planilha_geral(gerar_novo=False, enviar_pasta_sebrae=True)
    gerar_planilha_erros(planilha_geral)
    gerar_planilhas_uf(planilha_geral, combinado, municipios, port_ue, proj_emp, port_emp, port_me, gerar_novo=False, enviar_pasta_sebrae=True)

    # Levando arquivos para o SharePoint
    print("Passo 4/4: Levando planilhas para o SharePoint")

    for nome_arquivo in os.listdir(STEP3):
            caminho_do_arquivo = os.path.join(STEP3, nome_arquivo)
            if os.path.isfile(caminho_do_arquivo):
                sp.upload_file_to_folder(caminho_do_arquivo, 'DWPII/sebrae_ufs')

if __name__ == "__main__":
    main()