import os
import sys
import pandas as pd
from dotenv import load_dotenv
from scripts.query_clickhouse import query_clickhouse

load_dotenv()

ROOT = os.getenv('ROOT_SEBRAE_UFS')
sys.path.append(ROOT)

HOST = os.getenv('HOST')
PORT = os.getenv('PORT')
USER = os.getenv('USER')
PASSWORD = os.getenv('CLICKHOUSE_PASSWORD')
STEP1 = os.path.abspath(os.path.join(ROOT, 'step_1_data_raw'))
STEP3 = os.path.abspath(os.path.join(ROOT, 'step_3_data_processed'))


def srinfo_company_company():
    """
    Dados de empresas, sem inativação e sem CNPJ nulo
    """
    query = """
        SELECT DISTINCT
            company.*
        FROM db_bronze_srinfo.company_company AS company
        WHERE data_inativacao IS NULL
        AND NOT cnpj IS NULL
    """
    nome_arquivo = "company_company"
    query_clickhouse(HOST, PORT, USER, PASSWORD, query, nome_arquivo)

    # Carregar arquivo
    path_file_raw = os.path.abspath(os.path.join(ROOT, STEP1 ,f"{nome_arquivo}.csv"))
    df_raw = pd.read_csv(path_file_raw)

    # Carregar arquivo
    path_file_raw = os.path.abspath(os.path.join(ROOT, STEP1 ,f"{nome_arquivo}.csv"))
    df_raw = pd.read_csv(path_file_raw)

    # Salvar em formato Excel
    path_file_processed = os.path.abspath(os.path.join(ROOT, STEP3, f"srinfo_{nome_arquivo}.xlsx"))
    df_raw.to_excel(path_file_processed, index=False)
