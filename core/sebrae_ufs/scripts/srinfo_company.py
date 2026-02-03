import os
import pandas as pd
from dotenv import load_dotenv
from databricks import sql

load_dotenv()
ROOT = os.getenv('ROOT_SEBRAE_UFS')
STEP1 = os.path.abspath(os.path.join(ROOT, 'step_1_data_raw'))
DATABRICKS_SERVER_HOSTNAME = os.getenv('DATABRICKS_SERVER_HOSTNAME')
DATABRICKS_HTTP_PATH = os.getenv('DATABRICKS_HTTP_PATH')
DATABRICKS_TOKEN = os.getenv('DATABRICKS_TOKEN')

connection = sql.connect(
    server_hostname=DATABRICKS_SERVER_HOSTNAME,
    http_path=DATABRICKS_HTTP_PATH,
    access_token=DATABRICKS_TOKEN
)


def srinfo_company_company():
    """
    Dados de empresas, sem inativação e sem CNPJ nulo
    """
    query = """
        SELECT DISTINCT company.*
        FROM lhdw_srinfo.fd_srinfo_empresa AS company
        WHERE NOT cnpj IS NULL
    """
    nome_arquivo = "company_company"
    df = pd.read_sql(query, connection)

    path_file_processed = os.path.abspath(os.path.join(ROOT, STEP1, f"srinfo_{nome_arquivo}.xlsx"))
    df.to_excel(path_file_processed, index=False)