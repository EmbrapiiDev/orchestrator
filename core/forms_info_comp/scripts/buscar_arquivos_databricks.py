import os
import pandas as pd
from dotenv import load_dotenv
from databricks import sql

load_dotenv()
ROOT = os.getenv('ROOT_FORMS')
STEP1 = os.path.abspath(os.path.join(ROOT, 'step_1_data_raw'))
STEP2 = os.path.abspath(os.path.join(ROOT, 'step_2_stage_area'))
STEP3 = os.path.abspath(os.path.join(ROOT, 'step_3_data_processed'))
DATABRICKS_SERVER_HOSTNAME = os.getenv('DATABRICKS_SERVER_HOSTNAME')
DATABRICKS_HTTP_PATH = os.getenv('DATABRICKS_HTTP_PATH')
DATABRICKS_TOKEN = os.getenv('DATABRICKS_TOKEN')

connection = sql.connect(
    server_hostname=DATABRICKS_SERVER_HOSTNAME,
    http_path=DATABRICKS_HTTP_PATH,
    access_token=DATABRICKS_TOKEN
)


def srinfo_partnership_fundsapproval():
    query = """
            SELECT *
            FROM lhdw_srinfo.fd_srinfo_partnership_fundsapproval
            """

    df = pd.read_sql(query, connection)
    nome_arquivo = "partnership_fundsapproval"

    # Salvar em formato Excel
    path_file_processed = os.path.abspath(os.path.join(STEP1, f"srinfo_{nome_arquivo}.xlsx"))
    df.to_excel(path_file_processed, index=False)


def respostas_formularios_sebrae():
    query = """
            SELECT *
            FROM lhdw_srinfo.fd_srinfo_respostas_formularios_sebrae
            """

    df = pd.read_sql(query, connection)
    nome_arquivo = "respostas_formularios_sebrae_anterior"

    # Salvar em formato Excel
    path_file_processed = os.path.abspath(os.path.join(STEP1, f"{nome_arquivo}.xlsx"))
    df.to_excel(path_file_processed, index=False)