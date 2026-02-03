import os
import pandas as pd
from dotenv import load_dotenv
from databricks import sql

load_dotenv()
ROOT = os.getenv('ROOT_SEBRAE_UFS')
STEP1 = os.path.abspath(os.path.join(ROOT, 'step_1_data_raw'))
STEP3 = os.path.abspath(os.path.join(ROOT, 'step_3_data_processed'))
DATABRICKS_SERVER_HOSTNAME = os.getenv('DATABRICKS_SERVER_HOSTNAME')
DATABRICKS_HTTP_PATH = os.getenv('DATABRICKS_HTTP_PATH')
DATABRICKS_TOKEN = os.getenv('DATABRICKS_TOKEN')

connection = sql.connect(
    server_hostname=DATABRICKS_SERVER_HOSTNAME,
    http_path=DATABRICKS_HTTP_PATH,
    access_token=DATABRICKS_TOKEN
)

def srinfo_sebrae_sourceamount():
    query = """
            SELECT *
            FROM lhdw_srinfo.fd_srinfo_sebrae_sourceamount
            """

    df = pd.read_sql(query, connection)
    nome_arquivo = "sebrae_sourceamount"

    # Salvar em formato Excel
    path_file_processed = os.path.abspath(os.path.join(ROOT, STEP1, f"srinfo_{nome_arquivo}.xlsx"))
    df.to_excel(path_file_processed, index=False)