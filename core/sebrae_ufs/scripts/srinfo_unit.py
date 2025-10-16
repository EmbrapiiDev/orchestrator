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

def srinfo_unit():
    query = """
        SELECT DISTINCT
            main.id AS proj_id,
            main.name,
            main.initials,
            main.ticket_project_name,
            main.action_plan_signature,
            main.inst_type,
            main.presentation_text,
            main.institutional_responsable,
            main.phone_responsable,
            main.cell_responsable,
            main.phone_inst_responsable,
            main.cell_inst_responsable,
            main.comercial_responsable,
            main.phone_comerc_responsable,
            main.cell_comerc_responsable,
            main.accreditation_status,
            main.uf_state,
            main.city,
            main.zip_code,
            main.address,
            main.latitude,
            main.longitude,
            main.website,
            main.instagram,
            main.linkedin,
            main.updated_at,
            main.data_carga,
        FROM db_bronze_srinfo.ue_unit AS main
        WHERE main.data_inativacao IS NULL
    """
    nome_arquivo = "unit"
    query_clickhouse(HOST, PORT, USER, PASSWORD, query, nome_arquivo)

    # Carregar arquivo
    path_file_raw = os.path.abspath(os.path.join(ROOT, STEP1 ,f"{nome_arquivo}.csv"))
    df_raw = pd.read_csv(path_file_raw)

    # Carregar arquivo
    path_file_raw = os.path.abspath(os.path.join(ROOT, STEP1 ,f"{nome_arquivo}.csv"))
    df_raw = pd.read_csv(path_file_raw)

    # Substituir "main." nos nomes das colunas
    df_raw.columns = [col.replace("main.", "") for col in df_raw.columns]

    # Salvar em formato Excel
    path_file_processed = os.path.abspath(os.path.join(ROOT, STEP3, f"srinfo_{nome_arquivo}.xlsx"))
    df_raw.to_excel(path_file_processed, index=False)