import os
from dotenv import load_dotenv
from core.portfolio2.start_clean import start_clean
from core.portfolio2.connection.sharepoint import get_files_from_sharepoint, sharepoint_post
from core.portfolio2.processar_portfolio2 import processar_portfolio2
from core.portfolio2.connection.levar_postgresql import processar_portfolio2_para_postgres, levar_postgresql

load_dotenv()

ROOT = os.getenv("ROOT_PORTFOLIO2")
STEP_3_DATA_PROCESSED = os.getenv("STEP_3_DATA_PROCESSED")
PORTFOLIO2 = os.path.abspath(os.path.join(ROOT, STEP_3_DATA_PROCESSED, 'portfolio2.xlsx'))


def main_portfolio2():
    start_clean()
    get_files_from_sharepoint()
    processar_portfolio2()
    sharepoint_post()

    df = processar_portfolio2_para_postgres(PORTFOLIO2)
    levar_postgresql(
        df,
        nome_banco='db_srinfo_projeto',
        nome_schema='sc_srinfo_projeto',
        nome_tabela='tb_portfolio_2',
        truncate=True        
    )


if __name__ == "__main__":
    main_portfolio2()