import os
from dotenv import load_dotenv
import pandas as pd
import psycopg2
from io import StringIO
import inspect

load_dotenv()

POSTGRES_HOST = os.getenv('POSTGRES_HOST', default='10.0.93.100')
POSTGRES_PORT = os.getenv('POSTGRES_PORT', default='5432')
POSTGRES_USER = os.getenv('POSTGRES_USER')
POSTGRES_PASS = os.getenv('POSTGRES_PASS')

def processar_portfolio2_para_postgres(caminho_excel):
    """
    Lê o Excel do portfolio2, ajusta tipos de dados e strings, e retorna um DataFrame pronto para PostgreSQL.
    """
    print("🟡 " + inspect.currentframe().f_code.co_name)
    try:
        # Ler Excel
        df = pd.read_excel(caminho_excel)
        
        # TODAS as colunas esperadas pela tabela PostgreSQL (na ordem correta, exceto co_seq_projeto que é auto)
        colunas_esperadas = [
            'id_codigo_projeto', 'rede_projeto_codigo', 'rede_projeto_papel_ue', 
            'st_status_projeto', 'negociacao_id', 'negociacao_negociacao', 
            'dt_data_negociacao', 'dt_data_projeto_contrato', 'dt_data_projeto_inicio', 
            'dt_data_projeto_termino', 'desc_projeto', 'desc_titulo', 
            'desc_titulo_publico', 'desc_descricao_publica', 'desc_objetivo', 
            'ue_unidade_embrapii', 'ue_status', 'ue_uf', 'ue_tipo_instituicao', 
            'ue_competencias_tecnicas', 'emp_empresas', 'emp_n_empresas', 
            'fin_parceria_programa', 'fin_call', 'fin_modalidade_financiamento', 
            'fin_uso_recurso', 'fin_parceiro_id', 'fin_parceiro', 'fin_contrato', 
            'fin_contrato_eixo', 'fin_contrato_eixo_regra', 'fin_macrocontrato', 
            'fin_sebrae', 'fin_sebrae_contrato', 'fin_sebrae_contrato_eixo', 
            'fin_sebrae_eixo_regra', 'fin_sebrae_eixo_regra2', 'fin_sebrae_macrocontrato', 
            'class_tipo_projeto', 'class_trl_inicial', 'class_trl_final', 
            'class_area_aplicacao', 'class_tecnologia_habilitadora', 'class_nib', 
            'class_bmaisp', 'class_energia_renovavel', 'class_economia_circular', 
            'class_tecnologia_verde', 'val_nominal_embrapii', 'val_nominal_empresa', 
            'val_nominal_sebrae', 'val_nominal_unidade', 'val_nominal_total', 
            'val_ipca_embrapii', 'val_ipca_empresa', 'val_ipca_sebrae', 
            'val_ipca_total', 'val_ipca_unidade', 'macroentrega_numero', 
            'macroentrega_aceites', 'macroentrega_macroentregas', 'pi_numero_pedidos', 
            'aval_nota', 'aval_data_avaliacao', 'outros_cooperacao_internacional', 
            'outros_observacoes', 'outros_tags', 'outros_data_extracao'
        ]
        
        # Identificar colunas faltantes
        colunas_faltantes = [col for col in colunas_esperadas if col not in df.columns]
        if colunas_faltantes:
            print(f"⚠️  Colunas faltantes ({len(colunas_faltantes)}): {colunas_faltantes}")
            # Adiciona colunas faltantes com None
            for col in colunas_faltantes:
                df[col] = None
        
        # Identificar colunas extras (que estão no Excel mas não na tabela)
        colunas_extras = [col for col in df.columns if col not in colunas_esperadas]
        if colunas_extras:
            print(f"⚠️  Colunas extras no Excel (serão removidas): {colunas_extras}")
            df = df.drop(columns=colunas_extras)
        
        # Reordenar colunas para corresponder à ordem da tabela
        df = df[colunas_esperadas]
        
        # Verificar e remover duplicatas na chave primária
        total_antes = len(df)
        duplicatas = df[df.duplicated(subset=['id_codigo_projeto'], keep=False)]
        
        if len(duplicatas) > 0:            
            # Remove duplicatas, mantendo a PRIMEIRA ocorrência
            df = df.drop_duplicates(subset=['id_codigo_projeto'], keep='first')

        # Ajustar strings: remover espaços desnecessários e garantir None em células vazias
        # IMPORTANTE: Substituir tabs, quebras de linha e outros caracteres problemáticos
        for col in df.select_dtypes(include='object').columns:
            df[col] = df[col].apply(lambda x: 
                x.strip()
                .replace('\t', ' ')      # Substitui tabs por espaço
                .replace('\n', ' ')      # Substitui quebras de linha por espaço
                .replace('\r', ' ')      # Substitui retorno de carro por espaço
                if isinstance(x, str) 
                else None if pd.isna(x) 
                else x
            )

        # Ajustar datas: garantir que todas sejam datetime ou None
        colunas_data = [
            'dt_data_negociacao', 'dt_data_projeto_contrato', 'dt_data_projeto_inicio',
            'dt_data_projeto_termino', 'aval_data_avaliacao', 'outros_data_extracao'
        ]
        for col in colunas_data:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

        # Ajustar floats: garantir apenas duas casas decimais para notas e percentuais
        float_cols_2dec = [
            'macroentrega_aceites', 'aval_nota',
            'val_nominal_embrapii', 'val_nominal_empresa', 'val_nominal_sebrae',
            'val_nominal_unidade', 'val_nominal_total',
            'val_ipca_embrapii', 'val_ipca_empresa', 'val_ipca_sebrae',
            'val_ipca_total', 'val_ipca_unidade'
        ]
        for col in float_cols_2dec:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').round(2)

        # Ajustar inteiros: preencher NaN com 0 ou None se necessário
        int_cols = ['emp_n_empresas', 'macroentrega_numero', 'pi_numero_pedidos', 'fin_parceiro_id']
        for col in int_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

        print("🟢 " + inspect.currentframe().f_code.co_name)
        return df
    
    except Exception as e:
        print(f"🔴 Erro: {e}")
        raise

def levar_postgresql(df, nome_banco, nome_schema, nome_tabela, truncate=True):
    """
    Função geral para enviar qualquer DataFrame para PostgreSQL rapidamente usando COPY.
    Se truncate=True, apaga a tabela antes de inserir.
    """
    print("🟡 " + inspect.currentframe().f_code.co_name)
    try:
        conn = psycopg2.connect(
            host=POSTGRES_HOST,
            port=POSTGRES_PORT,
            user=POSTGRES_USER,
            password=POSTGRES_PASS,
            database=nome_banco
        )
        cur = conn.cursor()

        # Verificações
        cur.execute("SELECT current_database();")
        print(f"📌 Banco conectado: {cur.fetchone()[0]}")
        
        cur.execute(f"""
            SELECT EXISTS (
                SELECT 1 
                FROM information_schema.tables 
                WHERE table_schema = '{nome_schema}' 
                AND table_name = '{nome_tabela}'
            );
        """)
        existe = cur.fetchone()[0]
        
        if not existe:
            print(f"❌ A tabela não existe! Verifique o schema e o banco de dados.")
            cur.close()
            conn.close()
            return

        # Configure o search_path ANTES de qualquer operação
        cur.execute(f"SET search_path TO {nome_schema};")
        conn.commit()

        if truncate:
            cur.execute(f'TRUNCATE TABLE {nome_tabela} RESTART IDENTITY;')
            conn.commit()
            print(f"✅ Tabela {nome_tabela} truncada")

        # Converte o DataFrame para CSV em memória
        buffer = StringIO()
        df.to_csv(buffer, index=False, header=False, sep='\t', na_rep='\\N', encoding='utf-8')
        buffer.seek(0)

        # Lista de colunas do DataFrame
        colunas = list(df.columns)

        # Usa COPY SEM o schema, pois o search_path já está configurado
        cur.copy_from(
            buffer, 
            nome_tabela,
            sep='\t', 
            null='\\N', 
            columns=colunas
        )
        conn.commit()

        # Verifica quantas linhas foram inseridas
        cur.execute(f"SELECT COUNT(*) FROM {nome_tabela};")
        count = cur.fetchone()[0]
        print(f"✅ {count} linhas inseridas na tabela {nome_tabela}")

        cur.close()
        conn.close()
        print("🟢 " + inspect.currentframe().f_code.co_name)
    
    except Exception as e:
        print(f"🔴 Erro: {e}")
        if 'conn' in locals():
            conn.rollback()
        raise