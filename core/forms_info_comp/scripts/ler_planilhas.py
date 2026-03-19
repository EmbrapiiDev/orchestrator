import os
import pandas as pd
import hashlib
import urllib.parse
import warnings
from autentica_azure.requisicao import RequisicaoMS
import json
from dotenv import load_dotenv
import re

# carregar .env
load_dotenv()
ROOT = os.getenv("ROOT_FORMS")

STEP1= os.path.abspath(os.path.join(ROOT, "step_1_data_raw"))
STEP2 = os.path.abspath(os.path.join(ROOT, "step_2_stage_area"))
STEP3 = os.path.abspath(os.path.join(ROOT, "step_3_data_processed"))
BACKUP = os.path.abspath(os.path.join(ROOT, "backup"))

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def hash_arquivo(caminho_arquivo):
    """Calcula o hash MD5 de um arquivo."""
    hash_md5 = hashlib.md5()
    with open(caminho_arquivo, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()

def consolidar_planilhas(pasta_downloads):
    arquivos = [f for f in os.listdir(pasta_downloads) if f.endswith(('.xls', '.xlsx', '.xlsm', '.XLS', '.XLSX', '.XLSM'))]

    # Excluindo arquivos que não são formulário de informações complementares
    arquivo_excluir = ["srinfo_partnership_fundsapproval.xlsx", "respostas_formularios_sebrae_anterior.xlsx"]
    arquivos = [f for f in arquivos if f not in arquivo_excluir]

    registros = []
    # hashes_processados = set()

    arquivo_n = 0
    for arquivo in arquivos:
        arquivo_n += 1
        print(f"Arquivo {arquivo_n} de {len(arquivos)}")

        nome_legivel = urllib.parse.unquote(arquivo)
        caminho = os.path.join(pasta_downloads, nome_legivel)
        try:
            # hash_atual = hash_arquivo(caminho)
            # if hash_atual in hashes_processados:
            #     print(f"Aviso: {nome_legivel} é duplicado e será ignorado.")
            #     continue
            # hashes_processados.add(hash_atual)

            # Lendo os arquivos
            df = pd.read_excel(caminho, header=None)

            # Eliminando linhas totalmente vazias
            df.dropna(how='all', inplace=True)

            # Verificando se há ao menos três colunas para permitir deslocamento
            if df.shape[1] >= 3:
                # Verificando se a primeira coluna (coluna 0) está vazia nas primeiras linhas
                primeiras_linhas = df.head(10)
                if primeiras_linhas.iloc[:, 0].isna().sum() >= 8:  # mais de 80% das 10 primeiras células vazias
                    # Assumindo que os dados estão deslocados para colunas 1 e 2 (B e C)
                    df_recorte = df.iloc[:, 1:3].copy()
                    df_recorte.columns = ['pergunta', 'resposta']
                    df_recorte['arquivo'] = nome_legivel
                    registros.append(df_recorte)
                else:
                    # Assumindo que os dados estão nas colunas padrão 0 e 1 (A e B)
                    df_recorte = df.iloc[:, :2].copy()
                    df_recorte.columns = ['pergunta', 'resposta']
                    df_recorte['arquivo'] = nome_legivel
                    registros.append(df_recorte)
            elif df.shape[1] >= 2:
                df_recorte = df.iloc[:, :2].copy()
                df_recorte.columns = ['pergunta', 'resposta']
                df_recorte['arquivo'] = nome_legivel
                registros.append(df_recorte)

            else:
                print(f"Aviso: {nome_legivel} tem menos de 2 colunas relevantes.")
        except Exception as e:
            print(f"Erro ao processar {nome_legivel}: {e}")

    # Transformando a lista de dicionários em DataFrame
    df_final = pd.concat(registros, ignore_index=True)

    return df_final

def ajustes(df, pasta_saida):
    df_final = df

    # Removendo linhas com valor em branco ou NaN
    df_final = df_final[df_final['resposta'].notna() & (df_final['resposta'].astype(str).str.strip() != '')]

    # Retirando quebra de linha do nome das chaves
    df_final['pergunta'] = df_final['pergunta'].astype(str).str.replace('\n', ' ').str.strip()

    # Dicionário com renomeações das chaves
    renomear_chaves = {
        '% VALOR APORTADO EMBRAPII/BNDES:': 'pct_aporte_embrapii_bndes',
        '% VALOR APORTADO EMBRAPII:': 'pct_aporte_embrapii',
        '% VALOR APORTADO PEL0 SEBRAE:': 'pct_aporte_sebrae',
        '% VALOR APORTADO PELA UNIDADE EMBRAPII:': 'pct_aporte_unidade',
        '% VALOR APORTADO PELA(S) EMPRESA(S):': 'pct_aporte_empresa',
        '% VALOR APORTADO PELA(S) EMPRESAS(S):': 'pct_aporte_empresa',
        '% VALOR APORTADO PELO SEBRAE:': 'pct_aporte_sebrae',
        '1ª TECNOLOGIA HABILITADORA ASSOCIADA AO PROJETO P,D&I:': 'tecnologias_habilitadoras',
        '1ª TECNOLOGIA HABILITADORA ASSOCIADA AO PROJETO P,D&i:': 'tecnologias_habilitadoras',
        '1ª ÁREA DE APLICAÇÃO ASSOCIADA AO PROJETO P,D&I:': 'areas_aplicacao',
        '2ª TECNOLOGIA HABILITADORA ASSOCIADA AO PROJETO P,D&I:': 'tecnologias_habilitadoras',
        '2ª ÁREA DE APLICAÇÃO ASSOCIADA AO PROJETO P,D&I:': 'areas_aplicacao',
        '3ª TECNOLOGIA HABILITADORA ASSOCIADA AO PROJETO P,D&I:': 'tecnologias_habilitadoras',
        '3ª ÁREA DE APLICAÇÃO ASSOCIADA AO PROJETO P,D&I:': 'areas_aplicacao',
        '4ª TECNOLOGIA HABILITADORA ASSOCIADA AO PROJETO P,D&I:': 'tecnologias_habilitadoras',
        '4ª ÁREA DE APLICAÇÃO ASSOCIADA AO PROJETO P,D&I:': 'areas_aplicacao',
        'AMPLIAR A GAMA DE BENS OU SERVIÇOS OFERTADOS:': 'Ampliar a gama de bens ou serviços ofertados',
        'AMPLIAR A PARTICIPAÇÃO DA EMPRESA NO MERCADO:': 'Ampliar a participação da empresa no mercado',
        'AUMENTAR A CAPACIDADE DE PRODUÇÃO OU DE PRESTAÇÃO DE SERVIÇOS:': 'Aumentar a capacidade de produção ou de prestação de serviços',
        'AUMENTAR A FLEXIBILIDADE DA PRODUÇÃO OU DA PRESTAÇÃO DE SERVIÇOS:': 'Aumentar a flexibilidade da produção ou da prestação de serviços',
        'CASO TENHA ESCOLHIDO ""OUTROS"", PREENCHA AO LADO': 'outros',
        'CASO TENHA ESCOLHIDO "OUTROS", PREENCHA AO LADO': 'outros',
        'CNAE (GRUPO 3 DÍGITOS) DA EMPRESA DA 1ª EMPRESA:': 'cnae',
        'CNAE (GRUPO 3 DÍGITOS) DA EMPRESA DA 2ª EMPRESA:': 'cnae',
        'CNAE (GRUPO 3 DÍGITOS) DA EMPRESA DA 3ª EMPRESA:': 'cnae',
        'CNAE (GRUPO 3 DÍGITOS) DA EMPRESA DA EMPRESA:': 'cnae',
        'CNAE (GRUPO 3 DÍGITOS) DA EMPRESA:': 'cnae',
        'CNAE de idustrialização (GRUPO 3 DÍGITOS) DA EMPRESA:': 'cnae',
        'CNPJ DA 1ª EMPRESA:': 'cnpj',
        'CNPJ DA 2ª EMPRESA:': 'cnpj',
        'CNPJ DA 3ª EMPRESA:': 'cnpj',
        'CNPJ DA EMPRESA:': 'cnpj',
        'CNPJ:': 'cnpj',
        'CÓDIGO DE NEGOCIAÇÃO:': 'codigo_negociacao',
        'ENQUADRAR EM REGULAÇÕES E NORMAS-PADRÃO RELATIVAS AO MERCADO INTERNO OU EXTERNO:': 'Enquadrar em regulações e normas-padrão relativas ao mercado interno ou externo',
        'ESCALA DA TRL NA ÚLTIMA MACROENTREGA (ESPERADO NA CONCLUSÃO DO PROJETO):': 'trl_final',
        'ESCALA TRL DA 1ª MACROENTREGA DO PROJETO (NO INÍCIO DA SUA EXECUÇÃO):': 'trl_inicial',
        'ESCALA TRL DA ÚLTIMA MACROENTREGA (ESPERADO NA CONCLUSÃO DO PROJETO):': 'trl_final',
        'EXPECTATIVA DE IMPACTO ESPERADO DA(S) TECNOLOGIA(S) QUE SERÁ(ÃO) DESENVOLVIDAS(S) - BAIXO, MÉDIO OUALTO/DISRUPTIVO:': 'impacto_tecnologia',
        'EXPECTATIVA DE IMPACTO ESPERADO DA(S) TECNOLOGIA(S) QUE SERÁ(ÃO) DESENVOLVIDAS(S) - BAIXO, MÉDIO OU ALTO/DISRUPTIVO:': 'impacto_tecnologia',
        'EXPECTATIVA DE IMPACTO ESPERADO DA(S) TECNOLOGIA(S) QUE SERÁ(ÃO) DESENVOLVIDAS(S) - BAIXO,': 'impacto_tecnologia',
        'EXPECTATIVA DE TEMPO ESPERADO DA(S) TECNOLOGIA(S) QUE SERÁ(ÂO) DESENVOLVIDA(S) - BAIXO, MÉDIO OU ALTO/DISRUPTIVO:': 'tempo_tecnologia',
        'EXPECTATIVA DE TEMPO ESPERADO PARA QUE A TECNOLOGIA CHEGUE AO MERCADO (EM Nº DE MESES APÓS ACONCLUSÃO DO PROJETO):': 'tempo_chegada_mercado',
        'EXPECTATIVA DE TEMPO ESPERADO PARA QUE A TECNOLOGIA CHEGUE AO MERCADO (EM Nº DE MESES APÓS A CONCLUSÃO DO PROJETO):': 'tempo_chegada_mercado',
        'EXPECTATIVA DE TEMPO ESPERADO PARA QUE A TECNOLOGIA CHEGUE AO MERCADO (EM NÚMERO DE MESES APÓS A CONCLUSÃO DO PROJETO):': 'tempo_chegada_mercado',
        'FAIXA DE ROB NO ÚLTIMO ANO DA 1ª EMPRESA:': 'rob',
        'FAIXA DE ROB NO ÚLTIMO ANO DA 2ª EMPRESA:': 'rob',
        'FAIXA DE ROB NO ÚLTIMO ANO DA 3ª EMPRESA:': 'rob',
        'FAIXA DE ROB NO ÚLTIMO ANO DA EMPRESA:': 'rob',
        'FAIXA DE ROB NO ÚLTIMO ANO:': 'rob',
        'FOCO DO CONTRATO BNDES/EMBRAPII DA SOLICITAÇÃO DE RESERVA:': 'foco',
        'FONTE DE RECURSO SECUNDÁRIA:': 'fonte_secundaria',
        'MELHORAR A QUALIDADE DOS BENS OU SERIÇOS:': 'Melhorar a qualidade dos bens ou serviços',
        'MELHORAR A QUALIDADE DOS BENS OU SERVIÇOS:': 'Melhorar a qualidade dos bens ou serviços',
        'MODALIDADE DE APORTE DO PROJETO:': 'modalidade_aporte',
        'MODALIDADE DE FINANCIAMENTO DO PROJETO:': 'modalidade_financiamento',
        'MODALIDADE DE PROJETO:': 'modalidade_financiamento',
        'NOME DO PROJETO:': 'projeto',
        'NOME FANTASIA DA 1ª EMPRESA:': 'nome_fantasia',
        'NOME FANTASIA DA 2ª EMPRESA:': 'nome_fantasia',
        'NOME FANTASIA DA 3ª EMPRESA:': 'nome_fantasia',
        'NOME FANTASIA DA EMPRESA:': 'nome_fantasia',
        'NOME FANTASIA:': 'nome_fantasia',
        'Nº DE MACROENTREGAS PLANEJADAS:': 'num_macroentregas',
        'NÚMERO DE FUNCIONÁRIOS': 'num_funcionarios',
        'NÚMERO DE FUNCIONÁRIOS NO ÚLTIMO ANO DA 1ª EMPRESA:': 'num_funcionarios',
        'NÚMERO DE FUNCIONÁRIOS NO ÚLTIMO ANO DA 2ª EMPRESA:': 'num_funcionarios',
        'NÚMERO DE FUNCIONÁRIOS NO ÚLTIMO ANO DA 3ª EMPRESA:': 'num_funcionarios',
        'NÚMERO DE FUNCIONÁRIOS NO ÚLTIMO ANO DA EMPRESA:': 'num_funcionarios',
        'NÚMERO DE FUNCIONÁRIOS NO ÚLTIMO ANO:': 'num_funcionarios',
        'O PROJETO IRÁ USAR RECURSOS DOS CONTRATOS SEBRAE': 'recursos_sebrae',
        'OBJETIVO DO PROJETO:': 'objetivo',
        'PERMITIR ABERTURA DE NOVOS MERCADOS:': 'Permitir abertura de novos mercados',
        'PERMITIR ABRERTURA DE NOVOS MERCADOS:': 'Permitir abertura de novos mercados',
        'PERMITIR CONTROLAR ASPECTOS LIGADOS À SAÚDE E/OU À SEGURANÇA:': 'Permitir controlar aspectos ligados à saúde e/ou à segurança',
        'PERMITIR MANTER A PARTICIPAÇÃO DA EMPRESA NO MERCADO:': 'Permitir manter a participação da empresa no mercado',
        'PERMITIR REDUZIR O IMPACTO SOBRE O MEIO AMBIENTE:': 'Permitir reduzir o impacto sobre o meio ambiente',
        'PORTE EMPRESA POR NÚMERO DE FUNCIONÁRIOS:': 'porte_num_funcionarios',
        'QUAL É A EXPECTATIVA DE SIGNIFICÂNCIA DA(S) INOVAÇÃO(ÕES) QUE SERÁ(ÃO) GERADA(S) NO PROJETO?': 'significancia_inovacao',
        'QUAL É A EXPECTATIVA DE SIGNIFICÂNCIA DA(S) INOVAÇÃO(ÕES) QUE SERÁ(ÃO) GERADA(S) NO': 'significancia_inovacao',
        'QUAL É O GRAU DE IMPACTO ESPERADO DO PROJETO (NA EMPRESA E NO MERCADO), EM RELAÇÃO AOSSEGUINTES PONTOS:': 'QUAL É O GRAU DE IMPACTO ESPERADO DO PROJETO (NA EMPRESA E NO MERCADO), EM RELAÇÃO AOS SEGUINTES PONTOS:',
        'QUAL É O GRAU DE IMPACTO ESPERADO DO PROJETO (NA EMPRESA E NO MERCADO), EM RELAÇÃO AOS SEGUINTES PONTOS': 'QUAL É O GRAU DE IMPACTO ESPERADO DO PROJETO (NA EMPRESA E NO MERCADO), EM RELAÇÃO AOS SEGUINTES PONTOS:',
        'QUAL É O GRAU DE IMPACTO ESPERADO DO PROJETO (NA EMPRESA E NO MERCADO), EM RELAÇÃO AOS SEGUINTES PONTOS:': 'QUAL É O GRAU DE IMPACTO ESPERADO DO PROJETO (NA EMPRESA E NO MERCADO), EM RELAÇÃO AOS SEGUINTES PONTOS:',
        'QUAL É O GRAU DE IMPACTO ESPERADO DO PROJETO (NA EMPRESA E NO MERCADO), EM RELAÇÃO AOS SEGUINTES PONTOS: ': 'QUAL É O GRAU DE IMPACTO ESPERADO DO PROJETO (NA EMPRESA E NO MERCADO), EM RELAÇÃO AOS SEGUINTES PONTOS:',
        'RAZÃO SOCIAL DA 10º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 11º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 12º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 13º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 14º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 15º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 1ª EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 1º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 2ª EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 2º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 3ª EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 3º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 4º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 5º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 6º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 7º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 8º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA 9º EMPRESA:': 'empresa',
        'RAZÃO SOCIAL DA EMPRESA:': 'empresa',
        'RECICLAGEM DE RESÍDUOS, ÁGUAS RESIDUAIS OU MATERIAIS PARA VENDA E/OU REUTILIZAÇÃO:': 'Reciclagem de resíduos, águas residuais ou materiais para venda e/ou reutilização',
        'REDUZIR O CONSUMO DE ENERGIA:': 'Reduzir o consumo de energia',
        'REDUZIR O CONSUMO DE MATÉRIAS-PRIMAS:': 'Reduzir o consumo de matérias-primas',
        'REDUZIR O CONSUMO DE ÁGUA:': 'Reduzir o consumo de água',
        'REDUZIR OS CUSTOS DE PRODUÇÃO OU DOS SERVIÇOS PRESTADOS': 'Reduzir os custos de produção ou dos serviços prestados',
        'REDUZIR OS CUSTOS DE PRODUÇÃO OU DOS SERVIÇOS PRESTADOS:': 'Reduzir os custos de produção ou dos serviços prestados',
        'REDUZIR OS CUSTOS DO TRABALHO:': 'Reduzir os custos do trabalho',
        'REDUZIR OS CUSTOS DO TRABALHO: ': 'Reduzir os custos do trabalho',
        'REDUZIR RUÍDOS OU A CONTAMINAÇÃO DO SOLO, DA ÁGUA OU DO AR:': 'Reduzir ruídos ou a contaminação do solo, da água ou do ar',
        'REDUÇÃO DA \'PEGADA\' DE CO (PRODUÇÃO TOTAL DE CO) DE SUA EMPRESA:': 'Redução da pegada de CO (produção total de CO) de sua empresa',
        'RESULTADOS ESPERADOS COM A CONCLUSÃO DO PROJETO (DESCRITIVO - MÁX DE 500 CARACTERES):': 'resultados_esperados',
        'RESULTADOS ESPERADOS COM A CONCLUSÃO DO PROJETO (DESCRITIVO - MÁXIMO DE 500 CARACTERES):': 'resultados_esperados',
        'SUBSTITUIR (TOTAL OU PARCIAL) ENERGIA PROVENIENTE DE COMBUSTÍVEIS FÓSSEIS POR FONTES DE ENERGIARENOVÁVEIS:': 'Substituir (total ou parcial) energia proveniente de combustíveis fósseis por fontes de energia renováveis',
        'SUBSTITUIR (TOTAL OU PARCIAL) ENERGIA PROVENIENTE DE COMBUSTÍVEIS FÓSSEIS POR FONTES DE ENERGIA RENOVÁVEIS:': 'Substituir (total ou parcial) energia proveniente de combustíveis fósseis por fontes de energia renováveis',
        'SUBSTITUIR (TOTAL OU PARCIAL) ENERGIA PROVENIENTE DE COMBUSTÍVEIS FÓSSEIS POR FONTES DE': 'Substituir (total ou parcial) energia proveniente de combustíveis fósseis por fontes de energia renováveis',
        'SUBSTITUIR (TOTAL OU PARCIAL) MATÉRIAS-PRIMAS POR OUTRAS MENOS CONTAMINANTES OU PERIGOSAS:': 'Substituir (total ou parcial) matérias-primas por outras menos contaminantes ou perigosas',
        'SUBSTITUIR (TOTAL OU PARCIAL) MATÉRIAS-PRIMAS POR OUTRAS MENOS CONTAMINANTES OU': 'Substituir (total ou parcial) matérias-primas por outras menos contaminantes ou perigosas',
        'TIPO DE IMPACTO PRODUTIVO ESPERADO COM O PROJETO:': 'impacto_produtivo',
        'TÍTULO DO PROJETO:': 'projeto',
        'UF DO CNPJ DA 1ª EMPRESA:': 'uf',
        'UF DO CNPJ DA 2ª EMPRESA:': 'uf',
        'UF DO CNPJ DA 3ª EMPRESA:': 'uf',
        'UF DO CNPJ DA EMPRESA:': 'uf',
        'UF DO CNPJ:': 'uf',
        'UNIDADE EMBRAPII:': 'unidade_embrapii',
        'UNIDADE EMBRAPII: ': 'unidade_embrapii',
        'VALOR APORTADO PELA EMBRAPII/BNDES:': 'valor_embrapii_bndes',
        'VALOR APORTADO PELA EMBRAPII:': 'valor_embrapii',
        'VALOR APORTADO PELA UNIDADE EMBRAPII:': 'valor_unidade_embrapii',
        'VALOR APORTADO PELA(S) EMPRESA(S):': 'valor_empresa',
        'VALOR APORTADO PELA(S) EMPRESAS(S):': 'valor_empresa',
        'VALOR APORTADO PELO SEBRAE:': 'valor_sebrae',
        'VALOR TOTAL:': 'valor_total',
        'Ampliar a gama de bens ou serviços ofertados': 'Ampliar a gama de bens ou serviços ofertados',
        'Ampliar a participação da empresa no mercado': 'Ampliar a participação da empresa no mercado',
        'Aumentar a capacidade de produção ou de prestação de serviços': 'Aumentar a capacidade de produção ou de prestação de serviços',
        'Enquadrar em regulações e normas-padrão relativas ao mercado interno ou externo': 'Enquadrar em regulações e normas-padrão relativas ao mercado interno ou externo',
        'Melhorar a qualidade dos bens ou serviços': 'Melhorar a qualidade dos bens ou serviços',
        'Permitir a abertura de novos mercados': 'Permitir abertura de novos mercados',
        'Permitir controlar aspectos ligados à saúde e/ou à segurança': 'Permitir controlar aspectos ligados à saúde e/ou à segurança',
        'Permitir manter a participação da empresa no mercado': 'Permitir manter a participação da empresa no mercado',
        'Permitir reduzir o impacto sobre o meio ambiente': 'Permitir reduzir o impacto sobre o meio ambiente',
        'Reduzir o consumo de água': 'Reduzir o consumo de água',
        'Reduzir o consumo de energia': 'Reduzir o consumo de energia',
        'Reduzir o consumo de matérias-primas': 'Reduzir o consumo de matérias-primas',
        'Reduzir os custos de produção ou dos serviços prestados': 'Reduzir os custos de produção ou dos serviços prestados',
        'Reduzir os custos do trabalho': 'Reduzir os custos do trabalho',
        'Substituir total ou parcial matérias-primas por outras menos contaminantes ou perigosas': 'Substituir (total ou parcial) matérias-primas por outras menos contaminantes ou perigosas',
        'Substituir total ou parcial energias provenientes de combustíveis fósseis por fontes de energia renováveis': 'Substituir (total ou parcial) energia proveniente de combustíveis fósseis por fontes de energia renováveis',
        'Reduzir ruídos ou a contaminação do solo, da água ou do ar': 'Reduzir ruídos ou a contaminação do solo, da água ou do ar',
        'Reciclagem de resíduos, águas residuárias ou materiais para venda e/ou reutilização': 'Reciclagem de resíduos, águas residuais ou materiais para venda e/ou reutilização',
        "Redução da 'pegada' de CO (produção total de CO) da sua empresa": 'Redução da pegada de CO (produção total de CO) de sua empresa',
        'Expectativa de impacto esperado da(s) tecnologia(s) que será(ão) desenvolvida(s) - baixo, médio ou alto/disruptivo': 'impacto_tecnologia',
    }

    # Aplicando as renomeações nas chaves
    df_final['pergunta'] = df_final['pergunta'].replace(renomear_chaves)

    # Criando a coluna de tickets
    df_final['ticket'] = df_final['arquivo'].str.extract(r'_(\d+)\.')


    ## SALVANDO O ARQUIVO COM AS PERGUNTAS COMPLEMENTARES SOBRE O PROJETO
    # Escolhendo as chaves desejadas
    chaves_desejadas = [
        'Ampliar a gama de bens ou serviços ofertados',
        'Ampliar a participação da empresa no mercado',
        'Aumentar a capacidade de produção ou de prestação de serviços',
        'Aumentar a flexibilidade da produção ou da prestação de serviços',
        'Enquadrar em regulações e normas-padrão relativas ao mercado interno ou externo',
        'Melhorar a qualidade dos bens ou serviços',
        'Permitir abertura de novos mercados',
        'Permitir controlar aspectos ligados à saúde e/ou à segurança',
        'Permitir manter a participação da empresa no mercado',
        'Permitir reduzir o impacto sobre o meio ambiente',
        'Reduzir o consumo de água',
        'Reduzir o consumo de energia',
        'Reduzir o consumo de matérias-primas',
        'Reduzir os custos de produção ou dos serviços prestados',
        'Reduzir os custos do trabalho',
        'Substituir (total ou parcial) energia proveniente de combustíveis fósseis por fontes de energia renováveis',
        'Substituir (total ou parcial) matérias-primas por outras menos contaminantes ou perigosas',
        'Reduzir ruídos ou a contaminação do solo, da água ou do ar',
        'Reciclagem de resíduos, águas residuais ou materiais para venda e/ou reutilização',
        'Redução da pegada de CO (produção total de CO) de sua empresa'
    ]

    # Filtrar o DataFrame
    df_filtrado = df_final[df_final['pergunta'].isin(chaves_desejadas)]

    # Reordenando
    df_filtrado = df_filtrado[['arquivo', 'ticket', 'pergunta', 'resposta']]

    # Modificando palavras específicas
    palavras_alvo = {
        'Alto': 'Alta',
        # 'Além': 'Alta',
        # 'Sim': 'Alta',
        # 'SIM': 'Alta',
        # 'Significativo': 'Média',
        'Media': 'Média',
        'Medio': 'Média',
        # 'Apesar': 'Média',
        'Baixo': 'Baixa',
        'Pouco': 'Baixa',
        'Não': 'Não relevante',
        'não': 'Não relevante',
        'na': 'Não relevante',
        'NÃO SE APLICA': 'Não relevante'
    }

    # Aplica substituições
    for palavra, novo_valor in palavras_alvo.items():
        mask = df_filtrado['resposta'].str.contains(palavra, na=False)
        df_filtrado.loc[mask, 'resposta'] = novo_valor

    df_filtrado = df_filtrado.drop_duplicates(subset=['ticket', 'pergunta', 'resposta'])
    # incluir impacto_tecnologia em df_filtrado como nova linha
    impacto_tecnologia = df_final[df_final['pergunta'] == 'impacto_tecnologia'][['arquivo', 'ticket', 'resposta']]

    # Modificando palavras específicas
    palavras_alvo_tecnologia = {
        'Alto': 'ALTO/DISRUPTIVO',
        'Alta': 'ALTO/DISRUPTIVO',
        'alto/disruptivo': 'ALTO/DISRUPTIVO',
        'alto': 'ALTO/DISRUPTIVO',
        'alta': 'ALTO/DISRUPTIVO',
        'media': 'MÉDIO',
        'Média': 'MÉDIO',
        'Media': 'MÉDIO',
        'medio': 'MÉDIO',
        'Medio': 'MÉDIO',
        'baixo': 'BAIXO',
        'Baixo': 'BAIXO',
        'Baixa': 'BAIXO',
        'pequena': 'BAIXO',
        'pouco': 'BAIXO',
        'Pouco': 'BAIXO',
    }

    # Aplica substituições
    for palavra, novo_valor in palavras_alvo_tecnologia.items():
        mask2 = impacto_tecnologia['resposta'].str.contains(palavra, na=False)
        impacto_tecnologia.loc[mask2, 'resposta'] = novo_valor

    df_filtrado = pd.concat([df_filtrado, impacto_tecnologia.assign(pergunta='Expectativa de impacto esperado da(s) tecnologia(s) que será(ão) desenvolvida(s)')], ignore_index=True)
    df_filtrado = df_filtrado.drop_duplicates(subset=['ticket', 'pergunta', 'resposta'])

    # Salvando na pasta de saída
    df_filtrado.to_excel(os.path.join(pasta_saida, "info_complementares.xlsx"), index=False)


    ## SALVANDO O ARQUIVO COM INFORMAÇÕES GERAIS DO PROJETO, EMPRESAS ETC
    # Agrupando por Arquivo + Chave e juntando os valores com "; " se for o caso
    df_agrupado = df_final.groupby(['arquivo', 'ticket', 'pergunta'])['resposta'].apply(lambda x: '; '.join(x.astype(str))).reset_index()

    # Pivotando
    df_pivot = df_agrupado.pivot(index=['arquivo', 'ticket'], columns='pergunta', values='resposta').reset_index()

    # Modificando palavras específicas
    df_pivot.loc[df_pivot['areas_aplicacao'].astype(str).str.contains(r'[1-9]', na=False), 'areas_aplicacao'] = ''
    df_pivot.loc[df_pivot['outros'].astype(str).str.contains(r'[1-9]', na=False), 'outros'] = ''
    df_pivot['cnpj'] = df_pivot['cnpj'].astype(str).str.replace('cnpj ', '', case=False)
    df_pivot.loc[df_pivot['num_macroentregas'].astype(str).str.contains('duas', case=False, na=False), 'num_macroentregas'] = 2
    df_pivot.loc[df_pivot['num_macroentregas'].astype(str).str.contains('três', case=False, na=False), 'num_macroentregas'] = 3

    trl = [
        '3. Estabelecimento de função crítica de forma analítica, experimental e/ou prova de conceito',
        '4. Validação funcional dos componentes em ambiente de laboratório',
        '5. Validação das funções críticas dos componentes em ambiente relevante',
        '6. Demonstração de funções críticas do protótipo em ambiente relevante',
        '7. Demonstração de protótipo do sistema em ambiente operacional',
        '8. Sistema qualificado e finalizado',
        '9. Sistema operando e comprovado em todos os aspectos de sua missão operacional'
    ]
    for i in range(len(trl)):
        df_pivot.loc[df_pivot['trl_inicial'].astype(str).str.contains(str(i+3), na=False), 'trl_inicial'] = trl[i]
        df_pivot.loc[df_pivot['trl_final'].astype(str).str.contains(str(i+3), na=False), 'trl_final'] = trl[i]

    valores = [
        'pct_aporte_embrapii',
        'pct_aporte_embrapii_bndes',
        'pct_aporte_empresa',
        'pct_aporte_sebrae',
        'pct_aporte_unidade',
        'valor_embrapii',
        'valor_embrapii_bndes',
        'valor_empresa',
        'valor_sebrae',
        'valor_unidade_embrapii',
        'valor_total'
    ]

    for i in range(len(valores)):
        df_pivot[valores[i]] = df_pivot[valores[i]].astype(str).apply(lambda x: x.replace('.', '') if 'R$ ' in x else x)
        df_pivot[valores[i]] = df_pivot[valores[i]].str.replace('R$ ', '', case=False)
        df_pivot[valores[i]] = df_pivot[valores[i]].str.replace(',', '.', case=False)
        df_pivot[valores[i]] = df_pivot[valores[i]].apply(lambda x: x.split(';')[0].strip() if ';' in x else x)
        # df_pivot[valores[i]] = df_pivot[valores[i]].astype(float)


    # Escolhendo as colunas
    colunas_informacoes = [
        'arquivo',
        'ticket',
        'codigo_negociacao',
        'unidade_embrapii',
        'modalidade_financiamento',
        'modalidade_aporte',
        'foco',
        'trl_inicial',
        'trl_final',
        'num_macroentregas',
        'projeto',
        'objetivo',
        'recursos_sebrae',
        'fonte_secundaria',
        'impacto_produtivo',
        'areas_aplicacao',
        'tecnologias_habilitadoras',
        'outros',
        'resultados_esperados',
        'tempo_chegada_mercado',
        'impacto_tecnologia',
        'significancia_inovacao',
        'cnpj',
        'empresa',
        'nome_fantasia',
        'uf',
        'cnae',
        'rob',
        'num_funcionarios',
        'pct_aporte_embrapii',
        # 'pct_aporte_embrapii_bndes',
        'pct_aporte_empresa',
        'pct_aporte_sebrae',
        'pct_aporte_unidade',
        'valor_embrapii',
        'valor_embrapii_bndes',
        'valor_empresa',
        'valor_sebrae',
        'valor_unidade_embrapii',
        'valor_total'
        ]

    df_pivot = df_pivot[colunas_informacoes]

    colunas_para_usar = df_pivot.columns.difference(['arquivo'])
    df_pivot = df_pivot.drop_duplicates(subset=colunas_para_usar)

    df_pivot.to_excel(os.path.join(pasta_saida, 'info_gerais.xlsx'), index=False)
    print(f'Arquivos salvos com sucesso na pasta {pasta_saida}')





MAPA_PERGUNTAS = {
    "impactos_ampliar_gama_bens_servicos": "Ampliar a gama de bens ou serviços ofertados",
    "impactos_ampliar_participacao_mercado": "Ampliar a participação da empresa no mercado",
    "impactos_aumentar_capacidade_producao": "Aumentar a capacidade de produção ou de prestação de serviços",
    "impactos_aumentar_flexibilidade": "Aumentar a flexibilidade da produção ou da prestação de serviços",
    "impactos_enquadrar_regulacoes": "Enquadrar em regulações e normas-padrão relativas ao mercado interno ou externo",
    "impactos_melhorar_qualidade": "Melhorar a qualidade dos bens ou serviços",
    "impactos_abrir_novos_mercados": "Permitir abertura de novos mercados",
    "impactos_controlar_saude_seguranca": "Permitir controlar aspectos ligados à saúde e/ou à segurança",
    "impactos_manter_participacao": "Permitir manter a participação da empresa no mercado",
    "impactos_reduzir_impacto_ambiental": "Permitir reduzir o impacto sobre o meio ambiente",
    "impactos_reduzir_consumo_agua": "Reduzir o consumo de água",
    "impactos_reduzir_consumo_energia": "Reduzir o consumo de energia",
    "impactos_reduzir_consumo_materias_primas": "Reduzir o consumo de matérias-primas",
    "impactos_reduzir_custos_producao": "Reduzir os custos de produção ou dos serviços prestados",
    "impactos_reduzir_custos_trabalho": "Reduzir os custos do trabalho",
    "impactos_substituir_energia_fossil": "Substituir (total ou parcial) energia proveniente de combustíveis fósseis por fontes de energia renováveis",
    "impactos_substituir_materias_primas": "Substituir (total ou parcial) matérias-primas por outras menos contaminantes ou perigosas",
    "impactos_reduzir_ruidos_contaminacao": "Reduzir ruídos ou a contaminação do solo, da água ou do ar",
    "impactos_reciclagem_residuos": "Reciclagem de resíduos, águas residuais ou materiais para venda e/ou reutilização",
    "impactos_reducao_pegada_co": "Redução da pegada de CO (produção total de CO) de sua empresa"
}

MAPA_IMPACTO_TECNOLOGIA = {
    "BAIXA": "BAIXO",
    "MÉDIA": "MÉDIO",
    "ALTA": "ALTO/DISRUPTIVO"
}


def excel_para_texto_colunas(caminho):
    df = pd.read_excel(caminho, header=None)

    linhas = []
    rotulo_pendente = None

    for _, row in df.iterrows():
        col_a = str(row[0]).strip() if pd.notna(row[0]) else None
        col_b = str(row[1]).strip() if len(row) > 1 and pd.notna(row[1]) else None

        # Caso clássico: A = pergunta, B = resposta
        if col_a and col_b:
            linhas.append(f"A: {col_a}")
            linhas.append(f"B: {col_b}")
            rotulo_pendente = None

        # A tem pergunta, B vazio → guarda
        elif col_a and not col_b:
            linhas.append(f"A: {col_a}")
            rotulo_pendente = col_a

        # A vazio, B tem valor → resposta da última pergunta
        elif not col_a and col_b and rotulo_pendente:
            linhas.append(f"B: {col_b}")
            rotulo_pendente = None

        # Tudo vazio → ignora
        else:
            continue

    return "\n".join(linhas)

def parse_json_from_text(text):
        """
        Tenta extrair JSON válido de um texto (remove cercas de código, captura primeiro bloco {}).
        """
        # Remover cercas de código se existirem
        if '```' in text:
            parts = text.split('```')
            # Busca o primeiro trecho que pareça JSON
            for i in range(1, len(parts), 2):
                candidate = parts[i].strip()
                if candidate.startswith('json'):
                    candidate = candidate[4:].strip()
                if candidate.startswith('{') and candidate.endswith('}'):
                    try:
                        return json.loads(candidate)
                    except Exception:
                        pass
            # Se não achou, tenta concatenar removendo as cercas
            text = ''.join(p for idx, p in enumerate(parts) if idx % 2 == 0)

        # Tenta parse direto
        try:
            return json.loads(text)
        except Exception:
            pass

        # Tenta localizar o primeiro bloco { ... }
        try:
            start = text.find('{')
            end = text.rfind('}')
            if start != -1 and end != -1 and end > start:
                return json.loads(text[start:end+1])
        except Exception:
            pass

        raise json.JSONDecodeError('Não foi possível extrair JSON da resposta do modelo', text, 0)

def requisicao(prompt, modelo = 'o4-mini'):
    requisicao_ms = RequisicaoMS()

    response_text = requisicao_ms.requisita(
        mensagem=prompt,
        model=modelo
    )
    
    if response_text == "erro":
        print(f"Erro na requisição")
        return None
    
    # Faz parse do JSON
    try:
        resposta = parse_json_from_text(response_text)
        
        # # Armazenar justificativas específicas com prefixo
        # if 'justificativa' in classificacao:
        #     classificacao[f'{prompt_key}_justificativa'] = classificacao['justificativa']
        # if 'confianca' in classificacao:
        #     classificacao[f'{prompt_key}_confianca'] = classificacao['confianca']
        
        # print(f"Classificação {prompt_key} realizada com sucesso")
        return resposta
    except json.JSONDecodeError as e:
        print(f"Erro ao fazer parse do JSON: {e}")
        return None

def gerar_prompt(texto):
    prompt_base = """
O texto abaixo foi extraído de um EXCEL que pode estar mal formatado.
Existem perguntas e respostas desalinhadas. A ideia é que as perguntas estejam na coluna A e as respostas na coluna B.
Associe semanticamente cada resposta à pergunta correta.

Sua tarefa:
1. Interpretar semanticamente o texto
2. Identificar perguntas e seus respectivos valores
3. Associar corretamente cada resposta à sua pergunta
4. Normalizar os dados conforme o schema abaixo
5. Retornar APENAS um JSON válido
6. Não inventar informações
7. Se houver ambiguidade, usar null
----

Texto a ser interpretado:
{texto}

----
Campos esperados (normalize exatamente estes nomes):

Informações gerais:
- projeto
- objetivo
- codigo_negociacao
- unidade_embrapii
- modalidade_financiamento
- modalidade_aporte
- foco
- trl_inicial
- trl_final
- num_macroentregas
- recursos_sebrae
- fonte_secundaria
- impacto_produtivo
- impacto_tecnologia
- significancia_inovacao
- tempo_chegada_mercado
- resultados_esperados

Empresas (lista):
- empresas: [
    {{
      empresa,
      nome_fantasia,
      cnpj,
      uf,
      cnae,
      rob,
      num_funcionarios
    }}
  ]

Áreas e tecnologias:
- areas_aplicacao
- tecnologias_habilitadoras
- outros

Financeiro:
- pct_aporte_embrapii
- pct_aporte_embrapii_bndes
- pct_aporte_empresa
- pct_aporte_sebrae
- pct_aporte_unidade
- valor_embrapii
- valor_embrapii_bndes
- valor_empresa
- valor_sebrae
- valor_unidade_embrapii
- valor_total

Impactos esperados (classifique exatamente como):
- Alta
- Média
- Baixa
- Não relevante

Campos de impacto:
- ampliar_gama_bens_servicos
- ampliar_participacao_mercado
- aumentar_capacidade_producao
- aumentar_flexibilidade
- enquadrar_regulacoes
- melhorar_qualidade
- abrir_novos_mercados
- controlar_saude_seguranca
- manter_participacao
- reduzir_impacto_ambiental
- reduzir_consumo_agua
- reduzir_consumo_energia
- reduzir_consumo_materias_primas
- reduzir_custos_producao
- reduzir_custos_trabalho
- substituir_energia_fossil
- substituir_materias_primas
- reduzir_ruidos_contaminacao
- reciclagem_residuos
- reducao_pegada_co


Formato de saída (exemplo mínimo, APENAS em JSON válido):

{{
  "projeto": "...",
  "empresas": [...],
  "financeiro": {{...}},
  "impactos": {{...}}
}}
"""
    
    variables = {
        "texto": texto
    }

    try:
        formatted_prompt = prompt_base.format(**variables)
        return formatted_prompt
    except Exception as e:
        raise ValueError(f"Erro ao formatar o prompt: {str(e)}")

def gerar_planilhas_ia(pasta, arquivos, pasta_saida):
    # Excluindo arquivos em branco ou que não são formulário de informações complementares
    arquivo_excluir = ["srinfo_partnership_fundsapproval.xlsx", "info_complementares.xlsx", "info_gerais.xlsx"]
    arquivos = [f for f in arquivos if f not in arquivo_excluir]

    # arquivos = arquivos[:2]  # Limitando para os 2 primeiros arquivos para teste
    # arquivos = ['formulario_53698.xlsx']

    resultados = []

    for arquivo in arquivos:
        print(f"Processando arquivo: {arquivo}")

        texto = excel_para_texto_colunas(os.path.join(pasta, arquivo))
        prompt = gerar_prompt(texto)
        resposta = requisicao(prompt=prompt, modelo='o4-mini')
        resposta["arquivo"] = os.path.basename(arquivo)

        resultados.append(resposta)

    df_final = pd.json_normalize(resultados, sep="_")
    df_final.to_excel(os.path.join(pasta_saida, "resultados_raw_ia.xlsx"), index=False)

    df_final = pd.read_excel(os.path.join(pasta_saida, "resultados_raw_ia.xlsx"))
    linhas = []

    for _, row in df_final.iterrows():
        arquivo = row["arquivo"]
        match = re.search(r"(\d+)", arquivo)
        ticket = match.group(1) if match else None

        for campo, pergunta in MAPA_PERGUNTAS.items():
            resposta = row.get(campo)

            linhas.append({
                "arquivo": arquivo,
                "ticket": ticket,
                "pergunta": pergunta,
                "resposta": resposta
            })

    df_impactos = pd.DataFrame(linhas)
    df_impactos = df_impactos.dropna(subset=['resposta'])  #tirar nulos

    df_impacto_tecnologia = df_final[['arquivo', 'impacto_tecnologia']].copy()
    df_impacto_tecnologia['ticket'] = df_impacto_tecnologia['arquivo'].str.extract(r'_(\d+)\.')
    df_impacto_tecnologia = df_impacto_tecnologia.rename(columns={'impacto_tecnologia': 'pergunta'})
    df_impacto_tecnologia['resposta'] = df_impacto_tecnologia['pergunta']
    df_impacto_tecnologia['pergunta'] = 'Expectativa de impacto esperado da(s) tecnologia(s) que será(ão) desenvolvida(s)'
    df_impacto_tecnologia = df_impacto_tecnologia.dropna(subset=['resposta'])  #tirar nulos
    df_impacto_tecnologia['resposta'] = df_impacto_tecnologia['resposta'].str.upper().apply(lambda x: MAPA_IMPACTO_TECNOLOGIA.get(x, x))

    # Juntar df_impactos com df_impacto_tecnologia
    df_impactos = pd.concat([df_impactos, df_impacto_tecnologia], ignore_index=True)

    caminho_saida = os.path.join(pasta_saida, "info_complementares_ia.xlsx")
    df_impactos.to_excel(caminho_saida, index=False)

    return df_final, df_impactos