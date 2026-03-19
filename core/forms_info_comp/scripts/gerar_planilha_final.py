import os
import pandas as pd
from dotenv import load_dotenv

load_dotenv()
ROOT = os.getenv("ROOT_FORMS")
STEP1 = os.path.abspath(os.path.join(ROOT, "step_1_data_raw"))
STEP2 = os.path.abspath(os.path.join(ROOT, "step_2_stage_area"))
STEP3 = os.path.abspath(os.path.join(ROOT, "step_3_data_processed"))

def juntar_incremental(df_anterior, df_novo, coluna_ticket):
    """
    - Mantém todo o histórico de tickets válidos
    - Inclui apenas tickets novos
    - Para tickets NaN, mantém somente os do df_novo
    """

    # 1️⃣ separa tickets válidos e inválidos (NaN)
    anterior_validos = df_anterior[df_anterior[coluna_ticket].notna()]
    novo_validos = df_novo[df_novo[coluna_ticket].notna()]

    anterior_nan = df_anterior[df_anterior[coluna_ticket].isna()]
    novo_nan = df_novo[df_novo[coluna_ticket].isna()]

    # 2️⃣ tickets já existentes
    tickets_anteriores = set(anterior_validos[coluna_ticket])

    # 3️⃣ somente tickets realmente novos
    novo_validos_filtrado = novo_validos[
        ~novo_validos[coluna_ticket].isin(tickets_anteriores)
    ]

    # 4️⃣ concatenação final
    df_final = pd.concat(
        [
            anterior_validos,
            novo_validos_filtrado,
            novo_nan  # ⚠️ mantém apenas NaN do novo
        ],
        ignore_index=True
    )

    return df_final

def atualizar_repasses(rep_anterior, rep_novo):
    anterior = rep_anterior.copy()
    novo = rep_novo.copy()

    colunas_manter = ["arquivo", "tipo_arquivo", "status"]

    resultado = []

    for _, row_novo in novo.iterrows():
        ticket = row_novo["ticket_acompanhamento"]
        id_solicitacao = row_novo["id_solicitacao"]

        # -------------------------
        # CASO 1: ticket NÃO existe
        # -------------------------
        if ticket not in anterior["ticket_acompanhamento"].values:
            # verifica se id_solicitacao já existe
            existe_id = anterior[anterior["id_solicitacao"] == id_solicitacao]

            if not existe_id.empty:
                # reescreve completamente
                resultado.append(row_novo)
            else:
                # novo registro
                resultado.append(row_novo)

        # -------------------------
        # CASO 2: ticket JÁ existe
        # -------------------------
        else:
            row_anterior = anterior.loc[
                anterior["ticket_acompanhamento"] == ticket
            ].iloc[0]

            row_final = row_novo.copy()

            # mantém SOMENTE essas colunas do anterior
            for col in colunas_manter:
                if col in anterior.columns:
                    row_final[col] = row_anterior[col]

            resultado.append(row_final)

    # transforma em DataFrame
    df_resultado = pd.DataFrame(resultado)

    return df_resultado

def atualizar_formularios(form_anterior, form_novo):
    tickets_anteriores = set(form_anterior['ticket_acompanhamento'].dropna())

    novos = form_novo[
        ~form_novo['ticket_acompanhamento'].isin(tickets_anteriores)
    ]

    formularios_final = pd.concat(
        [form_anterior, novos],
        ignore_index=True
    )

    return formularios_final

def atualizar_info_complementares(info_anterior, info_novo):
    tickets_anteriores = set(info_anterior['ticket_acompanhamento'].dropna())

    novos = info_novo[
        ~info_novo['ticket_acompanhamento'].isin(tickets_anteriores)
    ]

    info_final = pd.concat(
        [info_anterior, novos],
        ignore_index=True
    )

    return info_final

def gerar_planilha_final(sebrae = True):
    ## RESPOSTAS DOS FORMULÁRIOS
    df_pdf = pd.read_excel(os.path.join(STEP2, 'dados_extraidos.xlsx'))
    info = pd.read_excel(os.path.join(STEP2, 'info_complementares.xlsx'))
    info['passou_ia'] = "Não"
    info.rename(columns={'ticket': 'ticket_acompanhamento'}, inplace=True)
    neg = pd.read_excel(os.path.join(STEP1, 'srinfo_partnership_fundsapproval.xlsx'))
    neg = neg[['id_solicitacao', 'co_negociacao', 'status_negociacao', 'parceria', 'ticket_acompanhamento', 'co_projeto',  'status', 'data_contrato', 'status_repasse', 'log_data_extracao_dados']]
    
    # Lê arquivos IA se existirem
    if os.path.exists(os.path.join(STEP2, 'info_complementares_ia.xlsx')):
        info_ia = pd.read_excel(os.path.join(STEP2, 'info_complementares_ia.xlsx'))
        info_ia['passou_ia'] = "Sim"
        info2 = pd.concat([info, info_ia], ignore_index=True)
    else:
        info2 = info.copy()

    info2 = neg[['co_negociacao', 'parceria', 'ticket_acompanhamento', 'log_data_extracao_dados']].merge(info2, on='ticket_acompanhamento', how='right')

    renomear_chaves = {
        "AMPLIAR A GAMA DE BENS OU SERVIÇOS OFERTADOS": 'Ampliar a gama de bens ou serviços ofertados',
        "AMPLIAR A PARTICIPAÇÃO DA EMPRESA NO MERCADO": 'Ampliar a participação da empresa no mercado',
        "AUMENTAR A CAPACIDADE DE PRODUÇÃO OU DE PRESTAÇÃO DE SERVIÇOS": 'Aumentar a capacidade de produção ou de prestação de serviços',
        "AUMENTAR A CAPACIDADE DE PRODUÇÃO OU DE PRESTAÇÃO DE": 'Aumentar a capacidade de produção ou de prestação de serviços',
        "AUMENTAR A FLEXIBILIDADE DA PRODUÇÃO OU DA PRESTAÇÃO DE SERVIÇOS": 'Aumentar a flexibilidade da produção ou da prestação de serviços',
        "ENQUADRAR EM REGULAÇÕES E NORMAS-PADRÃO RELATIVAS AO MERCADO INTERNO OU EXTERNO": 'Enquadrar em regulações e normas-padrão relativas ao mercado interno ou externo',
        "MERCADO INTERNO OU EXTERNO": 'Enquadrar em regulações e normas-padrão relativas ao mercado interno ou externo',
        "MELHORAR A QUALIDADE DOS BENS OU SERVIÇOS": 'Melhorar a qualidade dos bens ou serviços',
        "PERMITIR ABRERTURA DE NOVOS MERCADOS": 'Permitir abertura de novos mercados',
        "PERMITIR ABERTURA DE NOVOS MERCADOS": 'Permitir abertura de novos mercados',
        "PERMITIR CONTROLAR ASPECTOS LIGADOS À SAÚDE E/OU À SEGURANÇA": 'Permitir controlar aspectos ligados à saúde e/ou à segurança',
        "SEGURANÇA": 'Permitir controlar aspectos ligados à saúde e/ou à segurança',
        "PERMITIR MANTER A PARTICIPAÇÃO DA EMPRESA NO MERCADO": 'Permitir manter a participação da empresa no mercado',
        "PERMITIR REDUZIR O IMPACTO SOBRE O MEIO AMBIENTE": 'Permitir reduzir o impacto sobre o meio ambiente',
        "REDUZIR O CONSUMO DE ÁGUA": 'Reduzir o consumo de água',
        "REDUZIR O CONSUMO DE ENERGIA": 'Reduzir o consumo de energia',
        "REDUZIR O CONSUMO DE MATÉRIAS-PRIMAS": 'Reduzir o consumo de matérias-primas',
        "REDUZIR OS CUSTOS DE PRODUÇÃO OU DOS SERVIÇOS PRESTADOS": 'Reduzir os custos de produção ou dos serviços prestados',
        "REDUZIR OS CUSTOS DO TRABALHO": 'Reduzir os custos do trabalho',
        "SUBSTITUIR (TOTAL OU PARCIAL) MATÉRIAS-PRIMAS POR OUTRAS MENOS CONTAMINANTES OU PERIGOSAS": 'Substituir (total ou parcial) matérias-primas por outras menos contaminantes ou perigosas',
        "MENOS CONTAMINANTES OU PERIGOSAS": 'Substituir (total ou parcial) matérias-primas por outras menos contaminantes ou perigosas',
        "CONTAMINANTES OU PERIGOSAS": 'Substituir (total ou parcial) matérias-primas por outras menos contaminantes ou perigosas',
        "PERIGOSAS": 'Substituir (total ou parcial) matérias-primas por outras menos contaminantes ou perigosas',
        "SUBSTITUIR (TOTAL OU PARCIAL) MATÉRIAS-PRIMAS POR OUTRAS MENOS CONTAMINANTES OU": 'Substituir (total ou parcial) matérias-primas por outras menos contaminantes ou perigosas',
        "SUBSTITUIR (TOTAL OU PARCIAL) MATÉRIAS-PRIMAS POR OUTRAS MENOS CONTAMINANTES OU ": 'Substituir (total ou parcial) matérias-primas por outras menos contaminantes ou perigosas',
        "SUBSTITUIR (TOTAL OU PARCIAL) MATÉRIAS-PRIMAS POR OUTRAS MENOS": 'Substituir (total ou parcial) matérias-primas por outras menos contaminantes ou perigosas',
        "SUBSTITUIR (TOTAL OU PARCIAL) ENERGIA PROVENIENTE DE COMBUSTÍVEIS FÓSSEIS POR FONTES DE ENERGIA RENOVÁVEIS": 'Substituir (total ou parcial) energia proveniente de combustíveis fósseis por fontes de energia renováveis',
        "SUBSTITUIR (TOTAL OU PARCIAL) ENERGIA PROVENIENTE DE COMBUSTÍVEIS FÓSSEIS POR FONTES DE ENERGIA": 'Substituir (total ou parcial) energia proveniente de combustíveis fósseis por fontes de energia renováveis',
        "SUBSTITUIR (TOTAL OU PARCIAL) ENERGIA PROVENIENTE DE COMBUSTÍVEIS FÓSSEIS POR FONTES DE": 'Substituir (total ou parcial) energia proveniente de combustíveis fósseis por fontes de energia renováveis',
        "ENERGIA RENOVÁVEIS": 'Substituir (total ou parcial) energia proveniente de combustíveis fósseis por fontes de energia renováveis',
        "RENOVÁVEIS": 'Substituir (total ou parcial) energia proveniente de combustíveis fósseis por fontes de energia renováveis',
        "COMBUSTÍVEIS FÓSSEIS POR FONTES DE ENERGIA RENOVÁVEIS": 'Substituir (total ou parcial) energia proveniente de combustíveis fósseis por fontes de energia renováveis',
        "REDUZIR RUÍDOS OU A CONTAMINAÇÃO DO SOLO, DA ÁGUA OU DO AR": 'Reduzir ruídos ou a contaminação do solo, da água ou do ar',
        "RECICLAGEM DE RESÍDUOS, ÁGUAS RESIDUAIS OU MATERIAIS PARA VENDA E/OU REUTILIZAÇÃO": 'Reciclagem de resíduos, águas residuais ou materiais para venda e/ou reutilização',
        "VENDA E/OU REUTILIZAÇÃO": 'Reciclagem de resíduos, águas residuais ou materiais para venda e/ou reutilização',
        "REDUÇÃO DA 'PEGADA' DE CO (PRODUÇÃO TOTAL DE CO) DE SUA EMPRESA": 'Redução da pegada de CO (produção total de CO) de sua empresa',
        "EMPRESA:": 'Redução da pegada de CO (produção total de CO) de sua empresa',
        "EXPECTATIVA DE IMPACTO ESPERADO DA(S) TECNOLOGIA(S) QUE": "Expectativa de impacto esperado da(s) tecnologia(s) que será(ão) desenvolvida(s)",
        "EXPECTATIVA DE IMPACTO ESPERADO DA(S) TECNOLOGIA(S) QUE SERÁ(ÃO) DESENVOLVIDAS(S) - BAIXO,": "Expectativa de impacto esperado da(s) tecnologia(s) que será(ão) desenvolvida(s)",
        "EXPECTATIVA DE IMPACTO ESPERADO DA(S) TECNOLOGIA(S) QUE SERÁ(ÃO) DESENVOLVIDAS(S) - BAIXO, MÉDIO": "Expectativa de impacto esperado da(s) tecnologia(s) que será(ão) desenvolvida(s)",
        "EXPECTATIVA DE IMPACTO ESPERADO DA(S) TECNOLOGIA(S) QUE SERÁ(ÃO) DESENVOLVIDAS(S) - BAIXO, MÉDIO OU": "Expectativa de impacto esperado da(s) tecnologia(s) que será(ão) desenvolvida(s)",
        "ALTO/DISRUPTIVO:": "Expectativa de impacto esperado da(s) tecnologia(s) que será(ão) desenvolvida(s)",
        "SERÁ(ÃO) DESENVOLVIDAS(S) - BAIXO, MÉDIO OU ALTO/DISRUPTIVO": "Expectativa de impacto esperado da(s) tecnologia(s) que será(ão) desenvolvida(s)",
    }

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
        'Redução da pegada de CO (produção total de CO) de sua empresa',
        'Expectativa de impacto esperado da(s) tecnologia(s) que será(ão) desenvolvida(s)'
    ]

    # Aplicando as renomeações nas chaves
    df_pdf['pergunta'] = df_pdf['pergunta'].replace(renomear_chaves)

    # Filtrar o DataFrame
    df_filtrado = df_pdf[df_pdf['pergunta'].isin(chaves_desejadas)]

    # # Salvando na pasta de saída
    # df_filtrado.to_excel(os.path.join(STEP2, "info_complementares_pdfs.xlsx"), index=False)

    # Juntando os dois
    df_final2 = neg[['co_negociacao', 'parceria', 'ticket_acompanhamento', 'log_data_extracao_dados']].merge(df_filtrado[['arquivo', 'pergunta', 'resposta', 'ticket_acompanhamento']],
                                                                                                             on='ticket_acompanhamento', how='right')
    df_final2 = df_final2[['arquivo', 'ticket_acompanhamento', 'parceria', 'co_negociacao', 'pergunta', 'resposta', 'log_data_extracao_dados']]
    df_final2['passou_ia'] = "Não"

    if os.path.exists(os.path.join(STEP2, 'info_complementares_pdf_ia.xlsx')):
        df_ia = pd.read_excel(os.path.join(STEP2, 'info_complementares_pdf_ia.xlsx'))
        df_ia['passou_ia'] = "Sim"
        df_final2 = pd.concat([df_final2, df_ia], ignore_index=True)

    info_final = pd.concat([info2, df_final2], ignore_index=True)
    info_final = info_final.drop_duplicates(subset=['ticket_acompanhamento', 'pergunta', 'resposta'])
    # criando a coluna de status: "Resposta segue padrão" quando resposta for "Alta", "Média", "Baixa", "Não relevante", "ALTO/DISRUPTIVO", "MÉDIO" OU "BAIXO", "Resposta não segue padrão" caso contrário
    info_final['status_resposta'] = info_final['resposta'].apply(
        lambda x: 'Resposta segue padrão' if str(x).strip().upper() in ['ALTA', 'MÉDIA', 'BAIXA', 'NÃO RELEVANTE', 'ALTO/DISRUPTIVO', 'MÉDIO', 'BAIXO'] else 'Resposta não segue padrão'
    )
    # info_final.to_excel(os.path.join(STEP2, 'info_complementares_novo.xlsx'), index=False)
    info_final = info_final.rename(columns={'log_data_extracao_dados': 'data_extracao'})

    info_final['modificado'] = ""
    info_final['observacao'] = ""

    colunas_ordem = [
        'co_negociacao',
        'parceria',
        'ticket_acompanhamento',
        'arquivo',
        # 'tipo_arquivo',
        'pergunta',
        'resposta',
        'passou_ia',
        'status_resposta',
        'modificado',
        'observacao',
        'data_extracao'
    ]

    info_final = info_final[colunas_ordem]

    if sebrae:
        info_final = info_final[info_final['parceria'].str.contains('SEBRAE', case=False, na=False)]

    ## FORMULÁRIOS
    arquivos = os.listdir(STEP1)
    arquivos = [a for a in arquivos if not a in ['problemas.xlsx', 'srinfo_partnership_fundsapproval.xlsx', 'respostas_formularios_sebrae_anterior.xlsx', 'formularios_anteriores.zip']]
    # salvar nomes dos arquivos em excel
    df_arquivos = pd.DataFrame(arquivos, columns=["arquivo"])

    # incluir colunas de contagem de valores em info_complementares_novo.xlsx
    contagem_info = info_final.groupby('ticket_acompanhamento').size().reset_index(name='info_encontradas')

    # incluindo novas colunas (tipo de arquivo, codigo_negociaca, ticket, link, info_encontradas, passou_ia, status, data_extracao)
    df_arquivos['tipo'] = df_arquivos['arquivo'].apply(lambda x: '.pdf' if x.lower().endswith('.pdf') else '.xlsx' if x.lower().endswith('.xlsx') else '.xlsm' if x.lower().endswith('.xlsm') else '.xls' if x.lower().endswith('.xls') else 'outro')
    # para encontrar o ticket temos que usar o nome do arquivo, tirar tudo que vem antes do primeiro _ e tudo que vem depois do .
    df_arquivos['ticket_acompanhamento'] = df_arquivos['arquivo'].str.extract(r'_(\d+)\.').astype(float)
    df_arquivos = df_arquivos.merge(neg[['co_negociacao', 'parceria', 'ticket_acompanhamento', 'log_data_extracao_dados']], on='ticket_acompanhamento', how='left')
    df_arquivos['link'] = 'https://tickets.embrapii.org.br/issues/' + df_arquivos['ticket_acompanhamento'].astype(str).str.replace('.0', '', regex=False)
    # incluindo coluna de contagem de informações encontradas
    df_arquivos = df_arquivos.merge(contagem_info, on='ticket_acompanhamento', how='left')
    df_arquivos['info_encontradas'] = df_arquivos['info_encontradas'].fillna(0).astype(int)
    # incluindo a coluna de passou_ia
    df_arquivos = df_arquivos.merge(info_final[['ticket_acompanhamento', 'passou_ia']].drop_duplicates(), on='ticket_acompanhamento', how='left')
    # df_arquivos.to_excel(os.path.join(STEP2, "arquivos2.xlsx"), index=False)
    df_arquivos['status'] = df_arquivos.apply(
        lambda row: 'Sem informações' if row['info_encontradas'] == 0
        else 'Incompleto' if row['info_encontradas'] < 21
        else 'Completo' if row['info_encontradas'] == 21
        else 'Necessita verificação',
        axis=1
    )
    df_arquivos = df_arquivos.rename(columns={'log_data_extracao_dados': 'data_extracao'})

    colunas_ordem = [
        'arquivo',
        'tipo',
        'co_negociacao',
        'parceria',
        'ticket_acompanhamento',
        'link',
        'info_encontradas',
        'passou_ia',
        'status',
        'data_extracao'
    ]

    df_arquivos = df_arquivos[colunas_ordem]

    if sebrae:
        df_arquivos = df_arquivos[df_arquivos['parceria'].str.contains('SEBRAE', case=False, na=False)]

    ## REPASSES
    # obter o link do ticket a partir do ticket_acompanhamento, transformar em string e retirar os decimais
    # se não tiver ticket_acompanhamento, deixar o link em branco
    neg['link'] = neg['ticket_acompanhamento'].apply(
        lambda x: 'https://tickets.embrapii.org.br/issues/' + str(x).replace('.0', '')
        if not pd.isna(x) else None
    )
    neg['arquivo'] = neg['ticket_acompanhamento'].apply(
        lambda x: df_arquivos[df_arquivos['ticket_acompanhamento'] == x]['arquivo'].values[0] if x in df_arquivos['ticket_acompanhamento'].values
        else None
    )
    neg['tipo_arquivo'] = neg['arquivo'].apply(
        lambda x: '.pdf' if isinstance(x, str) and x.lower().endswith('.pdf')
        else '.xlsx' if isinstance(x, str) and x.lower().endswith('.xlsx')
        else '.xlsm' if isinstance(x, str) and x.lower().endswith('.xlsm')
        else '.xls' if isinstance(x, str) and x.lower().endswith('.xls')
        else 'outro' if isinstance(x, str)
        else None
    )
    neg = neg.merge(contagem_info, on='ticket_acompanhamento', how='left')
    neg['info_encontradas'] = neg['info_encontradas'].fillna(0).astype(int)
    neg.rename(columns={'status': 'status_projeto', 'log_data_extracao_dados': 'data_extracao'}, inplace=True)
    # criando a coluna de status: "Sem ticket" quando ticket_acompanhamento for nulo, "Arquivo não encontrado" quando não houver arquivo, "Sem informações" quando info_encontradas for 0, "Incompleto" quando info_encontradas for menor que 21, "Necessita verificação" quando info_encontradas for mairo que 21 e "Completo" quando info_encontradas for igual a 21
    neg['status'] = neg.apply(
        lambda row: 'Sem ticket' if pd.isna(row['ticket_acompanhamento'])
        else 'Arquivo não encontrado' if pd.isna(row['arquivo'])
        else 'Sem informações' if row['info_encontradas'] == 0
        else 'Incompleto' if row['info_encontradas'] < 21
        else 'Completo' if row['info_encontradas'] == 21
        else 'Necessita verificação',
        axis=1
    )

    colunas_ordem = [
        'id_solicitacao', 
        'co_negociacao', 
        'status_negociacao', 
        'parceria', 
        'ticket_acompanhamento', 
        'link', 
        'co_projeto', 
        'status_projeto', 
        'data_contrato', 
        'status_repasse', 
        'arquivo', 
        'tipo_arquivo', 
        'status',
        'data_extracao'
    ]
    neg = neg[colunas_ordem]
    # neg.to_excel(os.path.join(STEP2, "repasses.xlsx"), index=False)

    if sebrae:
        neg = neg[neg['parceria'].str.contains('SEBRAE', case=False, na=False)]


    ## JUNTANDO AS PLANILHAS EM UM ÚNICO EXCEL
    with pd.ExcelWriter(os.path.join(STEP2, "respostas_formularios_sebrae.xlsx")) as writer:
        neg.to_excel(writer, sheet_name='repasses', index=False)
        df_arquivos.to_excel(writer, sheet_name='formularios', index=False)
        info_final.to_excel(writer, sheet_name='info_complementares', index=False)

    
    ## VERIFICANDO SE EXISTE UMA PLANILHA ANTERIOR E JUNTANDO COM ELA
    if os.path.exists(os.path.join(STEP1, 'respostas_formularios_sebrae_anterior.xlsx')):
        repasses_anterior = pd.read_excel(os.path.join(STEP1, 'respostas_formularios_sebrae_anterior.xlsx'), sheet_name='repasses')
        formularios_anterior = pd.read_excel(os.path.join(STEP1, 'respostas_formularios_sebrae_anterior.xlsx'), sheet_name='formularios')
        info_complementares_anterior = pd.read_excel(os.path.join(STEP1, 'respostas_formularios_sebrae_anterior.xlsx'), sheet_name='info_complementares')

        repasses_final = atualizar_repasses(repasses_anterior, neg)
        formularios_final = atualizar_formularios(formularios_anterior, df_arquivos)
        info_complementares_final = atualizar_info_complementares(info_complementares_anterior, info_final)

        with pd.ExcelWriter(os.path.join(STEP3, "respostas_formularios_sebrae.xlsx")) as writer:
            repasses_final.to_excel(writer, sheet_name='repasses', index=False)
            formularios_final.to_excel(writer, sheet_name='formularios', index=False)
            info_complementares_final.to_excel(writer, sheet_name='info_complementares', index=False)
    ## SE NÃO, MOVER A PLANILHA PARA A PASTA FINAL
    else:
        os.rename(os.path.join(STEP2, "respostas_formularios_sebrae.xlsx"), os.path.join(STEP3, "respostas_formularios_sebrae.xlsx"))