import os
import fitz
import pandas as pd
from dotenv import load_dotenv
import pytesseract
from PIL import Image
import io
import json
from autentica_azure.requisicao import RequisicaoMS
import re

# Carregar variáveis de ambiente do arquivo .env
load_dotenv()
ROOT = os.getenv("ROOT_FORMS")
STEP1 = os.path.abspath(os.path.join(ROOT, "step_1_data_raw"))
STEP2 = os.path.abspath(os.path.join(ROOT, "step_2_stage_area"))
STEP3 = os.path.abspath(os.path.join(ROOT, "step_3_data_processed"))

# Processa os PDFs normalmente
def ler_pdfs():
    # Lista de frases de interesse
    frases_interesse = [
        "UNIDADE EMBRAPII",
        "CÓDIGO DE NEGOCIAÇÃO",
        "MODALIDADE DE FINANCIAMENTO DO PROJETO",
        "FOCO DO CONTRATO BNDES/EMBRAPII DA SOLICITAÇÃO DE RESERVA",
        "FONTE DE RECURSO SECUNDÁRIA",
        "EMPRESA:",
        "RAZÃO SOCIAL DA 1ª EMPRESA",
        "RAZÃO SOCIAL DA 1º EMPRESA",
        "RAZÃO SOCIAL DA 2ª EMPRESA",
        "RAZÃO SOCIAL DA 3ª EMPRESA",
        "RAZÃO SOCIAL DA 4ª EMPRESA",
        "NOME FANTASIA",
        "CNPJ",
        "UF DO CNPJ",
        "NÚMERO DE FUNCIONÁRIOS NO ÚLTIMO ANO",
        "FAIXA DE ROB NO ÚLTIMO ANO",
        "CNAE (GRUPO 3 DÍGITOS) DA EMPRESA",
        "VALOR TOTAL",
        "VALOR APORTADO PELA EMBRAPII",
        "VALOR APORTADO PELA EMBRAPII/BNDES",
        "% VALOR APORTADO EMBRAPII",
        "% VALOR APORTADO EMBRAPII/BNDES",
        "VALOR APORTADO PELA(S) EMPRESA(S)",
        "% VALOR APORTADO PELA(S) EMPRESA(S)",
        "VALOR APORTADO PELO SEBRAE",
        "% VALOR APORTADO PEL0 SEBRAE",
        "% VALOR APORTADO PELO SEBRAE",
        "VALOR APORTADO PELA UNIDADE EMBRAPII",
        "% VALOR APORTADO PELA UNIDADE EMBRAPII",
        "NOME DO PROJETO",
        "OBJETIVO DO PROJETO",
        "TIPO DE IMPACTO PRODUTIVO ESPERADO COM O PROJETO",
        "1ª ÁREA DE APLICAÇÃO ASSOCIADA AO PROJETO P,D&I",
        "2ª ÁREA DE APLICAÇÃO ASSOCIADA AO PROJETO P,D&I",
        "3ª ÁREA DE APLICAÇÃO ASSOCIADA AO PROJETO P,D&I",
        "1ª TECNOLOGIA HABILITADORA ASSOCIADA AO PROJETO P,D&I",
        "2ª TECNOLOGIA HABILITADORA ASSOCIADA AO PROJETO P,D&I",
        "3ª TECNOLOGIA HABILITADORA ASSOCIADA AO PROJETO P,D&I",
        "CASO TENHA ESCOLHIDO \"OUTROS\", PREENCHA AO LADO",
        "Nº DE MACROENTREGAS PLANEJADAS",
        "ESCALA TRL DA 1ª MACROENTREGA DO PROJETO (NO INÍCIO DA SUA EXECUÇÃO)",
        "ESCALA TRL DA 1ª MACROENTREGA DO PROJETO (NO INÍCIO DA SUA",
        "EXECUÇÃO)",
        "ESCALA TRL DA ÚLTIMA MACROENTREGA (ESPERADO NA CONCLUSÃO DO PROJETO)",
        "ESCALA TRL DA ÚLTIMA MACROENTREGA (ESPERADO NA CONCLUSÃO DO",
        "ESCALA TRL DA ÚLTIMA MACROENTREGA (ESPERADO NA CONCLUSÃO",
        "RESULTADOS ESPERADOS COM A CONCLUSÃO DO PROJETO (DESCRITIVO - MÁX DE 500 CARACTERES)",
        "RESULTADOS ESPERADOS COM A CONCLUSÃO DO PROJETO (DESCRITIVO -",
        "RESULTADOS ESPERADOS COM A CONCLUSÃO DO PROJETO",
        "(DESCRITIVO - MÁX DE 500 CARACTERES)",
        "MÁX DE 500 CARACTERES)",
        "A CONCLUSÃO DO PROJETO)",
        "CONCLUSÃO DO PROJETO)",
        "EXPECTATIVA DE TEMPO ESPERADO PARA QUE A TECNOLOGIA CHEGUE",
        "EXPECTATIVA DE TEMPO ESPERADO PARA QUE A TECNOLOGIA CHEGUE AO MERCADO (EM Nº DE MESES",
        "EXPECTATIVA DE TEMPO ESPERADO PARA QUE A TECNOLOGIA CHEGUE AO MERCADO (EM Nº DE MESES APÓS",
        "EXPECTATIVA DE TEMPO ESPERADO PARA QUE A TECNOLOGIA CHEGUE AO MERCADO (EM Nº DE MESES APÓS A",
        "EXPECTATIVA DE TEMPO ESPERADO PARA QUE A TECNOLOGIA CHEGUE AO MERCADO (EM Nº DE MESES APÓS A CONCLUSÃ",
        "EXPECTATIVA DE TEMPO ESPERADO PARA QUE A TECNOLOGIA CHEGUE AO MERCADO (EM Nº DE MESES",
        "EXPECTATIVA DE TEMPO ESPERADO PARA QUE A TECNOLOGIA CHEGUE AO MERCADO (EM Nº DE MESES APÓS A CONCLUSÃO",
        "EXPECTATIVA DE TEMPO ESPERADO PARA QUE A TECNOLOGIA CHEGUE AO MERCADO (EM Nº DE MESES APÓS A CONCLUSÃO DO PROJETO)",
        "EXPECTATIVA DE IMPACTO ESPERADO DA(S) TECNOLOGIA(S) QUE SERÁ(ÃO) DESENVOLVIDAS(S) - BAIXO,",
        "EXPECTATIVA DE IMPACTO ESPERADO DA(S) TECNOLOGIA(S) QUE SERÁ(ÃO) DESENVOLVIDAS(S) - BAIXO, MÉDIO",
        "EXPECTATIVA DE IMPACTO ESPERADO DA(S) TECNOLOGIA(S) QUE SERÁ(ÃO) DESENVOLVIDAS(S) - BAIXO, MÉDIO OU",
        "EXPECTATIVA DE IMPACTO ESPERADO DA(S) TECNOLOGIA(S) QUE SERÁ(ÃO) DESENVOLVIDAS(S) - BAIXO, MÉDIO OU ALTO/DISRUPTIVO",
        "EXPECTATIVA DE IMPACTO ESPERADO DA(S) TECNOLOGIA(S) QUE",
        "SERÁ(ÃO) DESENVOLVIDAS(S) - BAIXO, MÉDIO OU ALTO/DISRUPTIVO",
        "MÉDIO OU ALTO/DISRUPTIVO"
        "OU ALTO/DISRUPTIVO",
        "ALTO/DISRUPTIVO:",
        "QUAL É A EXPECTATIVA DE SIGNIFICÂNCIA DA(S) INOVAÇÃO(ÕES) QUE",
        "QUAL É A EXPECTATIVA DE SIGNIFICÂNCIA DA(S) INOVAÇÃO(ÕES) QUE SERÁ(ÃO) GERADA(S) NO PROJETO?",
        "O PROJETO IRÁ USAR RECURSOS DOS CONTRATOS SEBRAE",
        "MODALIDADE DE APORTE DO PROJETO",
        "AMPLIAR A GAMA DE BENS OU SERVIÇOS OFERTADOS",
        "AMPLIAR A PARTICIPAÇÃO DA EMPRESA NO MERCADO",
        "AUMENTAR A CAPACIDADE DE PRODUÇÃO OU DE PRESTAÇÃO DE SERVIÇOS",
        "AUMENTAR A CAPACIDADE DE PRODUÇÃO OU DE PRESTAÇÃO DE",
        "AUMENTAR A FLEXIBILIDADE DA PRODUÇÃO OU DA PRESTAÇÃO DE SERVIÇOS",
        "ENQUADRAR EM REGULAÇÕES E NORMAS-PADRÃO RELATIVAS AO MERCADO INTERNO OU EXTERNO",
        "MERCADO INTERNO OU EXTERNO",
        "MELHORAR A QUALIDADE DOS BENS OU SERVIÇOS",
        "PERMITIR ABRERTURA DE NOVOS MERCADOS",
        "PERMITIR ABERTURA DE NOVOS MERCADOS",
        "PERMITIR CONTROLAR ASPECTOS LIGADOS À SAÚDE E/OU À SEGURANÇA",
        "SEGURANÇA",
        "PERMITIR MANTER A PARTICIPAÇÃO DA EMPRESA NO MERCADO",
        "PERMITIR REDUZIR O IMPACTO SOBRE O MEIO AMBIENTE",
        "REDUZIR O CONSUMO DE ÁGUA",
        "REDUZIR O CONSUMO DE ENERGIA",
        "REDUZIR O CONSUMO DE MATÉRIAS-PRIMAS",
        "REDUZIR OS CUSTOS DE PRODUÇÃO OU DOS SERVIÇOS PRESTADOS",
        "REDUZIR OS CUSTOS DO TRABALHO",
        "SUBSTITUIR (TOTAL OU PARCIAL) MATÉRIAS-PRIMAS POR OUTRAS MENOS CONTAMINANTES OU PERIGOSAS",
        "MENOS CONTAMINANTES OU PERIGOSAS",
        "CONTAMINANTES OU PERIGOSAS",
        "PERIGOSAS",
        "SUBSTITUIR (TOTAL OU PARCIAL) MATÉRIAS-PRIMAS POR OUTRAS MENOS CONTAMINANTES OU",
        "SUBSTITUIR (TOTAL OU PARCIAL) MATÉRIAS-PRIMAS POR OUTRAS MENOS CONTAMINANTES OU ",
        "SUBSTITUIR (TOTAL OU PARCIAL) ENERGIA PROVENIENTE DE COMBUSTÍVEIS FÓSSEIS POR FONTES DE ENERGIA RENOVÁVEIS",
        "SUBSTITUIR (TOTAL OU PARCIAL) ENERGIA PROVENIENTE DE COMBUSTÍVEIS FÓSSEIS POR FONTES DE ENERGIA",
        "SUBSTITUIR (TOTAL OU PARCIAL) ENERGIA PROVENIENTE DE COMBUSTÍVEIS FÓSSEIS POR FONTES DE",
        "ENERGIA RENOVÁVEIS",
        "RENOVÁVEIS",
        "COMBUSTÍVEIS FÓSSEIS POR FONTES DE ENERGIA RENOVÁVEIS",
        "SUBSTITUIR (TOTAL OU PARCIAL) MATÉRIAS-PRIMAS POR OUTRAS MENOS",
        "REDUZIR RUÍDOS OU A CONTAMINAÇÃO DO SOLO, DA ÁGUA OU DO AR",
        "RECICLAGEM DE RESÍDUOS, ÁGUAS RESIDUAIS OU MATERIAIS PARA VENDA E/OU REUTILIZAÇÃO",
        "VENDA E/OU REUTILIZAÇÃO",
        "REDUÇÃO DA 'PEGADA' DE CO (PRODUÇÃO TOTAL DE CO) DE SUA EMPRESA"
    ]

    # Ordenar frases de interesse da maior para a menor para evitar sobreposição (ex: "UF DO CNPJ" antes de "CNPJ")
    frases_interesse = sorted(frases_interesse, key=lambda x: -len(x))

    # Frases que devem ser ignoradas
    frases_excluir = [
        "Assinatura do responsável pela Unidade EMBRAPII",
        "RESPOSTAS",
        "INFORMAÇÕES GERAIS",
        "FORMULÁRIO PARA SOLICITAÇÃO DE RESERVA RECURSOS", 
        "CONTRATO SEBRAE/EMBRAPII",
        "Página"
    ]

    # Lista para armazenar os dados
    coletado = []
    textos_pdf = {} 
    arquivo_n = 0

    for arquivo in os.listdir(STEP1):
        if arquivo.lower().endswith('.pdf'):
            arquivo_n += 1
            caminho_arquivo = os.path.join(STEP1, arquivo)
            nome_arquivo = arquivo
            doc = fitz.open(caminho_arquivo)

            textos_pdf[nome_arquivo] = []

            for pagina_num in range(len(doc)):
                pagina = doc.load_page(pagina_num)

                # Tenta extrair texto diretamente
                texto = pagina.get_text()
                if texto is None or len(texto.strip()) < 800:
                    # Muito pouco texto: faz OCR
                    print(f"OCR necessário na página {pagina_num+1} do arquivo {nome_arquivo}")
                    pix = pagina.get_pixmap(dpi=300)
                    img = Image.open(io.BytesIO(pix.tobytes()))
                    texto = pytesseract.image_to_string(img, lang='por')

                else:
                    print(f"Texto extraído normalmente da página {pagina_num+1} do arquivo {nome_arquivo}")

                textos_pdf[nome_arquivo].append(texto) 

                linhas = texto.split('\n')

                i = 0
                while i < len(linhas):
                    linha = linhas[i].strip()

                    if not linha or any(f.lower() in linha.lower() for f in frases_excluir):
                        i += 1
                        continue

                    for frase in frases_interesse:
                        if frase.lower() in linha.lower():
                            pergunta = frase
                            resto = linha.replace(frase, '').strip().lstrip(':').strip()

                            if not resto:
                                # Pega próxima linha
                                j = i + 1
                                while j < len(linhas):
                                    proxima = linhas[j].strip()
                                    if proxima:
                                        resto = proxima
                                        break
                                    j += 1
                                i = j - 1

                            coletado.append({
                                'arquivo': nome_arquivo,
                                'pagina': pagina_num + 1,
                                'pergunta': pergunta,
                                'resposta': resto
                            })
                            break
                    i += 1

    print(f"Arquivos lidos: {arquivo_n}")

    df = pd.DataFrame(coletado) if coletado else pd.DataFrame()

    if not df.empty:
        df['ticket_acompanhamento'] = df['arquivo'].str.extract(r'_(\d+)\.')
        df.to_excel(os.path.join(STEP2, "dados_extraidos.xlsx"), index=False)
        print("Extração concluída. Dados salvos em 'dados_extraidos.xlsx'")
    else:
        print("Nenhum dado coletado.")

    # 🔍 Identificar PDFs problemáticos
    textos_problematicos = {}

    for arquivo, textos in textos_pdf.items():
        qtd = df[df["arquivo"] == arquivo].shape[0] if not df.empty else 0
        if qtd < 20:
            textos_problematicos[arquivo] = "\n\n".join(textos)

    return textos_problematicos

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
O texto abaixo foi extraído de um PDF de formulário EMBRAPII.
O documento pode conter:
- Quebras de linha incorretas
- Texto desalinhado
- OCR com erros
- Perguntas e respostas em linhas separadas ou fora de ordem

As perguntas seguem o padrão dos formulários EMBRAPII, mas podem variar levemente na redação.

Sua tarefa:
1. Interpretar semanticamente o texto
2. Identificar perguntas e seus respectivos valores
3. Associar corretamente cada resposta à sua pergunta
4. Normalizar os dados conforme o schema abaixo
5. Retornar APENAS um JSON válido
6. Não inventar informações
7. Se houver ambiguidade ou ausência de resposta, usar null
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

{
  "projeto": "...",
  "empresas": [...],
  "financeiro": {...},
  "impactos": {...}
}
"""

    
    variables = {
        "texto": texto
    }

    try:
        formatted_prompt = prompt_base.format(**variables)
        return formatted_prompt
    except Exception as e:
        raise ValueError(f"Erro ao formatar o prompt: {str(e)}")

# Processa os PDFs problemáticos usando IA para interpretar o texto e extrair as informações
def ler_pdfs_ia(textos_problematicos, pasta_saida):

    resultados = []

    for arquivo, texto in textos_problematicos.items():
        print(f"Processando arquivo: {arquivo}")
        print(f"Texto extraído: {texto[:500]}...")  # Exibe os primeiros 500 caracteres para referência

        prompt = gerar_prompt(texto)

        resposta = requisicao(
            prompt=prompt,
            modelo='o4-mini'
        )

        if resposta is None:
            continue

        resposta["arquivo"] = arquivo
        resultados.append(resposta)

    # ===============================
    # JSON → DataFrame normalizado
    # ===============================
    df_final = pd.json_normalize(resultados, sep="_")
    df_final.to_excel(
        os.path.join(pasta_saida, "resultados_pdf_raw_ia.xlsx"),
        index=False
    )

    # ===============================
    # Explodir impactos
    # ===============================
    linhas = []

    for _, row in df_final.iterrows():
        arquivo = row["arquivo"]
        match = re.search(r"(\d+)", arquivo)
        ticket = match.group(1) if match else None

        for campo, pergunta in MAPA_PERGUNTAS.items():
            resposta = row.get(campo)

            if pd.notna(resposta):
                linhas.append({
                    "arquivo": arquivo,
                    "ticket_acompanhamento": ticket,
                    "pergunta": pergunta,
                    "resposta": resposta
                })

    df_impactos = pd.DataFrame(linhas)

    # ===============================
    # Impacto tecnologia
    # ===============================
    if "impacto_tecnologia" in df_final.columns:
        df_impacto_tecnologia = df_final[['arquivo', 'impacto_tecnologia']].copy()
        df_impacto_tecnologia['ticket_acompanhamento'] = df_impacto_tecnologia['arquivo'].str.extract(r"(\d+)")
        df_impacto_tecnologia['resposta'] = (
            df_impacto_tecnologia['impacto_tecnologia']
            .str.upper()
            .map(MAPA_IMPACTO_TECNOLOGIA)
        )

        df_impacto_tecnologia['pergunta'] = (
            'Expectativa de impacto esperado da(s) tecnologia(s) que será(ão) desenvolvida(s)'
        )

        df_impacto_tecnologia = df_impacto_tecnologia.dropna(subset=['resposta'])

        df_impactos = pd.concat(
            [df_impactos, df_impacto_tecnologia[['arquivo', 'ticket_acompanhamento', 'pergunta', 'resposta']]],
            ignore_index=True
        )

    caminho_saida = os.path.join(pasta_saida, "info_complementares_pdf_ia.xlsx")
    df_impactos.to_excel(caminho_saida, index=False)

    return df_final, df_impactos
