# https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0&preserve-view=true
# pip install azure-identity
# pip install azure-keyvault-secrets
# pip install openai

import traceback
from decouple import config
from openai import AzureOpenAI
from autentica_azure.get_key import CofreChavesMS


class RequisicaoMS:
    tenant_id = config('MS_TENANT_ID')
    client_id = config('MS_CLIENT_ID')
    secret_value = config('MS_SECRET_VALUE')

    key_vault_name = config('AZURE_KEY_VAULT_NAME') # Cofre de chaves
    secret_name = config('AZURE_KEY_SECRET_NAME') # Chave a ser buscada

    def __init__(self):
        self.access_token = CofreChavesMS().get_access_token()

    def requisita(self, mensagem, model="o4-mini", contexto=""):
        try:

            endpoint = "https://azure-openai-equipe-dados.openai.azure.com/"
            subscription_key = self.access_token
            api_version = "2024-12-01-preview"

            client = AzureOpenAI(
                api_version=api_version,
                azure_endpoint=endpoint,
                api_key=subscription_key,
            )

            response = client.chat.completions.create(
                messages=[
                    {"role": "system", "content": contexto},
                    {"role": "user", "content": mensagem}
                ],
                max_completion_tokens=40000,
                model=model
            )

            return response.choices[0].message.content

        except Exception as e:
            print(f"Houve um erro na requisição: {e}")
            traceback.print_exc()
            return "erro"
        


if __name__ == "__main__":
    mensagem = "O que é a EMBRAPII?"
    contexto = """Você é um assistente virtual da EMBRAPII e só deve dar informações da EMBRAPII. 
        Para informações de outros órgãos ou empresas informe que não tem autorização"""
    modelo = "o4-mini"

    requisita = RequisicaoMS()

    response = requisita.requisita(mensagem=mensagem,
                                   contexto=contexto,
                                   model=modelo)

    print(response)
