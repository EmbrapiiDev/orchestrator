# https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0&preserve-view=true
# pip install azure-identity
# pip install azure-keyvault-secrets
# pip install openai

import traceback
from decouple import config
from autentica_azure.autenticacao import AutenticaMS
from azure.keyvault.secrets import SecretClient
from azure.identity import ClientSecretCredential


class CofreChavesMS:
    tenant_id = config('MS_TENANT_ID')
    client_id = config('MS_CLIENT_ID')
    secret_value = config('MS_SECRET_VALUE')

    key_vault_name = config('AZURE_KEY_VAULT_NAME') # Cofre de chaves
    secret_name = config('AZURE_KEY_SECRET_NAME') # Chave a ser buscada

    def __init__(self):
        self.access_token = AutenticaMS().get_access_token()
        self.KVUri = f"https://{self.key_vault_name}.vault.azure.net"

    def instancia_cliente(self):

        credential = ClientSecretCredential(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.secret_value
        )
        client = SecretClient(vault_url=self.KVUri, credential=credential)

        return client

    def get_access_token(self):
        try:
            client = self.instancia_cliente()

            retrieved_secret = client.get_secret(self.secret_name).value

            return retrieved_secret

        except Exception as e:
            print(f"Houve um erro na requisição da chave: {e}")
            traceback.print_exc()
            return "erro"


if __name__ == "__main__":
    cofre = CofreChavesMS()
    access_token = cofre.get_access_token()
    # print(f"Access Token: {access_token}")
