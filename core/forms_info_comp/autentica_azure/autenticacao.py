# https://learn.microsoft.com/en-us/graph/api/overview?view=graph-rest-1.0&preserve-view=true

import requests
import traceback
from decouple import config


class AutenticaMS:
    def __init__(self):
        self.tenant_id = config('MS_TENANT_ID')
        self.client_id = config('MS_CLIENT_ID')
        self.secret_value = config('MS_SECRET_VALUE')

    def get_access_token(self):
        try:
            url_requisicao = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
            headers = {
                "Content-Type": f"application/x-www-form-urlencoded",
            }

            params = {
                "grant_type": f"client_credentials",
                "scope": "https://graph.microsoft.com/.default",
                "client_id": self.client_id,
                "client_secret": self.secret_value,
            }

            # print(url_requisicao)
            # print(headers)
            # print(params)

            response = requests.post(url=url_requisicao, headers=headers, data=params)
            try:
                dicionario = response.json()
                if "access_token" in dicionario:
                    access_token = dicionario["access_token"]
                    # print(f"Access token obtido com sucesso: {access_token}")
                else:
                    print("Erro na obtenção do access token")
                    print(response.json())
                    raise Exception(f"Erro na obtenção do access token: {response.json()}")

            except Exception as e:
                raise Exception(f"Erro na obtenção do access token: {e}")

            # print("Access token obtido com sucesso")
            return access_token

        except Exception as e:
            print(f"Houve um erro na obtenção do access token: {e}")
            traceback.print_exc()
            return None


if __name__ == "__main__":
    access_token = AutenticaMS().get_access_token()
