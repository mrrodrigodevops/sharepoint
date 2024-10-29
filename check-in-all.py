import requests

# Configurações
tenant_id = "<TENANT_ID>"
client_id = "<CLIENT_ID>"
client_secret = "<CLIENT_SECRET>"
site_id = '<SITE_ID>'
library_path = '/Homol'  # Caminho da biblioteca no SharePoint

# Função para obter token de acesso
def get_access_token(tenant_id, client_id, client_secret):
    token_url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
    data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'
    }
    response = requests.post(token_url, data=data)
    
    if response.status_code == 200:
        return response.json().get('access_token')
    else:
        print("Erro ao obter token:", response.json())
        return None

# Função para listar itens na biblioteca
def list_items(access_token, site_id, library_path):
    items_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{library_path}:/children'
    headers = {
        'Authorization': f'Bearer {access_token}'
    }

    response = requests.get(items_url, headers=headers)

    if response.status_code == 200:
        return response.json().get('value', [])
    else:
        print("Erro ao listar itens:", response.status_code)
        print("Response Body:", response.json())
        return []

# Função para realizar check-out
def checkout_file(access_token, site_id, item_id):
    checkout_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/checkout'
    headers = {
        'Authorization': f'Bearer {access_token}'
    }

    response = requests.post(checkout_url, headers=headers)

    if response.status_code in (200, 201, 204):
        print("Arquivo check-out realizado com sucesso!")
    else:
        print("Erro ao realizar check-out:", response.status_code)
        print("Response Body:", response.json())

# Função para realizar check-in
def check_in_file(access_token, site_id, item_id):
    check_in_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/checkin'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    data = {
        "comment": "Checked in via API",
        "checkInType": "majorCheckIn"  # Ou "minorCheckIn"
    }

    response = requests.post(check_in_url, headers=headers, json=data)

    if response.status_code == 204:  # Verifica se o status é 204
        print(f"Arquivo {item_id} check-in realizado com sucesso!")
    else:
        print(f"Erro ao realizar check-in para {item_id}:", response.status_code)
        try:
            print("Response Body:", response.json())
        except ValueError:
            print("Response Body não é um JSON válido ou está vazio.")

# Execução do script
access_token = get_access_token(tenant_id, client_id, client_secret)

if access_token:
    # Listar itens na biblioteca
    items = list_items(access_token, site_id, library_path)
    
    for item in items:
        if item.get('folder') is None:  # Verifica se não é uma pasta
            item_id = item['id']
            print(f"Processando arquivo: {item['name']}")
            
            # Realiza o check-out do arquivo
            checkout_file(access_token, site_id, item_id)
            
            # Se precisar, faça o upload de um novo arquivo aqui
            # upload_file_to_sharepoint(access_token, site_id, local_file_path, upload_path)
            
            # Realiza o check-in do arquivo
            check_in_file(access_token, site_id, item_id)