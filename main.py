import requests

# Configurações
tenant_id = "<TENANT_ID>"
client_id = "<CLIENT_ID>"
client_secret = "<CLIENT_SECRET>"
site_id = '<SITE_ID>'
file_path = r'arquivo.png'
upload_path = '/Homol/arquivo1.png'

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

# Função para fazer upload do arquivo
def upload_file_to_sharepoint(access_token, site_id, file_path, upload_path):
    with open(file_path, 'rb') as file:
        content = file.read()
        
    upload_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:{upload_path}:/content'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'text/plain'  # Tipo de conteúdo para arquivo .txt
    }

    response = requests.put(upload_url, headers=headers, data=content)

    if response.status_code in (200, 201):
        print("Arquivo enviado com sucesso!")
        return response.json()  # Retorna os detalhes do arquivo enviado
    else:
        print("Erro ao enviar arquivo:", response.status_code)
        print("Response Body:", response.json())
        return None

# Função para Check-In do arquivo
def check_in_file(access_token, site_id, item_id):
    check_in_url = f'https://graph.microsoft.com/v1.0/sites/{site_id}/drive/items/{item_id}/checkin'
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    data = {
        "comment": "Checked in via API",
        "checkInType": "majorCheckIn"  # Ou "minorCheckIn" conforme necessário
    }

    response = requests.post(check_in_url, headers=headers, json=data)

    if response.status_code == 204:  # Verifica se o status é 204
        print("Arquivo check-in realizado com sucesso!")
    else:
        print("Erro ao realizar check-in:", response.status_code)
        try:
            print("Response Body:", response.json())
        except ValueError:  # Caso a resposta não seja um JSON válido
            print("Response Body não é um JSON válido ou está vazio.")

# Execução do script
access_token = get_access_token(tenant_id, client_id, client_secret)

if access_token:
    uploaded_file = upload_file_to_sharepoint(access_token, site_id, file_path, upload_path)
    if uploaded_file:
        # Realiza o Check-In com o ID do item retornado
        check_in_file(access_token, site_id, uploaded_file['id'])