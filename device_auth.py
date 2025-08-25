import msal
import requests
import time
import config

class DeviceCodeAuthenticator:
    def __init__(self):
        self.client_id = config.CLIENT_ID
        self.tenant_id = config.TENANT_ID
        self.scopes = config.SCOPES
        self.access_token = None
        
        if not all([self.client_id, self.tenant_id]):
            raise ValueError("CLIENT_ID e TENANT_ID são obrigatórios no arquivo .env")
        
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.app = msal.PublicClientApplication(
            self.client_id,
            authority=self.authority
        )
    
    def get_access_token(self):
        if self.access_token:
            return self.access_token
            
        return self._device_code_login()
    
    def _device_code_login(self):
        print("🔐 Iniciando autenticação via Device Code...")
        
        # Iniciar fluxo device code
        device_flow = self.app.initiate_device_flow(scopes=self.scopes)
        
        if "user_code" not in device_flow:
            raise Exception("Falha ao iniciar device flow")
        
        # Mostrar instruções para o usuário
        print("\n" + "="*60)
        print("📱 INSTRUÇÕES DE AUTENTICAÇÃO")
        print("="*60)
        print(f"1. Abra seu navegador e vá para: {device_flow['verification_uri']}")
        print(f"2. Digite o código: {device_flow['user_code']}")
        print("3. Faça login com sua conta Microsoft")
        print("4. Aguarde... (não feche este terminal)")
        print("="*60)
        
        # Aguardar autenticação
        result = self.app.acquire_token_by_device_flow(device_flow)
        
        if "access_token" in result:
            self.access_token = result["access_token"]
            print("✅ Autenticação bem-sucedida!")
            return self.access_token
        else:
            error = result.get("error_description", result.get("error", "Erro desconhecido"))
            raise Exception(f"Falha na autenticação: {error}")
    
    def get_headers(self):
        token = self.get_access_token()
        return {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
    
    def test_connection(self):
        try:
            headers = self.get_headers()
            response = requests.get(f"{config.GRAPH_ENDPOINT}/me", headers=headers)
            
            if response.status_code == 200:
                user_data = response.json()
                display_name = user_data.get('displayName', 'Usuário')
                email = user_data.get('userPrincipalName', 'N/A')
                return True, f"Autenticado como: {display_name} ({email})"
            else:
                return False, f"Falha na autenticação: {response.status_code} - {response.text}"
        except Exception as e:
            return False, f"Erro no teste: {str(e)}"