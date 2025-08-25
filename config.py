import os
from dotenv import load_dotenv

load_dotenv()

CLIENT_ID = os.getenv('CLIENT_ID')
TENANT_ID = os.getenv('TENANT_ID')
MODE = os.getenv('MODE', 'prod').lower()  # 'test' ou 'prod'

OUTPUT_DIR = os.getenv('OUTPUT_DIR', './exports')
EXPORT_FORMAT = os.getenv('EXPORT_FORMAT', 'json')
EXPORT_ATTACHMENTS = os.getenv('EXPORT_ATTACHMENTS', 'true').lower() == 'true'
MAX_MESSAGES_PER_REQUEST = int(os.getenv('MAX_MESSAGES_PER_REQUEST', '50'))  # Limite da API para chats

GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0'
REDIRECT_URI = 'http://localhost:8080'

# Scopes para autenticação delegada
SCOPES = [
    'https://graph.microsoft.com/Chat.Read',
    'https://graph.microsoft.com/Chat.ReadBasic',
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/offline_access'
]