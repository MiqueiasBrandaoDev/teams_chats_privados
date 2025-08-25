#!/usr/bin/env python3

import os
import json
import requests
import re
from urllib.parse import urlparse, unquote
from datetime import datetime
import time
from device_auth import DeviceCodeAuthenticator
import config

class AttachmentDownloader:
    def __init__(self):
        self.authenticator = DeviceCodeAuthenticator()
        self.headers = self.authenticator.get_headers()
        self.output_dir = os.path.join(config.OUTPUT_DIR, "attachments")
        self.ensure_output_directory()
        
    def ensure_output_directory(self):
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
    
    def sanitize_filename(self, filename):
        """Remove caracteres inválidos do nome do arquivo"""
        # Remover caracteres inválidos para nomes de arquivo
        sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename)
        # Limitar tamanho do nome
        if len(sanitized) > 200:
            name, ext = os.path.splitext(sanitized)
            sanitized = name[:190] + ext
        return sanitized
    
    def download_file(self, url, filename, chat_id=None):
        """Baixar arquivo de uma URL"""
        try:
            # Fazer request para baixar o arquivo
            response = requests.get(url, headers=self.headers, stream=True)
            
            if response.status_code == 200:
                # Criar subdiretório por chat se fornecido
                if chat_id:
                    chat_dir = os.path.join(self.output_dir, self.sanitize_filename(chat_id))
                    if not os.path.exists(chat_dir):
                        os.makedirs(chat_dir)
                    filepath = os.path.join(chat_dir, self.sanitize_filename(filename))
                else:
                    filepath = os.path.join(self.output_dir, self.sanitize_filename(filename))
                
                # Evitar sobrescrita - adicionar número se já existir
                counter = 1
                original_filepath = filepath
                while os.path.exists(filepath):
                    name, ext = os.path.splitext(original_filepath)
                    filepath = f"{name}_{counter}{ext}"
                    counter += 1
                
                # Salvar arquivo
                with open(filepath, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                
                print(f"✅ Baixado: {os.path.basename(filepath)}")
                return True, filepath
            else:
                print(f"❌ Erro ao baixar {filename}: HTTP {response.status_code}")
                return False, None
                
        except Exception as e:
            print(f"❌ Erro ao baixar {filename}: {str(e)}")
            return False, None
    
    def extract_hosted_content_urls(self, message_content):
        """Extrair URLs de imagens do conteúdo HTML da mensagem"""
        urls = []
        # Regex para encontrar URLs de imagens do Graph API
        pattern = r'https://graph\.microsoft\.com/v1\.0/chats/[^"]+/hostedContents/[^"]+/\$value'
        matches = re.findall(pattern, message_content)
        
        for match in matches:
            # Extrair ID do item da URL para criar nome do arquivo
            item_match = re.search(r'itemid="([^"]+)"', message_content)
            if item_match:
                item_id = item_match.group(1)
                filename = f"image_{item_id}.jpg"  # Assumir JPG por padrão
            else:
                # Fallback: usar parte da URL como nome
                url_parts = match.split('/')
                filename = f"image_{url_parts[-3]}.jpg"
            
            urls.append((match, filename))
        
        return urls
    
    def process_sharepoint_attachment(self, attachment, chat_id):
        """Processar anexos do SharePoint"""
        content_url = attachment.get('contentUrl', '')
        if content_url and 'sharepoint.com' in content_url:
            # Extrair nome do arquivo da URL
            parsed_url = urlparse(content_url)
            filename = unquote(os.path.basename(parsed_url.path))
            
            if not filename or filename == '/':
                filename = f"sharepoint_file_{attachment['id']}.pdf"
            
            print(f"📁 Tentando baixar arquivo SharePoint: {filename}")
            
            # Tentar baixar diretamente
            success, filepath = self.download_file(content_url, filename, chat_id)
            
            if not success:
                # Se falhou, tentar através da Graph API
                graph_url = f"{config.GRAPH_ENDPOINT}/me/drive/root:{parsed_url.path}:/content"
                success, filepath = self.download_file(graph_url, filename, chat_id)
            
            return success, filepath
        
        return False, None
    
    def download_attachments_from_messages(self, json_file_path):
        """Baixar todos os anexos de um arquivo JSON de mensagens"""
        print(f"📥 Carregando mensagens de: {json_file_path}")
        
        try:
            with open(json_file_path, 'r', encoding='utf-8') as f:
                messages = json.load(f)
        except Exception as e:
            print(f"❌ Erro ao carregar JSON: {str(e)}")
            return
        
        print(f"📊 Encontradas {len(messages)} mensagens para processar")
        
        total_downloaded = 0
        total_failed = 0
        
        for i, message in enumerate(messages, 1):
            chat_info = message.get('chatInfo', {})
            chat_id = chat_info.get('topic', f"chat_{chat_info.get('id', 'unknown')}")
            
            print(f"\n[{i}/{len(messages)}] Processando mensagem em: {chat_id}")
            
            # 1. Processar anexos estruturados
            attachments = message.get('attachments', [])
            for attachment in attachments:
                content_type = attachment.get('contentType', '')
                
                if content_type == 'reference':
                    # Arquivo do SharePoint ou OneDrive
                    success, filepath = self.process_sharepoint_attachment(attachment, chat_id)
                    if success:
                        total_downloaded += 1
                    else:
                        total_failed += 1
                
                elif content_type == 'messageReference':
                    # Referência a outra mensagem - pular por enquanto
                    continue
            
            # 2. Extrair imagens do conteúdo HTML
            body_content = message.get('body', {}).get('content', '')
            if body_content:
                image_urls = self.extract_hosted_content_urls(body_content)
                
                for url, filename in image_urls:
                    print(f"🖼️ Baixando imagem: {filename}")
                    success, filepath = self.download_file(url, filename, chat_id)
                    if success:
                        total_downloaded += 1
                    else:
                        total_failed += 1
            
            # Pausa para evitar rate limiting
            if i % 10 == 0:
                time.sleep(0.5)
        
        print(f"\n🎯 RESULTADO FINAL:")
        print(f"✅ Arquivos baixados: {total_downloaded}")
        print(f"❌ Falhas: {total_failed}")
        print(f"📁 Diretório: {self.output_dir}")
    
    def list_available_exports(self):
        """Listar arquivos JSON disponíveis para download"""
        json_files = []
        exports_dir = config.OUTPUT_DIR
        
        if os.path.exists(exports_dir):
            for file in os.listdir(exports_dir):
                if file.endswith('.json') and 'private_chats' in file:
                    filepath = os.path.join(exports_dir, file)
                    json_files.append(filepath)
        
        return sorted(json_files, reverse=True)  # Mais recentes primeiro

def main():
    try:
        downloader = AttachmentDownloader()
        
        print("🚀 DOWNLOADER DE ANEXOS DO TEAMS")
        print("=" * 50)
        
        # Testar autenticação
        success, auth_message = downloader.authenticator.test_connection()
        if not success:
            print(f"❌ {auth_message}")
            return
        
        print(f"✅ {auth_message}")
        
        # Listar exports disponíveis
        json_files = downloader.list_available_exports()
        
        if not json_files:
            print("❌ Nenhum arquivo de export encontrado!")
            print("   Execute primeiro: python device_chat_exporter.py")
            return
        
        print(f"\n📂 Arquivos de export disponíveis:")
        for i, filepath in enumerate(json_files, 1):
            filename = os.path.basename(filepath)
            file_size = os.path.getsize(filepath) / 1024 / 1024  # MB
            print(f"   {i}. {filename} ({file_size:.1f} MB)")
        
        # Usar o arquivo mais recente por padrão
        selected_file = json_files[0]
        print(f"\n🎯 Usando arquivo mais recente: {os.path.basename(selected_file)}")
        
        # Baixar anexos
        start_time = datetime.now()
        downloader.download_attachments_from_messages(selected_file)
        end_time = datetime.now()
        duration = end_time - start_time
        
        print(f"\n⏱️ Tempo total: {duration}")
        
    except KeyboardInterrupt:
        print("\n⚠️ Download interrompido pelo usuário")
    except Exception as e:
        print(f"❌ Erro durante o download: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()