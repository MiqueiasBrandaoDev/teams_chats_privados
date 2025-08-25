#!/usr/bin/env python3

import os
import json
import requests
import pandas as pd
from datetime import datetime
import time
import re
from urllib.parse import urlparse, unquote

from device_auth import DeviceCodeAuthenticator
import config

class DeviceChatExporter:
    def __init__(self):
        self.authenticator = DeviceCodeAuthenticator()
        self.headers = self.authenticator.get_headers()
        self.output_dir = config.OUTPUT_DIR
        self.attachments_dir = os.path.join(self.output_dir, "attachments")
        self.ensure_output_directory()
        
    def ensure_output_directory(self):
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
        if not os.path.exists(self.attachments_dir):
            os.makedirs(self.attachments_dir)
    
    def make_request(self, url, params=None):
        try:
            response = requests.get(url, headers=self.headers, params=params)
            
            if response.status_code == 429:
                retry_after = int(response.headers.get('Retry-After', 60))
                print(f"⏳ Rate limit atingido. Aguardando {retry_after} segundos...")
                time.sleep(retry_after)
                return self.make_request(url, params)
            
            if response.status_code == 401:
                print("🔄 Token expirado. Renovando...")
                self.authenticator.access_token = None
                self.headers = self.authenticator.get_headers()
                return self.make_request(url, params)
            
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.RequestException as e:
            print(f"❌ Erro na requisição: {str(e)}")
            if hasattr(e, 'response') and e.response is not None:
                print(f"   Status: {e.response.status_code}")
                print(f"   Resposta: {e.response.text[:200]}")
            return None
    
    def get_my_chats(self):
        print("💬 Obtendo lista de chats privados...")
        url = f"{config.GRAPH_ENDPOINT}/me/chats"
        params = {
            '$expand': 'members',
            '$top': 50  # Limite da API para chats
        }
        
        chats = []
        while url:
            data = self.make_request(url, params)
            if not data:
                break
                
            chats.extend(data.get('value', []))
            url = data.get('@odata.nextLink')
            params = None
            
        print(f"✅ Encontrados {len(chats)} chats")
        return chats
    
    def get_messages_from_chat(self, chat_id, chat_info):
        url = f"{config.GRAPH_ENDPOINT}/me/chats/{chat_id}/messages"
        params = {
            '$top': 50,  # API limit para messages é 50
            '$orderby': 'createdDateTime desc'
        }
        
        messages = []
        while url:
            data = self.make_request(url, params)
            if not data:
                break
                
            batch_messages = data.get('value', [])
            for msg in batch_messages:
                msg['chatInfo'] = chat_info
                msg['sourceType'] = 'private_chat'
                
            messages.extend(batch_messages)
            url = data.get('@odata.nextLink')
            params = None
            
            # Para evitar rate limits em chats com muitas mensagens
            if len(messages) > 0:
                time.sleep(0.1)
                
        return messages
    
    def format_chat_info(self, chat):
        """Formatar informações do chat para exibição"""
        chat_type = chat.get('chatType', 'unknown')
        topic = chat.get('topic', 'Sem título')
        
        members = []
        for member in chat.get('members', []):
            user = member.get('displayName', 'Usuário desconhecido')
            members.append(user)
        
        if chat_type == 'oneOnOne':
            # Para chats 1:1, mostrar o outro participante
            other_members = [m for m in members if m != 'Eu']  # Filtrar o próprio usuário se necessário
            if other_members:
                return f"1:1 com {other_members[0]}"
            else:
                return f"1:1 ({topic})"
        elif chat_type == 'group':
            return f"Grupo: {topic} ({len(members)} membros)"
        else:
            return f"{chat_type}: {topic}"
    
    def export_private_chats(self):
        print("\n📱 Iniciando exportação de chats privados...")
        chats = self.get_my_chats()
        
        if not chats:
            print("⚠️  Nenhum chat encontrado")
            return []
        
        # Modo teste: apenas primeira conversa
        if config.MODE == 'test':
            chats = chats[:1]
            print(f"🧪 MODO TESTE: Processando apenas 1 conversa (primeira)")
        else:
            print(f"🚀 MODO PRODUÇÃO: Processando todas as {len(chats)} conversas")
        
        all_messages = []
        
        print(f"\n🔄 Exportando {len(chats)} conversas...")
        print("=" * 60)
        
        for i, chat in enumerate(chats, 1):
            chat_display = self.format_chat_info(chat)
            
            # Mostrar progresso atual
            progress_percent = (i / len(chats)) * 100
            print(f"\n[{i:2d}/{len(chats)}] ({progress_percent:5.1f}%) 📨 {chat_display}")
            
            # Exportar mensagens do chat
            chat_messages = self.get_messages_from_chat(chat['id'], chat)
            all_messages.extend(chat_messages)
            
            # Resultado da exportação
            if chat_messages:
                print(f"           ✅ {len(chat_messages):4d} mensagens | Total acumulado: {len(all_messages):5d}")
            else:
                print(f"           ℹ️   0 mensagens | Total acumulado: {len(all_messages):5d}")
            
            # Pequena pausa para rate limiting
            time.sleep(0.2)
        
        print("\n" + "=" * 60)
        print(f"🎉 Exportação concluída! Total: {len(all_messages)} mensagens")
        
        return all_messages
    
    def save_to_json(self, data, filename):
        filepath = os.path.join(self.output_dir, f"{filename}.json")
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2, default=str)
        print(f"💾 JSON salvo: {filepath}")
    
    def save_to_excel(self, messages, filename):
        if not messages:
            return
            
        processed_messages = []
        for msg in messages:
            # Tratamento seguro do campo 'from' que pode ser None
            from_info = msg.get('from') or {}
            user_info = from_info.get('user') or {}
            
            # Tratamento seguro do campo 'body' que pode ser None
            body_info = msg.get('body') or {}
            
            processed_msg = {
                'id': msg.get('id'),
                'createdDateTime': msg.get('createdDateTime'),
                'lastModifiedDateTime': msg.get('lastModifiedDateTime'),
                'messageType': msg.get('messageType'),
                'importance': msg.get('importance'),
                'subject': msg.get('subject', ''),
                'body': body_info.get('content', ''),
                'body_contentType': body_info.get('contentType', ''),
                'from_displayName': user_info.get('displayName', ''),
                'from_email': user_info.get('userPrincipalName', ''),
                'attachments_count': len(msg.get('attachments', [])),
                'reactions_count': len(msg.get('reactions', [])),
                'mentions_count': len(msg.get('mentions', []))
            }
            
            # Informações do chat - tratamento seguro
            chat_info = msg.get('chatInfo') or {}
            processed_msg.update({
                'chat_id': chat_info.get('id', ''),
                'chat_topic': chat_info.get('topic', ''),
                'chat_type': chat_info.get('chatType', 'unknown'),
                'chat_display': self.format_chat_info(chat_info)
            })
            
            processed_messages.append(processed_msg)
        
        df = pd.DataFrame(processed_messages)
        filepath = os.path.join(self.output_dir, f"{filename}.xlsx")
        df.to_excel(filepath, index=False)
        print(f"📊 Excel salvo: {filepath}")
    
    def sanitize_filename(self, filename):
        """Remove caracteres inválidos do nome do arquivo"""
        sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename)
        if len(sanitized) > 200:
            name, ext = os.path.splitext(sanitized)
            sanitized = name[:190] + ext
        return sanitized
    
    def download_file(self, url, filename, chat_id=None):
        """Baixar arquivo de uma URL"""
        try:
            response = requests.get(url, headers=self.headers, stream=True)
            
            if response.status_code == 200:
                if chat_id:
                    chat_dir = os.path.join(self.attachments_dir, self.sanitize_filename(chat_id))
                    if not os.path.exists(chat_dir):
                        os.makedirs(chat_dir)
                    filepath = os.path.join(chat_dir, self.sanitize_filename(filename))
                else:
                    filepath = os.path.join(self.attachments_dir, self.sanitize_filename(filename))
                
                counter = 1
                original_filepath = filepath
                while os.path.exists(filepath):
                    name, ext = os.path.splitext(original_filepath)
                    filepath = f"{name}_{counter}{ext}"
                    counter += 1
                
                with open(filepath, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                
                return True, filepath
            else:
                return False, None
                
        except Exception as e:
            return False, None
    
    def extract_hosted_content_urls(self, message_content):
        """Extrair URLs de imagens do conteúdo HTML da mensagem"""
        urls = []
        pattern = r'https://graph\.microsoft\.com/v1\.0/chats/[^"]+/hostedContents/[^"]+/\$value'
        matches = re.findall(pattern, message_content)
        
        for match in matches:
            item_match = re.search(r'itemid="([^"]+)"', message_content)
            if item_match:
                item_id = item_match.group(1)
                filename = f"image_{item_id}.jpg"
            else:
                url_parts = match.split('/')
                filename = f"image_{url_parts[-3]}.jpg"
            
            urls.append((match, filename))
        
        return urls
    
    def process_sharepoint_attachment(self, attachment, chat_id):
        """Processar anexos do SharePoint"""
        content_url = attachment.get('contentUrl', '')
        if content_url and 'sharepoint.com' in content_url:
            parsed_url = urlparse(content_url)
            filename = unquote(os.path.basename(parsed_url.path))
            
            if not filename or filename == '/':
                filename = f"sharepoint_file_{attachment['id']}.pdf"
            
            success, filepath = self.download_file(content_url, filename, chat_id)
            
            if not success:
                graph_url = f"{config.GRAPH_ENDPOINT}/me/drive/root:{parsed_url.path}:/content"
                success, filepath = self.download_file(graph_url, filename, chat_id)
            
            return success, filepath
        
        return False, None
    
    def download_attachments_from_messages(self, messages):
        """Baixar todos os anexos das mensagens"""
        print(f"\n📥 Iniciando download de anexos...")
        
        total_downloaded = 0
        total_failed = 0
        
        for message in messages:
            chat_info = message.get('chatInfo', {})
            chat_id = chat_info.get('topic', f"chat_{chat_info.get('id', 'unknown')}")
            
            # 1. Processar anexos estruturados
            attachments = message.get('attachments', [])
            for attachment in attachments:
                content_type = attachment.get('contentType', '')
                
                if content_type == 'reference':
                    success, filepath = self.process_sharepoint_attachment(attachment, chat_id)
                    if success:
                        total_downloaded += 1
                        print(f"✅ Baixado: {os.path.basename(filepath)}")
                    else:
                        total_failed += 1
            
            # 2. Extrair imagens do conteúdo HTML
            body_content = message.get('body', {}).get('content', '')
            if body_content:
                image_urls = self.extract_hosted_content_urls(body_content)
                
                for url, filename in image_urls:
                    success, filepath = self.download_file(url, filename, chat_id)
                    if success:
                        total_downloaded += 1
                        print(f"✅ Imagem baixada: {os.path.basename(filepath)}")
                    else:
                        total_failed += 1
        
        print(f"\n📁 ANEXOS BAIXADOS:")
        print(f"✅ Arquivos baixados: {total_downloaded}")
        print(f"❌ Falhas: {total_failed}")
        print(f"📂 Diretório: {self.attachments_dir}")
        
        return total_downloaded, total_failed
    
    def export_all(self):
        print("🚀 Exportador de Chats Privados - Device Code")
        print("👤 Usando autenticação device code")
        
        # Mostrar modo atual
        if config.MODE == 'test':
            print("🧪 MODO: TESTE (apenas primeira conversa)")
        else:
            print("🚀 MODO: PRODUÇÃO (todas as conversas)")
            
        print(f"📁 Diretório de saída: {self.output_dir}")
        
        start_time = datetime.now()
        
        # Testar autenticação
        success, auth_message = self.authenticator.test_connection()
        if not success:
            print(f"❌ {auth_message}")
            return
        
        print(f"✅ {auth_message}")
        
        # Exportar chats (não duplicar - export_private_chats já chama get_my_chats)
        messages = self.export_private_chats()
        
        if not messages:
            print("\n⚠️  Nenhuma mensagem encontrada para exportar")
            return
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Salvar dados com sufixo do modo
        mode_suffix = "test" if config.MODE == 'test' else "prod"
        self.save_to_json(messages, f"private_chats_{mode_suffix}_{timestamp}")
        self.save_to_excel(messages, f"private_chats_{mode_suffix}_{timestamp}")
        
        # Baixar anexos automaticamente
        if config.EXPORT_ATTACHMENTS:
            downloaded, failed = self.download_attachments_from_messages(messages)
        else:
            print("\n⚠️  Download de anexos desabilitado (EXPORT_ATTACHMENTS=false)")
            downloaded, failed = 0, 0
        
        end_time = datetime.now()
        duration = end_time - start_time
        
        print(f"\n🎯 RESUMO FINAL")
        print("=" * 40)
        print(f"💬 Total de mensagens: {len(messages):,}")
        if config.EXPORT_ATTACHMENTS:
            print(f"📎 Anexos baixados: {downloaded:,}")
            print(f"❌ Falhas no download: {failed:,}")
        print(f"⏱️  Tempo total: {duration}")
        print(f"📁 Arquivos salvos em: {self.output_dir}")
        if config.EXPORT_ATTACHMENTS and downloaded > 0:
            print(f"📂 Anexos salvos em: {self.attachments_dir}")
        
        # Calcular estatísticas dos chats exportados
        chat_types = {}
        for msg in messages:
            chat_type = msg.get('chatInfo', {}).get('chatType', 'unknown')
            chat_types[chat_type] = chat_types.get(chat_type, 0) + 1
        
        print(f"\n📊 Mensagens por tipo:")
        for chat_type, count in chat_types.items():
            print(f"   {chat_type}: {count:,} mensagens")

def main():
    try:
        exporter = DeviceChatExporter()
        exporter.export_all()
    except KeyboardInterrupt:
        print("\n⚠️  Exportação interrompida pelo usuário")
    except Exception as e:
        print(f"❌ Erro durante a exportação: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()