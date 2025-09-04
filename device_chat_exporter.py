#!/usr/bin/env python3

import os
import json
import requests
import pandas as pd
from datetime import datetime
import time
import re
from urllib.parse import urlparse, unquote
import csv

from device_auth import DeviceCodeAuthenticator
import config

class DeviceChatExporter:
    def __init__(self):
        self.authenticator = DeviceCodeAuthenticator()
        self.headers = self.authenticator.get_headers()
        self.base_output_dir = config.OUTPUT_DIR
        self.user_email = None
        self.user_output_dir = None
    
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
    
    def get_user_info(self):
        """Obter informações do usuário conectado"""
        print("👤 Obtendo informações do usuário...")
        url = f"{config.GRAPH_ENDPOINT}/me"
        
        data = self.make_request(url)
        if data:
            self.user_email = data.get('userPrincipalName', data.get('mail', 'usuario_desconhecido'))
            display_name = data.get('displayName', 'Usuário')
            print(f"✅ Conectado como: {display_name} ({self.user_email})")
            
            # Configurar diretório do usuário
            safe_email = re.sub(r'[<>:"/\\|?*]', '_', self.user_email)
            self.user_output_dir = os.path.join(self.base_output_dir, safe_email)
            
            if not os.path.exists(self.user_output_dir):
                os.makedirs(self.user_output_dir)
                
            return True
        else:
            print("❌ Não foi possível obter informações do usuário")
            return False
    
    def ensure_chat_directory(self, chat_display):
        """Criar diretório para uma conversa específica"""
        safe_chat_name = self.sanitize_filename(chat_display)
        chat_dir = os.path.join(self.user_output_dir, safe_chat_name)
        attachments_dir = os.path.join(chat_dir, "attachments")
        
        if not os.path.exists(chat_dir):
            os.makedirs(chat_dir)
        if not os.path.exists(attachments_dir):
            os.makedirs(attachments_dir)
            
        return chat_dir, attachments_dir
    
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
    
    
    def save_chat_to_excel(self, messages, chat_dir, chat_display):
        """Salvar mensagens de um chat específico em Excel"""
        if not messages:
            return None
            
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
        safe_filename = self.sanitize_filename(chat_display)
        filepath = os.path.join(chat_dir, f"{safe_filename}.xlsx")
        df.to_excel(filepath, index=False)
        print(f"📊 Excel salvo: {os.path.basename(filepath)}")
        return filepath
    
    def sanitize_filename(self, filename):
        """Remove caracteres inválidos do nome do arquivo"""
        sanitized = re.sub(r'[<>:"/\\|?*]', '_', filename)
        if len(sanitized) > 200:
            name, ext = os.path.splitext(sanitized)
            sanitized = name[:190] + ext
        return sanitized
    
    def extract_owner_from_url(self, url):
        """Extrair o dono do arquivo da URL do SharePoint"""
        if not url:
            return 'desconhecido'
        
        # Procurar padrão: personal/usuario_dominio_com_br/
        personal_match = re.search(r'personal/([^/]+)/', url)
        if personal_match:
            user_encoded = personal_match.group(1)
            # Decodificar: ti01_camozziconsultoria_com_br -> ti01@camozziconsultoria.com.br
            user_decoded = user_encoded.replace('_', '.').replace('.', '@', 1)
            return user_decoded
        
        return 'sistema'
    
    def download_file(self, url, filename, attachments_dir):
        """Baixar arquivo de uma URL"""
        try:
            response = requests.get(url, headers=self.headers, stream=True)
            
            if response.status_code == 200:
                filepath = os.path.join(attachments_dir, self.sanitize_filename(filename))
                
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
                print(f"         ❌ ERRO HTTP {response.status_code} para {filename}")
                print(f"         📄 URL: {url[:100]}...")
                if response.text:
                    print(f"         📝 Resposta: {response.text[:200]}")
                return False, None
                
        except Exception as e:
            print(f"         ❌ EXCEÇÃO para {filename}: {str(e)}")
            print(f"         📄 URL: {url[:100]}...")
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
    
    def process_sharepoint_attachment(self, attachment, attachments_dir):
        """Processar anexos do SharePoint"""
        content_url = attachment.get('contentUrl', '')
        if content_url:
            parsed_url = urlparse(content_url)
            filename = unquote(os.path.basename(parsed_url.path))
            
            if not filename or filename == '/' or filename == '':
                attachment_name = attachment.get('name', 'arquivo_sem_nome')
                filename = f"{attachment_name}_{attachment.get('id', 'unknown')}"
                if not filename.endswith('.pdf'):
                    filename += '.pdf'
            
            # Tentar 1: Via Graph API como item compartilhado (PRIMEIRO - funciona melhor)
            if 'sharepoint.com' in content_url:
                print(f"         🔄 Tentando via Graph API shared items...")
                import base64
                encoded_url = base64.b64encode(content_url.encode()).decode().rstrip('=')
                shared_url = f"{config.GRAPH_ENDPOINT}/shares/u!{encoded_url}/driveItem/content"
                success, filepath = self.download_file(shared_url, filename, attachments_dir)
            else:
                success = False
            
            if not success:
                print(f"         🔄 Tentando URL direta...")
                # Tentar 2: URL direta
                success, filepath = self.download_file(content_url, filename, attachments_dir)
                
                if not success and 'sharepoint.com' in content_url:
                    print(f"         🔄 Tentando via Graph API me/drive...")
                    # Tentar 3: Via Graph API tradicional
                    graph_url = f"{config.GRAPH_ENDPOINT}/me/drive/root:{parsed_url.path}:/content"
                    success, filepath = self.download_file(graph_url, filename, attachments_dir)
            
            return success, filepath
        else:
            print(f"         ❌ Anexo sem contentUrl: {attachment}")
        
        return False, None
    
    def download_chat_attachments(self, messages, attachments_dir):
        """Baixar anexos de um chat específico"""
        downloaded = 0
        failed = 0
        
        for message in messages:
            # 1. Processar anexos estruturados (SharePoint files)
            attachments = message.get('attachments', [])
            for attachment in attachments:
                content_type = attachment.get('contentType', '')
                
                if content_type == 'reference':
                    success, filepath = self.process_sharepoint_attachment(attachment, attachments_dir)
                    if success:
                        downloaded += 1
                        print(f"         ✅ Baixado: {os.path.basename(filepath)}")
                    else:
                        failed += 1
                        print(f"         ❌ Falha ao baixar anexo SharePoint: {attachment.get('name', 'sem nome')}")
            
            # 2. Procurar por URLs de arquivo no conteúdo da mensagem
            body_content = message.get('body', {}).get('content', '')
            if body_content:
                # Procurar por URLs que apontam para arquivos reais (não imagens)
                file_patterns = [
                    r'https://[^"]*\.(?:pdf|doc|docx|xls|xlsx|ppt|pptx|zip|rar|txt)',
                    r'https://[^"]*sharepoint[^"]*',
                    r'https://[^"]*onedrive[^"]*'
                ]
                
                for pattern in file_patterns:
                    file_matches = re.findall(pattern, body_content, re.IGNORECASE)
                    for match in file_matches:
                        if 'hostedContents' not in match:  # Evitar imagens incorporadas
                            filename_from_url = unquote(os.path.basename(urlparse(match).path))
                            if not filename_from_url:
                                filename_from_url = 'arquivo_extraido.file'
                                
                            success, filepath = self.download_file(match, filename_from_url, attachments_dir)
                            if success:
                                downloaded += 1
                                print(f"         ✅ URL baixada: {os.path.basename(filepath)}")
                            else:
                                failed += 1
                                print(f"         ❌ Falha ao baixar URL: {filename_from_url}")
        
        return downloaded, failed
    
    def extract_chat_attachments_info(self, messages):
        """Extrair informações dos anexos de um chat específico para CSV"""
        attachments_info = []
        
        for message in messages:
            chat_info = message.get('chatInfo', {})
            chat_display = self.format_chat_info(chat_info)
            message_id = message.get('id', '')
            message_date = message.get('createdDateTime', '')
            from_info = message.get('from')
            if from_info and isinstance(from_info, dict):
                user_info = from_info.get('user')
                if user_info and isinstance(user_info, dict):
                    from_user = user_info.get('displayName', 'Usuário desconhecido')
                else:
                    from_user = 'Usuário desconhecido'
            else:
                from_user = 'Usuário desconhecido'
            
            # 1. Anexos estruturados (arquivos do SharePoint, etc.)
            attachments = message.get('attachments', [])
            for idx, attachment in enumerate(attachments):
                attachment_info = {
                    'tipo': 'attachment',
                    'message_id': message_id,
                    'message_date': message_date,
                    'from_user': from_user,
                    'attachment_id': attachment.get('id', f'att_{idx}'),
                    'attachment_name': attachment.get('name', 'Sem nome'),
                    'content_type': attachment.get('contentType', ''),
                    'content_url': attachment.get('contentUrl', ''),
                    'web_url': attachment.get('contentUrl', ''),
                    'size_bytes': attachment.get('size', ''),
                    'observacoes': f"Anexo estruturado - {attachment.get('contentType', 'tipo desconhecido')} - Dono: {self.extract_owner_from_url(attachment.get('contentUrl', ''))}"
                }
                attachments_info.append(attachment_info)
            
            # 2. Imagens incorporadas (hostedContents)
            body_content = message.get('body', {}).get('content', '')
            if body_content:
                image_urls = self.extract_hosted_content_urls(body_content)
                
                for url, filename in image_urls:
                    attachment_info = {
                        'tipo': 'hosted_image',
                        'message_id': message_id,
                        'message_date': message_date,
                        'from_user': from_user,
                        'attachment_id': filename,
                        'attachment_name': filename,
                        'content_type': 'image',
                        'content_url': url,
                        'web_url': url,
                        'size_bytes': '',
                        'observacoes': 'Imagem incorporada na mensagem'
                    }
                    attachments_info.append(attachment_info)
                
                # 3. Outros links de arquivo no conteúdo (opcional)
                # Procurar por outros padrões de URL que possam ser arquivos
                file_patterns = [
                    r'https://[^"]*\.(?:pdf|doc|docx|xls|xlsx|ppt|pptx|zip|rar|txt)',
                    r'https://[^"]*sharepoint[^"]*',
                    r'https://[^"]*onedrive[^"]*'
                ]
                
                for pattern in file_patterns:
                    file_matches = re.findall(pattern, body_content, re.IGNORECASE)
                    for match in file_matches:
                        if 'hostedContents' not in match:  # Evitar duplicatas das imagens já processadas
                            filename_from_url = unquote(os.path.basename(urlparse(match).path))
                            if not filename_from_url:
                                filename_from_url = 'arquivo_extraido.file'
                                
                            attachment_info = {
                                'tipo': 'url_file',
                                'message_id': message_id,
                                'message_date': message_date,
                                'from_user': from_user,
                                'attachment_id': f'url_{len(attachments_info)}',
                                'attachment_name': filename_from_url,
                                'content_type': 'file_url',
                                'content_url': match,
                                'web_url': match,
                                'size_bytes': '',
                                'observacoes': 'URL de arquivo encontrada no conteúdo da mensagem'
                            }
                            attachments_info.append(attachment_info)
        
        return attachments_info
    
    def save_chat_attachments_to_csv(self, attachments_info, chat_dir, chat_display):
        """Salvar informações dos anexos de um chat específico em CSV"""
        if not attachments_info:
            return None
        
        safe_filename = self.sanitize_filename(chat_display)
        filepath = os.path.join(chat_dir, f"anexos_{safe_filename}.csv")
        
        fieldnames = [
            'tipo',
            'message_date',
            'from_user',
            'attachment_name',
            'content_type',
            'content_url',
            'web_url',
            'size_bytes',
            'observacoes',
            'message_id',
            'attachment_id'
        ]
        
        with open(filepath, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(attachments_info)
        
        print(f"📋 CSV anexos: {os.path.basename(filepath)} ({len(attachments_info)} itens)")
        return filepath
    
    def export_all(self):
        print("🚀 Exportador de Chats Privados - Device Code")
        print("👤 Usando autenticação device code")
        
        # Mostrar modo atual
        if config.MODE == 'test':
            print("🧪 MODO: TESTE (apenas primeira conversa)")
        else:
            print("🚀 MODO: PRODUÇÃO (todas as conversas)")
        
        # Mostrar modo de anexos
        if config.EXPORT_ATTACHMENTS:
            if config.EXPORT_ATTACHMENTS_MODE == 'csv':
                print("📋 ANEXOS: Modo CSV (lista de links para download manual)")
            elif config.EXPORT_ATTACHMENTS_MODE == 'download':
                print("📥 ANEXOS: Modo download automático")
            elif config.EXPORT_ATTACHMENTS_MODE == 'both':
                print("📋📥 ANEXOS: Modo completo (CSV + download automático)")
            else:
                print("📥 ANEXOS: Modo download automático (padrão)")
        else:
            print("❌ ANEXOS: Desabilitado")
            
        print(f"📁 Diretório base: {self.base_output_dir}")
        
        start_time = datetime.now()
        
        # Testar autenticação
        success, auth_message = self.authenticator.test_connection()
        if not success:
            print(f"❌ {auth_message}")
            return
        
        print(f"✅ {auth_message}")
        
        # Obter informações do usuário
        if not self.get_user_info():
            return
        
        # Obter lista de chats
        chats = self.get_my_chats()
        
        if not chats:
            print("⚠️  Nenhum chat encontrado")
            return
        
        # Modo teste: apenas primeira conversa
        if config.MODE == 'test':
            chats = chats[:1]
            print(f"🧪 MODO TESTE: Processando apenas 1 conversa (primeira)")
        else:
            print(f"🚀 MODO PRODUÇÃO: Processando todas as {len(chats)} conversas")
        
        print(f"\n🔄 Exportando {len(chats)} conversas...")
        print("=" * 60)
        
        total_messages = 0
        total_attachments = 0
        exported_chats = 0
        
        for i, chat in enumerate(chats, 1):
            chat_display = self.format_chat_info(chat)
            
            # Mostrar progresso atual
            progress_percent = (i / len(chats)) * 100
            print(f"\n[{i:2d}/{len(chats)}] ({progress_percent:5.1f}%) 📨 {chat_display}")
            
            # Criar diretórios para este chat
            chat_dir, attachments_dir = self.ensure_chat_directory(chat_display)
            
            # Exportar mensagens do chat
            chat_messages = self.get_messages_from_chat(chat['id'], chat)
            
            if chat_messages:
                # Salvar Excel do chat
                self.save_chat_to_excel(chat_messages, chat_dir, chat_display)
                
                # Processar anexos se habilitado
                if config.EXPORT_ATTACHMENTS:
                    attachments_info = self.extract_chat_attachments_info(chat_messages)
                    downloaded_count = 0
                    failed_count = 0
                    
                    # Salvar CSV com informações dos anexos (sempre que há anexos)
                    if attachments_info and config.EXPORT_ATTACHMENTS_MODE in ['csv', 'both']:
                        self.save_chat_attachments_to_csv(attachments_info, chat_dir, chat_display)
                    
                    # Fazer download dos anexos se modo download ou both
                    if config.EXPORT_ATTACHMENTS_MODE in ['download', 'both']:
                        downloaded_count, failed_count = self.download_chat_attachments(chat_messages, attachments_dir)
                    
                    if attachments_info:
                        total_attachments += len(attachments_info)
                        if downloaded_count > 0:
                            print(f"           📥 {downloaded_count} arquivos baixados, {failed_count} falhas")
                    else:
                        print(f"           ℹ️  Sem anexos")
                
                total_messages += len(chat_messages)
                exported_chats += 1
                print(f"           ✅ {len(chat_messages):4d} mensagens exportadas")
            else:
                print(f"           ℹ️   0 mensagens encontradas")
            
            # Pequena pausa para rate limiting
            time.sleep(0.2)
        
        end_time = datetime.now()
        duration = end_time - start_time
        
        print("\n" + "=" * 60)
        print(f"🎉 Exportação concluída!")
        print(f"\n🎯 RESUMO FINAL")
        print("=" * 40)
        print(f"👤 Usuário: {self.user_email}")
        print(f"📨 Conversas exportadas: {exported_chats:,}")
        print(f"💬 Total de mensagens: {total_messages:,}")
        
        if config.EXPORT_ATTACHMENTS and total_attachments > 0:
            print(f"📋 Total de anexos catalogados: {total_attachments:,}")
            if config.EXPORT_ATTACHMENTS_MODE == 'csv':
                print(f"ℹ️  Use os CSVs nas pastas para baixar anexos manualmente")
            elif config.EXPORT_ATTACHMENTS_MODE == 'download':
                print(f"ℹ️  Arquivos baixados automaticamente nas pastas attachments/")
            elif config.EXPORT_ATTACHMENTS_MODE == 'both':
                print(f"ℹ️  CSVs criados + arquivos baixados automaticamente")
        
        print(f"⏱️  Tempo total: {duration}")
        print(f"📁 Estrutura criada em: {self.user_output_dir}")
        print(f"\n📂 Estrutura:")
        print(f"   {self.user_email}/")
        print(f"   ├── [conversa1]/")
        print(f"   │   ├── conversa1.xlsx")
        if config.EXPORT_ATTACHMENTS:
            if config.EXPORT_ATTACHMENTS_MODE in ['csv', 'both']:
                print(f"   │   ├── anexos_conversa1.csv")
                if config.EXPORT_ATTACHMENTS_MODE == 'both':
                    print(f"   │   └── attachments/")
                    print(f"   │       ├── documento1.pdf")
                    print(f"   │       └── arquivo2.xlsx")
            elif config.EXPORT_ATTACHMENTS_MODE == 'download':
                print(f"   │   └── attachments/")
                print(f"   │       ├── documento1.pdf")
                print(f"   │       └── arquivo2.xlsx")
        print(f"   └── [conversa2]/...")
        
        # Calcular estatísticas dos chats exportados
        if exported_chats > 0:
            print(f"\n📊 Estatísticas:")
            print(f"   Média de mensagens por conversa: {total_messages/exported_chats:.0f}")
            if config.EXPORT_ATTACHMENTS and total_attachments > 0:
                print(f"   Média de anexos por conversa: {total_attachments/exported_chats:.0f}")

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