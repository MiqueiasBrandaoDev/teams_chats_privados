#!/usr/bin/env python3

import os
import json
import requests
import pandas as pd
from datetime import datetime
import time

from device_auth import DeviceCodeAuthenticator
import config

class DeviceChatExporter:
    def __init__(self):
        self.authenticator = DeviceCodeAuthenticator()
        self.headers = self.authenticator.get_headers()
        self.output_dir = config.OUTPUT_DIR
        self.ensure_output_directory()
        
    def ensure_output_directory(self):
        if not os.path.exists(self.output_dir):
            os.makedirs(self.output_dir)
    
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
            processed_msg = {
                'id': msg.get('id'),
                'createdDateTime': msg.get('createdDateTime'),
                'lastModifiedDateTime': msg.get('lastModifiedDateTime'),
                'messageType': msg.get('messageType'),
                'importance': msg.get('importance'),
                'subject': msg.get('subject', ''),
                'body': msg.get('body', {}).get('content', ''),
                'body_contentType': msg.get('body', {}).get('contentType', ''),
                'from_displayName': msg.get('from', {}).get('user', {}).get('displayName', ''),
                'from_email': msg.get('from', {}).get('user', {}).get('userPrincipalName', ''),
                'attachments_count': len(msg.get('attachments', [])),
                'reactions_count': len(msg.get('reactions', [])),
                'mentions_count': len(msg.get('mentions', []))
            }
            
            # Informações do chat
            chat_info = msg.get('chatInfo', {})
            processed_msg.update({
                'chat_id': chat_info.get('id'),
                'chat_topic': chat_info.get('topic'),
                'chat_type': chat_info.get('chatType'),
                'chat_display': self.format_chat_info(chat_info)
            })
            
            processed_messages.append(processed_msg)
        
        df = pd.DataFrame(processed_messages)
        filepath = os.path.join(self.output_dir, f"{filename}.xlsx")
        df.to_excel(filepath, index=False)
        print(f"📊 Excel salvo: {filepath}")
    
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
        
        end_time = datetime.now()
        duration = end_time - start_time
        
        print(f"\n🎯 RESUMO FINAL")
        print("=" * 40)
        print(f"💬 Total de mensagens: {len(messages):,}")
        print(f"⏱️  Tempo total: {duration}")
        print(f"📁 Arquivos salvos em: {self.output_dir}")
        
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