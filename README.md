# Teams Chat Exporter

🚀 **Exportador de Chats Privados do Microsoft Teams**

Uma aplicação Python para exportar conversas privadas (1:1 e grupos) do Microsoft Teams usando a Microsoft Graph API com autenticação Device Code Flow.

## 📋 Visão Geral

Esta ferramenta permite:
- ✅ Exportar **todos** os chats privados do Teams
- ✅ Incluir conversas 1:1 e grupos privados
- ✅ Salvar em formato **JSON** e **Excel**
- ✅ Funcionar em **WSL/Linux** sem problemas de navegador
- ✅ Modo **teste** para validação antes da exportação completa
- ✅ Rate limiting automático para evitar bloqueios da API
- ✅ Renovação automática de tokens

## 🎯 Características

- **Autenticação Segura**: Device Code Flow (sem necessidade de servidor web local)
- **Compatibilidade WSL**: Funciona perfeitamente no Windows Subsystem for Linux
- **Exportação Completa**: Inclui metadados, anexos, reações e menções
- **Formato Duplo**: JSON para dados brutos + Excel para análise
- **Rate Limiting**: Controle automático de velocidade das requisições
- **Modo Teste**: Validação com apenas o primeiro chat antes da exportação completa
- **Logs Detalhados**: Progresso em tempo real com estatísticas

## 🛠️ Tecnologias Utilizadas

- **Python 3.7+**
- **Microsoft Graph API v1.0**
- **MSAL (Microsoft Authentication Library)**
- **pandas** - Processamento de dados e Excel
- **requests** - HTTP requests
- **python-dotenv** - Gerenciamento de variáveis de ambiente

## 📦 Estrutura do Projeto

```
teams_chats_privados/
├── device_chat_exporter.py    # 🎯 Script principal
├── device_auth.py             # 🔐 Autenticação Device Code
├── config.py                  # ⚙️ Configurações
├── test_device_auth.py        # 🧪 Teste de autenticação  
├── requirements.txt           # 📚 Dependências Python
├── .env                       # 🔒 Variáveis de ambiente (criar)
├── exports/                   # 📁 Pasta de saída dos arquivos
└── README.md                  # 📖 Esta documentação
```

## 🚀 Instalação e Configuração

### 1. Pré-requisitos

- Python 3.7 ou superior
- Conta Microsoft com acesso ao Teams
- Aplicação registrada no Azure AD (veja seção de configuração)

### 2. Clonar e Instalar

```bash
# Clonar o repositório
cd sua_pasta_de_projetos
git clone <url_do_repo> teams_chats_privados
cd teams_chats_privados

# Criar ambiente virtual (recomendado)
python -m venv env
source env/bin/activate  # Linux/WSL
# ou
env\Scripts\activate     # Windows

# Instalar dependências
pip install -r requirements.txt
```

### 3. Configuração do Azure AD

#### 3.1 Registrar Aplicação no Azure Portal

1. Acesse [Azure Portal](https://portal.azure.com)
2. Vá em **Azure Active Directory** > **App registrations**
3. Clique **New registration**
4. Configure:
   - **Name**: Teams Chat Exporter
   - **Supported account types**: Accounts in this organizational directory only
   - **Redirect URI**: Deixe em branco para Device Code Flow
5. Clique **Register**

#### 3.2 Configurar Permissões

1. Na aplicação criada, vá em **API permissions**
2. Clique **Add a permission** > **Microsoft Graph** > **Delegated permissions**
3. Adicione as permissões:
   - `Chat.Read` - Ler chats do usuário
   - `Chat.ReadBasic` - Ler informações básicas dos chats
   - `User.Read` - Ler perfil do usuário
   - `offline_access` - Refresh token
4. Clique **Add permissions**
5. **IMPORTANTE**: Clique **Grant admin consent** (se você for admin)

#### 3.3 Habilitar Public Client

1. Vá em **Authentication**
2. Na seção **Advanced settings**
3. Marque **Yes** em "Allow public client flows"
4. Salve as alterações

#### 3.4 Obter IDs Necessários

1. Na página **Overview** da aplicação, copie:
   - **Application (client) ID**
   - **Directory (tenant) ID**

### 4. Configurar Variáveis de Ambiente

#### 4.1 Criar arquivo de configuração

O projeto inclui um arquivo `.env.example` com todas as configurações disponíveis:

```bash
# Copiar o template para seu arquivo de configuração
cp .env.example .env
```

#### 4.2 Preencher as configurações obrigatórias

Abra o arquivo `.env` e preencha **obrigatoriamente**:

```bash
# Valores obtidos no Azure Portal (seção 3.4)
CLIENT_ID=sua_application_client_id_aqui
TENANT_ID=seu_directory_tenant_id_aqui
```

#### 4.3 Configurações opcionais (já com valores padrão)

```bash
# Configurações já preenchidas no .env.example
MODE=test                    # 'test' ou 'prod' (padrão: prod)
OUTPUT_DIR=./exports         # Pasta de saída (padrão: ./exports)
EXPORT_FORMAT=json          # Formato adicional (padrão: json)
EXPORT_ATTACHMENTS=true     # Exportar anexos (padrão: true)
MAX_MESSAGES_PER_REQUEST=50 # Limite por requisição (padrão: 50)
```

**⚠️ IMPORTANTE**: 
- Nunca commite o arquivo `.env` para o git!
- O `.gitignore` já está configurado para ignorar o `.env`
- Use sempre o `.env.example` como referência

## 📖 Como Usar

### 1. Teste de Autenticação

Antes de tudo, teste se a autenticação está funcionando:

```bash
python test_device_auth.py
```

**O que acontece:**
1. Será exibido um código de device e um link
2. Acesse o link em qualquer dispositivo/navegador
3. Digite o código mostrado
4. Faça login com sua conta Microsoft
5. Se bem-sucedido, verá suas informações de usuário

### 2. Modo Teste (Recomendado Primeiro)

Configure `MODE=test` no arquivo `.env` e execute:

```bash
python device_chat_exporter.py
```

**Modo Teste exporta apenas o primeiro chat** para validar se tudo funciona corretamente.

### 3. Exportação Completa

Configure `MODE=prod` no arquivo `.env` e execute:

```bash
python device_chat_exporter.py
```

**Modo Produção exporta todos os chats** encontrados.

### 4. Exemplo de Execução

```bash
$ python device_chat_exporter.py

🚀 Exportador de Chats Privados - Device Code
👤 Usando autenticação device code
🧪 MODO: TESTE (apenas primeira conversa)
📁 Diretório de saída: ./exports

🔐 Iniciando autenticação via Device Code...
🖥️  Visite: https://microsoft.com/devicelogin
🔑 E digite o código: A1B2C3D4E
⏳ Aguardando autenticação...
✅ Autenticação bem-sucedida!
✅ Autenticado como: João Silva (joao.silva@empresa.com)

💬 Obtendo lista de chats privados...
✅ Encontrados 15 chats

🔄 Exportando 1 conversas...
============================================================

[ 1/ 1] ( 100.0%) 📨 1:1 com Maria Santos
           ✅   42 mensagens | Total acumulado:    42

============================================================
🎉 Exportação concluída! Total: 42 mensagens

💾 JSON salvo: ./exports/private_chats_test_20240125_143022.json
📊 Excel salvo: ./exports/private_chats_test_20240125_143022.xlsx

🎯 RESUMO FINAL
========================================
💬 Total de mensagens: 42
⏱️  Tempo total: 0:00:08.123456
📁 Arquivos salvos em: ./exports

📊 Mensagens por tipo:
   oneOnOne: 42 mensagens
```

## 📊 Formato dos Dados Exportados

### JSON Export
Dados brutos da API do Microsoft Graph com informações adicionais:

```json
{
  "id": "message_id",
  "createdDateTime": "2024-01-25T14:30:22Z",
  "lastModifiedDateTime": "2024-01-25T14:30:22Z",
  "messageType": "message",
  "body": {
    "content": "Conteúdo da mensagem...",
    "contentType": "html"
  },
  "from": {
    "user": {
      "displayName": "João Silva",
      "userPrincipalName": "joao.silva@empresa.com"
    }
  },
  "chatInfo": {
    "id": "chat_id",
    "topic": "Chat topic",
    "chatType": "oneOnOne"
  },
  "sourceType": "private_chat"
}
```

### Excel Export
Dados tabulares otimizados para análise:

| Coluna | Descrição |
|--------|-----------|
| `id` | ID único da mensagem |
| `createdDateTime` | Data/hora de criação |
| `messageType` | Tipo da mensagem |
| `body` | Conteúdo da mensagem |
| `from_displayName` | Nome do remetente |
| `from_email` | Email do remetente |
| `chat_type` | Tipo do chat (oneOnOne, group) |
| `chat_display` | Nome amigável do chat |
| `attachments_count` | Número de anexos |
| `reactions_count` | Número de reações |
| `mentions_count` | Número de menções |

## ⚙️ Configurações Avançadas

### Variáveis de Ambiente Disponíveis

```bash
# Obrigatórias
CLIENT_ID=                    # ID da aplicação Azure AD
TENANT_ID=                    # ID do tenant Azure AD

# Opcionais
MODE=prod                     # 'test' | 'prod'
OUTPUT_DIR=./exports          # Diretório de saída
EXPORT_FORMAT=json            # Formato adicional
EXPORT_ATTACHMENTS=true       # Exportar anexos
MAX_MESSAGES_PER_REQUEST=50   # Limite da API (max: 50)
```

### Modos de Operação

- **`MODE=test`**: Exporta apenas o primeiro chat (para testes)
- **`MODE=prod`**: Exporta todos os chats (padrão)

### Rate Limiting

A aplicação implementa rate limiting automático:
- Pausa entre chats: 200ms
- Pausa entre mensagens do mesmo chat: 100ms
- Retry automático em caso de HTTP 429
- Renovação automática de token em caso de HTTP 401

## 🔧 Troubleshooting

### Erro de Autenticação

**Problema**: "Erro na autenticação: invalid_client"
**Solução**:
1. Verifique se CLIENT_ID e TENANT_ID estão corretos no `.env`
2. Confirme se "Allow public client flows" está habilitado no Azure
3. Verifique se as permissões foram concedidas

### Erro de Permissões

**Problema**: "Insufficient privileges to complete the operation"
**Solução**:
1. Adicione as permissões necessárias no Azure Portal:
   - Chat.Read
   - Chat.ReadBasic  
   - User.Read
   - offline_access
2. Clique em "Grant admin consent" se você for administrador
3. Ou solicite ao admin para aprovar as permissões

### Rate Limiting

**Problema**: Muitos erros 429 (Too Many Requests)
**Solução**:
1. A aplicação já trata automaticamente
2. Se persistir, aumente as pausas no código
3. Execute em horários de menor uso

### WSL/Linux Issues

**Problema**: Browser não abre automaticamente
**Solução**:
1. Copie o link mostrado no terminal
2. Abra manualmente no navegador (Windows)
3. Digite o código de device mostrado

### Arquivo .env não encontrado

**Problema**: "CLIENT_ID e TENANT_ID são obrigatórios"
**Solução**:
1. Crie o arquivo `.env` na raiz do projeto
2. Adicione as variáveis obrigatórias
3. Reinicie a aplicação

## 🔒 Segurança e Privacidade

- **Tokens**: Armazenados apenas em memória durante a execução
- **Credenciais**: Nunca logadas ou salvas em arquivos
- **Device Code**: Expira automaticamente após alguns minutos
- **Permissões**: Apenas leitura, nunca modificação
- **Dados**: Salvos localmente, nunca enviados para terceiros

## 🤝 Contribuições

Contribuições são bem-vindas! Por favor:

1. Fork o projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

## 📄 Licença

Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.

## 🆘 Suporte

Para suporte e dúvidas:
1. Verifique a seção **Troubleshooting** acima
2. Abra uma **Issue** no GitHub com detalhes do erro
3. Inclua logs relevantes (sem informações sensíveis)

---

⭐ **Gostou do projeto? Deixe uma estrela no GitHub!**