# PROMPT PARA IA: Sistema de Download de Anexos do Teams

## O QUE EU QUERO
Tenho um script Python que exporta chats privados do Microsoft Teams para Excel. Preciso que ele também baixe automaticamente os anexos (arquivos SharePoint) que estão nas conversas e organize tudo em uma estrutura de pastas limpa.

## ESTRUTURA DE PASTAS DESEJADA
```
exports/
└── usuario@empresa.com/
    ├── Conversa_1/
    │   ├── Conversa_1.xlsx
    │   └── attachments/
    │       ├── documento1.pdf
    │       ├── planilha2.xlsx
    │       └── anexos_Conversa_1.csv
    ├── Conversa_2/
    │   ├── Conversa_2.xlsx
    │   └── attachments/
    │       ├── arquivo3.docx
    │       └── anexos_Conversa_2.csv
    └── Conversa_3/
        ├── Conversa_3.xlsx
        └── attachments/
            └── anexos_Conversa_3.csv
```

## COMO DEVE FUNCIONAR

### 1. Organização
- Uma pasta para o email do usuário logado
- Dentro dela, uma pasta para cada conversa 
- Cada conversa tem: arquivo Excel + pasta "attachments"
- Na pasta attachments: arquivos baixados + CSV com lista completa

### 2. Tipos de Anexos
O Teams tem 3 tipos de anexos nas mensagens:
- **`attachment`** - Arquivos do SharePoint (PDFs, DOCs, etc) → **BAIXAR ESTES**
- **`hosted_image`** - Imagens incorporadas → **NÃO BAIXAR** (não funcionam)
- **`url_file`** - Links de arquivos no texto → **BAIXAR ESTES**

### 3. Downloads
- Baixar apenas arquivos reais (attachment e url_file)
- Ignorar as imagens incorporadas (hosted_image)
- Salvar os arquivos na pasta attachments/ de cada conversa
- Se o download direto falhar, tentar método alternativo via Graph API

### 4. CSV de Controle
Criar um CSV em cada pasta attachments/ com TODOS os anexos encontrados:
```csv
tipo,attachment_name,content_url,from_user,message_date,observacoes
attachment,documento.pdf,https://sharepoint.com/doc.pdf,João Silva,2024-01-15,Anexo estruturado - reference
hosted_image,image123.jpg,https://graph.microsoft.com/hostedContents/123,Maria Santos,2024-01-15,Imagem incorporada na mensagem
url_file,planilha.xlsx,https://sharepoint.com/plan.xlsx,João Silva,2024-01-16,URL de arquivo encontrada no conteúdo da mensagem
```

## CONFIGURAÇÃO
No arquivo .env:
```bash
EXPORT_ATTACHMENTS=true
EXPORT_ATTACHMENTS_MODE=both  # 'csv', 'download', ou 'both'
```

## RESULTADO ESPERADO
Quando executar o script:
1. Cria estrutura de pastas organizada por usuário/conversa
2. Exporta Excel de cada conversa individualmente  
3. Extrai todos os anexos de cada conversa
4. Baixa apenas os arquivos reais (SharePoint)
5. Gera CSV com lista completa dos anexos
6. Mostra estatísticas de sucessos/falhas

**OBJETIVO**: Ter cada conversa do Teams organizada em sua própria pasta, com Excel + arquivos baixados + CSV de controle, tudo organizadamente.