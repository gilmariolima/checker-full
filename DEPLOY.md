# Publicar na Vercel

1. Na Vercel, escolha **Add New → Project** e importe `gilmariolima/checker-full`.
2. Use a branch `main`, diretório raiz `./` e framework **FastAPI**.
3. Mantenha os comandos de build e saída nos padrões do framework. Não use o Dockerfile nem configure `uvicorn` como Build Command.
4. Configure as variáveis de ambiente (seção abaixo).
5. Clique em **Deploy**.

O arquivo `app.py` exporta o aplicativo de `servidor.py`. A versão Python está fixada em 3.12 por `.python-version`; as dependências ficam em `requirements.txt`. O frontend continua em `/` e seus recursos em `/static/`, com a API em `/conferir_caixa` no mesmo domínio.

## Variáveis de ambiente

O acesso (login) e a leitura do Google Sheets exigem cinco variáveis. Localmente elas ficam
no `.env` (ignorado pelo Git); na Vercel, em **Settings → Environment Variables**.

| Variável | Origem |
|----------|--------|
| `CHECKER_USERNAME` | usuário compartilhado dos donos |
| `CHECKER_PASSWORD_HASH` | hash PBKDF2 da senha (gerado, a senha não é gravada) |
| `CHECKER_SESSION_SECRET` | segredo aleatório da sessão (≥ 32 caracteres) |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | JSON da chave da conta de serviço, em **uma linha** |
| `GOOGLE_SHEETS_ID` | ID da planilha BASES (padrão já embutido) |

Gerar as três primeiras + embutir o JSON do Google:

```bash
python configurar_acesso.py
```

Ele lê `service-account.json` (se existir na raiz) e escreve o `.env` completo.
Para a Vercel, copie cada valor **sem as aspas externas** que aparecem no arquivo.

### Conta de serviço do Google

1. Google Cloud Console → projeto `checker-full` → ativar **Google Sheets API**.
2. Criar uma conta de serviço, sem papéis, e gerar uma chave **JSON**.
3. Salvar a chave como `service-account.json` na raiz do projeto (ignorada pelo Git).
4. **Compartilhar a planilha BASES** com o `client_email` da conta como **Leitor**.

Sem `GOOGLE_SERVICE_ACCOUNT_JSON` o app continua funcionando: o botão do Sheets
responde que a conexão não foi configurada e o upload de Excel segue disponível.

## Limites

A interface limita a soma dos arquivos a 4.000.000 bytes para reservar margem dentro do limite de 4,5 MB por requisição da Vercel. Essa checagem também se aplica ao uso local. A plataforma ainda pode rejeitar requisições por tamanho ou tempo; a interface apresenta uma mensagem para tentar novamente com menos arquivos.

O processamento não grava `pdf_debug.txt` nem imprime nomes, valores ou conteúdo dos extratos. Arquivos de debug, ambientes locais e testes são excluídos do upload pela `.vercelignore`.

Depois da publicação, valide com um PDF BB/C6 e uma planilha pequenos: login, upload, conferência, exibição dos setores e exportação PDF. O teste local não substitui a validação do build e da função na Vercel.

Documentação: https://vercel.com/docs/frameworks/backend/fastapi
