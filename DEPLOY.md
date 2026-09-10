# Publicar na Vercel

1. Na Vercel, escolha **Add New → Project** e importe `gilmariolima/checker-full`.
2. Use a branch `main`, diretório raiz `./` e framework **FastAPI**.
3. Mantenha os comandos de build e saída nos padrões do framework. Não use o Dockerfile nem configure `uvicorn` como Build Command.
4. Clique em **Deploy**.

O arquivo `app.py` exporta o aplicativo de `servidor.py`. A versão Python está fixada em 3.12 por `.python-version`; as dependências ficam em `requirements.txt`. O frontend continua em `/` e seus recursos em `/static/`, com a API em `/conferir_caixa` no mesmo domínio. Não são necessárias variáveis de ambiente para o funcionamento atual.

A interface limita a soma dos arquivos a 4.000.000 bytes para reservar margem dentro do limite de 4,5 MB por requisição da Vercel. Essa checagem também se aplica ao uso local. A plataforma ainda pode rejeitar requisições por tamanho ou tempo; a interface apresenta uma mensagem para tentar novamente com menos arquivos.

O processamento não grava mais `pdf_debug.txt` nem imprime nomes, valores ou conteúdo dos extratos. Arquivos de debug, ambientes locais e testes são excluídos do upload pela `.vercelignore`.

Depois da publicação, valide com um PDF BB/C6 e uma planilha pequenos: upload, conferência, exibição dos setores e exportação PDF. O teste local não substitui a validação do build e da função na Vercel.

Documentação: https://vercel.com/docs/frameworks/backend/fastapi
