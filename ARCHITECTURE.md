# Arquitetura

Ferramenta de conferência de PIX. Cruza extratos bancários em PDF (Banco do Brasil,
C6 Bank) com planilhas Excel de controle por agente e reporta o que bateu e o que
está faltando de cada lado. Operador único (Gilmario Lima), rodada de forma avulsa ao
fim do turno. Sem banco de dados, sem estado, uma requisição por rodada.

Ver `PRODUCT.md` para contexto de produto e `DESIGN.md` para o design system
"Register Tape" (Fita de Caixa Registradora).

## Rodar

```bash
uvicorn servidor:app --reload        # dev, http://127.0.0.1:8000
docker build -t checker . && docker run -p 8080:8080 checker
```

`GET /` serve `frontend/leitor-extratos.html`; `/static` monta `frontend/`.

## Arquitetura

Backend é um arquivo só: `servidor.py` (FastAPI). Frontend é JS puro estático, sem build.

**Fluxo da requisição** — `POST /conferir_caixa`, multipart: `pdfs[]`, `excels[]`,
`data` opcional (ignorado, mantido por compatibilidade), `senha` (senha do PDF C6).

1. `processar_pdf` — detecta o banco por match de string no texto extraído e despacha:
   - `detalhe_bb` — parser ancorado por linha. Ancora em cada linha `Pix - Recebido` e
     lê as linhas vizinhas (linha de valor antes, linha de detalhe depois, ou uma
     linha combinada). NÃO é um regex único de texto corrido de propósito — a ordem
     de exportação do PDF varia e a abordagem antiga perdia registros.
   - `detalhe_c6` + `extrair_pix_c6` — regex sobre texto colapsado. Frágil, não
     verificado a fundo contra exportações reais do C6. Sem desbloqueio de PDF:
     arquivo protegido → erro pedindo PDF já desbloqueado.
2. `processar_excel` — pandas/openpyxl. Linhas contendo `AGENTE` definem o contexto de
   agente/setor para as linhas seguintes. Colunas de dados: `0` nome, `1` hora, `4` valor.
3. Matching (corpo de `conferir_caixa`) — monta uma matriz de pontuação Excel × PDF e
   resolve como problema de atribuição ótima via `scipy.optimize.linear_sum_assignment`
   (algoritmo húngaro). NÃO é guloso — a ordem das linhas do Excel não pode afetar o
   resultado.
   - Pontuação = similaridade de nome (`SequenceMatcher` sobre nomes sem acento +
     bônus/supressões por sobreposição de tokens) + proximidade de valor (corte rígido
     em diferença de R$0,50) + bônus de janela de horário (só quando os dois lados têm hora).
   - Aceita um par se `sim >= 0.70`, ou `sim >= 0.55` com valor exato. Uma diferença de
     horário acima de `GAP_HORA_MAX` (2h) veta um par que seria aceito.
4. Resposta: `conferidos`, `faltando_no_pdf` (no Excel, sem par no PDF — cada item
   carrega um `motivo` explicando o porquê), `faltando_no_excel` (linhas do PDF que
   ninguém consumiu).

**Frontend** (`frontend/scriptfinal.js`) — renderiza cards recolhíveis por agente,
anéis de progresso, exportação em PDF (`html2pdf`). Ajuste manual: marcar/desmarcar um
par, ou preencher uma "base manual" (nome + valor, com autocomplete puxado da lista de
"faltando no Excel" da própria página) para confirmar um par na mão.

## Arquivos

| Caminho | O que é |
|---------|---------|
| `servidor.py` | Backend inteiro: parsers + matching + rotas |
| `frontend/leitor-extratos.html` | Página única |
| `frontend/scriptfinal.js` | Toda a lógica do frontend (cache furado pelo `?v=` na tag script) |
| `frontend/style.css` | Estilos — segue `DESIGN.md` |
| `requirements.txt` / `Dockerfile` | Deploy |

## Convenções e restrições

- **Escopo travado em BB e C6.** Não adicionar outros bancos/formatos sem pedido.
- **Só pt-BR.** UI, mensagens, strings de `motivo` — sem i18n.
- **Conteúdo de extrato/planilha é sensível.** Sem nova persistência, sem log desse
  conteúdo, sem chamada externa com ele.
- **Atribuição "Gilmario Lima"** fica no rodapé da UI, no rodapé do PDF exportado e no
  `<meta name="author">`.
- Sem suíte de testes. Não adicionar frameworks de teste sem pedido.
- Mensagens de commit são curtas e genéricas.
- UI usa Bootstrap Icons, nunca emoji. IBM Plex Mono em todo número/data/hora/valor.

## Dívida técnica conhecida

- `detalhe_bb` grava todo o texto extraído do PDF em `pdf_debug.txt` a cada rodada, e
  os dois parsers dão `print` de cada nome/valor/hora no stdout. É anterior à restrição
  de privacidade acima — sinalizar, não tratar como normal. (`pdf_debug.txt` está no
  gitignore.)
- Precisão do parser C6 não verificada contra um conjunto real de amostras.
