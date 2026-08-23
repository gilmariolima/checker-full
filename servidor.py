from fastapi import FastAPI, File, UploadFile, Form
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import io, re, pdfplumber
from datetime import datetime, date
from difflib import SequenceMatcher
import unicodedata
import numpy as np
from scipy.optimize import linear_sum_assignment
from typing import List
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

# ==========================================================
# 🚀 Configuração principal
# ==========================================================
app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==========================================================
# 🌐 Servir o Frontend (HTML, CSS, JS e ícone)
# ==========================================================
app.mount("/static", StaticFiles(directory="frontend"), name="static")

@app.get("/")
def home():
    return FileResponse("frontend/leitor-extratos.html")


# ==========================================================
# 📅 Funções auxiliares para lidar com datas
# ==========================================================
def try_parse_date(s: str) -> date | None:
    """Tenta converter várias strings de data comuns em date."""
    if not s or not isinstance(s, str):
        return None
    s = s.strip()
    patterns = ["%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y"]
    for p in patterns:
        try:
            return datetime.strptime(s, p).date()
        except Exception:
            pass
    m = re.search(r"(\d{2}/\d{2}/\d{4})", s)
    if m:
        try:
            return datetime.strptime(m.group(1), "%d/%m/%Y").date()
        except Exception:
            pass
    return None


# ==========================================================
# 🔹 Função auxiliar para parsear valores numéricos
# ==========================================================
def parse_valor_robusto(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        try:
            return round(float(v), 2)
        except:
            return 0.0
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none", "-"):
        return 0.0
    s = s.replace("R$", "").replace("r$", "").replace(" ", "")
    if re.match(r'^\d{1,3}(\.\d{3})+,\d{1,}$', s):
        s_num = s.replace('.', '').replace(',', '.')
        return round(float(s_num), 2)
    if re.match(r'^\d+,\d+$', s):
        return round(float(s.replace(',', '.')), 2)
    if re.match(r'^\d+\.\d+$', s):
        return round(float(s), 2)
    if re.match(r'^\d{1,3}(\.\d{3})+$', s):
        return round(float(s.replace('.', '')), 2)
    if re.match(r'^\d+$', s):
        return round(float(s), 2)
    cleaned = re.sub(r'[^\d\.,\-]', '', s)
    if '.' in cleaned and ',' in cleaned:
        return round(float(cleaned.replace('.', '').replace(',', '.')), 2)
    if ',' in cleaned:
        return round(float(cleaned.replace(',', '.')), 2)
    if '.' in cleaned:
        return round(float(cleaned), 2)
    fallback = re.sub(r'[^\d\.\-]', '', cleaned)
    return round(float(fallback or 0.0), 2)


# ==========================================================
# 🟡 DETALHE BANCO DO BRASIL (versão final consolidada)
# ==========================================================
async def detalhe_bb(file_bytes: bytes):
    """
    Parser linha-a-linha para extratos do Banco do Brasil.

    O extrato lista cada PIX recebido em 2 ou 3 linhas, mas a ORDEM dessas
    linhas varia conforme a exportação do PDF (já vimos os dois casos):

    Formato A (valor antes do rótulo):
        13/07/2026 14397 130921147643521 1.055,74 (+)
        Pix - Recebido
        13/07 09:21 08123628374 Gilmario de Li

    Formato B (tudo depois do rótulo, numa linha só):
        Pix - Recebido
        13/07/2026 14397 130921147643521 13/07 09:21 08123628374 Gilmario de Li 1.055,74 (+)

    Por isso ancoramos a busca em cada ocorrência de "Pix - Recebido" e
    olhamos para as linhas vizinhas (antes e depois) em vez de assumir uma
    ordem fixa via regex de texto corrido — o formato antigo, ao colapsar
    tudo numa única linha, perdia registros sempre que a ordem real do PDF
    não era exatamente a esperada.
    """
    print("\n========== [DEBUG] INÍCIO DA LEITURA PDF BANCO DO BRASIL ==========\n")

    try:
        paginas_texto = []
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                paginas_texto.append(page.extract_text() or "")
        texto_total = "\n".join(paginas_texto)

        with open("pdf_debug.txt", "w", encoding="utf-8") as f:
            f.write(texto_total)

    except Exception as e:
        print(f"\n⚠️ Erro ao ler PDF: {e}")
        return {"erro": f"Falha ao processar PDF ({e})"}

    print("\n========== [DEBUG] PARSE LINHA A LINHA ==========\n")

    linhas = [l.strip() for l in texto_total.split("\n") if l.strip()]
    n = len(linhas)

    texto_join = " ".join(linhas)
    ano_padrao = datetime.now().year
    ano_match = re.search(r"Per[íi]odo:\s*\d{2}\s+a\s+\d{2}/\d{2}/(\d{4})", texto_join, re.IGNORECASE)
    if ano_match:
        ano_padrao = int(ano_match.group(1))

    re_label_recebido = re.compile(r"^Pix\s*-\s*Recebido\b", re.IGNORECASE)
    re_valor_row = re.compile(r"^(\d{2}/\d{2}/\d{4})\b.*?([\d\.]+,\d{2})\s*\(\+\)\s*$")
    re_detail_row = re.compile(r"^(\d{2}/\d{2})\s+(\d{2}:\d{2})\s+[0-9]{5,}\s+(.+?)\s*$")
    re_combined_row = re.compile(
        r"^(?:(\d{2}/\d{2}/\d{4})\s+)?(?:\d+\s+\d+\s+)?"
        r"(\d{2}/\d{2})\s+(\d{2}:\d{2})\s+[0-9]{5,}\s+"
        r"(.+?)\s+([\d\.]+,\d{2})\s*\(\+\)\s*$"
    )

    dados = []

    for i, linha in enumerate(linhas):
        if not re_label_recebido.match(linha):
            continue

        data_linha = hora = nome_raw = valor_txt = None

        # Formato A: valor numa das 2 linhas anteriores, detalhe numa das 2 seguintes
        for back in (1, 2):
            if i - back < 0:
                continue
            m_val = re_valor_row.match(linhas[i - back])
            if m_val:
                data_linha, valor_txt = m_val.groups()
                break

        for fwd in (1, 2):
            if i + fwd >= n:
                continue
            m_det = re_detail_row.match(linhas[i + fwd])
            if m_det:
                _, hora, nome_raw = m_det.groups()
                break

        # Formato B (fallback): tudo combinado numa das linhas seguintes
        if not (hora and valor_txt):
            for fwd in (1, 2):
                if i + fwd >= n:
                    continue
                m_comb = re_combined_row.match(linhas[i + fwd])
                if m_comb:
                    data_full, data_curta, hora, nome_raw, valor_txt = m_comb.groups()
                    data_linha = data_full or f"{data_curta}/{ano_padrao}"
                    break

        if not (hora and valor_txt and nome_raw):
            print(f"⚠️ [linha {i}] PIX - Recebido sem par válido de valor/detalhe nas vizinhanças, ignorado.")
            continue

        nome = re.sub(r"\s{2,}", " ", nome_raw.strip()).title()
        nome = re.sub(r"(?i)\b(Agência|Conta|Saldo|Pix)\b.*", "", nome).strip()
        if not nome:
            continue

        try:
            valor = float(valor_txt.replace(".", "").replace(",", "."))
        except Exception:
            continue
        if valor <= 0:
            continue

        nome_limpo = nome.strip()
        if re.fullmatch(r"[0-9.\s]{7,}", nome_limpo):
            cnpj_num = re.sub(r"\D", "", nome_limpo)
            nome_limpo = f"Cliente CNPJ {cnpj_num}"
        nome = nome_limpo

        dados.append({
            "data": data_linha or f"??/??/{ano_padrao}",
            "hora": hora,
            "nome": nome,
            "valor": valor,
            "banco": "BB",
        })

    unicos = []
    vistos = set()
    for d in dados:
        chave = (d["hora"], round(d["valor"], 2), d["nome"])
        if chave not in vistos:
            unicos.append(d)
            vistos.add(chave)
    dados = sorted(unicos, key=lambda d: d["hora"])

    print(f"\n========== [LOG - PIX RECEBIDOS BANCO DO BRASIL - FINAL] ==========")
    print(f"Total detectado: {len(dados)}\n")
    for i, d in enumerate(dados, start=1):
        print(f"[{i:03}] {d['data']} {d['hora']} | {d['nome']} | R${d['valor']:.2f}")
    print("=" * 100 + "\n")

    if not dados:
        return {"erro": "Nenhum lançamento PIX identificado no PDF do Banco do Brasil."}
    return {"banco": "bb", "dados": dados}


# ==========================================================
# 🟢 DETALHE BANCO C6 (SEM DESBLOQUEIO)
# ==========================================================
async def detalhe_c6(file_bytes: bytes, senha: str = None):
    """
    Extrai transações PIX de PDFs do C6 Bank.
    ❌ Sem desbloqueio: se o PDF estiver protegido, retorna erro pedindo PDF desbloqueado.
    """
    texto_total = ""
    try:
        with pdfplumber.open(io.BytesIO(file_bytes), password=senha or None) as pdf:
            for page in pdf.pages:
                texto_total += "\n" + (page.extract_text() or "")
    except Exception as e:
        erro_str = str(e).lower()
        if any(word in erro_str for word in ["password", "encrypt", "decrypt", "permiss"]):
            return {"erro": "O PDF C6 está protegido. Envie o PDF já desbloqueado."}
        return {"erro": f"Erro ao abrir PDF C6: {e}"}

    return {"dados": extrair_pix_c6(texto_total)}


# ==========================================================
# 🧩 PARSER DE TEXTO - C6
# ==========================================================
def extrair_pix_c6(texto_total: str):
    texto_limpo = re.sub(r"\s+", " ", texto_total)

    padrao = re.compile(
        r"(\d{2}/\d{2})(?:/\d{4})?.{0,30}?Pix\s+recebid[oa](?:\s+c6)?\s+(?:de\s+)?([A-Za-zÀ-ÿ0-9\.\-\,\s]+?)\s+R\$?\s*([\d\.,]+)(?:\s+às\s+(\d{2}:\d{2}))?",
        re.IGNORECASE
    )

    dados = []
    ano_atual = datetime.now().year

    for m in padrao.finditer(texto_limpo):
        data_curta, nome_raw, valor_txt, hora = m.groups()

        if data_curta is None:
            data = ""
        elif re.fullmatch(r"\d{2}/\d{2}/\d{4}", data_curta):
            data = data_curta
        elif re.fullmatch(r"\d{2}/\d{2}", data_curta):
            data = f"{data_curta}/{ano_atual}"
        else:
            data = ""

        nome = re.sub(r"\s{2,}", " ", (nome_raw or "").strip()).title()
        nome = re.sub(r"(?i)\b(Agência|Conta|Saldo|Pix)\b.*", "", nome).strip()
        if not nome:
            continue

        valor_txt_norm = (valor_txt or "").replace(".", "").replace(",", ".")
        try:
            valor = float(valor_txt_norm)
        except:
            continue
        if valor <= 0:
            continue

        hora = hora or ""
        dados.append({"data": data, "hora": hora, "nome": nome, "valor": valor})

    if not dados:
        linhas = re.split(r"\n+", texto_limpo)
        ano = ano_atual
        for ln in linhas:
            m2 = re.search(r"(\d{2}/\d{2})(?:/\d{4})?.{0,30}?(?:Pix\s+recebid[oa]).*?R\$?\s*([\d\.,]+)", ln, re.IGNORECASE)
            if m2:
                data_curta = m2.group(1)
                valor_txt = m2.group(2)
                nome_match = re.search(r"Pix\s+recebid[oa].*?de\s+(.*?)\s+R\$", ln, re.IGNORECASE)
                nome = (nome_match.group(1).strip().title() if nome_match else "(sem nome)")
                if re.fullmatch(r"\d{2}/\d{2}/\d{4}", data_curta):
                    data = data_curta
                else:
                    data = f"{data_curta}/{ano}"
                try:
                    valor = float(valor_txt.replace(".", "").replace(",", "."))
                except:
                    continue
                dados.append({"data": data, "nome": nome, "valor": valor, "hora": ""})

    unicos = []
    vistos = set()
    for d in dados:
        chave = (d.get("data", ""), d.get("hora", ""), round(d.get("valor", 0.0), 2), d.get("nome", ""))
        if chave not in vistos:
            unicos.append(d)
            vistos.add(chave)
    dados = unicos

    def sort_key(d):
        try:
            dt = try_parse_date(d.get("data", "")) or datetime.min.date()
        except:
            dt = datetime.min.date()
        hora = d.get("hora") or ""
        return (dt, hora)

    dados = sorted(dados, key=sort_key)

    print(f"\n========== [LOG - PIX RECEBIDOS C6 BANK - FINAL] ==========")
    print(f"Total detectado: {len(dados)}\n")
    for i, d in enumerate(dados, start=1):
        print(f"[{i:03}] {d.get('data','')} {d.get('hora','')} | {d.get('nome')} | R${d.get('valor'):.2f}")
    print("=" * 100 + "\n")

    return dados


# ==========================================================
# 🔍 PROCESSAR PDF → Detecta e chama o parser correto
# ==========================================================
async def processar_pdf(file_bytes: bytes, senha: str = None):
    texto_total = ""
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                texto_total += "\n" + (page.extract_text() or "")
    except Exception as e:
        return {"erro": f"Erro ao processar PDF: {e}"}

    texto_total = re.sub(r"\s+", " ", texto_total)
    upper = texto_total.upper()

    if "C6" in upper or "C6BANK" in upper:
        banco = "c6"
        resp = await detalhe_c6(file_bytes, senha)
        if "erro" in resp:
            return resp
        dados = resp.get("dados", [])
        for d in dados:
            d["banco"] = "C6"

    elif "BANCO DO BRASIL" in upper or "EXTRATO DE CONTA" in upper or "BB S.A" in upper:
        banco = "bb"
        resp = await detalhe_bb(file_bytes)
        if "erro" in resp:
            return resp
        dados = resp.get("dados", [])
        for d in dados:
            d["banco"] = "BB"

    else:
        banco = "desconhecido"
        return {"erro": "Banco não reconhecido no PDF."}

    if not dados:
        return {"erro": f"Nenhum lançamento PIX identificado no PDF do banco {banco.upper()}."}

    return {"banco": banco, "dados": dados}


async def processar_excel(file_bytes: bytes):
    def normalizar_hora_excel(h: str) -> str:
        """Aceita 7h58, 758, 07:58, 07.58, 7, 07:58:00 → retorna HH:MM"""
        if not h:
            return ""
        h = str(h).strip().lower().replace(" ", "")
        h = h.replace(".", ":").replace("h", ":")
        if re.fullmatch(r"^\d{1,2}$", h):
            return f"{int(h):02d}:00"
        if re.fullmatch(r"^\d{3,4}$", h):
            return f"{int(h[:-2]):02d}:{int(h[-2:]):02d}"
        if re.fullmatch(r"^\d{1,2}:\d{1,2}$", h):
            partes = h.split(":")
            return f"{int(partes[0]):02d}:{int(partes[1]):02d}"
        if re.fullmatch(r"^\d{2}:\d{2}:\d{2}$", h):
            return h[:5]
        return ""

    try:
        excel = pd.ExcelFile(io.BytesIO(file_bytes))
    except Exception as e:
        return {"erro": "Erro ao abrir Excel: " + str(e)}

    todas_linhas = []
    agente_atual = None

    for aba in excel.sheet_names:
        df = pd.read_excel(excel, aba, header=None, dtype=object)
        for _, linha in df.iterrows():
            texto = " ".join(str(x) for x in linha if pd.notna(x)).strip()
            if "AGENTE" in texto.upper():
                m_ag = re.search(r"AGENTE[:/]\s*([A-Za-zÀ-ÿ0-9\s]+)", texto, re.IGNORECASE)
                nome_agente = ""
                if m_ag:
                    nome_agente = re.sub(r"\d+", "", m_ag.group(1)).strip().upper()

                setor = ""
                try:
                    ultima_coluna = str(linha.iloc[-1]).strip()
                    if ultima_coluna and not re.search(r"\d{2}/\d{2}/\d{4}", ultima_coluna):
                        setor = ultima_coluna.upper()
                except:
                    pass

                if nome_agente and setor:
                    agente_atual = f"{nome_agente} - {setor}"
                else:
                    agente_atual = nome_agente

                continue
            if not agente_atual:
                continue

            nome, hora, raw_val = "", "", ""
            if 0 in linha.index and pd.notna(linha.iloc[0]):
                nome = str(linha.iloc[0]).strip()
            if 1 in linha.index and pd.notna(linha.iloc[1]):
                hora = normalizar_hora_excel(str(linha.iloc[1]))
            if 4 in linha.index and pd.notna(linha.iloc[4]):
                raw_val = linha.iloc[4]

            if not nome or re.search(r"TOTAL|NOME", nome, re.IGNORECASE):
                continue

            valor = parse_valor_robusto(raw_val)
            todas_linhas.append({
                "agente": agente_atual,
                "nome": nome.title(),
                "hora": hora,
                "valor": valor
            })

    if not todas_linhas:
        return {"erro": "Nenhum dado válido encontrado na planilha."}
    return {"tabela": todas_linhas}


@app.post("/conferir_caixa")
async def conferir_caixa(
    pdfs: List[UploadFile] = File(...),
    excels: List[UploadFile] = File(...),
    data: str = Form(None),   # recebido apenas por compatibilidade, sem filtrar
    senha: str = Form(None)
):
    todos_pdf = []
    bancos_detectados = set()

    # ============================
    # PROCESSAR PDFs
    # ============================
    for pdf in pdfs:
        try:
            pdf_bytes = await pdf.read()
            pdf_resp = await processar_pdf(pdf_bytes, senha)

            if "erro" in pdf_resp:
                continue

            bancos_detectados.add(pdf_resp.get("banco", "").upper())
            todos_pdf.extend(pdf_resp.get("dados", []))
        except:
            pass

    if not todos_pdf:
        return {"erro": "Nenhum PDF válido ou sem PIX encontrado."}

    dados_pdf = todos_pdf

    # ============================
    # PROCESSAR EXCELS
    # ============================
    dados_excel = []
    for excel in excels:
        excel_bytes = await excel.read()
        excel_resp = await processar_excel(excel_bytes)
        if "tabela" in excel_resp:
            dados_excel.extend(excel_resp["tabela"])

    if not dados_excel:
        return {"erro": "Nenhum dado válido encontrado nas planilhas enviadas."}

    # ============================
    # FUNÇÕES AUXILIARES
    # ============================
    def normalizar(s: str):
        s = unicodedata.normalize("NFKD", s or "")
        s = "".join(c for c in s if not unicodedata.combining(c))
        return s.lower().strip()

    def similaridade(a, b):
        return SequenceMatcher(None, normalizar(a), normalizar(b)).ratio()

    def normalizar_hora(h: str) -> str:
        if not h:
            return ""
        h = h.strip().lower().replace("h", ":").replace(".", ":")

        if re.fullmatch(r"\d{1,2}$", h):
            return f"{int(h):02d}:00"

        if re.fullmatch(r"\d{3,4}$", h):
            return f"{int(h[:-2]):02d}:{int(h[-2:]):02d}"

        if re.fullmatch(r"\d{1,2}:\d{1,2}$", h):
            p = h.split(":")
            return f"{int(p[0]):02d}:{int(p[1]):02d}"

        if re.fullmatch(r"\d{2}:\d{2}:\d{2}$", h):
            return h[:5]

        return ""

    usados_pdf = set()
    usado_por = {}
    conferidos = []
    faltando_no_pdf = []
    faltando_no_excel = []

    n_excel = len(dados_excel)
    n_pdf = len(dados_pdf)

    # ==========================================================
    # MATCH Excel → PDF (atribuição ótima — algoritmo húngaro)
    # ==========================================================
    # Antes: cada lançamento do Excel era casado na ordem em que aparecia na
    # planilha e "travava" o melhor PDF disponível na hora — um casamento
    # guloso. Isso deixa o resultado dependente da ordem das linhas: um
    # lançamento no topo da planilha podia roubar o PDF que seria o par
    # perfeito de outro lançamento mais abaixo, gerando divergências
    # evitáveis. Aqui montamos uma matriz de pontuação Excel × PDF e
    # resolvemos como problema de atribuição (scipy.optimize.linear_sum_assignment),
    # encontrando a combinação que maximiza a soma de todas as pontuações ao
    # mesmo tempo — não apenas a melhor escolha local de cada linha.
    if n_excel and n_pdf:
        excel_info = [
            {
                "nome": item["nome"],
                "valor": round(item.get("valor") or 0.0, 2),
                "hora": normalizar_hora(item.get("hora", "")),
            }
            for item in dados_excel
        ]
        pdf_info = [
            {
                "nome": p["nome"],
                "valor": round(p.get("valor") or 0.0, 2),
                "hora": normalizar_hora(p.get("hora", "")),
            }
            for p in dados_pdf
        ]

        CUSTO_INVIAVEL = 1e6  # par com valor incompatível: nunca deve ser escolhido

        def pontuar(e, p):
            dif_valor = abs(e["valor"] - p["valor"])
            if dif_valor > 0.50:
                return None

            sim = similaridade(e["nome"], p["nome"])
            ne, npn = normalizar(e["nome"]), normalizar(p["nome"])
            tok_e, tok_p = set(ne.split()), set(npn.split())

            if ne == npn:
                sim = max(sim, 0.95)
            elif (ne in npn or npn in ne) and len(tok_e) >= 2 and len(tok_p) >= 2:
                sim = max(sim, 0.90)
            else:
                comuns = tok_e & tok_p
                if len(comuns) >= 2:
                    # nome E sobrenome batendo: sinal forte de ser a mesma pessoa
                    sim = max(sim, min(0.85, sim + 0.20))
                elif len(tok_e) >= 2 and len(tok_p) >= 2:
                    # Nomes com várias palavras que só coincidem em NO MÁXIMO
                    # um token — tipicamente um primeiro nome ou sobrenome
                    # comum (Maria, José, Ana, Souza...). Isso não é sinal
                    # confiável de ser a mesma pessoa, mesmo quando o
                    # SequenceMatcher bruto dá um número enganosamente alto
                    # só por coincidência de tamanho/caracteres — foi
                    # exatamente esse caso que gerou falsos positivos reais
                    # (ex: "Ana Gabrielly Cunha Mesquita" casando com "Ana
                    # Rafaele", "Maria Maia" com "Maria Eduarda"). Suprime em
                    # vez de confiar.
                    sim = min(sim, 0.45)

            # O horário só ajuda ou atrapalha quando os DOIS lados o
            # informam — ausência de horário fica neutra (nem bônus, nem
            # penalidade). Quando os dois são conhecidos e divergem muito,
            # a penalidade cresce; um nome+valor batendo não deve mascarar
            # um horário completamente incompatível.
            hora_bonus = 0
            hora_delta = None
            if e["hora"] and p["hora"]:
                try:
                    t1 = datetime.strptime(e["hora"], "%H:%M")
                    t2 = datetime.strptime(p["hora"], "%H:%M")
                    hora_delta = abs((t1 - t2).total_seconds())
                    if hora_delta <= 10:
                        hora_bonus = 50
                    elif hora_delta <= 60:
                        hora_bonus = 35
                    elif hora_delta <= 300:
                        hora_bonus = 20
                    elif hora_delta <= 600:
                        hora_bonus = 10
                    elif hora_delta <= 3600:
                        hora_bonus = -20
                    else:
                        hora_bonus = -45
                except Exception:
                    hora_delta = None

            valor_score = 40 if dif_valor < 0.01 else max(0.0, 40 - dif_valor * 60)
            score = (sim * 100) + valor_score + hora_bonus
            return score, sim, dif_valor, hora_delta

        custo = np.full((n_excel, n_pdf), CUSTO_INVIAVEL)
        detalhes = {}
        for i, e in enumerate(excel_info):
            for j, p in enumerate(pdf_info):
                resultado = pontuar(e, p)
                if resultado is not None:
                    score, sim, dif_valor, hora_delta = resultado
                    custo[i, j] = -score
                    detalhes[(i, j)] = (score, sim, dif_valor, hora_delta)

        linhas, colunas = linear_sum_assignment(custo)

        GAP_HORA_MAX = 2 * 3600  # 2h: nome+valor batendo não vence um horário muito diferente

        aceito_excel = {}
        aceito_pdf = {}
        for i, j in zip(linhas, colunas):
            par = detalhes.get((i, j))
            if par is None:
                continue  # atribuição forçada pela matriz (nenhum par viável)
            score, sim, dif_valor, hora_delta = par
            aceito = sim >= 0.70 or (sim >= 0.55 and dif_valor < 0.01)
            if aceito and hora_delta is not None and hora_delta > GAP_HORA_MAX:
                aceito = False
            if aceito:
                aceito_excel[i] = j
                aceito_pdf[j] = i

        for i, item in enumerate(dados_excel):
            e = excel_info[i]

            if i in aceito_excel:
                j = aceito_excel[i]
                score, sim, dif_valor, hora_delta = detalhes[(i, j)]
                p_completo = dados_pdf[j]
                usados_pdf.add(j)
                usado_por[j] = {
                    "agente": item.get("agente", ""),
                    "nome_excel": e["nome"],
                    "valor": e["valor"],
                    "hora": e["hora"],
                }
                conferidos.append({
                    "agente": item.get("agente", ""),
                    "nome_excel": e["nome"],
                    "nome_pdf": pdf_info[j]["nome"],
                    "valor_excel": e["valor"],
                    "valor_pdf": pdf_info[j]["valor"],
                    "hora_excel": e["hora"],
                    "hora_pdf": pdf_info[j]["hora"],
                    "data_pdf": p_completo.get("data"),
                    "similaridade": round(sim, 2),
                    "analise": "ok",
                    "banco": p_completo.get("banco"),
                })
                continue

            # Não casou: acha o melhor candidato entre TODOS os PDFs (mesmo
            # os rejeitados ou já usados por outro agente) só para explicar
            # o motivo ao usuário.
            melhor_j, melhor_score, melhor_sim, melhor_dif = None, float("-inf"), 0.0, None
            for j in range(n_pdf):
                par = detalhes.get((i, j))
                if par is None:
                    continue
                score, sim, dif_valor, _hora_delta = par
                if score > melhor_score:
                    melhor_j, melhor_score, melhor_sim, melhor_dif = j, score, sim, dif_valor

            if melhor_j is not None and melhor_j in aceito_pdf:
                outro_i = aceito_pdf[melhor_j]
                outro_agente = dados_excel[outro_i].get("agente", "(desconhecido)")
                p = pdf_info[melhor_j]
                item["motivo"] = (
                    f"PIX já foi conferido por outro agente: {outro_agente} "
                    f"— {p['nome']} R${p['valor']:.2f} • {p['hora']}"
                )
                item["banco"] = dados_pdf[melhor_j].get("banco", "")
            elif melhor_j is not None:
                p = pdf_info[melhor_j]
                val_msg = (
                    "igual" if melhor_dif < 0.01 else
                    "próximo" if melhor_dif <= 0.50 else
                    "diferente"
                )
                hora_msg = ""
                if e["hora"] and p["hora"]:
                    if e["hora"] == p["hora"]:
                        hora_msg = f", horário igual ({p['hora']})"
                    else:
                        hora_msg = f", horários diferentes (Excel {e['hora']} ≠ PDF {p['hora']})"
                item["motivo"] = (
                    f"Nome semelhante encontrado: '{p['nome']}' "
                    f"(Sim={melhor_sim:.2f}), valor {val_msg} (R${p['valor']:.2f})"
                    f"{hora_msg}."
                )
                item["banco"] = dados_pdf[melhor_j].get("banco", "")
            else:
                item["motivo"] = "Nenhum parecido encontrado no PDF (ou já consumido por outro agente)."
                item["banco"] = ""

            faltando_no_pdf.append(item)
    else:
        for item in dados_excel:
            item["motivo"] = "Nenhum lançamento no PDF para comparar."
            item["banco"] = ""
            faltando_no_pdf.append(item)

    # ============================
    # PDF → Excel (não usados)
    # ============================
    for i, p in enumerate(dados_pdf):
        if i not in usados_pdf:
            faltando_no_excel.append({
                "nome": p["nome"],
                "hora": normalizar_hora(p.get("hora", "")),
                "valor": round(p.get("valor", 0), 2),
                "data": p.get("data"),
                "banco": p.get("banco", "")
            })

    return {
        "banco": ", ".join(bancos_detectados),
        "conferidos": conferidos,
        "faltando_no_pdf": faltando_no_pdf,
        "faltando_no_excel": faltando_no_excel
    }
