from fastapi import FastAPI, File, UploadFile, Form
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import io, re, pdfplumber
from datetime import datetime, date
from difflib import SequenceMatcher
import unicodedata
from typing import List, Dict, Any
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


def filter_items_by_date(items: List[Dict[str, Any]], selected: date) -> List[Dict[str, Any]]:
    """Filtra lista de dicionários onde exista algum campo de data que bata com selected."""
    if not selected:
        return items[:]
    keys_to_try = ["data", "date", "data_pdf", "data_excel", "dia"]
    filtered = []
    for it in items:
        matched = False
        for k in keys_to_try:
            if k in it and it[k]:
                dt = try_parse_date(str(it[k]))
                if dt and dt == selected:
                    matched = True
                    break
        if not matched:
            for v in it.values():
                if isinstance(v, str):
                    dt = try_parse_date(v)
                    if dt and dt == selected:
                        matched = True
                        break
        if matched:
            filtered.append(it)
    return filtered


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
    Parser robusto e filtrado para extratos do Banco do Brasil.
    ✅ Captura todos os PIX RECEBIDOS (com '(+)')
    ✅ Corrige PIX quebrados entre páginas
    ✅ Reconstrói PIX com CNPJ sem nome
    🚫 Ignora ruídos como '5 Pix - Recebido' ou cabeçalhos incompletos.
    """
    print("\n========== [DEBUG] INÍCIO DA LEITURA PDF BANCO DO BRASIL ==========\n")

    try:
        texto_total = ""
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for i, page in enumerate(pdf.pages):
                texto_pagina = page.extract_text() or ""
                texto_total += "\n" + texto_pagina

        with open("pdf_debug.txt", "w", encoding="utf-8") as f:
            f.write(texto_total)

    except Exception as e:
        print(f"\n⚠️ Erro ao ler PDF: {e}")
        return {"erro": f"Falha ao processar PDF ({e})"}

    print("\n========== [DEBUG] LIMPEZA E NORMALIZAÇÃO ==========\n")

    texto_total = re.sub(r"\s+", " ", texto_total)
    texto_limpo = texto_total

    texto_limpo = re.sub(r"Extrato de Conta Corrente.*?Valor", " ", texto_limpo, flags=re.IGNORECASE)
    texto_limpo = re.sub(r"----- Página \d+ -----", " ", texto_limpo)
    texto_limpo = re.sub(r"\s+", " ", texto_limpo)

    texto_limpo = re.sub(r"\(\-\)", "", texto_limpo)
    texto_limpo = re.sub(r"(?i)Pix\s*-\s*Enviado", "", texto_limpo)

    # juntar pix cortados
    corrigidos = []
    texto_expandido = texto_limpo

    pix_soltos = list(re.finditer(
        r"Pix\s*-\s*Recebido\s+([\d\.,]+)\s*\(\+\)(?!\s*[A-ZÀ-ÿ])",
        texto_expandido,
        flags=re.IGNORECASE
    ))
    for m in pix_soltos:
        valor_txt = m.group(1)
        pos_fim = m.end()
        trecho_proximo = texto_expandido[pos_fim:pos_fim + 800]
        trecho_proximo = re.sub(
            r"Extrato de Conta Corrente|Cliente\s+[A-ZÀ-ÿ\s]+|Ag[êe]ncia:\s*\d+-\d+\s*Conta:\s*\d+-\d+|Lançamentos|Dia\s+Lote\s+Documento\s+Histórico\s+Valor",
            " ",
            trecho_proximo,
            flags=re.IGNORECASE
        )

        prox = re.search(
            r"(\d{2}/\d{2})\s+(\d{2}:\d{2})\s+[0-9]{6,}\s+([A-Za-zÀ-ÿ\s]{3,}?)\s+\d{1,3}",
            trecho_proximo
        )

        if prox:
            data_curta, hora, nome_raw = prox.groups()
            nome = re.sub(r"\s{2,}", " ", nome_raw.strip()).title()
            bloco_completo = f"Pix - Recebido {data_curta}/2025 {hora} {nome} {valor_txt} (+)"
            corrigidos.append((m.start(), bloco_completo))
            print(f"🧩 [RECONSTRUÍDO] {data_curta}/2025 {hora} | {nome} | R${valor_txt}")

    if corrigidos:
        partes = []
        ultimo_fim = 0
        for inicio, bloco_corrigido in corrigidos:
            partes.append(texto_expandido[ultimo_fim:inicio])
            partes.append(bloco_corrigido)
            ultimo_fim = inicio
        texto_limpo = "".join(partes) + texto_expandido[ultimo_fim:]

    blocos = re.split(r"(?=Pix\s*-\s*Recebido)", texto_limpo, flags=re.IGNORECASE)

    # conserto local
    corrigido = []
    padrao_pix_valor = re.compile(r"Pix\s*-\s*Recebido\s+([\d\.,]+)\s*\(\+\)", re.IGNORECASE)
    padrao_nome_hora = re.compile(r"(\d{2}/\d{2})\s+(\d{2}:\d{2})\s+[0-9]{6,}\s+([A-ZÀ-ÿ\s]{3,})\s+\d+", re.IGNORECASE)
    for i, bloco in enumerate(blocos[:-1]):
        m_val = padrao_pix_valor.search(bloco)
        if not m_val:
            corrigido.append(bloco)
            continue
        valor_txt = m_val.group(1)
        trecho_entre = blocos[i + 1][:150]
        m_nome = padrao_nome_hora.search(trecho_entre)
        if m_nome:
            data_curta, hora, nome_raw = m_nome.groups()
            nome = re.sub(r"\s{2,}", " ", nome_raw.strip()).title()
            bloco_corrigido = f"Pix - Recebido {data_curta}/2025 {hora} {nome} {valor_txt} (+)"
            corrigido.append(bloco_corrigido)
        else:
            corrigido.append(bloco)
    blocos = corrigido

    dados = []
    print(f"🔍 Total de blocos detectados: {len(blocos)}\n")

    padrao_pix = re.compile(
        r"Pix\s*-\s*Recebido.*?"
        r"(?:(\d{2}/\d{2}/\d{4})|(\d{2}/\d{2}))?\s*"
        r"(\d{2}:\d{2})\s+"
        r"(?:[0-9]{5,}\s+)?"
        r"([A-ZÀ-ÿ0-9\s\.]{3,}?)\s+"
        r"([\d\.]+,\d{2})\s*\(\+\)",
        re.IGNORECASE,
    )

    data_atual = None

    for bloco in blocos:
        m = padrao_pix.search(bloco)
        if not m:
            continue

        data_full, data_partial, hora, nome_raw, valor_txt = m.groups()

        if data_full:
            data_atual = data_full
        elif data_partial:
            data_atual = f"{data_partial}/2025"
        else:
            m2 = re.search(r"(\d{2}/\d{2})(?:/\d{4})?", bloco)
            if m2:
                data_atual = f"{m2.group(1)}/2025"

        if not data_atual:
            idx = texto_limpo.find(bloco)
            if idx != -1:
                antes = texto_limpo[max(0, idx - 100):idx]
                m3 = re.search(r"(\d{2}/\d{2})(?:/\d{4})?", antes)
                if m3:
                    data_atual = f"{m3.group(1)}/2025"

        if not data_atual:
            continue

        data = data_atual

        nome = re.sub(r"\s{2,}", " ", nome_raw.strip()).title()
        nome = re.sub(r"(?i)\b(Agência|Conta|Saldo|Pix)\b.*", "", nome).strip()
        if not nome:
            continue
        try:
            valor = float(valor_txt.replace(".", "").replace(",", "."))
        except:
            continue
        if valor <= 0:
            continue

        nome_limpo = nome.strip()
        if re.fullmatch(r"[0-9.\s]{7,}", nome_limpo):
            cnpj_num = re.sub(r"\D", "", nome_limpo)
            nome_limpo = f"Cliente CNPJ {cnpj_num}"
        nome = nome_limpo

        dados.append({
            "data": data,
            "hora": hora,
            "nome": nome,
            "valor": valor,
            "banco": "BB"
        })

    padrao_cnpj = re.compile(
        r"(\d{2}/\d{2})\s+(\d{2}:\d{2})\s+[0-9]{6,}\s+(\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}|\d{11,14})\s+([\d\.,]+)\s*\(\+\)",
        re.IGNORECASE,
    )
    for m in padrao_cnpj.finditer(texto_limpo):
        data_curta, hora, cnpj_raw, valor_txt = m.groups()
        nome = f"Cliente CNPJ {re.sub(r'[^0-9]', '', cnpj_raw)}"
        try:
            valor = float(valor_txt.replace(".", "").replace(",", "."))
        except:
            continue
        dados.append({"data": f"{data_curta}/2025", "hora": hora, "nome": nome, "valor": valor, "banco": "BB"})

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
                # Extrai nome do agente
                m_ag = re.search(r"AGENTE[:/]\s*([A-Za-zÀ-ÿ0-9\s]+)", texto, re.IGNORECASE)
                nome_agente = ""
                if m_ag:
                    nome_agente = re.sub(r"\d+", "", m_ag.group(1)).strip().upper()

                # Tenta pegar setor na última coluna da linha
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
    data: str = Form(None),
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
    # FILTRAR POR DATA
    # ============================
    selected_date = None
    if data:
        try:
            selected_date = datetime.strptime(data.strip(), "%Y-%m-%d").date()
        except:
            selected_date = try_parse_date(data.strip())

    if selected_date:
        dados_pdf = filter_items_by_date(dados_pdf, selected_date)

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

    # ==========================================================
    # ✅ REGRA: 1 PIX DO PDF SÓ PODE SER USADO 1 VEZ
    # ==========================================================
    usados_pdf = set()     # índices do pdf já consumidos
    usado_por = {}         # idx -> info de quem consumiu
    conferidos = []
    faltando_no_pdf = []
    faltando_no_excel = []

    # ============================
    # MATCH Excel → PDF
    # ============================
    for item in dados_excel:
        nome_excel = item["nome"]
        valor_excel = round(item.get("valor") or 0.0, 2)
        hora_excel = normalizar_hora(item.get("hora", ""))
        agente_excel = item.get("agente", "")

        escolhido = None
        candidatos = []
        melhor_ja_usado = None

        # ----------------------------
        # candidatos com valor igual
        # ----------------------------
        for idx, p in enumerate(dados_pdf):
            nome_pdf = p["nome"]
            valor_pdf = round(p.get("valor") or 0.0, 2)
            hora_pdf = normalizar_hora(p.get("hora", ""))

            if abs(valor_excel - valor_pdf) < 0.01:
                sim = similaridade(nome_excel, nome_pdf)

                ne = normalizar(nome_excel)
                np = normalizar(nome_pdf)

                if ne in np:
                    sim = max(sim, 0.90)
                elif np in ne:
                    sim = max(sim, 0.90)
                else:
                    if set(ne.split()).intersection(set(np.split())):
                        sim = max(sim, min(0.75, sim + 0.20))

                hora_ok = True
                hora_delta = 999999

                if hora_excel and hora_pdf:
                    try:
                        t1 = datetime.strptime(hora_excel, "%H:%M")
                        t2 = datetime.strptime(hora_pdf, "%H:%M")
                        hora_delta = abs((t1 - t2).total_seconds())
                        hora_ok = hora_delta <= 600
                    except:
                        pass

                # ✅ se já foi usado, NÃO deixa virar candidato
                if idx in usados_pdf:
                    score_dup = (sim * 100) + (20 if hora_ok else 0) - (hora_delta / 1000)
                    if (melhor_ja_usado is None) or (score_dup > melhor_ja_usado["score"]):
                        melhor_ja_usado = {
                            "idx": idx,
                            "score": score_dup,
                            "sim": sim,
                            "hora_ok": hora_ok,
                            "hora_delta": hora_delta,
                            "valor_pdf": valor_pdf,
                            "nome_pdf": nome_pdf,
                            "hora_pdf": hora_pdf,
                            "data_pdf": p.get("data"),
                            "usado_por": usado_por.get(idx),
                        }
                    continue

                candidatos.append({
                    "idx": idx,
                    "sim": sim,
                    "hora_ok": hora_ok,
                    "hora_delta": hora_delta,
                    "valor_pdf": valor_pdf,
                    "nome_pdf": nome_pdf,
                    "hora_pdf": hora_pdf,
                    "data_pdf": p.get("data"),
                })

        if candidatos:
            candidatos.sort(key=lambda x: (
                not x["hora_ok"],
                -x["sim"],
                x["hora_delta"]
            ))
            escolhido = candidatos[0]

        # ----------------------------
        # confirma
        # ----------------------------
        if escolhido and (
            escolhido["sim"] >= 0.70 or
            (escolhido["sim"] >= 0.55 and abs(valor_excel - escolhido["valor_pdf"]) < 0.01)
        ):
            idx_escolhido = escolhido["idx"]
            usados_pdf.add(idx_escolhido)
            usado_por[idx_escolhido] = {
                "agente": agente_excel,
                "nome_excel": nome_excel,
                "valor": valor_excel,
                "hora": hora_excel,
            }

            conferidos.append({
                "agente": agente_excel,
                "nome_excel": nome_excel,
                "nome_pdf": escolhido["nome_pdf"],
                "valor_excel": valor_excel,
                "valor_pdf": escolhido["valor_pdf"],
                "hora_excel": hora_excel,
                "hora_pdf": escolhido["hora_pdf"],
                "data_pdf": escolhido["data_pdf"],
                "similaridade": round(escolhido["sim"], 2),
                "analise": "ok",
                "banco": dados_pdf[idx_escolhido].get("banco"),
            })
            continue

        # ----------------------------
        # sugestão valor próximo (✅ não reutiliza usados)
        # ----------------------------
        melhor_pontuacao = -999999
        possivel = None

        candidatos_valor = []
        for idx, p in enumerate(dados_pdf):
            if idx in usados_pdf:
                continue
            if abs(valor_excel - round(p.get("valor", 0), 2)) <= 0.50:
                candidatos_valor.append((idx, p))

        lista_busca = candidatos_valor if candidatos_valor else [(idx, p) for idx, p in enumerate(dados_pdf) if idx not in usados_pdf]

        for idx, p in lista_busca:
            nome_pdf = p["nome"]
            valor_pdf = round(p.get("valor") or 0.0, 2)
            hora_pdf = normalizar_hora(p.get("hora", ""))

            sim = similaridade(nome_excel, nome_pdf)

            ne = normalizar(nome_excel)
            np = normalizar(nome_pdf)

            if ne in np:
                sim = max(sim, 0.90)
            elif np in ne:
                sim = max(sim, 0.90)
            else:
                if set(ne.split()).intersection(set(np.split())):
                    sim = max(sim, min(0.75, sim + 0.20))

            dif_valor = abs(valor_excel - valor_pdf)

            hora_bonus = 0
            if hora_excel and hora_pdf:
                try:
                    t1 = datetime.strptime(hora_excel, "%H:%M")
                    t2 = datetime.strptime(hora_pdf, "%H:%M")
                    delta = abs((t1 - t2).total_seconds())

                    if delta <= 10:
                        hora_bonus = 50
                    elif delta <= 60:
                        hora_bonus = 35
                    elif delta <= 300:
                        hora_bonus = 20
                    elif delta <= 600:
                        hora_bonus = 10
                except:
                    pass

            pontuacao = (sim * 100) - dif_valor + hora_bonus

            if pontuacao > melhor_pontuacao:
                melhor_pontuacao = pontuacao
                possivel = (idx, p)

        if possivel:
            idx_p, p = possivel
            valor_pdf = round(p.get("valor") or 0.0, 2)
            val_dif = abs(valor_excel - valor_pdf)
            val_msg = (
                "igual" if val_dif < 0.01 else
                "próximo" if val_dif <= 0.50 else
                "diferente"
            )
            hora_pdf = normalizar_hora(p.get("hora", ""))

            hora_msg = ""
            if hora_excel and hora_pdf:
                if hora_excel == hora_pdf:
                    hora_msg = f", horário igual ({hora_pdf})"
                else:
                    hora_msg = f", horários diferentes (Excel {hora_excel} ≠ PDF {hora_pdf})"

            motivo = (
                f"Nome semelhante encontrado: '{p.get('nome','')}' "
                f"(Sim={similaridade(nome_excel, p.get('nome','')):.2f}), "
                f"valor {val_msg} (R${valor_pdf:.2f})"
                f"{hora_msg}."
            )

            item["banco"] = p.get("banco", "")
            item["motivo"] = motivo
            faltando_no_pdf.append(item)
            continue

        # se não achou disponível, mas tinha um match que já estava usado
        if melhor_ja_usado and melhor_ja_usado.get("usado_por"):
            up = melhor_ja_usado["usado_por"]
            motivo = (
                f"PIX já foi conferido por outro agente: {up.get('agente','(desconhecido)')} "
                f"— {melhor_ja_usado.get('nome_pdf','')} "
                f"R${melhor_ja_usado.get('valor_pdf',0):.2f} • {melhor_ja_usado.get('hora_pdf','')}"
            )
            item["motivo"] = motivo
            item["banco"] = dados_pdf[melhor_ja_usado["idx"]].get("banco", "")
        else:
            item["motivo"] = "Nenhum parecido encontrado no PDF (ou já consumido por outro agente)."
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
