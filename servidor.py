from fastapi import FastAPI, File, UploadFile, Form
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
import io, re, pdfplumber
from datetime import datetime, date   # ✅ <-- ADICIONADO
from difflib import SequenceMatcher
import unicodedata
from typing import List, Dict, Any    # ✅ tipos usados nas funções
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
    # tentar extrair trecho dd/mm/yyyy dentro da string (ex: '03/11/2025 07:11')
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
        return items[:]  # sem filtro -> retorna tudo
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
        # se nenhuma das chaves padrão existir, tentar achar uma data dentro dos valores de texto
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
                #print(f"\n----- Página {i+1} -----\n")
                #print(texto_pagina[:1500])  # preview
                texto_total += "\n" + texto_pagina

        with open("pdf_debug.txt", "w", encoding="utf-8") as f:
            f.write(texto_total)

    except Exception as e:
        print(f"\n⚠️ Erro ao ler PDF: {e}")
        return {"erro": f"Falha ao processar PDF ({e})"}

    print("\n========== [DEBUG] LIMPEZA E NORMALIZAÇÃO ==========\n")

    texto_total = re.sub(r"\s+", " ", texto_total)
    texto_limpo = texto_total

    # Remove cabeçalhos e rodapés
    texto_limpo = re.sub(r"Extrato de Conta Corrente.*?Valor", " ", texto_limpo, flags=re.IGNORECASE)
    texto_limpo = re.sub(r"----- Página \d+ -----", " ", texto_limpo)
    texto_limpo = re.sub(r"\s+", " ", texto_limpo)

    # 🔹 Mantém apenas lançamentos com "(+)" = Pix recebido
    texto_limpo = re.sub(r"\(\-\)", "", texto_limpo)  # remove saídas
    texto_limpo = re.sub(r"(?i)Pix\s*-\s*Enviado", "", texto_limpo)  # remove envios

    # ==========================================================
        # ==========================================================
    # 🧠 Correção global: juntar PIX que ficaram cortados entre páginas
    # ==========================================================
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

            # 🟡 DEBUG EXTRA: loga no terminal quando reconstruir
            print(f"🧩 [RECONSTRUÍDO] {data_curta}/2025 {hora} | {nome} | R${valor_txt}")



    if corrigidos:
        partes = []
        ultimo_fim = 0
        for inicio, bloco_corrigido in corrigidos:
            partes.append(texto_expandido[ultimo_fim:inicio])
            partes.append(bloco_corrigido)
            ultimo_fim = inicio
        texto_limpo = "".join(partes) + texto_expandido[ultimo_fim:]

    # 🔹 Divide o texto em blocos que comecem em "Pix - Recebido"
    blocos = re.split(r"(?=Pix\s*-\s*Recebido)", texto_limpo, flags=re.IGNORECASE)

    # ==========================================================
    # 🔧 REGRA ESPECIAL: consertar PIX quebrados entre páginas
    # ==========================================================
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

    # ==========================================================
    # 🎯 REGEX PRINCIPAL: captura Pix válidos
    # ==========================================================
    padrao_pix = re.compile(
        r"Pix\s*-\s*Recebido.*?"
        r"(?:(\d{2}/\d{2}/\d{4})|(\d{2}/\d{2}))?\s*"  # data completa ou parcial
        r"(\d{2}:\d{2})\s+"                            # hora
        r"(?:[0-9]{5,}\s+)?"                           # ID transação (opcional)
        r"([A-ZÀ-ÿ0-9\s\.]{3,}?)\s+"                   # nome (pode conter números)
        r"([\d\.]+,\d{2})\s*\(\+\)",                   # valor (+)
        re.IGNORECASE,
    )

    data_atual = None  # mantém última data válida, se realmente faltar

    for bloco in blocos:
        m = padrao_pix.search(bloco)
        if not m:
            continue

        data_full, data_partial, hora, nome_raw, valor_txt = m.groups()

        # 1️⃣ tenta pegar a data do próprio bloco
        if data_full:
            data_atual = data_full
        elif data_partial:
            data_atual = f"{data_partial}/2025"
        else:
            # 2️⃣ tenta achar uma data próxima (dentro do próprio bloco)
            m2 = re.search(r"(\d{2}/\d{2})(?:/\d{4})?", bloco)
            if m2:
                data_atual = f"{m2.group(1)}/2025"

        # 3️⃣ se ainda não achou, busca até 100 caracteres antes do bloco
        if not data_atual:
            idx = texto_limpo.find(bloco)
            if idx != -1:
                antes = texto_limpo[max(0, idx - 100):idx]
                m3 = re.search(r"(\d{2}/\d{2})(?:/\d{4})?", antes)
                if m3:
                    data_atual = f"{m3.group(1)}/2025"

        # 4️⃣ ainda não tem data? ignora o bloco
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

        # 🔧 Correção automática de nomes "fora do padrão"
        nome_limpo = nome.strip()

        # Detecta CNPJs isolados (como 44.735.327 ou 00061162798300)
        if re.fullmatch(r"[0-9.\s]{7,}", nome_limpo):
            cnpj_num = re.sub(r"\D", "", nome_limpo)
            nome_limpo = f"Cliente CNPJ {cnpj_num}"

        # Corrige nomes truncados comuns (terminam abruptamente com 2 letras)
        if re.match(r".{3,}\s+[A-ZÀ-Úa-zà-ú]{1,2}$", nome_limpo):
            nome_limpo = nome_limpo + " (incompleto)"

        nome = nome_limpo

        dados.append({
            "data": data,
            "hora": hora,
            "nome": nome,
            "valor": valor,
            "banco": "BB"
        })


    # ==========================================================
    # 🧾 Captura complementar: PIX com CNPJ (sem nome)
    # ==========================================================
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

    # ==========================================================
    # 🧹 Limpeza final: remover duplicatas e ordenar
    # ==========================================================
    unicos = []
    vistos = set()
    for d in dados:
        chave = (d["hora"], round(d["valor"], 2), d["nome"])
        if chave not in vistos:
            unicos.append(d)
            vistos.add(chave)
    dados = sorted(unicos, key=lambda d: d["hora"])

    # ==========================================================
    # 🪵 LOG FINAL
    # ==========================================================
    print(f"\n========== [LOG - PIX RECEBIDOS BANCO DO BRASIL - FINAL] ==========")
    print(f"Total detectado: {len(dados)}\n")
    for i, d in enumerate(dados, start=1):
        print(f"[{i:03}] {d['data']} {d['hora']} | {d['nome']} | R${d['valor']:.2f}")
    print("=" * 100 + "\n")

    if not dados:
        return {"erro": "Nenhum lançamento PIX identificado no PDF do Banco do Brasil."}
    return {"banco": "bb", "dados": dados}

# ==========================================================
# 🟢 DETALHE BANCO C6 (versão robusta com desbloqueio automático)
# ==========================================================
async def detalhe_c6(file_bytes: bytes, senha: str = None):
    """
    Extrai transações PIX de PDFs do C6 Bank.
    Se o PDF estiver bloqueado, tenta desbloquear automaticamente com pikepdf.
    """
    texto_total = ""

    # 1️⃣ Tenta abrir diretamente
    try:
        with pdfplumber.open(io.BytesIO(file_bytes), password=senha or None) as pdf:
            for page in pdf.pages:
                texto_total += "\n" + (page.extract_text() or "")

    except Exception as e:
        erro_str = str(e).lower()
        print("⚠️ Falha ao abrir PDF C6:", erro_str)

        # 2️⃣ Se bloqueado, tenta desbloquear com pikepdf
        if any(word in erro_str for word in ["password", "encrypt", "decrypt", "permiss"]):
            print("🔓 Tentando desbloquear PDF C6 com pikepdf...")

            try:
                pdf_desbloqueado = pikepdf.open(io.BytesIO(file_bytes))
                buffer = io.BytesIO()
                pdf_desbloqueado.save(buffer)
                buffer.seek(0)

                with pdfplumber.open(buffer) as pdf:
                    for page in pdf.pages:
                        texto_total += "\n" + (page.extract_text() or "")
                print("✅ PDF C6 desbloqueado e processado com sucesso.")

            except pikepdf._qpdf.PasswordError:
                return {"erro": "O PDF C6 está protegido por senha. Informe a senha para continuar."}
            except Exception as e2:
                return {"erro": f"Falha ao desbloquear PDF C6: {e2}"}
        else:
            return {"erro": f"Erro ao abrir PDF C6: {e}"}

    # 3️⃣ Continua com o parser normal
    return {"dados": extrair_pix_c6(texto_total)}


# ==========================================================
# 🧩 PARSER DE TEXTO - C6 (idêntico ao seu, com refinamento)
# ==========================================================
def extrair_pix_c6(texto_total: str):
    texto_limpo = re.sub(r"\s+", " ", texto_total)

    padrao = re.compile(
        r"(\d{2}/\d{2})(?:/\d{4})?.{0,30}?Pix\s+recebid[oa](?:\s+c6)?\s+(?:de\s+)?([A-Za-zÀ-ÿ0-9\.\-\,\s]+?)\s+R\$?\s*([\d\.,]+)(?:\s+às\s+(\d{2}:\d{2}))?",
        re.IGNORECASE
    )

    dados = []
    for m in padrao.finditer(texto_limpo):
        data_curta, nome_raw, valor_txt, hora = m.groups()
        ano = datetime.now().year
        data = f"{data_curta}/" + (data_curta if "/" in data_curta else str(ano))
        nome = re.sub(r"\s{2,}", " ", (nome_raw or "").strip()).title()
        valor_txt = (valor_txt or "").replace(".", "").replace(",", ".")
        try:
            valor = float(valor_txt)
        except:
            continue
        hora = hora or ""

        if not any(d["nome"] == nome and abs(d["valor"] - valor) < 0.01 and d.get("hora", "") == hora for d in dados):
            dados.append({"data": data, "nome": nome, "valor": valor, "hora": hora, "banco": "C6"})

    if not dados:
        linhas = re.split(r"\n+", texto_limpo)
        for ln in linhas:
            m2 = re.search(r"(\d{2}/\d{2})(?:/\d{4})?.{0,30}?(?:Pix\s+recebid[oa]).*?R\$?\s*([\d\.,]+)", ln, re.IGNORECASE)
            if m2:
                data_curta = m2.group(1)
                valor_txt = m2.group(2)
                nome_match = re.search(r"Pix\s+recebid[oa].*?de\s+(.*?)\s+R\$", ln, re.IGNORECASE)
                nome = (nome_match.group(1).strip().title() if nome_match else "(sem nome)")
                data = f"{data_curta}/{ano}"
                try:
                    valor = float(valor_txt.replace(".", "").replace(",", "."))
                except:
                    continue
                dados.append({"data": data, "nome": nome, "valor": valor, "hora": ""})

    # LOG FINAL
    print(f"\n========== [LOG - PIX RECEBIDOS C6 BANK - FINAL] ==========")
    print(f"Total detectado: {len(dados)}\n")
    for i, d in enumerate(dados, start=1):
        print(f"[{i:03}] {d['data']} {d.get('hora','')} | {d['nome']} | R${d['valor']:.2f}")
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
        # ✅ Chamar com await e passando os bytes corretos
        resp = await detalhe_c6(file_bytes, senha)
        dados = resp.get("dados", [])
        for d in dados:
            d["banco"] = "C6"

    elif "BANCO DO BRASIL" in upper or "EXTRATO DE CONTA" in upper or "BB S.A" in upper:
        banco = "bb"
        # 🟢 Chama corretamente a função assíncrona detalhe_bb (que lê os bytes)
        resp = await detalhe_bb(file_bytes)
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
        # 7 -> 07:00
        if re.fullmatch(r"^\d{1,2}$", h):
            return f"{int(h):02d}:00"
        # 758 -> 07:58
        if re.fullmatch(r"^\d{3,4}$", h):
            return f"{int(h[:-2]):02d}:{int(h[-2:]):02d}"
        # 7:5 -> 07:05
        if re.fullmatch(r"^\d{1,2}:\d{1,2}$", h):
            partes = h.split(":")
            return f"{int(partes[0]):02d}:{int(partes[1]):02d}"
        # 07:58:00 -> 07:58
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
                if m_ag:
                    agente_atual = re.sub(r"\d+", "", m_ag.group(1)).strip().upper()
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



from typing import List
from datetime import datetime, timedelta

# ==========================================================
# 🧾 ROTA PRINCIPAL /conferir_caixa — MULTIPLOS PDF
# ==========================================================
@app.post("/conferir_caixa")
async def conferir_caixa(
    pdfs: List[UploadFile] = File(...),   # AGORA ACEITA MÚLTIPLOS PDFs
    excels: List[UploadFile] = File(...),
    data: str = Form(None),
    senha: str = Form(None)
):
    todos_pdf = []       # ← PIX DO C6 + BB
    bancos_detectados = set()

    # ==========================================================
    # 1️⃣ PROCESSAR CADA PDF ENVIADO
    # ==========================================================
    for pdf in pdfs:
        try:
            pdf_bytes = await pdf.read()
            pdf_resp = await processar_pdf(pdf_bytes, senha)

            if "erro" in pdf_resp:
                print(f"⚠️ PDF ignorado ({pdf.filename}): {pdf_resp['erro']}")
                continue

            bancos_detectados.add(pdf_resp.get("banco", "").upper())
            dados_pdf_local = pdf_resp.get("dados", [])
            todos_pdf.extend(dados_pdf_local)

            print(f"\n🏦 PDF: {pdf.filename}")
            print(f"Banco detectado: {pdf_resp.get('banco')}")
            print(f"PIX encontrados: {len(dados_pdf_local)}")

        except Exception as e:
            print(f"❌ Erro lendo PDF {pdf.filename}: {e}")

    if not todos_pdf:
        return {"erro": "Nenhum PDF válido ou sem PIX encontrado."}

    print(f"\n📌 TOTAL GERAL DE PIX (C6 + BB): {len(todos_pdf)}\n")

    dados_pdf = todos_pdf

    # ==========================================================
    # 2️⃣ PROCESSAR EXCELS
    # ==========================================================
    dados_excel = []
    for excel in excels:
        try:
            excel_bytes = await excel.read()
            excel_resp = await processar_excel(excel_bytes)
            if isinstance(excel_resp, dict) and "tabela" in excel_resp:
                dados_excel.extend(excel_resp["tabela"])
            print(f"✅ Excel {excel.filename}: {len(excel_resp.get('tabela', []))} linhas")
        except Exception as e:
            print(f"❌ Erro ao ler Excel {excel.filename}: {e}")

    if not dados_excel:
        return {"erro": "Nenhum dado válido encontrado nas planilhas enviadas."}

    # ==========================================================
    # 3️⃣ FILTRO POR DATA
    # ==========================================================
    selected_date = None
    if data:
        try:
            selected_date = datetime.strptime(data.strip(), "%Y-%m-%d").date()
        except:
            selected_date = try_parse_date(data.strip())

    if selected_date:
        total_pdf_antes = len(dados_pdf)
        dados_pdf = filter_items_by_date(dados_pdf, selected_date)
        print(f"\n📅 Filtro aplicado no PDF ({selected_date.strftime('%d/%m/%Y')})")
        print(f"   PDF: {total_pdf_antes} → {len(dados_pdf)} após filtro")
        print("=" * 60)

    # ==========================================================
    # 4️⃣ CONFERÊNCIA PRINCIPAL
    # ==========================================================
    def normalizar(s: str):
        s = unicodedata.normalize("NFKD", s or "")
        s = "".join(c for c in s if not unicodedata.combining(c))
        return s.lower().strip()

    def similaridade(n1: str, n2: str):
        return SequenceMatcher(None, normalizar(n1), normalizar(n2)).ratio()

    def normalizar_hora(h: str) -> str:
        if not h:
            return ""
        h = str(h).strip().lower().replace(" ", "")
        h = h.replace("h", ":")
        if re.fullmatch(r"^\d{1,2}[:\.]\d{1,2}$", h):
            p = re.split(r"[:\.]", h)
            return f"{int(p[0]):02d}:{int(p[1]):02d}"
        if re.fullmatch(r"^\d{3,4}$", h):
            return f"{int(h[:-2]):02d}:{int(h[-2:]):02d}"
        if re.fullmatch(r"^\d{1,2}$", h):
            return f"{int(h):02d}:00"
        if re.fullmatch(r"^\d{2}:\d{2}$", h):
            return h
        return ""

    usados_pdf = set()
    conferidos = []
    faltando_no_pdf = []
    faltando_no_excel = []

    # ==========================================================
    # 4.1️⃣ Excel → PDF
    # ==========================================================
    for item in dados_excel:
        nome_excel = item.get("nome", "").strip()
        valor_excel = round(item.get("valor") or 0.0, 2)
        hora_excel = normalizar_hora(item.get("hora", ""))
        agente_excel = (item.get("agente") or "").strip()

        escolhido = None
        candidatos = []

        for idx, p in enumerate(dados_pdf):
            nome_pdf = (p.get("nome") or "").strip()
            valor_pdf = round(p.get("valor") or 0.0, 2)
            hora_pdf = normalizar_hora(p.get("hora", ""))
            data_pdf = try_parse_date(str(p.get("data", "")))
            usado = idx in usados_pdf

            if abs(valor_excel - valor_pdf) < 0.01:
                sim = similaridade(nome_excel, nome_pdf)

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

                candidatos.append({
                    "idx": idx,
                    "nome_pdf": nome_pdf,
                    "valor_pdf": valor_pdf,
                    "hora_pdf": hora_pdf,
                    "data_pdf": data_pdf,
                    "sim": sim,
                    "hora_ok": hora_ok,
                    "hora_delta": hora_delta,
                    "usado": usado
                })

        if candidatos:
            candidatos = sorted(
                candidatos,
                key=lambda x: (x["usado"], not x["hora_ok"], -x["sim"], x["hora_delta"])
            )
            escolhido = candidatos[0]

        # CASO CONFIRA
        if escolhido and abs(valor_excel - escolhido["valor_pdf"]) < 0.01 and (
            escolhido["sim"] >= 0.55 or
            normalizar(escolhido["nome_pdf"]).startswith(normalizar(nome_excel)) or
            normalizar(nome_excel).startswith(normalizar(escolhido["nome_pdf"]))
        ):
            usados_pdf.add(escolhido["idx"])
            conferidos.append({
                "agente": agente_excel,
                "nome_excel": nome_excel,
                "nome_pdf": escolhido["nome_pdf"],
                "valor_excel": valor_excel,
                "valor_pdf": escolhido["valor_pdf"],
                "hora_excel": hora_excel,
                "hora_pdf": escolhido["hora_pdf"],
                "data_pdf": escolhido["data_pdf"].strftime("%d/%m/%Y") if escolhido["data_pdf"] else "",
                "similaridade": round(escolhido["sim"], 2),
                "analise": "ok",
                "banco": dados_pdf[escolhido["idx"]].get("banco")  # ok
            })
        else:
            possivel = None
            melhor_sim = 0
            for p in dados_pdf:
                sim_nome = similaridade(item.get("nome", ""), p.get("nome", ""))
                if sim_nome > melhor_sim:
                    melhor_sim = sim_nome
                    possivel = p

            motivo = "Nenhum parecido encontrado no PDF."
            # --- CORREÇÃO AQUI: NÃO USAR pdf_resp (fora do escopo) ---
            banco_possivel = ""
            if possivel:
                val_dif = abs((item.get("valor") or 0) - (possivel.get("valor") or 0))
                val_msg = "igual" if val_dif < 0.01 else ("próximo" if val_dif < 0.20 else "diferente")
                motivo = (
                    f"Nome semelhante encontrado: '{possivel.get('nome')}' "
                    f"(Sim={melhor_sim:.2f}), valor {val_msg} "
                    f"(R${(possivel.get('valor') or 0):.2f})."
                )
                banco_possivel = possivel.get("banco", "") or ""

            # preferir banco do 'possivel', senão manter qualquer banco já no item, senão vazio
            item["motivo"] = motivo
            item["banco_possivel"] = banco_possivel
            item["banco"] = banco_possivel or item.get("banco", "") or ""
            faltando_no_pdf.append(item)



    # ==========================================================
    # 4.2️⃣ PDF → Excel (sobras do PDF)
    # ==========================================================
    for i, p in enumerate(dados_pdf):
        if i not in usados_pdf:
            faltando_no_excel.append({
                "nome": p.get("nome"),
                "hora": normalizar_hora(p.get("hora")),
                "valor": round(p.get("valor", 0.0), 2),
                "data": p.get("data", ""),
                "banco": p.get("banco", "")
            })

    # ==========================================================
    # 5️⃣ RETORNO
    # ==========================================================
    return {
        "banco": ", ".join(bancos_detectados),
        "conferidos": conferidos,
        "faltando_no_pdf": faltando_no_pdf,
        "faltando_no_excel": faltando_no_excel,
        "meta": {
            "data_filtrada": selected_date.isoformat() if selected_date else None,
            "filtro_aplicado_no": "PDF",
            "criterios": {
                "tolerancia_valor": "exato",
                "similaridade_minima": 0.6,
                "tolerancia_hora": "±10min"
            }
        }
    }
