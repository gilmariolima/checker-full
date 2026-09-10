"""Leitura privada da base Google Sheets; credenciais somente no servidor."""

import json
import os
import re

from fastapi import APIRouter
from fastapi.responses import JSONResponse, Response
from google.auth.transport.requests import AuthorizedSession
from google.oauth2 import service_account
from pydantic import BaseModel
from datetime import date
from typing import Literal

from relatorio_sheets import AGENTES_AGENCIA, gerar_excel, montar_blocos, normalizar_nome

router = APIRouter()
DEFAULT_SHEET = "1NGvUJ8Ssv8Ym79CSXoLaZclVGvcvxFwhAAMe8lmVSoA"
SCOPE = "https://www.googleapis.com/auth/spreadsheets.readonly"


class ImportRequest(BaseModel):
    data: date
    tipo: Literal["geral", "agencia"] = "geral"


def ler_base(tipo):
    info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"])
    # Não aceitar endpoints alternativos vindos de um JSON mal configurado.
    info["token_uri"] = "https://oauth2.googleapis.com/token"
    credentials = service_account.Credentials.from_service_account_info(info, scopes=[SCOPE])
    sheet_id = os.getenv("GOOGLE_SHEETS_ID", DEFAULT_SHEET)
    if not re.fullmatch(r"[A-Za-z0-9_-]+", sheet_id):
        raise ValueError("ID inválido")
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}"
    with AuthorizedSession(credentials, refresh_timeout=20) as session:
        response = session.get(url, params={"fields": "sheets.properties.title"}, timeout=20)
        response.raise_for_status()
        nomes = [s["properties"]["title"] for s in response.json().get("sheets", [])
                 if s["properties"]["title"].upper().startswith("BASE_")]
        if tipo == "agencia":
            allowed = {normalizar_nome(n) for n in AGENTES_AGENCIA}
            nomes = [n for n in nomes if normalizar_nome(n[5:]) in allowed]
        abas = []
        for i in range(0, len(nomes), 30):
            lote = nomes[i:i + 30]
            ranges = ["'" + n.replace("'", "''") + "'!A3:G" for n in lote]
            response = session.get(url + "/values:batchGet", params={
                "ranges": ranges, "valueRenderOption": "UNFORMATTED_VALUE",
                "dateTimeRenderOption": "SERIAL_NUMBER",
            }, timeout=20)
            response.raise_for_status()
            for nome, values in zip(lote, response.json().get("valueRanges", [])):
                abas.append({"titulo": nome, "linhas": values.get("values", [])})
        return abas


@router.post("/api/sheets/importar")
def importar(data: ImportRequest):
    if not os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON"):
        return JSONResponse({"erro": "A conexão com o Google Sheets ainda não foi configurada. Você pode enviar uma planilha Excel."}, status_code=503)
    try:
        blocos = montar_blocos(ler_base(data.tipo), data.data.isoformat(), data.tipo)
        if not blocos:
            return JSONResponse({"erro": "Nenhum lançamento encontrado para a data e o grupo selecionados."}, status_code=404)
        arquivo = gerar_excel(blocos, data.data.isoformat())
        if len(arquivo) > 3_500_000:
            return JSONResponse({"erro": "O relatório é grande demais para a conferência. Tente o grupo Agência."}, status_code=413)
        return Response(arquivo,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Cache-Control": "no-store", "X-Agentes": str(len(blocos)),
                     "X-Lancamentos": str(sum(len(b["linhas"]) for b in blocos))})
    except Exception:
        # Não devolver erros de bibliotecas com dados/credenciais Google.
        return JSONResponse({"erro": "Não foi possível acessar a base. Verifique a conexão e a permissão da conta de serviço e tente novamente."}, status_code=502)
