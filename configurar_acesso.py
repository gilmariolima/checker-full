"""Gera configuração privada local sem exibir a senha no terminal."""

import getpass
import json
import secrets
from pathlib import Path

from acesso import password_hash


def main():
    arquivo = Path(__file__).resolve().parent / ".env"
    if arquivo.exists():
        raise SystemExit("Já existe um .env. Edite-o ou guarde uma cópia antes de gerar outro.")
    usuario = input("Usuário compartilhado dos donos: ").strip()
    if not usuario or len(usuario) > 128:
        raise SystemExit("Informe um usuário de até 128 caracteres.")
    senha = getpass.getpass("Senha (mínimo de 12 caracteres): ")
    if len(senha) < 12 or len(senha) > 1024:
        raise SystemExit("Use uma senha entre 12 e 1024 caracteres.")
    if senha != getpass.getpass("Confirme a senha: "):
        raise SystemExit("As senhas não coincidem.")
    valores = {
        "CHECKER_USERNAME": usuario,
        "CHECKER_PASSWORD_HASH": password_hash(senha),
        "CHECKER_SESSION_SECRET": secrets.token_urlsafe(48),
    }

    conta = Path(__file__).resolve().parent / "service-account.json"
    if conta.exists():
        info = json.loads(conta.read_text(encoding="utf-8"))
        if info.get("type") != "service_account" or "private_key" not in info:
            raise SystemExit("service-account.json não parece ser uma chave de conta de serviço.")
        # Uma linha só: json.loads no servidor entende, e o .env não quebra.
        valores["GOOGLE_SERVICE_ACCOUNT_JSON"] = json.dumps(info, ensure_ascii=False, separators=(",", ":"))
        valores["GOOGLE_SHEETS_ID"] = "1NGvUJ8Ssv8Ym79CSXoLaZclVGvcvxFwhAAMe8lmVSoA"

    with arquivo.open("x", encoding="utf-8") as stream:
        for chave, valor in valores.items():
            stream.write(chave + "=" + json.dumps(valor, ensure_ascii=False) + "\n")
    print("Configuração salva no .env (ignorado pelo Git). Copie os valores para as variáveis da Vercel, sem as aspas externas.")
    if "GOOGLE_SERVICE_ACCOUNT_JSON" in valores:
        print("A chave do Google foi incluída a partir de service-account.json.")
    else:
        print("service-account.json não encontrado: o Sheets fica indisponível até configurar GOOGLE_SERVICE_ACCOUNT_JSON.")
    print("A senha original não foi gravada. Veja DEPLOY.md para o deploy na Vercel.")


if __name__ == "__main__":
    main()
