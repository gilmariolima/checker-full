"""Acesso compartilhado: senha com hash e sessão assinada de oito horas."""

import hashlib
import hmac
import os
import secrets
import threading
import time
from collections import OrderedDict
from urllib.parse import urlsplit

from fastapi import APIRouter, Request
from fastapi.responses import JSONResponse, RedirectResponse
from itsdangerous import BadSignature, URLSafeTimedSerializer
from pydantic import BaseModel, Field

router = APIRouter()
COOKIE = "checker_session"
SESSION_SECONDS = 8 * 60 * 60
ITERATIONS = 600_000
_attempts = OrderedDict()
_lock = threading.Lock()


def password_hash(password):
    salt = secrets.token_hex(16)
    digest = hashlib.pbkdf2_hmac("sha256", password.encode(), salt.encode(), ITERATIONS).hex()
    return f"pbkdf2_sha256${ITERATIONS}${salt}${digest}"


def configured():
    return bool(os.getenv("CHECKER_USERNAME") and valid_hash(os.getenv("CHECKER_PASSWORD_HASH", ""))
                and len(os.getenv("CHECKER_SESSION_SECRET", "")) >= 32)


def valid_hash(value):
    try:
        algorithm, iterations, salt, digest = value.split("$")
        return (algorithm == "pbkdf2_sha256" and ITERATIONS <= int(iterations) <= 2_000_000
                and len(salt) >= 16 and len(bytes.fromhex(digest)) == 32)
    except (ValueError, TypeError):
        return False


def serializer():
    # Trocar usuário, hash da senha ou segredo invalida sessões anteriores.
    version = hashlib.sha256((os.environ["CHECKER_USERNAME"] + os.environ["CHECKER_PASSWORD_HASH"]).encode()).hexdigest()
    return URLSafeTimedSerializer(os.environ["CHECKER_SESSION_SECRET"], salt="checker-session-" + version)


def authenticated(request):
    if not configured():
        return False
    try:
        return serializer().loads(request.cookies.get(COOKIE, ""), max_age=SESSION_SECONDS) == "owner"
    except BadSignature:
        return False


async def protect(request, call_next):
    path = request.url.path
    public = {"/login", "/auth/login", "/static/style.css", "/static/login.js",
              "/static/icone.png", "/static/favicon.png"}
    if request.method not in ("GET", "HEAD", "OPTIONS"):
        origin = request.headers.get("origin")
        if (request.headers.get("sec-fetch-site") == "cross-site"
                or (origin and urlsplit(origin).netloc != request.url.netloc)):
            return JSONResponse({"erro": "Origem da solicitação não permitida."}, status_code=403)
    if path not in public and not authenticated(request):
        if path == "/":
            return RedirectResponse("/login", status_code=303)
        return JSONResponse({"erro": "Entre novamente para continuar."}, status_code=401,
                            headers={"Cache-Control": "no-store"})
    response = await call_next(request)
    response.headers["Cache-Control"] = "no-store"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "same-origin"
    return response


class Login(BaseModel):
    usuario: str = Field(min_length=1, max_length=128)
    senha: str = Field(min_length=1, max_length=1024)


def allow_attempt(key):
    """Limite local por instância; pode ser reforçado pelo firewall da hospedagem."""
    now = time.monotonic()
    with _lock:
        recent = [t for t in _attempts.pop(key, []) if now - t < 60]
        allowed = len(recent) < 5
        if allowed:
            recent.append(now)
        _attempts[key] = recent
        if len(_attempts) > 4096:
            _attempts.popitem(last=False)
        return allowed


@router.post("/auth/login")
def login(data: Login, request: Request):
    if not configured():
        return JSONResponse({"erro": "O acesso ainda não foi configurado pelo responsável pelo site."}, status_code=503)
    key = request.client.host if request.client else "unknown"
    if not allow_attempt(key):
        return JSONResponse({"erro": "Muitas tentativas. Aguarde um minuto e tente novamente."}, status_code=429,
                            headers={"Retry-After": "60"})
    _, iterations, salt, expected = os.environ["CHECKER_PASSWORD_HASH"].split("$")
    actual = hashlib.pbkdf2_hmac("sha256", data.senha.encode(), salt.encode(), int(iterations)).hex()
    password_ok = hmac.compare_digest(actual, expected)
    username_ok = hmac.compare_digest(data.usuario.encode(), os.environ["CHECKER_USERNAME"].encode())
    if not (password_ok and username_ok):
        return JSONResponse({"erro": "Usuário ou senha incorretos."}, status_code=401)
    response = JSONResponse({"ok": True})
    response.set_cookie(COOKIE, serializer().dumps("owner"), max_age=SESSION_SECONDS,
                        httponly=True, secure=bool(os.getenv("VERCEL")) or request.url.scheme == "https",
                        samesite="strict", path="/")
    return response


@router.post("/auth/logout")
def logout():
    response = JSONResponse({"ok": True})
    response.delete_cookie(COOKIE, path="/", httponly=True, samesite="strict")
    return response
