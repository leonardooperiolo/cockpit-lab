# -*- coding: utf-8 -*-
"""
cripto_arquivos.py
==================
Criptografia usada pelo painel: AES-256-GCM com chave derivada da senha (PBKDF2-SHA256, 600 mil voltas).

- Arquivos de dados: criptografados no SEU computador (preparar_planilha.py --senha) e enviados ao GitHub
  como NOME.enc. O GitHub Actions abre com a senha guardada no segredo DASHBOARD_SENHA.
- Painel (index.html): o gerador criptografa os dados com a mesma senha; o navegador pede a senha e abre.

Precisa da biblioteca 'cryptography':  pip install cryptography
"""
import base64
import gzip
import os
import sys
from functools import lru_cache

ITERACOES = 600_000
SAL_ARQUIVOS = b"cockpit-laboratorio/arquivos/v1"
MAGICO = b"CLP1"


def _lib():
    try:
        from cryptography.exceptions import InvalidTag
        from cryptography.hazmat.primitives import hashes
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
        from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    except ImportError:
        sys.exit("ERRO: falta a biblioteca 'cryptography'. Instale com:  pip install cryptography")
    return hashes, AESGCM, PBKDF2HMAC, InvalidTag


@lru_cache(maxsize=8)
def _chave(senha: str, sal: bytes, iteracoes: int) -> bytes:
    hashes, _, PBKDF2HMAC, _ = _lib()
    return PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=sal, iterations=iteracoes).derive(senha.encode("utf-8"))


def criptografar_arquivo(dados: bytes, senha: str) -> bytes:
    """bytes -> MAGICO + nonce(12) + texto cifrado."""
    _, AESGCM, _, _ = _lib()
    nonce = os.urandom(12)
    return MAGICO + nonce + AESGCM(_chave(senha, SAL_ARQUIVOS, ITERACOES)).encrypt(nonce, dados, None)


def abrir_arquivo(dados: bytes, senha: str):
    """Devolve os bytes originais, ou None se a senha estiver errada / o arquivo estiver corrompido."""
    _, AESGCM, _, InvalidTag = _lib()
    if dados[:4] != MAGICO or len(dados) < 32:
        return None
    try:
        return AESGCM(_chave(senha, SAL_ARQUIVOS, ITERACOES)).decrypt(dados[4:16], dados[16:], None)
    except InvalidTag:
        return None


def criptografar_para_pagina(texto: str, senha: str) -> dict:
    """Texto (JSON dos dados) -> objeto que o navegador sabe abrir (gzip + AES-GCM, base64)."""
    _, AESGCM, _, _ = _lib()
    sal, nonce = os.urandom(16), os.urandom(12)
    cifrado = AESGCM(_chave(senha, sal, ITERACOES)).encrypt(nonce, gzip.compress(texto.encode("utf-8"), 9), None)
    b64 = lambda b: base64.b64encode(b).decode("ascii")
    return {"enc": 1, "it": ITERACOES, "s": b64(sal), "n": b64(nonce), "c": b64(cifrado)}


def abrir_pagina(obj: dict, senha: str) -> str:
    """Inverso de criptografar_para_pagina (usado nos testes)."""
    _, AESGCM, _, _ = _lib()
    d = base64.b64decode
    return gzip.decompress(AESGCM(_chave(senha, d(obj["s"]), obj["it"])).decrypt(d(obj["n"]), d(obj["c"]), None)).decode("utf-8")
