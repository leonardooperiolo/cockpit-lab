#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_dashboard.py  -  Cockpit Executivo do Laboratório
=======================================================
Lê os arquivos da pasta data/ e escreve o index.html (a partir do template.html).

ARQUIVOS EM data/
  2020.csv.gz ... 2026.csv.gz   -> um arquivo por ano (nome = o ano)
  24092026.csv.gz               -> um arquivo por dia (nome = ddmmaaaa), enviado todo dia
  lista_medico.xlsx             -> col. A = código, col. B = nome
  (aceita também .csv, .zip com csv e .xlsx; maiúsculas/minúsculas não importam)

QUANDO OS ARQUIVOS SE SOBREPÕEM
  Para cada atendimento em cada dia, vale o arquivo mais recente (o diário vence o anual;
  entre dois diários, vence o de data maior). Assim reenviar um dia corrige o dia,
  sem contar em dobro.

COLUNAS USADAS DA PLANILHA DE EXAMES (as demais são descartadas):
  Convênio, Nome Convênio, Entrada, Código Requisição, Setor, Cód. Exame,
  Nome Exame, Nome Médico Solicitante (código), Total Bruto, Descontos, Líquido

REGRAS
  Unidade          = 2º bloco do Código Requisição (001-00X-........)
  Paciente/atend.  = Código Requisição único (cada cadastro = 1 atendimento)
  Nº de exames     = nº de linhas (Qtde é sempre 1)
  Faturamento      = coluna Líquido
  Ticket médio     = Líquido / atendimentos
  Ficam de fora    = unidades fora de UNIDADES, setores fora de SETORES, linhas sem código de exame

PRIVACIDADE: o nome do paciente NÃO é lido nem gravado no index.html.
"""
import atexit
import csv
import gzip
import json
import os
import re
import shutil
import statistics
import tempfile
import unicodedata
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

RAIZ = Path(__file__).resolve().parent
PASTA_DADOS = RAIZ / "data"
# Se existir o segredo DASHBOARD_SENHA (GitHub: Settings > Secrets and variables > Actions):
#   - arquivos NOME.enc em data/ são abertos com essa senha;
#   - o painel sai criptografado e pede a senha ao abrir.
SENHA = os.environ.get("DASHBOARD_SENHA", "").strip()
_TMP = []


def _abrir_enc(p: Path) -> Path:
    """Arquivo .enc -> arquivo aberto numa pasta temporária (com o nome sem .enc)."""
    if not p.name.lower().endswith(".enc"):
        return p
    if not SENHA:
        sys.exit(f"ERRO: {p.name} está criptografado, mas o segredo DASHBOARD_SENHA não está definido no GitHub "
                 "(Settings > Secrets and variables > Actions).")
    from cripto_arquivos import abrir_arquivo
    dados = abrir_arquivo(p.read_bytes(), SENHA)
    if dados is None:
        sys.exit(f"ERRO: não consegui abrir {p.name}: a senha do segredo DASHBOARD_SENHA está diferente da usada ao "
                 "criptografar, ou o arquivo está corrompido.")
    if not _TMP:
        _TMP.append(Path(tempfile.mkdtemp(prefix="cockpit_")))
        atexit.register(shutil.rmtree, _TMP[0], ignore_errors=True)
    destino = _TMP[0] / p.name[:-4]
    destino.write_bytes(dados)
    return destino
TEMPLATE = RAIZ / "template.html"
SAIDA = RAIZ / "index.html"

COLUNAS = [
    "Convênio", "Nome Convênio", "Entrada", "Código Requisição", "Setor",
    "Cód. Exame", "Nome Exame", "Nome Médico Solicitante",
    "Total Bruto", "Descontos", "Líquido",
]

# 2º bloco do Código Requisição -> nome da unidade
UNIDADES = {
    "001": "Matriz", "002": "Colônia", "003": "Quintino", "004": "Pinhão",
    "007": "Bonsucesso", "008": "C. Rocha", "010": "Pitanga", "013": "Pitanga 2",
}

# Setores (print 1)
SETORES = {
    1: "BIOQUÍMICA", 2: "HEMATOLOGIA", 3: "URINÁLISE", 4: "PARASITOLOGIA",
    5: "IMUNO E HORMÔNIOS", 6: "HORMÔNIOS-ENDOCRINOLOGIA", 7: "MARCADORES TUMORAIS",
    8: "MICROBIOLOGIA", 9: "MICOLOGIA", 10: "DROGAS DE ABUSO", 11: "IMUNOQUÍMICA",
    20: "COLETA", 30: "LABORATÓRIO APOIO", 31: "COAGULAÇÃO", 32: "TOXICOLOGIA",
    33: "TURBIDIMETRIA", 34: "HEMOGRAMA", 35: "RECEPÇÃO", 38: "GASOMETRIA",
    39: "INTERFACE", 40: "APOIO HOSPITALAR", 41: "TESTE RÁPIDO", 50: "HEMOCULTURAS",
    51: "CENTRIFUGAÇÃO",
}


# --------------------------------------------------------------------------- leitura
_EXT = (".enc", ".gz", ".csv", ".zip", ".xlsx")


def _tronco(nome: str) -> str:
    """'24092026.CSV.gz' -> '24092026'."""
    n = nome
    while True:
        base, ext = os.path.splitext(n)
        if ext.lower() in _EXT:
            n = base
        else:
            return n


def achar_arquivos_dados():
    """Devolve [(caminho, prioridade, rotulo)] em ordem crescente de prioridade.
    Nome = ano (2026) -> arquivo anual, prioridade 0.
    Nome = ddmmaaaa (24092026) -> arquivo do dia, prioridade maior quanto mais recente."""
    achados, avisos = [], []
    for p in sorted(PASTA_DADOS.iterdir()):
        nome = p.name
        if not p.is_file() or nome.startswith(".") or not nome.lower().endswith((".csv", ".gz", ".zip", ".xlsx", ".enc")):
            continue
        if nome.lower().startswith("lista_medico"):
            continue
        t = _tronco(nome)
        if re.fullmatch(r"(19|20)\d{2}", t):
            achados.append((_abrir_enc(p), 0, t, int(t)))
            continue
        m = re.fullmatch(r"(\d{2})[-_. ]?(\d{2})[-_. ]?((?:19|20)\d{2})", t)
        if m:
            try:
                d = datetime(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                achados.append((_abrir_enc(p), 1 + d.toordinal() * 10, d.strftime("%d/%m/%Y"), d.toordinal()))
                continue
            except ValueError:
                pass
        avisos.append(nome)
    for n in avisos:
        print(f"AVISO: '{n}' ignorado. O nome deve ser o ano (2026) ou a data ddmmaaaa (24092026).")
    if not achados:
        sys.exit("ERRO: não há arquivos de exames em data/. Coloque 2020.csv.gz, 2021.csv.gz... e os arquivos diários (ex.: 24092026.csv.gz).")
    achados.sort(key=lambda a: (a[1], a[3], a[0].name))
    # desempate: dois arquivos do mesmo dia (ex.: .csv e .csv.gz) -> o último em ordem alfabética vence
    fim, ult = [], None
    for k, (p, prio, rot, _) in enumerate(achados):
        if prio > 0 and prio == ult:
            prio = fim[-1][1] + 1
        fim.append((p, prio, rot))
        ult = achados[k][1]
    return fim


def _ler_csv(fonte):
    cabec = fonte.read(4096) if hasattr(fonte, "read") else b""
    if hasattr(fonte, "seek"):
        fonte.seek(0)
    primeira = cabec.decode("utf-8-sig", errors="ignore").splitlines()[0] if cabec else ""
    sep = ";" if primeira.count(";") > primeira.count(",") else ","
    texto = {c: str for c in ("Entrada", "Código Requisição", "Cód. Exame", "Nome Exame", "Nome Convênio",
                              "Total Bruto", "Descontos", "Líquido")}
    for enc in ("utf-8-sig", "latin-1"):
        try:
            if hasattr(fonte, "seek"):
                fonte.seek(0)
            return pd.read_csv(fonte, sep=sep, encoding=enc, dtype=texto,
                               usecols=lambda c: c.strip() in COLUNAS, low_memory=False)
        except UnicodeDecodeError:
            continue
    raise SystemExit("ERRO: não consegui ler o CSV (encoding).")


def carregar_planilha(caminho: Path) -> pd.DataFrame:
    """Devolve um DataFrame só com as colunas úteis (sem nome de paciente)."""
    nome = caminho.name.lower()
    if nome.endswith(".gz"):
        with gzip.open(caminho, "rb") as f:
            df = _ler_csv(f)
    elif nome.endswith(".csv"):
        with open(caminho, "rb") as f:
            df = _ler_csv(f)
    elif nome.endswith(".zip"):
        with zipfile.ZipFile(caminho) as z:
            alvo = [n for n in z.namelist() if n.lower().endswith(".csv")]
            if not alvo:
                sys.exit("ERRO: o .zip precisa conter um .csv.")
            with z.open(alvo[0]) as f:
                df = _ler_csv(f)
    else:
        from openpyxl import load_workbook
        wb = load_workbook(caminho, read_only=True, data_only=True)
        melhor = None
        for ws in wb.worksheets:
            it = ws.iter_rows(values_only=True)
            cab = next(it, None)
            if not cab or not set(COLUNAS) <= set(cab):
                continue  # ignora abas de tabela dinâmica / resumo
            idx = [cab.index(c) for c in COLUNAS]
            linhas = [tuple(r[i] for i in idx) for r in it if r and r[idx[3]] is not None]
            if melhor is None or len(linhas) > len(melhor):
                melhor = linhas  # a aba com mais linhas é a base
        if melhor is None:
            sys.exit(f"ERRO: nenhuma aba de {caminho.name} tem o cabeçalho esperado.")
        df = pd.DataFrame(melhor, columns=COLUNAS)

    df.columns = [str(c).strip() for c in df.columns]
    faltam = [c for c in COLUNAS if c not in df.columns]
    if faltam:
        sys.exit(f"ERRO: faltam colunas em {caminho.name}: {faltam}")
    return df[COLUNAS].copy()


def _tag(t):
    return t.split("}")[-1]


def _ler_lista_medicos_xml(caminho):
    """Leitor de reserva: o Excel às vezes salva em formato 'Strict Open XML',
    que o openpyxl não abre. Aqui lemos o XML direto (só biblioteca padrão)."""
    z = zipfile.ZipFile(caminho)
    ss = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")):
            ss.append("".join((t.text or "") for t in si.iter() if _tag(t.tag) == "t"))
    planilha = sorted(n for n in z.namelist() if n.startswith("xl/worksheets/sheet"))[0]
    linhas = []
    for row in ET.fromstring(z.read(planilha)).iter():
        if _tag(row.tag) != "row":
            continue
        d = {}
        for c in row:
            if _tag(c.tag) != "c":
                continue
            v = [x for x in c if _tag(x.tag) == "v"]
            if not v:
                continue
            valor = ss[int(v[0].text)] if c.get("t") == "s" else v[0].text
            d["".join(ch for ch in c.get("r") if ch.isalpha())] = valor
        linhas.append((d.get("A"), d.get("B")))
    return linhas[1:]  # pula cabeçalho


def carregar_filiais():
    """Devolve dict código (como string) -> nome da filial, lido de data/filiais.csv (ou
    .xlsx), col. A = FILCOD, col. B = nome. Cadastro simples e genérico: qualquer aba do
    painel que precise mostrar o nome de uma filial (não só o Financeiro) usa esse mesmo
    dicionário (dados["filiais"]) em vez de cada consulta ter que buscar/juntar isso
    de novo em alguma tabela do CONCENT."""
    pares = None
    pc = PASTA_DADOS / "filiais.csv"
    if not pc.exists() and (PASTA_DADOS / "filiais.csv.enc").exists():
        pc = _abrir_enc(PASTA_DADOS / "filiais.csv.enc")
    if pc.exists():
        pares = [(r[0], r[1]) for r in _linhas_arquivo(pc) if len(r) >= 2 and _cel(r[0]).strip()]
    else:
        p = PASTA_DADOS / "filiais.xlsx"
        if not p.exists() and (PASTA_DADOS / "filiais.xlsx.enc").exists():
            p = _abrir_enc(PASTA_DADOS / "filiais.xlsx.enc")
        if not p.exists():
            print("AVISO: data/filiais não encontrado - filiais aparecerão só pelo número.")
            return {}
        try:
            from openpyxl import load_workbook
            wb = load_workbook(p, read_only=True)
            if wb.worksheets:
                pares = [(r[0], r[1]) for r in wb.worksheets[0].iter_rows(min_row=2, values_only=True)]
        except Exception:
            pares = None
    mapa = {}
    for cod, nome in pares or []:
        cod = _cel(cod).strip()
        nome = re.sub(r"\s+", " ", str(nome or "")).strip()
        if not cod or not nome or not cod.isdigit():
            continue  # pula linha de cabeçalho (ex.: "FILCOD,FILNOME"), se vier uma
        mapa[cod] = nome
    return mapa


def carregar_medicos():
    """Devolve dict codigo -> [nomes distintos]. Códigos com mais de um nome ficam ambíguos."""
    pares = None
    pc = PASTA_DADOS / "lista_medico.csv"
    if not pc.exists() and (PASTA_DADOS / "lista_medico.csv.enc").exists():
        pc = _abrir_enc(PASTA_DADOS / "lista_medico.csv.enc")
    if pc.exists():
        pares = [(r[0], r[1]) for r in _linhas_arquivo(pc) if len(r) >= 2 and _cel(r[0]).strip()]
    else:
        p = PASTA_DADOS / "lista_medico.xlsx"
        if not p.exists() and (PASTA_DADOS / "lista_medico.xlsx.enc").exists():
            p = _abrir_enc(PASTA_DADOS / "lista_medico.xlsx.enc")
        if not p.exists():
            print("AVISO: data/lista_medico não encontrado - médicos aparecerão só pelo código.")
            return {}
        try:
            from openpyxl import load_workbook
            wb = load_workbook(p, read_only=True)
            if wb.worksheets:
                pares = [(r[0], r[1]) for r in wb.worksheets[0].iter_rows(min_row=2, values_only=True)]
        except Exception:
            pares = None
        if not pares:
            pares = _ler_lista_medicos_xml(p)
    mapa = {}
    for cod, nome in pares:
        try:
            cod = int(float(cod))
        except (TypeError, ValueError):
            continue
        nome = re.sub(r"\s+", " ", str(nome or "")).strip()
        if not nome:
            continue
        lst = mapa.setdefault(cod, [])
        if nome not in lst:
            lst.append(nome)
    return mapa


# --------------------------------------------------------------------------- tratamento
def _datas(ent: pd.Series) -> pd.Series:
    """Converte a coluna Entrada. Aceita data do Excel, número serial, texto ISO
    (2023-08-31 10:00:00) e texto brasileiro com ano de 4 ou 2 dígitos
    (31/08/2023 10:00, 31/08/23 10:00)."""
    if pd.api.types.is_datetime64_any_dtype(ent):
        return ent
    if pd.api.types.is_numeric_dtype(ent):
        return pd.to_datetime(ent, unit="D", origin="1899-12-30", errors="coerce")
    s = ent.astype(str).str.strip()
    res = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    regras = (
        (r"^\d{4}-\d{2}-\d{2}", ["ISO8601"]),
        (r"^\d{2}/\d{2}/\d{4}", ["%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y"]),
        (r"^\d{2}/\d{2}/\d{2}(?!\d)", ["%d/%m/%y %H:%M:%S", "%d/%m/%y %H:%M", "%d/%m/%y"]),
    )
    def _atribuir(mask, valores):
        """Como res[mask] = valores, mas uma data absurda (ex.: ano 0001, vinda de
        registro antigo com campo em branco no sistema de origem) vira NaT em vez
        de travar a leitura do arquivo inteiro."""
        try:
            res[mask] = valores
        except (pd.errors.OutOfBoundsDatetime, OverflowError, ValueError):
            valores = pd.to_datetime(valores, errors="coerce")
            dentro = valores.notna() & (valores >= pd.Timestamp.min + pd.Timedelta(days=1)) & (valores <= pd.Timestamp.max - pd.Timedelta(days=1))
            valores = valores.where(dentro)
            res[mask] = valores

    for rx, formatos in regras:
        casa = s.str.match(rx)
        for f in formatos:
            falta = casa & res.isna()
            if not falta.any():
                break
            _atribuir(falta, pd.to_datetime(s[falta], format=f, errors="coerce"))
    resto = res.isna() & s.ne("") & s.ne("nan")
    if resto.any():
        _atribuir(resto, pd.to_datetime(s[resto], errors="coerce", format="mixed", dayfirst=True))
    return res


def _num(sr: pd.Series) -> pd.Series:
    """Número no padrão brasileiro ('3.473,00', '59', '32,00') ou com ponto decimal."""
    if pd.api.types.is_numeric_dtype(sr):
        return pd.to_numeric(sr, errors="coerce").fillna(0.0)
    t = sr.astype(str).str.strip()
    br = t.str.contains(",", regex=False)
    t = t.where(~br, t.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(t, errors="coerce").fillna(0.0)


def preparar(df: pd.DataFrame, rotulo: str = "") -> pd.DataFrame:
    df = df.copy()
    df["Entrada"] = _datas(df["Entrada"])
    ruins = df["Entrada"].isna().sum()
    if ruins:
        print(f"AVISO: {rotulo} {ruins} linhas sem data válida foram ignoradas.")
    df = df[df["Entrada"].notna()].copy()

    df["dia"] = df["Entrada"].dt.normalize()
    df["req"] = df["Código Requisição"].astype(str).str.strip()
    partes = df["req"].str.extract(r"^(\d{3})[-.](\d{3})")
    df["uni"] = partes[1].fillna("???")

    for c in ("Total Bruto", "Descontos", "Líquido"):
        df[c] = _num(df[c])
    df["Convênio"] = pd.to_numeric(df["Convênio"], errors="coerce").fillna(-1).astype(int)
    df["Setor"] = pd.to_numeric(df["Setor"], errors="coerce").fillna(-1).astype(int)
    df["med"] = pd.to_numeric(df["Nome Médico Solicitante"], errors="coerce").fillna(-1).astype(int)
    df["Nome Convênio"] = df["Nome Convênio"].astype(str).str.strip()
    df["Nome Exame"] = df["Nome Exame"].astype(str).str.strip()
    cod = df["Cód. Exame"].astype(object).where(df["Cód. Exame"].notna(), None)
    df["exame"] = [str(c).strip() if c is not None else "S/COD" for c in cod]
    return df.drop(columns=["Código Requisição", "Nome Médico Solicitante", "Cód. Exame"])


def resolver_sobreposicao(df: pd.DataFrame):
    """Para cada (atendimento, dia) vale só o arquivo mais recente (maior 'prio')."""
    if df["prio"].nunique() <= 1:
        return df, 0
    mx = df.groupby(["req", "dia"], sort=False)["prio"].transform("max")
    manter = df["prio"] == mx
    return df[manter].copy(), int((~manter).sum())


def aplicar_exclusoes(df: pd.DataFrame):
    """Regras do laboratório: ficam de fora do painel
      - unidades que não existem mais (código fora de UNIDADES);
      - setores fora da lista SETORES (são taxas de coleta domiciliar, não exames);
      - linhas sem código de exame (mesmo caso das taxas).
    Tudo o que sai é somado e mostrado na aba Qualidade dos Dados."""
    m_uni = ~df["uni"].isin(UNIDADES)
    m_set = ~df["Setor"].isin(SETORES)
    m_cod = df["exame"] == "S/COD"
    fora = m_uni | m_set | m_cod
    liq = lambda m: round(float(df.loc[m, "Líquido"].sum()), 2)
    excl = {
        "linhas": int(fora.sum()),
        "valor": liq(fora),
        "unidades": [{"cod": u, "linhas": int((df["uni"] == u).sum()), "valor": liq(df["uni"] == u)}
                     for u in sorted(df.loc[m_uni, "uni"].unique())],
        "setores": [{"id": int(sid), "linhas": int((df["Setor"] == sid).sum()), "valor": liq(df["Setor"] == sid),
                     "nome": str(df.loc[df["Setor"] == sid, "Nome Exame"].mode().iloc[0])}
                    for sid in sorted(df.loc[m_set, "Setor"].unique())],
        "sem_codigo": {"linhas": int(m_cod.sum()), "valor": liq(m_cod)},
    }
    if excl["linhas"]:
        print(f"  {excl['linhas']:,} linhas desconsideradas (unidade extinta / taxas sem código de exame)".replace(",", "."))
    return df[~fora].copy(), excl


def construir_dados(df: pd.DataFrame, medicos: dict, excl: dict, arquivos: list, substituidas: int) -> dict:
    # ---------- dimensões
    dias = sorted(df["dia"].unique())
    dia_idx = {d: i for i, d in enumerate(dias)}
    meses = sorted({d.strftime("%Y-%m") for d in dias})
    mes_idx = {m: i for i, m in enumerate(meses)}

    unis = sorted(df["uni"].unique())
    uni_dim = [{"cod": u, "nome": UNIDADES.get(u, f"Unidade {u} (não mapeada)"),
                "ok": u in UNIDADES} for u in unis]
    uni_idx = {u: i for i, u in enumerate(unis)}

    conv = (df.groupby("Convênio")["Nome Convênio"].first().sort_index().map(_nome_curto))
    repetidos = conv.value_counts()
    repetidos = set(repetidos[repetidos > 1].index)
    conv_dim = [{"id": int(i), "nome": (f"{n} ({i})" if n in repetidos else n)} for i, n in conv.items()]
    conv_idx = {c["id"]: i for i, c in enumerate(conv_dim)}

    volume_med = df.groupby("med").size()
    med_dim, med_idx = [], {}
    n_sem_nome = n_ambiguo = 0
    for cod in sorted(volume_med.index):
        nomes = medicos.get(int(cod), [])
        if not nomes:
            nome, q = f"Cód. {cod} (sem nome na lista)", 1
        elif len(nomes) == 1:
            nome, q = nomes[0], 0
        else:
            extra = f" +{len(nomes) - 2}" if len(nomes) > 2 else ""
            nome, q = f"{nomes[0]} / {nomes[1]}{extra} (cód. {cod}, ambíguo)", 2
        med_idx[int(cod)] = len(med_dim)
        med_dim.append({"id": int(cod), "nome": nome, "q": q})

    setores_usados = sorted(df["Setor"].unique())
    set_dim = [{"id": int(s), "nome": SETORES.get(int(s), f"Setor {s} (não mapeado)")} for s in setores_usados]
    set_idx = {s["id"]: i for i, s in enumerate(set_dim)}

    ex = df.groupby("exame").agg(nome=("Nome Exame", "first"), setor=("Setor", "first")).reset_index()
    ex = ex.sort_values("exame")
    ex_dim = [{"cod": r.exame, "nome": r.nome, "s": set_idx[int(r.setor)]} for r in ex.itertuples()]
    ex_idx = {r["cod"]: i for i, r in enumerate(ex_dim)}

    # ---------- fato A: dia x unidade x convênio x médico
    # linhas/valores usam os atributos de cada linha; o nº de atendimentos é atribuído
    # à 1ª linha de cada requisição, para o total geral nunca contar em dobro.
    chave = ["dia", "uni", "Convênio", "med"]
    a = df.groupby(chave).agg(e=("req", "size"), b=("Total Bruto", "sum"),
                              x=("Descontos", "sum"), l=("Líquido", "sum")).reset_index()
    prim = df.sort_values(["req", "Entrada"], kind="stable").drop_duplicates("req")
    r = prim.groupby(chave).size().rename("r").reset_index()
    a = a.merge(r, on=chave, how="outer").fillna({"e": 0, "b": 0.0, "x": 0.0, "l": 0.0, "r": 0})
    a = a.sort_values(chave)
    A = {
        "d": [dia_idx[d] for d in a["dia"]],
        "u": [uni_idx[u] for u in a["uni"]],
        "c": [conv_idx[int(c)] for c in a["Convênio"]],
        "m": [med_idx[int(m)] for m in a["med"]],
        "r": [int(v) for v in a["r"]],
        "e": [int(v) for v in a["e"]],
        "b": [round(float(v), 2) for v in a["b"]],
        "x": [round(float(v), 2) for v in a["x"]],
        "l": [round(float(v), 2) for v in a["l"]],
    }

    # ---------- fato B: mês x unidade x convênio x exame (aba Exames & Setores)
    df["ym"] = df["dia"].dt.strftime("%Y-%m")
    b = df.groupby(["ym", "uni", "Convênio", "exame"]).agg(e=("req", "size"), l=("Líquido", "sum")).reset_index()
    B = {
        "mo": [mes_idx[v] for v in b["ym"]],
        "u": [uni_idx[v] for v in b["uni"]],
        "c": [conv_idx[int(v)] for v in b["Convênio"]],
        "x": [ex_idx[v] for v in b["exame"]],
        "e": [int(v) for v in b["e"]],
        "l": [round(float(v), 2) for v in b["l"]],
    }

    # ---------- qualidade dos dados
    total_l, total_r = len(df), float(df["Líquido"].sum())
    q_med = df["med"].map(lambda c: med_dim[med_idx[int(c)]]["q"])
    por_req = df.groupby("req").agg(nc=("Convênio", "nunique"), nm=("med", "nunique"),
                                    nd=("dia", "nunique"))
    qual = {
        "linhas": int(total_l),
        "requisicoes": int(len(por_req)),
        "pct_linhas_medico_sem_nome": round(100 * float((q_med == 1).mean()), 1),
        "pct_linhas_medico_ambiguo": round(100 * float((q_med == 2).mean()), 1),
        "pct_fat_medico_sem_nome": round(100 * float(df.loc[q_med == 1, "Líquido"].sum()) / total_r, 1) if total_r else 0,
        "medicos_usados": int(len(med_dim)),
        "medicos_sem_nome": int(sum(1 for m in med_dim if m["q"] == 1)),
        "excluidos": excl,
        "arquivos": arquivos,
        "linhas_substituidas": substituidas,
        "linhas_valor_zero": int((df["Líquido"] == 0).sum()),
        "zero_por_convenio": [{"nome": conv_dim[conv_idx[int(c)]]["nome"], "linhas": int(n)}
                              for c, n in df.loc[df["Líquido"] == 0, "Convênio"].value_counts().head(6).items()],
        "req_multi_convenio": int((por_req["nc"] > 1).sum()),
        "req_multi_medico": int((por_req["nm"] > 1).sum()),
        "req_multi_dia": int((por_req["nd"] > 1).sum()),
        "convenios_nome_repetido": sorted(repetidos),
    }

    agora = datetime.now()
    try:
        from zoneinfo import ZoneInfo
        agora = datetime.now(ZoneInfo("America/Sao_Paulo"))
    except Exception:
        pass
    return {
        "meta": {
            "geradoEm": agora.strftime("%d/%m/%Y %H:%M"),
            "dataMin": dias[0].strftime("%Y-%m-%d"),
            "dataMax": dias[-1].strftime("%Y-%m-%d"),
        },
        "dias": [d.strftime("%Y-%m-%d") for d in dias],
        "meses": meses,
        "unidades": uni_dim, "convenios": conv_dim, "medicos": med_dim,
        "setores": set_dim, "exames": ex_dim,
        "A": A, "B": B, "qualidade": qual,
    }


# --------------------------------------------------------------------------- laboratório de apoio
PASTA_APOIO = PASTA_DADOS / "apoio"
# papel -> como o nome do arquivo começa (sem acento, minúsculas, espaço = _)
ARQ_APOIO = {
    "materiais": ("lista_materiais", "materiais"),   # opcional: código -> nome do material
    "precos_conv": ("preco_de_exame_por_convenio", "precos_convenio", "preco_convenio", "precos_por_convenio"),  # opcional
    "prazos_db": ("amostra_do_portf", "portfolio_de_exames", "portfolio_exames", "prazo_db", "prazos_db"),  # opcional
    "lista": ("lista",),
    "de_db": ("de_para_concent_db",),
    "de_hp": ("de_para_concent_hp",),
    "tab_db": ("tabela_db",),
    "tab_hp": ("tabela_pardini", "tabela_hp", "tabela_padini"),
}
APOIO_OPCIONAL = {"materiais", "precos_conv", "prazos_db"}
NOME_PAPEL = {
    "materiais": "Lista de materiais biológicos",
    "precos_conv": "Relatório de convênios por exame (preços)",
    "prazos_db": "Portfólio de exames e prazos do DB",
    "lista": "Lista de exames CONCENT (lab. apoio)", "de_db": "De-para CONCENT → DB",
    "de_hp": "De-para CONCENT → HP", "tab_db": "Tabela de preços DB", "tab_hp": "Tabela de preços HP (Pardini)",
}


def _sa(s) -> str:
    """minúsculas, sem acento, sem espaços nas pontas."""
    s = unicodedata.normalize("NFD", str(s)).encode("ascii", "ignore").decode()
    return s.strip().lower()


def _cel(v) -> str:
    """Célula -> texto (125.0 vira '125'; vazio vira '')."""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _nome_curto(s):
    """Nomes de convênio/cliente/fornecedor às vezes vêm do sistema com um preenchimento
    de pontos no final (campo de tamanho fixo do legado), tipo 'MEDPREV - ...................'
    - corta tudo a partir do ' - ' quando o que sobra depois é só pontuação de
    preenchimento, deixando só o nome de verdade. Nome sem esse padrão não é alterado."""
    if not s:
        return s
    t = s.rstrip(". ")
    if t.endswith("-"):
        return t.rstrip("- ").strip() or s.strip()
    return s.strip()


def _preco(v):
    """'1.234,56', '14,52', 45.78 -> float; vazio/ruim -> None."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return round(float(v), 2)
    t = re.sub(r"[^\d,.\-]", "", str(v))
    if not t:
        return None
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return round(float(t), 2)
    except ValueError:
        return None


def _texto_arquivo(p: Path) -> str:
    b = gzip.open(p, "rb").read() if p.name.lower().endswith(".gz") else p.read_bytes()
    try:
        return b.decode("utf-8-sig")
    except UnicodeDecodeError:
        return b.decode("latin-1")


def _linhas_arquivo(p: Path):
    if p.name.lower().endswith(".xlsx"):
        from openpyxl import load_workbook
        ws = load_workbook(p, read_only=True, data_only=True).worksheets[0]
        return [list(r) for r in ws.iter_rows(values_only=True)]
    linhas = _texto_arquivo(p).splitlines()
    sep = ";" if linhas and linhas[0].count(";") >= linhas[0].count(",") else ","
    return [r for r in csv.reader(linhas, delimiter=sep)]


def achar_arquivos_apoio():
    achados = {}
    if not PASTA_APOIO.is_dir():
        return achados
    for p in sorted(PASTA_APOIO.iterdir()):
        if not p.is_file() or p.name.startswith(".") or not p.name.lower().endswith((".xlsx", ".csv", ".gz", ".pdf", ".enc")):
            continue
        nome = _sa(p.name).replace(" ", "_")
        for papel, inicios in ARQ_APOIO.items():
            if papel not in achados and nome.startswith(inicios):
                achados[papel] = _abrir_enc(p)
                break
    return achados


def _ler_lista_concent(p: Path):
    """Lista de exames do CONCENT. Tolera ';' dentro do nome do exame."""
    out = []
    if p.name.lower().endswith(".xlsx"):
        linhas = _linhas_arquivo(p)
        cab = [_sa(c) for c in linhas[0]]
        ic = next(i for i, c in enumerate(cab) if c.startswith("cod"))
        inm = next(i for i, c in enumerate(cab) if c.startswith("nome"))
        ip = next((i for i, c in enumerate(cab) if c.startswith("prazo")), None)
        for r in linhas[1:]:
            if r and _cel(r[ic]):
                out.append((_cel(r[ic]), _cel(r[inm]), _cel(r[ip]) if ip is not None else ""))
    else:
        for ln in _texto_arquivo(p).splitlines()[1:]:
            partes = ln.split(";")
            if partes and partes[-1].strip() == "":
                partes = partes[:-1]
            if len(partes) < 4 or not partes[0].strip():
                continue
            out.append((partes[0].strip(), ";".join(partes[1:-2]).strip(), partes[-1].strip()))
    return out


def _ler_depara(p: Path):
    linhas = _linhas_arquivo(p)
    cab = [_sa(c) for c in linhas[0]]
    ie = cab.index("exame")
    ia = cab.index("exame no apoio")
    im = cab.index("material")
    nomes = [i for i, c in enumerate(cab) if c == "nome"]
    inc = cab.index("nome concent") if "nome concent" in cab else nomes[0]
    ina = nomes[-1]
    idesc = cab.index("descricao") if "descricao" in cab else None
    out = []
    for r in linhas[1:]:
        if not r or not _cel(r[ie]) or not _cel(r[ia]):
            continue
        out.append({"cod": _cel(r[ie]), "nome_c": _cel(r[inc]), "mat": _cel(r[im]) or "0",
                    "apo": _cel(r[ia]), "nome_ap": _cel(r[ina]) if ina != inc else "",
                    "desc": _cel(r[idesc]) if idesc is not None else ""})
    return out


def _ler_precos_convenio(p: Path):
    """CSV exportado direto do banco do CONCENT (convênio, plano, exame, valor, ativo).
    Devolve uma lista de (cod_exame_concent, nome_exame, cod_convenio, nome_convenio,
    ativo_bool, valor)."""
    saida = []
    for r in _linhas_arquivo(p):
        if len(r) < 8:
            continue
        cod_conv = _cel(r[0]).rstrip(".").strip()
        nome_conv = _nome_curto(_cel(r[1]).strip())
        cod_exame = _cel(r[4]).strip()
        nome_exame = _cel(r[5]).strip()
        valor = _preco(_cel(r[6]))
        ativo = _cel(r[7]).strip().upper() == "S"
        if not cod_conv.isdigit() or not cod_exame or not nome_exame:
            continue
        saida.append((cod_exame, nome_exame, int(cod_conv), nome_conv, ativo, valor or 0.0))
    return saida


def _ler_materiais(p: Path):
    """Lista de materiais biológicos (código;nome) -> {código: nome}."""
    out = {}
    for r in _linhas_arquivo(p):
        if len(r) >= 2 and _cel(r[0]).isdigit() and _cel(r[1]):
            out[_cel(r[0])] = re.sub(r"\s+", " ", _cel(r[1]))
    return out


def _ler_prazos_db(p: Path):
    """Portfólio de exames do DB (código/mnemônico -> prazo de entrega, ex.: '2 dias úteis')."""
    out = {}
    linhas = _linhas_arquivo(p)
    for r in linhas[1:] if linhas else []:
        if len(r) >= 4 and _cel(r[0]) and _cel(r[3]):
            out[_cel(r[0])] = re.sub(r"\s+", " ", _cel(r[3]))
    return out


def _ler_precos(p: Path):
    """Tabela de preços: devolve {código: {'precos': [..], 'nome': ..}} e a lista de códigos repetidos."""
    linhas = _linhas_arquivo(p)
    h = next(i for i, r in enumerate(linhas)
             if r and _sa(_cel(r[0])) == "codigo" and any(_sa(_cel(c)).startswith("preco") for c in r))
    cab = [_sa(_cel(c)) for c in linhas[h]]
    ip = next(i for i, c in enumerate(cab) if c.startswith("preco"))
    tab = {}
    for r in linhas[h + 1:]:
        if not r or len(r) <= ip or not _cel(r[0]):
            continue
        v = _preco(r[ip])
        if v is None:
            continue
        d = tab.setdefault(_cel(r[0]).upper(), {"cod": _cel(r[0]), "nome": _cel(r[1]) if len(r) > 1 else "", "precos": []})
        if v not in d["precos"]:
            d["precos"].append(v)
    return tab


def _opcoes_apoio(depara, tabela):
    """{cód. CONCENT: [{mat, apo, nome, p, e}]} sem repetir (material, código do apoio)."""
    out, vistos = {}, set()
    for r in depara:
        chave = (r["cod"], r["mat"], r["apo"].upper())
        if chave in vistos:
            continue
        vistos.add(chave)
        t = tabela.get(r["apo"].upper())
        ps = sorted(p for p in t["precos"] if p > 0) if t else []
        out.setdefault(r["cod"], []).append({
            "mat": r["mat"], "apo": r["apo"], "nome": r["nome_ap"] or (t["nome"] if t else ""),
            "p": [ps[0], ps[-1]] if ps else None,
            "e": None if ps else ("z" if t else "s")})   # z = R$ 0,00 na tabela; s = código sem preço na tabela
    return out


def _cruzar(mat, d, h):
    n = max(len(d), len(h), 1)
    return [(mat, d[i] if i < len(d) else None, h[i] if i < len(h) else None) for i in range(n)]


def _variantes(dopts, hopts):
    """Junta as opções do DB e do HP do mesmo exame, casando pelo material
    (material 0 = não especificado, vale para qualquer material)."""
    mats = sorted({o["mat"] for o in dopts + hopts if o["mat"] != "0"}, key=lambda x: (len(x), x))
    dw = [o for o in dopts if o["mat"] == "0"]
    hw = [o for o in hopts if o["mat"] == "0"]
    if not mats:
        return _cruzar("0", dw, hw)
    out, usou_d, usou_h = [], False, False
    for m in mats:
        d = [o for o in dopts if o["mat"] == m]
        h = [o for o in hopts if o["mat"] == m]
        if not d and dw:
            d, usou_d = dw, True
        if not h and hw:
            h, usou_h = hw, True
        out += _cruzar(m, d, h)
    resto_d = [] if usou_d else dw
    resto_h = [] if usou_h else hw
    if resto_d or resto_h:
        out += _cruzar("0", resto_d, resto_h)
    return out


def _comparar(d, h):
    """Quem cobra menos. Só decide quando não há dúvida (preços repetidos na tabela podem se sobrepor)."""
    pd_ = d["p"] if d else None
    ph = h["p"] if h else None
    if pd_ and ph:
        if pd_[1] < ph[0]:
            return "d", ph[0] - pd_[1], (ph[0] - pd_[1]) / ph[0] * 100
        if ph[1] < pd_[0]:
            return "h", pd_[0] - ph[1], (pd_[0] - ph[1]) / pd_[0] * 100
        if pd_[0] == pd_[1] == ph[0] == ph[1]:
            return "=", 0.0, 0.0
        return "?", None, None
    if pd_:
        return "sd", None, None
    if ph:
        return "sh", None, None
    return ("sp" if (d or h) else "nm"), None, None


def carregar_apoio():
    arqs = achar_arquivos_apoio()
    faltam = [NOME_PAPEL[k] for k in ARQ_APOIO if k not in arqs and k not in APOIO_OPCIONAL]
    if faltam:
        if arqs:
            print("AVISO: laboratório de apoio ignorado; faltam em data/apoio/: " + "; ".join(faltam))
        return None
    lista = _ler_lista_concent(arqs["lista"])
    dep_db, dep_hp = _ler_depara(arqs["de_db"]), _ler_depara(arqs["de_hp"])
    tab_db, tab_hp = _ler_precos(arqs["tab_db"]), _ler_precos(arqs["tab_hp"])
    op_db, op_hp = _opcoes_apoio(dep_db, tab_db), _opcoes_apoio(dep_hp, tab_hp)

    mat_nome = _ler_materiais(arqs["materiais"]) if "materiais" in arqs else {}
    prazos_db = _ler_prazos_db(arqs["prazos_db"]) if "prazos_db" in arqs else {}
    for r in dep_hp:   # reserva: descrição do material no de-para do HP
        if r["desc"] and r["mat"] not in mat_nome:
            mat_nome[r["mat"]] = r["desc"]
    nomes_c = {r["cod"]: r["nome_c"] for r in dep_db + dep_hp}
    na_lista = {c for c, _, _ in lista}
    exames = [(c, n, pz, 1) for c, n, pz in lista]
    extras = sorted((set(op_db) | set(op_hp)) - na_lista)
    exames += [(c, nomes_c.get(c, c), "", 0) for c in extras]

    def lado(o, pz_map=None):
        if not o:
            return None
        d = {"c": o["apo"], "n": o["nome"], "p": o["p"]}
        if o["e"]:
            d["e"] = o["e"]
        if pz_map:
            pz = pz_map.get(o["apo"])
            if pz:
                d["pz"] = pz
        return d

    linhas = []
    for cod, nome, pz, na in exames:
        do_exame = {}   # materiais diferentes com o mesmo resultado viram uma linha só
        for mat, d, h in _variantes(op_db.get(cod, []), op_hp.get(cod, [])):
            w, x, y = _comparar(d, h)
            rot = mat_nome.get(mat, f"Material {mat}") if mat != "0" else ""
            chave = json.dumps([lado(d), lado(h)], sort_keys=True)
            if chave in do_exame:
                if rot and rot not in do_exame[chave]["m"].split(" · "):
                    do_exame[chave]["m"] = (do_exame[chave]["m"] + " · " + rot) if do_exame[chave]["m"] else rot
                continue
            do_exame[chave] = {"c": cod, "n": nome, "m": rot, "l": na, "d": lado(d, prazos_db), "h": lado(h), "w": w,
                               "x": None if x is None else round(x, 2), "y": None if y is None else round(y, 1)}
        linhas += list(do_exame.values())
    linhas.sort(key=lambda r: (_sa(r["n"]), r["c"], r["m"]))

    # ---- pontos de atenção nos cadastros
    dup_hp = [{"c": t["cod"], "n": t["nome"], "p": sorted(t["precos"])} for t in tab_hp.values() if len(t["precos"]) > 1]
    dup_hp.sort(key=lambda x: x["c"])
    mesmo_mat = []
    for nome, op in (("DB", op_db), ("HP", op_hp)):
        for cod, ops in op.items():
            por_mat = {}
            for o in ops:
                por_mat.setdefault(o["mat"], []).append(o["apo"])
            for m, a in por_mat.items():
                if len(a) > 1:
                    mesmo_mat.append({"l": nome, "c": cod, "n": nomes_c.get(cod, cod), "a": a})
    sem_mapa = [{"c": c, "n": n} for c, n, _, na in exames if na and c not in op_db and c not in op_hp]
    zero_map = sorted({o["apo"] for ops in op_db.values() for o in ops if o["e"] == "z"} |
                      {o["apo"] for ops in op_hp.values() for o in ops if o["e"] == "z"})
    sem_preco = {"DB": sorted({o["apo"] for ops in op_db.values() for o in ops if o["e"] == "s"}),
                 "HP": sorted({o["apo"] for ops in op_hp.values() for o in ops if o["e"] == "s"})}
    usados_m = {o["mat"] for op in (op_db, op_hp) for ops in op.values() for o in ops if o["mat"] != "0"}
    mat_sem_nome = sorted((m for m in usados_m if m not in mat_nome), key=lambda x: (len(x), x))
    na_l = [r for r in linhas if r["l"]]
    cont = {k: sum(1 for r in na_l if r["w"] == k) for k in ("d", "h", "=", "?", "sd", "sh", "sp", "nm")}
    cods_l = {r["c"] for r in na_l}
    tem_db = {r["c"] for r in na_l if r["d"] and r["d"]["p"]}
    tem_hp = {r["c"] for r in na_l if r["h"] and r["h"]["p"]}
    return {
        "rows": linhas,
        "arquivos": [{"papel": NOME_PAPEL[k], "nome": arqs[k].name} for k in ARQ_APOIO if k in arqs],
        "resumo": {"exames": len(cods_l), "extras": len(extras), "linhas": len(na_l), "com_db": len(tem_db), "com_hp": len(tem_hp),
                   "ambos": len(tem_db & tem_hp), "cont": cont},
        "atencao": {"sem_mapa": sem_mapa, "dup_hp": dup_hp, "mesmo_mat": mesmo_mat, "zero_map": zero_map, "sem_preco": sem_preco,
                    "extras": [{"c": c, "n": nomes_c.get(c, c)} for c in extras], "mat_sem_nome": mat_sem_nome},
    }


# --------------------------------------------------------------------------- área técnica (prazo de entrega)
# Exceção de privacidade: esta é a ÚNICA aba do painel que mostra nome de
# paciente (campo "paciente" abaixo). Decisão combinada: o objetivo aqui é
# o técnico achar a amostra física na bancada, e sem nome isso não dá pra
# fazer. Em nenhum outro lugar do dashboard o nome de paciente aparece.
#
# O status (Não Coletado / Triado / Digitado / Digitado Parcial / Conferido)
# não existe como campo pronto no banco - é calculado aqui a partir de flags
# espalhados em 3 tabelas (ver PLANO_extracao_area_tecnica.md para o
# levantamento). A extração (coletar_area_tecnica.py, roda só na VM do
# Leo) já filtra para "ainda não liberado, não cancelado, entrada nos
# últimos ~120 dias" - ou seja, o CSV já É o backlog em aberto, não um
# recorte por prazo; aqui só resta calcular o status e separar quem está
# aguardando coleta (sem prazo ainda, nada a cobrar) de quem já está em
# andamento (tem prazo, pode estar atrasado).
PASTA_AREA_TEC = PASTA_DADOS / "area_tecnica"
ARQ_AREA_TEC = {
    "exames": ("area_tecnica_exames",),
    "digitados": ("area_tecnica_digitados",),
    "obrigatorios": ("area_tecnica_obrigatorios",),
}
NOME_PAPEL_AT = {
    "exames": "Exames pendentes (área técnica)",
    "digitados": "Atributos digitados por exame",
    "obrigatorios": "Atributos obrigatórios por tipo de exame",
}
STATUS_AT = {
    "nc": "Não Coletado", "t": "Triado", "col": "Coletado",
    "dp": "Digitado Parcial", "d": "Digitado", "c": "Conferido",
}


def achar_arquivos_area_tecnica():
    achados = {}
    if not PASTA_AREA_TEC.is_dir():
        return achados
    for p in sorted(PASTA_AREA_TEC.iterdir()):
        if not p.is_file() or p.name.startswith(".") or not p.name.lower().endswith((".csv", ".gz", ".enc")):
            continue
        nome = _sa(p.name).replace(" ", "_")
        for papel, inicios in ARQ_AREA_TEC.items():
            if papel not in achados and nome.startswith(inicios):
                achados[papel] = _abrir_enc(p)
                break
    return achados


def _linhas_db2(p: Path):
    """Export do DB2 (MODIFIED BY STRIPLZEROS DECPLUSBLANK): CSV sem cabeçalho,
    separado por vírgula, strings entre aspas."""
    texto = _texto_arquivo(p) if not p.name.lower().endswith(".gz") else gzip.open(p, "rt", encoding="latin-1", errors="ignore").read()
    return [r for r in csv.reader(texto.splitlines()) if r]


def _dbnum(s):
    """' 665334.' -> 665334; '' -> None (campo DECIMAL do export DB2)."""
    s = (s or "").strip().rstrip(".")
    if not s:
        return None
    try:
        return int(s)
    except ValueError:
        try:
            return float(s)
        except ValueError:
            return None


def _data_db2(s):
    """20260605 -> '2026-06-05'; 00010101 (data nula do DB2) -> None."""
    s = (s or "").strip()
    if not s or s == "00010101" or len(s) != 8:
        return None
    return f"{s[0:4]}-{s[4:6]}-{s[6:8]}"


def _hora_db2(s):
    """'0001-01-01-17.00.00.000000' -> '17:00'; sem hora de verdade -> None."""
    s = (s or "").strip()
    m = re.search(r"-(\d{2})\.(\d{2})\.\d{2}\.\d+$", s)
    if not m or s.startswith("0001-01-01-00.00.00"):
        return None
    return f"{m.group(1)}:{m.group(2)}"


def _data_hora_db2(s):
    """'2026-06-11-09.39.41.000000' -> '2026-06-11T09:39:41'; data nula -> None."""
    s = (s or "").strip()
    if not s or s.startswith("0001-01-01"):
        return None
    m = re.match(r"(\d{4}-\d{2}-\d{2})-(\d{2})\.(\d{2})\.(\d{2})", s)
    return f"{m.group(1)}T{m.group(2)}:{m.group(3)}:{m.group(4)}" if m else None


def carregar_area_tecnica():
    arqs = achar_arquivos_area_tecnica()
    faltam = [NOME_PAPEL_AT[k] for k in ARQ_AREA_TEC if k not in arqs]
    if faltam:
        if arqs:
            print("AVISO: área técnica ignorada; faltam em data/area_tecnica/: " + "; ".join(faltam))
        return None

    # quantos atributos obrigatórios cada tipo de exame tem (molde, não muda por requisição)
    obrig = {}
    for r in _linhas_db2(arqs["obrigatorios"]):
        if len(r) >= 2 and r[0].strip():
            obrig[r[0].strip()] = _dbnum(r[1]) or 0

    # quantos atributos já foram digitados, por exame de verdade (fil+req+seq+cod)
    digit = {}
    for r in _linhas_db2(arqs["digitados"]):
        if len(r) >= 5:
            chave = (_dbnum(r[0]), _dbnum(r[1]), _dbnum(r[2]), r[3].strip())
            digit[chave] = _dbnum(r[4]) or 0

    pendentes, aguardando = [], []
    cont_status = defaultdict(int)
    cont_setor = defaultdict(int)
    atrasados = 0
    # datetime.now() sozinho pega a hora do servidor do GitHub Actions, que roda em UTC
    # (3h à frente do horário de Brasília) - sem isso, exame com prazo às 17:30 aparecia
    # "atrasado" já no início da tarde. RQEXDTPROMESSA/RQEXHRPROMESSA vêm do DB2 como
    # horário local (sem fuso), então precisamos do agora também como horário local, sem
    # tzinfo (senão o Python não deixa subtrair aware de naive).
    try:
        from zoneinfo import ZoneInfo
        agora = datetime.now(ZoneInfo("America/Sao_Paulo")).replace(tzinfo=None)
    except Exception:
        agora = datetime.now()
    # Alguns exames (ex.: GPO/curva glicêmica, LAC/intolerância a lactose - testes com
    # vários "tempos"/materiais de coleta) vêm com MAIS de uma linha física no export
    # pra um mesmo exame (fil+req+seq+cod) - provavelmente um join que multiplica
    # (LABEXAME ou LABCOLTRI com mais de um registro para o mesmo exame). Sem agrupar
    # aqui, o mesmo exame aparecia 4-5x repetido no painel. Juntamos tudo numa linha só
    # por (fil,req,seq,cod), usando "o mais avançado" entre as cópias (se qualquer cópia
    # diz coletado/triado, vale coletado/triado).
    agrupado = {}
    for r in _linhas_db2(arqs["exames"]):
        if len(r) < 18:
            continue
        (fil, req, seq, cod, dt_entrada, dt_promessa, hr_promessa, conferido, liberado, cancelado,
         paciente, exame, strcod, setor, coletado, triado, hr_coleta, hr_triado) = r[:18]
        fil, req, seq = _dbnum(fil), _dbnum(req), _dbnum(seq)
        cod = cod.strip()
        chave = (fil, req, seq, cod)
        atual = agrupado.get(chave)
        if atual is None:
            agrupado[chave] = {
                "dt_entrada": dt_entrada, "dt_promessa": dt_promessa, "hr_promessa": hr_promessa,
                "conferido": conferido, "liberado": liberado, "cancelado": cancelado,
                "paciente": paciente, "exame": exame, "setor": setor,
                "coletado": coletado, "triado": triado, "hr_coleta": hr_coleta, "hr_triado": hr_triado,
            }
        else:
            s_s = lambda v: (v or "").strip().upper() == "S"
            if s_s(conferido):
                atual["conferido"] = conferido
            if s_s(coletado):
                atual["coletado"], atual["hr_coleta"] = coletado, (atual["hr_coleta"] or hr_coleta)
            if s_s(triado):
                atual["triado"], atual["hr_triado"] = triado, (atual["hr_triado"] or hr_triado)
            if s_s(liberado):
                atual["liberado"] = liberado
            if s_s(cancelado):
                atual["cancelado"] = cancelado

    for (fil, req, seq, cod), v in agrupado.items():
        dt_entrada, dt_promessa, hr_promessa = v["dt_entrada"], v["dt_promessa"], v["hr_promessa"]
        conferido, liberado, cancelado = v["conferido"], v["liberado"], v["cancelado"]
        paciente, exame, setor = v["paciente"], v["exame"], v["setor"]
        coletado, triado, hr_coleta, hr_triado = v["coletado"], v["triado"], v["hr_coleta"], v["hr_triado"]
        if (liberado or "").strip().upper() == "S" or (cancelado or "").strip().upper() == "S":
            continue  # não deveria vir na extração, mas por segurança não entra no painel
        ob, dg = obrig.get(cod, 0), digit.get((fil, req, seq, cod), 0)
        conferido_s = (conferido or "").strip().upper() == "S"
        triado_s = (triado or "").strip().upper() == "S"
        coletado_s = (coletado or "").strip().upper() == "S"
        if conferido_s:
            status = "c"
        elif ob > 0 and dg >= ob:
            status = "d"
        elif dg > 0:
            status = "dp"
        elif triado_s:
            status = "t"
        elif coletado_s:
            status = "col"
        else:
            status = "nc"
        cont_status[status] += 1
        setor_nome = (setor or "").strip() or "Sem setor"
        cont_setor[setor_nome] += 1
        linha = {
            "fil": fil, "req": req, "seq": seq, "cod": cod, "exame": (exame or "").strip(),
            "paciente": (paciente or "").strip(), "setor": setor_nome, "status": status,
            "entrada": _data_db2(dt_entrada), "coletadoEm": _data_hora_db2(hr_coleta),
            "triadoEm": _data_hora_db2(hr_triado),
        }
        if status == "nc":
            aguardando.append(linha)
            continue
        prazo_data = _data_db2(dt_promessa)
        prazo_hora = _hora_db2(hr_promessa)
        linha["prazo"] = prazo_data
        linha["prazoHora"] = prazo_hora
        if prazo_data:
            prazo_dt = datetime.fromisoformat(prazo_data + "T" + (prazo_hora or "23:59") + ":00")
            atrasado = agora > prazo_dt
            linha["atrasado"] = atrasado
            linha["horasRestantes"] = round((prazo_dt - agora).total_seconds() / 3600, 1)
            if atrasado:
                atrasados += 1
        else:
            linha["atrasado"] = None
            linha["horasRestantes"] = None
        if ob:
            linha["digitados"], linha["obrigatorios"] = dg, ob
        pendentes.append(linha)

    pendentes.sort(key=lambda r: (r["horasRestantes"] is None, r["horasRestantes"] if r["horasRestantes"] is not None else 0))
    return {
        "rows": pendentes,
        "aguardandoColeta": aguardando,
        "arquivos": [{"papel": NOME_PAPEL_AT[k], "nome": arqs[k].name} for k in ARQ_AREA_TEC if k in arqs],
        "resumo": {
            "total": len(pendentes) + len(aguardando),
            "emAndamento": len(pendentes), "aguardandoColeta": len(aguardando), "atrasados": atrasados,
            "porStatus": {k: cont_status.get(k, 0) for k in STATUS_AT},
            "porSetor": dict(sorted(cont_setor.items(), key=lambda kv: -kv[1])),
        },
    }


# --------------------------------------------------------------------------- reajustes de preço do DB
# Estes 2 arquivos são gerados automaticamente pelo próprio script (não são
# planilhas que você envia), e como o repositório é público, ficam
# criptografados com a mesma senha do painel (DASHBOARD_SENHA) sempre que ela
# estiver definida — nome termina em .json.enc nesse caso. Sem senha (uso
# local/teste), ficam em .json normal, sem criptografia.
CACHE_PRECOS_DB = PASTA_APOIO / ("_cache_precos_db.json.enc" if SENHA else "_cache_precos_db.json")
HIST_REAJUSTES_DB = PASTA_APOIO / ("_historico_reajustes_db.json.enc" if SENHA else "_historico_reajustes_db.json")
MAX_EVENTOS_REAJUSTE_DB = 30


def _carregar_json(p: Path, padrao):
    """Lê um .json (ou .json.enc, se SENHA estiver definida) gerado por nós mesmos."""
    if not p.exists():
        return padrao
    try:
        dados = p.read_bytes()
        if p.name.endswith(".enc"):
            if not SENHA:
                return padrao
            from cripto_arquivos import abrir_arquivo
            dados = abrir_arquivo(dados, SENHA)
            if dados is None:
                return padrao
        return json.loads(dados.decode("utf-8"))
    except Exception:
        return padrao


def _salvar_json(p: Path, obj):
    """Grava um .json (ou .json.enc, se SENHA estiver definida)."""
    dados = json.dumps(obj, ensure_ascii=False).encode("utf-8")
    if p.name.endswith(".enc"):
        from cripto_arquivos import criptografar_arquivo
        dados = criptografar_arquivo(dados, SENHA)
    p.write_bytes(dados)


def carregar_reajustes_db():
    """Compara a tabela_db de hoje com a última tabela_db processada (guardada em
    data/apoio/_cache_precos_db.json) e, se algo mudou, registra um "evento" de
    reajuste em data/apoio/_historico_reajustes_db.json (código, nome, preço
    antigo, preço novo, % de variação), além de códigos removidos/adicionados.
    Sempre atualiza o cache para a tabela do dia, pronto para a próxima comparação.
    Devolve o histórico completo de eventos (o mais recente por último), para o
    painel mostrar "Reajuste de Preços DB"."""
    arqs = achar_arquivos_apoio()
    if "tab_db" not in arqs:
        return _carregar_json(HIST_REAJUSTES_DB, [])
    tab_db = _ler_precos(arqs["tab_db"])
    atual = {cod: {"nome": t["nome"], "preco": min(t["precos"])} for cod, t in tab_db.items() if t["precos"]}

    cache_anterior = _carregar_json(CACHE_PRECOS_DB, None)
    if cache_anterior is None and SENHA:
        # bootstrap: aceita tambem a "foto" semente ainda em .json puro (sem
        # criptografia), pra nao obrigar a criptografar so pra subir 1 vez;
        # da proxima rodada em diante ja salva criptografado (CACHE_PRECOS_DB)
        cache_anterior = _carregar_json(PASTA_APOIO / "_cache_precos_db.json", None)
    historico = _carregar_json(HIST_REAJUSTES_DB, [])

    if cache_anterior is not None:
        comuns = set(cache_anterior) & set(atual)
        mudou = [c for c in comuns if round(cache_anterior[c]["preco"], 2) != round(atual[c]["preco"], 2)]
        removidos = sorted(set(cache_anterior) - set(atual))
        adicionados = sorted(set(atual) - set(cache_anterior))
        if mudou or removidos or adicionados:
            linhas = []
            for c in mudou:
                pa, pn = cache_anterior[c]["preco"], atual[c]["preco"]
                linhas.append({"c": c, "n": atual[c]["nome"], "pa": round(pa, 2), "pn": round(pn, 2),
                               "d": round(pn - pa, 2), "p": round((pn - pa) / pa * 100, 2) if pa else None})
            linhas.sort(key=lambda r: -(r["p"] or 0))
            evento = {
                "data": datetime.now().strftime("%Y-%m-%d"),
                "arquivo": arqs["tab_db"].name,
                "linhas": linhas,
                "removidos": [{"c": c, "n": cache_anterior[c]["nome"], "pa": round(cache_anterior[c]["preco"], 2)} for c in removidos],
                "adicionados": [{"c": c, "n": atual[c]["nome"], "pn": round(atual[c]["preco"], 2)} for c in adicionados],
            }
            historico.append(evento)
            historico = historico[-MAX_EVENTOS_REAJUSTE_DB:]
            _salvar_json(HIST_REAJUSTES_DB, historico)
            print(f"Reajuste de preços DB detectado: {len(mudou)} exames mudaram, "
                  f"{len(removidos)} saíram, {len(adicionados)} entraram (evento de {evento['data']}).")

    _salvar_json(CACHE_PRECOS_DB, atual)
    return historico


# --------------------------------------------------------------------------- financeiro (contas a pagar / DRE)
PASTA_FINANCEIRO = PASTA_DADOS / "financeiro"
ARQ_FINANCEIRO = {
    "aberto": ("aberto", "contas_a_pagar_aberto", "contas_pagar_aberto"),
    "categoria": ("categoria", "contas_a_pagar_categoria", "contas_pagar_categoria"),
    "aberto_receber": ("contas_a_receber_aberto", "contas_receber_aberto"),
    "categoria_receber": ("contas_a_receber_categoria", "contas_receber_categoria"),
    "baixas": ("contas_a_pagar_baixas", "contas_pagar_baixas"),
    "baixas_receber": ("contas_a_receber_baixas", "contas_receber_baixas"),
    "nao_integrados": ("contas_a_receber_nao_integrados", "contas_receber_nao_integrados"),
    # valor total de cada titulo (soma de todas as parcelas) - usado só pelo DRE, pra
    # ratear a contribuição de um título entre os meses em que ele foi pago (regime de
    # caixa). Sem esses dois arquivos o DRE fica None (mesmo padrão dos outros opcionais).
    "valor_titulo": ("contas_a_pagar_valor_titulo", "contas_pagar_valor_titulo"),
    "valor_titulo_receber": ("contas_a_receber_valor_titulo", "contas_receber_valor_titulo"),
    # lançamentos de caixa (módulo CXB) sem título nenhum por trás - crédito rotativo,
    # aluguel recebido direto no portador, toxicológico/DNA, etc. (ver _parse_movcxb_dre).
    # Também opcional - sem esse arquivo o DRE só fica sem esses lançamentos, como já
    # acontecia antes dessa fonte existir.
    "movcxb": ("contas_movcxb", "movcxb"),
}
# contas a pagar (aberto/categoria) sao obrigatorios; contas a receber, as baixas
# (pagamentos ja efetivados, usados na sub-aba "Contas a Pagar") e o valor_titulo (usado
# só pelo DRE) sao opcionais - enquanto nao forem enviados, os cards/abas que dependem
# deles ficam "sem dados ainda"
FINANCEIRO_OBRIGATORIOS = {"aberto", "categoria"}

# mapa_dre.json mora em data/apoio (mesma pasta usada pra outros arquivos de apoio, ver
# PASTA_APOIO/ARQ_APOIO acima) - não é sensível (só o de-para código->conta do DRE), por
# isso não precisa ir criptografado feito o resto de data/financeiro.
ARQ_MAPA_DRE = PASTA_APOIO / "mapa_dre.json"


def carregar_mapa_dre():
    """De-para TPDRCOD (código do CONCENT) -> conta do DRE, de data/apoio/mapa_dre.json.
    Código novo que não estiver nesse mapa cai em "(sem classificação)" no DRE - nunca
    adivinhar a conta certa, sempre perguntar pro Leo antes de adicionar uma entrada
    nova nesse arquivo."""
    if not ARQ_MAPA_DRE.exists():
        print("AVISO: data/apoio/mapa_dre.json não encontrado - DRE ficará sem classificação de contas.")
        return {}
    try:
        with open(ARQ_MAPA_DRE, encoding="utf-8") as f:
            conteudo = json.load(f)
        return conteudo.get("mapa", {})
    except Exception as e:
        print(f"AVISO: não consegui ler data/apoio/mapa_dre.json ({e}) - DRE ficará sem classificação de contas.")
        return {}
MESES_JANELA_RECORRENCIA = 12
MIN_MESES_RECORRENTE = 10  # aparece em pelo menos 10 dos ultimos 12 meses = considerado recorrente


def achar_arquivos_financeiro():
    achados = {}
    if not PASTA_FINANCEIRO.is_dir():
        return achados
    for p in sorted(PASTA_FINANCEIRO.iterdir()):
        if not p.is_file() or p.name.startswith(".") or not p.name.lower().endswith((".xlsx", ".csv", ".gz", ".enc")):
            continue
        nome = _sa(p.name).replace(" ", "_")
        for papel, inicios in ARQ_FINANCEIRO.items():
            if papel not in achados and nome.startswith(inicios):
                achados[papel] = _abrir_enc(p)
                break
    return achados


def _data_aaaammdd(s):
    """'20260930' -> date(2026,9,30); tolera lixo/vazio -> None."""
    s = str(s).strip()
    if not s or len(s) != 8 or not s.isdigit():
        return None
    ano, mes, dia = int(s[:4]), int(s[4:6]), int(s[6:8])
    try:
        d = date(ano, mes, dia)
    except ValueError:
        return None
    if ano < 2015 or ano > 2035:
        return None
    return d


def _fim_mes(d):
    if d.month == 12:
        return date(d.year, 12, 31)
    return date(d.year, d.month + 1, 1) - timedelta(days=1)


def _add_meses(ano, mes, k):
    m = mes - 1 + k
    return ano + m // 12, m % 12 + 1


def carregar_financeiro(movimento_diario=None):
    """Contas a pagar (aberto.csv) + lançamentos por categoria (categoria.csv,
    também usado no DRE) exportados do CONCENT. Monta o resumo da aba
    Financeiro / Visão Geral: vencidas, vencendo hoje, resto do mês, total a
    pagar até fim do mês, e a lista de fornecedores/categorias recorrentes
    (aparecem em pelo menos 10 dos últimos 12 meses) que ainda não tiveram
    nenhum lançamento no mês atual. Devolve None se os arquivos não foram
    enviados para data/financeiro/.

    movimento_diario: {filial: {ym: valor}} com o Líquido faturado (vindo da
    base de exames, não dos títulos) - usado só no DRE por posto individual,
    pra substituir a receita (ver _montar_dre)."""
    arqs = achar_arquivos_financeiro()
    faltam = [k for k in FINANCEIRO_OBRIGATORIOS if k not in arqs]
    if faltam:
        if arqs:
            print("AVISO: aba Financeiro ignorada; faltam em data/financeiro/: " + ", ".join(faltam))
        return None

    hoje = datetime.now().date()
    fim_mes_atual = _fim_mes(hoje)

    aberto = []
    for row in _linhas_arquivo(arqs["aberto"]):
        row = [_cel(c) for c in row]
        if len(row) < 10:
            continue
        _modulo, filcod, titnr, _serie, fornecedor, parcela, dtvencto, _vlr, saldo, _hist = row[:10]
        # o DB2 exporta FILCOD (coluna decimal) às vezes com um ponto sobrando no final
        # quando não tem casas decimais (ex.: "100." em vez de "100") - sem isso o
        # cruzamento com data/filiais não batia e mostrava só o número (com o ponto).
        filcod = filcod.strip().rstrip(".")
        d = _data_aaaammdd(dtvencto)
        aberto.append({"titulo": titnr, "fornecedor": _nome_curto(fornecedor), "parcela": parcela,
                        "vencimento": d, "saldo": _preco(saldo) or 0.0, "filial": filcod})

    vencidas = [a for a in aberto if a["vencimento"] and a["vencimento"] < hoje]
    venc_hoje = [a for a in aberto if a["vencimento"] and a["vencimento"] == hoje]
    resto_mes = [a for a in aberto if a["vencimento"] and hoje < a["vencimento"] <= fim_mes_atual]
    ate_fim_mes = [a for a in aberto if a["vencimento"] and a["vencimento"] <= fim_mes_atual]

    def resumo_grupo(lst):
        return {"n": len(lst), "v": round(sum(x["saldo"] for x in lst), 2)}

    tabela = []
    for a in sorted(vencidas + venc_hoje, key=lambda x: x["vencimento"]):
        dias = (hoje - a["vencimento"]).days
        tabela.append({"f": a["fornecedor"], "t": a["titulo"], "p": a["parcela"], "fil": a["filial"],
                        "v": a["vencimento"].strftime("%Y-%m-%d"), "d": dias, "s": round(a["saldo"], 2),
                        "st": "vencido" if dias > 0 else "hoje"})

    categoria = []
    for row in _linhas_arquivo(arqs["categoria"]):
        row = [_cel(c) for c in row]
        if len(row) < 10:
            continue
        _filcod, _titnr, _serie, fornecedor, dtemissao, _centro, _tpcod, tipo, _tptipo, _valor = row[:10]
        fornecedor = _nome_curto(fornecedor)
        d = _data_aaaammdd(dtemissao)
        if d:
            categoria.append({"fornecedor": fornecedor, "tipo": tipo, "data": d})

    mes_atual_str = hoje.strftime("%Y-%m")
    meses_janela = set()
    for k in range(1, MESES_JANELA_RECORRENCIA + 1):
        ay, am = _add_meses(hoje.year, hoje.month, -k)
        meses_janela.add(f"{ay:04d}-{am:02d}")

    por_par_meses = defaultdict(set)
    par_no_mes_atual = set()
    for c in categoria:
        chave = (c["fornecedor"], c["tipo"])
        m = c["data"].strftime("%Y-%m")
        if m in meses_janela:
            por_par_meses[chave].add(m)
        if m == mes_atual_str:
            par_no_mes_atual.add(chave)

    recorrentes = {par for par, meses in por_par_meses.items() if len(meses) >= MIN_MESES_RECORRENTE}
    nao_lancadas = sorted(recorrentes - par_no_mes_atual)

    # dia do mês "normal" de cada recorrência ainda não lançada (usa a mesma janela de
    # histórico da categoria - a data de emissão é a melhor aproximação que temos hoje
    # do "costuma vencer perto do dia X", já que ainda não temos o vencimento histórico
    # de títulos já baixados/removidos, só do que está em aberto agora).
    dias_por_par = defaultdict(list)
    for c in categoria:
        m = c["data"].strftime("%Y-%m")
        if m in meses_janela:
            dias_por_par[(c["fornecedor"], c["tipo"])].append(c["data"].day)

    def _dia_estimado(chave):
        dias = dias_por_par.get(chave)
        return round(statistics.median(dias)) if dias else None

    # baixas (pagamentos já efetivados, MOVTITULO com MVTITMOVTO='BXA') - opcional;
    # alimenta "pago até hoje" e a projeção pela média dos últimos 2 anos. Enquanto
    # contas_pagar_baixas.csv não for enviado, esses cards ficam "sem dados ainda",
    # do mesmo jeito que contas a receber já fazia.
    # fornecedor por título (filial+numero+série), a partir do histórico de categoria -
    # que continua tendo o lançamento mesmo depois do título ser baixado/pago, ao
    # contrário do aberto.csv (que só tem o que ainda está em aberto agora). É o que
    # permite mostrar o nome do fornecedor numa baixa antiga.
    fornecedor_por_titulo = {}
    for row in _linhas_arquivo(arqs["categoria"]):
        row = [_cel(c) for c in row]
        if len(row) < 10:
            continue
        filcod_c, titnr_c, serie_c, fornecedor_c = row[0], row[1], row[2], row[3]
        fornecedor_por_titulo[(filcod_c.strip().rstrip("."), titnr_c, serie_c)] = _nome_curto(fornecedor_c)

    baixas = []
    if "baixas" in arqs:
        for row in _linhas_arquivo(arqs["baixas"]):
            row = [_cel(c) for c in row]
            if len(row) < 6:
                continue
            filcod_b, titnr_b, serie_b, parcela_b, dtmov, valor = row[:6]
            filcod_b = filcod_b.strip().rstrip(".")
            d = _data_aaaammdd(dtmov)
            if d:
                baixas.append({
                    "data": d, "valor": _preco(valor) or 0.0, "titulo": titnr_b, "parcela": parcela_b,
                    "filial": filcod_b,
                    "fornecedor": fornecedor_por_titulo.get((filcod_b, titnr_b, serie_b), ""),
                })

    pago_mes = None
    projecao_mes = None
    if baixas:
        pago_mes = round(sum(b["valor"] for b in baixas
                              if b["data"].strftime("%Y-%m") == mes_atual_str and b["data"] <= hoje), 2)
        totais_por_ano = defaultdict(float)
        for b in baixas:
            if b["data"].month == hoje.month and b["data"].year < hoje.year:
                totais_por_ano[b["data"].year] += b["valor"]
        anos_recentes = sorted(totais_por_ano)[-2:]
        if anos_recentes:
            projecao_mes = round(sum(totais_por_ano[a] for a in anos_recentes) / len(anos_recentes), 2)

    janela_7d = hoje + timedelta(days=7)
    janela_15d = hoje + timedelta(days=15)
    a_pagar_7d = [a for a in aberto if a["vencimento"] and a["vencimento"] <= janela_7d]
    a_pagar_15d = [a for a in aberto if a["vencimento"] and a["vencimento"] <= janela_15d]
    # títulos de contas a receber vencidos há mais de 15 dias não entram mais nos KPIs de
    # "a receber" (fim do mês / 7 dias / 15 dias) - atraso grande é outro tipo de problema,
    # não "a receber" no sentido normal do card
    janela_receber_desde = hoje - timedelta(days=15)

    recebimentos_mes = None
    aberto_r = []
    if "aberto_receber" in arqs:
        for row in _linhas_arquivo(arqs["aberto_receber"]):
            row = [_cel(c) for c in row]
            if len(row) < 10:
                continue
            _modulo, filcod_r, titnr_r, serie_r, cliente, parcela_r, dtvencto, _vlr, saldo, _hist = row[:10]
            tipo_r = row[10].strip() if len(row) > 10 else ""
            filcod_r = filcod_r.strip().rstrip(".")
            d = _data_aaaammdd(dtvencto)
            aberto_r.append({"titulo": titnr_r, "cliente": _nome_curto(cliente), "parcela": parcela_r,
                              "vencimento": d, "saldo": _preco(saldo) or 0.0, "filial": filcod_r,
                              "serie": serie_r.strip(), "tipo": "C" if tipo_r == "C" else "P"})
        a_receber_fim_mes = [a["saldo"] for a in aberto_r if a["vencimento"] and janela_receber_desde <= a["vencimento"] <= fim_mes_atual]
        recebimentos_mes = {"n": len(a_receber_fim_mes), "v": round(sum(a_receber_fim_mes), 2)}

    # baixas de contas a receber (mesma tabela MOVTITULO, módulo CRE) - pra "recebido até
    # hoje" e a projeção, espelhando exatamente a lógica que já existe pro contas a pagar
    # acima. O nome do cliente por título vem do categoria_receber (segue existindo depois
    # do título ser baixado, ao contrário do aberto_receber).
    cliente_por_titulo = {}
    if "categoria_receber" in arqs:
        for row in _linhas_arquivo(arqs["categoria_receber"]):
            row = [_cel(c) for c in row]
            if len(row) < 10:
                continue
            filcod_c, titnr_c, serie_c, cliente_c = row[0], row[1], row[2], row[3]
            cliente_por_titulo[(filcod_c.strip().rstrip("."), titnr_c, serie_c)] = _nome_curto(cliente_c)

    baixas_r = []
    if "baixas_receber" in arqs:
        for row in _linhas_arquivo(arqs["baixas_receber"]):
            row = [_cel(c) for c in row]
            if len(row) < 6:
                continue
            filcod_b, titnr_b, serie_b, parcela_b, dtmov, valor = row[:6]
            tipo_b = row[6].strip() if len(row) > 6 else ""
            filcod_b = filcod_b.strip().rstrip(".")
            d = _data_aaaammdd(dtmov)
            if d:
                baixas_r.append({
                    "data": d, "valor": _preco(valor) or 0.0, "titulo": titnr_b, "parcela": parcela_b,
                    "filial": filcod_b, "tipo": "C" if tipo_b == "C" else "P",
                    "cliente": cliente_por_titulo.get((filcod_b, titnr_b, serie_b), ""),
                })

    recebido_mes = None
    projecao_mes_r = None
    if baixas_r:
        recebido_mes = round(sum(b["valor"] for b in baixas_r
                                  if b["data"].strftime("%Y-%m") == mes_atual_str and b["data"] <= hoje), 2)
        totais_por_ano_r = defaultdict(float)
        for b in baixas_r:
            if b["data"].month == hoje.month and b["data"].year < hoje.year:
                totais_por_ano_r[b["data"].year] += b["valor"]
        anos_recentes_r = sorted(totais_por_ano_r)[-2:]
        if anos_recentes_r:
            projecao_mes_r = round(sum(totais_por_ano_r[a] for a in anos_recentes_r) / len(anos_recentes_r), 2)

    def resumo_grupo_r(lst):
        return {"n": len(lst), "v": round(sum(x["saldo"] for x in lst), 2)}

    receber_7d = [a for a in aberto_r if a["vencimento"] and janela_receber_desde <= a["vencimento"] <= janela_7d]
    receber_15d = [a for a in aberto_r if a["vencimento"] and janela_receber_desde <= a["vencimento"] <= janela_15d]

    financeiro_receber = None
    if aberto_r or baixas_r:
        financeiro_receber = {
            "resumo": {
                "aReceberMes": recebimentos_mes or {"n": 0, "v": 0.0},
                "recebidoMes": recebido_mes,
                "projecaoMes": projecao_mes_r,
                "aReceber7d": resumo_grupo_r(receber_7d),
                "aReceber15d": resumo_grupo_r(receber_15d),
            },
            "abertos": [{"f": a["cliente"], "t": a["titulo"], "p": a["parcela"], "fil": a["filial"],
                         "tp": a["tipo"],
                         "v": a["vencimento"].strftime("%Y-%m-%d") if a["vencimento"] else None,
                         "s": round(a["saldo"], 2)} for a in aberto_r],
            "baixados": [{"f": b["cliente"], "t": b["titulo"], "p": b["parcela"], "fil": b["filial"],
                          "tp": b["tipo"],
                          "v": b["data"].strftime("%Y-%m-%d"), "s": round(b["valor"], 2)} for b in baixas_r],
        }

    # títulos "não integrados": exame já faturado (RQEXDTFATURA preenchida no CONCENT),
    # não cancelado, sem NENHUM dos 3 títulos (convênio/particular/cartão) vinculado -
    # dinheiro que já saiu pro convênio mas nunca virou conta a receber rastreável.
    # Validado direto no CONCENT (ver comentário da SQL_NAO_INTEGRADOS no coletar) antes
    # de virar número do painel.
    # Por ora, a SQL_NAO_INTEGRADOS já desconsidera convênio tipo C (à vista/caixa) na
    # origem - ver comentário no coletar_contas_pagar.py. O Leo vai confirmar mais coisas
    # direto na CONCENT antes de decidir um critério mais fino (ex: taxa de integração por
    # convênio); até lá, o que chega aqui já é só o que deve mesmo virar título de convênio.
    def _bloco_nao_integrados(linhas, top_n=12):
        """Monta resumo + por-convênio + matriz mensal por não integrados."""
        por_convenio = defaultdict(lambda: {"n": 0, "v": 0.0})
        for r in linhas:
            g = por_convenio[r["convenio"]]
            g["n"] += 1
            g["v"] += r["valor"]
        porConvenio = sorted(
            [{"convenio": c, "n": g["n"], "v": round(g["v"], 2)} for c, g in por_convenio.items()],
            key=lambda x: -x["v"])

        # evolução mensal por convênio (mesma ideia da planilha dinâmica que o Leo já
        # fazia manualmente: convênio nas linhas, mês do FATURAMENTO nas colunas, soma do
        # valor não integrado em cada célula) - mostra quando esse dinheiro ficou parado,
        # não só o total acumulado. Só os top_n convênios de maior valor viram linha
        # própria; o resto entra agrupado em "Outros convênios" pra não estourar a tabela.
        por_convenio_mes = defaultdict(lambda: defaultdict(float))
        meses_set = set()
        for r in linhas:
            if not r["faturado"]:
                continue
            ym = r["faturado"][:7]
            por_convenio_mes[r["convenio"]][ym] += r["valor"]
            meses_set.add(ym)
        meses = sorted(meses_set)
        top_convs = [c["convenio"] for c in porConvenio[:top_n]]
        outros_convs = [c["convenio"] for c in porConvenio[top_n:]]

        def linha_matriz(nome, lista_convs):
            valores = [round(sum(por_convenio_mes[c].get(m, 0.0) for c in lista_convs), 2) for m in meses]
            return {"convenio": nome, "valores": valores, "total": round(sum(valores), 2)}

        # linhas dos top_n sempre aparecem; as dos "outros" convênios também são calculadas
        # individualmente (não só agregadas) pra permitir expandir a tabela e ver cada uma -
        # por padrão o painel mostra só o resumo agregado ("Outros convênios (28)") e o
        # usuário decide se quer expandir.
        linhas_top = [linha_matriz(c, [c]) for c in top_convs]
        linhas_outros = [linha_matriz(c, [c]) for c in outros_convs]
        outros_resumo = linha_matriz(f"Outros convênios ({len(outros_convs)})", outros_convs) if outros_convs else None
        todas_linhas = linhas_top + linhas_outros
        total_mes = [round(sum(l["valores"][i] for l in todas_linhas), 2) for i in range(len(meses))]

        return {
            "resumo": {"n": len(linhas), "v": round(sum(r["valor"] for r in linhas), 2)},
            "porConvenio": porConvenio,
            "linhas": sorted(linhas, key=lambda r: -(r["dias"] or 0)),
            "evolucaoMensal": {"meses": meses, "linhasTop": linhas_top, "linhasOutros": linhas_outros,
                                "outrosResumo": outros_resumo, "totalMes": total_mes,
                                "totalGeral": round(sum(total_mes), 2)},
        }

    nao_integrados = None
    if "nao_integrados" in arqs:
        linhas_ni = []
        for row in _linhas_arquivo(arqs["nao_integrados"]):
            row = [_cel(c) for c in row]
            if len(row) < 8:
                continue
            filcod_n, reqnum_n, exaseq_n, exacod_n, _convcod_n, convnome_n, dtfat_n, valor_n = row[:8]
            filcod_n = filcod_n.strip().rstrip(".")
            reqnum_n = reqnum_n.strip().rstrip(".")
            exaseq_n = exaseq_n.strip().rstrip(".")
            d = _data_aaaammdd(dtfat_n)
            linhas_ni.append({
                "filial": filcod_n, "requisicao": reqnum_n, "seq": exaseq_n, "exame": exacod_n,
                "convenio": _nome_curto(convnome_n) or "SEM CONVÊNIO CADASTRADO",
                "faturado": d.strftime("%Y-%m-%d") if d else None,
                "dias": (hoje - d).days if d else None,
                "valor": _preco(valor_n) or 0.0,
            })
        nao_integrados = _bloco_nao_integrados(linhas_ni)

    dre = _montar_dre(arqs, movimento_diario)

    return {
        "hoje": hoje.strftime("%Y-%m-%d"),
        "resumo": {
            "vencidas": resumo_grupo(vencidas), "hoje": resumo_grupo(venc_hoje),
            "restoMes": resumo_grupo(resto_mes), "aPagarMes": resumo_grupo(ate_fim_mes),
            "recebimentosMes": recebimentos_mes,
            "pagoMes": pago_mes, "projecaoMes": projecao_mes,
            "aPagar7d": resumo_grupo(a_pagar_7d), "aPagar15d": resumo_grupo(a_pagar_15d),
        },
        "tabela": tabela,
        "recorrencia": {
            "naoLancadas": [{"f": f, "t": t, "dia": _dia_estimado((f, t))} for f, t in nao_lancadas],
            "n": len(nao_lancadas),
        },
        # listas completas (não só vencidas/hoje) pra sub-aba "Contas a Pagar", com
        # filtro por data/status no próprio painel
        "abertos": [{"f": a["fornecedor"], "t": a["titulo"], "p": a["parcela"], "fil": a["filial"],
                     "v": a["vencimento"].strftime("%Y-%m-%d") if a["vencimento"] else None,
                     "s": round(a["saldo"], 2)} for a in aberto],
        "baixados": [{"f": b["fornecedor"], "t": b["titulo"], "p": b["parcela"], "fil": b["filial"],
                      "v": b["data"].strftime("%Y-%m-%d"), "s": round(b["valor"], 2)} for b in baixas],
        # aba Contas a Receber (espelha a de Contas a Pagar acima) - None enquanto os
        # arquivos de aberto_receber/baixas_receber não forem enviados
        "receber": financeiro_receber,
        # seção de destaque "títulos não integrados" - None enquanto
        # contas_receber_nao_integrados.csv não for enviado
        "naoIntegrados": nao_integrados,
        # DRE (Demonstrativo de Resultado) em regime de caixa, por posto + consolidado
        # (soma dos postos 01-13, sem o 100) + geral (soma de tudo, incluindo o 100) -
        # None enquanto contas_a_pagar_valor_titulo/contas_a_receber_valor_titulo não
        # forem enviados (ver _montar_dre logo abaixo pra metodologia completa)
        "dre": dre,
    }


# ------------------------------------------------------------------- DRE (regime de caixa)
def _chave_dre(filcod, titnr, serie, emitcod, espdccod):
    """Chave que identifica um título de forma única no CONCENT. FILCOD+TITNRDOCTO+SERIECOD
    sozinhos NÃO bastam: títulos manuais (SERIECOD='M') às vezes usam um número de
    referência parecido com uma data (ex.: "16072026") que fornecedores diferentes no
    mesmo posto podem repetir - sem EMITCOD/ESPDCCOD, títulos de fornecedores distintos
    acabam se misturando e o rateio por mês sai completamente errado. Bug encontrado e
    corrigido em 30/09/2026 validando o DRE contra os relatórios reais de 2023 do
    Claudio (ex-consultor) - ver diagnostico_dre_mes.py."""
    return (filcod.strip().rstrip("."), titnr.strip(), serie.strip(),
            emitcod.strip().rstrip("."), espdccod.strip().rstrip("."))


def _carregar_valor_titulo_dre(arqs, papel):
    """Valor total de cada título (soma de todas as parcelas), de contas_a_pagar/
    contas_a_receber_valor_titulo.csv. É o denominador da fração de rateio por mês."""
    valores = {}
    if papel not in arqs:
        return valores
    for row in _linhas_arquivo(arqs[papel]):
        row = [_cel(c) for c in row]
        if len(row) < 6:
            continue
        filcod, titnr, serie, emit, esp, total = row[:6]
        valores[_chave_dre(filcod, titnr, serie, emit, esp)] = _preco(total) or 0.0
    return valores


def _parse_rateio_dre(arqs, papel):
    """Uma linha por (título, código DRE) rateado, a partir de contas_a_pagar/
    contas_a_receber_categoria.csv (já traz o rateio via join com RATEIOTITULO na
    coleta da VM). Layout: FILCOD,TITNRDOCTO,SERIECOD,EMITNOME,dtemissao,centro,
    TPDRCOD,TPDRDESCRICAO,TPDRTIPO,RATITVLR,historico,EMITCOD,ESPDCCOD."""
    linhas = []
    if papel not in arqs:
        return linhas
    for row in _linhas_arquivo(arqs[papel]):
        row = [_cel(c) for c in row]
        if len(row) < 13:
            continue
        filcod, titnr, serie, fornecedor, _dtemissao, _centro, tpdrcod, tpdrdesc, tpdrtipo, valor, historico, emit, esp = row[:13]
        linhas.append({
            "chave": _chave_dre(filcod, titnr, serie, emit, esp),
            "filial": filcod.strip().rstrip("."),
            "titulo": titnr.strip(),
            "fornecedor": _nome_curto(fornecedor),
            "tpdrcod": tpdrcod.strip().rstrip("."),
            "descricao": tpdrdesc.strip(),
            "tipo": tpdrtipo.strip(),
            "valor": _preco(valor) or 0.0,
            "historico": historico.strip(),
        })
    return linhas


def _parse_baixas_dre(arqs, papel):
    """Uma linha por baixa (pagamento/recebimento), a partir de contas_a_pagar/
    contas_a_receber_baixas.csv. Layout: FILCOD,TITNRDOCTO,SERIECOD,PATITPARCELA,
    dtpgto,valor,C/P,EMITCOD,ESPDCCOD (o C/P só é usado em baixas_receber, mas a coluna
    existe nos dois - ver SQL_BAIXAS no coletar_contas_pagar.py)."""
    linhas = []
    if papel not in arqs:
        return linhas
    for row in _linhas_arquivo(arqs[papel]):
        row = [_cel(c) for c in row]
        if len(row) < 9:
            continue
        filcod, titnr, serie, _parcela, dtmov, valor, _flag, emit, esp = row[:9]
        d = _data_aaaammdd(dtmov)
        if not d:
            continue
        linhas.append({"chave": _chave_dre(filcod, titnr, serie, emit, esp),
                        "mes": d.strftime("%Y-%m"), "valor": _preco(valor) or 0.0})
    return linhas


def _agrupar_pago_por_mes(baixas_dre):
    pago = defaultdict(lambda: defaultdict(float))  # chave -> {mes: valor pago naquele mes}
    for b in baixas_dre:
        pago[b["chave"]][b["mes"]] += b["valor"]
    return pago


def _eventos_dre(rateio_linhas, pago_por_mes, valor_titulo, mapa_dre, nao_classificados):
    """Regime de caixa: pra cada rateio (título x código DRE), distribui a contribuição
    entre os meses em que o título teve alguma baixa, proporcional a
    (pago naquele mês / valor total do título) - assim um título pago em parcelas em
    meses diferentes entra em cada mês só com a fatia paga naquele mês. Metodologia
    validada em 30/09/2026 contra os relatórios reais de maio/2023 e novembro/2023 do
    Claudio (ex-consultor) - bateu centavo a centavo depois de corrigida a chave do
    título (ver _chave_dre) e entendido que o posto 100 é o hub administrativo/rateio
    (fica de fora do "Consolidado" mas entra no "Geral")."""
    eventos = []
    for r in rateio_linhas:
        meses_pagos = pago_por_mes.get(r["chave"])
        if not meses_pagos:
            continue  # título sem nenhuma baixa ainda - não entra no regime de caixa
        total_titulo = valor_titulo.get(r["chave"])
        info = mapa_dre.get(r["tpdrcod"])
        if info is None:
            nao_classificados.add((r["tpdrcod"], r["descricao"]))
            conta = "(sem classificação)"
        elif info.get("regra_especial") == "sinal":
            conta = info["conta_dre_positivo"] if r["valor"] >= 0 else info["conta_dre_negativo"]
        elif info.get("conta_dre"):
            conta = info["conta_dre"]
        else:
            # existe em mapa_dre.json mas com "conta_dre": null (placeholder nunca
            # preenchido, ex.: entrada criada de antemão esperando um código que só veio a
            # aparecer de verdade depois) - trata igual a "não encontrado", pra aparecer no
            # aviso de não classificados em vez de desaparecer silenciosamente.
            nao_classificados.add((r["tpdrcod"], info.get("descricao") or r["descricao"]))
            conta = "(sem classificação)"
        for mes, pago in meses_pagos.items():
            if total_titulo:
                fracao = pago / total_titulo
            else:
                # sem valor_titulo (ex.: arquivo ainda não enviado) - assume que a baixa
                # cobre o título inteiro em vez de descartar a linha
                fracao = 1.0
            contribuicao = round(r["valor"] * fracao, 2)
            if not contribuicao:
                continue
            eventos.append({"mes": mes, "filial": r["filial"], "conta": conta, "tipo": r["tipo"],
                             "valor": contribuicao, "titulo": r["titulo"], "fornecedor": r["fornecedor"],
                             "historico": r["historico"]})
    return eventos


def _parse_movcxb_dre(arqs, mapa_dre, nao_classificados):
    """Lançamentos de caixa (módulo CXB) que não passam por título nenhum - ex.: saques/
    pagamentos do crédito rotativo, aluguéis recebidos direto no portador, exames de
    toxicológico/DNA cujo recebimento não gera conta a receber no sistema. Vêm prontos
    da VM em contas_movcxb.csv (ver SQL_MOVCXB em coletar_contas_pagar.py), já filtrados
    só pros lançamentos "reais" (MVCXBTIPO='TS' - as transferências entre portadores,
    MVCXBTIPO='TR', são só troco de bolso entre contas internas e não entram no DRE;
    confirmado que 100% das TR do ano têm TPDRCOD=0, nunca aparecem aqui).

    Diferente de título (que pode ser pago em parcelas em meses diferentes - regime de
    caixa via _eventos_dre), cada linha de MOVCXB já É o movimento de caixa, pronto e
    definitivo: sem rateio por mês, sem fração de baixa - o valor da linha é o evento
    inteiro. Layout: FILCOD,data(AAAAMMDD),MVCXBNRO,MVCXBMOVTO(E/S),valor,TPDRCOD,
    TPDRDESCRICAO,TPDRTIPO,historico,EMITCOD,EMITNOME."""
    eventos = []
    if "movcxb" not in arqs:
        return eventos
    for row in _linhas_arquivo(arqs["movcxb"]):
        row = [_cel(c) for c in row]
        if len(row) < 11:
            continue
        filcod, data, _nro, movto, valor, tpdrcod, tpdrdesc, tpdrtipo, historico, _emit, fornecedor = row[:11]
        d = _data_aaaammdd(data)
        if not d:
            continue
        v = _preco(valor) or 0.0
        if not v:
            continue
        # MVCXBMOVTO: 'E' = entrada (dinheiro entrando no portador), 'S' = saída - dá o
        # sinal direto, sem precisar inferir nada do texto do histórico.
        valor_assinado = v if movto.strip().upper() == "E" else -v
        tpdrcod = tpdrcod.strip().rstrip(".")
        info = mapa_dre.get(tpdrcod)
        if info is None:
            nao_classificados.add((tpdrcod, tpdrdesc.strip()))
            conta = "(sem classificação)"
        elif info.get("regra_especial") == "sinal":
            conta = info["conta_dre_positivo"] if valor_assinado >= 0 else info["conta_dre_negativo"]
        elif info.get("conta_dre"):
            conta = info["conta_dre"]
        else:
            # mesma situação do _eventos_dre acima: existe em mapa_dre.json mas com
            # "conta_dre": null - trata como não classificado em vez de sumir.
            nao_classificados.add((tpdrcod, info.get("descricao") or tpdrdesc.strip()))
            conta = "(sem classificação)"
        eventos.append({"mes": d.strftime("%Y-%m"), "filial": filcod.strip().rstrip("."), "conta": conta,
                         "tipo": tpdrtipo.strip(), "valor": round(valor_assinado, 2), "titulo": "",
                         "fornecedor": _nome_curto(fornecedor), "historico": historico.strip()})
    return eventos


def _resumo_dre(eventos):
    """Agrupa eventos em {mes: {conta_dre: {tipo, valor, lancamentos: [...]}}}, pronto
    pro painel abrir uma conta do DRE e ver os lançamentos que a compõem naquele mês."""
    out = defaultdict(lambda: defaultdict(lambda: {"tipo": None, "valor": 0.0, "lancamentos": []}))
    for e in eventos:
        bucket = out[e["mes"]][e["conta"]]
        bucket["tipo"] = e["tipo"]
        bucket["valor"] = round(bucket["valor"] + e["valor"], 2)
        bucket["lancamentos"].append({"titulo": e["titulo"], "fornecedor": e["fornecedor"],
                                       "valor": e["valor"], "historico": e["historico"],
                                       "filial": e["filial"]})
    for contas in out.values():
        for c in contas.values():
            c["lancamentos"].sort(key=lambda l: -abs(l["valor"]))
    return {mes: dict(contas) for mes, contas in out.items()}


# --------------------------------------------------------------------------- DRE: hierarquia
# Estrutura do DRE igual à planilha "Acompanhamento Realizado" que o Claudio (ex-consultor)
# usava (aba "Celape Geral"), reproduzida linha por linha a partir do arquivo real de
# novembro/2023 que o Leo enviou em 30/09/2026. Só entram aqui os nomes de conta que já
# existem em data/apoio/mapa_dre.json (99 dos 101 bateram exatamente com o texto da
# planilha do Claudio) - conferido programaticamente antes de escrever esta lista, não
# chutado. Uma conta nova que apareça em mapa_dre.json sem estar em nenhum desses grupos
# cai automaticamente no grupo "(-) Outras contas (a posicionar na hierarquia)" no fim,
# com aviso no log - nunca inventamos onde ela entra.
_DRE_GRUPO_RECEITA_BRUTA = ["(+) Receita unidades", "(+) Receita Fat Convênio", "(-) Cancelamento de serviços ou glosas"]
_DRE_GRUPO_IMPOSTOS_RECEITA = ["(-) ISSQN", "(-) PIS", "(-) COFINS", "(-) Simples Nacional", "(-) Imposto autônomo"]
_DRE_GRUPO_CUSTOS_VARIAVEIS = ["(-) Compras de Kit Reagente", "(-) Aluguel Interface", "(-) Insumos", "(-) Terceirização exames - Lab Apoio"]
_DRE_SUBGRUPOS_DESPESAS_FIXAS = [
    ("Desp. com pessoal", ["(-) Salarios e ordenados", "(-) Horas extras", "(-) Adicional noturno",
        "(-) Adicional de insalubridade", "(-) Autonomos", "(-) Estagiarios", "(-) 13º Salário", "(-) Férias",
        "(-) INSS s/ folha", "(-) FGTS", "(-) Bônus, prêmios e gratificações", "(-) Aviso prévio e indenizações",
        "(-) Seguro de vida", "(-) Vale alimentação", "(-) Vale transporte", "(-) Assistencia médica",
        "(-) Uniformes e EPI´s", "(-) Contribuição sindical", "(-) Contribuição patronal",
        "(-) Cursos e treinamentos", "(-) Comissões", "(-) Auxilio educação", "(-) Medicina ocupacional",
        "(-) Confraternização", "(-) Lanches e refeições", "(-) Outras despesas com pessoal",
        "(-) Pró labore", "(-) Distribuição a sócios"]),
    ("Desp. com infraestrutura", ["(-) Energia elétrica", "(-) Água e esgoto", "(-) Aluguéis e condomínios",
        "(-) Limpeza / Diaristas", "(-) Telefone e internet", "(-) Coleta de resíduos e detetização",
        "(-) IPTU, alvará, taxa de lixo e VISA", "(-) Segurança e vigilância",
        "(-) Manutenção e conservação de estrutura"]),
    ("Desp. com marketing e comercial", ["(-) Marketing e propaganda", "(-) Tráfego Pago", "(-) Eventos",
        "(-) Viagens comerciais", "(-) Doações, brindes e patrocinios"]),
    ("Desp. com equipamentos", ["(-) Locação de  equipamento", "(-) Manutenção e conservação de  equipamento",
        "(-) Manutenção e conservação de equipamento", "(-) Compra de equipamentos"]),
    ("Desp. administrativas", ["(-) Material de consumo/expediente", "(-) Material de higiêne/limpeza",
        "(-) Material de copa/cozinha", "(-) Despesas com viagens / NET", "(-) Cartórios e registros",
        "(-) Combustível", "(-) Entidades de Classe", "(-) Mercado", "(-) Cartões Empresariais",
        "(-) Outras Despesas", "???"]),
    ("Desp. com logística", ["(-) Despesas e conservação de veículos", "(-) Combustíveis e lubrificantes",
        "(-) Correios e transportes", "(-) Fretes e carretos", "(-) IPVA, seguro e multas"]),
    ("Desp. com serviços contratados", ["(-) Assessoria advocatícia", "(-) Assessoria contábil",
        "(-) Consultoria e auditoria", "(-) Softwares e informática", "(-) Controle de qualidade externo"]),
]
_DRE_GRUPO_RECEITAS_FINANCEIRAS = ["(+) Juros recebidos", "(+) Descontos auferidos"]
_DRE_GRUPO_DESPESAS_FINANCEIRAS = ["(-) Juros", "(-) Taxas Getnet", "(-) Descontos concedidos", "(-) Tarifas bancárias"]
_DRE_GRUPO_IMPOSTOS_LUCRO = ["(-) Contribuição Social", "(-) Imposto de Renda"]
_DRE_GRUPO_OUTROS_MOVIMENTOS = ["(+) Adiantamento de Clientes", "(-) Investimentos", "(+) Empréstimos Recebidos",
                                 "(-) Empréstimos Pagos", "(-) Diferença Cartão"]
_DRE_GRUPO_CONTA_1000 = ["(-) Cta 1000 Energia elétrica", "(-) Cta 1000 Aluguéis e condomínios",
    "(-) Cta 1000 Telefone e internet", "(-) Cta 1000 Sky", "(-) Cta 1000 IPTU, alvará, taxa de lixo e VISA",
    "(-) Cta 1000 Manutenção e conservação de estrutura", "(-) Cta 1000 Pedágios", "(-) Cta 1000 Cartões",
    "(-) Cta 1000 Clube de Tiro", "(-) Cta 1000 Outras Despesas Particulares", "(-) Cta 1000 Unimed",
    "(-) Cta 1000 Farmácia", "(-) Cta 1000 Supermercado", "(-) Cta 1000 IPVA, seguro e multas",
    "(-) Cta 1000 Aplicações Sócios", "(-) Cta 1000 Investimentos CILA", "(-) Cta 1000 Investimentos OuroCap"]
_DRE_GRUPO_FINAIS = ["(+) Alugueis recebidos", "(-) Cta 2000 Leonardo"]
_DRE_CONTAS_CONHECIDAS = set(
    _DRE_GRUPO_RECEITA_BRUTA + _DRE_GRUPO_IMPOSTOS_RECEITA + _DRE_GRUPO_CUSTOS_VARIAVEIS
    + [c for _, cs in _DRE_SUBGRUPOS_DESPESAS_FIXAS for c in cs]
    + _DRE_GRUPO_RECEITAS_FINANCEIRAS + _DRE_GRUPO_DESPESAS_FINANCEIRAS + _DRE_GRUPO_IMPOSTOS_LUCRO
    + _DRE_GRUPO_OUTROS_MOVIMENTOS + _DRE_GRUPO_CONTA_1000 + _DRE_GRUPO_FINAIS)


def _dre_valor_assinado(info):
    return info["valor"] if info["tipo"] == "R" else -info["valor"]


def _dre_grupo_simples(nome, contas_lista, contas_mes, usadas):
    itens, soma = [], 0.0
    for c in contas_lista:
        info = contas_mes.get(c)
        if not info:
            continue
        usadas.add(c)
        itens.append({"tipo": "conta", "nome": c, "tipoDre": info["tipo"], "valor": info["valor"],
                      "lancamentos": info["lancamentos"]})
        soma += _dre_valor_assinado(info)
    return {"tipo": "grupo", "nome": nome, "valor": round(soma, 2), "contas": itens}


def _dre_grupo_composto(nome, subgrupos_def, contas_mes, usadas):
    subitens, soma = [], 0.0
    for sub_nome, lista in subgrupos_def:
        sub = _dre_grupo_simples(sub_nome, lista, contas_mes, usadas)
        if sub["contas"]:
            subitens.append(sub)
            soma += sub["valor"]
    return {"tipo": "grupo", "nome": nome, "valor": round(soma, 2), "contas": subitens}


def _dre_formula(nome, valor):
    return {"tipo": "formula", "nome": nome, "valor": round(valor, 2)}


def _dre_leaf_movimento(valor):
    return {"tipo": "conta", "nome": "(+) Receita Movimento Diário", "tipoDre": "R",
            "valor": round(valor, 2), "lancamentos": []}


def _hierarquia_dre(contas_mes, movimento_mes=None):
    """Aplica a estrutura acima sobre {conta: {tipo,valor,lancamentos}} de um mês/escopo e
    devolve a lista ordenada de blocos (grupo/subgrupo com contas e drill-down, ou formula
    - subtotal calculado, sem lançamentos próprios) pronta pro painel renderizar.

    movimento_mes: quando informado (só acontece no DRE de um posto individual), substitui
    as contas normais de receita bruta por uma única linha "Receita Movimento Diário" com o
    Líquido faturado da base de exames daquele posto/mês - porque, posto a posto, o
    faturamento por título não é confiável (fica tudo junto no "conjunto" entre as
    filiais), diferente do Geral/Consolidado, que usam os títulos de verdade."""
    usadas = set()
    if movimento_mes is not None:
        for c in _DRE_GRUPO_RECEITA_BRUTA:
            if c in contas_mes:
                usadas.add(c)
        receita_bruta = {"tipo": "grupo", "nome": "Receita bruta operacional",
                          "valor": round(movimento_mes, 2), "contas": [_dre_leaf_movimento(movimento_mes)]}
    else:
        receita_bruta = _dre_grupo_simples("Receita bruta operacional", _DRE_GRUPO_RECEITA_BRUTA, contas_mes, usadas)
    impostos_receita = _dre_grupo_simples("(-) Impostos sobre a receita", _DRE_GRUPO_IMPOSTOS_RECEITA, contas_mes, usadas)
    receita_liquida_v = receita_bruta["valor"] + impostos_receita["valor"]
    custos_variaveis = _dre_grupo_simples("(-) Custos de produção - variáveis", _DRE_GRUPO_CUSTOS_VARIAVEIS, contas_mes, usadas)
    margem_contribuicao_v = receita_liquida_v + custos_variaveis["valor"]
    despesas_fixas = _dre_grupo_composto("(-) Despesas operacionais - fixas", _DRE_SUBGRUPOS_DESPESAS_FIXAS, contas_mes, usadas)
    ebitda_v = margem_contribuicao_v + despesas_fixas["valor"]
    receitas_financeiras = _dre_grupo_simples("Receitas financeiras", _DRE_GRUPO_RECEITAS_FINANCEIRAS, contas_mes, usadas)
    despesas_financeiras = _dre_grupo_simples("(-) Despesas financeiras", _DRE_GRUPO_DESPESAS_FINANCEIRAS, contas_mes, usadas)
    resultado_financeiro_v = receitas_financeiras["valor"] + despesas_financeiras["valor"]
    impostos_lucro = _dre_grupo_simples("(-) Impostos sobre o lucro", _DRE_GRUPO_IMPOSTOS_LUCRO, contas_mes, usadas)
    resultado_liquido_op_v = ebitda_v + resultado_financeiro_v + impostos_lucro["valor"]
    outros_movimentos = _dre_grupo_simples("Outros movimentos", _DRE_GRUPO_OUTROS_MOVIMENTOS, contas_mes, usadas)
    conta_1000 = _dre_grupo_simples("Total Conta 1000", _DRE_GRUPO_CONTA_1000, contas_mes, usadas)
    finais = _dre_grupo_simples("Outros lançamentos", _DRE_GRUPO_FINAIS, contas_mes, usadas)
    resultado_liquido_final_v = resultado_liquido_op_v + outros_movimentos["valor"] + conta_1000["valor"] + finais["valor"]

    hierarquia = [receita_bruta, impostos_receita, _dre_formula("Receita líquida", receita_liquida_v),
                  custos_variaveis, _dre_formula("Margem de contribuição", margem_contribuicao_v),
                  despesas_fixas, _dre_formula("EBITDA", ebitda_v)]
    if receitas_financeiras["contas"] or despesas_financeiras["contas"]:
        hierarquia += [receitas_financeiras, despesas_financeiras,
                       _dre_formula("Resultado financeiro", resultado_financeiro_v)]
    if impostos_lucro["contas"]:
        hierarquia.append(impostos_lucro)
    hierarquia.append(_dre_formula("Resultado líquido operacional", resultado_liquido_op_v))
    extras_finais = [g for g in (outros_movimentos, conta_1000, finais) if g["contas"]]
    if extras_finais:
        hierarquia += extras_finais
        hierarquia.append(_dre_formula("Resultado líquido", resultado_liquido_final_v))
    else:
        # sem lançamentos "fora do operacional" neste mês/escopo - o resultado líquido
        # operacional já é o final, não repete a mesma linha duas vezes
        hierarquia[-1] = _dre_formula("Resultado líquido", resultado_liquido_op_v)

    sobrando = {c: info for c, info in contas_mes.items() if c not in usadas}
    if sobrando:
        extra = {"tipo": "grupo", "nome": "(-) Outras contas (a posicionar na hierarquia)",
                 "valor": round(sum(_dre_valor_assinado(i) for i in sobrando.values()), 2),
                 "contas": [{"tipo": "conta", "nome": c, "tipoDre": i["tipo"], "valor": i["valor"],
                             "lancamentos": i["lancamentos"]} for c, i in sobrando.items()]}
        hierarquia.append(extra)
        print("AVISO: DRE - conta(s) sem posição definida na hierarquia (caíram em "
              "'(-) Outras contas (a posicionar na hierarquia)', ajustar em gerar_dashboard.py): "
              + ", ".join(sorted(sobrando)))
    return hierarquia


def _aplicar_hierarquia(flat_por_mes, movimento_por_mes=None):
    return {mes: _hierarquia_dre(contas, (movimento_por_mes or {}).get(mes))
            for mes, contas in flat_por_mes.items()}


# Postos com DRE individual - só os 7 que o Claudio acompanhava separadamente (os outros,
# Quintino/Laranjeiras/Administrativo, não tem botão nem filtro próprio, mas continuam
# entrando no Geral - e Quintino/Laranjeiras também no Consolidado, como já era).
_DRE_POSTOS_INDIVIDUAIS = {"1", "2", "4", "7", "8", "10", "13"}


def _montar_dre(arqs, movimento_diario=None):
    """Monta o DRE completo (por posto + Consolidado [01-13, sem o 100] + Geral [tudo]),
    em regime de caixa. Devolve None enquanto os arquivos de valor_titulo não tiverem
    sido enviados pela VM (contas_a_pagar/contas_a_receber_valor_titulo.csv)."""
    if "valor_titulo" not in arqs and "valor_titulo_receber" not in arqs:
        return None

    mapa_dre = carregar_mapa_dre()

    valor_titulo_pagar = _carregar_valor_titulo_dre(arqs, "valor_titulo")
    valor_titulo_receber = _carregar_valor_titulo_dre(arqs, "valor_titulo_receber")

    rateio_pagar = _parse_rateio_dre(arqs, "categoria")
    rateio_receber = _parse_rateio_dre(arqs, "categoria_receber")

    pago_mes_pagar = _agrupar_pago_por_mes(_parse_baixas_dre(arqs, "baixas"))
    pago_mes_receber = _agrupar_pago_por_mes(_parse_baixas_dre(arqs, "baixas_receber"))

    nao_classificados = set()
    eventos = (_eventos_dre(rateio_pagar, pago_mes_pagar, valor_titulo_pagar, mapa_dre, nao_classificados)
               + _eventos_dre(rateio_receber, pago_mes_receber, valor_titulo_receber, mapa_dre, nao_classificados)
               + _parse_movcxb_dre(arqs, mapa_dre, nao_classificados))

    if nao_classificados:
        print("AVISO: DRE - código(s) TPDRCOD sem classificação em data/apoio/mapa_dre.json "
              "(perguntar pro Leo antes de adicionar, nunca chutar): "
              + ", ".join(f"{c} ({d})" for c, d in sorted(nao_classificados)))

    if not eventos:
        return None

    postos = sorted({e["filial"] for e in eventos if e["filial"] in _DRE_POSTOS_INDIVIDUAIS},
                     key=lambda f: (len(f), f))
    por_posto = {p: _aplicar_hierarquia(_resumo_dre([e for e in eventos if e["filial"] == p]),
                                         (movimento_diario or {}).get(p)) for p in postos}
    consolidado = _aplicar_hierarquia(_resumo_dre([e for e in eventos if e["filial"] != "100"]))
    geral = _aplicar_hierarquia(_resumo_dre(eventos))
    meses = sorted({e["mes"] for e in eventos})

    return {
        "meses": meses,
        "postos": postos,
        "porPosto": por_posto,
        "consolidado": consolidado,
        "geral": geral,
        "naoClassificados": sorted(f"{c} - {d}" for c, d in nao_classificados),
    }


# --------------------------------------------------------------------------- main
def carregar_precos_convenio(apoio):
    """Preços por convênio, exportados direto do banco do CONCENT, cruzados com o
    custo do laboratório de apoio pelo código do exame. Quando um convênio tem mais
    de um plano para o mesmo exame, fica só o de maior valor (o plano em si não é
    mostrado no painel). Devolve None se o arquivo não foi enviado."""
    arqs = achar_arquivos_apoio()
    if "precos_conv" not in arqs:
        return None
    print(f"Lendo {arqs['precos_conv'].name}...")
    linhas = _ler_precos_convenio(arqs["precos_conv"])

    custo_db, custo_hp = {}, {}
    if apoio:
        for r in apoio["rows"]:
            for lado, alvo in (("d", custo_db), ("h", custo_hp)):
                o = r[lado]
                if o and o["p"]:
                    v = min(o["p"])
                    alvo[r["c"]] = min(alvo.get(r["c"], v), v)

    # agrupa por (exame, convênio): entre os planos do mesmo convênio, fica o de maior
    # valor - preferindo um plano ativo quando existir algum ativo nesse convênio.
    grupos = {}  # (cod_exame, cod_conv) -> {"nome":..., "convn":..., "ativo":bool, "valor":float}
    for cod_ex, nome_ex, cod_conv, convn, ativo, valor in linhas:
        chave = (cod_ex, cod_conv)
        atual = grupos.get(chave)
        if atual is None:
            grupos[chave] = {"nome": nome_ex, "convn": convn, "ativo": ativo, "valor": valor}
            continue
        # um plano ativo sempre vence um inativo; dentro do mesmo status, fica o maior valor
        troca = ativo if ativo != atual["ativo"] else valor > atual["valor"]
        if troca:
            atual["ativo"], atual["valor"] = ativo, valor

    exames, exames_cod, ex_idx = [], [], {}
    custo_apoio_db, custo_apoio_hp = [], []
    convenios, cv_idx = [], {}
    linhas_out = []
    for (cod_ex, cod_conv), g in grupos.items():
        if cod_ex not in ex_idx:
            ex_idx[cod_ex] = len(exames)
            exames.append(g["nome"])
            exames_cod.append(cod_ex)
            custo_apoio_db.append(custo_db.get(cod_ex))
            custo_apoio_hp.append(custo_hp.get(cod_ex))
        if cod_conv not in cv_idx:
            cv_idx[cod_conv] = len(convenios)
            convenios.append({"c": cod_conv, "n": g["convn"]})
        linhas_out.append([ex_idx[cod_ex], cv_idx[cod_conv], 1 if g["ativo"] else 0, round(g["valor"], 2)])
    linhas_out.sort(key=lambda r: (r[0], r[1]))

    com_apoio = sum(1 for i in range(len(exames)) if custo_apoio_db[i] is not None or custo_apoio_hp[i] is not None)
    print(f"  {len(exames):,} exames, {len(convenios):,} convênios, {len(linhas_out):,} preços "
          f"(veio de {len(linhas):,} linhas do banco, antes de juntar planos) "
          f"({com_apoio:,} exames com custo de apoio para comparar)".replace(",", "."))
    return {
        "exames": exames,
        "exameCod": exames_cod,
        "custoDB": custo_apoio_db,
        "custoHP": custo_apoio_hp,
        "convenios": convenios,
        "linhas": linhas_out,
        "arquivo": arqs["precos_conv"].name,
    }


def main():
    arquivos = achar_arquivos_dados()
    partes, resumo = [], []
    for caminho, prio, rotulo in arquivos:
        print(f"Lendo {caminho.name} ...")
        d = preparar(carregar_planilha(caminho), caminho.name)
        d["prio"] = prio
        print(f"  {len(d):,} linhas ({d['dia'].min():%d/%m/%Y} a {d['dia'].max():%d/%m/%Y})".replace(",", "."))
        resumo.append({"nome": caminho.name, "tipo": "dia" if prio else "ano", "linhas": int(len(d)),
                       "de": d["dia"].min().strftime("%Y-%m-%d"), "ate": d["dia"].max().strftime("%Y-%m-%d")})
        partes.append(d)
    df = pd.concat(partes, ignore_index=True)
    del partes
    df, subst = resolver_sobreposicao(df)
    if subst:
        print(f"  {subst:,} linhas de arquivos mais antigos foram substituídas por versões mais novas".replace(",", "."))
    df, excl = aplicar_exclusoes(df)
    medicos = carregar_medicos()
    dados = construir_dados(df, medicos, excl, resumo, subst)
    dados["apoio"] = carregar_apoio()
    if dados["apoio"]:
        r = dados["apoio"]["resumo"]
        print(f"Laboratório de apoio: {r['exames']} exames da lista, {r['com_db']} com preço DB, {r['com_hp']} com preço HP, {r['ambos']} nos dois")
    dados["precosConv"] = carregar_precos_convenio(dados["apoio"])
    dados["areaTecnica"] = carregar_area_tecnica()
    if dados["areaTecnica"]:
        r = dados["areaTecnica"]["resumo"]
        print(f"Área técnica: {r['emAndamento']} exames em andamento ({r['atrasados']} atrasados), "
              f"{r['aguardandoColeta']} aguardando coleta.")
    dados["reajustesDB"] = carregar_reajustes_db()
    if dados["reajustesDB"]:
        ult = dados["reajustesDB"][-1]
        print(f"Histórico de reajustes DB: {len(dados['reajustesDB'])} evento(s); último em {ult['data']} "
              f"({len(ult['linhas'])} exames mudaram).")
    # Líquido faturado por posto/mês (base de exames, não títulos) - usado no DRE por
    # posto individual, onde o faturamento por título vem todo junto no "conjunto" e não
    # dá pra atribuir com confiança a uma filial (ver _montar_dre/_hierarquia_dre).
    mov = df.groupby(["uni", "ym"])["Líquido"].sum().reset_index()
    movimento_diario = {}
    for r in mov.itertuples():
        filial = str(int(r.uni))
        movimento_diario.setdefault(filial, {})[r.ym] = round(float(r.Líquido), 2)
    dados["financeiro"] = carregar_financeiro(movimento_diario)
    if dados["financeiro"]:
        rf = dados["financeiro"]["resumo"]
        print(f"Financeiro: {rf['vencidas']['n']} contas vencidas (R$ {rf['vencidas']['v']:,.2f}), "
              f"{rf['hoje']['n']} vencendo hoje, {dados['financeiro']['recorrencia']['n']} recorrentes sem lançamento no mês.".replace(",", "."))
    dados["filiais"] = carregar_filiais()
    if dados["filiais"]:
        print(f"Filiais: {len(dados['filiais'])} cadastradas em data/filiais.")
    payload = json.dumps(dados, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = TEMPLATE.read_text(encoding="utf-8")
    if "/*__DATA__*/null" not in html:
        sys.exit("ERRO: o template.html não tem o marcador /*__DATA__*/null")
    if SENHA:
        from cripto_arquivos import criptografar_para_pagina
        payload = json.dumps(criptografar_para_pagina(payload, SENHA), separators=(",", ":"))
        print("Painel protegido por senha (dados criptografados no index.html).")
    SAIDA.write_text(html.replace("/*__DATA__*/null", payload), encoding="utf-8")
    m = dados["meta"]
    print(f"OK -> {SAIDA.name} | {m['dataMin']} a {m['dataMax']} | "
          f"{len(dados['A']['d']):,} linhas de resumo | {SAIDA.stat().st_size/1e6:.1f} MB".replace(",", "."))


if __name__ == "__main__":
    main()
