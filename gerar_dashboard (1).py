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
import tempfile
import unicodedata
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
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


def carregar_medicos():
    """Devolve dict codigo -> [nomes distintos]. Códigos com mais de um nome ficam ambíguos."""
    p = PASTA_DADOS / "lista_medico.xlsx"
    if not p.exists() and (PASTA_DADOS / "lista_medico.xlsx.enc").exists():
        p = _abrir_enc(PASTA_DADOS / "lista_medico.xlsx.enc")
    if not p.exists():
        print("AVISO: data/lista_medico.xlsx não encontrado - médicos aparecerão só pelo código.")
        return {}
    pares = None
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

    conv = (df.groupby("Convênio")["Nome Convênio"].first().sort_index())
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
    "lista": ("lista",),
    "de_db": ("de_para_concent_db",),
    "de_hp": ("de_para_concent_hp",),
    "tab_db": ("tabela_db",),
    "tab_hp": ("tabela_pardini", "tabela_hp", "tabela_padini"),
}
APOIO_OPCIONAL = {"materiais", "precos_conv"}
NOME_PAPEL = {
    "materiais": "Lista de materiais biológicos",
    "precos_conv": "Relatório de convênios por exame (preços)",
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
        nome_conv = _cel(r[1]).strip()
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
    for r in dep_hp:   # reserva: descrição do material no de-para do HP
        if r["desc"] and r["mat"] not in mat_nome:
            mat_nome[r["mat"]] = r["desc"]
    nomes_c = {r["cod"]: r["nome_c"] for r in dep_db + dep_hp}
    na_lista = {c for c, _, _ in lista}
    exames = [(c, n, pz, 1) for c, n, pz in lista]
    extras = sorted((set(op_db) | set(op_hp)) - na_lista)
    exames += [(c, nomes_c.get(c, c), "", 0) for c in extras]

    def lado(o):
        if not o:
            return None
        d = {"c": o["apo"], "n": o["nome"], "p": o["p"]}
        if o["e"]:
            d["e"] = o["e"]
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
            do_exame[chave] = {"c": cod, "n": nome, "m": rot, "l": na, "d": lado(d), "h": lado(h), "w": w,
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

    exames, ex_idx = [], {}
    custo_apoio_db, custo_apoio_hp = [], []
    convenios, cv_idx = [], {}
    linhas_out = []
    for (cod_ex, cod_conv), g in grupos.items():
        if cod_ex not in ex_idx:
            ex_idx[cod_ex] = len(exames)
            exames.append(g["nome"])
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
