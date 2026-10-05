"""Versão de DEMONSTRAÇÃO do painel: "espelho disfarçado".

Pega os dados reais já agregados (o mesmo dicionário que vai criptografado para o
index.html) e troca tudo que identifica alguém ou algum valor real:

- laboratório, unidades, convênios, médicos, fornecedores e pacientes viram nomes
  inventados (sempre os MESMOS nomes para o mesmo original, porque a troca é feita
  por hash: a demo é regerada toda hora e não pode mudar de nomes a cada atualização);
- todos os valores em dinheiro são multiplicados por um fator fixo (FATOR_DINHEIRO) e
  as quantidades (atendimentos, exames, pacientes) por outro (FATOR_QTD), para que
  nem o ticket médio seja o real;
- preços de tabela (apoio, convênios) têm um fator próprio (FATOR_PRECO).

A estrutura continua a de um laboratório de verdade (sazonalidade, convênio dominante,
exames mais comuns), que é o que faz a demo convencer. Nenhum nome ou número real sai.

Uso: gerar_dashboard.py chama disfarcar(dados) quando DASHBOARD_DEMO=1.
"""
import hashlib
import re

LAB_NOME = "Laboratório Horizonte"
LAB_LINHA = "Horizonte · Análises Clínicas"
FATOR_DINHEIRO = 0.82
FATOR_QTD = 0.71
FATOR_PRECO = 0.82   # igual ao do dinheiro, para a margem com o apoio continuar parecida com a real

UNIDADES = ["Centro", "Bairro Alto", "Jardim América", "Vila Nova", "São Pedro",
            "Industrial", "Lagoa", "Parque das Flores", "Boa Vista", "Santa Rita",
            "Alvorada", "Primavera", "Horizonte Sul", "Administrativo"]

CONVENIOS = ["Plano Norte", "Vida Saúde", "Amparo", "Bem Viver", "Saúde Total", "Cooperativa Médica",
             "Previdência Municipal", "Medicina do Trabalho ABC", "Clube de Benefícios", "Seguro Vida",
             "Consórcio Regional", "Assistência Serrana", "Plano Família", "Saúde Empresarial",
             "Fundação dos Servidores", "Caixa de Assistência", "Plano Rural", "Mais Saúde", "Protege",
             "Viva Bem", "Rede Cuidar", "Aliança Saúde", "União Médica", "Plano Litoral", "Saúde Já",
             "Prevenir", "Cooperação Sul", "Plano Universitário", "Instituto Municipal", "Bem Estar",
             "Assistência Ferroviária", "Clínica Popular", "Saúde Metal", "Plano Docente", "Vida Plena",
             "Convênio Industrial", "Programa Sênior", "Sindicato Comércio", "Plano Bancário", "Saúde Cidadã"]
CONV_MANTER = {"PARTICULAR", "CORTESIA", "SUS", "CHECK UP", "CONVÊNIO DIVERSOS", "CONVENIO DIVERSOS"}

NOMES = ["Ana", "Bruno", "Carla", "Daniel", "Eduarda", "Felipe", "Gabriela", "Henrique", "Isabela", "João",
         "Karina", "Lucas", "Mariana", "Nicolas", "Olívia", "Pedro", "Rafaela", "Samuel", "Tatiana", "Vinícius",
         "Beatriz", "Caio", "Débora", "Enzo", "Fernanda", "Gustavo", "Helena", "Igor", "Juliana", "Leonardo",
         "Larissa", "Marcelo", "Natália", "Otávio", "Patrícia", "Renato", "Sofia", "Thiago", "Valentina", "William",
         "Alice", "Bernardo", "Camila", "Diego", "Elisa", "Fábio", "Giovana", "Heitor", "Ingrid", "Júlio",
         "Letícia", "Murilo", "Nina", "Paulo", "Raquel", "Sérgio", "Tereza", "Vitor", "Yasmin", "André"]
SOBRENOMES = ["Almeida", "Barbosa", "Cardoso", "Dias", "Esteves", "Ferreira", "Gomes", "Hoffmann", "Ibrahim", "Jesus",
              "Klein", "Lima", "Martins", "Nogueira", "Oliveira", "Pereira", "Queiroz", "Ribeiro", "Santos", "Teixeira",
              "Uchoa", "Vieira", "Weber", "Xavier", "Zanella", "Araújo", "Bastos", "Carvalho", "Duarte", "Fonseca",
              "Guimarães", "Henriques", "Lopes", "Machado", "Nunes", "Pacheco", "Rocha", "Siqueira", "Tavares", "Vargas",
              "Andrade", "Borges", "Castro", "Dantas", "Farias", "Garcia", "Lacerda", "Medeiros", "Neves", "Peixoto",
              "Ramos", "Sales", "Toledo", "Valente", "Azevedo", "Brito", "Campos", "Freitas", "Moraes", "Pinto"]
EMPRESAS_A = ["Comercial", "Distribuidora", "Serviços", "Tecnologia", "Transportes", "Suprimentos", "Engenharia",
              "Manutenção", "Limpeza", "Segurança", "Gráfica", "Papelaria", "Informática", "Energia", "Água e Saneamento",
              "Telecom", "Contabilidade", "Imobiliária", "Laboratório", "Equipamentos", "Logística", "Consultoria",
              "Uniformes", "Reagentes", "Descartáveis", "Climatização", "Elétrica", "Hidráulica", "Vidraçaria", "Marcenaria"]
EMPRESAS_B = ["Horizonte", "Boa Vista", "Primavera", "Serrana", "do Vale", "Central", "Nacional", "Regional", "Sul",
              "Norte", "Alfa", "Beta", "Delta", "Ômega", "Prime", "Ideal", "Moderna", "Única", "Real", "Nova"]
PAGAMENTOS = {"", "TED", "DEBITO", "DÉBITO", "CREDITO", "CRÉDITO", "PIX", "DINHEIRO", "CHEQUE", "BOLETO", "CARTAO", "CARTÃO"}
IMPOSTOS = {"INSS", "ISSQN", "FGTS", "IRRF", "PIS", "COFINS", "CSLL", "IRPJ", "SIMPLES", "DAS", "MULTAS", "ICMS", "IPTU", "ALVARÁ", "ALVARA"}


def _h(texto, mod):
    return int(hashlib.md5(texto.encode("utf-8")).hexdigest(), 16) % mod


def pessoa(nome):
    """Nome fictício, sempre o mesmo para o mesmo original."""
    n = str(nome or "").strip()
    if not n or n.upper() == "PARTICULAR":
        return n
    k = n.upper()
    a = NOMES[_h("a" + k, len(NOMES))]
    b = SOBRENOMES[_h("b" + k, len(SOBRENOMES))]
    c = SOBRENOMES[_h("c" + k, len(SOBRENOMES))]
    return f"{a} {b} {c}".upper() if n.isupper() else f"{a} {b} {c}"


def empresa(nome):
    n = str(nome or "").strip()
    if not n or n.upper() in PAGAMENTOS:
        return n
    k = n.upper()
    for imp in IMPOSTOS:
        if k.startswith(imp):
            return n           # imposto é imposto em qualquer laboratório
    return f"{EMPRESAS_A[_h('e' + k, len(EMPRESAS_A))]} {EMPRESAS_B[_h('f' + k, len(EMPRESAS_B))]} Ltda".upper()


_RE_SUFIXO = re.compile(r"\s*\((\d+)\)\s*$")


# lista grande de nomes de convênio (base x variação) para não repetir nome entre convênios diferentes
_CONV_LISTA = CONVENIOS + [b + " " + v for v in ("Regional", "Plus", "Mais", "Norte", "Sul") for b in CONVENIOS]
_conv_mapa, _conv_usados = {}, set()


def convenio(nome):
    """Nome fictício de convênio, sem repetir entre convênios diferentes: começa no índice do
    hash e, se já estiver em uso, pega o próximo livre. Como a ordem de chegada é a mesma a
    cada geração, o resultado é estável."""
    n = str(nome or "").strip()
    if not n:
        return n
    m = _RE_SUFIXO.search(n)
    base, suf = (n[:m.start()].strip(), m.group(0)) if m else (n, "")
    k = base.upper()
    if k in CONV_MANTER or base.startswith("Outros convênios"):
        return n
    if k not in _conv_mapa:
        i = _h("c" + k, len(_CONV_LISTA))
        while _CONV_LISTA[i % len(_CONV_LISTA)] in _conv_usados:
            i += 1
        _conv_mapa[k] = _CONV_LISTA[i % len(_CONV_LISTA)]
        _conv_usados.add(_conv_mapa[k])
    return _conv_mapa[k] + suf


def _medico_especial(nome):
    n = str(nome or "").lower()
    return bool(re.search(r"solicitac|propri|sem medico|sem médico|^cod\.", n))


def medico(nome):
    return nome if _medico_especial(nome) else pessoa(nome)


def _rd(v, f, casas=2):
    return None if v is None else round(v * f, casas)


def _q(v, f):
    return None if v is None else int(round(v * f))


def _escala_rec(o, chave=None):
    """Percorre dicionários/listas do financeiro escalando dinheiro e trocando nomes."""
    if isinstance(o, dict):
        for k, v in o.items():
            if k in ("s", "v", "valor", "total", "totalGeral", "pagoMes", "projecaoMes", "recebidoMes") and isinstance(v, (int, float)) and not isinstance(v, bool):
                o[k] = _rd(v, FATOR_DINHEIRO)
            elif k in ("valores", "totalMes") and isinstance(v, list):
                o[k] = [_rd(x, FATOR_DINHEIRO) for x in v]
            elif k == "n" and isinstance(v, int) and not isinstance(v, bool):
                o[k] = _q(v, FATOR_QTD)
            elif k == "f":
                o[k] = pessoa(v) if o.get("tp") == "P" else empresa(v)
            elif k == "fornecedor":
                o[k] = empresa(v)
            elif k == "paciente":
                o[k] = pessoa(v)
            elif k == "convenio":
                o[k] = convenio(v)
            elif k in ("t", "titulo") and isinstance(v, str) and re.search(r"[A-Za-zÀ-ú]{3,}", v):
                o[k] = "DOC-" + str(_h("t" + v, 900000) + 100000)   # título com texto pode carregar nome
            elif k == "historico":
                o[k] = ""          # texto livre: pode ter nome de gente, de unidade, de tudo
            elif k == "requisicao" and isinstance(v, str) and v.isdigit():
                o[k] = str(int(v) + 100000)
            else:
                _escala_rec(v, k)
    elif isinstance(o, list):
        for x in o:
            _escala_rec(x, chave)


def disfarcar(dados):
    fm, fq, fp = FATOR_DINHEIRO, FATOR_QTD, FATOR_PRECO
    dados["meta"]["demo"] = True
    dados["meta"]["labNome"] = LAB_NOME

    # unidades / filiais
    for i, u in enumerate(dados.get("unidades") or []):
        u["nome"] = UNIDADES[i % len(UNIDADES)]
    if dados.get("filiais"):
        for i, k in enumerate(sorted(dados["filiais"], key=lambda x: int(x) if str(x).isdigit() else 999)):
            dados["filiais"][k] = UNIDADES[i % len(UNIDADES)]

    # convênios, médicos
    for c in dados.get("convenios") or []:
        c["nome"] = convenio(c["nome"])
    for m in dados.get("medicos") or []:
        m["nome"] = medico(m["nome"])

    # fatos
    A = dados["A"]
    A["r"] = [_q(v, fq) for v in A["r"]]
    A["e"] = [_q(v, fq) for v in A["e"]]
    for k in ("b", "x", "l"):
        A[k] = [_rd(v, fm) for v in A[k]]
    B = dados["B"]
    B["e"] = [_q(v, fq) for v in B["e"]]
    B["l"] = [_rd(v, fm) for v in B["l"]]
    if dados.get("AP"):
        AP = dados["AP"]
        AP["e"] = [_q(v, fq) for v in AP["e"]]
        AP["l"] = [_rd(v, fm) for v in AP["l"]]

    # qualidade dos dados
    Q = dados.get("qualidade") or {}
    for k in ("linhas", "requisicoes", "linhas_substituidas", "linhas_valor_zero", "req_multi_convenio", "req_multi_medico", "req_multi_dia"):
        if k in Q:
            Q[k] = _q(Q[k], fq)
    for z in Q.get("zero_por_convenio") or []:
        z["nome"] = convenio(z["nome"]); z["linhas"] = _q(z["linhas"], fq)
    Q["convenios_nome_repetido"] = [convenio(x) for x in Q.get("convenios_nome_repetido") or []]
    if Q.get("excluidos"):
        ex = Q["excluidos"]
        ex["linhas"] = _q(ex.get("linhas", 0), fq); ex["valor"] = _rd(ex.get("valor", 0), fm)
        for u in ex.get("unidades") or []:
            u["linhas"] = _q(u["linhas"], fq); u["valor"] = _rd(u["valor"], fm)
    for a in Q.get("arquivos") or []:
        if "linhas" in a:
            a["linhas"] = _q(a["linhas"], fq)

    # apoio (preços de tabela)
    ap = dados.get("apoio")
    if ap:
        for r in ap["rows"]:
            for lado in ("d", "h"):
                o = r.get(lado)
                if o and o.get("p"):
                    o["p"] = [_rd(v, fp) for v in o["p"]]
        for d in (ap.get("atencao") or {}).get("dup_hp") or []:
            if d.get("p"):
                d["p"] = [_rd(v, fp) for v in d["p"]]

    pc = dados.get("precosConv")
    if pc:
        pc["custoDB"] = [_rd(v, fp) for v in pc["custoDB"]]
        pc["custoHP"] = [_rd(v, fp) for v in pc["custoHP"]]
        for c in pc["convenios"]:
            c["n"] = convenio(c["n"])
        for l in pc["linhas"]:
            l[3] = _rd(l[3], fm)

    for ev in dados.get("reajustesDB") or []:
        for l in ev.get("linhas") or []:
            for k in ("pa", "pn", "d", "p"):
                if k in l and l[k] is not None:
                    l[k] = _rd(l[k], fp)

    # área técnica: nomes de pacientes
    at = dados.get("areaTecnica")
    if at:
        for lista in ("rows", "aguardandoColeta"):
            for r in at.get(lista) or []:
                if "paciente" in r:
                    r["paciente"] = pessoa(r["paciente"])
                if "req" in r and isinstance(r["req"], int):
                    r["req"] += 100000

    te = dados.get("tempoEntrega")
    if te:
        for l in te["linhas"]:
            l[3] = max(1, _q(l[3], fq)); l[4] = min(l[3], _q(l[4], fq))
            l[5] = _rd(l[5], fq, 1); l[6] = _rd(l[6], fq, 1)   # somas de horas acompanham a quantidade

    pa = dados.get("pacientes")
    if pa:
        for l in pa["linhas"]:
            for i in (2, 3, 4, 5):
                l[i] = _q(l[i], fq)

    # financeiro: recursivo
    if dados.get("financeiro"):
        _escala_rec(dados["financeiro"])
    return dados
