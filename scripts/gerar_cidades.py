"""Gera cidades.json: postos com gasolina comum por municipio (ANP semanal, Brasil todo)
+ cidades fora da pesquisa com postos do cadastro oficial da ANP (sem preco).
Dados publicos da ANP. Sem segredos."""
import io, json, re, sys, datetime, urllib.request
import openpyxl

PAGE = "https://www.gov.br/anp/pt-br/assuntos/precos-e-defesa-da-concorrencia/precos/levantamento-de-precos-de-combustiveis-ultimas-semanas-pesquisadas"
CADASTRO = "https://www.gov.br/anp/pt-br/centrais-de-conteudo/dados-abertos/arquivos/arquivos-dados-cadastrais-dos-revendedores-varejistas-de-combustiveis-automotivos/dados-cadastrais-revendedores-varejistas-combustiveis-automoveis.csv"
# cidades fora da pesquisa semanal que entram via cadastro oficial (sem preco):
CADASTRO_CIDADES = {"JANDAIA DO SUL|PR"}
# postos cujo municipio oficial difere mas que pertencem a cidade na pratica (ex.: BR-369, "saida para Jandaia do Sul"):
CADASTRO_EXTRA = {"JANDAIA DO SUL|PR": {"06243880000103"}}  # P.H.G. Fernandes Dias (oficial: Bom Sucesso)
# postos sempre incluidos na propria cidade mesmo sem coleta na semana (ultimo preco disponivel):
EXTRAS = {"APUCARANA|PR": ["10601388000130"]}  # Posto Brasil Sul II
NOMES = {"10601388000130": "Posto Brasil Sul II (Brz Sul Comercio De Combustiveis Ltda)"}
MAX_SEMANAS = 8
UA = {"User-Agent": "Mozilla/5.0"}

def get(url):
    return urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=180).read()

def links():
    html = get(PAGE).decode("utf-8", "ignore")
    found = re.findall(r'href="([^"]*revendas_lpc_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})\.xlsx)"', html)
    uniq = {}
    for u, a, b in found:
        uniq[a] = (u, a, b)
    return [uniq[k] for k in sorted(uniq, reverse=True)]

def rows(url):
    """{key 'MUNICIPIO|UF': {cnpj: posto}} da semana, so gasolina comum."""
    wb = openpyxl.load_workbook(io.BytesIO(get(url)), read_only=True)
    out = {}
    for r in wb.active.iter_rows(min_row=11, values_only=True):
        if not r or r[0] is None:
            continue
        if str(r[11]).strip().upper() != "GASOLINA COMUM":
            continue
        muni = str(r[8]).strip().upper()
        uf = str(r[9]).strip().upper()
        uf = {"PARANA": "PR"}.get(uf, uf)  # arquivo usa nome por extenso
        UFMAP = {"ACRE":"AC","ALAGOAS":"AL","AMAPA":"AP","AMAZONAS":"AM","BAHIA":"BA","CEARA":"CE","DISTRITO FEDERAL":"DF","ESPIRITO SANTO":"ES","GOIAS":"GO","MARANHAO":"MA","MATO GROSSO":"MT","MATO GROSSO DO SUL":"MS","MINAS GERAIS":"MG","PARA":"PA","PARAIBA":"PB","PARANA":"PR","PERNAMBUCO":"PE","PIAUI":"PI","RIO DE JANEIRO":"RJ","RIO GRANDE DO NORTE":"RN","RIO GRANDE DO SUL":"RS","RONDONIA":"RO","RORAIMA":"RR","SANTA CATARINA":"SC","SAO PAULO":"SP","SERGIPE":"SE","TOCANTINS":"TO"}
        uf = UFMAP.get(uf, uf)
        cnpj = str(int(r[0])).zfill(14)
        d = r[14]
        posto = {
            "cnpj": cnpj,
            "nome": (str(r[2] or r[1] or "")).strip().title(),
            "bandeira": (str(r[10] or "")).strip().title(),
            "endereco": f"{str(r[3] or '').strip().title()}, {r[4]}".strip(", "),
            "bairro": (str(r[6] or "")).strip().title(),
            "preco": float(r[13]),
            "coleta": d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d),
        }
        out.setdefault(f"{muni}|{uf}", {})[cnpj] = posto
    return out

def cadastro():
    """Postos oficiais (autorizados) das cidades sem pesquisa semanal. Sem preco."""
    raw = get(CADASTRO)
    try:
        txt = raw.decode("utf-8")
    except UnicodeDecodeError:
        txt = raw.decode("latin-1", "ignore")
    out = {}
    for line in txt.splitlines()[1:]:
        p = line.split(";")
        if len(p) < 13:
            continue
        key = f"{p[10].strip().upper()}|{p[9].strip().upper()}"
        if key not in CADASTRO_CIDADES:
            cnpj_row = str(p[4]).strip().zfill(14)
            alt = [k for k, cnpjs in CADASTRO_EXTRA.items() if cnpj_row in cnpjs]
            if not alt:
                continue
            key = alt[0]
        out.setdefault(key, []).append({
            "cnpj": str(p[4]).strip().zfill(14),
            "nome": p[3].strip().title(),
            "bandeira": (p[12].strip().title() or "Bandeira Branca").replace("Bandeira Branca", "Bandeira branca"),
            "endereco": f"{p[5].strip().title()}, {p[6].strip()}".strip(", "),
            "bairro": p[7].strip().title(),
        })
    return out

def titulo(muni):
    low = {"do","da","de","dos","das","e"}
    return " ".join(w if w.lower() not in low or i == 0 else w.lower()
                    for i, w in enumerate(muni.title().split()))

def main():
    ls = links()
    if not ls:
        sys.exit("nenhum link de planilha encontrado")
    cur_url, a, b = ls[0]
    cur = rows(cur_url)
    if not cur:
        sys.exit("planilha semanal vazia")
    prev = rows(ls[1][0]) if len(ls) > 1 else {}
    cache = {0: cur, 1: prev}
    def week(i):
        if i not in cache:
            try:
                cache[i] = rows(ls[i][0])
            except Exception:
                cache[i] = {}
        return cache[i]
    # extras (ex.: Brasil Sul II em Apucarana) com ultimo preco disponivel
    for key, cnpjs in EXTRAS.items():
        city = cur.setdefault(key, {})
        for cnpj in cnpjs:
            if cnpj in city:
                continue
            for i in range(1, min(len(ls), MAX_SEMANAS)):
                if cnpj in week(i).get(key, {}):
                    p = dict(week(i)[key][cnpj])
                    p["desatualizado"] = True
                    p["nome"] = NOMES.get(cnpj, p["nome"])
                    old = None
                    for j in range(i + 1, min(len(ls), MAX_SEMANAS + 1)):
                        if cnpj in week(j).get(key, {}):
                            old = week(j)[key][cnpj]["preco"]
                            break
                    p["_prev"] = old
                    city[cnpj] = p
                    break
    cidades = {}
    for key, postos_map in sorted(cur.items()):
        muni, uf = key.split("|")
        prev_map = prev.get(key, {})
        postos = []
        for cnpj, p in sorted(postos_map.items(), key=lambda kv: kv[1]["preco"]):
            if "_prev" in p:
                old = p.pop("_prev")
            else:
                old = prev_map.get(cnpj, {}).get("preco")
            p["preco_anterior"] = old
            if old is None:
                p["variacao"] = "novo"
            elif round(p["preco"] - old, 2) > 0:
                p["variacao"] = "subiu"
            elif round(p["preco"] - old, 2) < 0:
                p["variacao"] = "baixou"
            else:
                p["variacao"] = "sem alteração"
            postos.append(p)
        cidades[key] = {"label": titulo(muni), "uf": uf, "postos": postos}
    cad = cadastro()
    cad_out = {}
    for key, postos in sorted(cad.items()):
        muni, uf = key.split("|")
        cad_out[key] = {"label": titulo(muni), "uf": uf,
                        "fonte": "Cadastro oficial de postos da ANP",
                        "postos": postos}
    def pt(d):
        return d[8:10] + "/" + d[5:7] + "/" + d[0:4]
    now = datetime.datetime.now(datetime.timezone.utc)
    data = {
        "combustivel": "Gasolina comum",
        "periodo_inicio": a,
        "periodo_fim": b,
        "semana_label": pt(a) + " a " + pt(b),
        "atualizado_label": now.strftime("%d/%m/%Y"),
        "fonte": "ANP - Levantamento de Preços de Combustíveis (semanal)",
        "fonte_url": PAGE,
        "arquivo": cur_url,
        "gerado_em": now.isoformat(timespec="seconds"),
        "cidades": cidades,
        "cadastro": cad_out,
    }
    new = json.dumps(data, ensure_ascii=False, separators=(",", ":")) + "\n"
    try:
        old_data = json.load(open("cidades.json", encoding="utf-8"))
        old_data.pop("gerado_em", None); old_data.pop("atualizado_label", None)
        cmp = dict(data); cmp.pop("gerado_em"); cmp.pop("atualizado_label")
        if old_data == cmp:
            print("sem mudanca"); return
    except Exception:
        pass
    open("cidades.json", "w", encoding="utf-8").write(new)
    print("atualizado", a, b, len(cidades), "cidades +", len(cad_out), "via cadastro")

main()
