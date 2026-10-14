"""Baixa a planilha semanal da ANP e gera postos.json (gasolina comum, Apucarana-PR).
Dados publicos da ANP. Sem segredos."""
import io, json, re, sys, datetime, urllib.request
import openpyxl

PAGE = "https://www.gov.br/anp/pt-br/assuntos/precos-e-defesa-da-concorrencia/precos/levantamento-de-precos-de-combustiveis-ultimas-semanas-pesquisadas"
CIDADE = "APUCARANA"
UF = "PARANA"
UA = {"User-Agent": "Mozilla/5.0"}
# Postos sempre incluidos (mesmo sem coleta na semana atual): usa o ultimo preco ANP disponivel.
EXTRAS = ["10601388000130"]  # BRZ SUL / Posto Brasil Sul II, Av. Parana 196
NOMES = {"10601388000130": "Posto Brasil Sul II (Brz Sul Comercio De Combustiveis Ltda)"}
MAX_SEMANAS = 8

def get(url):
    return urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=120).read()

def links():
    html = get(PAGE).decode("utf-8", "ignore")
    found = re.findall(r'href="([^"]*revendas_lpc_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})\.xlsx)"', html)
    uniq = {}
    for u, a, b in found:
        uniq[a] = (u, a, b)
    return [uniq[k] for k in sorted(uniq, reverse=True)]

def rows(url):
    wb = openpyxl.load_workbook(io.BytesIO(get(url)), read_only=True)
    out = {}
    for r in wb.active.iter_rows(min_row=11, values_only=True):
        if not r or r[0] is None:
            continue
        if str(r[8]).strip().upper() != CIDADE or str(r[9]).strip().upper() != UF:
            continue
        if str(r[11]).strip().upper() != "GASOLINA COMUM":
            continue
        cnpj = str(int(r[0])).zfill(14)
        d = r[14]
        out[cnpj] = {
            "cnpj": cnpj,
            "nome": (r[2] or r[1] or "").strip().title(),
            "razao": (r[1] or "").strip().title(),
            "bandeira": (r[10] or "").strip().title(),
            "endereco": f"{str(r[3] or '').strip().title()}, {r[4]}".strip(", "),
            "bairro": (r[6] or "").strip().title(),
            "preco": float(r[13]),
            "coleta": d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d),
        }
    return out

def main():
    ls = links()
    if not ls:
        sys.exit("nenhum link de planilha encontrado")
    cur_url, a, b = ls[0]
    cur = rows(cur_url)
    if not cur:
        sys.exit("nenhum posto de Apucarana na planilha")
    prev = rows(ls[1][0]) if len(ls) > 1 else {}
    cache = {0: cur}
    if len(ls) > 1:
        cache[1] = prev
    def week(i):
        if i not in cache:
            try:
                cache[i] = rows(ls[i][0])
            except Exception:
                cache[i] = {}
        return cache[i]
    for cnpj in EXTRAS:
        if cnpj in cur:
            continue
        for i in range(1, min(len(ls), MAX_SEMANAS)):
            if cnpj in week(i):
                p = dict(week(i)[cnpj])
                p["desatualizado"] = True
                p["nome"] = NOMES.get(cnpj, p["nome"])
                prev_price = None
                for j in range(i + 1, min(len(ls), MAX_SEMANAS + 1)):
                    if cnpj in week(j):
                        prev_price = week(j)[cnpj]["preco"]
                        break
                p["_prev"] = prev_price
                cur[cnpj] = p
                break
    postos = []
    for cnpj, p in sorted(cur.items(), key=lambda kv: kv[1]["preco"]):
        if "_prev" in p:
            old = p.pop("_prev")
        else:
            old = prev.get(cnpj, {}).get("preco")
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
    data = {
        "combustivel": "Gasolina comum",
        "cidade": "Apucarana-PR",
        "periodo_inicio": a,
        "periodo_fim": b,
        "fonte": "ANP - Levantamento de Preços de Combustíveis (semanal)",
        "fonte_url": PAGE,
        "arquivo": cur_url,
        "gerado_em": datetime.datetime.now(datetime.timezone.utc).isoformat(timespec="seconds"),
        "postos": postos,
    }
    new = json.dumps(data, ensure_ascii=False, indent=2) + "\n"
    try:
        old_data = json.load(open("postos.json", encoding="utf-8"))
        old_data.pop("gerado_em", None)
        cmp = dict(data); cmp.pop("gerado_em")
        if old_data == cmp:
            print("sem mudanca"); return
    except Exception:
        pass
    open("postos.json", "w", encoding="utf-8").write(new)
    print("atualizado", a, b, len(postos), "postos")

main()
