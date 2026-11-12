import pandas as pd
import re
import unicodedata
from openpyxl import load_workbook

ABAS = ("Globais", "Cenarios", "Equipamentos")
OBRIGATORIOS = {
    "marr": "taxa mínima de atratividade (MARR)",
    "vida_projeto": "vida útil do projeto",
}


def _valor(v):
    """Converte células vazias (NaN) em texto vazio."""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return v


def ler_planilha_tea(arquivo, modo_evte="formulas"):
    """
    Lê a planilha-modelo (somente valores digitados) e devolve (globais, cenarios).

    Abas:
      Globais      : chave | variavel | unidade | valor
      Cenarios     : chave | variavel | unidade | <um nome de cenário por coluna>
      Equipamentos : cenario | equipamento | quantidade | custo_unidade_usd | capacidade |
                     unidade_capacidade | consumo_kwh_ano | material | fonte | ajustar_cepci
    """
    xls = pd.ExcelFile(arquivo)

    # A planilha EVTE original usa abas/células diferentes do modelo do site.
    # Identifique-a antes de validar as abas do modelo simplificado.
    if {"Tecno Assumptions", "Custo dos Equipamentos", "Capital Investment"}.issubset(xls.sheet_names):
        arquivo.seek(0)
        return ler_planilha_evte(arquivo, modo_evte)

    faltando = [a for a in ABAS if a not in xls.sheet_names]
    if faltando:
        raise ValueError("Abas não encontradas na planilha: " + ", ".join(faltando))

    # ---- globais ----
    globais = {}
    for _, linha in xls.parse("Globais").iterrows():
        chave = _valor(linha.get("chave"))
        if chave:
            globais[str(chave).strip()] = _valor(linha.get("valor"))

    ausentes = [rotulo for chave, rotulo in OBRIGATORIOS.items() if str(globais.get(chave, "")).strip() == ""]
    if ausentes:
        raise ValueError("Faltam valores na aba Globais: " + ", ".join(ausentes))

    # ---- cenários (uma coluna por cenário) ----
    df_c = xls.parse("Cenarios")
    colunas = [c for c in df_c.columns
               if c not in ("chave", "variavel", "unidade") and not str(c).startswith("Unnamed")]
    if not colunas:
        raise ValueError("Nenhum cenário encontrado na aba Cenarios.")

    cenarios, por_nome = [], {}
    for col in colunas:
        cenario = {"nome": str(col).strip(), "equipamentos": []}
        for _, linha in df_c.iterrows():
            chave = _valor(linha.get("chave"))
            if chave:
                cenario[str(chave).strip()] = _valor(linha[col])
        cenarios.append(cenario)
        por_nome[cenario["nome"]] = cenario

    # ---- equipamentos ----
    for _, linha in xls.parse("Equipamentos").iterrows():
        nome_eq = _valor(linha.get("equipamento"))
        nome_cen = str(_valor(linha.get("cenario"))).strip()
        if not nome_eq and not nome_cen:
            continue
        if nome_cen == "" and len(cenarios) == 1:
            destino = cenarios[0]
        elif nome_cen in por_nome:
            destino = por_nome[nome_cen]
        else:
            raise ValueError(f"A aba Equipamentos cita o cenário '{nome_cen}', que não existe na aba Cenarios.")
        destino["equipamentos"].append({
            "nome": nome_eq,
            "quantidade": _valor(linha.get("quantidade")),
            "custo_unidade": _valor(linha.get("custo_unidade_usd")),
            "capacidade": _valor(linha.get("capacidade")),
            "unidade_capacidade": _valor(linha.get("unidade_capacidade")),
            "consumo_kwh_ano": _valor(linha.get("consumo_kwh_ano")),
            "material": _valor(linha.get("material")),
            "fonte": _valor(linha.get("fonte")),
            "ajustar_cepci": _valor(linha.get("ajustar_cepci")),
        })

    return globais, cenarios


def _num_planilha(valor, padrao=0.0):
    """Extrai números de valores numéricos ou textos como '8h/dia' e '60% do Buildings'."""
    valor = _valor(valor)
    if valor == "":
        return padrao
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace(" ", "")
    # Aceita números no padrão brasileiro e no padrão internacional.
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    else:
        texto = texto.replace(",", ".")
    encontrado = re.search(r"[-+]?\d+(?:\.\d+)?", texto)
    return float(encontrado.group()) if encontrado else padrao


# ---------------------------------------------------------------------------
# Importador do formato "Modelo Dados Processo EVTE"
#
# Em vez de ler endereços fixos (D6, F8...), procura pelos RÓTULOS das linhas, então continua
# funcionando se linhas forem inseridas/movidas. Lê também o que a planilha EFETIVAMENTE aplicou
# nas fórmulas (ex.: mão de obra a US$ 10/h) e compara com a tabela de premissas (US$ 25/h): a
# planilha original tem várias dessas divergências e elas mudam o veredito de viabilidade.
# ---------------------------------------------------------------------------

def _norm(texto):
    texto = unicodedata.normalize("NFKD", str(texto if texto is not None else ""))
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", " ", texto.lower()).strip()


def _grade(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _cel(grade, r, c):
    if r is None or c is None or r < 0 or c < 0 or r >= len(grade) or c >= len(grade[r]):
        return None
    return grade[r][c]


def _eh_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _acha(grade, rotulo, col, inicio=0, fim=None, exato=False):
    """Índice da 1ª linha (entre inicio e fim) cujo texto na coluna 'col' casa com o rótulo."""
    alvo = _norm(rotulo)
    fim = len(grade) if fim is None else fim
    for r in range(max(inicio, 0), min(fim, len(grade))):
        t = _norm(_cel(grade, r, col))
        if t and (t == alvo if exato else t.startswith(alvo)):
            return r
    return None


def _val(grade, rotulo, col_rotulo, col_valor, inicio=0, fim=None, exato=False):
    r = _acha(grade, rotulo, col_rotulo, inicio, fim, exato)
    v = _cel(grade, r, col_valor) if r is not None else None
    return _num_planilha(v, None) if v not in (None, "") else None


def _pct(valor):
    """Aceita 0,34 ou 34 e devolve em %."""
    return valor * 100 if valor is not None and 0 < valor <= 1 else valor


def _razao(a, b):
    return a / b if (a is not None and b not in (None, 0)) else None


def _tecno(sheet, nomes_cenarios=None):
    """Lê a aba Tecno Assumptions (premissas digitadas) por rótulo."""
    G = _grade(sheet)
    c_lab = 1
    r_var = _acha(G, "variable", c_lab) or 0
    r_unit = _acha(G, "unit cost", c_lab) or len(G)
    r_fin = next((r for r in range(len(G)) if _norm(_cel(G, r, c_lab)).startswith(("finacial criteria", "financial criteria"))), None)
    r_custo = next((r for r in range(len(G)) if _norm(_cel(G, r, c_lab)).startswith("cost criteri")), None)

    # colunas de cenário = colunas (a partir da D) com capacidade processada numérica
    r_cap = _acha(G, "processed capacity", c_lab)
    colunas = [c for c in range(3, max(len(l) for l in G)) if _eh_num(_cel(G, r_cap, c)) and _cel(G, r_cap, c) > 0]
    nomes = {}
    for c in colunas:
        partes = [str(_cel(G, r, c)).strip() for r in range(0, r_var)
                  if isinstance(_cel(G, r, c), str) and _norm(_cel(G, r, c)) not in ("", "cenarios")]
        nomes[c] = " ".join(partes)

    def qtd(rotulo, c, exato=True):      # quantidades: antes do bloco 'Unit cost'
        return _val(G, rotulo, c_lab, c, 0, r_unit, exato)

    def preco(rotulo, c=3, exato=False):  # custos unitários: depois de 'Unit cost'
        return _val(G, rotulo, c_lab, c, r_unit, r_fin, exato)

    def crit(rotulo, c=3, exato=False):
        return _val(G, rotulo, c_lab, c, r_custo or 0, None, exato)

    def fin(rotulo, c=3, exato=False):
        return _val(G, rotulo, c_lab, c, r_fin or 0, r_custo, exato)

    escalares = {
        "custo_lupulo": preco("lupulo cost"),
        "etanol_preco": preco("etanol", exato=True), "co2_preco": preco("co2", exato=True),
        "energia_eletrica": preco("electric energy"), "residuos": preco("waste disposal"),
        "mao_obra_hora": preco("operating labor cost"), "agua": preco("water", exato=True),
        "vida_equipamentos": fin("equipment", exato=True), "vida_edificacoes": fin("buildings", exato=True),
        "imposto": _pct(fin("local tax")), "marr": fin("minimum acceptable"), "wacc": fin("weighted average"),
        "valor_residual_pct": fin("salvage value"), "vida_projeto": fin("life time project"),
        "manutencao": crit("maintenance and repairs"), "suprimentos": crit("operating supplies"),
        "laboratorio": crit("laboratory charges"), "patentes": crit("patents and royalties"),
        "contingencia": crit("contingency"), "impostos_locais": crit("local taxes"),
        "seguro": crit("insurance"), "embalagem": crit("packaging"),
        "mao_obra_indireta": crit("indirect labor"), "armazenagem": crit("warehouse and storage"),
        "administrativas": crit("administrative expenses"), "vendas": crit("selling and distribution"),
        "pd": crit("research and development"),
    }
    cenarios = []
    for c in colunas:
        cenarios.append({
            "col": c,
            "nome": nomes[c] or f"Cenário {len(cenarios) + 1}",
            "capacidade": qtd("processed capacity", c, False),
            "horas_dia": qtd("shift length", c, False), "dias_ano": qtd("production schedule", c, False),
            "operadores": qtd("labor", c, True),
            "produto1": qtd("product 1", c, False), "produto2": qtd("product 2", c, False),
            "etanol": qtd("etanol", c), "co2": qtd("co2", c),
            "agua": qtd("water", c), "residuos": qtd("waste disposal", c),
            "preco_produto1_t": preco("preco produto 1", c), "preco_produto2_t": preco("preco produto 2", c),
        })
    return escalares, cenarios


def _equipamentos(eq_ws, n_cenarios):
    """Lê CEPCI e as listas de equipamentos (um bloco por cenário, de 'via ...' até 'Total')."""
    G = _grade(eq_ws)
    # CEPCI: cabeçalho "Índice CEPCI"; abaixo, ano-base e ano atual
    cepci = {}
    for r, linha in enumerate(G):
        for c, v in enumerate(linha):
            if isinstance(v, str) and _norm(v).startswith("indice cepci"):
                cepci = {"cepci_ano_base": _cel(G, r + 1, c - 1), "cepci_base": _cel(G, r + 1, c),
                         "cepci_ano_atual": _cel(G, r + 2, c - 1), "cepci_atual": _cel(G, r + 2, c)}
                break
        if cepci:
            break
    razao = _razao(_num_planilha(cepci.get("cepci_atual"), None), _num_planilha(cepci.get("cepci_base"), None)) or 1.0

    r_cab = _acha(G, "equipamento", 1, exato=True)
    cab = {}
    if r_cab is not None:
        for c, v in enumerate(G[r_cab]):
            t = _norm(v)
            for chave, rot in (("qtd", "quantidade"), ("custo", "custo unidade"), ("ajust", "valor ajustado"),
                               ("cap", "capacidades"), ("unid", "especificacao"), ("kwh", "consumo eletrico"),
                               ("mat", "material"), ("fonte", "fonte")):
                if t.startswith(rot) and chave not in cab:
                    cab[chave] = c
    blocos, atual = [], None
    for r in range((r_cab or 0) + 1, len(G)):
        a = _norm(_cel(G, r, 0))
        if a == "total":
            if atual is not None:
                blocos.append(atual)
            atual = None
            continue
        if a and atual is None:
            atual = []
        if atual is None:
            continue
        conteudo = [_cel(G, r, c) for c in range(1, 12)]
        if any(v not in (None, "") for v in conteudo):
            atual.append(r)
    if atual:
        blocos.append(atual)

    resultado = []
    for linhas in blocos[:n_cenarios]:
        itens = []
        for r in linhas:
            nome = _cel(G, r, 1)
            custo = _num_planilha(_cel(G, r, cab.get("custo", 3)))
            ajust = _num_planilha(_cel(G, r, cab.get("ajust", 4)))
            ajustado = custo > 0 and razao != 1 and abs(ajust - custo * razao) < 1e-6 * max(1, abs(ajust))
            c_fonte = cab.get("fonte", 10)
            itens.append({
                "nome": str(nome).strip() if nome else f"Equipamento (linha {r + 1})",
                "quantidade": _num_planilha(_cel(G, r, cab.get("qtd", 2))),
                "custo_unidade": custo,
                "capacidade": _num_planilha(_cel(G, r, cab.get("cap", 6))),
                "unidade_capacidade": str(_cel(G, r, cab.get("unid", 7)) or ""),
                "consumo_kwh_ano": _num_planilha(_cel(G, r, cab.get("kwh", 8))),
                "material": str(_cel(G, r, cab.get("mat", 9)) or ""),
                "fonte": str(_cel(G, r, c_fonte) or _cel(G, r, c_fonte + 1) or ""),
                "ajustar_cepci": "sim" if ajustado else "não",
            })
        resultado.append(itens)
    return cepci, resultado


def _capital(cap_ws):
    """Um bloco por cenário: cada bloco começa numa coluna que tem o rótulo 'Total Capital Investment'."""
    G = _grade(cap_ws)
    ancoras = sorted({c for linha in G for c, v in enumerate(linha) if _norm(v) == "total capital investment"})
    blocos = []
    for c in ancoras:
        def v(rotulo, exato=True):
            return _val(G, rotulo, c, c + 1, exato=exato)
        blocos.append({
            "fci": v("total fixed capital investment"), "fci_direto": v("direct fixed capital investment"),
            "tci": v("total capital investment"), "edificacoes": v("buildings including services", False),
        })
    return blocos


def _custos_aplicados(pc_ws, n):
    """Valores que a planilha realmente calculou em 'Production Cost and Sales' (uma coluna por cenário)."""
    G = _grade(pc_ws)
    r_cab = _acha(G, "costs item", 0)
    r_fim = next((r for r in range(len(G)) if _norm(_cel(G, r, 0)).startswith("cost criteri")), len(G))
    colunas = [c for c, v in enumerate(G[r_cab]) if c > 0 and _norm(v) not in ("", "%")][:n] if r_cab is not None else []
    out = []
    for c in colunas:
        d = {}
        for chave, rot in (("mao_obra", "operating labor"), ("manutencao", "maintenance and repairs"),
                           ("suprimentos", "operating supplies"), ("laboratorio", "laboratory"),
                           ("patentes", "patents and royalties"), ("impostos_locais", "local taxes"),
                           ("seguro", "insurance"), ("embalagem", "packaging"),
                           ("mao_obra_indireta", "indirect labor"), ("armazenagem", "warehouse"),
                           ("administrativas", "administrative expenses"), ("vendas", "selling and distribution"),
                           ("pd", "research and development"), ("agua", "water"),
                           ("manufatura", "manufacturing costs"), ("contingencia", "contingency")):
            d[chave] = _val(G, rot, 0, c, 0, r_fim)
        out.append(d)
    return out


def _financeiro(fin_ws, vendas_ws, n):
    """Resultados que o Excel calculou (VPL, TIR, payback), preço efetivo, imposto aplicado e cronograma."""
    F, S = _grade(fin_ws), _grade(vendas_ws)
    res = {"excel": [], "preco1": [], "preco2": [], "receita": [], "imposto_aplicado": [], "cronograma": None}
    for i in range(n):
        c = 2 + i
        ebit = _val(F, "ebit", 0, c, exato=True)
        imp = _val(F, "income tax", 0, c)
        res["imposto_aplicado"].append(_razao(imp, ebit) * 100 if _razao(imp, ebit) else None)
        res["excel"].append({"vpl": _val(F, "vpl", 0, c, exato=True), "tir": _val(F, "tir", 0, c, exato=True),
                             "payback": _val(F, "payback", 0, c, exato=True)})
        res["preco1"].append(_val(S, "preco produto 1", 0, c))
        res["preco2"].append(_val(S, "preco produto 2", 0, c))
        res["receita"].append(_val(S, "total revenue", 0, c))
    # cronograma de desembolso: dois números consecutivos na coluna D, entre 0 e 1, somando 1 (ex.: 0,7 e 0,3)
    for r in range(len(F) - 1):
        a, b = _cel(F, r, 3), _cel(F, r + 1, 3)
        if _eh_num(a) and _eh_num(b) and 0 < a < 1 and 0 < b < 1 and abs(a + b - 1) < 1e-9:
            res["cronograma"] = a * 100
            break
    return res


def _salvar_param(divs, nome, tabela, aplicado, unidade, tol=0.02):
    """Registra divergência entre a tabela de premissas e o que as fórmulas aplicam."""
    if tabela is None or aplicado is None:
        return False
    if abs(aplicado - tabela) <= tol * max(abs(tabela), abs(aplicado), 1e-12):
        return False
    divs.append({"parametro": nome, "premissa": tabela, "aplicado": aplicado, "unidade": unidade})
    return True


def ler_planilha_evte(arquivo, modo="formulas"):
    """
    Importa o formato 'Modelo Dados Processo EVTE'. Devolve (globais, cenarios).

    modo="formulas"  : usa o que as fórmulas da planilha EFETIVAMENTE aplicam (reproduz o VPL/TIR do Excel).
    modo="premissas" : usa a tabela 'Tecno Assumptions' ao pé da letra.
    Os dois cálculos ficam em globais["_conferencia"] para o site comparar.
    """
    wb = load_workbook(arquivo, data_only=True, read_only=True)
    escalares, cen_t = _tecno(wb["Tecno Assumptions"])
    n = len(cen_t)
    cepci, equips = _equipamentos(wb["Custo dos Equipamentos"], n)
    capitais = _capital(wb["Capital Investment"])
    aplicados = _custos_aplicados(wb["Production Cost and Sales"], n) if "Production Cost and Sales" in wb.sheetnames else []
    fin = _financeiro(wb["Financial Analisys"], wb["Sales"], n) if {"Financial Analisys", "Sales"} <= set(wb.sheetnames) else None

    base = {k: (v if v is not None else 0.0) for k, v in escalares.items()}
    base["etanol_desconto"] = 0  # os valores do arquivo original já incluem os descontos de compra
    base["co2_desconto"] = 0
    for k, v in cepci.items():
        base[k] = _num_planilha(v)
    if fin and fin["cronograma"]:
        base["invest_ano0_pct"] = fin["cronograma"]

    cenarios = []
    for i, t in enumerate(cen_t):
        eqs = equips[i] if i < len(equips) else []
        cap = capitais[i] if i < len(capitais) else {}
        cenarios.append({
            "nome": t["nome"], "capacidade": t["capacidade"] or 0.0, "horas_dia": t["horas_dia"] or 0.0,
            "dias_ano": t["dias_ano"] or 0.0, "operadores": t["operadores"] or 0.0,
            "produto1": t["produto1"] or 0.0, "produto2": t["produto2"] or 0.0,
            "preco_produto1": (t["preco_produto1_t"] or 0.0) / 1000, "preco_produto2": (t["preco_produto2_t"] or 0.0) / 1000,
            "etanol": t["etanol"] or 0.0, "co2": t["co2"] or 0.0, "agua": t["agua"] or 0.0,
            "residuos": t["residuos"] or 0.0,
            "fci": cap.get("fci") or 0.0, "fci_direto": cap.get("fci_direto") or 0.0,
            "investimento_inicial": cap.get("tci") or 0.0, "edificacoes": cap.get("edificacoes") or 0.0,
            "equipamentos": eqs,
        })

    ausentes = [rotulo for chave, rotulo in OBRIGATORIOS.items() if not base.get(chave)]
    if ausentes:
        raise ValueError("Faltam valores obrigatórios na planilha EVTE: " + ", ".join(ausentes))
    if not cenarios:
        raise ValueError("Nenhum cenário com 'Processed capacity' numérica foi encontrado em Tecno Assumptions.")

    # ---- o que as fórmulas realmente aplicam -------------------------------
    globais_formulas, cenarios_formulas = dict(base), [dict(c, equipamentos=c["equipamentos"]) for c in cenarios]
    divs = []
    if aplicados and fin:
        cen0 = cenarios[0]
        horas_mo = (cen0["operadores"] or 0) * (cen0["horas_dia"] or 0) * (cen0["dias_ano"] or 0)
        receita0 = fin["receita"][0]
        implicito = {
            "mao_obra_hora": ("Mão de obra operacional", "US$/h", _razao(aplicados[0]["mao_obra"], horas_mo)),
            "manutencao": ("Manutenção e reparos (% do FCI direto)", "%", _pct_de(_razao(aplicados[0]["manutencao"], cenarios[0]["fci_direto"]))),
            "suprimentos": ("Suprimentos operacionais (% da manutenção)", "%", _pct_de(_razao(aplicados[0]["suprimentos"], aplicados[0]["manutencao"]))),
            "laboratorio": ("Encargos de laboratório (% da mão de obra)", "%", _pct_de(_razao(aplicados[0]["laboratorio"], aplicados[0]["mao_obra"]))),
            "patentes": ("Patentes e royalties (% da receita)", "%", _pct_de(_razao(aplicados[0]["patentes"], receita0))),
            "impostos_locais": ("Impostos locais (% do FCI)", "%", _pct_de(_razao(aplicados[0]["impostos_locais"], cenarios[0]["fci"]))),
            "seguro": ("Seguro (% do FCI)", "%", _pct_de(_razao(aplicados[0]["seguro"], cenarios[0]["fci"]))),
            "vendas": ("Vendas e distribuição (% da receita)", "%", _pct_de(_razao(aplicados[0]["vendas"], receita0))),
            "pd": ("Pesquisa e desenvolvimento (% da receita)", "%", _pct_de(_razao(aplicados[0]["pd"], receita0))),
        }
        base_mo_man = (aplicados[0]["mao_obra"] or 0) + (aplicados[0]["manutencao"] or 0)
        implicito["embalagem"] = ("Embalagem (% de mão de obra + manutenção)", "%", _pct_de(_razao(aplicados[0]["embalagem"], base_mo_man)))
        implicito["mao_obra_indireta"] = ("Mão de obra indireta (% de mão de obra + manutenção)", "%", _pct_de(_razao(aplicados[0]["mao_obra_indireta"], base_mo_man)))
        implicito["armazenagem"] = ("Armazém e estocagem (% de mão de obra + manutenção)", "%", _pct_de(_razao(aplicados[0]["armazenagem"], base_mo_man)))
        implicito["administrativas"] = ("Despesas administrativas (% do custo de manufatura)", "%", _pct_de(_razao(aplicados[0]["administrativas"], aplicados[0]["manufatura"])))
        for chave, (rotulo, unidade, valor) in implicito.items():
            if valor is not None and _salvar_param(divs, rotulo, base.get(chave), valor, unidade):
                globais_formulas[chave] = valor
        # contingência: a tabela de premissas tem, mas não existe linha de contingência nos custos calculados
        if aplicados[0].get("contingencia") is None and base.get("contingencia"):
            divs.append({"parametro": "Contingência (% do FCI)", "premissa": base["contingencia"], "aplicado": 0.0, "unidade": "%"})
            globais_formulas["contingencia"] = 0.0
        # imposto sobre o lucro
        if fin["imposto_aplicado"] and fin["imposto_aplicado"][0] is not None and \
                _salvar_param(divs, "Imposto sobre o lucro", base.get("imposto"), fin["imposto_aplicado"][0], "%"):
            globais_formulas["imposto"] = fin["imposto_aplicado"][0]
        # a planilha calcula o valor residual mas não o inclui no VPL
        if base.get("valor_residual_pct"):
            divs.append({"parametro": "Valor residual (% das edificações)", "premissa": base["valor_residual_pct"], "aplicado": 0.0, "unidade": "% (não entra no VPL do Excel)"})
            globais_formulas["valor_residual_pct"] = 0.0
        for i, c in enumerate(cenarios_formulas):
            # preço efetivamente usado em 'Sales' (a planilha aplica 10% de desconto que não está nas premissas)
            for k, ref, rot in (("preco_produto1", fin["preco1"][i], "Preço do produto 1"), ("preco_produto2", fin["preco2"][i], "Preço do produto 2")):
                if ref is not None and _salvar_param(divs, f"{rot} — {c['nome']}", c[k] * 1000, ref, "US$/t"):
                    c[k] = ref / 1000
            # água: a planilha usa a quantidade do 1º cenário nos dois
            if i < len(aplicados) and aplicados[i]["agua"] is not None and base.get("agua"):
                qtd = aplicados[i]["agua"] / (1000 * base["agua"])
                if _salvar_param(divs, f"Água — {c['nome']}", c["agua"], qtd, "t/ano"):
                    c["agua"] = qtd
            # a planilha deprecia só os equipamentos comprados (e as edificações)
            rz = _razao(base.get("cepci_atual"), base.get("cepci_base")) or 1.0
            c["base_equip_depreciavel"] = sum(
                e["quantidade"] * e["custo_unidade"] * (rz if e["ajustar_cepci"] == "sim" else 1.0)
                for e in c["equipamentos"])
    # ---- modo escolhido ------------------------------------------------------
    conferencia = {
        "excel": fin["excel"] if fin else [],
        "divergencias": _agrupar_divs(divs),
        "modo": modo,
        "formulas": (globais_formulas, cenarios_formulas),
        "premissas": (dict(base), [dict(c) for c in cenarios]),
    }
    escolhido_g, escolhido_c = conferencia["formulas" if modo == "formulas" else "premissas"]
    escolhido_g = dict(escolhido_g)
    escolhido_g["_conferencia"] = conferencia
    return escolhido_g, escolhido_c


def _pct_de(razao):
    return razao * 100 if razao is not None else None


def _agrupar_divs(divs):
    """Remove duplicadas (mesmo parâmetro/valores) mantendo a ordem."""
    vistos, out = set(), []
    for d in divs:
        chave = (d["parametro"], round(d["premissa"], 6), round(d["aplicado"], 6))
        if chave not in vistos:
            vistos.add(chave)
            out.append(d)
    return out
