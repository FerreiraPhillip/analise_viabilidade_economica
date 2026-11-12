"""
Análise tecno-econômica por cenário.

Parte 1 — o que a planilha Teste.xlsx já calculava (reproduzido aqui):
  * razão CEPCI                      = índice atual / índice do ano-base
  * valor ajustado do equipamento    = custo unitário (× razão CEPCI, se for custo de ano antigo)
  * custo total do equipamento       = valor ajustado × quantidade
  * custo total e consumo elétrico do cenário (somas)
  * rendimento do processo           = produto / matéria-prima
  * preços finais digitados como fórmula (etanol/CO₂ com desconto; US$/kg → US$/t)

Parte 2 — o que o site acrescenta para dar o veredito de viabilidade:
  receita, custos anuais, depreciação, imposto, fluxo de caixa e indicadores
  (VPL, TIR, payback), reutilizando as funções de analysis.py.
"""

from analysis import (numero, analisar_viabilidade_investimento, valor_presente_liquido,
                      aplicar_cronograma_investimento)


def fmt_num(valor, casas=2):
    """Formata no padrão brasileiro: 1.234,56"""
    if valor is None:
        return "—"
    texto = f"{valor:,.{casas}f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def razao_cepci(indice_base, indice_atual):
    base = numero(indice_base)
    return numero(indice_atual) / base if base else 1.0


def _sim(valor):
    return str(valor).strip().lower() in ("sim", "s", "true", "1", "x", "yes")


def calcular_equipamentos(equipamentos, razao):
    itens = []
    custo_total = 0.0
    consumo_total = 0.0
    sem_custo = []

    for eq in equipamentos:
        nome = (eq.get("nome") or "Equipamento").strip() if isinstance(eq.get("nome"), str) else "Equipamento"
        quantidade = numero(eq.get("quantidade"))
        custo_unidade = numero(eq.get("custo_unidade"))
        ajustar = _sim(eq.get("ajustar_cepci"))
        valor_ajustado = custo_unidade * (razao if ajustar else 1.0)
        total = valor_ajustado * quantidade
        consumo = numero(eq.get("consumo_kwh_ano"))

        if quantidade <= 0 or custo_unidade <= 0:
            sem_custo.append(nome)

        itens.append({
            "nome": nome,
            "quantidade": quantidade,
            "custo_unidade": custo_unidade,
            "ajustado_cepci": ajustar,
            "valor_ajustado": valor_ajustado,
            "total": total,
            "capacidade": eq.get("capacidade"),
            "unidade_capacidade": eq.get("unidade_capacidade") or "",
            "consumo_kwh_ano": consumo,
            "material": eq.get("material") or "",
            "fonte": eq.get("fonte") or "",
        })
        custo_total += total
        consumo_total += consumo

    return {
        "itens": itens,
        "custo_total": custo_total,
        "consumo_kwh_ano": consumo_total,
        "sem_custo": sem_custo,
    }


def _pct(globais, chave):
    return numero(globais.get(chave)) / 100.0


# Chaves dos fatores de investimento (% do custo dos equipamentos comprados).
FATORES_DIRETOS = ("capex_instalacao", "capex_instrumentacao", "capex_tubulacao", "capex_eletrica",
                   "capex_terreno", "capex_escritorio", "capex_edificacoes")
FATORES_INDIRETOS = ("capex_engenharia", "capex_construcao", "capex_juridico",
                     "capex_empreiteiro", "capex_contingencia")


def calcular_capex_por_fatores(g, custo_equipamentos):
    """
    Estima o investimento a partir do custo dos equipamentos comprados, como a planilha EVTE faz
    na aba 'Capital Investment': cada item é um percentual do custo dos equipamentos.
      FCI direto   = equipamentos × (1 + Σ fatores diretos)
      FCI indireto = equipamentos × Σ fatores indiretos
      FCI          = direto + indireto
      TCI          = FCI + capital de giro (% do FCI) + partida (% do FCI)
    Retorna None se nenhum fator foi informado (aí vale o comportamento antigo).
    """
    chaves = FATORES_DIRETOS + FATORES_INDIRETOS + ("capex_capital_giro", "capex_partida")
    if not any(str(g.get(k, "")).strip() != "" for k in chaves):
        return None

    direto = custo_equipamentos * (1 + sum(_pct(g, k) for k in FATORES_DIRETOS))
    indireto = custo_equipamentos * sum(_pct(g, k) for k in FATORES_INDIRETOS)
    fci = direto + indireto
    capital_giro = fci * _pct(g, "capex_capital_giro")
    partida = fci * _pct(g, "capex_partida")
    return {
        "fci": fci,
        "fci_direto": direto,
        "edificacoes": custo_equipamentos * _pct(g, "capex_edificacoes"),
        "capital_giro": capital_giro,
        "partida": partida,
        "tci": fci + capital_giro + partida,
    }


def calcular_cenario(g, c, razao):
    """Calcula um cenário completo. g = parâmetros globais; c = dados do cenário."""
    capacidade = numero(c.get("capacidade"))
    produto1 = numero(c.get("produto1"))
    produto2 = numero(c.get("produto2"))
    horas = numero(c.get("horas_dia"))
    dias = numero(c.get("dias_ano"))
    operadores = numero(c.get("operadores"))

    # ---- o que o Excel calculava -----------------------------------------
    equip = calcular_equipamentos(c.get("equipamentos", []), razao)
    rendimento1 = produto1 / capacidade if capacidade else 0.0
    rendimento2 = produto2 / capacidade if capacidade else 0.0
    energia_kwh = equip["consumo_kwh_ano"]

    preco_etanol = numero(g.get("etanol_preco")) * (1 - _pct(g, "etanol_desconto"))
    preco_co2 = numero(g.get("co2_preco")) * (1 - _pct(g, "co2_desconto"))
    preco_p1_t = numero(c.get("preco_produto1")) * 1000  # US$/kg -> US$/t
    preco_p2_t = numero(c.get("preco_produto2")) * 1000

    # ---- receita ---------------------------------------------------------
    receita = produto1 * preco_p1_t + produto2 * preco_p2_t

    # ---- investimento ----------------------------------------------------
    fci_informado = numero(c.get("fci"))
    investimento_informado = numero(c.get("investimento_inicial"))
    edificacoes = numero(c.get("edificacoes"))
    fci_direto_informado = numero(c.get("fci_direto"))

    capex = None
    if fci_informado <= 0 and equip["custo_total"] > 0:
        capex = calcular_capex_por_fatores(g, equip["custo_total"])

    if fci_informado > 0:
        fci = fci_informado
        fci_direto = fci_direto_informado if fci_direto_informado > 0 else fci
    elif capex:
        fci, fci_direto = capex["fci"], capex["fci_direto"]
        if edificacoes <= 0:
            edificacoes = capex["edificacoes"]
    else:
        fci = equip["custo_total"]
        fci_direto = fci

    if investimento_informado > 0:
        investimento_inicial = investimento_informado
    elif capex:
        investimento_inicial = capex["tci"]
    else:
        investimento_inicial = fci

    # ---- custos anuais (cada linha traz a fórmula usada) -----------------
    custos = []

    def add(nome, valor, formula):
        custos.append({"nome": nome, "valor": valor, "formula": formula})
        return valor

    mp = add("Matéria-prima (lúpulo)", capacidade * numero(g.get("custo_lupulo")),
             f"{fmt_num(capacidade)} t/ano × US$ {fmt_num(numero(g.get('custo_lupulo')))}/t")
    etanol = add("Etanol", numero(c.get("etanol")) * preco_etanol,
                 f"{fmt_num(numero(c.get('etanol')), 3)} t/ano × US$ {fmt_num(preco_etanol)}/t (preço de lista com desconto)")
    co2 = add("CO₂", numero(c.get("co2")) * preco_co2,
              f"{fmt_num(numero(c.get('co2')), 3)} t/ano × US$ {fmt_num(preco_co2)}/t (preço de lista com desconto)")
    eletr = add("Energia elétrica", energia_kwh * numero(g.get("energia_eletrica")),
                f"{fmt_num(energia_kwh)} kWh/ano (soma dos equipamentos) × US$ {fmt_num(numero(g.get('energia_eletrica')), 3)}/kWh")
    resid = add("Descarte de resíduos", numero(c.get("residuos")) * numero(g.get("residuos")),
                f"{fmt_num(numero(c.get('residuos')), 3)} t/ano × US$ {fmt_num(numero(g.get('residuos')))}/t")
    agua = add("Água", numero(c.get("agua")) * 1000 * numero(g.get("agua")),
               f"{fmt_num(numero(c.get('agua')), 2)} t/ano × 1.000 kg/t × US$ {fmt_num(numero(g.get('agua')), 3)}/kg")
    mao_obra = add("Mão de obra operacional", operadores * horas * dias * numero(g.get("mao_obra_hora")),
                   f"{fmt_num(operadores, 0)} operadores × {fmt_num(horas, 0)} h/dia × {fmt_num(dias, 0)} dias × US$ {fmt_num(numero(g.get('mao_obra_hora')))}/h")

    manutencao = add("Manutenção e reparos", fci_direto * _pct(g, "manutencao"),
                     f"{fmt_num(numero(g.get('manutencao')))}% × FCI direto (US$ {fmt_num(fci_direto)})")
    add("Suprimentos operacionais", manutencao * _pct(g, "suprimentos"),
        f"{fmt_num(numero(g.get('suprimentos')))}% × manutenção")
    add("Encargos de laboratório", mao_obra * _pct(g, "laboratorio"),
        f"{fmt_num(numero(g.get('laboratorio')))}% × mão de obra operacional")
    add("Patentes e royalties", receita * _pct(g, "patentes"),
        f"{fmt_num(numero(g.get('patentes')))}% × receita")
    add("Contingência", fci * _pct(g, "contingencia"),
        f"{fmt_num(numero(g.get('contingencia')))}% × FCI")
    add("Impostos locais", fci * _pct(g, "impostos_locais"),
        f"{fmt_num(numero(g.get('impostos_locais')))}% × FCI")
    add("Seguro", fci * _pct(g, "seguro"),
        f"{fmt_num(numero(g.get('seguro')))}% × FCI")

    base_mo_manut = mao_obra + manutencao
    add("Embalagem", base_mo_manut * _pct(g, "embalagem"),
        f"{fmt_num(numero(g.get('embalagem')))}% × (mão de obra operacional + manutenção)")
    add("Mão de obra indireta", base_mo_manut * _pct(g, "mao_obra_indireta"),
        f"{fmt_num(numero(g.get('mao_obra_indireta')))}% × (mão de obra operacional + manutenção)")
    add("Armazém e estocagem", base_mo_manut * _pct(g, "armazenagem"),
        f"{fmt_num(numero(g.get('armazenagem')))}% × (mão de obra operacional + manutenção)")

    vida_eq = numero(g.get("vida_equipamentos"))
    vida_ed = numero(g.get("vida_edificacoes"))
    base_informada = numero(c.get("base_equip_depreciavel"))
    base_equip = base_informada if base_informada > 0 else max(fci - edificacoes, 0.0)
    depreciacao = (base_equip / vida_eq if vida_eq else 0.0) + (edificacoes / vida_ed if vida_ed else 0.0)
    rotulo_base = "equipamentos comprados" if base_informada > 0 else "(FCI − edificações)"
    add("Depreciação (linear)", depreciacao,
        f"{rotulo_base} ÷ {fmt_num(vida_eq, 0)} anos + edificações ÷ {fmt_num(vida_ed, 0)} anos")

    custo_manufatura = sum(item["valor"] for item in custos)
    add("Despesas administrativas", custo_manufatura * _pct(g, "administrativas"),
        f"{fmt_num(numero(g.get('administrativas')))}% × custo de manufatura (US$ {fmt_num(custo_manufatura)})")
    add("Vendas e distribuição", receita * _pct(g, "vendas"),
        f"{fmt_num(numero(g.get('vendas')))}% × receita")
    add("Pesquisa e desenvolvimento", receita * _pct(g, "pd"),
        f"{fmt_num(numero(g.get('pd')))}% × receita")

    custo_total_anual = sum(item["valor"] for item in custos)

    # ---- resultado e fluxo de caixa -------------------------------------
    lucro_antes_imposto = receita - custo_total_anual
    imposto = max(lucro_antes_imposto, 0.0) * _pct(g, "imposto")
    lucro_liquido = lucro_antes_imposto - imposto
    fluxo_anual = lucro_liquido + depreciacao

    anos = int(round(numero(g.get("vida_projeto"))))
    fluxos = [fluxo_anual] * max(anos, 0)
    valor_residual = edificacoes * numero(g.get("valor_residual_pct")) / 100.0

    # Parcela do investimento paga no ano 0; o restante vai para o ano 1, antes da operação
    # (a planilha EVTE usa 70% / 30%). Em branco ou 100 = tudo no ano 0.
    ano0 = numero(g.get("invest_ano0_pct"))
    cronograma = [ano0 / 100.0, 1 - ano0 / 100.0] if 0 < ano0 < 100 else None

    marr = numero(g.get("marr"))
    wacc = numero(g.get("wacc")) or marr
    viabilidade = analisar_viabilidade_investimento(
        investimento_inicial=investimento_inicial,
        fluxos_caixa=fluxos,
        taxa=marr,
        valor_residual=valor_residual,
        cronograma_investimento=cronograma,
    )
    inv_ano0, fluxos_completos, _ = aplicar_cronograma_investimento(investimento_inicial, fluxos, cronograma)
    vpl_wacc = valor_presente_liquido(inv_ano0, fluxos_completos, wacc, valor_residual)

    # ---- avisos ----------------------------------------------------------
    avisos = []
    if capex:
        avisos.append(
            f"Investimento estimado pelo método dos fatores (como a aba Capital Investment da planilha EVTE): "
            f"FCI = US$ {fmt_num(fci)}; investimento total (TCI, inclui capital de giro e partida) = US$ {fmt_num(capex['tci'])}. "
            "Os fatores podem ser ajustados nas premissas avançadas."
        )
    if fci_informado <= 0 and not capex:
        avisos.append(
            f"FCI não informado: foi usado o custo de equipamentos (US$ {fmt_num(equip['custo_total'])}) como base dos custos e da depreciação. "
            "A planilha não traz o capital fixo completo (instalação, tubulações, edificações etc.), "
            "então esses valores provavelmente estão subestimados."
        )
    if investimento_informado <= 0 and not capex:
        avisos.append(
            f"Investimento total inicial (TCI) não informado: foi usado o FCI (US$ {fmt_num(fci)}) como investimento inicial. "
            "Capital de giro e custos de partida podem estar fora do cálculo do VPL, TIR e payback."
        )
    if edificacoes <= 0 and not capex:
        avisos.append("Custo das edificações não informado: sem depreciação de edificações e sem valor residual.")
    if razao == 1.0 and any(i["ajustado_cepci"] for i in equip["itens"]):
        avisos.append("Há equipamentos marcados para ajuste pelo CEPCI, mas os índices não foram informados "
                      "(ou são iguais): nenhum ajuste foi aplicado.")
    if equip["sem_custo"]:
        avisos.append("Equipamentos sem quantidade ou custo (não entram no total): " + ", ".join(equip["sem_custo"]) + ".")
    if capacidade <= 0:
        avisos.append("Matéria-prima processada igual a zero: rendimentos não puderam ser calculados.")
    if anos <= 0:
        avisos.append("Vida útil do projeto não informada: não há fluxo de caixa para calcular VPL e TIR.")

    return {
        "rendimento1": rendimento1,
        "rendimento2": rendimento2,
        "energia_eletrica_kwh": energia_kwh,
        "equipamentos": equip,
        "precos": {"etanol": preco_etanol, "co2": preco_co2, "produto1_t": preco_p1_t, "produto2_t": preco_p2_t},
        "receita": receita,
        "custos": custos,
        "custo_manufatura": custo_manufatura,
        "custo_total_anual": custo_total_anual,
        "depreciacao": depreciacao,
        "lucro_antes_imposto": lucro_antes_imposto,
        "imposto": imposto,
        "lucro_liquido": lucro_liquido,
        "fluxo_caixa_anual": fluxo_anual,
        "anos": anos,
        "fci": fci,
        "fci_informado": fci_informado > 0,
        "investimento_inicial": investimento_inicial,
        "investimento_informado": investimento_informado > 0,
        "edificacoes": edificacoes,
        "valor_residual": valor_residual,
        "viabilidade": viabilidade,
        "vpl_wacc": vpl_wacc,
        "fluxos_operacionais": fluxos,
        "cronograma_investimento": cronograma,
        "avisos": avisos,
    }


def analisar(globais, cenarios):
    """Roda todos os cenários e indica o de maior VPL."""
    razao = razao_cepci(globais.get("cepci_base"), globais.get("cepci_atual"))
    resultados = []
    for i, c in enumerate(cenarios, start=1):
        r = calcular_cenario(globais, c, razao)
        r["nome"] = (c.get("nome") or "").strip() or f"Cenário {i}"
        resultados.append(r)

    melhor = None
    if len(resultados) > 1:
        melhor = max(resultados, key=lambda r: r["viabilidade"]["vpl"])["nome"]

    return {"razao_cepci": razao, "cenarios": resultados, "melhor": melhor,
            "taxa_imposto": numero(globais.get("imposto"))}
