"""Cálculo genérico de viabilidade econômica por cenário."""

from analysis import analisar_viabilidade_investimento, numero, valor_presente_liquido


def _normalizar_itens(items, rotulo):
    normalizados = []
    for item in items:
        quantidade = numero(item.get("quantidade", 1))
        unitario = numero(item.get("valor_unitario"))
        valor = numero(item.get("valor")) if item.get("valor") is not None else quantidade * unitario
        normalizados.append({
            "nome": item.get("nome") or rotulo,
            "quantidade": quantidade,
            "valor_unitario": unitario,
            "valor": valor,
        })
    return normalizados


def _items_total(items):
    return sum(item["valor"] for item in items)


def analisar_projeto(globais, cenarios):
    """Avalia cenários com benefícios/receitas, custos, investimento e fluxo anual."""
    taxa = numero(globais.get("taxa_desconto"))
    if 0 < taxa <= 1:
        taxa *= 100
    wacc = numero(globais.get("wacc", taxa))
    if 0 < wacc <= 1:
        wacc *= 100
    imposto = numero(globais.get("taxa_imposto"))
    if 0 < imposto <= 1:
        imposto *= 100
    anos = max(0, int(round(numero(globais.get("vida_projeto")))))
    resultados = []

    for i, c in enumerate(cenarios, start=1):
        investimentos = _normalizar_itens(c.get("investimentos", []), "Investimento")
        beneficios = _normalizar_itens(c.get("beneficios", []), "Benefício")
        custos = _normalizar_itens(c.get("custos", []), "Custo")
        total_beneficios = _items_total(beneficios)
        total_custos = _items_total(custos)
        ativos = []
        for ativo in c.get("ativos_depreciaveis", []):
            base = numero(ativo.get("base"))
            vida = numero(ativo.get("vida_util"))
            ativos.append({
                "nome": ativo.get("nome") or "Ativo",
                "base": base,
                "vida_util": vida,
                "depreciacao_anual": base / vida if base > 0 and vida > 0 else 0,
            })
        depreciacao = sum(a["depreciacao_anual"] for a in ativos)
        fluxos = c.get("fluxos_caixa")
        lucro_antes_imposto = total_beneficios - total_custos - depreciacao
        imposto_anual = max(lucro_antes_imposto, 0) * imposto / 100
        if fluxos is None:
            fluxo_anual = lucro_antes_imposto - imposto_anual + depreciacao
            fluxos = [fluxo_anual] * anos
        else:
            fluxos = [numero(x) for x in fluxos]
            # Um cronograma direto já deve conter os fluxos líquidos após impostos.
            fluxo_anual = sum(fluxos) / len(fluxos) if fluxos else 0
            depreciacao = 0
            imposto_anual = 0
        investimento_itens = _items_total(investimentos)
        investimento = investimento_itens if investimento_itens > 0 else numero(c.get("investimento_inicial"))
        residual = numero(c.get("valor_residual"))
        viabilidade = analisar_viabilidade_investimento(
            investimento_inicial=investimento,
            fluxos_caixa=fluxos,
            taxa=taxa,
            valor_residual=residual,
        )
        resultados.append({
            "nome": (c.get("nome") or "").strip() or f"Cenário {i}",
            "investimento_inicial": investimento,
            "investimentos": investimentos,
            "valor_residual": residual,
            "beneficios": beneficios,
            "custos": custos,
            "beneficios_anuais": total_beneficios,
            "custos_anuais": total_custos,
            "fluxo_liquido_anual": total_beneficios - total_custos,
            "depreciacao_anual": depreciacao,
            "lucro_antes_imposto": lucro_antes_imposto,
            "imposto_anual": imposto_anual,
            "fluxo_anual": fluxo_anual,
            "ativos_depreciaveis": ativos,
            "fluxos_caixa": fluxos,
            "fluxo_direto": c.get("fluxos_caixa") is not None,
            "anos": len(fluxos),
            "viabilidade": viabilidade,
            "vpl_wacc": valor_presente_liquido(
                investimento, fluxos, wacc, residual
            ),
            "avisos": list(c.get("avisos", [])),
        })
        if c.get("fluxos_caixa") is None and imposto <= 0:
            resultados[-1]["avisos"].append("Impostos não incluídos: o fluxo anual está antes dos impostos.")
        elif c.get("fluxos_caixa") is None and imposto > 0 and depreciacao <= 0:
            resultados[-1]["avisos"].append("Nenhum ativo depreciável foi informado; o imposto foi estimado sem dedução de depreciação.")

    melhor = max(resultados, key=lambda x: x["viabilidade"]["vpl"])["nome"] if len(resultados) > 1 else None
    return {"taxa_desconto": taxa, "cenarios": resultados, "melhor": melhor}


def converter_resultado_tea(resultado_tea):
    """Apresenta uma análise TEA já calculada usando a tela genérica de resultados."""
    cenarios = []
    for c in resultado_tea["cenarios"]:
        cenarios.append({
            "nome": c["nome"],
            "investimento_inicial": c["investimento_inicial"],
            "valor_residual": c["valor_residual"],
            "beneficios": [{"nome": "Receita anual", "valor": c["receita"]}],
            "custos": [{"nome": x["nome"], "valor": x["valor"]} for x in c["custos"]],
            "fluxos_caixa": [c["fluxo_caixa_anual"]] * c["anos"],
            "avisos": c["avisos"],
        })
    result = analisar_projeto(
        {"taxa_desconto": resultado_tea["cenarios"][0]["viabilidade"]["taxa_desconto"] if cenarios else 0},
        cenarios,
    )
    for target, source in zip(result["cenarios"], resultado_tea["cenarios"]):
        target["viabilidade"] = source["viabilidade"]
        target["vpl_wacc"] = source["vpl_wacc"]
        target["depreciacao_anual"] = source["depreciacao"]
        target["lucro_antes_imposto"] = source["lucro_antes_imposto"]
        target["imposto_anual"] = source["imposto"]
        target["fluxo_anual"] = source["fluxo_caixa_anual"]
        target["fluxo_direto"] = False
    return result
