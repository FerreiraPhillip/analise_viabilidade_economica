"""Prepara os dados do dashboard interativo (viram JSON na página de resultado)."""


def _top_custos(c, limite=8):
    itens = [{"nome": x["nome"], "valor": x["valor"]} for x in c.get("custos", []) if x["valor"] > 0]
    if not c.get("custos_incluem_depreciacao") and c.get("depreciacao_anual", 0) > 0 and not c.get("fluxo_direto"):
        itens.append({"nome": "Depreciação (linear)", "valor": c["depreciacao_anual"]})
    itens.sort(key=lambda x: x["valor"], reverse=True)
    if len(itens) > limite:
        outros = sum(x["valor"] for x in itens[limite - 1:])
        itens = itens[:limite - 1] + [{"nome": "Outros", "valor": outros}]
    return itens


def montar_dashboard(resultado):
    """Resumo por cenário com tudo que o navegador precisa para simular (receita, custos, imposto...)."""
    cenarios = []
    for c in resultado["cenarios"]:
        v = c["viabilidade"]
        decomposto = not c.get("fluxo_direto")
        dep = c.get("depreciacao_anual", 0.0) if decomposto else 0.0
        custo_caixa = c.get("custos_anuais", 0.0) - (dep if c.get("custos_incluem_depreciacao") else 0.0)
        cron = c.get("cronograma_investimento") or [1.0]
        soma = sum(cron) or 1.0
        cenarios.append({
            "nome": c["nome"],
            "investimento": c["investimento_inicial"],
            "cronograma": [x / soma for x in cron],
            "residual": c["valor_residual"],
            "taxa": v["taxa_desconto"],
            "imposto": c.get("imposto_pct", 0.0) if decomposto else 0.0,
            "decomposto": decomposto,
            "receita": c.get("beneficios_anuais", 0.0),
            "custo_caixa": custo_caixa,
            "depreciacao": dep,
            "anos": c["anos"],
            "fluxos_op": c["fluxos_caixa"],
            "custos": _top_custos(c),
            "base": {
                "vpl": v["vpl"], "tir": v["tir"], "payback_simples": v["payback_simples"],
                "payback_descontado": v["payback_descontado"], "viavel": v["viavel"],
            },
        })
    return {"cenarios": cenarios, "melhor": resultado.get("melhor")}
