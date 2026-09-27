import math


def numero(valor):
    """Converte valores recebidos do formulário para float."""
    try:
        if valor is None or valor == "":
            return 0.0

        if isinstance(valor, (int, float)):
            return float(valor)

        valor = str(valor).replace(",", ".")
        return float(valor)

    except (ValueError, TypeError):
        return 0.0


def calcular_estatisticas(valores):
    """
    Calcula média, desvio padrão amostral e RSD.
    """

    valores = [numero(v) for v in valores]

    if not valores:
        return {
            "media": 0,
            "desvio_padrao": 0,
            "rsd": 0,
            "n": 0
        }

    media = sum(valores) / len(valores)

    if len(valores) > 1:
        variancia = sum((x - media) ** 2 for x in valores) / (len(valores) - 1)
        desvio = math.sqrt(variancia)
    else:
        desvio = 0

    rsd = (desvio / media * 100) if media != 0 else 0

    return {
        "media": media,
        "desvio_padrao": desvio,
        "rsd": rsd,
        "n": len(valores)
    }


def calcular_projeto(dados):
    """
    Calcula os custos do projeto acadêmico.
    """

    materiais = dados.get("materiais", [])
    equipamentos = dados.get("equipamentos", [])

    custo_materiais = 0
    custo_equipamentos = 0

    # Materiais
    for material in materiais:

        quantidade = numero(material.get("quantidade"))
        preco = numero(material.get("preco"))

        material["custo"] = quantidade * preco

        custo_materiais += material["custo"]

    # Equipamentos
    for equipamento in equipamentos:

        horas = numero(equipamento.get("horas"))
        custo_hora = numero(equipamento.get("custo_hora"))

        equipamento["custo"] = horas * custo_hora

        custo_equipamentos += equipamento["custo"]

    energia = numero(dados.get("energia"))
    mao_obra = numero(dados.get("mao_obra"))
    outros = numero(dados.get("outros"))

    custo_total = (
        custo_materiais
        + custo_equipamentos
        + energia
        + mao_obra
        + outros
    )

    quantidade_produzida = numero(
        dados.get("quantidade_produzida")
    )

    if quantidade_produzida > 0:
        custo_unitario = custo_total / quantidade_produzida
    else:
        custo_unitario = 0

    # Experimentos
    experimentos = dados.get("experimentos", [])

    estatisticas = calcular_estatisticas(experimentos)

    return {
        "custo_materiais": custo_materiais,
        "custo_equipamentos": custo_equipamentos,
        "energia": energia,
        "mao_obra": mao_obra,
        "outros": outros,
        "custo_total": custo_total,
        "quantidade_produzida": quantidade_produzida,
        "custo_unitario": custo_unitario,
        "estatisticas": estatisticas
    }


def comparar_alternativas(projeto, alternativas):
    """
    Compara o projeto principal com alternativas cadastradas.
    """

    resultados = []

    resultados.append({
        "nome": projeto.get("nome", "Projeto proposto"),
        "custo": numero(projeto.get("custo_total")),
        "tempo": numero(projeto.get("tempo")),
        "rendimento": numero(projeto.get("rendimento"))
    })

    for alternativa in alternativas:

        resultados.append({
            "nome": alternativa.get("nome", "Alternativa"),
            "custo": numero(alternativa.get("custo")),
            "tempo": numero(alternativa.get("tempo")),
            "rendimento": numero(alternativa.get("rendimento"))
        })

    return resultados


def calcular_cenarios(custo):

    custo = numero(custo)

    return {
        "otimista": custo * 0.90,
        "esperado": custo,
        "pessimista": custo * 1.20
    }


def analise_sensibilidade(custo, variacoes=None):

    custo = numero(custo)

    if variacoes is None:
        variacoes = [-20, -10, 0, 10, 20]

    resultado = []

    for variacao in variacoes:

        novo_custo = custo * (1 + variacao / 100)

        resultado.append({
            "variacao": variacao,
            "custo": novo_custo
        })

    return resultado


# ---------------------------------------------------------------------------
# Análise de investimento (VP, VF, VPL, TIR, Payback e viabilidade econômica)
# ---------------------------------------------------------------------------

def valor_presente(valor_futuro, taxa, periodos):
    """
    VP = VF / (1 + i)^n
    'taxa' em % ao período.
    """
    valor_futuro = numero(valor_futuro)
    taxa_decimal = numero(taxa) / 100
    periodos = numero(periodos)

    return valor_futuro / ((1 + taxa_decimal) ** periodos)


def valor_futuro(valor_presente_aplicado, taxa, periodos):
    """
    VF = VP * (1 + i)^n
    'taxa' em % ao período.
    """
    valor_presente_aplicado = numero(valor_presente_aplicado)
    taxa_decimal = numero(taxa) / 100
    periodos = numero(periodos)

    return valor_presente_aplicado * ((1 + taxa_decimal) ** periodos)


def valor_presente_liquido(investimento_inicial, fluxos_caixa, taxa, valor_residual=0):
    """
    VPL = -Investimento + soma( FC_t / (1+i)^t ) + ValorResidual / (1+i)^n

    fluxos_caixa: lista com o fluxo de caixa líquido de cada período (1, 2, 3, ...).
    """
    investimento_inicial = numero(investimento_inicial)
    taxa_decimal = numero(taxa) / 100
    valor_residual = numero(valor_residual)

    vpl = -investimento_inicial

    for periodo, fluxo in enumerate(fluxos_caixa, start=1):
        vpl += numero(fluxo) / ((1 + taxa_decimal) ** periodo)

    n = len(fluxos_caixa)
    if valor_residual and n > 0:
        vpl += valor_residual / ((1 + taxa_decimal) ** n)

    return vpl


def taxa_interna_retorno(investimento_inicial, fluxos_caixa, valor_residual=0,
                          precisao=0.0001, max_iteracoes=1000):
    """
    Calcula a TIR (%) por bisseção: procura a taxa que zera o VPL.
    Retorna None se não existir raiz no intervalo pesquisado (-99% a 1000%).
    """
    investimento_inicial = numero(investimento_inicial)
    valor_residual = numero(valor_residual)
    fluxos = [numero(f) for f in fluxos_caixa]

    if investimento_inicial <= 0 or not fluxos:
        return None

    def vpl_para_taxa(taxa_pct):
        taxa = taxa_pct / 100
        vpl = -investimento_inicial
        for periodo, fluxo in enumerate(fluxos, start=1):
            vpl += fluxo / ((1 + taxa) ** periodo)
        n = len(fluxos)
        if valor_residual and n > 0:
            vpl += valor_residual / ((1 + taxa) ** n)
        return vpl

    taxa_min, taxa_max = -99.0, 1000.0
    vpl_min = vpl_para_taxa(taxa_min)
    vpl_max = vpl_para_taxa(taxa_max)

    if vpl_min == 0:
        return taxa_min
    if vpl_max == 0:
        return taxa_max
    if vpl_min * vpl_max > 0:
        return None  # não há troca de sinal no intervalo pesquisado

    taxa_meio = (taxa_min + taxa_max) / 2

    for _ in range(max_iteracoes):
        taxa_meio = (taxa_min + taxa_max) / 2
        vpl_meio = vpl_para_taxa(taxa_meio)

        if abs(vpl_meio) < precisao:
            return taxa_meio

        if vpl_min * vpl_meio < 0:
            taxa_max = taxa_meio
            vpl_max = vpl_meio
        else:
            taxa_min = taxa_meio
            vpl_min = vpl_meio

    return taxa_meio


def payback_simples(investimento_inicial, fluxos_caixa):
    """
    Nº de períodos (fracionário) até recuperar o investimento, sem descontar
    os fluxos. Retorna None se o investimento nunca é recuperado.
    """
    investimento_inicial = numero(investimento_inicial)
    saldo = -investimento_inicial

    for periodo, fluxo in enumerate(fluxos_caixa, start=1):
        fluxo = numero(fluxo)
        saldo_anterior = saldo
        saldo += fluxo

        if saldo >= 0:
            if fluxo == 0:
                return periodo
            return (periodo - 1) + (-saldo_anterior / fluxo)

    return None


def payback_descontado(investimento_inicial, fluxos_caixa, taxa):
    """
    Igual ao payback simples, mas descontando cada fluxo de caixa pela taxa
    informada antes de somar — considera o valor do dinheiro no tempo.
    """
    investimento_inicial = numero(investimento_inicial)
    taxa_decimal = numero(taxa) / 100
    saldo = -investimento_inicial

    for periodo, fluxo in enumerate(fluxos_caixa, start=1):
        fluxo_descontado = numero(fluxo) / ((1 + taxa_decimal) ** periodo)
        saldo_anterior = saldo
        saldo += fluxo_descontado

        if saldo >= 0:
            if fluxo_descontado == 0:
                return periodo
            return (periodo - 1) + (-saldo_anterior / fluxo_descontado)

    return None


def analisar_viabilidade_investimento(investimento_inicial, fluxos_caixa, taxa, valor_residual=0):
    """
    Reúne VP, VF, VPL, TIR e Payback num único resultado e emite o veredito
    de viabilidade econômica do investimento, com os motivos da decisão.
    """
    investimento_inicial = numero(investimento_inicial)
    taxa_num = numero(taxa)
    valor_residual = numero(valor_residual)
    n = len(fluxos_caixa)

    vp_fluxos = sum(
        valor_presente(fluxo, taxa_num, periodo)
        for periodo, fluxo in enumerate(fluxos_caixa, start=1)
    )
    vp_residual = valor_presente(valor_residual, taxa_num, n) if (valor_residual and n > 0) else 0
    vf_investimento = valor_futuro(investimento_inicial, taxa_num, n) if n > 0 else 0

    vpl = valor_presente_liquido(investimento_inicial, fluxos_caixa, taxa_num, valor_residual)
    tir = taxa_interna_retorno(investimento_inicial, fluxos_caixa, valor_residual)
    payback_s = payback_simples(investimento_inicial, fluxos_caixa)
    payback_d = payback_descontado(investimento_inicial, fluxos_caixa, taxa_num)

    vpl_positivo = vpl > 0
    tir_acima_da_taxa = (tir is not None) and (tir > taxa_num)

    viavel = bool(investimento_inicial > 0 and vpl_positivo and tir_acima_da_taxa)

    motivos = []

    if investimento_inicial <= 0:
        motivos.append("não há um valor de investimento inicial válido para ser analisado")
    else:
        if vpl_positivo:
            motivos.append(
                "o Valor Presente Líquido (VPL) é positivo, ou seja, os retornos trazidos a "
                "valor presente superam o investimento inicial"
            )
        else:
            motivos.append(
                "o Valor Presente Líquido (VPL) é negativo ou nulo, ou seja, o investimento "
                "não se paga considerando a taxa mínima de atratividade informada"
            )

        if tir is None:
            motivos.append(
                "não foi possível calcular uma Taxa Interna de Retorno (TIR) real para os "
                "fluxos de caixa informados"
            )
        elif tir_acima_da_taxa:
            motivos.append(
                f"a Taxa Interna de Retorno (TIR) de {tir:.2f}% é maior que a taxa mínima de "
                f"atratividade de {taxa_num:.2f}%"
            )
        else:
            motivos.append(
                f"a Taxa Interna de Retorno (TIR) de {tir:.2f}% é menor que a taxa mínima de "
                f"atratividade de {taxa_num:.2f}%"
            )

        if payback_d is None:
            motivos.append(
                "o investimento não se paga dentro do horizonte de tempo informado, mesmo "
                "descontando os fluxos de caixa"
            )
        else:
            motivos.append(
                f"o investimento se paga em aproximadamente {payback_d:.2f} período(s), "
                "considerando o valor do dinheiro no tempo (payback descontado)"
            )

    return {
        "investimento_inicial": investimento_inicial,
        "taxa_desconto": taxa_num,
        "valor_residual": valor_residual,
        "numero_periodos": n,
        "valor_presente_fluxos": vp_fluxos,
        "valor_presente_residual": vp_residual,
        "valor_futuro_investimento": vf_investimento,
        "vpl": vpl,
        "tir": tir,
        "payback_simples": payback_s,
        "payback_descontado": payback_d,
        "viavel": viavel,
        "motivos": motivos,
    }
