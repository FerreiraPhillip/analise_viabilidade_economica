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