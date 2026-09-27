import pandas as pd


def _aba_para_registros(xls, nome_aba):
    """Lê uma aba da planilha e retorna uma lista de dicionários (uma por linha)."""
    if nome_aba not in xls.sheet_names:
        return []

    df = xls.parse(nome_aba)
    df = df.dropna(how="all")

    return df.to_dict(orient="records")


def ler_planilha_projeto(arquivo):
    """
    Lê uma planilha Excel (.xlsx) com o modelo do projeto e devolve
    (dados, alternativas, investimento) no formato esperado por analysis.py.

    Abas esperadas:
      - Parametros    : nome, tempo, rendimento, energia, mao_obra, outros,
                        quantidade_produzida  (uma única linha)
      - Materiais     : nome, quantidade, preco
      - Equipamentos  : nome, horas, custo_hora
      - Experimentos  : valor
      - Alternativas  : nome, custo, tempo, rendimento  (opcional)
      - Investimento  : investimento_inicial, taxa_desconto, valor_residual
                        (uma única linha; investimento_inicial pode ficar em
                        branco para usar o custo total calculado do projeto)
      - FluxoCaixa    : periodo, valor  (um fluxo de caixa líquido por período)
    """

    xls = pd.ExcelFile(arquivo)

    parametros_registros = _aba_para_registros(xls, "Parametros")
    parametros = parametros_registros[0] if parametros_registros else {}

    materiais = _aba_para_registros(xls, "Materiais")
    equipamentos = _aba_para_registros(xls, "Equipamentos")

    experimentos_registros = _aba_para_registros(xls, "Experimentos")
    experimentos = [
        r.get("valor") for r in experimentos_registros
        if r.get("valor") is not None
    ]

    alternativas = _aba_para_registros(xls, "Alternativas")

    dados = {
        "nome": parametros.get("nome", "Projeto proposto"),
        "materiais": materiais,
        "equipamentos": equipamentos,
        "energia": parametros.get("energia", 0),
        "mao_obra": parametros.get("mao_obra", 0),
        "outros": parametros.get("outros", 0),
        "quantidade_produzida": parametros.get("quantidade_produzida", 0),
        "tempo": parametros.get("tempo", 0),
        "rendimento": parametros.get("rendimento", 0),
        "experimentos": experimentos,
    }

    investimento_registros = _aba_para_registros(xls, "Investimento")
    investimento_parametros = investimento_registros[0] if investimento_registros else {}

    fluxo_caixa_registros = _aba_para_registros(xls, "FluxoCaixa")
    fluxo_caixa_registros.sort(key=lambda r: r.get("periodo", 0))
    fluxos_caixa = [
        r.get("valor") for r in fluxo_caixa_registros if r.get("valor") is not None
    ]

    investimento = {
        "investimento_inicial": investimento_parametros.get("investimento_inicial"),
        "taxa_desconto": investimento_parametros.get("taxa_desconto", 0),
        "valor_residual": investimento_parametros.get("valor_residual", 0),
        "fluxos_caixa": fluxos_caixa,
    }

    return dados, alternativas, investimento
