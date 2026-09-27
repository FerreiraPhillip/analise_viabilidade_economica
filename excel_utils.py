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
    (dados, alternativas) no formato esperado por analysis.py.

    Abas esperadas:
      - Parametros    : nome, tempo, rendimento, energia, mao_obra, outros,
                        quantidade_produzida  (uma única linha)
      - Materiais     : nome, quantidade, preco
      - Equipamentos  : nome, horas, custo_hora
      - Experimentos  : valor
      - Alternativas  : nome, custo, tempo, rendimento  (opcional)
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

    return dados, alternativas
