import pandas as pd
import re
from openpyxl import load_workbook

ABAS = ("Globais", "Cenarios", "Equipamentos")
OBRIGATORIOS = {
    "cepci_base": "índice CEPCI do ano-base",
    "cepci_atual": "índice CEPCI do ano atual",
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


def ler_planilha_tea(arquivo):
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
        return ler_planilha_evte(arquivo)

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


def ler_planilha_evte(arquivo):
    """Importa o formato do arquivo 'Modelo Dados Processo EVTE'."""
    wb = load_workbook(arquivo, data_only=True, read_only=True)
    ws = wb["Tecno Assumptions"]
    eq = wb["Custo dos Equipamentos"]
    cap = wb["Capital Investment"]

    def v(sheet, cell):
        return _valor(sheet[cell].value)

    globais = {
        "cepci_ano_base": _num_planilha(v(eq, "E8")),
        "cepci_base": _num_planilha(v(eq, "F8")),
        "cepci_ano_atual": _num_planilha(v(eq, "E9")),
        "cepci_atual": _num_planilha(v(eq, "F9")),
        "custo_lupulo": _num_planilha(v(ws, "D39")),
        # Os valores do arquivo original já incluem os descontos de compra.
        "etanol_preco": _num_planilha(v(ws, "D33")), "etanol_desconto": 0,
        "co2_preco": _num_planilha(v(ws, "D34")), "co2_desconto": 0,
        "energia_eletrica": _num_planilha(v(ws, "D35")),
        "residuos": _num_planilha(v(ws, "D36")),
        "mao_obra_hora": _num_planilha(v(ws, "D37")),
        "agua": _num_planilha(v(ws, "D38")),
        "vida_equipamentos": _num_planilha(v(ws, "D44")),
        "vida_edificacoes": _num_planilha(v(ws, "D45")),
        "imposto": _num_planilha(v(ws, "D46")) * 100,
        "marr": _num_planilha(v(ws, "D47")),
        "wacc": _num_planilha(v(ws, "D48")),
        "valor_residual_pct": _num_planilha(v(ws, "D49")),
        "vida_projeto": _num_planilha(v(ws, "D50")),
        "manutencao": _num_planilha(v(ws, "D53")),
        "suprimentos": _num_planilha(v(ws, "D54")),
        "laboratorio": _num_planilha(v(ws, "D55")),
        "patentes": _num_planilha(v(ws, "D56")),
        "contingencia": _num_planilha(v(ws, "D57")),
        "impostos_locais": _num_planilha(v(ws, "D58")),
        "seguro": _num_planilha(v(ws, "D59")),
        "embalagem": _num_planilha(v(ws, "D60")),
        "mao_obra_indireta": _num_planilha(v(ws, "D61")),
        "armazenagem": _num_planilha(v(ws, "D62")),
        "administrativas": _num_planilha(v(ws, "D63")),
        "vendas": _num_planilha(v(ws, "D64")),
        "pd": _num_planilha(v(ws, "D65")),
    }

    # A planilha apresenta cada processo numa coluna: D (EtOH) e E (CO2).
    cenarios = []
    for col, cap_fci, cap_tci, edif_fci, equipamentos in (
        ("D", "B27", "B30", "B18", range(13, 19)),
        ("E", "G25", "G28", "G16", range(20, 24)),
    ):
        nome = " ".join(str(v(ws, f"{col}{r}")).strip() for r in (2, 3) if v(ws, f"{col}{r}") != "")
        if not nome:
            nome = "Extração via solvente EtOH puro" if col == "D" else "CO2 Supercrítico"
        c = {
            "nome": nome,
            "capacidade": _num_planilha(v(ws, f"{col}6")),
            "horas_dia": _num_planilha(v(ws, f"{col}7")),
            "dias_ano": _num_planilha(v(ws, f"{col}8")),
            "operadores": _num_planilha(v(ws, f"{col}28")),
            "produto1": _num_planilha(v(ws, f"{col}13")),
            "preco_produto1": _num_planilha(v(ws, f"{col}40")) / 1000,
            "produto2": _num_planilha(v(ws, f"{col}14")),
            "preco_produto2": _num_planilha(v(ws, f"{col}41")) / 1000,
            "etanol": _num_planilha(v(ws, f"{col}20")),
            "co2": _num_planilha(v(ws, f"{col}21")),
            "agua": _num_planilha(v(ws, f"{col}26")),
            "residuos": _num_planilha(v(ws, f"{col}25")),
            "fci": _num_planilha(v(cap, cap_fci)),
            "investimento_inicial": _num_planilha(v(cap, cap_tci)),
            "edificacoes": _num_planilha(v(cap, edif_fci)),
            "equipamentos": [],
        }
        for row in equipamentos:
            nome_eq = v(eq, f"B{row}")
            if not nome_eq:
                nome_eq = f"Equipamento (linha {row})"
            c["equipamentos"].append({
                "nome": str(nome_eq),
                "quantidade": _num_planilha(v(eq, f"C{row}")),
                "custo_unidade": _num_planilha(v(eq, f"D{row}")),
                "capacidade": _num_planilha(v(eq, f"G{row}")),
                "unidade_capacidade": str(v(eq, f"H{row}") or ""),
                "consumo_kwh_ano": _num_planilha(v(eq, f"I{row}")),
                "material": str(v(eq, f"J{row}") or ""),
                "fonte": str(v(eq, f"K{row}") or v(eq, f"L{row}") or ""),
                "ajustar_cepci": "sim" if row == 18 and col == "D" else "não",
            })
        cenarios.append(c)

    ausentes = [rotulo for chave, rotulo in OBRIGATORIOS.items()
                if str(globais.get(chave, "")).strip() == ""]
    if ausentes:
        raise ValueError("Faltam valores obrigatórios na planilha EVTE: " + ", ".join(ausentes))
    return globais, cenarios
