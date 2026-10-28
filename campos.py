"""
Campos que o USUÁRIO digita (extraídos das células inseridas manualmente em Teste.xlsx).
Tudo o que a planilha calculava (rendimentos, razão CEPCI, valor ajustado, totais,
consumo elétrico somado, preços finais) fica em tea.py e NÃO é pedido ao usuário.

Cada campo: (chave, rótulo, unidade)
"""

GLOBAIS_GRUPOS = [
    ("Ajuste de custo de equipamentos (índice CEPCI)", [
        ("cepci_ano_base", "Ano-base do custo de referência", "ano"),
        ("cepci_base", "Índice CEPCI do ano-base", "índice"),
        ("cepci_ano_atual", "Ano atual", "ano"),
        ("cepci_atual", "Índice CEPCI do ano atual", "índice"),
    ]),
    ("Custos unitários", [
        ("custo_lupulo", "Matéria-prima (lúpulo)", "US$/t"),
        ("etanol_preco", "Etanol — preço de lista", "US$/t"),
        ("etanol_desconto", "Etanol — desconto", "%"),
        ("co2_preco", "CO₂ — preço de lista", "US$/t"),
        ("co2_desconto", "CO₂ — desconto", "%"),
        ("energia_eletrica", "Energia elétrica", "US$/kWh"),
        ("residuos", "Descarte de resíduos", "US$/t"),
        ("mao_obra_hora", "Mão de obra operacional (com encargos)", "US$/h"),
        ("agua", "Água", "US$/kg"),
    ]),
    ("Critérios financeiros", [
        ("vida_equipamentos", "Vida útil dos equipamentos (depreciação linear)", "anos"),
        ("vida_edificacoes", "Vida útil das edificações (depreciação linear)", "anos"),
        ("imposto", "Imposto local sobre o lucro antes do imposto", "%"),
        ("marr", "Taxa mínima de atratividade (MARR)", "% a.a."),
        ("wacc", "Custo médio ponderado de capital (WACC)", "% a.a."),
        ("valor_residual_pct", "Valor residual (% do custo das edificações)", "%"),
        ("vida_projeto", "Vida útil do projeto", "anos"),
    ]),
    ("Critérios de custo (percentuais)", [
        ("manutencao", "Manutenção e reparos", "% do capital fixo (FCI)"),
        ("suprimentos", "Suprimentos operacionais", "% da manutenção"),
        ("laboratorio", "Encargos de laboratório", "% da mão de obra operacional"),
        ("patentes", "Patentes e royalties", "% da receita"),
        ("contingencia", "Contingência", "% do capital fixo (FCI)"),
        ("impostos_locais", "Impostos locais", "% do capital fixo (FCI)"),
        ("seguro", "Seguro", "% do capital fixo (FCI)"),
        ("embalagem", "Embalagem", "% da mão de obra operacional + manutenção"),
        ("mao_obra_indireta", "Mão de obra indireta", "% da mão de obra operacional + manutenção"),
        ("armazenagem", "Armazém e estocagem", "% da mão de obra operacional + manutenção"),
        ("administrativas", "Despesas administrativas", "% do custo de manufatura"),
        ("vendas", "Despesas de vendas e distribuição", "% da receita"),
        ("pd", "Pesquisa e desenvolvimento", "% da receita"),
    ]),
]

CENARIO_GRUPOS = [
    ("Operação", [
        ("capacidade", "Matéria-prima processada (entrada)", "t/ano"),
        ("horas_dia", "Duração do turno", "horas/dia"),
        ("dias_ano", "Dias de operação por ano", "dias/ano"),
        ("operadores", "Operadores por turno", "nº"),
    ]),
    ("Produtos e receitas", [
        ("produto1", "Produto 1 (saída) — extrato", "t/ano"),
        ("preco_produto1", "Preço do produto 1", "US$/kg"),
        ("produto2", "Produto 2 (saída) — farelo desengordurado", "t/ano"),
        ("preco_produto2", "Preço do produto 2", "US$/kg"),
    ]),
    ("Insumos e utilidades", [
        ("etanol", "Reagente: etanol", "t/ano"),
        ("co2", "Reagente: CO₂", "t/ano"),
        ("agua", "Água", "t/ano"),
        ("residuos", "Descarte de resíduos", "t/ano"),
    ]),
    ("Investimento (opcional no preenchimento manual)", [
        ("fci", "Investimento de capital fixo — FCI", "US$"),
        ("investimento_inicial", "Investimento total inicial — TCI (inclui capital de giro e partida)", "US$"),
        ("edificacoes", "Custo das edificações", "US$"),
    ]),
]

EQUIP_CAMPOS = [
    ("nome", "Equipamento", "text"),
    ("quantidade", "Qtd.", "number"),
    ("custo_unidade", "Custo unitário (US$)", "number"),
    ("capacidade", "Capacidade", "number"),
    ("unidade_capacidade", "Unidade", "text"),
    ("consumo_kwh_ano", "Consumo elétrico (kWh/ano)", "number"),
    ("material", "Material", "text"),
    ("fonte", "Fonte do preço", "text"),
    ("ajustar_cepci", "Ajustar pelo CEPCI?", "select"),
]

# ---------------------------------------------------------------------------
# Valores de Teste.xlsx — usados só como exemplo (botão "Carregar exemplo")
# e para gerar a planilha-modelo.
# ---------------------------------------------------------------------------
EXEMPLO_GLOBAIS = {
    "cepci_ano_base": 2014, "cepci_base": 576.1, "cepci_ano_atual": 2024, "cepci_atual": 799.1,
    "custo_lupulo": 500, "etanol_preco": 500, "etanol_desconto": 10,
    "co2_preco": 100, "co2_desconto": 10, "energia_eletrica": 0.14, "residuos": 0.53,
    "mao_obra_hora": 25, "agua": 0.008,
    "vida_equipamentos": 10, "vida_edificacoes": 25, "imposto": 34, "marr": 20, "wacc": 20,
    "valor_residual_pct": 60, "vida_projeto": 10,
    "manutencao": 6, "suprimentos": 10, "laboratorio": 10, "patentes": 2, "contingencia": 9,
    "impostos_locais": 2, "seguro": 1, "embalagem": 6, "mao_obra_indireta": 40,
    "armazenagem": 3.5, "administrativas": 8, "vendas": 10, "pd": 2,
}

EXEMPLO_CENARIOS = [
    {
        "nome": "Extração via solvente EtOH puro",
        "capacidade": 20, "horas_dia": 8, "dias_ano": 330, "operadores": 2,
        "produto1": 0.8884316390400001, "preco_produto1": 250,
        "produto2": 0, "preco_produto2": 0,
        "etanol": 103.02, "co2": 0, "agua": 135.57, "residuos": 0.1600104,
        "fci": "", "edificacoes": "",
        "equipamentos": [
            {"nome": "Lavador (sem cotação)", "quantidade": 0, "custo_unidade": 0, "capacidade": 100, "unidade_capacidade": "kg/h", "consumo_kwh_ano": "", "material": "aço inox", "fonte": "macanuda.com.br (lavador de alface)", "ajustar_cepci": "não"},
            {"nome": "Centrífuga (sem cotação)", "quantidade": 0, "custo_unidade": 0, "capacidade": 30, "unidade_capacidade": "kg", "consumo_kwh_ano": "", "material": "aço inox", "fonte": "eBay", "ajustar_cepci": "não"},
            {"nome": "Secador / desidratador (sem cotação)", "quantidade": 0, "custo_unidade": 0, "capacidade": 10.8, "unidade_capacidade": "m2", "consumo_kwh_ano": "", "material": "aço inox", "fonte": "macanuda.com.br (desidratador)", "ajustar_cepci": "não"},
            {"nome": "Extrator etanol", "quantidade": 1, "custo_unidade": 22500, "capacidade": 37.6, "unidade_capacidade": "kg/h", "consumo_kwh_ano": 990, "material": "aço inox", "fonte": "laboao", "ajustar_cepci": "não"},
            {"nome": "Evaporador", "quantidade": 1, "custo_unidade": 33800, "capacidade": 100, "unidade_capacidade": "L/h", "consumo_kwh_ano": 400.47103488816964, "material": "aço inox tubos", "fonte": "timerhaus", "ajustar_cepci": "não"},
            {"nome": "Secador (ábaco, sem valor)", "quantidade": 0, "custo_unidade": 0, "capacidade": 10, "unidade_capacidade": "ft2", "consumo_kwh_ano": "", "material": "aço carbono", "fonte": "timerhaus / matche.com", "ajustar_cepci": "sim"},
        ],
    },
    {
        "nome": "CO2 Supercrítico",
        "capacidade": 20, "horas_dia": 8, "dias_ano": 330, "operadores": 2,
        "produto1": 0.62442458496, "preco_produto1": 512.22,
        "produto2": 3.6425194150399998, "preco_produto2": 0.03,
        "etanol": 0, "co2": 112.58203735679999, "agua": 100.0065, "residuos": 0.1600104,
        "fci": "", "edificacoes": "",
        "equipamentos": [
            {"nome": "Lavador (sem cotação)", "quantidade": 0, "custo_unidade": 0, "capacidade": 100, "unidade_capacidade": "kg/h", "consumo_kwh_ano": "", "material": "aço inox", "fonte": "macanuda.com.br (lavador de alface)", "ajustar_cepci": "não"},
            {"nome": "Centrífuga", "quantidade": 1, "custo_unidade": 10999, "capacidade": 30, "unidade_capacidade": "kg", "consumo_kwh_ano": 247.5, "material": "aço inox", "fonte": "eBay", "ajustar_cepci": "não"},
            {"nome": "Secador", "quantidade": 1, "custo_unidade": 5052.55, "capacidade": 10.8, "unidade_capacidade": "m2", "consumo_kwh_ano": 6097.926527168509, "material": "aço inox", "fonte": "macanuda.com.br (desidratador)", "ajustar_cepci": "não"},
            {"nome": "Extrator supercrítico", "quantidade": 1, "custo_unidade": 55805, "capacidade": 15, "unidade_capacidade": "L/h", "consumo_kwh_ano": 740.8309258350001, "material": "aço inox", "fonte": "Labsolutions", "ajustar_cepci": "não"},
        ],
    },
]
