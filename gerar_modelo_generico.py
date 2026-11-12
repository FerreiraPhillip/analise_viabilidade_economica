"""Gera static/modelo_planilha_generico.xlsx: modelo simples para qualquer projeto (não só processo industrial)."""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

CAB = PatternFill("solid", fgColor="13293D")
CAB_FONT = Font(bold=True, color="FFFFFF")
ENTRADA = PatternFill("solid", fgColor="FFF4CC")   # células a preencher


def cabecalho(ws, titulos):
    for i, t in enumerate(titulos, start=1):
        c = ws.cell(row=1, column=i, value=t)
        c.fill, c.font = CAB, CAB_FONT


def larguras(ws, valores):
    for i, w in enumerate(valores, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


wb = Workbook()

ws = wb.active
ws.title = "Leia-me"
linhas = [
    "Modelo para análise de viabilidade econômica de qualquer projeto",
    "",
    "• Preencha só as células amarelas. VPL, TIR, payback, depreciação e imposto são calculados pelo site — não precisa calcular nada aqui.",
    "• Aba Premissas: nome, investimento inicial, taxa mínima de atratividade e vida útil são OBRIGATÓRIOS. Valor residual e imposto são opcionais.",
    "• Informe os dados de UM dos dois jeitos: (a) aba 'Receitas e custos' com os valores de um ano típico (o site repete pelo horizonte), ou (b) aba 'Fluxo de caixa' com o fluxo líquido de cada ano, quando ele muda ao longo do tempo. Se preencher os dois, vale o fluxo de caixa.",
    "• Use a mesma moeda e a mesma periodicidade (anual) em tudo. Fórmulas, se houver, precisam estar calculadas e salvas.",
    "• Para comparar alternativas: numa aba, escreva 'Cenário' na célula A1 e os nomes das alternativas nas células ao lado (B1, C1...); abaixo, uma linha por dado (Investimento inicial, Receita anual, Custo anual, Taxa de desconto, Vida útil do projeto) com o valor de cada alternativa na sua coluna.",
    "• Não renomeie os rótulos da coluna A: o site encontra os dados por eles (aceita também em inglês, como 'Initial investment', 'Discount rate', 'Project life').",
]
for i, t in enumerate(linhas, start=1):
    ws.cell(row=i, column=1, value=t).alignment = Alignment(wrap_text=True, vertical="top")
ws["A1"].font = Font(bold=True, size=14)
ws.column_dimensions["A"].width = 120

ws = wb.create_sheet("Premissas")
cabecalho(ws, ["Item", "Valor", "Observação"])
dados = [
    ("Nome do projeto", "Meu projeto", "Texto livre"),
    ("Investimento inicial", 100000, "Obrigatório. Total gasto antes de começar a operar (equipamentos, obras, capital de giro...)."),
    ("Taxa mínima de atratividade (%)", 12, "Obrigatório. Retorno mínimo exigido, em % ao ano (ex.: 12)."),
    ("Vida útil do projeto (anos)", 5, "Obrigatório se usar a aba 'Receitas e custos'. Anos de operação analisados."),
    ("Valor residual", 10000, "Opcional. Quanto o projeto vale no último ano (venda de equipamentos etc.)."),
    ("Taxa de imposto (%)", None, "Opcional. Imposto sobre o lucro, em %. Em branco = análise antes dos impostos."),
]
for r, (a, b, c) in enumerate(dados, start=2):
    ws.cell(row=r, column=1, value=a)
    cel = ws.cell(row=r, column=2, value=b)
    cel.fill = ENTRADA
    ws.cell(row=r, column=3, value=c).alignment = Alignment(wrap_text=True)
larguras(ws, [36, 18, 80])

ws = wb.create_sheet("Receitas e custos")
cabecalho(ws, ["Tipo", "Item", "Valor anual"])
itens = [("Receita", "Vendas", 60000), ("Custo", "Pessoal", 18000), ("Custo", "Insumos", 12000)]
for r, linha in enumerate(itens, start=2):
    for col, v in enumerate(linha, start=1):
        c = ws.cell(row=r, column=col, value=v)
        c.fill = ENTRADA
ws["E1"] = "Tipo: escreva Receita (ou Benefício) / Custo (ou Despesa). Uma linha por item; adicione quantas precisar."
larguras(ws, [14, 34, 16, 4, 90])

ws = wb.create_sheet("Fluxo de caixa")
cabecalho(ws, ["Ano", "Fluxo de caixa líquido"])
for r in range(2, 7):
    ws.cell(row=r, column=1, value=None).fill = ENTRADA
    ws.cell(row=r, column=2, value=None).fill = ENTRADA
ws["D1"] = "Opcional: preencha (Ano = 1, 2, 3...) só se o fluxo líquido mudar de um ano para o outro. Já deve ser o valor após impostos."
larguras(ws, [10, 26, 4, 90])

wb.save("static/modelo_planilha_generico.xlsx")
print("modelo genérico gerado")
