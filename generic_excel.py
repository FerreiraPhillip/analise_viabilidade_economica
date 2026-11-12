"""Importador genérico de planilhas de viabilidade com rótulos identificáveis."""

import re
import unicodedata
from collections import defaultdict

from openpyxl import load_workbook

from analysis import numero
from excel_utils import ler_planilha_tea
from project_analysis import analisar_projeto, converter_resultado_tea
from tea import analisar as analisar_tea


ALIASES = {
    "project": ("nome do projeto", "project name", "project title", "projeto", "project"),
    "investment": ("investimento inicial", "investimento total", "initial investment",
                   "initial investment required", "investment required", "total capital investment",
                   "total project investment", "initial capital requirement", "initial cost", "capital total", "capex"),
    "rate": ("taxa minima de atratividade", "taxa de desconto", "taxa minima",
             "discount rate", "minimum acceptable rate", "marr", "wacc", "hurdle rate"),
    "tax": ("taxa de imposto", "imposto sobre lucro", "tax rate", "income tax rate"),
    "horizon": ("vida util do projeto", "vida do projeto", "horizonte do projeto",
                "project life", "life-time project", "project horizon", "analysis period", "horizon"),
    "residual": ("valor residual", "salvage value", "residual value"),
    "benefit": ("receita anual", "receitas anuais", "beneficio anual", "beneficios anuais",
                "annual revenue", "annual revenues", "annual benefit", "annual benefits",
                "annual income", "annual savings"),
    "cost": ("custo anual", "custos anuais", "despesa anual", "despesas anuais",
             "annual cost", "annual costs", "operating cost", "operating costs", "operating expenses", "annual expenses"),
    "net_flow": ("fluxo de caixa liquido", "fluxo de caixa anual", "net cash flow",
                 "annual cash flow", "cash flow"),
}


def _norm(value):
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", " ", text.lower()).strip()


def _number(value):
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value or "").strip().replace(" ", "")
    if not text:
        return None
    match = re.search(r"[-+]?\d[\d.,]*", text)
    if not match:
        return None
    return numero(match.group())


def _alias_key(text):
    normalized = _norm(text)
    # Prefer the most specific match where labels overlap.
    for key, aliases in ALIASES.items():
        if key == "project":
            continue
        for alias in sorted(aliases, key=len, reverse=True):
            if normalized == _norm(alias):
                return key
    melhor, tamanho = None, 0
    for key, aliases in ALIASES.items():
        for alias in aliases:
            normalized_alias = _norm(alias)
            if len(normalized_alias) > 5 and normalized_alias in normalized and len(normalized_alias) > tamanho:
                melhor, tamanho = key, len(normalized_alias)
    return melhor


def _has_any(text, phrases):
    return any(_norm(phrase) in text for phrase in phrases)


def _formulas_sem_valor(arquivo):
    """True se há células com fórmula cujo resultado não foi salvo no arquivo."""
    try:
        formulas = load_workbook(arquivo, data_only=False, read_only=True)
        arquivo.seek(0)
        valores = load_workbook(arquivo, data_only=True, read_only=True)
        for ws in formulas.worksheets:
            wv = valores[ws.title]
            for linha_f, linha_v in zip(ws.iter_rows(), wv.iter_rows()):
                for cf, cv in zip(linha_f, linha_v):
                    if isinstance(cf.value, str) and cf.value.startswith("=") and cv.value is None:
                        return True
    except Exception:
        return False
    return False


def _legacy_result(arquivo, modo_evte="formulas"):
    """Converte os formatos já aceitos pelo site para a saída genérica."""
    arquivo.seek(0)
    globais, cenarios = ler_planilha_tea(arquivo, modo_evte)
    arquivo.seek(0)
    conferencia = globais.pop("_conferencia", None)
    resultado = converter_resultado_tea(analisar_tea(globais, cenarios))
    if conferencia:
        resultado["conferencia"] = _montar_conferencia(conferencia, resultado)
    return resultado


def _montar_conferencia(conf, resultado):
    """Compara o resultado do site com o do Excel e com o cálculo pela tabela de premissas."""
    outro = "premissas" if conf["modo"] == "formulas" else "formulas"
    g_alt, c_alt = conf[outro]
    alt = converter_resultado_tea(analisar_tea(dict(g_alt), c_alt))
    linhas = []
    for i, c in enumerate(resultado["cenarios"]):
        excel = conf["excel"][i] if i < len(conf["excel"]) else {}
        v, va = c["viabilidade"], alt["cenarios"][i]["viabilidade"]
        linhas.append({
            "nome": c["nome"],
            "excel_vpl": excel.get("vpl"), "excel_tir": excel.get("tir"), "excel_payback": excel.get("payback"),
            "site_vpl": v["vpl"], "site_tir": v["tir"], "site_viavel": v["viavel"],
            "alt_vpl": va["vpl"], "alt_tir": va["tir"], "alt_viavel": va["viavel"],
        })
    return {
        "modo": conf["modo"],
        "modo_alternativo": outro,
        "linhas": linhas,
        "divergencias": conf["divergencias"],
    }


def ler_planilha_generica(arquivo, modo_evte="formulas"):
    """Lê a planilha do site, a planilha EVTE ou tabelas financeiras rotuladas."""
    try:
        return _legacy_result(arquivo, modo_evte)
    except (ValueError, KeyError, TypeError, AttributeError):
        arquivo.seek(0)

    wb = load_workbook(arquivo, data_only=True, read_only=True)
    sheets = []
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        sheets.append((ws.title, rows))

    assumptions = {"project": "Projeto", "investment": None, "rate": None, "tax": 0,
                   "horizon": None, "residual": 0.0, "benefit": None, "cost": None}
    line_items = defaultdict(lambda: {"beneficios": [], "custos": []})
    flows_by_scenario = defaultdict(dict)
    inputs_by_scenario = defaultdict(dict)
    scenario_names = set()
    detected_rows = 0

    for title, rows in sheets:
        scenario_columns = {}
        for header_i, row in enumerate(rows[:15]):
            cells = list(row or ())
            marker = next((i for i, value in enumerate(cells)
                           if isinstance(value, str) and _has_any(_norm(value), ("cenario", "scenario"))), None)
            if marker is None:
                continue
            candidates = [cells]
            candidates.extend(list(x or ()) for x in rows[header_i + 1:header_i + 4])
            names_row = max(candidates, key=lambda candidate: sum(
                1 for value in candidate[marker:] if isinstance(value, str) and _norm(value)
                and not _has_any(_norm(value), ("cenario", "scenario"))))
            for col in range(marker + 1, len(names_row)):
                name = names_row[col]
                if isinstance(name, str) and _norm(name):
                    scenario_columns[col] = name.strip()
            if len(scenario_columns) >= 2:
                break

        for row in rows:
            cells = list(row or ())
            if not cells:
                continue
            # Lê pares rótulo/valor em qualquer aba, sem depender do nome da aba.
            valor_consumido_ate = -1  # o valor de um rótulo não pode ser lido como outro rótulo
            for i, cell in enumerate(cells):
                if i <= valor_consumido_ate or not isinstance(cell, str):
                    continue
                if len(_norm(cell).split()) > 10:  # rótulos são curtos; textos longos são observações
                    continue
                key = _alias_key(cell)
                if not key:
                    continue
                for col, scenario_name in scenario_columns.items():
                    if col < len(cells):
                        scenario_value = cells[col]
                        if key == "project":
                            if scenario_value is not None and str(scenario_value).strip():
                                inputs_by_scenario[scenario_name][key] = str(scenario_value).strip()
                        else:
                            parsed_scenario = _number(scenario_value)
                            if parsed_scenario is not None:
                                if key in ("rate", "tax") and 0 < parsed_scenario <= 1:
                                    parsed_scenario *= 100
                                inputs_by_scenario[scenario_name][key] = parsed_scenario
                if key == "project":
                    pos = next((j for j in range(i + 1, len(cells)) if cells[j] is not None and str(cells[j]).strip()), None)
                else:
                    pos = next((j for j in range(i + 1, len(cells)) if _number(cells[j]) is not None), None)
                if pos is None:
                    continue
                value = cells[pos]
                valor_consumido_ate = pos
                if key == "project":
                    assumptions[key] = str(value).strip()
                    detected_rows += 1
                elif key in ("investment", "rate", "tax", "horizon", "residual", "benefit", "cost"):
                    parsed = _number(value)
                    if parsed is not None:
                        if key in ("rate", "tax") and 0 < parsed <= 1:
                            parsed *= 100
                        assumptions[key] = parsed
                        detected_rows += 1

        # Tabela longa: identifica colunas de ano + fluxo líquido ou categoria + valor.
        for header_i, row in enumerate(rows):
            headers = [_norm(x) for x in (row or ())]
            if not headers:
                continue
            col_year = next((i for i, h in enumerate(headers) if h in ("ano", "year", "periodo", "period") or _has_any(h, ("project year", "ano do projeto"))), None)
            col_flow = next((i for i, h in enumerate(headers) if _has_any(h, ("cash flow", "fluxo de caixa", "fluxo caixa"))), None)
            col_scenario = next((i for i, h in enumerate(headers) if _has_any(h, ("cenario", "scenario", "alternativa", "alternative"))), None)
            if col_year is not None and col_flow is not None:
                for data_row in rows[header_i + 1:]:
                    if not data_row or max(col_year, col_flow) >= len(data_row):
                        continue
                    year = _number(data_row[col_year])
                    amount = _number(data_row[col_flow])
                    if year is None or amount is None or year < 1:
                        continue
                    scenario = str(data_row[col_scenario]).strip() if col_scenario is not None and col_scenario < len(data_row) and data_row[col_scenario] else "Projeto"
                    flows_by_scenario[scenario][int(year)] = amount
                    scenario_names.add(scenario)

            col_kind = next((i for i, h in enumerate(headers) if _has_any(h, ("tipo", "type", "categoria", "category", "natureza"))), None)
            col_amount = next((i for i, h in enumerate(headers) if _has_any(h, ("valor", "value", "amount", "total anual", "annual amount"))), None)
            col_name = next((i for i, h in enumerate(headers) if _has_any(h, ("item", "descricao", "description", "nome", "name"))), None)
            if col_kind is not None and col_amount is not None:
                for data_row in rows[header_i + 1:]:
                    if max(col_kind, col_amount) >= len(data_row or ()):
                        continue
                    kind = _norm(data_row[col_kind])
                    amount = _number(data_row[col_amount])
                    if amount is None:
                        continue
                    name = str(data_row[col_name]).strip() if col_name is not None and col_name < len(data_row) and data_row[col_name] else (str(data_row[col_kind]).strip() or "Item")
                    bucket = "beneficios" if any(x in kind for x in ("receita", "beneficio", "revenue", "income", "benefit", "entrada")) else "custos" if any(x in kind for x in ("custo", "despesa", "cost", "expense", "saida")) else None
                    if bucket:
                        line_items["Projeto"][bucket].append({"nome": name, "valor": amount})
                        detected_rows += 1

    cenarios = []
    names = sorted(scenario_names or inputs_by_scenario.keys()) if (scenario_names or inputs_by_scenario) else [str(assumptions["project"] or "Projeto")]
    for name in names:
        scenario_inputs = inputs_by_scenario.get(name, {})
        flows = flows_by_scenario.get(name)
        benefits = list(line_items[name]["beneficios"] or line_items["Projeto"]["beneficios"])
        costs = list(line_items[name]["custos"] or line_items["Projeto"]["custos"])
        benefit = scenario_inputs.get("benefit", assumptions["benefit"])
        cost = scenario_inputs.get("cost", assumptions["cost"])
        # Valor digitado na coluna do próprio cenário vale mais que a lista de itens compartilhada.
        if "benefit" in scenario_inputs or (not benefits and benefit is not None):
            benefits = [{"nome": "Receitas e benefícios anuais", "valor": benefit}]
        if "cost" in scenario_inputs or (not costs and cost is not None):
            costs = [{"nome": "Custos e despesas anuais", "valor": cost}]
        project_years = scenario_inputs.get("horizon", assumptions["horizon"])
        if flows:
            last_year = max(flows)
            cash_flows = [flows.get(year, 0.0) for year in range(1, last_year + 1)]
        else:
            # Sem tabela de fluxo: o site monta o fluxo (e aplica imposto/depreciação) a partir de
            # receitas, custos e horizonte. Antes o fluxo já saía pronto e o imposto lido era ignorado.
            cash_flows = None
        nome_cenario = scenario_inputs.get("project")
        if not nome_cenario:
            nome_cenario = assumptions["project"] if (len(names) == 1 and assumptions["project"] not in (None, "Projeto")) else name
        cenarios.append({
            "nome": str(nome_cenario),
            "anos": int(round(project_years)) if project_years else None,
            "investimento_inicial": scenario_inputs.get("investment", assumptions["investment"]),
            "valor_residual": scenario_inputs.get("residual", assumptions["residual"]) or 0,
            "beneficios": benefits,
            "custos": costs,
            "fluxos_caixa": cash_flows,
        })

    missing = []
    if any(c["investimento_inicial"] is None for c in cenarios):
        missing.append("investimento inicial")
    if assumptions["rate"] is None and not all("rate" in x for x in inputs_by_scenario.values()):
        missing.append("taxa de desconto/MARR")
    if assumptions["horizon"] is None and not all("horizon" in x for x in inputs_by_scenario.values()) and not any(c["fluxos_caixa"] for c in cenarios):
        missing.append("vida útil do projeto ou anos de fluxo de caixa")
    has_annual_values = any(c["beneficios"] or c["custos"] for c in cenarios)
    has_flows = any(c["fluxos_caixa"] for c in cenarios)
    if not has_annual_values and not has_flows:
        missing.append("receitas/benefícios e custos anuais, ou uma tabela de fluxo de caixa líquido")
    if missing or detected_rows == 0:
        arquivo.seek(0)
        if _formulas_sem_valor(arquivo):
            raise ValueError(
                "Esta planilha tem fórmulas sem valor calculado salvo (provavelmente foi gerada por um programa). "
                "Abra no Excel, salve o arquivo e envie novamente."
            )
        detail = ", ".join(missing) if missing else "dados reconhecíveis"
        raise ValueError(
            "Não consegui identificar todos os dados necessários nesta planilha: " + detail + ". "
            "Use rótulos como investimento inicial, taxa de desconto/MARR, vida útil, receitas/benefícios e custos anuais; "
            "ou uma tabela com Ano e Fluxo de caixa líquido."
        )

    horizon = assumptions["horizon"] or max(len(c["fluxos_caixa"] or []) for c in cenarios)
    rate = assumptions["rate"]
    if rate is None:
        rate = next(iter(inputs_by_scenario.values()))["rate"]
    return analisar_projeto({"taxa_desconto": rate, "vida_projeto": horizon,
                             "taxa_imposto": assumptions["tax"]}, cenarios)
