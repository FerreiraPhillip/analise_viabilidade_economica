import math
import re

from main import app
from flask import render_template, request, redirect, url_for, flash

from analysis import numero
from generic_excel import ler_planilha_generica
from project_analysis import analisar_projeto
from tea import fmt_num

app.add_template_filter(fmt_num, "num")

@app.route('/')
def homepage():
    return render_template('homepage.html')


# ---------- Inserção manual ----------

@app.route('/analise-manual')
def analise_manual():
    globais = {"taxa_desconto": "", "vida_projeto": "", "taxa_imposto": ""}
    cenario = {"nome": "", "investimento_inicial": "", "valor_residual": "",
               "investimentos": [{}],
               "beneficios": [{"nome": "", "valor": ""}],
               "custos": [{"nome": "", "valor": ""}], "ativos_depreciaveis": [{}]}
    return render_template('formulario.html', g=globais, cens=[cenario])


@app.route('/analise-manual/resultado', methods=['POST'])
def analise_manual_resultado():
    globais = {
        "taxa_desconto": request.form.get("taxa_desconto", ""),
        "vida_projeto": request.form.get("vida_projeto", ""),
        "taxa_imposto": request.form.get("taxa_imposto", ""),
    }
    cenarios = _extrair_projetos(request.form)

    erro = _validar_projetos(globais, cenarios)
    if erro:
        flash(erro)
        return redirect(url_for('analise_manual'))

    return render_template('resultado_generico.html', r=analisar_projeto(globais, cenarios))


# ---------- Importação de planilha ----------

@app.route('/importar-planilha')
def importar_planilha():
    return render_template('importar.html')


@app.route('/importar-planilha/resultado', methods=['POST'])
def importar_planilha_resultado():
    arquivo = request.files.get('planilha')

    if not arquivo or arquivo.filename == '':
        flash('Selecione um arquivo de planilha (.xlsx) para continuar.')
        return redirect(url_for('importar_planilha'))

    if not arquivo.filename.lower().endswith('.xlsx'):
        flash('O arquivo precisa estar no formato .xlsx.')
        return redirect(url_for('importar_planilha'))

    try:
        resultado = ler_planilha_generica(arquivo)
    except ValueError as e:
        flash(str(e))
        return redirect(url_for('importar_planilha'))
    except Exception:
        flash('Não foi possível ler a planilha. Confira se ela está no formato .xlsx e use uma das planilhas compatíveis ou a planilha-modelo do site.')
        return redirect(url_for('importar_planilha'))

    return render_template('resultado_generico.html', r=resultado)


# ---------- Funções auxiliares ----------

def _extrair_projetos(form):
    indices = sorted({
        int(m.group(1)) for k in form.keys() if (m := re.match(r'c(\d+)_nome$', k))
    })
    projetos = []
    for i in indices:
        c = {
            "nome": form.get(f"c{i}_nome", "").strip(),
            "investimento_inicial": form.get(f"c{i}_investimento", ""),
            "valor_residual": form.get(f"c{i}_residual", ""),
            "investimentos": [], "beneficios": [], "custos": [],
        }
        cap_names = form.getlist(f"c{i}_capital_nome")
        cap_quantities = form.getlist(f"c{i}_capital_quantidade")
        cap_units = form.getlist(f"c{i}_capital_unitario")
        c["investimentos"] = [
            {"nome": name.strip() or "Investimento", "quantidade": quantity, "valor_unitario": unit}
            for name, quantity, unit in zip(cap_names, cap_quantities, cap_units)
            if str(name).strip() or str(quantity).strip() or str(unit).strip()
        ]
        for kind in ("beneficio", "custo"):
            names = form.getlist(f"c{i}_{kind}_nome")
            quantities = form.getlist(f"c{i}_{kind}_quantidade")
            unit_values = form.getlist(f"c{i}_{kind}_unitario")
            items = []
            for name, quantity, unit_value in zip(names, quantities, unit_values):
                if not (str(name).strip() or str(quantity).strip() or str(unit_value).strip()):
                    continue
                items.append({
                    "nome": name.strip() or ("Benefício" if kind == "beneficio" else "Custo"),
                    "quantidade": quantity,
                    "valor_unitario": unit_value,
                })
            c["beneficios" if kind == "beneficio" else "custos"] = items
        nomes_ativos = form.getlist(f"c{i}_ativo_nome")
        bases_ativos = form.getlist(f"c{i}_ativo_base")
        vidas_ativos = form.getlist(f"c{i}_ativo_vida")
        c["ativos_depreciaveis"] = [
            {"nome": nome, "base": base, "vida_util": vida}
            for nome, base, vida in zip(nomes_ativos, bases_ativos, vidas_ativos)
            if str(nome).strip() or str(base).strip() or str(vida).strip()
        ]
        if c["nome"] or c["investimento_inicial"] or c["investimentos"] or c["beneficios"] or c["custos"]:
            projetos.append(c)
    return projetos


def _validar_projetos(globais, projetos):
    if not projetos:
        return "Informe pelo menos um projeto ou cenário."
    if not _numero_valido(globais.get("taxa_desconto")):
        return "Informe a taxa mínima de atratividade ou taxa de desconto."
    if not _numero_valido(globais.get("vida_projeto")) or numero(globais.get("vida_projeto")) <= 0:
        return "Informe um horizonte de análise maior que zero."
    if numero(globais.get("vida_projeto")) % 1:
        return "Informe o horizonte em anos inteiros."
    if str(globais.get("taxa_imposto", "")).strip():
        if not _numero_valido(globais.get("taxa_imposto")) or not 0 <= numero(globais.get("taxa_imposto")) <= 100:
            return "A taxa de imposto opcional deve estar entre 0 e 100%."
    for i, p in enumerate(projetos, 1):
        for j, item in enumerate(p.get("investimentos", []), 1):
            if not _numero_valido(item.get("quantidade")) or not _numero_valido(item.get("valor_unitario")):
                return f"Preencha quantidade e valor unitário do item de investimento {j} no cenário {i}."
            if numero(item["quantidade"]) < 0 or numero(item["valor_unitario"]) < 0:
                return f"Quantidade e valor unitário do investimento não podem ser negativos (cenário {i})."
        total_investido = sum(numero(x.get("quantidade")) * numero(x.get("valor_unitario")) for x in p.get("investimentos", []))
        investimento_direto = p.get("investimento_inicial")
        if total_investido <= 0 and (not _numero_valido(investimento_direto) or numero(investimento_direto) <= 0):
            return f"Informe ao menos um item de investimento inicial no cenário {i}."
        if str(p.get("valor_residual", "")).strip() and not _numero_valido(p.get("valor_residual")):
            return f"Confira o valor residual informado no cenário {i}."
        if numero(p.get("valor_residual")) < 0:
            return f"O valor residual não pode ser negativo (cenário {i})."
        for j, ativo in enumerate(p.get("ativos_depreciaveis", []), 1):
            if not _numero_valido(ativo.get("base")) or not _numero_valido(ativo.get("vida_util")):
                return f"Preencha a base depreciável e a vida útil do ativo {j} no cenário {i}, ou remova a linha."
            if numero(ativo["base"]) < 0 or numero(ativo["vida_util"]) <= 0:
                return f"A base depreciável não pode ser negativa e a vida útil deve ser maior que zero (cenário {i})."
        if not p["beneficios"] and not p["custos"]:
            return f"Informe ao menos um benefício/receita ou custo/despesa anual no cenário {i}."
        for tipo, itens in (("benefício/receita", p["beneficios"]), ("custo/despesa", p["custos"])):
            for j, item in enumerate(itens, 1):
                if not _numero_valido(item.get("quantidade")) or not _numero_valido(item.get("valor_unitario")):
                    return f"Preencha quantidade e valor unitário do item {j} de {tipo} no cenário {i}."
                if numero(item["quantidade"]) < 0 or numero(item["valor_unitario"]) < 0:
                    return f"Quantidade e valor unitário não podem ser negativos (cenário {i})."
    return None


def _numero_valido(valor):
    texto = str(valor if valor is not None else "").strip().replace(" ", "")
    if not texto:
        return False
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".") if texto.rfind(",") > texto.rfind(".") else texto.replace(",", "")
    else:
        texto = texto.replace(",", ".")
    try:
        return math.isfinite(float(texto))
    except ValueError:
        return False
