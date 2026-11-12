import math
import re

from main import app
from flask import render_template, request, redirect, url_for, flash

from analysis import numero
from campos import (GLOBAIS_GRUPOS, CENARIO_GRUPOS, EQUIP_CAMPOS, EXEMPLO_GLOBAIS, EXEMPLO_CENARIOS,
                    GRUPOS_AVANCADOS)
from dashboard import montar_dashboard
from generic_excel import ler_planilha_generica
from project_analysis import analisar_projeto, converter_resultado_tea
from tea import fmt_num, analisar as analisar_tea

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

    return _mostrar_resultado(analisar_projeto(globais, cenarios))


# ---------- Inserção manual: processo industrial (equipamentos, insumos, critérios) ----------

OBRIGATORIOS_PROCESSO = ("marr", "vida_projeto")


@app.route('/analise-manual/processo')
def analise_manual_processo():
    exemplo = request.args.get('exemplo') == '1'
    # Premissas avançadas já vêm preenchidas com os valores-padrão; o resto só no exemplo.
    avancadas = {chave for titulo, campos in GLOBAIS_GRUPOS if titulo in GRUPOS_AVANCADOS for chave, _, _ in campos}
    g = {k: v for k, v in EXEMPLO_GLOBAIS.items() if exemplo or k in avancadas}
    cens = EXEMPLO_CENARIOS if exemplo else [{"nome": "", "equipamentos": [{}]}]
    return render_template('formulario_processo.html', grupos=GLOBAIS_GRUPOS, cen_grupos=CENARIO_GRUPOS,
                           equip=EQUIP_CAMPOS, avancados=GRUPOS_AVANCADOS, g=g, cens=cens,
                           obrigatorios=OBRIGATORIOS_PROCESSO)


@app.route('/analise-manual/processo/resultado', methods=['POST'])
def analise_manual_processo_resultado():
    globais, cenarios = _extrair_processo(request.form)
    erro = _validar_processo(globais, cenarios)
    if erro:
        flash(erro)
        return redirect(url_for('analise_manual_processo'))
    return _mostrar_resultado(converter_resultado_tea(analisar_tea(globais, cenarios)))


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

    modo_evte = request.form.get('modo_evte', 'formulas')
    if modo_evte not in ('formulas', 'premissas'):
        modo_evte = 'formulas'

    try:
        resultado = ler_planilha_generica(arquivo, modo_evte)
    except ValueError as e:
        flash(str(e))
        return redirect(url_for('importar_planilha'))
    except Exception:
        flash('Não foi possível ler a planilha. Confira se ela está no formato .xlsx e use uma das planilhas compatíveis ou a planilha-modelo do site.')
        return redirect(url_for('importar_planilha'))

    return _mostrar_resultado(resultado)


# ---------- Funções auxiliares ----------

def _mostrar_resultado(resultado):
    """Página única de resultado: veredito, dashboard interativo e tabelas."""
    return render_template('resultado_generico.html', r=resultado, dash=montar_dashboard(resultado))


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
        cap_lives = form.getlist(f"c{i}_capital_vida")
        cap_lives += [""] * (len(cap_names) - len(cap_lives))
        c["investimentos"] = [
            {"nome": name.strip() or "Investimento", "quantidade": quantity, "valor_unitario": unit, "vida_util": life}
            for name, quantity, unit, life in zip(cap_names, cap_quantities, cap_units, cap_lives)
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
            vida = item.get("vida_util")
            if str(vida if vida is not None else "").strip() and (not _numero_valido(vida) or numero(vida) <= 0):
                return f"A vida útil do item de investimento {j} no cenário {i} deve ser um número maior que zero (ou deixe em branco)."
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


def _extrair_processo(form):
    """Lê o formulário de processo industrial: g_<chave> (globais), c<i>_<chave> e c<i>_eq_<campo>."""
    globais = {chave: form.get(f"g_{chave}", "").strip()
               for _, campos in GLOBAIS_GRUPOS for chave, _, _ in campos}
    indices = sorted({int(m.group(1)) for k in form.keys() if (m := re.match(r'c(\d+)_nome$', k))})
    cenarios = []
    for i in indices:
        c = {"nome": form.get(f"c{i}_nome", "").strip()}
        for _, campos in CENARIO_GRUPOS:
            for chave, _, _ in campos:
                c[chave] = form.get(f"c{i}_{chave}", "").strip()
        listas = {campo: form.getlist(f"c{i}_eq_{campo}") for campo, _, _ in EQUIP_CAMPOS}
        n = max((len(v) for v in listas.values()), default=0)
        equipamentos = []
        for j in range(n):
            linha = {campo: (listas[campo][j].strip() if j < len(listas[campo]) else "") for campo in listas}
            if any(linha[campo] for campo in linha if campo != "ajustar_cepci"):
                equipamentos.append(linha)
        c["equipamentos"] = equipamentos
        if c["nome"] or equipamentos or c.get("capacidade"):
            cenarios.append(c)
    return globais, cenarios


def _validar_processo(globais, cenarios):
    if not cenarios:
        return "Informe pelo menos um cenário."
    for chave in OBRIGATORIOS_PROCESSO:
        if not _numero_valido(globais.get(chave)) or numero(globais.get(chave)) <= 0:
            rotulo = {"marr": "a taxa mínima de atratividade (MARR)", "vida_projeto": "a vida útil do projeto"}[chave]
            return f"Informe {rotulo} (número maior que zero)."
    if numero(globais.get("vida_projeto")) % 1:
        return "Informe a vida útil do projeto em anos inteiros."
    ano0 = globais.get("invest_ano0_pct", "")
    if ano0 and (not _numero_valido(ano0) or not 0 < numero(ano0) <= 100):
        return "A parcela do investimento paga no ano 0 deve estar entre 0 e 100%."
    for chave, valor in globais.items():
        if valor and not _numero_valido(valor):
            return f"Confira o valor digitado em “{chave}”: use apenas números."
        if valor and numero(valor) < 0:
            return f"O valor de “{chave}” não pode ser negativo."
    for i, c in enumerate(cenarios, 1):
        for chave, valor in c.items():
            if chave in ("nome", "equipamentos"):
                continue
            if valor and (not _numero_valido(valor) or numero(valor) < 0):
                return f"Confira o campo “{chave}” do cenário {i}: use apenas números não negativos."
        if numero(c.get("capacidade")) <= 0:
            return f"Informe a matéria-prima processada (t/ano) do cenário {i}."
        validos = 0
        for j, e in enumerate(c["equipamentos"], 1):
            for campo in ("quantidade", "custo_unidade", "capacidade", "consumo_kwh_ano"):
                if e.get(campo) and (not _numero_valido(e[campo]) or numero(e[campo]) < 0):
                    return f"Confira o campo “{campo}” do equipamento {j} no cenário {i}."
            if numero(e.get("quantidade")) > 0 and numero(e.get("custo_unidade")) > 0:
                validos += 1
        if validos == 0 and numero(c.get("fci")) <= 0:
            return f"Informe ao menos um equipamento com quantidade e custo (ou o FCI) no cenário {i}."
    return None
