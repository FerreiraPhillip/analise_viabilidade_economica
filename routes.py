from main import app
from flask import render_template, request, redirect, url_for, flash

from analysis import (
    calcular_projeto,
    comparar_alternativas,
    calcular_cenarios,
    analise_sensibilidade,
)
from excel_utils import ler_planilha_projeto


@app.route('/')
def homepage():
    return render_template('homepage.html')


# ---------- Inserção manual ----------

@app.route('/analise-manual')
def analise_manual():
    return render_template('formulario.html')


@app.route('/analise-manual/resultado', methods=['POST'])
def analise_manual_resultado():
    dados = _extrair_dados_formulario(request.form)
    alternativas = _extrair_alternativas_formulario(request.form)

    resultado = _processar(dados, alternativas)
    return render_template('resultado.html', **resultado)


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
        dados, alternativas = ler_planilha_projeto(arquivo)
    except Exception:
        flash('Não foi possível ler a planilha. Confira se ela segue o modelo esperado.')
        return redirect(url_for('importar_planilha'))

    resultado = _processar(dados, alternativas)
    return render_template('resultado.html', **resultado)


# ---------- Funções auxiliares ----------

def _processar(dados, alternativas):
    """Roda toda a análise de viabilidade para um conjunto de dados."""
    projeto = calcular_projeto(dados)

    projeto['nome'] = dados.get('nome', 'Projeto proposto')
    projeto['tempo'] = dados.get('tempo', 0)
    projeto['rendimento'] = dados.get('rendimento', 0)
    projeto['materiais'] = dados.get('materiais', [])
    projeto['equipamentos'] = dados.get('equipamentos', [])

    cenarios = calcular_cenarios(projeto['custo_total'])
    sensibilidade = analise_sensibilidade(projeto['custo_total'])
    comparacao = comparar_alternativas(projeto, alternativas)

    return {
        'projeto': projeto,
        'cenarios': cenarios,
        'sensibilidade': sensibilidade,
        'comparacao': comparacao,
    }


def _extrair_dados_formulario(form):
    """Monta o dicionário 'dados' a partir dos campos do formulário manual."""

    nomes = form.getlist('material_nome')
    quantidades = form.getlist('material_quantidade')
    precos = form.getlist('material_preco')
    materiais = [
        {'nome': n, 'quantidade': q, 'preco': p}
        for n, q, p in zip(nomes, quantidades, precos)
        if n or q or p
    ]

    eq_nomes = form.getlist('equipamento_nome')
    eq_horas = form.getlist('equipamento_horas')
    eq_custos = form.getlist('equipamento_custo_hora')
    equipamentos = [
        {'nome': n, 'horas': h, 'custo_hora': c}
        for n, h, c in zip(eq_nomes, eq_horas, eq_custos)
        if n or h or c
    ]

    experimentos = [
        v for v in form.getlist('experimento_valor') if v not in (None, '')
    ]

    return {
        'nome': form.get('nome_projeto') or 'Projeto proposto',
        'materiais': materiais,
        'equipamentos': equipamentos,
        'energia': form.get('energia', 0),
        'mao_obra': form.get('mao_obra', 0),
        'outros': form.get('outros', 0),
        'quantidade_produzida': form.get('quantidade_produzida', 0),
        'tempo': form.get('tempo', 0),
        'rendimento': form.get('rendimento', 0),
        'experimentos': experimentos,
    }


def _extrair_alternativas_formulario(form):
    nomes = form.getlist('alt_nome')
    custos = form.getlist('alt_custo')
    tempos = form.getlist('alt_tempo')
    rendimentos = form.getlist('alt_rendimento')

    return [
        {'nome': n, 'custo': c, 'tempo': t, 'rendimento': r}
        for n, c, t, r in zip(nomes, custos, tempos, rendimentos)
        if n or c or t or r
    ]