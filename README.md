Link para o uso do site:

https://analise-viabilidade-economica.onrender.com

## Entradas aceitas

- **Inserção manual — duas abas:** *Projeto geral* (itens de investimento, receitas e custos) e *Processo industrial* (equipamentos, insumos e critérios, como na planilha EVTE). No processo industrial o usuário digita só o que é medido/cotado; o site calcula razão CEPCI, custo e consumo dos equipamentos, rendimentos, investimento total (método dos fatores), cada linha de custo, depreciação, imposto, fluxo de caixa, VPL, TIR e payback.

- Preenchimento manual de projetos de qualquer área: itens de investimento, quantidades e valores unitários de receitas/benefícios e custos/despesas, horizonte e taxa de desconto. O site multiplica e soma os itens; se o investimento já estiver totalizado, basta cadastrá-lo como uma linha com quantidade 1.
- Imposto e ativos depreciáveis são opcionais. Se informados, o site calcula depreciação linear e fluxo de caixa após impostos; sem eles, o resultado é antes dos impostos.
- Upload de planilhas `.xlsx` com rótulos em português ou inglês. O importador procura os dados pelo conteúdo das células, sem exigir nomes fixos para as abas.
- Tabelas de itens anuais com tipo/categoria e valor, ou tabelas de fluxo líquido com ano e valor.
- Compatibilidade preservada com a planilha EVTE fornecida e com o modelo antigo do site.

Na análise manual, as quantidades e valores unitários são tratados como constantes ao longo do horizonte. Para variações por ano, importe uma tabela de fluxo de caixa líquido anual. O parecer é financeiro e depende dos dados e das premissas informadas; benefícios acadêmicos sem estimativa monetária não entram no VPL. Fórmulas do Excel precisam estar calculadas e salvas antes do upload.

## Resultado e dashboard

A página de resultado mostra o veredito (viável / não viável) por cenário e um dashboard interativo (Chart.js via CDN): indicadores, fluxo acumulado e descontado, fluxo anual, curva VPL × taxa, sensibilidade, composição dos custos, margem de segurança e um simulador "E se...?" (receita, custos, investimento e taxa). A simulação roda no navegador e não altera os dados enviados.

## Planilha EVTE

O importador procura os dados pelos **rótulos** (não por endereços fixos) e reproduz o VPL/TIR do Excel usando o que as fórmulas realmente aplicam. Como a tabela de premissas da planilha diverge das fórmulas em vários itens (mão de obra, imposto, mão de obra indireta, preços...), o resultado traz a conferência lado a lado e a importação permite escolher "fórmulas" ou "tabela de premissas".

## Modelos de planilha

`gerar_modelo.py` gera `static/modelo_planilha.xlsx` (processo industrial) e `gerar_modelo_generico.py` gera `static/modelo_planilha_generico.xlsx` (qualquer projeto). Os dois estão linkados na tela de importação.
