import streamlit as st

st.set_page_config(
    page_title="Comitê Técnico - Excel & Power Query",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
.block-container {
    padding-top: 1rem !important;
    padding-left: 1rem !important;
    padding-right: 1rem !important;
    max-width: 100% !important;
}
div[data-testid="stVerticalBlock"] {
    gap: 0.25rem !important;
}
div.stButton > button {
    width: 100% !important;
    min-height: 44px !important;
    white-space: normal !important;
    font-size: 11px !important;
    line-height: 1.2 !important;
}
.card {
    background-color: #ffffff;
    border: 1px solid #dcdcdc;
    border-left: 5px solid #1f77b4;
    border-radius: 6px;
    padding: 12px;
    margin-bottom: 8px;
}
.card-green {
    border-left-color: #18a558;
    background-color: #eefaf1;
}
.card-yellow {
    border-left-color: #f0ad4e;
    background-color: #fff8e6;
}
.card-red {
    border-left-color: #d9534f;
    background-color: #fff0f0;
}
.code-box {
    background-color: #f7f7f7;
    border: 1px solid #ddd;
    border-radius: 6px;
    padding: 10px;
    font-family: monospace;
    font-size: 12px;
    white-space: pre-wrap;
}
.category-header {
    font-size: 12px;
    font-weight: bold;
    text-transform: uppercase;
    border-bottom: 2px solid #ddd;
    margin-bottom: 6px;
}
</style>
""", unsafe_allow_html=True)


DATA_MAPPING = {
    1: {
        "category": "NÍVEL 1 - Fundamentos",
        "title": "Excel como base analítica",
        "tag": "EXCEL-BASE",
        "objective": "Entender quando o Excel é suficiente e como organizar uma base para análise.",
        "concept": "Excel é adequado para bases moderadas, análises rápidas, controles operacionais e protótipos.",
        "when_use": "Use Excel quando há poucos usuários, baixa complexidade, baixo volume e necessidade de resposta rápida.",
        "attention": "Evite usar Excel como sistema oficial quando houver múltiplos usuários editando, dados sensíveis, histórico crítico ou alto volume.",
        "formulas": [
            '=SE(A2="";"Sem informação";A2)',
            '=ARRUMAR(A2)',
            '=MAIÚSCULA(A2)',
            '=PROCX(A2;TabelaProdutos[Produto];TabelaProdutos[Categoria])',
            '=SOMASES(TabelaVendas[Valor];TabelaVendas[Produto];A2)'
        ],
        "powerquery": [
            'Text.Trim([Produto])',
            'Text.Clean([Produto])',
            'Text.Upper([Produto])'
        ],
        "exercise": "Crie uma tabela com Produto, Cliente, Data e Valor. Padronize os textos, remova vazios e crie uma coluna de classificação por valor.",
        "qa": [
            {
                "q": "Quando Excel é suficiente?",
                "a": "Quando o volume é controlado, a lógica é simples ou média, há poucos usuários e o risco operacional é baixo."
            },
            {
                "q": "Quando devo migrar para Power Query?",
                "a": "Quando há repetição, bases sujas, múltiplas fontes ou necessidade de atualização automática."
            }
        ]
    },

    2: {
        "category": "NÍVEL 2 - Power Query",
        "title": "Limpeza e padronização de textos",
        "tag": "PQ-CLEAN",
        "objective": "Aprender a limpar campos textuais e corrigir inconsistências comuns.",
        "concept": "Power Query usa linguagem M, não DAX. DAX é usado em Power Pivot e Power BI para medidas analíticas.",
        "when_use": "Use Power Query para importar, limpar, transformar, combinar e automatizar bases.",
        "attention": "Sempre preserve a coluna original antes de aplicar correções, para garantir rastreabilidade.",
        "formulas": [
            '=ARRUMAR(A2)',
            '=MAIÚSCULA(A2)',
            '=SUBSTITUIR(A2;"  ";" ")'
        ],
        "powerquery": [
            'Text.Trim(Text.Clean([Produto]))',
            'Text.Upper(Text.Trim(Text.Clean([Produto])))',
            'Text.Proper(Text.Trim([Cliente]))'
        ],
        "exercise": "Receba uma base com nomes escritos como ' skol ', 'Skoll', 'BRAHMAA' e crie uma coluna Produto_Limpo.",
        "qa": [
            {
                "q": "Por que usar Text.Trim e Text.Clean juntos?",
                "a": "Text.Trim remove espaços externos. Text.Clean remove caracteres invisíveis que podem causar erro em cruzamentos."
            },
            {
                "q": "Qual é a boa prática antes de limpar uma coluna?",
                "a": "Duplicar a coluna original para manter auditoria."
            }
        ]
    },

    3: {
        "category": "NÍVEL 2 - Power Query",
        "title": "Remoção de grafias incorretas",
        "tag": "PQ-GRAFIAS",
        "objective": "Criar uma lógica reutilizável para corrigir nomes de produtos, clientes, cidades ou fornecedores.",
        "concept": "A melhor solução é combinar limpeza automática com uma tabela De/Para.",
        "when_use": "Use quando o mesmo item aparece com várias grafias diferentes.",
        "attention": "Evite corrigir tudo manualmente dentro do código. Prefira tabela de referência para manutenção.",
        "formulas": [
            '=PROCX(A2;DePara[Grafia_Incorreta];DePara[Nome_Correto];A2)',
            '=MAIÚSCULA(ARRUMAR(A2))'
        ],
        "powerquery": [
            'Text.Upper(Text.Trim(Text.Clean([Produto])))',
            'if [Produto_Correto] <> null then [Produto_Correto] else [Produto_Limpo]',
            'Table.NestedJoin(BaseVendas, {"Produto_Limpo"}, DeParaProdutos, {"Grafia_Incorreta"}, "Correcoes", JoinKind.LeftOuter)'
        ],
        "exercise": "Crie uma tabela DeParaProdutos com SKOLL -> SKOL, BRAHMAA -> BRAHMA e GUARANA ANTARTICA -> GUARANA ANTARCTICA.",
        "qa": [
            {
                "q": "Como corrigir grafias incorretas de forma profissional?",
                "a": "Criando uma coluna limpa, fazendo merge com uma tabela De/Para e gerando uma coluna final corrigida."
            }
        ]
    },

    4: {
        "category": "NÍVEL 3 - Data Quality",
        "title": "Validação de qualidade dos dados",
        "tag": "DQ-CHECK",
        "objective": "Identificar nulos, duplicidades, tipos incorretos e erros de preenchimento.",
        "concept": "Data Quality garante que dashboards e análises não propaguem erros operacionais.",
        "when_use": "Use em toda base que será fonte de indicadores, relatórios ou tomada de decisão.",
        "attention": "Um dashboard automatizado com dado errado apenas espalha erro mais rápido.",
        "formulas": [
            '=CONT.SE(A:A;A2)',
            '=SE(CONT.SE(A:A;A2)>1;"Duplicado";"Único")',
            '=SE(ÉCÉL.VAZIA(A2);"Nulo";"OK")'
        ],
        "powerquery": [
            'Table.SelectRows(Fonte, each [Produto] = null)',
            'Table.Distinct(Fonte)',
            'Table.RowCount(Fonte)'
        ],
        "exercise": "Crie uma rotina para contar linhas, remover duplicados e listar produtos sem preenchimento.",
        "qa": [
            {
                "q": "O que você checaria antes de entregar um relatório?",
                "a": "Nulos, duplicidades, tipos de dados, chaves inválidas, volume esperado e consistência com a regra de negócio."
            }
        ]
    },

    5: {
        "category": "NÍVEL 4 - Modelagem",
        "title": "Modelo analítico e Power Pivot",
        "tag": "MODELAGEM",
        "objective": "Entender fato, dimensão, relacionamento e uso de DAX.",
        "concept": "Power Query prepara dados. Power Pivot modela. DAX cria medidas analíticas.",
        "when_use": "Use Power Pivot/DAX quando houver relacionamento entre tabelas e necessidade de KPIs.",
        "attention": "Não confunda M com DAX. M transforma dados; DAX calcula métricas no modelo.",
        "formulas": [
            '=SOMASES(Vendas[Valor];Vendas[Produto];A2)',
            '=CONT.SES(Vendas[Cliente];A2;Vendas[Status];"Ativo")'
        ],
        "powerquery": [
            'Table.TransformColumnTypes(Fonte, {{"Data", type date}, {"Valor", type number}})',
            'Table.Group(Fonte, {"Produto"}, {{"Total", each List.Sum([Valor]), type number}})'
        ],
        "exercise": "Monte uma tabela fato de vendas e dimensões de produto, cliente e calendário.",
        "qa": [
            {
                "q": "Power Query usa DAX?",
                "a": "Não. Power Query usa linguagem M. DAX é usado no Power Pivot e Power BI para medidas."
            }
        ]
    },

    6: {
        "category": "NÍVEL 5 - Automação",
        "title": "Automação de planilhas",
        "tag": "AUTOMACAO",
        "objective": "Reduzir atualização manual e criar processo replicável.",
        "concept": "Automação começa com uma base bem estruturada e consultas reutilizáveis.",
        "when_use": "Use quando há rotina recorrente, arquivos mensais, consolidação manual ou retrabalho.",
        "attention": "Automatizar sem padronizar a origem pode gerar falhas recorrentes.",
        "formulas": [
            '=HOJE()',
            '=MÊS(A2)',
            '=ANO(A2)'
        ],
        "powerquery": [
            'Folder.Files("C:\\\\Bases\\\\Vendas")',
            'Table.Combine(Arquivos[Transformado])',
            'Table.SelectRows(Fonte, each [Valor] > 0)'
        ],
        "exercise": "Automatize a importação de todos os arquivos de uma pasta e consolide em uma única base.",
        "qa": [
            {
                "q": "Qual é o ganho da automação?",
                "a": "Menos retrabalho, menos erro manual, maior rastreabilidade e atualização mais rápida."
            }
        ]
    },

    7: {
        "category": "NÍVEL 6 - Teste Técnico",
        "title": "Simulação de case técnico",
        "tag": "CASE",
        "objective": "Treinar raciocínio técnico para prova prática ou entrevista.",
        "concept": "Um bom teste técnico avalia limpeza, modelagem, indicador e explicação do raciocínio.",
        "when_use": "Use este roteiro para treinar cases de Excel, Power Query, BI e Data Analyst.",
        "attention": "Não basta entregar resultado. É preciso explicar método, risco e validação.",
        "formulas": [
            '=SOMASES(Base[Valor];Base[Produto];A2)',
            '=SE(B2>=1000;"Alta";SE(B2>=500;"Média";"Baixa"))'
        ],
        "powerquery": [
            'if [Valor] >= 1000 then "Alta" else if [Valor] >= 500 then "Média" else "Baixa"',
            'Duration.Days([DataFim] - [DataInicio])',
            '[Cliente] & "-" & Text.From([Pedido])'
        ],
        "exercise": "Receba uma base suja, corrija grafias, crie indicadores de venda e entregue um resumo executivo.",
        "qa": [
            {
                "q": "Como você explicaria sua solução em uma entrevista?",
                "a": "Eu preservei a base original, criei uma camada de limpeza no Power Query, usei tabela De/Para para correções e gerei indicadores validados."
            }
        ]
    }
}


def render_code_list(title, items):
    st.markdown(f"**{title}**")
    for item in items:
        st.code(item, language="powerquery" if "Text." in item or "Table." in item or "if " in item else "excel")


if "active_id" not in st.session_state:
    st.session_state.active_id = 1

with st.sidebar:
    st.markdown("## Comitê Técnico")
    st.caption("Excel · Power Query · Data Quality · Automação · Teste Técnico")

    st.markdown("---")
    st.markdown("### Mapa Mental")
    st.markdown("""
    - Fundamentos de Excel  
    - Power Query e linguagem M  
    - Limpeza de grafias incorretas  
    - Data Quality  
    - Modelagem e DAX  
    - Automação  
    - Simulação de teste técnico  
    """)

    st.markdown("---")
    st.metric("Foco", "Aprovação", "Teste Técnico")
    st.metric("Stack", "Excel + PQ", "Data Analyst")

categories = list(dict.fromkeys([v["category"] for v in DATA_MAPPING.values()]))

cols = st.columns(len(categories))

for idx, category in enumerate(categories):
    with cols[idx]:
        st.markdown(f'<div class="category-header">{category}</div>', unsafe_allow_html=True)
        for item_id, item in DATA_MAPPING.items():
            if item["category"] == category:
                label = f"{item['tag']} · {item['title']}"
                if st.button(label, key=f"btn_{item_id}"):
                    st.session_state.active_id = item_id

active = DATA_MAPPING[st.session_state.active_id]

st.markdown("---")

left, right = st.columns([0.52, 0.48])

with left:
    st.markdown(f"""
    <div class="card card-green">
        <h3>{active['title']}</h3>
        <p><strong>Objetivo:</strong> {active['objective']}</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="card">
        <p><strong>Conceito essencial:</strong></p>
        <p>{active['concept']}</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="card card-yellow">
        <p><strong>Quando usar:</strong></p>
        <p>{active['when_use']}</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="card card-red">
        <p><strong>Ponto de atenção:</strong></p>
        <p>{active['attention']}</p>
    </div>
    """, unsafe_allow_html=True)

with right:
    render_code_list("Fórmulas Excel", active["formulas"])
    render_code_list("Códigos Power Query / M", active["powerquery"])

st.markdown("---")

col1, col2 = st.columns([0.5, 0.5])

with col1:
    st.markdown("## Exercício prático")
    st.markdown(f"""
    <div class="card">
        {active['exercise']}
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("## Perguntas de entrevista")
    for qa in active["qa"]:
        st.markdown(f"""
        <div class="card">
            <p><strong>Q:</strong> {qa['q']}</p>
            <p><strong>A:</strong> {qa['a']}</p>
        </div>
        """, unsafe_allow_html=True)


st.markdown("---")

st.markdown("## Bloco completo Power Query — correção de grafias incorretas")

st.code("""
let
    Fonte = Excel.CurrentWorkbook(){[Name="BaseVendas"]}[Content],

    TiposAlterados = Table.TransformColumnTypes(
        Fonte,
        {
            {"Produto", type text},
            {"Cliente", type text},
            {"Valor", type number}
        }
    ),

    ColunaOriginalPreservada = Table.DuplicateColumn(
        TiposAlterados,
        "Produto",
        "Produto_Original"
    ),

    ProdutoLimpo = Table.TransformColumns(
        ColunaOriginalPreservada,
        {
            {
                "Produto",
                each Text.Upper(Text.Trim(Text.Clean(_))),
                type text
            }
        }
    ),

    CorrecoesDiretas = Table.ReplaceValue(
        ProdutoLimpo,
        "SKOLL",
        "SKOL",
        Replacer.ReplaceText,
        {"Produto"}
    ),

    CorrecoesDiretas2 = Table.ReplaceValue(
        CorrecoesDiretas,
        "BRAHMAA",
        "BRAHMA",
        Replacer.ReplaceText,
        {"Produto"}
    ),

    CorrecoesDiretas3 = Table.ReplaceValue(
        CorrecoesDiretas2,
        "GUARANA ANTARTICA",
        "GUARANA ANTARCTICA",
        Replacer.ReplaceText,
        {"Produto"}
    ),

    ClassificacaoValor = Table.AddColumn(
        CorrecoesDiretas3,
        "Faixa_Valor",
        each if [Valor] >= 1000 then "Alta"
             else if [Valor] >= 500 then "Média"
             else "Baixa",
        type text
    ),

    ChaveAnalitica = Table.AddColumn(
        ClassificacaoValor,
        "Chave_Produto_Cliente",
        each [Produto] & "-" & Text.Upper(Text.Trim([Cliente])),
        type text
    )

in
    ChaveAnalitica
""", language="powerquery")
