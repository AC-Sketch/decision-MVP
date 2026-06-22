import streamlit as st
import pandas as pd
from textwrap import dedent

# ============================================================
# CONFIGURAÇÃO DA PÁGINA
# ============================================================

st.set_page_config(
    page_title="Comitê Técnico | Excel & Power Query",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CSS AUDITADO — CORREÇÃO DE DIMENSIONAMENTO
# ============================================================

st.markdown(
    """
    <style>
    html, body, [data-testid="stAppViewContainer"] {
        overflow-x: hidden !important;
    }

    .block-container {
        max-width: 1440px !important;
        padding-top: 1.25rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        padding-bottom: 2rem !important;
    }

    [data-testid="stSidebar"] {
        min-width: 280px !important;
        max-width: 320px !important;
    }

    div[data-testid="stHorizontalBlock"] {
        align-items: stretch !important;
    }

    .main-card {
        border: 1px solid #d9dee7;
        border-left: 5px solid #1f77b4;
        border-radius: 10px;
        padding: 18px;
        margin-bottom: 14px;
        background: #ffffff;
        min-height: 120px;
    }

    .green-card {
        border-left-color: #18a558;
        background: #eefaf1;
    }

    .yellow-card {
        border-left-color: #f0ad4e;
        background: #fff8e6;
    }

    .red-card {
        border-left-color: #d9534f;
        background: #fff0f0;
    }

    .blue-card {
        border-left-color: #1f77b4;
        background: #eef6ff;
    }

    .small-card {
        border: 1px solid #d9dee7;
        border-radius: 10px;
        padding: 14px;
        background: #ffffff;
        min-height: 150px;
        margin-bottom: 12px;
    }

    .section-title {
        font-size: 1.05rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        color: #1b2a41;
    }

    .tag {
        display: inline-block;
        padding: 4px 9px;
        border-radius: 999px;
        background: #e7f0ff;
        color: #0b5ed7;
        font-size: 0.78rem;
        font-weight: 700;
        margin-bottom: 8px;
    }

    pre, code {
        white-space: pre-wrap !important;
        word-break: break-word !important;
        overflow-wrap: anywhere !important;
        font-size: 0.82rem !important;
    }

    div[data-testid="stCodeBlock"] {
        max-width: 100% !important;
        overflow-x: auto !important;
    }

    button[kind="secondary"] {
        white-space: normal !important;
        min-height: 42px !important;
    }

    @media (max-width: 900px) {
        .block-container {
            padding-left: 0.8rem !important;
            padding-right: 0.8rem !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ============================================================
# BASE DE CONTEÚDO
# ============================================================

DATA_MAPPING = {
    "Fundamentos": {
        "tag": "EXCEL-BASE",
        "title": "Excel como base analítica",
        "objective": "Entender quando o Excel é suficiente e como estruturar uma base limpa para análise.",
        "concept": "Excel é adequado para bases moderadas, análises rápidas, controles operacionais e protótipos.",
        "when_use": "Use Excel quando há poucos usuários, volume controlado, baixa complexidade e necessidade de resposta rápida.",
        "attention": "Evite usar Excel como sistema oficial quando houver múltiplos usuários editando, alto risco operacional ou histórico crítico.",
        "excel": [
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
        "exercise": "Crie uma tabela com Produto, Cliente, Data e Valor. Padronize textos, remova vazios e classifique valores.",
        "qa": [
            ("Quando Excel é suficiente?", "Quando o volume é controlado, a lógica é simples ou média, há poucos usuários e o risco operacional é baixo."),
            ("Quando migrar para Power Query?", "Quando há repetição, múltiplas fontes, bases sujas ou necessidade de atualização automática.")
        ]
    },

    "Power Query": {
        "tag": "PQ-CLEAN",
        "title": "Limpeza e padronização de textos",
        "objective": "Aprender a limpar campos textuais e corrigir inconsistências comuns.",
        "concept": "Power Query usa linguagem M. Ele prepara, transforma e automatiza dados antes da análise.",
        "when_use": "Use Power Query para importar, limpar, transformar, combinar e automatizar bases.",
        "attention": "Sempre preserve a coluna original antes de aplicar correções para garantir rastreabilidade.",
        "excel": [
            '=ARRUMAR(A2)',
            '=MAIÚSCULA(A2)',
            '=SUBSTITUIR(A2;"  ";" ")'
        ],
        "powerquery": [
            'Text.Trim(Text.Clean([Produto]))',
            'Text.Upper(Text.Trim(Text.Clean([Produto])))',
            'Text.Proper(Text.Trim([Cliente]))'
        ],
        "exercise": "Receba uma base com nomes como ' skol ', 'Skoll', 'BRAHMAA' e crie uma coluna Produto_Limpo.",
        "qa": [
            ("Por que usar Text.Trim e Text.Clean juntos?", "Text.Trim remove espaços externos. Text.Clean remove caracteres invisíveis."),
            ("Qual boa prática antes de limpar uma coluna?", "Duplicar a coluna original para manter auditoria.")
        ]
    },

    "Grafias incorretas": {
        "tag": "PQ-GRAFIAS",
        "title": "Remoção de grafias incorretas",
        "objective": "Criar lógica reutilizável para corrigir nomes de produtos, clientes, cidades ou fornecedores.",
        "concept": "A melhor solução é combinar limpeza automática com tabela De/Para.",
        "when_use": "Use quando o mesmo item aparece com várias grafias diferentes na base.",
        "attention": "Evite corrigir tudo manualmente no código. Prefira tabela de referência para manutenção e rastreabilidade.",
        "excel": [
            '=PROCX(A2;DePara[Grafia_Incorreta];DePara[Nome_Correto];A2)',
            '=MAIÚSCULA(ARRUMAR(A2))'
        ],
        "powerquery": [
            'Text.Upper(Text.Trim(Text.Clean([Produto])))',
            'if [Produto_Correto] <> null then [Produto_Correto] else [Produto_Limpo]',
            'Table.NestedJoin(BaseVendas, {"Produto_Limpo"}, DeParaProdutos, {"Grafia_Incorreta"}, "Correcoes", JoinKind.LeftOuter)'
        ],
        "exercise": "Crie uma tabela DeParaProdutos com SKOLL → SKOL, BRAHMAA → BRAHMA e GUARANA ANTARTICA → GUARANA ANTARCTICA.",
        "qa": [
            ("Como corrigir grafias incorretas de forma profissional?", "Criando uma coluna limpa, fazendo merge com uma tabela De/Para e gerando uma coluna final corrigida.")
        ]
    },

    "Data Quality": {
        "tag": "DQ-CHECK",
        "title": "Validação de qualidade dos dados",
        "objective": "Identificar nulos, duplicidades, tipos incorretos e erros de preenchimento.",
        "concept": "Data Quality garante que dashboards e análises não propaguem erros operacionais.",
        "when_use": "Use em toda base que será fonte de indicadores, relatórios ou tomada de decisão.",
        "attention": "Um dashboard automatizado com dado errado apenas espalha erro mais rápido.",
        "excel": [
            '=CONT.SE(A:A;A2)',
            '=SE(CONT.SE(A:A;A2)>1;"Duplicado";"Único")',
            '=SE(ÉCÉL.VAZIA(A2);"Nulo";"OK")'
        ],
        "powerquery": [
            'Table.SelectRows(Fonte, each [Produto] = null)',
            'Table.Distinct(Fonte)',
            'Table.RowCount(Fonte)'
        ],
        "exercise": "Crie rotina para contar linhas, remover duplicados e listar produtos sem preenchimento.",
        "qa": [
            ("O que checar antes de entregar um relatório?", "Nulos, duplicidades, tipos de dados, chaves inválidas, volume esperado e consistência com a regra de negócio.")
        ]
    },

    "Modelagem": {
        "tag": "MODELAGEM",
        "title": "Modelo analítico e Power Pivot",
        "objective": "Entender fato, dimensão, relacionamento e uso de DAX.",
        "concept": "Power Query prepara dados. Power Pivot modela. DAX cria medidas analíticas.",
        "when_use": "Use Power Pivot/DAX quando houver relacionamento entre tabelas e necessidade de KPIs.",
        "attention": "Não confunda M com DAX. M transforma dados; DAX calcula métricas no modelo.",
        "excel": [
            '=SOMASES(Vendas[Valor];Vendas[Produto];A2)',
            '=CONT.SES(Vendas[Cliente];A2;Vendas[Status];"Ativo")'
        ],
        "powerquery": [
            'Table.TransformColumnTypes(Fonte, {{"Data", type date}, {"Valor", type number}})',
            'Table.Group(Fonte, {"Produto"}, {{"Total", each List.Sum([Valor]), type number}})'
        ],
        "exercise": "Monte uma tabela fato de vendas e dimensões de produto, cliente e calendário.",
        "qa": [
            ("Power Query usa DAX?", "Não. Power Query usa linguagem M. DAX é usado no Power Pivot e Power BI para medidas.")
        ]
    },

    "Automação": {
        "tag": "AUTOMACAO",
        "title": "Automação de planilhas",
        "objective": "Reduzir atualização manual e criar processo replicável.",
        "concept": "Automação começa com uma base bem estruturada e consultas reutilizáveis.",
        "when_use": "Use quando há rotina recorrente, arquivos mensais, consolidação manual ou retrabalho.",
        "attention": "Automatizar sem padronizar a origem pode gerar falhas recorrentes.",
        "excel": [
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
            ("Qual é o ganho da automação?", "Menos retrabalho, menos erro manual, maior rastreabilidade e atualização mais rápida.")
        ]
    },

    "Teste Técnico": {
        "tag": "CASE",
        "title": "Simulação de case técnico",
        "objective": "Treinar raciocínio técnico para prova prática ou entrevista.",
        "concept": "Um bom teste técnico avalia limpeza, modelagem, indicador e explicação do raciocínio.",
        "when_use": "Use este roteiro para treinar cases de Excel, Power Query, BI e Data Analyst.",
        "attention": "Não basta entregar resultado. É preciso explicar método, risco e validação.",
        "excel": [
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
            ("Como explicar sua solução em entrevista?", "Eu preservei a base original, criei uma camada de limpeza no Power Query, usei tabela De/Para e gerei indicadores validados.")
        ]
    }
}

EXEMPLO_APLICACAO = pd.DataFrame({
    "Produto Original": ["SKOLL", "Brahmaa", "Guaraná Antartica", "Skol Lataa"],
    "Grafia Incorreta": ["SKOLL", "BRAHMAA", "GUARANA ANTARTICA", "SKOL LATAA"],
    "Produto Corrigido": ["SKOL", "BRAHMA", "GUARANA ANTARCTICA", "SKOL LATA"]
})

BLOCO_POWER_QUERY_COMPLETO = dedent("""
let
    // 1. Fonte principal
    Fonte = Excel.CurrentWorkbook(){[Name="BaseVendas"]}[Content],

    // 2. Ajuste de tipos
    TiposAlterados = Table.TransformColumnTypes(
        Fonte,
        {
            {"Produto", type text},
            {"Cliente", type text},
            {"Valor", type number}
        }
    ),

    // 3. Preservar coluna original para auditoria
    ProdutoOriginal = Table.DuplicateColumn(
        TiposAlterados,
        "Produto",
        "Produto_Original"
    ),

    // 4. Criar coluna limpa
    ProdutoLimpo = Table.AddColumn(
        ProdutoOriginal,
        "Produto_Limpo",
        each Text.Upper(Text.Trim(Text.Clean([Produto]))),
        type text
    ),

    // 5. Carregar tabela De/Para
    DePara = Excel.CurrentWorkbook(){[Name="DeParaProdutos"]}[Content],

    // 6. Ajustar tipos da tabela De/Para
    DeParaTipos = Table.TransformColumnTypes(
        DePara,
        {
            {"Grafia_Incorreta", type text},
            {"Produto_Correto", type text}
        }
    ),

    // 7. Limpar tabela De/Para
    DeParaLimpo = Table.TransformColumns(
        DeParaTipos,
        {
            {"Grafia_Incorreta", each Text.Upper(Text.Trim(Text.Clean(_))), type text},
            {"Produto_Correto", each Text.Upper(Text.Trim(Text.Clean(_))), type text}
        }
    ),

    // 8. Mesclar base principal com tabela De/Para
    MergeCorrecoes = Table.NestedJoin(
        ProdutoLimpo,
        {"Produto_Limpo"},
        DeParaLimpo,
        {"Grafia_Incorreta"},
        "Correcoes",
        JoinKind.LeftOuter
    ),

    // 9. Expandir produto correto
    Expandido = Table.ExpandTableColumn(
        MergeCorrecoes,
        "Correcoes",
        {"Produto_Correto"},
        {"Produto_Correto"}
    ),

    // 10. Criar produto final
    ProdutoFinal = Table.AddColumn(
        Expandido,
        "Produto_Final",
        each if [Produto_Correto] <> null then [Produto_Correto] else [Produto_Limpo],
        type text
    ),

    // 11. Criar faixa de valor
    FaixaValor = Table.AddColumn(
        ProdutoFinal,
        "Faixa_Valor",
        each if [Valor] >= 1000 then "Alta"
             else if [Valor] >= 500 then "Média"
             else "Baixa",
        type text
    ),

    // 12. Criar chave analítica
    ChaveAnalitica = Table.AddColumn(
        FaixaValor,
        "Chave_Produto_Cliente",
        each [Produto_Final] & "-" & Text.Upper(Text.Trim(Text.Clean([Cliente]))),
        type text
    )

in
    ChaveAnalitica
""").strip()


# ============================================================
# FUNÇÕES DE RENDERIZAÇÃO
# ============================================================

def card(title, text, css_class="main-card"):
    st.markdown(
        f"""
        <div class="{css_class}">
            <div class="section-title">{title}</div>
            <div>{text}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def render_code_items(title, items, language):
    st.markdown(f"### {title}")
    for item in items:
        st.code(item, language=language)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.markdown("## 📊 Comitê Técnico")
    st.caption("Excel · Power Query · Data Quality · Automação · Teste Técnico")

    st.divider()

    st.markdown("### Mapa Mental")
    selected_topic = st.radio(
        "Selecione o tópico:",
        list(DATA_MAPPING.keys()),
        index=2,
        label_visibility="collapsed"
    )

    st.divider()

    st.markdown("### Indicadores")
    st.metric("Foco", "Aprovação", "Teste Técnico")
    st.metric("Stack", "Excel + PQ")
    st.info("Dica do comitê: preserve a coluna original antes de corrigir qualquer grafia.")

# ============================================================
# CONTEÚDO PRINCIPAL
# ============================================================

active = DATA_MAPPING[selected_topic]

st.title("Comitê Técnico — Excel & Power Query")
st.caption("Base de estudo auditada para teste técnico, automação de planilhas e Data Quality.")

st.markdown(f'<span class="tag">{active["tag"]}</span>', unsafe_allow_html=True)

# Abas substituem botões horizontais que causavam corte no topo
tabs = st.tabs([
    "Visão geral",
    "Fórmulas",
    "Power Query",
    "Exercício",
    "Entrevista",
    "Bloco completo"
])

with tabs[0]:
    col1, col2 = st.columns([1, 1], gap="large")

    with col1:
        card(active["title"], f"<strong>Objetivo:</strong> {active['objective']}", "main-card green-card")
        card("Conceito essencial", active["concept"], "main-card blue-card")

    with col2:
        card("Quando usar", active["when_use"], "main-card yellow-card")
        card("Ponto de atenção", active["attention"], "main-card red-card")

with tabs[1]:
    render_code_items("Fórmulas Excel", active["excel"], "excel")

with tabs[2]:
    render_code_items("Códigos Power Query / M", active["powerquery"], "powerquery")

    st.markdown("### Exemplo de aplicação")
    st.dataframe(EXEMPLO_APLICACAO, use_container_width=True, hide_index=True)

with tabs[3]:
    card("Exercício prático", active["exercise"], "main-card green-card")

with tabs[4]:
    for question, answer in active["qa"]:
        card(f"Q: {question}", f"<strong>A:</strong> {answer}", "main-card blue-card")

with tabs[5]:
    st.markdown("### Bloco completo Power Query — correção de grafias incorretas")
    st.code(BLOCO_POWER_QUERY_COMPLETO, language="powerquery")

    st.download_button(
        label="Baixar código M",
        data=BLOCO_POWER_QUERY_COMPLETO,
        file_name="powerquery_correcao_grafias.m",
        mime="text/plain"
    )
