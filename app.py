
import streamlit as st
import pandas as pd
from textwrap import dedent

# ============================================================
# COMITÊ TÉCNICO — EXCEL & POWER QUERY
# Versão revisada: intermediário + avançado, coringas,
# combinações de fórmulas, limpeza textual robusta e Data Quality
# ============================================================

st.set_page_config(
    page_title="Comitê Técnico | Excel & Power Query Avançado",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# CSS AUDITADO — SEM CORTE SUPERIOR / SEM QUEBRA HORIZONTAL
# ============================================================

st.markdown("""
<style>
html, body, [data-testid="stAppViewContainer"] {
    overflow-x: hidden !important;
}

.block-container {
    max-width: 1480px !important;
    padding-top: 1.15rem !important;
    padding-left: 1.35rem !important;
    padding-right: 1.35rem !important;
    padding-bottom: 2.5rem !important;
}

[data-testid="stSidebar"] {
    min-width: 285px !important;
    max-width: 330px !important;
}

[data-testid="stSidebarContent"] {
    padding-top: 1rem !important;
}

div[data-testid="stHorizontalBlock"] {
    align-items: stretch !important;
}

h1, h2, h3 {
    line-height: 1.2 !important;
}

.main-card {
    border: 1px solid #d9dee7;
    border-left: 5px solid #1f77b4;
    border-radius: 10px;
    padding: 16px;
    margin-bottom: 12px;
    background: #ffffff;
    min-height: 112px;
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

.purple-card {
    border-left-color: #7e57c2;
    background: #f5f0ff;
}

.section-title {
    font-size: 1.02rem;
    font-weight: 750;
    margin-bottom: 0.45rem;
    color: #17233c;
}

.tag {
    display: inline-block;
    padding: 5px 10px;
    border-radius: 999px;
    background: #e7f0ff;
    color: #0b5ed7;
    font-size: 0.78rem;
    font-weight: 750;
    margin-bottom: 8px;
}

.level-pill {
    display: inline-block;
    padding: 4px 9px;
    border-radius: 6px;
    background: #f1f3f5;
    color: #34495e;
    font-size: 0.76rem;
    font-weight: 700;
    margin-right: 4px;
}

pre, code {
    white-space: pre-wrap !important;
    word-break: break-word !important;
    overflow-wrap: anywhere !important;
    font-size: 0.80rem !important;
}

div[data-testid="stCodeBlock"] {
    max-width: 100% !important;
    overflow-x: auto !important;
}

button[kind="secondary"] {
    white-space: normal !important;
    min-height: 40px !important;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 0.25rem;
    flex-wrap: wrap;
}

.stTabs [data-baseweb="tab"] {
    height: auto;
    min-height: 38px;
    white-space: normal;
    padding: 8px 12px;
}

@media (max-width: 900px) {
    .block-container {
        padding-left: 0.75rem !important;
        padding-right: 0.75rem !important;
    }
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# DADOS DE EXEMPLO
# ============================================================

EXEMPLO_GRAFIAS = pd.DataFrame({
    "Produto Original": [
        " skol lata 350ml ",
        "SKOLL LATA",
        "Brahmaa duplo malte",
        "Guaraná Antartica",
        "agua tonica 350",
        "Cerv. Pilsen",
        "Heineken LongNeck",
        "Skol Lataa CX"
    ],
    "Após Limpeza": [
        "SKOL LATA 350ML",
        "SKOLL LATA",
        "BRAHMAA DUPLO MALTE",
        "GUARANA ANTARTICA",
        "AGUA TONICA 350",
        "CERV. PILSEN",
        "HEINEKEN LONGNECK",
        "SKOL LATAA CX"
    ],
    "Produto Final": [
        "SKOL LATA 350ML",
        "SKOL LATA",
        "BRAHMA DUPLO MALTE",
        "GUARANA ANTARCTICA",
        "AGUA TONICA 350ML",
        "CERVEJA PILSEN",
        "HEINEKEN LONG NECK",
        "SKOL LATA CX"
    ]
})

EXEMPLO_CORINGAS = pd.DataFrame({
    "Necessidade": [
        "Começa com SKOL",
        "Contém LATA",
        "Termina com 350ML",
        "Não contém ERRO",
        "Contém SKOL ou BRAHMA"
    ],
    "Excel": [
        '=CONT.SE(A2;"SKOL*")>0',
        '=CONT.SE(A2;"*LATA*")>0',
        '=CONT.SE(A2;"*350ML")>0',
        '=CONT.SE(A2;"*ERRO*")=0',
        '=OU(ÉNÚM(PROCURAR("SKOL";A2));ÉNÚM(PROCURAR("BRAHMA";A2)))'
    ],
    "Power Query / M": [
        'Text.StartsWith([Produto], "SKOL")',
        'Text.Contains([Produto], "LATA")',
        'Text.EndsWith([Produto], "350ML")',
        'not Text.Contains([Produto], "ERRO")',
        'Text.Contains([Produto], "SKOL") or Text.Contains([Produto], "BRAHMA")'
    ]
})


# ============================================================
# CONTEÚDO — EXCEL MUITO MAIS EXPLORADO
# ============================================================

DATA_MAPPING = {
    "Excel Básico": {
        "level": "Básico",
        "tag": "EXCEL-BASE",
        "title": "Fundamentos de base analítica",
        "objective": "Organizar dados, limpar textos simples, criar validações e evitar erros comuns.",
        "concept": "A base precisa estar em formato tabular: uma linha por registro, uma coluna por atributo, cabeçalhos únicos e sem células mescladas.",
        "when_use": "Controles operacionais, análises rápidas, pequenas bases e protótipos.",
        "attention": "Planilha visualmente bonita, mas mal estruturada, costuma falhar no Power Query, Power Pivot e BI.",
        "excel": [
            '=SE(A2="";"Sem informação";A2)',
            '=SEERRO(PROCX(A2;TabelaProdutos[Produto];TabelaProdutos[Categoria]);"Não localizado")',
            '=ARRUMAR(A2)',
            '=MAIÚSCULA(A2)',
            '=MINÚSCULA(A2)',
            '=PRI.MAIÚSCULA(A2)',
            '=SUBSTITUIR(A2;"  ";" ")',
            '=ESQUERDA(A2;3)',
            '=DIREITA(A2;4)',
            '=EXT.TEXTO(A2;4;6)',
            '=NÚM.CARACT(A2)',
            '=LOCALIZAR(" ";A2)',
            '=HOJE()',
            '=MÊS(A2)',
            '=ANO(A2)',
            '=TEXTO(A2;"mmm/aaaa")'
        ],
        "powerquery": [
            'Text.Trim([Produto])',
            'Text.Clean([Produto])',
            'Text.Upper([Produto])',
            'Text.Lower([Produto])',
            'Text.Proper([Produto])',
            'Date.Year([Data])',
            'Date.Month([Data])'
        ],
        "exercise": "Corrija uma base com espaços, caixa alta/baixa misturada, datas e valores nulos.",
        "qa": [
            ("Quando Excel basta?", "Quando há baixo volume, poucos usuários, baixa criticidade e necessidade de rapidez."),
            ("Qual regra de ouro antes de automatizar?", "Estruturar a base corretamente; automação em base ruim só acelera o erro.")
        ]
    },

    "Excel Intermediário": {
        "level": "Intermediário",
        "tag": "EXCEL-FORMULAS",
        "title": "Combinações de fórmulas e análise operacional",
        "objective": "Criar fórmulas combinadas para busca, classificação, consolidação e tratamento de erros.",
        "concept": "No nível intermediário, o diferencial é combinar funções: SE + E/OU, SEERRO + PROCX, SOMASES + critérios, FILTRO + CLASSIFICAR.",
        "when_use": "Relatórios recorrentes, conciliações, painéis simples, controles financeiros e operacionais.",
        "attention": "Fórmulas longas sem documentação dificultam manutenção. Nomeie tabelas e use colunas estruturadas.",
        "excel": [
            '=SE(E(B2>=1000;C2="Ativo");"Prioritário";"Normal")',
            '=SE(OU(A2="SKOL";A2="BRAHMA");"Cerveja";"Outros")',
            '=SEERRO(PROCX(A2;Produtos[SKU];Produtos[Categoria]);"Sem cadastro")',
            '=SOMASES(Vendas[Valor];Vendas[Produto];A2;Vendas[Canal];"Online")',
            '=CONT.SES(Vendas[Produto];A2;Vendas[Status];"Ativo")',
            '=MÉDIASES(Vendas[Valor];Vendas[Produto];A2;Vendas[Canal];"Distribuidor")',
            '=ÚNICO(Base[Produto])',
            '=CLASSIFICAR(ÚNICO(Base[Produto]))',
            '=FILTRO(Base;Base[Valor]>1000;"Sem registros")',
            '=CLASSIFICAR(FILTRO(Base;Base[Valor]>1000);3;-1)',
            '=TEXTOJUNTAR(" | ";VERDADEIRO;A2:C2)',
            '=LET(prod;ARRUMAR(MAIÚSCULA(A2));SE(prod="";"VAZIO";prod))',
            '=PROCX(1;(Produtos[Produto]=A2)*(Produtos[Canal]=B2);Produtos[Categoria];"Não localizado")',
            '=ÍNDICE(Tabela[Valor];CORRESP(1;(Tabela[Produto]=A2)*(Tabela[Canal]=B2);0))'
        ],
        "powerquery": [
            'Table.SelectRows(Fonte, each [Valor] > 1000 and [Status] = "Ativo")',
            'Table.Group(Fonte, {"Produto", "Canal"}, {{"Total", each List.Sum([Valor]), type number}})',
            'Table.Sort(Fonte, {{"Valor", Order.Descending}})',
            'Table.Distinct(Fonte, {"Produto"})'
        ],
        "exercise": "Monte um relatório com total por produto/canal, ranking, busca de categoria e tratamento de produtos sem cadastro.",
        "qa": [
            ("O que demonstra nível intermediário em Excel?", "Combinar funções com critérios, tratar erros, usar tabelas estruturadas e criar análises dinâmicas."),
            ("Por que usar LET?", "Para tornar fórmulas longas mais legíveis, reaproveitar cálculos e reduzir manutenção.")
        ]
    },

    "Excel Avançado": {
        "level": "Avançado",
        "tag": "EXCEL-ADV",
        "title": "Excel avançado para teste técnico",
        "objective": "Dominar matrizes dinâmicas, LET, LAMBDA, coringas, buscas compostas e fórmulas auditáveis.",
        "concept": "O Excel avançado combina clareza, performance e capacidade de resolver problemas sem criar planilhas frágeis.",
        "when_use": "Testes técnicos, modelos analíticos, conciliações complexas e automações sem VBA.",
        "attention": "Nem toda fórmula avançada é boa. Se o Power Query resolver melhor, use Power Query.",
        "excel": [
            '=CONT.SE(A:A;"*SKOL*")',
            '=CONT.SE(A:A;"SKOL*")',
            '=CONT.SE(A:A;"*350ML")',
            '=SOMASES(Base[Valor];Base[Produto];"*SKOL*";Base[Canal];"Online")',
            '=FILTRO(Base;ÉNÚM(PROCURAR("SKOL";Base[Produto]));"Sem SKOL")',
            '=PROCX("*SKOL*";Base[Produto];Base[Categoria];"Não localizado";2)',
            '=LET(txt;MAIÚSCULA(ARRUMAR(A2));SE(ÉNÚM(PROCURAR("SKOL";txt));"SKOL";txt))',
            '=LAMBDA(txt;MAIÚSCULA(ARRUMAR(txt)))(A2)',
            '=BYROW(A2:C10;LAMBDA(linha;SOMA(linha)))',
            '=MAP(A2:A10;LAMBDA(x;MAIÚSCULA(ARRUMAR(x))))',
            '=EMPILHARV(TabelaJan;TabelaFev;TabelaMar)',
            '=EMPILHARH(TabelaProdutos;TabelaCategorias)',
            '=ESCOLHERCOLS(Base;1;3;5)',
            '=DESCARTAR(Base;1)',
            '=PEGAR(CLASSIFICAR(Base;3;-1);10)',
            '=REDUZIR(0;Base[Valor];LAMBDA(acum;valor;acum+valor))',
            '=SCAN(0;Base[Valor];LAMBDA(acum;valor;acum+valor))'
        ],
        "powerquery": [
            'Text.Contains([Produto], "SKOL")',
            'Text.StartsWith([Produto], "SKOL")',
            'Text.EndsWith([Produto], "350ML")',
            'Table.SelectRows(Fonte, each Text.Contains([Produto], "SKOL"))',
            'Table.SelectRows(Fonte, each Text.StartsWith([Produto], "SKOL") or Text.Contains([Produto], "BRAHMA"))'
        ],
        "exercise": "Resolva uma base com busca parcial, critérios múltiplos, ranking top 10 e classificação automática por palavras-chave.",
        "qa": [
            ("Como usar coringas no Excel?", "Use * para qualquer sequência de caracteres e ? para um único caractere, especialmente em CONT.SE, SOMASES e PROCX com modo curinga."),
            ("Quando uma fórmula avançada vira problema?", "Quando fica difícil de auditar, lenta ou substitui uma transformação que seria melhor no Power Query.")
        ]
    },

    "Power Query Básico": {
        "level": "Básico",
        "tag": "PQ-M",
        "title": "Power Query e linguagem M",
        "objective": "Dominar transformações fundamentais de texto, número, data e tipo de dados.",
        "concept": "Power Query prepara dados. Ele usa linguagem M. Não usa DAX para transformação.",
        "when_use": "Bases recorrentes, bases sujas, múltiplas fontes, tratamento padronizado e atualização automática.",
        "attention": "A ordem dos passos importa. Etapas mal posicionadas podem quebrar consultas futuras.",
        "excel": [
            '=MAIÚSCULA(ARRUMAR(A2))',
            '=MINÚSCULA(ARRUMAR(A2))',
            '=PRI.MAIÚSCULA(ARRUMAR(A2))'
        ],
        "powerquery": [
            'Text.Upper(Text.Trim(Text.Clean([Produto])))',
            'Text.Lower(Text.Trim(Text.Clean([Produto])))',
            'Text.Proper(Text.Trim(Text.Clean([Cliente])))',
            'Number.Round([Valor], 2)',
            'Date.From([Data])',
            'Date.Year([Data])',
            'Date.Month([Data])',
            'Duration.Days([DataFim] - [DataInicio])',
            'Table.TransformColumnTypes(Fonte, {{"Data", type date}, {"Valor", type number}})'
        ],
        "exercise": "Crie colunas de produto em maiúsculo, cliente em formato próprio e diferença de dias entre datas.",
        "qa": [
            ("Power Query usa DAX?", "Não. Power Query usa M. DAX é para medidas no Power Pivot/Power BI."),
            ("Qual a principal vantagem do Power Query?", "Repetibilidade: você trata uma vez e atualiza sempre.")
        ]
    },

    "Power Query Intermediário": {
        "level": "Intermediário",
        "tag": "PQ-JOIN-GROUP",
        "title": "Merge, Append, Group By e regras condicionais",
        "objective": "Consolidar fontes, criar agrupamentos, aplicar regras e construir bases analíticas.",
        "concept": "O nível intermediário exige domínio de joins, agrupamentos, filtros condicionais e tabelas auxiliares.",
        "when_use": "Conciliação, cadastro De/Para, bases mensais, múltiplos arquivos e relatórios de fechamento.",
        "attention": "Antes de fazer Merge, padronize chaves dos dois lados.",
        "excel": [
            '=PROCX(A2;Cadastro[Chave];Cadastro[Nome];"Não localizado")',
            '=SOMASES(Base[Valor];Base[Produto];A2;Base[Mês];B2)'
        ],
        "powerquery": [
            'Table.NestedJoin(BaseVendas, {"Produto_Limpo"}, DeParaProdutos, {"Grafia_Incorreta"}, "Correcoes", JoinKind.LeftOuter)',
            'Table.ExpandTableColumn(MergeCorrecoes, "Correcoes", {"Produto_Correto"}, {"Produto_Correto"})',
            'Table.Combine({BaseJan, BaseFev, BaseMar})',
            'Table.Group(Fonte, {"Produto"}, {{"Total", each List.Sum([Valor]), type number}, {"Qtd", each Table.RowCount(_), Int64.Type}})',
            'Table.AddColumn(Fonte, "Faixa", each if [Valor] >= 1000 then "Alta" else if [Valor] >= 500 then "Média" else "Baixa", type text)',
            'Table.SelectRows(Fonte, each [Valor] <> null and [Valor] > 0)'
        ],
        "exercise": "Una três bases mensais, faça merge com cadastro de produtos e gere total por categoria.",
        "qa": [
            ("O que verificar antes de um Merge?", "Tipos de dados, espaços, caixa alta/baixa, acentos e duplicidade na chave da tabela de referência."),
            ("Quando usar Append?", "Quando as tabelas têm estrutura semelhante e precisam ser empilhadas.")
        ]
    },

    "Power Query Avançado": {
        "level": "Avançado",
        "tag": "PQ-ADV",
        "title": "Funções, listas, fuzzy matching e automação robusta",
        "objective": "Criar funções reutilizáveis, limpar textos complexos, aplicar listas de substituição e preparar automações escaláveis.",
        "concept": "Power Query avançado reduz código repetido usando funções, List.Accumulate, tabelas De/Para e regras de qualidade.",
        "when_use": "Bases muito sujas, grafias variadas, processos recorrentes e preparação para Power BI/Power Pivot.",
        "attention": "Fuzzy matching ajuda, mas deve ser validado. Em dado crítico, prefira De/Para auditável.",
        "excel": [
            '=LET(txt;MAIÚSCULA(ARRUMAR(A2));SEERRO(PROCX(txt;DePara[Erro];DePara[Correto]);txt))',
            '=PROCX("*"&A2&"*";Cadastro[Produto];Cadastro[Produto_Correto];"Validar";2)'
        ],
        "powerquery": [
            'List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))',
            'Table.FuzzyNestedJoin(Base, {"Produto_Limpo"}, Cadastro, {"Produto_Limpo"}, "Match", JoinKind.LeftOuter, [IgnoreCase=true, IgnoreSpace=true, Threshold=0.82])',
            'Table.Buffer(DeParaLimpo)',
            'Table.Profile(Fonte)',
            'Table.Schema(Fonte)',
            'try Number.From([Valor]) otherwise null',
            'try Date.From([Data]) otherwise null',
            'Table.SelectRows(Fonte, each List.Contains({"SKOL", "BRAHMA", "GUARANA"}, [Marca]))'
        ],
        "exercise": "Crie função de limpeza textual, remova acentos, aplique De/Para, use fuzzy matching e gere relatório de exceções.",
        "qa": [
            ("Quando usar Fuzzy Matching?", "Quando há pequenas diferenças textuais e uma tabela de referência confiável, sempre com validação."),
            ("Por que usar List.Accumulate?", "Para aplicar várias substituições em sequência sem criar dezenas de etapas manuais.")
        ]
    },

    "Grafias e Coringas": {
        "level": "Intermediário/Avançado",
        "tag": "TEXT-MATCH",
        "title": "Grafias incorretas, coringas e padrões textuais",
        "objective": "Corrigir variações como caixa alta/baixa, acentos, abreviações, erros comuns e correspondência parcial.",
        "concept": "A correção robusta combina quatro camadas: limpeza, normalização, substituições, De/Para e exceções.",
        "when_use": "Produtos, cidades, fornecedores, clientes, canais e descrições vindas de sistemas diferentes.",
        "attention": "Sempre mantenha Produto_Original, Produto_Limpo, Produto_Final e Status_Correcao.",
        "excel": [
            '=CONT.SE(A2;"*SKOL*")>0',
            '=SE(ÉNÚM(PROCURAR("SKOL";MAIÚSCULA(A2)));"SKOL";"OUTROS")',
            '=SEERRO(PROCX(MAIÚSCULA(ARRUMAR(A2));DePara[Grafia_Incorreta];DePara[Produto_Correto]);MAIÚSCULA(ARRUMAR(A2)))',
            '=PROCX("*"&A2&"*";Cadastro[Descrição];Cadastro[Produto];"Não localizado";2)',
            '=SE(OU(ÉNÚM(PROCURAR("LATA";A2));ÉNÚM(PROCURAR("LT";A2)));"LATA";"OUTRO")'
        ],
        "powerquery": [
            'Text.Contains([Produto_Limpo], "SKOL")',
            'Text.StartsWith([Produto_Limpo], "BRAHMA")',
            'Text.EndsWith([Produto_Limpo], "350ML")',
            'Text.Replace([Produto_Limpo], " LT ", " LATA ")',
            'Text.Remove([Produto_Limpo], {".", ",", ";", ":", "-", "_", "/", "\\\\", "(", ")"})',
            'if Text.Contains([Produto_Limpo], "SKOL") then "SKOL" else [Produto_Limpo]'
        ],
        "exercise": "Crie uma camada que padroniza SKOLL, SKOL LATAA, BRAHMAA, GUARANA ANTARTICA, CERV., LONGNECK e variações de LT/LATA.",
        "qa": [
            ("Qual arquitetura ideal para grafias incorretas?", "Original preservado, coluna limpa, remoção de acentos, substituições padronizadas, De/Para e status de correção."),
            ("Coringa no Excel equivale a quê no Power Query?", "No Excel usamos * e ?. No Power Query usamos Text.Contains, Text.StartsWith e Text.EndsWith.")
        ]
    },

    "Data Quality": {
        "level": "Intermediário",
        "tag": "DQ",
        "title": "Data Quality e auditoria",
        "objective": "Criar verificações de nulos, duplicidades, tipos, volume e exceções.",
        "concept": "Data Quality é o controle que impede que análises corretas sejam feitas sobre dados errados.",
        "when_use": "Sempre que a base alimentar dashboard, indicador, fechamento ou decisão.",
        "attention": "Não remova erros sem registrar. Crie relatório de exceções.",
        "excel": [
            '=SE(A2="";"Nulo";"OK")',
            '=SE(CONT.SE(A:A;A2)>1;"Duplicado";"Único")',
            '=SEERRO(VALOR(A2);"Erro numérico")',
            '=SE(E(B2<>"";C2>0);"OK";"Validar")',
            '=CONT.VALORES(A:A)',
            '=CONTAR.VAZIO(A:A)'
        ],
        "powerquery": [
            'Table.Profile(Fonte)',
            'Table.Schema(Fonte)',
            'Table.SelectRows(Fonte, each [Produto] = null or [Produto] = "")',
            'Table.Group(Fonte, {"Chave"}, {{"Qtd", each Table.RowCount(_), Int64.Type}})',
            'Table.SelectRows(Duplicados, each [Qtd] > 1)',
            'try Number.From([Valor]) otherwise null',
            'Table.AddColumn(Fonte, "Status_DQ", each if [Produto_Final] = null then "Erro Produto" else if [Valor] = null then "Erro Valor" else "OK", type text)'
        ],
        "exercise": "Crie um relatório de exceções com nulos, duplicados, produtos sem De/Para e valores inválidos.",
        "qa": [
            ("Como demonstrar Data Quality no teste?", "Mostre contagem de linhas, nulos, duplicados, exceções e explique impacto no indicador final.")
        ]
    }
}


# ============================================================
# BLOCO POWER QUERY COMPLETO — ROBUSTO
# ============================================================

BLOCO_POWER_QUERY_COMPLETO = dedent(r'''
let
    // ========================================================
    // 1. FONTE PRINCIPAL
    // Requer uma tabela Excel chamada BaseVendas
    // Colunas esperadas: Produto, Cliente, Valor, Data
    // ========================================================
    Fonte = Excel.CurrentWorkbook(){[Name="BaseVendas"]}[Content],

    TiposAlterados = Table.TransformColumnTypes(
        Fonte,
        {
            {"Produto", type text},
            {"Cliente", type text},
            {"Valor", type any},
            {"Data", type any}
        }
    ),

    // ========================================================
    // 2. FUNÇÃO INTERNA: REMOVER ACENTOS
    // ========================================================
    fnRemoveAcentos = (texto as nullable text) as nullable text =>
        let
            Entrada = if texto = null then null else texto,
            Substituicoes = {
                {"Á","A"},{"À","A"},{"Ã","A"},{"Â","A"},{"Ä","A"},
                {"É","E"},{"È","E"},{"Ê","E"},{"Ë","E"},
                {"Í","I"},{"Ì","I"},{"Î","I"},{"Ï","I"},
                {"Ó","O"},{"Ò","O"},{"Õ","O"},{"Ô","O"},{"Ö","O"},
                {"Ú","U"},{"Ù","U"},{"Û","U"},{"Ü","U"},
                {"Ç","C"},
                {"á","a"},{"à","a"},{"ã","a"},{"â","a"},{"ä","a"},
                {"é","e"},{"è","e"},{"ê","e"},{"ë","e"},
                {"í","i"},{"ì","i"},{"î","i"},{"ï","i"},
                {"ó","o"},{"ò","o"},{"õ","o"},{"ô","o"},{"ö","o"},
                {"ú","u"},{"ù","u"},{"û","u"},{"ü","u"},
                {"ç","c"}
            },
            Resultado = if Entrada = null then null else
                List.Accumulate(
                    Substituicoes,
                    Entrada,
                    (estado, atual) => Text.Replace(estado, atual{0}, atual{1})
                )
        in
            Resultado,

    // ========================================================
    // 3. FUNÇÃO INTERNA: LIMPEZA DE TEXTO
    // Remove acentos, invisíveis, pontuação comum e padroniza
    // ========================================================
    fnTextoLimpoMaiusculo = (texto as nullable text) as nullable text =>
        let
            Entrada = if texto = null then null else texto,
            SemAcentos = if Entrada = null then null else fnRemoveAcentos(Entrada),
            Limpo = if SemAcentos = null then null else Text.Clean(Text.Trim(SemAcentos)),
            Maiusculo = if Limpo = null then null else Text.Upper(Limpo),
            SemPontuacao = if Maiusculo = null then null else
                Text.Remove(Maiusculo, {".", ",", ";", ":", "-", "_", "/", "\\", "(", ")", "[", "]", "{", "}", "'", """"}),
            EspacosNormalizados = if SemPontuacao = null then null else
                Text.Combine(List.Select(Text.Split(SemPontuacao, " "), each _ <> ""), " ")
        in
            EspacosNormalizados,

    fnTextoLimpoMinusculo = (texto as nullable text) as nullable text =>
        let
            Base = fnTextoLimpoMaiusculo(texto),
            Resultado = if Base = null then null else Text.Lower(Base)
        in
            Resultado,

    // ========================================================
    // 4. PRESERVAR ORIGINAL
    // ========================================================
    ProdutoOriginal = Table.DuplicateColumn(
        TiposAlterados,
        "Produto",
        "Produto_Original"
    ),

    // ========================================================
    // 5. LIMPEZA EM MAIÚSCULO E MINÚSCULO
    // ========================================================
    ProdutoLimpo = Table.AddColumn(
        ProdutoOriginal,
        "Produto_Limpo",
        each fnTextoLimpoMaiusculo([Produto]),
        type text
    ),

    ProdutoLimpoMinusculo = Table.AddColumn(
        ProdutoLimpo,
        "Produto_Limpo_Min",
        each fnTextoLimpoMinusculo([Produto]),
        type text
    ),

    ClienteLimpo = Table.AddColumn(
        ProdutoLimpoMinusculo,
        "Cliente_Limpo",
        each fnTextoLimpoMaiusculo([Cliente]),
        type text
    ),

    // ========================================================
    // 6. SUBSTITUIÇÕES POR PADRÕES COMUNS
    // Exemplo: LT -> LATA, CX -> CAIXA, LONGNECK -> LONG NECK
    // ========================================================
    SubstituicoesPadrao = {
        {" LT ", " LATA "},
        {" LTA ", " LATA "},
        {" LATAA ", " LATA "},
        {" CX ", " CAIXA "},
        {" CERV ", " CERVEJA "},
        {" LONGNECK ", " LONG NECK "},
        {" GUARANA ANTARTICA ", " GUARANA ANTARCTICA "},
        {" BRAHMAA ", " BRAHMA "},
        {" SKOLL ", " SKOL "}
    },

    ProdutoPadronizado = Table.AddColumn(
        ClienteLimpo,
        "Produto_Padronizado",
        each
            let
                ComEspacos = " " & [Produto_Limpo] & " ",
                Corrigido = List.Accumulate(
                    SubstituicoesPadrao,
                    ComEspacos,
                    (estado, atual) => Text.Replace(estado, atual{0}, atual{1})
                ),
                Final = Text.Trim(Text.Combine(List.Select(Text.Split(Corrigido, " "), each _ <> ""), " "))
            in
                Final,
        type text
    ),

    // ========================================================
    // 7. CLASSIFICAÇÃO POR CORINGAS / PADRÕES TEXTUAIS
    // Power Query não usa * como Excel. Usa Text.Contains,
    // Text.StartsWith e Text.EndsWith.
    // ========================================================
    MarcaDetectada = Table.AddColumn(
        ProdutoPadronizado,
        "Marca_Detectada",
        each
            if [Produto_Padronizado] = null then "SEM PRODUTO"
            else if Text.Contains([Produto_Padronizado], "SKOL") then "SKOL"
            else if Text.Contains([Produto_Padronizado], "BRAHMA") then "BRAHMA"
            else if Text.Contains([Produto_Padronizado], "GUARANA") then "GUARANA"
            else if Text.Contains([Produto_Padronizado], "HEINEKEN") then "HEINEKEN"
            else "OUTROS",
        type text
    ),

    EmbalagemDetectada = Table.AddColumn(
        MarcaDetectada,
        "Embalagem_Detectada",
        each
            if [Produto_Padronizado] = null then "SEM EMBALAGEM"
            else if Text.Contains([Produto_Padronizado], "LATA") then "LATA"
            else if Text.Contains([Produto_Padronizado], "LONG NECK") then "LONG NECK"
            else if Text.Contains([Produto_Padronizado], "PET") then "PET"
            else if Text.Contains([Produto_Padronizado], "CAIXA") then "CAIXA"
            else "OUTROS",
        type text
    ),

    VolumeDetectado = Table.AddColumn(
        EmbalagemDetectada,
        "Volume_Detectado",
        each
            if [Produto_Padronizado] = null then null
            else if Text.Contains([Produto_Padronizado], "350") then "350ML"
            else if Text.Contains([Produto_Padronizado], "269") then "269ML"
            else if Text.Contains([Produto_Padronizado], "600") then "600ML"
            else if Text.Contains([Produto_Padronizado], "1L") then "1L"
            else null,
        type text
    ),

    // ========================================================
    // 8. TABELA DE/PARA AUDITÁVEL
    // Requer tabela Excel chamada DeParaProdutos com:
    // Grafia_Incorreta, Produto_Correto
    // ========================================================
    DeParaFonte = Excel.CurrentWorkbook(){[Name="DeParaProdutos"]}[Content],

    DeParaTipos = Table.TransformColumnTypes(
        DeParaFonte,
        {
            {"Grafia_Incorreta", type text},
            {"Produto_Correto", type text}
        }
    ),

    DeParaLimpo = Table.TransformColumns(
        DeParaTipos,
        {
            {"Grafia_Incorreta", each fnTextoLimpoMaiusculo(_), type text},
            {"Produto_Correto", each fnTextoLimpoMaiusculo(_), type text}
        }
    ),

    DeParaBuffer = Table.Buffer(DeParaLimpo),

    MergeCorrecoes = Table.NestedJoin(
        VolumeDetectado,
        {"Produto_Padronizado"},
        DeParaBuffer,
        {"Grafia_Incorreta"},
        "Correcoes",
        JoinKind.LeftOuter
    ),

    Expandido = Table.ExpandTableColumn(
        MergeCorrecoes,
        "Correcoes",
        {"Produto_Correto"},
        {"Produto_Correto"}
    ),

    ProdutoFinal = Table.AddColumn(
        Expandido,
        "Produto_Final",
        each if [Produto_Correto] <> null then [Produto_Correto] else [Produto_Padronizado],
        type text
    ),

    // ========================================================
    // 9. TRATAMENTO SEGURO DE VALOR E DATA
    // ========================================================
    ValorNumerico = Table.AddColumn(
        ProdutoFinal,
        "Valor_Numero",
        each try Number.From([Valor]) otherwise null,
        type number
    ),

    DataConvertida = Table.AddColumn(
        ValorNumerico,
        "Data_Convertida",
        each try Date.From([Data]) otherwise null,
        type date
    ),

    // ========================================================
    // 10. DATA QUALITY
    // ========================================================
    StatusCorrecao = Table.AddColumn(
        DataConvertida,
        "Status_Correcao",
        each
            if [Produto] = null or Text.Trim(Text.From([Produto])) = "" then "ERRO: PRODUTO VAZIO"
            else if [Produto_Correto] <> null then "CORRIGIDO POR DE/PARA"
            else if [Produto_Limpo] <> [Produto_Padronizado] then "CORRIGIDO POR REGRA"
            else "SEM ALTERAÇÃO",
        type text
    ),

    StatusDQ = Table.AddColumn(
        StatusCorrecao,
        "Status_DQ",
        each
            if [Produto_Final] = null or [Produto_Final] = "" then "ERRO PRODUTO"
            else if [Valor_Numero] = null then "ERRO VALOR"
            else if [Data_Convertida] = null then "ERRO DATA"
            else "OK",
        type text
    ),

    FaixaValor = Table.AddColumn(
        StatusDQ,
        "Faixa_Valor",
        each
            if [Valor_Numero] = null then "Sem valor"
            else if [Valor_Numero] >= 1000 then "Alta"
            else if [Valor_Numero] >= 500 then "Média"
            else "Baixa",
        type text
    ),

    ChaveAnalitica = Table.AddColumn(
        FaixaValor,
        "Chave_Produto_Cliente",
        each [Produto_Final] & "-" & [Cliente_Limpo],
        type text
    ),

    // ========================================================
    // 11. SELEÇÃO FINAL
    // ========================================================
    ResultadoFinal = Table.SelectColumns(
        ChaveAnalitica,
        {
            "Produto_Original",
            "Produto_Limpo",
            "Produto_Limpo_Min",
            "Produto_Padronizado",
            "Produto_Correto",
            "Produto_Final",
            "Marca_Detectada",
            "Embalagem_Detectada",
            "Volume_Detectado",
            "Cliente",
            "Cliente_Limpo",
            "Valor",
            "Valor_Numero",
            "Data",
            "Data_Convertida",
            "Faixa_Valor",
            "Chave_Produto_Cliente",
            "Status_Correcao",
            "Status_DQ"
        }
    )
in
    ResultadoFinal
''').strip()


BLOCO_RELATORIO_EXCECOES = dedent("""
let
    FonteTratada = ResultadoFinal,
    Excecoes = Table.SelectRows(
        FonteTratada,
        each [Status_DQ] <> "OK" or [Status_Correcao] <> "SEM ALTERAÇÃO"
    ),
    Resumo = Table.Group(
        Excecoes,
        {"Status_DQ", "Status_Correcao"},
        {{"Qtd", each Table.RowCount(_), Int64.Type}}
    )
in
    Resumo
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
    st.caption("Excel · Power Query · Data Quality · Coringas · Automação")

    st.divider()

    selected_topic = st.radio(
        "Mapa Mental",
        list(DATA_MAPPING.keys()),
        index=7
    )

    st.divider()

    st.metric("Meta", "Nota 8+")
    st.metric("Stack", "Excel + Power Query")
    st.info(
        "Dica: para teste técnico, mostre o raciocínio: original preservado, limpeza, regra, De/Para, exceções e validação."
    )


# ============================================================
# CONTEÚDO PRINCIPAL
# ============================================================

active = DATA_MAPPING[selected_topic]

st.title("Comitê Técnico — Excel & Power Query")
st.caption("Versão ampliada: fórmulas intermediárias/avançadas, coringas, Power Query robusto e Data Quality.")

st.markdown(
    f'<span class="tag">{active["tag"]}</span>'
    f'<span class="level-pill">Nível: {active["level"]}</span>',
    unsafe_allow_html=True
)

tabs = st.tabs([
    "Visão geral",
    "Excel",
    "Power Query",
    "Coringas",
    "Grafias",
    "Data Quality",
    "Exercício",
    "Entrevista",
    "Bloco completo M"
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
    render_code_items("Fórmulas Excel", active["excel"], "text")

with tabs[2]:
    render_code_items("Códigos Power Query / M", active["powerquery"], "powerquery")

with tabs[3]:
    st.markdown("### Coringas e equivalências")
    st.dataframe(EXEMPLO_CORINGAS, use_container_width=True, hide_index=True)
    st.markdown("### Fórmulas úteis com coringas")
    st.code('=CONT.SE(A:A;"*SKOL*")', language="text")
    st.code('=SOMASES(Base[Valor];Base[Produto];"*SKOL*")', language="text")
    st.code('=PROCX("*SKOL*";Base[Produto];Base[Categoria];"Não localizado";2)', language="text")
    st.code('=FILTRO(Base;ÉNÚM(PROCURAR("SKOL";Base[Produto]));"Sem registros")', language="text")

with tabs[4]:
    st.markdown("### Exemplo de limpeza de grafias")
    st.dataframe(EXEMPLO_GRAFIAS, use_container_width=True, hide_index=True)

    st.markdown("### Camadas recomendadas")
    st.markdown("""
    1. Preservar coluna original.  
    2. Criar coluna em maiúsculo.  
    3. Criar coluna em minúsculo, se necessário para comparação.  
    4. Remover acentos.  
    5. Remover pontuação.  
    6. Normalizar abreviações.  
    7. Aplicar tabela De/Para.  
    8. Criar status de correção.  
    9. Gerar relatório de exceções.  
    """)

with tabs[5]:
    st.markdown("### Data Quality")
    st.code("Table.Profile(Fonte)", language="powerquery")
    st.code("Table.Schema(Fonte)", language="powerquery")
    st.code('Table.SelectRows(Fonte, each [Status_DQ] <> "OK")', language="powerquery")
    st.code(BLOCO_RELATORIO_EXCECOES, language="powerquery")

with tabs[6]:
    card("Exercício prático", active["exercise"], "main-card green-card")
    st.markdown("### Entregáveis esperados")
    st.markdown("""
    - Base original preservada.  
    - Coluna limpa em maiúsculo.  
    - Coluna limpa em minúsculo.  
    - Produto final corrigido.  
    - Marca, embalagem e volume detectados.  
    - Status de correção.  
    - Status de Data Quality.  
    - Resumo executivo com exceções.  
    """)

with tabs[7]:
    for question, answer in active["qa"]:
        card(f"Q: {question}", f"<strong>A:</strong> {answer}", "main-card blue-card")

with tabs[8]:
    st.markdown("### Bloco completo Power Query — versão robusta")
    st.code(BLOCO_POWER_QUERY_COMPLETO, language="powerquery")

    st.download_button(
        label="Baixar código M",
        data=BLOCO_POWER_QUERY_COMPLETO,
        file_name="powerquery_grafias_robusto.m",
        mime="text/plain"
    )

    st.download_button(
        label="Baixar app Python",
        data=Path(__file__).read_text(encoding="utf-8") if "__file__" in globals() else "",
        file_name="comite_tecnico_excel_powerquery_v2.py",
        mime="text/x-python"
    )
