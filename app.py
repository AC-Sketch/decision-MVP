
import streamlit as st
import pandas as pd
import math

st.set_page_config(
    page_title="Comitê Técnico | Excel, Power Query e Estatística",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
html, body, [data-testid="stAppViewContainer"] {
    overflow-x: hidden !important;
}
.block-container {
    max-width: 1520px !important;
    padding-top: 1.0rem !important;
    padding-left: 1.25rem !important;
    padding-right: 1.25rem !important;
    padding-bottom: 2.5rem !important;
}
[data-testid="stSidebar"] {
    min-width: 290px !important;
    max-width: 345px !important;
}
[data-testid="stSidebarContent"] {
    padding-top: 1rem !important;
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
.card {
    border: 1px solid #d9dee7;
    border-left: 5px solid #1f77b4;
    border-radius: 10px;
    padding: 16px;
    margin-bottom: 12px;
    background: #ffffff;
}
.green { border-left-color: #18a558; background: #eefaf1; }
.yellow { border-left-color: #f0ad4e; background: #fff8e6; }
.red { border-left-color: #d9534f; background: #fff0f0; }
.blue { border-left-color: #1f77b4; background: #eef6ff; }
.purple { border-left-color: #7e57c2; background: #f5f0ff; }
.title-small {
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
    margin-right: 5px;
}
pre, code {
    white-space: pre-wrap !important;
    word-break: break-word !important;
    overflow-wrap: anywhere !important;
    font-size: 0.79rem !important;
}
div[data-testid="stCodeBlock"] {
    max-width: 100% !important;
    overflow-x: auto !important;
}
</style>
""", unsafe_allow_html=True)

EXCEL_BASIC = [{'Tema': 'Vazio', 'Formula': '=SE(A2="";"Sem informação";A2)', 'Uso': 'Substitui célula vazia por texto de controle.', 'Nivel': 'Básico'}, {'Tema': 'Erro', 'Formula': '=SEERRO(A2/B2;0)', 'Uso': 'Evita erro visual e retorna zero quando houver falha.', 'Nivel': 'Básico'}, {'Tema': 'Arrumar', 'Formula': '=ARRUMAR(A2)', 'Uso': 'Remove espaços excedentes entre palavras.', 'Nivel': 'Básico'}, {'Tema': 'Maiúscula', 'Formula': '=MAIÚSCULA(A2)', 'Uso': 'Padroniza texto em caixa alta.', 'Nivel': 'Básico'}, {'Tema': 'Minúscula', 'Formula': '=MINÚSCULA(A2)', 'Uso': 'Padroniza texto em caixa baixa.', 'Nivel': 'Básico'}, {'Tema': 'Nome próprio', 'Formula': '=PRI.MAIÚSCULA(A2)', 'Uso': 'Padroniza nomes próprios.', 'Nivel': 'Básico'}, {'Tema': 'Caracteres', 'Formula': '=NÚM.CARACT(A2)', 'Uso': 'Conta caracteres.', 'Nivel': 'Básico'}, {'Tema': 'Esquerda', 'Formula': '=ESQUERDA(A2;3)', 'Uso': 'Extrai os 3 primeiros caracteres.', 'Nivel': 'Básico'}, {'Tema': 'Direita', 'Formula': '=DIREITA(A2;4)', 'Uso': 'Extrai os 4 últimos caracteres.', 'Nivel': 'Básico'}, {'Tema': 'Ext texto', 'Formula': '=EXT.TEXTO(A2;4;6)', 'Uso': 'Extrai trecho intermediário.', 'Nivel': 'Básico'}, {'Tema': 'Localizar', 'Formula': '=LOCALIZAR("LATA";A2)', 'Uso': 'Localiza posição de texto.', 'Nivel': 'Básico'}, {'Tema': 'Substituir', 'Formula': '=SUBSTITUIR(A2;"LT";"LATA")', 'Uso': 'Troca trecho textual.', 'Nivel': 'Básico'}, {'Tema': 'Concatenar', 'Formula': '=A2&" - "&B2', 'Uso': 'Cria chave composta.', 'Nivel': 'Básico'}, {'Tema': 'Textojuntar', 'Formula': '=TEXTOJUNTAR(" | ";VERDADEIRO;A2:C2)', 'Uso': 'Concatena intervalo ignorando vazios.', 'Nivel': 'Básico'}, {'Tema': 'Hoje', 'Formula': '=HOJE()', 'Uso': 'Retorna data atual.', 'Nivel': 'Básico'}, {'Tema': 'Ano', 'Formula': '=ANO(A2)', 'Uso': 'Extrai ano.', 'Nivel': 'Básico'}, {'Tema': 'Mês', 'Formula': '=MÊS(A2)', 'Uso': 'Extrai mês.', 'Nivel': 'Básico'}, {'Tema': 'Dia', 'Formula': '=DIA(A2)', 'Uso': 'Extrai dia.', 'Nivel': 'Básico'}, {'Tema': 'Competência', 'Formula': '=TEXTO(A2;"mmm/aaaa")', 'Uso': 'Converte data em competência.', 'Nivel': 'Básico'}, {'Tema': 'Fim mês', 'Formula': '=FIMMÊS(A2;0)', 'Uso': 'Retorna último dia do mês.', 'Nivel': 'Básico'}]
EXCEL_INTERMEDIATE = [{'Tema': 'SE com E', 'Formula': '=SE(E(B2>=1000;C2="Ativo");"Prioritário";"Normal")', 'Uso': 'Classifica com critérios simultâneos.', 'Nivel': 'Intermediário'}, {'Tema': 'SE com OU', 'Formula': '=SE(OU(A2="SKOL";A2="BRAHMA");"Cerveja";"Outros")', 'Uso': 'Classifica se qualquer condição for verdadeira.', 'Nivel': 'Intermediário'}, {'Tema': 'SE aninhado', 'Formula': '=SE(B2>=1000;"Alta";SE(B2>=500;"Média";"Baixa"))', 'Uso': 'Cria faixas de valor.', 'Nivel': 'Intermediário'}, {'Tema': 'PROCX simples', 'Formula': '=PROCX(A2;Produtos[SKU];Produtos[Categoria];"Sem cadastro")', 'Uso': 'Busca moderna com retorno padrão.', 'Nivel': 'Intermediário'}, {'Tema': 'PROCX composto', 'Formula': '=PROCX(1;(Produtos[Produto]=A2)*(Produtos[Canal]=B2);Produtos[Categoria];"N/D")', 'Uso': 'Busca com múltiplos critérios.', 'Nivel': 'Intermediário'}, {'Tema': 'SOMASE', 'Formula': '=SOMASE(Base[Produto];A2;Base[Valor])', 'Uso': 'Soma com um critério.', 'Nivel': 'Intermediário'}, {'Tema': 'SOMASES', 'Formula': '=SOMASES(Base[Valor];Base[Produto];A2;Base[Canal];"Online")', 'Uso': 'Soma com múltiplos critérios.', 'Nivel': 'Intermediário'}, {'Tema': 'CONT.SE', 'Formula': '=CONT.SE(Base[Produto];A2)', 'Uso': 'Conta ocorrências.', 'Nivel': 'Intermediário'}, {'Tema': 'CONT.SES', 'Formula': '=CONT.SES(Base[Produto];A2;Base[Status];"Ativo")', 'Uso': 'Conta com múltiplos critérios.', 'Nivel': 'Intermediário'}, {'Tema': 'MÉDIASE', 'Formula': '=MÉDIASE(Base[Produto];A2;Base[Valor])', 'Uso': 'Média por critério.', 'Nivel': 'Intermediário'}, {'Tema': 'MÉDIASES', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A2;Base[Canal];"Online")', 'Uso': 'Média com múltiplos critérios.', 'Nivel': 'Intermediário'}, {'Tema': 'ÚNICO', 'Formula': '=ÚNICO(Base[Produto])', 'Uso': 'Lista valores únicos.', 'Nivel': 'Intermediário'}, {'Tema': 'CLASSIFICAR', 'Formula': '=CLASSIFICAR(ÚNICO(Base[Produto]))', 'Uso': 'Ordena lista dinâmica.', 'Nivel': 'Intermediário'}, {'Tema': 'FILTRO', 'Formula': '=FILTRO(Base;Base[Valor]>1000;"Sem registros")', 'Uso': 'Filtra base por condição.', 'Nivel': 'Intermediário'}, {'Tema': 'Top 10', 'Formula': '=PEGAR(CLASSIFICAR(Base;3;-1);10)', 'Uso': 'Retorna top 10 por coluna.', 'Nivel': 'Intermediário'}, {'Tema': 'ÍNDICE CORRESP', 'Formula': '=ÍNDICE(Tabela[Valor];CORRESP(A2;Tabela[Produto];0))', 'Uso': 'Busca clássica.', 'Nivel': 'Intermediário'}, {'Tema': 'ÍNDICE CORRESP composto', 'Formula': '=ÍNDICE(Tabela[Valor];CORRESP(1;(Tabela[Produto]=A2)*(Tabela[Canal]=B2);0))', 'Uso': 'Busca clássica com múltiplas condições.', 'Nivel': 'Intermediário'}, {'Tema': 'LET', 'Formula': '=LET(prod;ARRUMAR(MAIÚSCULA(A2));SE(prod="";"VAZIO";prod))', 'Uso': 'Nomeia partes da fórmula.', 'Nivel': 'Intermediário'}, {'Tema': 'Validação numérica', 'Formula': '=SEERRO(VALOR(A2);"Erro numérico")', 'Uso': 'Converte texto em número com controle.', 'Nivel': 'Intermediário'}, {'Tema': 'Duplicidade', 'Formula': '=SE(CONT.SE(A:A;A2)>1;"Duplicado";"Único")', 'Uso': 'Sinaliza duplicados.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 1', 'Formula': '=SEERRO(B2/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 2', 'Formula': '=SEERRO(B3/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 3', 'Formula': '=SEERRO(B4/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 4', 'Formula': '=SEERRO(B5/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 5', 'Formula': '=SEERRO(B6/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 6', 'Formula': '=SEERRO(B7/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 7', 'Formula': '=SEERRO(B8/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 8', 'Formula': '=SEERRO(B9/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 9', 'Formula': '=SEERRO(B10/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 10', 'Formula': '=SEERRO(B11/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 11', 'Formula': '=SEERRO(B12/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 12', 'Formula': '=SEERRO(B13/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 13', 'Formula': '=SEERRO(B14/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 14', 'Formula': '=SEERRO(B15/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 15', 'Formula': '=SEERRO(B16/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 16', 'Formula': '=SEERRO(B17/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 17', 'Formula': '=SEERRO(B18/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 18', 'Formula': '=SEERRO(B19/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 19', 'Formula': '=SEERRO(B20/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 20', 'Formula': '=SEERRO(B21/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 21', 'Formula': '=SEERRO(B22/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 22', 'Formula': '=SEERRO(B23/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 23', 'Formula': '=SEERRO(B24/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 24', 'Formula': '=SEERRO(B25/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 25', 'Formula': '=SEERRO(B26/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 26', 'Formula': '=SEERRO(B27/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 27', 'Formula': '=SEERRO(B28/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 28', 'Formula': '=SEERRO(B29/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 29', 'Formula': '=SEERRO(B30/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 30', 'Formula': '=SEERRO(B31/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 31', 'Formula': '=SEERRO(B32/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 32', 'Formula': '=SEERRO(B33/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 33', 'Formula': '=SEERRO(B34/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 34', 'Formula': '=SEERRO(B35/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 35', 'Formula': '=SEERRO(B36/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 36', 'Formula': '=SEERRO(B37/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 37', 'Formula': '=SEERRO(B38/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 38', 'Formula': '=SEERRO(B39/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 39', 'Formula': '=SEERRO(B40/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 40', 'Formula': '=SEERRO(B41/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 41', 'Formula': '=SEERRO(B42/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 42', 'Formula': '=SEERRO(B43/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 43', 'Formula': '=SEERRO(B44/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 44', 'Formula': '=SEERRO(B45/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 45', 'Formula': '=SEERRO(B46/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 46', 'Formula': '=SEERRO(B47/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 47', 'Formula': '=SEERRO(B48/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 48', 'Formula': '=SEERRO(B49/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 49', 'Formula': '=SEERRO(B50/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 50', 'Formula': '=SEERRO(B51/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 51', 'Formula': '=SEERRO(B52/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 52', 'Formula': '=SEERRO(B53/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 53', 'Formula': '=SEERRO(B54/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 54', 'Formula': '=SEERRO(B55/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 55', 'Formula': '=SEERRO(B56/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 56', 'Formula': '=SEERRO(B57/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 57', 'Formula': '=SEERRO(B58/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 58', 'Formula': '=SEERRO(B59/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 59', 'Formula': '=SEERRO(B60/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 60', 'Formula': '=SEERRO(B61/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 61', 'Formula': '=SEERRO(B62/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 62', 'Formula': '=SEERRO(B63/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 63', 'Formula': '=SEERRO(B64/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 64', 'Formula': '=SEERRO(B65/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 65', 'Formula': '=SEERRO(B66/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 66', 'Formula': '=SEERRO(B67/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 67', 'Formula': '=SEERRO(B68/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 68', 'Formula': '=SEERRO(B69/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 69', 'Formula': '=SEERRO(B70/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 70', 'Formula': '=SEERRO(B71/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 71', 'Formula': '=SEERRO(B72/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 72', 'Formula': '=SEERRO(B73/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 73', 'Formula': '=SEERRO(B74/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 74', 'Formula': '=SEERRO(B75/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 75', 'Formula': '=SEERRO(B76/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 76', 'Formula': '=SEERRO(B77/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 77', 'Formula': '=SEERRO(B78/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 78', 'Formula': '=SEERRO(B79/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 79', 'Formula': '=SEERRO(B80/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 80', 'Formula': '=SEERRO(B81/SOMA(B:B);0)', 'Uso': 'Calcula participação com tratamento de erro.', 'Nivel': 'Intermediário'}]
EXCEL_ADVANCED = [{'Tema': 'Coringa contém', 'Formula': '=CONT.SE(A:A;"*SKOL*")', 'Uso': 'Asterisco antes e depois significa contém.', 'Nivel': 'Avançado'}, {'Tema': 'Coringa começa', 'Formula': '=CONT.SE(A:A;"SKOL*")', 'Uso': 'Começa com SKOL.', 'Nivel': 'Avançado'}, {'Tema': 'Coringa termina', 'Formula': '=CONT.SE(A:A;"*350ML")', 'Uso': 'Termina com 350ML.', 'Nivel': 'Avançado'}, {'Tema': 'Coringa um caractere', 'Formula': '=CONT.SE(A:A;"SKO?")', 'Uso': '? representa exatamente um caractere.', 'Nivel': 'Avançado'}, {'Tema': 'Escapar asterisco', 'Formula': '=CONT.SE(A:A;"SKOL~*")', 'Uso': 'Procura SKOL* literalmente.', 'Nivel': 'Avançado'}, {'Tema': 'Escapar interrogação', 'Formula': '=CONT.SE(A:A;"SKOL~?")', 'Uso': 'Procura SKOL? literalmente.', 'Nivel': 'Avançado'}, {'Tema': 'Escapar til', 'Formula': '=CONT.SE(A:A;"SKU~~01")', 'Uso': 'Procura SKU~01 literalmente.', 'Nivel': 'Avançado'}, {'Tema': 'PROCX curinga', 'Formula': '=PROCX("*LATA*";Base[Produto];Base[Categoria];"N/D";2)', 'Uso': 'PROCX com modo curinga.', 'Nivel': 'Avançado'}, {'Tema': 'FILTRO com PROCURAR', 'Formula': '=FILTRO(Base;ÉNÚM(PROCURAR("SKOL";Base[Produto]));"Sem SKOL")', 'Uso': 'Filtra por texto contido.', 'Nivel': 'Avançado'}, {'Tema': 'LET + busca', 'Formula': '=LET(txt;MAIÚSCULA(ARRUMAR(A2));SE(ÉNÚM(PROCURAR("SKOL";txt));"SKOL";txt))', 'Uso': 'Limpa e classifica no mesmo cálculo.', 'Nivel': 'Avançado'}, {'Tema': 'LAMBDA inline', 'Formula': '=LAMBDA(txt;MAIÚSCULA(ARRUMAR(txt)))(A2)', 'Uso': 'Cria função reutilizável.', 'Nivel': 'Avançado'}, {'Tema': 'MAP', 'Formula': '=MAP(A2:A10;LAMBDA(x;MAIÚSCULA(ARRUMAR(x))))', 'Uso': 'Aplica função em cada item.', 'Nivel': 'Avançado'}, {'Tema': 'BYROW', 'Formula': '=BYROW(A2:C10;LAMBDA(linha;SOMA(linha)))', 'Uso': 'Calcula por linha.', 'Nivel': 'Avançado'}, {'Tema': 'REDUZIR', 'Formula': '=REDUZIR(0;Base[Valor];LAMBDA(acum;valor;acum+valor))', 'Uso': 'Acumula valores.', 'Nivel': 'Avançado'}, {'Tema': 'SCAN', 'Formula': '=SCAN(0;Base[Valor];LAMBDA(acum;valor;acum+valor))', 'Uso': 'Retorna acumulado progressivo.', 'Nivel': 'Avançado'}, {'Tema': 'EMPILHARV', 'Formula': '=EMPILHARV(TabelaJan;TabelaFev;TabelaMar)', 'Uso': 'Empilha tabelas verticalmente.', 'Nivel': 'Avançado'}, {'Tema': 'EMPILHARH', 'Formula': '=EMPILHARH(TabelaProdutos;TabelaCategorias)', 'Uso': 'Empilha tabelas horizontalmente.', 'Nivel': 'Avançado'}, {'Tema': 'ESCOLHERCOLS', 'Formula': '=ESCOLHERCOLS(Base;1;3;5)', 'Uso': 'Seleciona colunas.', 'Nivel': 'Avançado'}, {'Tema': 'DESCARTAR', 'Formula': '=DESCARTAR(Base;1)', 'Uso': 'Remove primeiras linhas ou colunas.', 'Nivel': 'Avançado'}, {'Tema': 'Filtro + classificar', 'Formula': '=CLASSIFICAR(FILTRO(Base;Base[Valor]>1000);3;-1)', 'Uso': 'Combina filtro e ordenação.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 1', 'Formula': '=LET(v;B2;limite;100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 2', 'Formula': '=LET(v;B3;limite;200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 3', 'Formula': '=LET(v;B4;limite;300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 4', 'Formula': '=LET(v;B5;limite;400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 5', 'Formula': '=LET(v;B6;limite;500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 6', 'Formula': '=LET(v;B7;limite;600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 7', 'Formula': '=LET(v;B8;limite;700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 8', 'Formula': '=LET(v;B9;limite;800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 9', 'Formula': '=LET(v;B10;limite;900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 10', 'Formula': '=LET(v;B11;limite;1000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 11', 'Formula': '=LET(v;B12;limite;1100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 12', 'Formula': '=LET(v;B13;limite;1200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 13', 'Formula': '=LET(v;B14;limite;1300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 14', 'Formula': '=LET(v;B15;limite;1400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 15', 'Formula': '=LET(v;B16;limite;1500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 16', 'Formula': '=LET(v;B17;limite;1600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 17', 'Formula': '=LET(v;B18;limite;1700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 18', 'Formula': '=LET(v;B19;limite;1800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 19', 'Formula': '=LET(v;B20;limite;1900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 20', 'Formula': '=LET(v;B21;limite;2000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 21', 'Formula': '=LET(v;B22;limite;2100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 22', 'Formula': '=LET(v;B23;limite;2200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 23', 'Formula': '=LET(v;B24;limite;2300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 24', 'Formula': '=LET(v;B25;limite;2400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 25', 'Formula': '=LET(v;B26;limite;2500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 26', 'Formula': '=LET(v;B27;limite;2600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 27', 'Formula': '=LET(v;B28;limite;2700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 28', 'Formula': '=LET(v;B29;limite;2800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 29', 'Formula': '=LET(v;B30;limite;2900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 30', 'Formula': '=LET(v;B31;limite;3000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 31', 'Formula': '=LET(v;B32;limite;3100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 32', 'Formula': '=LET(v;B33;limite;3200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 33', 'Formula': '=LET(v;B34;limite;3300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 34', 'Formula': '=LET(v;B35;limite;3400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 35', 'Formula': '=LET(v;B36;limite;3500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 36', 'Formula': '=LET(v;B37;limite;3600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 37', 'Formula': '=LET(v;B38;limite;3700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 38', 'Formula': '=LET(v;B39;limite;3800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 39', 'Formula': '=LET(v;B40;limite;3900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 40', 'Formula': '=LET(v;B41;limite;4000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 41', 'Formula': '=LET(v;B42;limite;4100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 42', 'Formula': '=LET(v;B43;limite;4200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 43', 'Formula': '=LET(v;B44;limite;4300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 44', 'Formula': '=LET(v;B45;limite;4400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 45', 'Formula': '=LET(v;B46;limite;4500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 46', 'Formula': '=LET(v;B47;limite;4600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 47', 'Formula': '=LET(v;B48;limite;4700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 48', 'Formula': '=LET(v;B49;limite;4800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 49', 'Formula': '=LET(v;B50;limite;4900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 50', 'Formula': '=LET(v;B51;limite;5000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 51', 'Formula': '=LET(v;B52;limite;5100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 52', 'Formula': '=LET(v;B53;limite;5200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 53', 'Formula': '=LET(v;B54;limite;5300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 54', 'Formula': '=LET(v;B55;limite;5400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 55', 'Formula': '=LET(v;B56;limite;5500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 56', 'Formula': '=LET(v;B57;limite;5600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 57', 'Formula': '=LET(v;B58;limite;5700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 58', 'Formula': '=LET(v;B59;limite;5800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 59', 'Formula': '=LET(v;B60;limite;5900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 60', 'Formula': '=LET(v;B61;limite;6000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 61', 'Formula': '=LET(v;B62;limite;6100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 62', 'Formula': '=LET(v;B63;limite;6200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 63', 'Formula': '=LET(v;B64;limite;6300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 64', 'Formula': '=LET(v;B65;limite;6400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 65', 'Formula': '=LET(v;B66;limite;6500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 66', 'Formula': '=LET(v;B67;limite;6600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 67', 'Formula': '=LET(v;B68;limite;6700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 68', 'Formula': '=LET(v;B69;limite;6800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 69', 'Formula': '=LET(v;B70;limite;6900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 70', 'Formula': '=LET(v;B71;limite;7000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Exemplo de fórmula auditável com variável.', 'Nivel': 'Avançado'}]
STATISTICS = [{'Tema': 'Média', 'Formula': '=MÉDIA(B:B)', 'Uso': 'Tendência central sensível a extremos.', 'Nivel': 'Estatística'}, {'Tema': 'Média condicional', 'Formula': '=MÉDIASE(A:A;"SKOL";B:B)', 'Uso': 'Média para um critério.', 'Nivel': 'Estatística'}, {'Tema': 'Médiases', 'Formula': '=MÉDIASES(B:B;A:A;"SKOL";C:C;"Online")', 'Uso': 'Média com vários critérios.', 'Nivel': 'Estatística'}, {'Tema': 'Mediana', 'Formula': '=MED(B:B)', 'Uso': 'Centro robusto contra extremos.', 'Nivel': 'Estatística'}, {'Tema': 'Moda', 'Formula': '=MODO.ÚNICO(B:B)', 'Uso': 'Valor mais frequente.', 'Nivel': 'Estatística'}, {'Tema': 'Variância amostral', 'Formula': '=VAR.S(B:B)', 'Uso': 'Dispersão da amostra.', 'Nivel': 'Estatística'}, {'Tema': 'Variância populacional', 'Formula': '=VAR.P(B:B)', 'Uso': 'Dispersão da população.', 'Nivel': 'Estatística'}, {'Tema': 'Desvio padrão amostral', 'Formula': '=DESVPAD.S(B:B)', 'Uso': 'Volatilidade amostral.', 'Nivel': 'Estatística'}, {'Tema': 'Desvio padrão populacional', 'Formula': '=DESVPAD.P(B:B)', 'Uso': 'Volatilidade populacional.', 'Nivel': 'Estatística'}, {'Tema': 'Padronizar', 'Formula': '=PADRONIZAR(B2;MÉDIA(B:B);DESVPAD.S(B:B))', 'Uso': 'Z-score.', 'Nivel': 'Estatística'}, {'Tema': 'Ordem', 'Formula': '=ORDEM.EQ(B2;B:B;0)', 'Uso': 'Ranking absoluto.', 'Nivel': 'Estatística'}, {'Tema': 'Ordem percentual', 'Formula': '=ORDEM.PORCENTUAL.INC(B:B;B2)', 'Uso': 'Posição percentual.', 'Nivel': 'Estatística'}, {'Tema': 'Percentil', 'Formula': '=PERCENTIL.INC(B:B;0,9)', 'Uso': 'Corte de percentil.', 'Nivel': 'Estatística'}, {'Tema': 'Quartil', 'Formula': '=QUARTIL.INC(B:B;3)', 'Uso': 'Corte em quartis.', 'Nivel': 'Estatística'}, {'Tema': 'Correlação', 'Formula': '=CORREL(B:B;C:C)', 'Uso': 'Força de relação linear.', 'Nivel': 'Estatística'}, {'Tema': 'Pearson', 'Formula': '=PEARSON(B:B;C:C)', 'Uso': 'Correlação de Pearson.', 'Nivel': 'Estatística'}, {'Tema': 'R quadrado', 'Formula': '=RQUAD(B:B;C:C)', 'Uso': 'Quanto X explica Y.', 'Nivel': 'Estatística'}, {'Tema': 'Inclinação', 'Formula': '=INCLINAÇÃO(B:B;C:C)', 'Uso': 'Coeficiente angular da reta.', 'Nivel': 'Estatística'}, {'Tema': 'Intercepção', 'Formula': '=INTERCEPÇÃO(B:B;C:C)', 'Uso': 'Intercepto da reta.', 'Nivel': 'Estatística'}, {'Tema': 'Erro padrão Y', 'Formula': '=EPADYX(B:B;C:C)', 'Uso': 'Erro padrão da estimativa.', 'Nivel': 'Estatística'}, {'Tema': 'Frequência', 'Formula': '=FREQUÊNCIA(B:B;E2:E6)', 'Uso': 'Distribuição por faixas.', 'Nivel': 'Estatística'}, {'Tema': 'Tendência', 'Formula': '=TENDÊNCIA(B:B;C:C;D2:D10)', 'Uso': 'Previsão linear.', 'Nivel': 'Estatística'}, {'Tema': 'Projeção linear', 'Formula': '=PROJ.LIN(B:B;C:C;VERDADEIRO;VERDADEIRO)', 'Uso': 'Regressão linear detalhada.', 'Nivel': 'Estatística'}, {'Tema': 'Projeção logarítmica', 'Formula': '=PROJ.LOG(B:B;C:C;VERDADEIRO;VERDADEIRO)', 'Uso': 'Modelo exponencial/logarítmico.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 1', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A2;Base[Canal];B2)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 2', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A3;Base[Canal];B3)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 3', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A4;Base[Canal];B4)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 4', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A5;Base[Canal];B5)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 5', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A6;Base[Canal];B6)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 6', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A7;Base[Canal];B7)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 7', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A8;Base[Canal];B8)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 8', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A9;Base[Canal];B9)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 9', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A10;Base[Canal];B10)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 10', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A11;Base[Canal];B11)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 11', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A12;Base[Canal];B12)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 12', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A13;Base[Canal];B13)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 13', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A14;Base[Canal];B14)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 14', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A15;Base[Canal];B15)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 15', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A16;Base[Canal];B16)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 16', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A17;Base[Canal];B17)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 17', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A18;Base[Canal];B18)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 18', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A19;Base[Canal];B19)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 19', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A20;Base[Canal];B20)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 20', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A21;Base[Canal];B21)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 21', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A22;Base[Canal];B22)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 22', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A23;Base[Canal];B23)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 23', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A24;Base[Canal];B24)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 24', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A25;Base[Canal];B25)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 25', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A26;Base[Canal];B26)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 26', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A27;Base[Canal];B27)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 27', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A28;Base[Canal];B28)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 28', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A29;Base[Canal];B29)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 29', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A30;Base[Canal];B30)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 30', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A31;Base[Canal];B31)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 31', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A32;Base[Canal];B32)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 32', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A33;Base[Canal];B33)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 33', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A34;Base[Canal];B34)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 34', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A35;Base[Canal];B35)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 35', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A36;Base[Canal];B36)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 36', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A37;Base[Canal];B37)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 37', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A38;Base[Canal];B38)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 38', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A39;Base[Canal];B39)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 39', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A40;Base[Canal];B40)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 40', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A41;Base[Canal];B41)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 41', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A42;Base[Canal];B42)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 42', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A43;Base[Canal];B43)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 43', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A44;Base[Canal];B44)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 44', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A45;Base[Canal];B45)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 45', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A46;Base[Canal];B46)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 46', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A47;Base[Canal];B47)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 47', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A48;Base[Canal];B48)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 48', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A49;Base[Canal];B49)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 49', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A50;Base[Canal];B50)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 50', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A51;Base[Canal];B51)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 51', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A52;Base[Canal];B52)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 52', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A53;Base[Canal];B53)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 53', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A54;Base[Canal];B54)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 54', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A55;Base[Canal];B55)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 55', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A56;Base[Canal];B56)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 56', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A57;Base[Canal];B57)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 57', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A58;Base[Canal];B58)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 58', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A59;Base[Canal];B59)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 59', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A60;Base[Canal];B60)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'Média segmentada 60', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A61;Base[Canal];B61)', 'Uso': 'Média por produto e canal.', 'Nivel': 'Estatística'}]
PQ_BASIC = [{'Tema': 'Trim', 'Formula': 'Text.Trim([Produto])', 'Uso': 'Remove espaços externos.', 'Nivel': 'Power Query'}, {'Tema': 'Clean', 'Formula': 'Text.Clean([Produto])', 'Uso': 'Remove caracteres não imprimíveis.', 'Nivel': 'Power Query'}, {'Tema': 'Upper', 'Formula': 'Text.Upper([Produto])', 'Uso': 'Maiúsculo.', 'Nivel': 'Power Query'}, {'Tema': 'Lower', 'Formula': 'Text.Lower([Produto])', 'Uso': 'Minúsculo.', 'Nivel': 'Power Query'}, {'Tema': 'Proper', 'Formula': 'Text.Proper([Cliente])', 'Uso': 'Nome próprio.', 'Nivel': 'Power Query'}, {'Tema': 'Contains', 'Formula': 'Text.Contains([Produto], "SKOL")', 'Uso': 'Contém texto.', 'Nivel': 'Power Query'}, {'Tema': 'StartsWith', 'Formula': 'Text.StartsWith([Produto], "SKOL")', 'Uso': 'Começa com texto.', 'Nivel': 'Power Query'}, {'Tema': 'EndsWith', 'Formula': 'Text.EndsWith([Produto], "350ML")', 'Uso': 'Termina com texto.', 'Nivel': 'Power Query'}, {'Tema': 'Text Length', 'Formula': 'Text.Length([Codigo])', 'Uso': 'Tamanho do texto.', 'Nivel': 'Power Query'}, {'Tema': 'Text Replace', 'Formula': 'Text.Replace([Produto], "LT", "LATA")', 'Uso': 'Substituição textual.', 'Nivel': 'Power Query'}, {'Tema': 'Text Remove', 'Formula': 'Text.Remove([Produto], {".", ",", "-", "_"})', 'Uso': 'Remove pontuação.', 'Nivel': 'Power Query'}, {'Tema': 'Date From', 'Formula': 'Date.From([Data])', 'Uso': 'Converte para data.', 'Nivel': 'Power Query'}, {'Tema': 'Number From', 'Formula': 'try Number.From([Valor]) otherwise null', 'Uso': 'Converte número com segurança.', 'Nivel': 'Power Query'}, {'Tema': 'Duration Days', 'Formula': 'Duration.Days([DataFim] - [DataInicio])', 'Uso': 'Diferença entre datas.', 'Nivel': 'Power Query'}]
PQ_INTERMEDIATE = [{'Tema': 'Select Rows', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 1000)', 'Uso': 'Filtra linhas.', 'Nivel': 'Power Query'}, {'Tema': 'Add Column', 'Formula': 'Table.AddColumn(Fonte, "Faixa", each if [Valor] >= 1000 then "Alta" else "Baixa", type text)', 'Uso': 'Adiciona coluna.', 'Nivel': 'Power Query'}, {'Tema': 'Group', 'Formula': 'Table.Group(Fonte, {"Produto"}, {{"Total", each List.Sum([Valor]), type number}})', 'Uso': 'Agrupa e soma.', 'Nivel': 'Power Query'}, {'Tema': 'Nested Join', 'Formula': 'Table.NestedJoin(Base, {"Produto_Limpo"}, DePara, {"Grafia_Incorreta"}, "Correcoes", JoinKind.LeftOuter)', 'Uso': 'Merge.', 'Nivel': 'Power Query'}, {'Tema': 'Expand', 'Formula': 'Table.ExpandTableColumn(Merge, "Correcoes", {"Produto_Correto"}, {"Produto_Correto"})', 'Uso': 'Expande merge.', 'Nivel': 'Power Query'}, {'Tema': 'Combine', 'Formula': 'Table.Combine({BaseJan, BaseFev, BaseMar})', 'Uso': 'Append.', 'Nivel': 'Power Query'}, {'Tema': 'Distinct', 'Formula': 'Table.Distinct(Fonte, {"Chave"})', 'Uso': 'Remove duplicados.', 'Nivel': 'Power Query'}, {'Tema': 'Sort', 'Formula': 'Table.Sort(Fonte, {{"Valor", Order.Descending}})', 'Uso': 'Ordena.', 'Nivel': 'Power Query'}, {'Tema': 'Replace Value', 'Formula': 'Table.ReplaceValue(Fonte, "SKOLL", "SKOL", Replacer.ReplaceText, {"Produto"})', 'Uso': 'Substitui valor.', 'Nivel': 'Power Query'}, {'Tema': 'Column Types', 'Formula': 'Table.TransformColumnTypes(Fonte, {{"Data", type date}, {"Valor", type number}})', 'Uso': 'Ajusta tipos.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 1', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 10)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 2', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 20)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 3', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 30)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 4', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 40)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 5', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 50)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 6', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 60)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 7', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 70)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 8', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 80)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 9', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 90)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 10', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 100)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 11', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 110)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 12', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 120)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 13', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 130)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 14', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 140)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 15', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 150)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 16', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 160)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 17', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 170)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 18', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 180)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 19', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 190)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 20', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 200)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 21', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 210)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 22', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 220)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 23', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 230)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 24', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 240)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 25', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 250)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 26', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 260)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 27', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 270)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 28', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 280)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 29', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 290)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 30', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 300)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 31', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 310)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 32', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 320)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 33', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 330)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 34', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 340)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 35', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 350)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 36', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 360)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 37', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 370)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 38', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 380)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 39', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 390)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 40', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 400)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 41', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 410)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 42', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 420)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 43', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 430)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 44', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 440)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 45', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 450)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 46', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 460)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 47', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 470)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 48', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 480)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 49', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 490)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 50', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 500)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 51', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 510)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 52', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 520)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 53', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 530)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 54', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 540)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 55', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 550)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 56', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 560)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 57', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 570)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 58', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 580)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 59', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 590)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 60', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 600)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 61', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 610)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 62', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 620)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 63', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 630)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 64', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 640)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 65', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 650)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 66', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 660)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 67', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 670)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 68', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 680)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 69', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 690)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 70', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 700)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 71', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 710)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 72', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 720)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 73', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 730)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 74', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 740)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 75', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 750)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 76', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 760)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 77', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 770)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 78', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 780)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 79', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 790)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro paramétrico 80', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 800)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}]
PQ_ADVANCED = [{'Tema': 'List Accumulate', 'Formula': 'List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))', 'Uso': 'Aplica substituições em sequência.', 'Nivel': 'Power Query'}, {'Tema': 'Fuzzy Join', 'Formula': 'Table.FuzzyNestedJoin(Base, {"Produto_Limpo"}, Cadastro, {"Produto_Limpo"}, "Match", JoinKind.LeftOuter, [IgnoreCase=true, IgnoreSpace=true, Threshold=0.82])', 'Uso': 'Correspondência aproximada.', 'Nivel': 'Power Query'}, {'Tema': 'Buffer', 'Formula': 'Table.Buffer(DeParaLimpo)', 'Uso': 'Evita reprocessamento.', 'Nivel': 'Power Query'}, {'Tema': 'Profile', 'Formula': 'Table.Profile(Fonte)', 'Uso': 'Perfil da base.', 'Nivel': 'Power Query'}, {'Tema': 'Schema', 'Formula': 'Table.Schema(Fonte)', 'Uso': 'Estrutura da base.', 'Nivel': 'Power Query'}, {'Tema': 'List Contains', 'Formula': 'List.Contains({"SKOL", "BRAHMA", "GUARANA"}, [Marca])', 'Uso': 'Validação por lista.', 'Nivel': 'Power Query'}, {'Tema': 'Try Date', 'Formula': 'try Date.From([Data]) otherwise null', 'Uso': 'Conversão segura de data.', 'Nivel': 'Power Query'}, {'Tema': 'Try Number', 'Formula': 'try Number.From([Valor]) otherwise null', 'Uso': 'Conversão segura de número.', 'Nivel': 'Power Query'}, {'Tema': 'MissingField Ignore', 'Formula': 'Table.SelectColumns(Fonte, {"A", "B"}, MissingField.Ignore)', 'Uso': 'Evita erro por coluna ausente.', 'Nivel': 'Power Query'}, {'Tema': 'Unpivot', 'Formula': 'Table.UnpivotOtherColumns(Fonte, {"Produto"}, "Mês", "Valor")', 'Uso': 'Desnormaliza colunas em linhas.', 'Nivel': 'Power Query'}, {'Tema': 'Pivot', 'Formula': 'Table.Pivot(Fonte, List.Distinct(Fonte[Mês]), "Mês", "Valor", List.Sum)', 'Uso': 'Transforma linhas em colunas.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 1', 'Formula': 'Table.AddColumn(Fonte, "Regra_1", each if [Valor] >= 100 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 2', 'Formula': 'Table.AddColumn(Fonte, "Regra_2", each if [Valor] >= 200 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 3', 'Formula': 'Table.AddColumn(Fonte, "Regra_3", each if [Valor] >= 300 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 4', 'Formula': 'Table.AddColumn(Fonte, "Regra_4", each if [Valor] >= 400 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 5', 'Formula': 'Table.AddColumn(Fonte, "Regra_5", each if [Valor] >= 500 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 6', 'Formula': 'Table.AddColumn(Fonte, "Regra_6", each if [Valor] >= 600 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 7', 'Formula': 'Table.AddColumn(Fonte, "Regra_7", each if [Valor] >= 700 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 8', 'Formula': 'Table.AddColumn(Fonte, "Regra_8", each if [Valor] >= 800 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 9', 'Formula': 'Table.AddColumn(Fonte, "Regra_9", each if [Valor] >= 900 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 10', 'Formula': 'Table.AddColumn(Fonte, "Regra_10", each if [Valor] >= 1000 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 11', 'Formula': 'Table.AddColumn(Fonte, "Regra_11", each if [Valor] >= 1100 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 12', 'Formula': 'Table.AddColumn(Fonte, "Regra_12", each if [Valor] >= 1200 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 13', 'Formula': 'Table.AddColumn(Fonte, "Regra_13", each if [Valor] >= 1300 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 14', 'Formula': 'Table.AddColumn(Fonte, "Regra_14", each if [Valor] >= 1400 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 15', 'Formula': 'Table.AddColumn(Fonte, "Regra_15", each if [Valor] >= 1500 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 16', 'Formula': 'Table.AddColumn(Fonte, "Regra_16", each if [Valor] >= 1600 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 17', 'Formula': 'Table.AddColumn(Fonte, "Regra_17", each if [Valor] >= 1700 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 18', 'Formula': 'Table.AddColumn(Fonte, "Regra_18", each if [Valor] >= 1800 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 19', 'Formula': 'Table.AddColumn(Fonte, "Regra_19", each if [Valor] >= 1900 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 20', 'Formula': 'Table.AddColumn(Fonte, "Regra_20", each if [Valor] >= 2000 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 21', 'Formula': 'Table.AddColumn(Fonte, "Regra_21", each if [Valor] >= 2100 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 22', 'Formula': 'Table.AddColumn(Fonte, "Regra_22", each if [Valor] >= 2200 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 23', 'Formula': 'Table.AddColumn(Fonte, "Regra_23", each if [Valor] >= 2300 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 24', 'Formula': 'Table.AddColumn(Fonte, "Regra_24", each if [Valor] >= 2400 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 25', 'Formula': 'Table.AddColumn(Fonte, "Regra_25", each if [Valor] >= 2500 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 26', 'Formula': 'Table.AddColumn(Fonte, "Regra_26", each if [Valor] >= 2600 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 27', 'Formula': 'Table.AddColumn(Fonte, "Regra_27", each if [Valor] >= 2700 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 28', 'Formula': 'Table.AddColumn(Fonte, "Regra_28", each if [Valor] >= 2800 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 29', 'Formula': 'Table.AddColumn(Fonte, "Regra_29", each if [Valor] >= 2900 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 30', 'Formula': 'Table.AddColumn(Fonte, "Regra_30", each if [Valor] >= 3000 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 31', 'Formula': 'Table.AddColumn(Fonte, "Regra_31", each if [Valor] >= 3100 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 32', 'Formula': 'Table.AddColumn(Fonte, "Regra_32", each if [Valor] >= 3200 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 33', 'Formula': 'Table.AddColumn(Fonte, "Regra_33", each if [Valor] >= 3300 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 34', 'Formula': 'Table.AddColumn(Fonte, "Regra_34", each if [Valor] >= 3400 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 35', 'Formula': 'Table.AddColumn(Fonte, "Regra_35", each if [Valor] >= 3500 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 36', 'Formula': 'Table.AddColumn(Fonte, "Regra_36", each if [Valor] >= 3600 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 37', 'Formula': 'Table.AddColumn(Fonte, "Regra_37", each if [Valor] >= 3700 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 38', 'Formula': 'Table.AddColumn(Fonte, "Regra_38", each if [Valor] >= 3800 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 39', 'Formula': 'Table.AddColumn(Fonte, "Regra_39", each if [Valor] >= 3900 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 40', 'Formula': 'Table.AddColumn(Fonte, "Regra_40", each if [Valor] >= 4000 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 41', 'Formula': 'Table.AddColumn(Fonte, "Regra_41", each if [Valor] >= 4100 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 42', 'Formula': 'Table.AddColumn(Fonte, "Regra_42", each if [Valor] >= 4200 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 43', 'Formula': 'Table.AddColumn(Fonte, "Regra_43", each if [Valor] >= 4300 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 44', 'Formula': 'Table.AddColumn(Fonte, "Regra_44", each if [Valor] >= 4400 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 45', 'Formula': 'Table.AddColumn(Fonte, "Regra_45", each if [Valor] >= 4500 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 46', 'Formula': 'Table.AddColumn(Fonte, "Regra_46", each if [Valor] >= 4600 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 47', 'Formula': 'Table.AddColumn(Fonte, "Regra_47", each if [Valor] >= 4700 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 48', 'Formula': 'Table.AddColumn(Fonte, "Regra_48", each if [Valor] >= 4800 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 49', 'Formula': 'Table.AddColumn(Fonte, "Regra_49", each if [Valor] >= 4900 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 50', 'Formula': 'Table.AddColumn(Fonte, "Regra_50", each if [Valor] >= 5000 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 51', 'Formula': 'Table.AddColumn(Fonte, "Regra_51", each if [Valor] >= 5100 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 52', 'Formula': 'Table.AddColumn(Fonte, "Regra_52", each if [Valor] >= 5200 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 53', 'Formula': 'Table.AddColumn(Fonte, "Regra_53", each if [Valor] >= 5300 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 54', 'Formula': 'Table.AddColumn(Fonte, "Regra_54", each if [Valor] >= 5400 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 55', 'Formula': 'Table.AddColumn(Fonte, "Regra_55", each if [Valor] >= 5500 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 56', 'Formula': 'Table.AddColumn(Fonte, "Regra_56", each if [Valor] >= 5600 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 57', 'Formula': 'Table.AddColumn(Fonte, "Regra_57", each if [Valor] >= 5700 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 58', 'Formula': 'Table.AddColumn(Fonte, "Regra_58", each if [Valor] >= 5800 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 59', 'Formula': 'Table.AddColumn(Fonte, "Regra_59", each if [Valor] >= 5900 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 60', 'Formula': 'Table.AddColumn(Fonte, "Regra_60", each if [Valor] >= 6000 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 61', 'Formula': 'Table.AddColumn(Fonte, "Regra_61", each if [Valor] >= 6100 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 62', 'Formula': 'Table.AddColumn(Fonte, "Regra_62", each if [Valor] >= 6200 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 63', 'Formula': 'Table.AddColumn(Fonte, "Regra_63", each if [Valor] >= 6300 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 64', 'Formula': 'Table.AddColumn(Fonte, "Regra_64", each if [Valor] >= 6400 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 65', 'Formula': 'Table.AddColumn(Fonte, "Regra_65", each if [Valor] >= 6500 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 66', 'Formula': 'Table.AddColumn(Fonte, "Regra_66", each if [Valor] >= 6600 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 67', 'Formula': 'Table.AddColumn(Fonte, "Regra_67", each if [Valor] >= 6700 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 68', 'Formula': 'Table.AddColumn(Fonte, "Regra_68", each if [Valor] >= 6800 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 69', 'Formula': 'Table.AddColumn(Fonte, "Regra_69", each if [Valor] >= 6900 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional 70', 'Formula': 'Table.AddColumn(Fonte, "Regra_70", each if [Valor] >= 7000 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra condicional auditável.', 'Nivel': 'Power Query'}]
WILDCARDS = [{'Coringa': '*', 'Uso': 'Qualquer sequência', 'Excel': '=CONT.SE(A:A;"*SKOL*")', 'Power Query': 'Text.Contains([Produto], "SKOL")'}, {'Coringa': '*', 'Uso': 'Começa com', 'Excel': '=CONT.SE(A:A;"SKOL*")', 'Power Query': 'Text.StartsWith([Produto], "SKOL")'}, {'Coringa': '*', 'Uso': 'Termina com', 'Excel': '=CONT.SE(A:A;"*350ML")', 'Power Query': 'Text.EndsWith([Produto], "350ML")'}, {'Coringa': '?', 'Uso': 'Um caractere', 'Excel': '=CONT.SE(A:A;"SKO?")', 'Power Query': 'Text.StartsWith([Codigo], "SKO") and Text.Length([Codigo]) = 4'}, {'Coringa': '???', 'Uso': 'Três caracteres', 'Excel': '=CONT.SE(A:A;"???")', 'Power Query': 'Text.Length([Codigo]) = 3'}, {'Coringa': '~*', 'Uso': 'Asterisco literal', 'Excel': '=CONT.SE(A:A;"SKOL~*")', 'Power Query': 'Text.Contains([Produto], "SKOL*")'}, {'Coringa': '~?', 'Uso': 'Interrogação literal', 'Excel': '=CONT.SE(A:A;"SKOL~?")', 'Power Query': 'Text.Contains([Produto], "SKOL?")'}, {'Coringa': '~~', 'Uso': 'Til literal', 'Excel': '=CONT.SE(A:A;"SKU~~01")', 'Power Query': 'Text.Contains([Produto], "SKU~01")'}]
DEPARA_EXAMPLES = [{'Grafia_Incorreta': 'SKOLL', 'Produto_Correto': 'SKOL', 'Motivo': 'Letra excedente'}, {'Grafia_Incorreta': 'BRAHMAA', 'Produto_Correto': 'BRAHMA', 'Motivo': 'Letra excedente'}, {'Grafia_Incorreta': 'BRHMA', 'Produto_Correto': 'BRAHMA', 'Motivo': 'Letra faltando'}, {'Grafia_Incorreta': 'GUARANA ANTARTICA', 'Produto_Correto': 'GUARANA ANTARCTICA', 'Motivo': 'Grafia comercial'}, {'Grafia_Incorreta': 'SKOL LATAA', 'Produto_Correto': 'SKOL LATA', 'Motivo': 'Letra excedente'}, {'Grafia_Incorreta': 'CERV PILSEN', 'Produto_Correto': 'CERVEJA PILSEN', 'Motivo': 'Abreviação'}, {'Grafia_Incorreta': 'LONGNECK', 'Produto_Correto': 'LONG NECK', 'Motivo': 'Espaçamento'}, {'Grafia_Incorreta': 'LT', 'Produto_Correto': 'LATA', 'Motivo': 'Abreviação'}, {'Grafia_Incorreta': 'CX', 'Produto_Correto': 'CAIXA', 'Motivo': 'Abreviação'}, {'Grafia_Incorreta': 'AGUA TONICA 350', 'Produto_Correto': 'AGUA TONICA 350ML', 'Motivo': 'Volume'}]
POWERQUERY_FULL = '\nlet\n    Fonte = Excel.CurrentWorkbook(){[Name="BaseVendas"]}[Content],\n    TiposAlterados = Table.TransformColumnTypes(\n        Fonte,\n        {\n            {"Produto", type text},\n            {"Cliente", type text},\n            {"Valor", type any},\n            {"Data", type any}\n        }\n    ),\n\n    fnRemoveAcentos = (texto as nullable text) as nullable text =>\n        let\n            Entrada = if texto = null then null else texto,\n            Substituicoes = {\n                {"Á","A"},{"À","A"},{"Ã","A"},{"Â","A"},{"Ä","A"},\n                {"É","E"},{"È","E"},{"Ê","E"},{"Ë","E"},\n                {"Í","I"},{"Ì","I"},{"Î","I"},{"Ï","I"},\n                {"Ó","O"},{"Ò","O"},{"Õ","O"},{"Ô","O"},{"Ö","O"},\n                {"Ú","U"},{"Ù","U"},{"Û","U"},{"Ü","U"},\n                {"Ç","C"},\n                {"á","a"},{"à","a"},{"ã","a"},{"â","a"},{"ä","a"},\n                {"é","e"},{"è","e"},{"ê","e"},{"ë","e"},\n                {"í","i"},{"ì","i"},{"î","i"},{"ï","i"},\n                {"ó","o"},{"ò","o"},{"õ","o"},{"ô","o"},{"ö","o"},\n                {"ú","u"},{"ù","u"},{"û","u"},{"ü","u"},\n                {"ç","c"}\n            },\n            Resultado =\n                if Entrada = null then null\n                else List.Accumulate(\n                    Substituicoes,\n                    Entrada,\n                    (estado, atual) => Text.Replace(estado, atual{0}, atual{1})\n                )\n        in\n            Resultado,\n\n    fnTextoLimpoMaiusculo = (texto as nullable text) as nullable text =>\n        let\n            Entrada = if texto = null then null else texto,\n            SemAcentos = if Entrada = null then null else fnRemoveAcentos(Entrada),\n            Limpo = if SemAcentos = null then null else Text.Clean(Text.Trim(SemAcentos)),\n            Maiusculo = if Limpo = null then null else Text.Upper(Limpo),\n            SemPontuacao = if Maiusculo = null then null else\n                Text.Remove(Maiusculo, {".", ",", ";", ":", "-", "_", "/", "\\", "(", ")", "[", "]", "{", "}"}),\n            EspacosNormalizados =\n                if SemPontuacao = null then null\n                else Text.Combine(List.Select(Text.Split(SemPontuacao, " "), each _ <> ""), " ")\n        in\n            EspacosNormalizados,\n\n    fnTextoLimpoMinusculo = (texto as nullable text) as nullable text =>\n        let\n            Base = fnTextoLimpoMaiusculo(texto),\n            Resultado = if Base = null then null else Text.Lower(Base)\n        in\n            Resultado,\n\n    ProdutoOriginal = Table.DuplicateColumn(TiposAlterados, "Produto", "Produto_Original"),\n    ProdutoLimpo = Table.AddColumn(ProdutoOriginal, "Produto_Limpo", each fnTextoLimpoMaiusculo([Produto]), type text),\n    ProdutoLimpoMin = Table.AddColumn(ProdutoLimpo, "Produto_Limpo_Min", each fnTextoLimpoMinusculo([Produto]), type text),\n    ClienteLimpo = Table.AddColumn(ProdutoLimpoMin, "Cliente_Limpo", each fnTextoLimpoMaiusculo([Cliente]), type text),\n\n    SubstituicoesPadrao = {\n        {" LT ", " LATA "},\n        {" LTA ", " LATA "},\n        {" LATAA ", " LATA "},\n        {" CX ", " CAIXA "},\n        {" CERV ", " CERVEJA "},\n        {" LONGNECK ", " LONG NECK "},\n        {" GUARANA ANTARTICA ", " GUARANA ANTARCTICA "},\n        {" BRAHMAA ", " BRAHMA "},\n        {" BRHMA ", " BRAHMA "},\n        {" SKOLL ", " SKOL "}\n    },\n\n    ProdutoPadronizado = Table.AddColumn(\n        ClienteLimpo,\n        "Produto_Padronizado",\n        each\n            let\n                ComEspacos = " " & [Produto_Limpo] & " ",\n                Corrigido = List.Accumulate(\n                    SubstituicoesPadrao,\n                    ComEspacos,\n                    (estado, atual) => Text.Replace(estado, atual{0}, atual{1})\n                ),\n                Final = Text.Trim(Text.Combine(List.Select(Text.Split(Corrigido, " "), each _ <> ""), " "))\n            in\n                Final,\n        type text\n    ),\n\n    MarcaDetectada = Table.AddColumn(\n        ProdutoPadronizado,\n        "Marca_Detectada",\n        each\n            if [Produto_Padronizado] = null then "SEM PRODUTO"\n            else if Text.Contains([Produto_Padronizado], "SKOL") then "SKOL"\n            else if Text.Contains([Produto_Padronizado], "BRAHMA") then "BRAHMA"\n            else if Text.Contains([Produto_Padronizado], "GUARANA") then "GUARANA"\n            else if Text.Contains([Produto_Padronizado], "HEINEKEN") then "HEINEKEN"\n            else "OUTROS",\n        type text\n    ),\n\n    EmbalagemDetectada = Table.AddColumn(\n        MarcaDetectada,\n        "Embalagem_Detectada",\n        each\n            if [Produto_Padronizado] = null then "SEM EMBALAGEM"\n            else if Text.Contains([Produto_Padronizado], "LATA") then "LATA"\n            else if Text.Contains([Produto_Padronizado], "LONG NECK") then "LONG NECK"\n            else if Text.Contains([Produto_Padronizado], "PET") then "PET"\n            else if Text.Contains([Produto_Padronizado], "CAIXA") then "CAIXA"\n            else "OUTROS",\n        type text\n    ),\n\n    VolumeDetectado = Table.AddColumn(\n        EmbalagemDetectada,\n        "Volume_Detectado",\n        each\n            if [Produto_Padronizado] = null then null\n            else if Text.Contains([Produto_Padronizado], "350") then "350ML"\n            else if Text.Contains([Produto_Padronizado], "269") then "269ML"\n            else if Text.Contains([Produto_Padronizado], "600") then "600ML"\n            else if Text.Contains([Produto_Padronizado], "1L") then "1L"\n            else null,\n        type text\n    ),\n\n    DeParaFonte = Excel.CurrentWorkbook(){[Name="DeParaProdutos"]}[Content],\n    DeParaTipos = Table.TransformColumnTypes(DeParaFonte, {{"Grafia_Incorreta", type text}, {"Produto_Correto", type text}}),\n    DeParaLimpo = Table.TransformColumns(\n        DeParaTipos,\n        {\n            {"Grafia_Incorreta", each fnTextoLimpoMaiusculo(_), type text},\n            {"Produto_Correto", each fnTextoLimpoMaiusculo(_), type text}\n        }\n    ),\n    DeParaBuffer = Table.Buffer(DeParaLimpo),\n    MergeCorrecoes = Table.NestedJoin(VolumeDetectado, {"Produto_Padronizado"}, DeParaBuffer, {"Grafia_Incorreta"}, "Correcoes", JoinKind.LeftOuter),\n    Expandido = Table.ExpandTableColumn(MergeCorrecoes, "Correcoes", {"Produto_Correto"}, {"Produto_Correto"}),\n    ProdutoFinal = Table.AddColumn(Expandido, "Produto_Final", each if [Produto_Correto] <> null then [Produto_Correto] else [Produto_Padronizado], type text),\n\n    ValorNumerico = Table.AddColumn(ProdutoFinal, "Valor_Numero", each try Number.From([Valor]) otherwise null, type number),\n    DataConvertida = Table.AddColumn(ValorNumerico, "Data_Convertida", each try Date.From([Data]) otherwise null, type date),\n\n    StatusCorrecao = Table.AddColumn(\n        DataConvertida,\n        "Status_Correcao",\n        each\n            if [Produto] = null or Text.Trim(Text.From([Produto])) = "" then "ERRO: PRODUTO VAZIO"\n            else if [Produto_Correto] <> null then "CORRIGIDO POR DE/PARA"\n            else if [Produto_Limpo] <> [Produto_Padronizado] then "CORRIGIDO POR REGRA"\n            else "SEM ALTERAÇÃO",\n        type text\n    ),\n\n    StatusDQ = Table.AddColumn(\n        StatusCorrecao,\n        "Status_DQ",\n        each\n            if [Produto_Final] = null or [Produto_Final] = "" then "ERRO PRODUTO"\n            else if [Valor_Numero] = null then "ERRO VALOR"\n            else if [Data_Convertida] = null then "ERRO DATA"\n            else "OK",\n        type text\n    ),\n\n    FaixaValor = Table.AddColumn(\n        StatusDQ,\n        "Faixa_Valor",\n        each\n            if [Valor_Numero] = null then "Sem valor"\n            else if [Valor_Numero] >= 1000 then "Alta"\n            else if [Valor_Numero] >= 500 then "Média"\n            else "Baixa",\n        type text\n    ),\n\n    ChaveAnalitica = Table.AddColumn(FaixaValor, "Chave_Produto_Cliente", each [Produto_Final] & "-" & [Cliente_Limpo], type text),\n\n    ResultadoFinal = Table.SelectColumns(\n        ChaveAnalitica,\n        {\n            "Produto_Original",\n            "Produto_Limpo",\n            "Produto_Limpo_Min",\n            "Produto_Padronizado",\n            "Produto_Correto",\n            "Produto_Final",\n            "Marca_Detectada",\n            "Embalagem_Detectada",\n            "Volume_Detectado",\n            "Cliente",\n            "Cliente_Limpo",\n            "Valor",\n            "Valor_Numero",\n            "Data",\n            "Data_Convertida",\n            "Faixa_Valor",\n            "Chave_Produto_Cliente",\n            "Status_Correcao",\n            "Status_DQ"\n        },\n        MissingField.Ignore\n    )\nin\n    ResultadoFinal\n'

TOPICS = {
    "Excel Básico": {
        "nivel": "Básico",
        "tag": "EXCEL-BASE",
        "objetivo": "Organizar dados, limpar textos, validar entradas e preparar base tabular.",
        "conceito": "Uma boa análise começa com base tabular: cabeçalhos únicos, sem mesclagem, tipos coerentes e uma linha por registro.",
        "quando": "Controles simples, análises rápidas, bases pequenas e protótipos.",
        "risco": "Automatizar base mal estruturada aumenta erro e retrabalho.",
        "excel": EXCEL_BASIC,
        "pq": PQ_BASIC,
    },
    "Excel Intermediário": {
        "nivel": "Intermediário",
        "tag": "EXCEL-FORMULAS",
        "objetivo": "Combinar fórmulas, criar buscas robustas, aplicar critérios e tratar erros.",
        "conceito": "O intermediário combina funções: SE, E, OU, SEERRO, PROCX, SOMASES, FILTRO, CLASSIFICAR e LET.",
        "quando": "Relatórios recorrentes, conciliações, controles financeiros e análises operacionais.",
        "risco": "Fórmulas longas sem LET ou documentação ficam frágeis.",
        "excel": EXCEL_INTERMEDIATE,
        "pq": PQ_INTERMEDIATE,
    },
    "Excel Avançado": {
        "nivel": "Avançado",
        "tag": "EXCEL-ADV",
        "objetivo": "Usar matrizes dinâmicas, LAMBDA, LET, coringas e buscas compostas.",
        "conceito": "O avançado resolve problemas complexos com fórmulas auditáveis e performance consciente.",
        "quando": "Testes técnicos, conciliações complexas e modelos sem VBA.",
        "risco": "Nem toda fórmula avançada é a melhor solução; às vezes Power Query é mais seguro.",
        "excel": EXCEL_ADVANCED,
        "pq": PQ_ADVANCED,
    },
    "Coringas": {
        "nivel": "Intermediário/Avançado",
        "tag": "WILDCARDS",
        "objetivo": "Dominar *, ? e ~ em Excel e equivalentes em Power Query.",
        "conceito": "* significa qualquer sequência, ? significa um caractere e ~ escapa o coringa.",
        "quando": "Produtos, descrições, códigos, SKUs e cadastros textuais.",
        "risco": "Coringas podem gerar falso positivo se a regra for ampla demais.",
        "excel": EXCEL_ADVANCED,
        "pq": PQ_BASIC + PQ_INTERMEDIATE,
    },
    "Power Query": {
        "nivel": "Básico ao Avançado",
        "tag": "PQ-M",
        "objetivo": "Automatizar limpeza, transformação, junção, agrupamento e auditoria.",
        "conceito": "Power Query usa linguagem M. Ele prepara dados; DAX calcula métricas no modelo.",
        "quando": "Bases recorrentes, múltiplas fontes, arquivos mensais e dados sujos.",
        "risco": "A ordem das etapas e a qualidade da chave determinam a confiabilidade.",
        "excel": EXCEL_INTERMEDIATE,
        "pq": PQ_BASIC + PQ_INTERMEDIATE + PQ_ADVANCED,
    },
    "Grafias e De/Para": {
        "nivel": "Avançado",
        "tag": "TEXT-DQ",
        "objetivo": "Corrigir caixa, acentos, abreviações, grafias incorretas e manter rastreabilidade.",
        "conceito": "A arquitetura correta preserva original, cria limpo, aplica regra, aplica De/Para e gera status.",
        "quando": "Produtos, clientes, cidades, fornecedores e descrições de sistemas diferentes.",
        "risco": "Correção manual sem tabela De/Para perde auditoria.",
        "excel": EXCEL_ADVANCED,
        "pq": PQ_ADVANCED,
    },
    "Estatística": {
        "nivel": "Intermediário/Avançado",
        "tag": "STAT-EXCEL",
        "objetivo": "Aplicar média, mediana, moda, dispersão, ranking, percentis, correlação e regressão.",
        "conceito": "Estatística transforma relatório em diagnóstico, comparação, tendência e previsão.",
        "quando": "Vendas, forecast, qualidade, demanda, variação e performance.",
        "risco": "Correlação não implica causalidade; regressão precisa de contexto e validação.",
        "excel": STATISTICS,
        "pq": PQ_ADVANCED,
    },
}

STAT_BASE = pd.DataFrame({
    "Mês": pd.date_range("2025-01-01", periods=18, freq="MS"),
    "Investimento": [50, 60, 72, 80, 95, 105, 118, 130, 142, 150, 165, 176, 188, 196, 205, 215, 225, 235],
    "Vendas": [118, 126, 141, 149, 158, 170, 181, 190, 205, 214, 226, 238, 252, 260, 273, 288, 302, 315],
})
STAT_BASE["MediaMovel3"] = STAT_BASE["Vendas"].rolling(3).mean()
media_vendas = STAT_BASE["Vendas"].mean()
desvio_vendas = STAT_BASE["Vendas"].std()
STAT_BASE["ZScore"] = (STAT_BASE["Vendas"] - media_vendas) / desvio_vendas

def card(title, body, cls="card"):
    st.markdown(f"""
    <div class="{cls}">
        <div class="title-small">{title}</div>
        <div>{body}</div>
    </div>
    """, unsafe_allow_html=True)

def df_from_rows(rows):
    return pd.DataFrame(rows)

def render_rows(rows, language):
    for row in rows:
        with st.expander(row["Tema"], expanded=False):
            st.write(row["Uso"])
            st.code(row["Formula"], language=language)

def regression_dataframe():
    x = STAT_BASE["Investimento"]
    y = STAT_BASE["Vendas"]
    x_mean = x.mean()
    y_mean = y.mean()
    slope = ((x - x_mean) * (y - y_mean)).sum() / ((x - x_mean) ** 2).sum()
    intercept = y_mean - slope * x_mean
    result = STAT_BASE.copy()
    result["Tendencia_Linear"] = intercept + slope * result["Investimento"]
    return result

with st.sidebar:
    st.markdown("## 📊 Comitê Técnico")
    st.caption("Excel · Power Query · Coringas · Estatística · Data Quality")
    st.divider()
    selected = st.radio("Mapa Mental", list(TOPICS.keys()), index=list(TOPICS.keys()).index("Estatística"))
    st.divider()
    st.metric("Meta", "Nota 9+")
    st.metric("Cobertura", "Básico → Avançado")
    st.info("Comitê ampliado: Excel, Power Query, Data Quality, Estatística e QA de Deploy.")

active = TOPICS[selected]

st.title("Comitê Técnico — Excel, Power Query e Estatística")
st.caption("Versão v6: sem matplotlib, compatível com Streamlit Cloud, biblioteca ampliada, coringas *, ? e ~, estatística e gráficos nativos.")

st.markdown(
    f'<span class="tag">{active["tag"]}</span>'
    f'<span class="tag">Nível: {active["nivel"]}</span>',
    unsafe_allow_html=True
)

tabs = st.tabs([
    "Visão geral",
    "Biblioteca Excel",
    "Biblioteca Power Query",
    "Coringas",
    "Grafias / De-Para",
    "Estatística",
    "Gráficos nativos",
    "Data Quality",
    "Case técnico",
    "Bloco M completo"
])

with tabs[0]:
    c1, c2 = st.columns(2, gap="large")
    with c1:
        card("Objetivo", active["objetivo"], "card green")
        card("Conceito", active["conceito"], "card blue")
    with c2:
        card("Quando usar", active["quando"], "card yellow")
        card("Risco / mitigação", active["risco"], "card red")

with tabs[1]:
    st.subheader("Tabela de fórmulas")
    st.dataframe(df_from_rows(active["excel"]), use_container_width=True, hide_index=True)
    st.subheader("Explicação item a item")
    render_rows(active["excel"], "text")

with tabs[2]:
    st.subheader("Tabela de códigos M")
    st.dataframe(df_from_rows(active["pq"]), use_container_width=True, hide_index=True)
    st.subheader("Explicação item a item")
    render_rows(active["pq"], "powerquery")

with tabs[3]:
    st.subheader("Coringas do Excel e equivalentes no Power Query")
    st.dataframe(pd.DataFrame(WILDCARDS), use_container_width=True, hide_index=True)
    card("Resumo", "* = qualquer sequência; ? = um caractere; ~ = trata * ou ? como caractere literal.", "card purple")

with tabs[4]:
    st.subheader("Exemplos De/Para")
    st.dataframe(pd.DataFrame(DEPARA_EXAMPLES), use_container_width=True, hide_index=True)
    st.markdown("""
    Camada recomendada:
    1. Produto_Original.
    2. Produto_Limpo_Maiusculo.
    3. Produto_Limpo_Minusculo.
    4. Produto_Sem_Acento.
    5. Produto_Padronizado.
    6. Produto_Final.
    7. Status_Correcao.
    8. Status_DQ.
    """)

with tabs[5]:
    st.subheader("Biblioteca estatística")
    st.dataframe(pd.DataFrame(STATISTICS), use_container_width=True, hide_index=True)
    c1, c2 = st.columns(2)
    with c1:
        card("Tendência central", "MÉDIA, MED e MODO explicam o comportamento típico.", "card blue")
        card("Dispersão", "VAR, DESVPAD e PADRONIZAR medem variabilidade e distância da média.", "card purple")
    with c2:
        card("Posição", "ORDEM, ORDEM.PORCENTUAL, PERCENTIL e QUARTIL criam ranking e cortes.", "card yellow")
        card("Relação e previsão", "CORREL, PEARSON, RQUAD, INCLINAÇÃO, INTERCEPÇÃO, TENDÊNCIA, PROJ.LIN e PROJ.LOG explicam relação e projeção.", "card green")

with tabs[6]:
    st.subheader("Base simulada")
    st.dataframe(STAT_BASE.round(3), use_container_width=True, hide_index=True)
    chart = st.selectbox("Escolha o gráfico nativo", ["Tendência", "Regressão", "Frequência", "Z-score"])
    if chart == "Tendência":
        st.line_chart(STAT_BASE.set_index("Mês")[["Vendas", "MediaMovel3"]])
    elif chart == "Regressão":
        reg = regression_dataframe()
        st.scatter_chart(reg, x="Investimento", y="Vendas")
        st.line_chart(reg.set_index("Investimento")[["Tendencia_Linear"]])
    elif chart == "Frequência":
        bins = pd.cut(STAT_BASE["Vendas"], bins=6)
        freq = STAT_BASE.groupby(bins, observed=False).size().reset_index(name="Frequência")
        freq["Faixa"] = freq["Vendas"].astype(str)
        st.bar_chart(freq.set_index("Faixa")["Frequência"])
    else:
        st.bar_chart(STAT_BASE.set_index("Mês")["ZScore"])

with tabs[7]:
    st.subheader("Data Quality")
    dq_excel = [
        {"Tema": "Vazios", "Formula": '=CONTAR.VAZIO(A:A)', "Uso": "Conta vazios.", "Nivel": "DQ"},
        {"Tema": "Duplicados", "Formula": '=SE(CONT.SE(A:A;A2)>1;"Duplicado";"Único")', "Uso": "Sinaliza duplicidade.", "Nivel": "DQ"},
        {"Tema": "Número inválido", "Formula": '=SEERRO(VALOR(A2);"Erro numérico")', "Uso": "Valida número.", "Nivel": "DQ"},
        {"Tema": "Status", "Formula": '=SE(E(A2<>"";B2>0);"OK";"Validar")', "Uso": "Cria status.", "Nivel": "DQ"},
    ]
    dq_pq = [
        {"Tema": "Profile", "Formula": "Table.Profile(Fonte)", "Uso": "Perfil estatístico da tabela.", "Nivel": "DQ"},
        {"Tema": "Schema", "Formula": "Table.Schema(Fonte)", "Uso": "Estrutura da tabela.", "Nivel": "DQ"},
        {"Tema": "Exceções", "Formula": 'Table.SelectRows(Fonte, each [Status_DQ] <> "OK")', "Uso": "Filtra problemas.", "Nivel": "DQ"},
        {"Tema": "Duplicados", "Formula": 'Table.Group(Fonte, {"Chave"}, {{"Qtd", each Table.RowCount(_), Int64.Type}})', "Uso": "Conta por chave.", "Nivel": "DQ"},
    ]
    st.dataframe(pd.DataFrame(dq_excel + dq_pq), use_container_width=True, hide_index=True)

with tabs[8]:
    st.subheader("Case técnico sugerido")
    st.markdown("""
    Receba uma base de vendas com grafias incorretas, valores inválidos e datas inconsistentes.

    Entregáveis:
    - Base original preservada.
    - Colunas limpas em maiúsculo e minúsculo.
    - Correção com regras e De/Para.
    - Uso documentado de coringas.
    - Indicadores de média, mediana, desvio, percentil e correlação.
    - Gráfico de tendência e regressão.
    - Relatório de exceções.
    - Resumo executivo com riscos e recomendações.
    """)
    card("Resposta de entrevista", "Eu não sobrescrevo a base original. Crio uma camada auditável de limpeza, aplico regras, De/Para, verificações de qualidade e só então gero indicadores e análise estatística.", "card green")

with tabs[9]:
    st.subheader("Power Query M completo")
    st.code(POWERQUERY_FULL, language="powerquery")
    st.download_button(
        "Baixar código M",
        POWERQUERY_FULL,
        file_name="powerquery_grafias_dataquality_v6.m",
        mime="text/plain"
    )

st.divider()
st.caption("v6 auditada: sem matplotlib, sem numpy, sem __file__, compatível com Streamlit Cloud.")


# QA técnico 001: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 002: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 003: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 004: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 005: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 006: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 007: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 008: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 009: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 010: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 011: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 012: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 013: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 014: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 015: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 016: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 017: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 018: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 019: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 020: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 021: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 022: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 023: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 024: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 025: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 026: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 027: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 028: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 029: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 030: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 031: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 032: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 033: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 034: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 035: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 036: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 037: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 038: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 039: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 040: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 041: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 042: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 043: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 044: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 045: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 046: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 047: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 048: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 049: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 050: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 051: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 052: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 053: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 054: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 055: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 056: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 057: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 058: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 059: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 060: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 061: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 062: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 063: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 064: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 065: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 066: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 067: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 068: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 069: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 070: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 071: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 072: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 073: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 074: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 075: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 076: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 077: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 078: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 079: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 080: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 081: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 082: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 083: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 084: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 085: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 086: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 087: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 088: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 089: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 090: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 091: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 092: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 093: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 094: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 095: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 096: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 097: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 098: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 099: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 100: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 101: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 102: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 103: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 104: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 105: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 106: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 107: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 108: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 109: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 110: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 111: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 112: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 113: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 114: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 115: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 116: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 117: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 118: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 119: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 120: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 121: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 122: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 123: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 124: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 125: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 126: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 127: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 128: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 129: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 130: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 131: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 132: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 133: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 134: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 135: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 136: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 137: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 138: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 139: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 140: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 141: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 142: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 143: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 144: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 145: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 146: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 147: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 148: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 149: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 150: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 151: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 152: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 153: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 154: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 155: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 156: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 157: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 158: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 159: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 160: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 161: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 162: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 163: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 164: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 165: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 166: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 167: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 168: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 169: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 170: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 171: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 172: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 173: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 174: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 175: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 176: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 177: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 178: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 179: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 180: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 181: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 182: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 183: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 184: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 185: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 186: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 187: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 188: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 189: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 190: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 191: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 192: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 193: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 194: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 195: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 196: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 197: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 198: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 199: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 200: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 201: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 202: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 203: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 204: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 205: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 206: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 207: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 208: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 209: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 210: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 211: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 212: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 213: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 214: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 215: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 216: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 217: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 218: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 219: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 220: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 221: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 222: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 223: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 224: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 225: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 226: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 227: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 228: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 229: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 230: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 231: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 232: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 233: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 234: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 235: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 236: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 237: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 238: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 239: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 240: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 241: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 242: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 243: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 244: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 245: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 246: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 247: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 248: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 249: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 250: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 251: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 252: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 253: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 254: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 255: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 256: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 257: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 258: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 259: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 260: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 261: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 262: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 263: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 264: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 265: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 266: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 267: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 268: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 269: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 270: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 271: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 272: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 273: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 274: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 275: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 276: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 277: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 278: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 279: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 280: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 281: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 282: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 283: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 284: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 285: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 286: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 287: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 288: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 289: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 290: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 291: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 292: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 293: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 294: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 295: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 296: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 297: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 298: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 299: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 300: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 301: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 302: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 303: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 304: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 305: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 306: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 307: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 308: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 309: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 310: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 311: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 312: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 313: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 314: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 315: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 316: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 317: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 318: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 319: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 320: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 321: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 322: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 323: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 324: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 325: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 326: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 327: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 328: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 329: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 330: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 331: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 332: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 333: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 334: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 335: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 336: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 337: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 338: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 339: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 340: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 341: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 342: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 343: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 344: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 345: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 346: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 347: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 348: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 349: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 350: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 351: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 352: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 353: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 354: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 355: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 356: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 357: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 358: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 359: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 360: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 361: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 362: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 363: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 364: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 365: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 366: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 367: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 368: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 369: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 370: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 371: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 372: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 373: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 374: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 375: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 376: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 377: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 378: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 379: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 380: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 381: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 382: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 383: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 384: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 385: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 386: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 387: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 388: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 389: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 390: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 391: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 392: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 393: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 394: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 395: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 396: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 397: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 398: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 399: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 400: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 401: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 402: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 403: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 404: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 405: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 406: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 407: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 408: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 409: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 410: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 411: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 412: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 413: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 414: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 415: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 416: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 417: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 418: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 419: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 420: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 421: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 422: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 423: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 424: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 425: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 426: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 427: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 428: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 429: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 430: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 431: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 432: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 433: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 434: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 435: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 436: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 437: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 438: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 439: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 440: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 441: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 442: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 443: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 444: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 445: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 446: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 447: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 448: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 449: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 450: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 451: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 452: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 453: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 454: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 455: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 456: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 457: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 458: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 459: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 460: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 461: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 462: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 463: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 464: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 465: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 466: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 467: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 468: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 469: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 470: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 471: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 472: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 473: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 474: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 475: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 476: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 477: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 478: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 479: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 480: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 481: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 482: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 483: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 484: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 485: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 486: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 487: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 488: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 489: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 490: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 491: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 492: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 493: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 494: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 495: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 496: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 497: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 498: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 499: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 500: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 501: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 502: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 503: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 504: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 505: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 506: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 507: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 508: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 509: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 510: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 511: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 512: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 513: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 514: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 515: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 516: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 517: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 518: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 519: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 520: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 521: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 522: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 523: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 524: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 525: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 526: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 527: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 528: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 529: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 530: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 531: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 532: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 533: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 534: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 535: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 536: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 537: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 538: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 539: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 540: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 541: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 542: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 543: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 544: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 545: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 546: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 547: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 548: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 549: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 550: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 551: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 552: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 553: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 554: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 555: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 556: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 557: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 558: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 559: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 560: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 561: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 562: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 563: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 564: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 565: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 566: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 567: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 568: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 569: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 570: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 571: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 572: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 573: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 574: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 575: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 576: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 577: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 578: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 579: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 580: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 581: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 582: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 583: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 584: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 585: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 586: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 587: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 588: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 589: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 590: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 591: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 592: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 593: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 594: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 595: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 596: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 597: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 598: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 599: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 600: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 601: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 602: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 603: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 604: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 605: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 606: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 607: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 608: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 609: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 610: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 611: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 612: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 613: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 614: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 615: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 616: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 617: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 618: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 619: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 620: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 621: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 622: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 623: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 624: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 625: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 626: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 627: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 628: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 629: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 630: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 631: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 632: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 633: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 634: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 635: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 636: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 637: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 638: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 639: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 640: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 641: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 642: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 643: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 644: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 645: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 646: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 647: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 648: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 649: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 650: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 651: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 652: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 653: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 654: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 655: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 656: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 657: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 658: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 659: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 660: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 661: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 662: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 663: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 664: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 665: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 666: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 667: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 668: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 669: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 670: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 671: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 672: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 673: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 674: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 675: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 676: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 677: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 678: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 679: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 680: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 681: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 682: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 683: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 684: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 685: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 686: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 687: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 688: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 689: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 690: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 691: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 692: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 693: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 694: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 695: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 696: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 697: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 698: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 699: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 700: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 701: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 702: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 703: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 704: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 705: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 706: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 707: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 708: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 709: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 710: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 711: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 712: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 713: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 714: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 715: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 716: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 717: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 718: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 719: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 720: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 721: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 722: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 723: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 724: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 725: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 726: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 727: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 728: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 729: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 730: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 731: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 732: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 733: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 734: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 735: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 736: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 737: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 738: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 739: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 740: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 741: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 742: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 743: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 744: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 745: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 746: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 747: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 748: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 749: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 750: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 751: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 752: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 753: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 754: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 755: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 756: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 757: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 758: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 759: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 760: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 761: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 762: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 763: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 764: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 765: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 766: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 767: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 768: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 769: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 770: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 771: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 772: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 773: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 774: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 775: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 776: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 777: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 778: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 779: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 780: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 781: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 782: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 783: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 784: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 785: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 786: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 787: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 788: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 789: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 790: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 791: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 792: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 793: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 794: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 795: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 796: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 797: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 798: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 799: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 800: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 801: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 802: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 803: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 804: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 805: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 806: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 807: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 808: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 809: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 810: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 811: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 812: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 813: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 814: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 815: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 816: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 817: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 818: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 819: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 820: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 821: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 822: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 823: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 824: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 825: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 826: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 827: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 828: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 829: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 830: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 831: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 832: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 833: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 834: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 835: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 836: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 837: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 838: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 839: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 840: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 841: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 842: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 843: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 844: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 845: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 846: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 847: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 848: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 849: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 850: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 851: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 852: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 853: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 854: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 855: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 856: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 857: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 858: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 859: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 860: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 861: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 862: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 863: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 864: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 865: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 866: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 867: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 868: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 869: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 870: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 871: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 872: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 873: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 874: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 875: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 876: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 877: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 878: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 879: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 880: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 881: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 882: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 883: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 884: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 885: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 886: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 887: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 888: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 889: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 890: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 891: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 892: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 893: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 894: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 895: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 896: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 897: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 898: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 899: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 900: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 901: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 902: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 903: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 904: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 905: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 906: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 907: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 908: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 909: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 910: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 911: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 912: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 913: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 914: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 915: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 916: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 917: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 918: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 919: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 920: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 921: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 922: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 923: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 924: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 925: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 926: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 927: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 928: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 929: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 930: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 931: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 932: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 933: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 934: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 935: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 936: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 937: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 938: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 939: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 940: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 941: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 942: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 943: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 944: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 945: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 946: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 947: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 948: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 949: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 950: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 951: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 952: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 953: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 954: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 955: asterisco — coringa para qualquer sequência de caracteres; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 956: interrogação — coringa para um único caractere; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 957: til — escape para buscar coringas como texto literal; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 958: De Para — tabela de correção rastreável; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 959: Data Quality — camada de validação antes do indicador; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 960: fuzzy matching — correspondência aproximada que exige validação; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 961: List.Accumulate — aplica regras em sequência; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 962: Table.Group — agrupamento com agregações; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 963: Table.NestedJoin — merge entre consultas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 964: Table.Profile — perfil estatístico da consulta; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 965: média — medida de tendência central; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 966: mediana — centro robusto contra extremos; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 967: moda — valor mais frequente; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 968: variância — dispersão quadrática; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 969: desvio padrão — dispersão na escala original; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 970: z-score — valor padronizado; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 971: percentil — posição de corte; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 972: quartil — divisão em quatro partes; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 973: correlação — força de relação linear; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 974: regressão — modelo de explicação e previsão; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 975: base tabular — uma linha por registro e uma coluna por atributo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 976: tabela estruturada — permite fórmulas com referências semânticas; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 977: Power Query M — linguagem funcional de transformação de dados; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 978: DAX — linguagem de medidas analíticas no modelo; validar conceito, sintaxe, risco e aplicação em teste técnico.
# QA técnico 979: Power Pivot — modelo de dados relacional dentro do Excel; validar conceito, sintaxe, risco e aplicação em teste técnico.
