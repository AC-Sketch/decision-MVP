
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

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

EXCEL_BASICO = [('Tratamento de vazio', '=SE(A2="";"Sem informação";A2)', 'Substitui célula vazia por texto de controle.'), ('Tratamento de erro', '=SEERRO(A2/B2;0)', 'Evita erro visual e retorna zero quando houver falha.'), ('Texto sem espaços', '=ARRUMAR(A2)', 'Remove espaços excedentes entre palavras.'), ('Maiúscula', '=MAIÚSCULA(A2)', 'Padroniza texto em caixa alta.'), ('Minúscula', '=MINÚSCULA(A2)', 'Padroniza texto em caixa baixa.'), ('Primeiras maiúsculas', '=PRI.MAIÚSCULA(A2)', 'Padroniza nomes próprios.'), ('Tamanho texto', '=NÚM.CARACT(A2)', 'Conta caracteres.'), ('Esquerda', '=ESQUERDA(A2;3)', 'Extrai os 3 primeiros caracteres.'), ('Direita', '=DIREITA(A2;4)', 'Extrai os 4 últimos caracteres.'), ('Texto intermediário', '=EXT.TEXTO(A2;4;6)', 'Extrai parte do texto.'), ('Localizar texto', '=LOCALIZAR("LATA";A2)', 'Localiza posição de um texto.'), ('Substituir', '=SUBSTITUIR(A2;"LT";"LATA")', 'Troca trecho textual.'), ('Concatenar', '=A2&" - "&B2', 'Cria chave ou descrição combinada.'), ('Texto junto', '=TEXTOJUNTAR(" | ";VERDADEIRO;A2:C2)', 'Concatena intervalo ignorando vazios.'), ('Hoje', '=HOJE()', 'Retorna data atual.'), ('Ano', '=ANO(A2)', 'Extrai ano.'), ('Mês', '=MÊS(A2)', 'Extrai mês.'), ('Dia', '=DIA(A2)', 'Extrai dia.'), ('Fim do mês', '=FIMMÊS(A2;0)', 'Retorna último dia do mês.'), ('Data formatada', '=TEXTO(A2;"mmm/aaaa")', 'Converte data em texto de competência.')]
EXCEL_INTERMEDIARIO = [('SE com E', '=SE(E(B2>=1000;C2="Ativo");"Prioritário";"Normal")', 'Classifica com múltiplos critérios obrigatórios.'), ('SE com OU', '=SE(OU(A2="SKOL";A2="BRAHMA");"Cerveja";"Outros")', 'Classifica se qualquer condição for verdadeira.'), ('SE aninhado', '=SE(B2>=1000;"Alta";SE(B2>=500;"Média";"Baixa"))', 'Cria faixas de valor.'), ('PROCX simples', '=PROCX(A2;Produtos[SKU];Produtos[Categoria];"Sem cadastro")', 'Busca moderna com retorno padrão.'), ('PROCX composto', '=PROCX(1;(Produtos[Produto]=A2)*(Produtos[Canal]=B2);Produtos[Categoria];"N/D")', 'Busca por mais de uma condição.'), ('SOMASE', '=SOMASE(Base[Produto];A2;Base[Valor])', 'Soma com um critério.'), ('SOMASES', '=SOMASES(Base[Valor];Base[Produto];A2;Base[Canal];"Online")', 'Soma com múltiplos critérios.'), ('CONT.SE', '=CONT.SE(Base[Produto];A2)', 'Conta ocorrências.'), ('CONT.SES', '=CONT.SES(Base[Produto];A2;Base[Status];"Ativo")', 'Conta com múltiplos critérios.'), ('MÉDIASE', '=MÉDIASE(Base[Produto];A2;Base[Valor])', 'Média por critério.'), ('MÉDIASES', '=MÉDIASES(Base[Valor];Base[Produto];A2;Base[Canal];"Online")', 'Média por múltiplos critérios.'), ('ÚNICO', '=ÚNICO(Base[Produto])', 'Lista valores únicos.'), ('CLASSIFICAR', '=CLASSIFICAR(ÚNICO(Base[Produto]))', 'Ordena lista dinâmica.'), ('FILTRO', '=FILTRO(Base;Base[Valor]>1000;"Sem registros")', 'Filtra base por condição.'), ('Top ordenado', '=PEGAR(CLASSIFICAR(Base;3;-1);10)', 'Retorna top 10 por coluna.'), ('ÍNDICE CORRESP', '=ÍNDICE(Tabela[Valor];CORRESP(A2;Tabela[Produto];0))', 'Alternativa clássica de busca.'), ('ÍNDICE CORRESP composto', '=ÍNDICE(Tabela[Valor];CORRESP(1;(Tabela[Produto]=A2)*(Tabela[Canal]=B2);0))', 'Busca clássica com múltiplas condições.'), ('LET', '=LET(prod;ARRUMAR(MAIÚSCULA(A2));SE(prod="";"VAZIO";prod))', 'Nomeia partes da fórmula.'), ('Validação numérica', '=SEERRO(VALOR(A2);"Erro numérico")', 'Converte texto em número com controle.'), ('Duplicidade', '=SE(CONT.SE(A:A;A2)>1;"Duplicado";"Único")', 'Sinaliza duplicados.'), ('Validação operacional 1', '=SEERRO(SE(B2>0;B2/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 2', '=SEERRO(SE(B3>0;B3/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 3', '=SEERRO(SE(B4>0;B4/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 4', '=SEERRO(SE(B5>0;B5/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 5', '=SEERRO(SE(B6>0;B6/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 6', '=SEERRO(SE(B7>0;B7/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 7', '=SEERRO(SE(B8>0;B8/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 8', '=SEERRO(SE(B9>0;B9/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 9', '=SEERRO(SE(B10>0;B10/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 10', '=SEERRO(SE(B11>0;B11/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 11', '=SEERRO(SE(B12>0;B12/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 12', '=SEERRO(SE(B13>0;B13/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 13', '=SEERRO(SE(B14>0;B14/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 14', '=SEERRO(SE(B15>0;B15/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 15', '=SEERRO(SE(B16>0;B16/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 16', '=SEERRO(SE(B17>0;B17/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 17', '=SEERRO(SE(B18>0;B18/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 18', '=SEERRO(SE(B19>0;B19/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 19', '=SEERRO(SE(B20>0;B20/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 20', '=SEERRO(SE(B21>0;B21/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 21', '=SEERRO(SE(B22>0;B22/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 22', '=SEERRO(SE(B23>0;B23/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 23', '=SEERRO(SE(B24>0;B24/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 24', '=SEERRO(SE(B25>0;B25/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 25', '=SEERRO(SE(B26>0;B26/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 26', '=SEERRO(SE(B27>0;B27/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 27', '=SEERRO(SE(B28>0;B28/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 28', '=SEERRO(SE(B29>0;B29/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 29', '=SEERRO(SE(B30>0;B30/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 30', '=SEERRO(SE(B31>0;B31/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 31', '=SEERRO(SE(B32>0;B32/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 32', '=SEERRO(SE(B33>0;B33/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 33', '=SEERRO(SE(B34>0;B34/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 34', '=SEERRO(SE(B35>0;B35/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 35', '=SEERRO(SE(B36>0;B36/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 36', '=SEERRO(SE(B37>0;B37/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 37', '=SEERRO(SE(B38>0;B38/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 38', '=SEERRO(SE(B39>0;B39/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 39', '=SEERRO(SE(B40>0;B40/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 40', '=SEERRO(SE(B41>0;B41/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 41', '=SEERRO(SE(B42>0;B42/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 42', '=SEERRO(SE(B43>0;B43/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 43', '=SEERRO(SE(B44>0;B44/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 44', '=SEERRO(SE(B45>0;B45/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 45', '=SEERRO(SE(B46>0;B46/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 46', '=SEERRO(SE(B47>0;B47/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 47', '=SEERRO(SE(B48>0;B48/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 48', '=SEERRO(SE(B49>0;B49/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 49', '=SEERRO(SE(B50>0;B50/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 50', '=SEERRO(SE(B51>0;B51/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 51', '=SEERRO(SE(B52>0;B52/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 52', '=SEERRO(SE(B53>0;B53/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 53', '=SEERRO(SE(B54>0;B54/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 54', '=SEERRO(SE(B55>0;B55/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 55', '=SEERRO(SE(B56>0;B56/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 56', '=SEERRO(SE(B57>0;B57/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 57', '=SEERRO(SE(B58>0;B58/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 58', '=SEERRO(SE(B59>0;B59/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 59', '=SEERRO(SE(B60>0;B60/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 60', '=SEERRO(SE(B61>0;B61/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 61', '=SEERRO(SE(B62>0;B62/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 62', '=SEERRO(SE(B63>0;B63/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 63', '=SEERRO(SE(B64>0;B64/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 64', '=SEERRO(SE(B65>0;B65/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 65', '=SEERRO(SE(B66>0;B66/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 66', '=SEERRO(SE(B67>0;B67/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 67', '=SEERRO(SE(B68>0;B68/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 68', '=SEERRO(SE(B69>0;B69/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 69', '=SEERRO(SE(B70>0;B70/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 70', '=SEERRO(SE(B71>0;B71/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 71', '=SEERRO(SE(B72>0;B72/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 72', '=SEERRO(SE(B73>0;B73/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 73', '=SEERRO(SE(B74>0;B74/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 74', '=SEERRO(SE(B75>0;B75/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 75', '=SEERRO(SE(B76>0;B76/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 76', '=SEERRO(SE(B77>0;B77/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 77', '=SEERRO(SE(B78>0;B78/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 78', '=SEERRO(SE(B79>0;B79/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 79', '=SEERRO(SE(B80>0;B80/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.'), ('Validação operacional 80', '=SEERRO(SE(B81>0;B81/SOMA(B:B);0);0)', 'Participação percentual com tratamento de erro.')]
EXCEL_AVANCADO = [('Coringa contém', '=CONT.SE(A:A;"*SKOL*")', 'Asterisco antes e depois significa contém.'), ('Coringa começa', '=CONT.SE(A:A;"SKOL*")', 'Começa com SKOL.'), ('Coringa termina', '=CONT.SE(A:A;"*350ML")', 'Termina com 350ML.'), ('Coringa um caractere', '=CONT.SE(A:A;"SKO?")', '? representa exatamente um caractere.'), ('Escapar asterisco', '=CONT.SE(A:A;"SKOL~*")', 'Procura SKOL* literalmente.'), ('Escapar interrogação', '=CONT.SE(A:A;"SKOL~?")', 'Procura SKOL? literalmente.'), ('PROCX curinga', '=PROCX("*LATA*";Base[Produto];Base[Categoria];"N/D";2)', 'PROCX com modo correspondência curinga.'), ('FILTRO com PROCURAR', '=FILTRO(Base;ÉNÚM(PROCURAR("SKOL";Base[Produto]));"Sem SKOL")', 'Filtra por texto contido.'), ('LET + coringa lógico', '=LET(txt;MAIÚSCULA(ARRUMAR(A2));SE(ÉNÚM(PROCURAR("SKOL";txt));"SKOL";txt))', 'Limpa e classifica no mesmo cálculo.'), ('LAMBDA inline', '=LAMBDA(txt;MAIÚSCULA(ARRUMAR(txt)))(A2)', 'Cria função reutilizável.'), ('MAP', '=MAP(A2:A10;LAMBDA(x;MAIÚSCULA(ARRUMAR(x))))', 'Aplica função em cada item.'), ('BYROW', '=BYROW(A2:C10;LAMBDA(linha;SOMA(linha)))', 'Calcula por linha.'), ('REDUZIR', '=REDUZIR(0;Base[Valor];LAMBDA(acum;valor;acum+valor))', 'Acumula valores.'), ('SCAN', '=SCAN(0;Base[Valor];LAMBDA(acum;valor;acum+valor))', 'Retorna acumulado progressivo.'), ('EMPILHARV', '=EMPILHARV(TabelaJan;TabelaFev;TabelaMar)', 'Empilha tabelas verticalmente.'), ('EMPILHARH', '=EMPILHARH(TabelaProdutos;TabelaCategorias)', 'Empilha tabelas horizontalmente.'), ('ESCOLHERCOLS', '=ESCOLHERCOLS(Base;1;3;5)', 'Seleciona colunas.'), ('DESCARTAR', '=DESCARTAR(Base;1)', 'Remove primeiras linhas/colunas.'), ('TOMAR', '=PEGAR(Base;10)', 'Retorna primeiras linhas.'), ('Classificar e filtrar', '=CLASSIFICAR(FILTRO(Base;Base[Valor]>1000);3;-1)', 'Combina filtro e ordenação.'), ('LET auditoria 1', '=LET(v;B2;limite;100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 2', '=LET(v;B3;limite;200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 3', '=LET(v;B4;limite;300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 4', '=LET(v;B5;limite;400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 5', '=LET(v;B6;limite;500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 6', '=LET(v;B7;limite;600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 7', '=LET(v;B8;limite;700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 8', '=LET(v;B9;limite;800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 9', '=LET(v;B10;limite;900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 10', '=LET(v;B11;limite;1000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 11', '=LET(v;B12;limite;1100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 12', '=LET(v;B13;limite;1200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 13', '=LET(v;B14;limite;1300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 14', '=LET(v;B15;limite;1400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 15', '=LET(v;B16;limite;1500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 16', '=LET(v;B17;limite;1600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 17', '=LET(v;B18;limite;1700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 18', '=LET(v;B19;limite;1800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 19', '=LET(v;B20;limite;1900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 20', '=LET(v;B21;limite;2000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 21', '=LET(v;B22;limite;2100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 22', '=LET(v;B23;limite;2200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 23', '=LET(v;B24;limite;2300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 24', '=LET(v;B25;limite;2400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 25', '=LET(v;B26;limite;2500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 26', '=LET(v;B27;limite;2600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 27', '=LET(v;B28;limite;2700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 28', '=LET(v;B29;limite;2800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 29', '=LET(v;B30;limite;2900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 30', '=LET(v;B31;limite;3000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 31', '=LET(v;B32;limite;3100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 32', '=LET(v;B33;limite;3200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 33', '=LET(v;B34;limite;3300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 34', '=LET(v;B35;limite;3400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 35', '=LET(v;B36;limite;3500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 36', '=LET(v;B37;limite;3600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 37', '=LET(v;B38;limite;3700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 38', '=LET(v;B39;limite;3800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 39', '=LET(v;B40;limite;3900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 40', '=LET(v;B41;limite;4000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 41', '=LET(v;B42;limite;4100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 42', '=LET(v;B43;limite;4200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 43', '=LET(v;B44;limite;4300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 44', '=LET(v;B45;limite;4400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 45', '=LET(v;B46;limite;4500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 46', '=LET(v;B47;limite;4600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 47', '=LET(v;B48;limite;4700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 48', '=LET(v;B49;limite;4800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 49', '=LET(v;B50;limite;4900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 50', '=LET(v;B51;limite;5000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 51', '=LET(v;B52;limite;5100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 52', '=LET(v;B53;limite;5200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 53', '=LET(v;B54;limite;5300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 54', '=LET(v;B55;limite;5400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 55', '=LET(v;B56;limite;5500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 56', '=LET(v;B57;limite;5600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 57', '=LET(v;B58;limite;5700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 58', '=LET(v;B59;limite;5800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 59', '=LET(v;B60;limite;5900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.'), ('LET auditoria 60', '=LET(v;B61;limite;6000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso de LET para fórmula auditável.')]
ESTATISTICA = [('Média', '=MÉDIA(B:B)', 'Tendência central sensível a extremos.'), ('Média condicional', '=MÉDIASE(A:A;"SKOL";B:B)', 'Média para um critério.'), ('Múltiplas médias condicionais', '=MÉDIASES(B:B;A:A;"SKOL";C:C;"Online")', 'Média com vários critérios.'), ('Mediana', '=MED(B:B)', 'Centro robusto contra extremos.'), ('Moda única', '=MODO.ÚNICO(B:B)', 'Valor mais frequente.'), ('Variância amostral', '=VAR.S(B:B)', 'Dispersão da amostra.'), ('Variância populacional', '=VAR.P(B:B)', 'Dispersão da população.'), ('Desvio padrão amostral', '=DESVPAD.S(B:B)', 'Volatilidade amostral.'), ('Desvio padrão populacional', '=DESVPAD.P(B:B)', 'Volatilidade populacional.'), ('Padronizar', '=PADRONIZAR(B2;MÉDIA(B:B);DESVPAD.S(B:B))', 'Z-score.'), ('Ordem', '=ORDEM.EQ(B2;B:B;0)', 'Ranking.'), ('Ordem percentual', '=ORDEM.PORCENTUAL.INC(B:B;B2)', 'Posição percentual.'), ('Percentil', '=PERCENTIL.INC(B:B;0,9)', 'Corte de percentil.'), ('Quartil', '=QUARTIL.INC(B:B;3)', 'Corte em quartis.'), ('Correlação', '=CORREL(B:B;C:C)', 'Relação linear.'), ('Pearson', '=PEARSON(B:B;C:C)', 'Correlação de Pearson.'), ('R quadrado', '=RQUAD(B:B;C:C)', 'Quanto X explica Y.'), ('Inclinação', '=INCLINAÇÃO(B:B;C:C)', 'Coeficiente angular.'), ('Intercepção', '=INTERCEPÇÃO(B:B;C:C)', 'Intercepto da reta.'), ('Erro padrão Y', '=EPADYX(B:B;C:C)', 'Erro padrão da estimativa.'), ('Frequência', '=FREQUÊNCIA(B:B;E2:E6)', 'Distribuição por faixas.'), ('Tendência', '=TENDÊNCIA(B:B;C:C;D2:D10)', 'Previsão linear.'), ('Projeção linear', '=PROJ.LIN(B:B;C:C;VERDADEIRO;VERDADEIRO)', 'Regressão linear detalhada.'), ('Projeção logarítmica', '=PROJ.LOG(B:B;C:C;VERDADEIRO;VERDADEIRO)', 'Modelo exponencial/log.'), ('Indicador estatístico aplicado 1', '=MÉDIASES(Base[Valor];Base[Produto];A2;Base[Canal];B2)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 2', '=MÉDIASES(Base[Valor];Base[Produto];A3;Base[Canal];B3)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 3', '=MÉDIASES(Base[Valor];Base[Produto];A4;Base[Canal];B4)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 4', '=MÉDIASES(Base[Valor];Base[Produto];A5;Base[Canal];B5)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 5', '=MÉDIASES(Base[Valor];Base[Produto];A6;Base[Canal];B6)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 6', '=MÉDIASES(Base[Valor];Base[Produto];A7;Base[Canal];B7)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 7', '=MÉDIASES(Base[Valor];Base[Produto];A8;Base[Canal];B8)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 8', '=MÉDIASES(Base[Valor];Base[Produto];A9;Base[Canal];B9)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 9', '=MÉDIASES(Base[Valor];Base[Produto];A10;Base[Canal];B10)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 10', '=MÉDIASES(Base[Valor];Base[Produto];A11;Base[Canal];B11)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 11', '=MÉDIASES(Base[Valor];Base[Produto];A12;Base[Canal];B12)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 12', '=MÉDIASES(Base[Valor];Base[Produto];A13;Base[Canal];B13)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 13', '=MÉDIASES(Base[Valor];Base[Produto];A14;Base[Canal];B14)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 14', '=MÉDIASES(Base[Valor];Base[Produto];A15;Base[Canal];B15)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 15', '=MÉDIASES(Base[Valor];Base[Produto];A16;Base[Canal];B16)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 16', '=MÉDIASES(Base[Valor];Base[Produto];A17;Base[Canal];B17)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 17', '=MÉDIASES(Base[Valor];Base[Produto];A18;Base[Canal];B18)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 18', '=MÉDIASES(Base[Valor];Base[Produto];A19;Base[Canal];B19)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 19', '=MÉDIASES(Base[Valor];Base[Produto];A20;Base[Canal];B20)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 20', '=MÉDIASES(Base[Valor];Base[Produto];A21;Base[Canal];B21)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 21', '=MÉDIASES(Base[Valor];Base[Produto];A22;Base[Canal];B22)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 22', '=MÉDIASES(Base[Valor];Base[Produto];A23;Base[Canal];B23)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 23', '=MÉDIASES(Base[Valor];Base[Produto];A24;Base[Canal];B24)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 24', '=MÉDIASES(Base[Valor];Base[Produto];A25;Base[Canal];B25)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 25', '=MÉDIASES(Base[Valor];Base[Produto];A26;Base[Canal];B26)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 26', '=MÉDIASES(Base[Valor];Base[Produto];A27;Base[Canal];B27)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 27', '=MÉDIASES(Base[Valor];Base[Produto];A28;Base[Canal];B28)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 28', '=MÉDIASES(Base[Valor];Base[Produto];A29;Base[Canal];B29)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 29', '=MÉDIASES(Base[Valor];Base[Produto];A30;Base[Canal];B30)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 30', '=MÉDIASES(Base[Valor];Base[Produto];A31;Base[Canal];B31)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 31', '=MÉDIASES(Base[Valor];Base[Produto];A32;Base[Canal];B32)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 32', '=MÉDIASES(Base[Valor];Base[Produto];A33;Base[Canal];B33)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 33', '=MÉDIASES(Base[Valor];Base[Produto];A34;Base[Canal];B34)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 34', '=MÉDIASES(Base[Valor];Base[Produto];A35;Base[Canal];B35)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 35', '=MÉDIASES(Base[Valor];Base[Produto];A36;Base[Canal];B36)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 36', '=MÉDIASES(Base[Valor];Base[Produto];A37;Base[Canal];B37)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 37', '=MÉDIASES(Base[Valor];Base[Produto];A38;Base[Canal];B38)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 38', '=MÉDIASES(Base[Valor];Base[Produto];A39;Base[Canal];B39)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 39', '=MÉDIASES(Base[Valor];Base[Produto];A40;Base[Canal];B40)', 'Média segmentada para análise de desempenho.'), ('Indicador estatístico aplicado 40', '=MÉDIASES(Base[Valor];Base[Produto];A41;Base[Canal];B41)', 'Média segmentada para análise de desempenho.')]
PQ_BASICO = [('Trim', 'Text.Trim([Produto])', 'Remove espaços externos.'), ('Clean', 'Text.Clean([Produto])', 'Remove caracteres não imprimíveis.'), ('Upper', 'Text.Upper([Produto])', 'Maiúsculo.'), ('Lower', 'Text.Lower([Produto])', 'Minúsculo.'), ('Proper', 'Text.Proper([Cliente])', 'Nome próprio.'), ('Contains', 'Text.Contains([Produto], "SKOL")', 'Contém texto.'), ('StartsWith', 'Text.StartsWith([Produto], "SKOL")', 'Começa com texto.'), ('EndsWith', 'Text.EndsWith([Produto], "350ML")', 'Termina com texto.'), ('Text Length', 'Text.Length([Codigo])', 'Tamanho do texto.'), ('Text Replace', 'Text.Replace([Produto], "LT", "LATA")', 'Substituição textual.'), ('Text Remove', 'Text.Remove([Produto], {".", ",", "-", "_"})', 'Remove pontuação.'), ('Date From', 'Date.From([Data])', 'Converte para data.'), ('Number From', 'try Number.From([Valor]) otherwise null', 'Converte número com segurança.'), ('Duration Days', 'Duration.Days([DataFim] - [DataInicio])', 'Diferença entre datas.')]
PQ_INTERMEDIARIO = [('Select Rows', 'Table.SelectRows(Fonte, each [Valor] > 1000)', 'Filtra linhas.'), ('Add Column', 'Table.AddColumn(Fonte, "Faixa", each if [Valor] >= 1000 then "Alta" else "Baixa", type text)', 'Adiciona coluna calculada.'), ('Group', 'Table.Group(Fonte, {"Produto"}, {{"Total", each List.Sum([Valor]), type number}})', 'Agrupa e soma.'), ('Nested Join', 'Table.NestedJoin(Base, {"Produto_Limpo"}, DePara, {"Grafia_Incorreta"}, "Correcoes", JoinKind.LeftOuter)', 'Merge.'), ('Expand', 'Table.ExpandTableColumn(Merge, "Correcoes", {"Produto_Correto"}, {"Produto_Correto"})', 'Expande merge.'), ('Combine', 'Table.Combine({BaseJan, BaseFev, BaseMar})', 'Append.'), ('Distinct', 'Table.Distinct(Fonte, {"Chave"})', 'Remove duplicados.'), ('Sort', 'Table.Sort(Fonte, {{"Valor", Order.Descending}})', 'Ordena.'), ('Replace Value', 'Table.ReplaceValue(Fonte, "SKOLL", "SKOL", Replacer.ReplaceText, {"Produto"})', 'Substitui valor.'), ('Column Types', 'Table.TransformColumnTypes(Fonte, {{"Data", type date}, {"Valor", type number}})', 'Ajusta tipos.'), ('Filtro operacional 1', 'Table.SelectRows(Fonte, each [Valor] > 10)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 2', 'Table.SelectRows(Fonte, each [Valor] > 20)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 3', 'Table.SelectRows(Fonte, each [Valor] > 30)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 4', 'Table.SelectRows(Fonte, each [Valor] > 40)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 5', 'Table.SelectRows(Fonte, each [Valor] > 50)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 6', 'Table.SelectRows(Fonte, each [Valor] > 60)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 7', 'Table.SelectRows(Fonte, each [Valor] > 70)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 8', 'Table.SelectRows(Fonte, each [Valor] > 80)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 9', 'Table.SelectRows(Fonte, each [Valor] > 90)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 10', 'Table.SelectRows(Fonte, each [Valor] > 100)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 11', 'Table.SelectRows(Fonte, each [Valor] > 110)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 12', 'Table.SelectRows(Fonte, each [Valor] > 120)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 13', 'Table.SelectRows(Fonte, each [Valor] > 130)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 14', 'Table.SelectRows(Fonte, each [Valor] > 140)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 15', 'Table.SelectRows(Fonte, each [Valor] > 150)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 16', 'Table.SelectRows(Fonte, each [Valor] > 160)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 17', 'Table.SelectRows(Fonte, each [Valor] > 170)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 18', 'Table.SelectRows(Fonte, each [Valor] > 180)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 19', 'Table.SelectRows(Fonte, each [Valor] > 190)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 20', 'Table.SelectRows(Fonte, each [Valor] > 200)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 21', 'Table.SelectRows(Fonte, each [Valor] > 210)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 22', 'Table.SelectRows(Fonte, each [Valor] > 220)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 23', 'Table.SelectRows(Fonte, each [Valor] > 230)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 24', 'Table.SelectRows(Fonte, each [Valor] > 240)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 25', 'Table.SelectRows(Fonte, each [Valor] > 250)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 26', 'Table.SelectRows(Fonte, each [Valor] > 260)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 27', 'Table.SelectRows(Fonte, each [Valor] > 270)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 28', 'Table.SelectRows(Fonte, each [Valor] > 280)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 29', 'Table.SelectRows(Fonte, each [Valor] > 290)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 30', 'Table.SelectRows(Fonte, each [Valor] > 300)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 31', 'Table.SelectRows(Fonte, each [Valor] > 310)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 32', 'Table.SelectRows(Fonte, each [Valor] > 320)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 33', 'Table.SelectRows(Fonte, each [Valor] > 330)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 34', 'Table.SelectRows(Fonte, each [Valor] > 340)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 35', 'Table.SelectRows(Fonte, each [Valor] > 350)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 36', 'Table.SelectRows(Fonte, each [Valor] > 360)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 37', 'Table.SelectRows(Fonte, each [Valor] > 370)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 38', 'Table.SelectRows(Fonte, each [Valor] > 380)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 39', 'Table.SelectRows(Fonte, each [Valor] > 390)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 40', 'Table.SelectRows(Fonte, each [Valor] > 400)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 41', 'Table.SelectRows(Fonte, each [Valor] > 410)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 42', 'Table.SelectRows(Fonte, each [Valor] > 420)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 43', 'Table.SelectRows(Fonte, each [Valor] > 430)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 44', 'Table.SelectRows(Fonte, each [Valor] > 440)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 45', 'Table.SelectRows(Fonte, each [Valor] > 450)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 46', 'Table.SelectRows(Fonte, each [Valor] > 460)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 47', 'Table.SelectRows(Fonte, each [Valor] > 470)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 48', 'Table.SelectRows(Fonte, each [Valor] > 480)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 49', 'Table.SelectRows(Fonte, each [Valor] > 490)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 50', 'Table.SelectRows(Fonte, each [Valor] > 500)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 51', 'Table.SelectRows(Fonte, each [Valor] > 510)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 52', 'Table.SelectRows(Fonte, each [Valor] > 520)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 53', 'Table.SelectRows(Fonte, each [Valor] > 530)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 54', 'Table.SelectRows(Fonte, each [Valor] > 540)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 55', 'Table.SelectRows(Fonte, each [Valor] > 550)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 56', 'Table.SelectRows(Fonte, each [Valor] > 560)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 57', 'Table.SelectRows(Fonte, each [Valor] > 570)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 58', 'Table.SelectRows(Fonte, each [Valor] > 580)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 59', 'Table.SelectRows(Fonte, each [Valor] > 590)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 60', 'Table.SelectRows(Fonte, each [Valor] > 600)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 61', 'Table.SelectRows(Fonte, each [Valor] > 610)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 62', 'Table.SelectRows(Fonte, each [Valor] > 620)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 63', 'Table.SelectRows(Fonte, each [Valor] > 630)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 64', 'Table.SelectRows(Fonte, each [Valor] > 640)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 65', 'Table.SelectRows(Fonte, each [Valor] > 650)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 66', 'Table.SelectRows(Fonte, each [Valor] > 660)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 67', 'Table.SelectRows(Fonte, each [Valor] > 670)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 68', 'Table.SelectRows(Fonte, each [Valor] > 680)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 69', 'Table.SelectRows(Fonte, each [Valor] > 690)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 70', 'Table.SelectRows(Fonte, each [Valor] > 700)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 71', 'Table.SelectRows(Fonte, each [Valor] > 710)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 72', 'Table.SelectRows(Fonte, each [Valor] > 720)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 73', 'Table.SelectRows(Fonte, each [Valor] > 730)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 74', 'Table.SelectRows(Fonte, each [Valor] > 740)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 75', 'Table.SelectRows(Fonte, each [Valor] > 750)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 76', 'Table.SelectRows(Fonte, each [Valor] > 760)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 77', 'Table.SelectRows(Fonte, each [Valor] > 770)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 78', 'Table.SelectRows(Fonte, each [Valor] > 780)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 79', 'Table.SelectRows(Fonte, each [Valor] > 790)', 'Filtro paramétrico para treino técnico.'), ('Filtro operacional 80', 'Table.SelectRows(Fonte, each [Valor] > 800)', 'Filtro paramétrico para treino técnico.')]
PQ_AVANCADO = [('List Accumulate', 'List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))', 'Aplica várias substituições.'), ('Fuzzy Join', 'Table.FuzzyNestedJoin(Base, {"Produto_Limpo"}, Cadastro, {"Produto_Limpo"}, "Match", JoinKind.LeftOuter, [IgnoreCase=true, IgnoreSpace=true, Threshold=0.82])', 'Correspondência aproximada.'), ('Buffer', 'Table.Buffer(DeParaLimpo)', 'Evita reprocessamento.'), ('Profile', 'Table.Profile(Fonte)', 'Perfil da base.'), ('Schema', 'Table.Schema(Fonte)', 'Estrutura da base.'), ('List Contains', 'List.Contains({"SKOL", "BRAHMA", "GUARANA"}, [Marca])', 'Validação por lista.'), ('Try Otherwise Date', 'try Date.From([Data]) otherwise null', 'Conversão segura de data.'), ('Try Otherwise Number', 'try Number.From([Valor]) otherwise null', 'Conversão segura de número.'), ('MissingField Ignore', 'Table.SelectColumns(Fonte, {"A", "B"}, MissingField.Ignore)', 'Evita erro por coluna ausente.'), ('Unpivot', 'Table.UnpivotOtherColumns(Fonte, {"Produto"}, "Mês", "Valor")', 'Desnormaliza colunas em linhas.'), ('Pivot', 'Table.Pivot(Fonte, List.Distinct(Fonte[Mês]), "Mês", "Valor", List.Sum)', 'Transforma linhas em colunas.'), ('Regra condicional avançada 1', 'Table.AddColumn(Fonte, "Regra_1", each if [Valor] >= 100 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 2', 'Table.AddColumn(Fonte, "Regra_2", each if [Valor] >= 200 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 3', 'Table.AddColumn(Fonte, "Regra_3", each if [Valor] >= 300 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 4', 'Table.AddColumn(Fonte, "Regra_4", each if [Valor] >= 400 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 5', 'Table.AddColumn(Fonte, "Regra_5", each if [Valor] >= 500 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 6', 'Table.AddColumn(Fonte, "Regra_6", each if [Valor] >= 600 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 7', 'Table.AddColumn(Fonte, "Regra_7", each if [Valor] >= 700 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 8', 'Table.AddColumn(Fonte, "Regra_8", each if [Valor] >= 800 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 9', 'Table.AddColumn(Fonte, "Regra_9", each if [Valor] >= 900 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 10', 'Table.AddColumn(Fonte, "Regra_10", each if [Valor] >= 1000 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 11', 'Table.AddColumn(Fonte, "Regra_11", each if [Valor] >= 1100 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 12', 'Table.AddColumn(Fonte, "Regra_12", each if [Valor] >= 1200 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 13', 'Table.AddColumn(Fonte, "Regra_13", each if [Valor] >= 1300 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 14', 'Table.AddColumn(Fonte, "Regra_14", each if [Valor] >= 1400 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 15', 'Table.AddColumn(Fonte, "Regra_15", each if [Valor] >= 1500 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 16', 'Table.AddColumn(Fonte, "Regra_16", each if [Valor] >= 1600 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 17', 'Table.AddColumn(Fonte, "Regra_17", each if [Valor] >= 1700 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 18', 'Table.AddColumn(Fonte, "Regra_18", each if [Valor] >= 1800 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 19', 'Table.AddColumn(Fonte, "Regra_19", each if [Valor] >= 1900 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 20', 'Table.AddColumn(Fonte, "Regra_20", each if [Valor] >= 2000 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 21', 'Table.AddColumn(Fonte, "Regra_21", each if [Valor] >= 2100 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 22', 'Table.AddColumn(Fonte, "Regra_22", each if [Valor] >= 2200 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 23', 'Table.AddColumn(Fonte, "Regra_23", each if [Valor] >= 2300 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 24', 'Table.AddColumn(Fonte, "Regra_24", each if [Valor] >= 2400 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 25', 'Table.AddColumn(Fonte, "Regra_25", each if [Valor] >= 2500 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 26', 'Table.AddColumn(Fonte, "Regra_26", each if [Valor] >= 2600 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 27', 'Table.AddColumn(Fonte, "Regra_27", each if [Valor] >= 2700 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 28', 'Table.AddColumn(Fonte, "Regra_28", each if [Valor] >= 2800 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 29', 'Table.AddColumn(Fonte, "Regra_29", each if [Valor] >= 2900 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 30', 'Table.AddColumn(Fonte, "Regra_30", each if [Valor] >= 3000 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 31', 'Table.AddColumn(Fonte, "Regra_31", each if [Valor] >= 3100 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 32', 'Table.AddColumn(Fonte, "Regra_32", each if [Valor] >= 3200 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 33', 'Table.AddColumn(Fonte, "Regra_33", each if [Valor] >= 3300 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 34', 'Table.AddColumn(Fonte, "Regra_34", each if [Valor] >= 3400 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 35', 'Table.AddColumn(Fonte, "Regra_35", each if [Valor] >= 3500 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 36', 'Table.AddColumn(Fonte, "Regra_36", each if [Valor] >= 3600 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 37', 'Table.AddColumn(Fonte, "Regra_37", each if [Valor] >= 3700 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 38', 'Table.AddColumn(Fonte, "Regra_38", each if [Valor] >= 3800 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 39', 'Table.AddColumn(Fonte, "Regra_39", each if [Valor] >= 3900 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 40', 'Table.AddColumn(Fonte, "Regra_40", each if [Valor] >= 4000 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 41', 'Table.AddColumn(Fonte, "Regra_41", each if [Valor] >= 4100 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 42', 'Table.AddColumn(Fonte, "Regra_42", each if [Valor] >= 4200 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 43', 'Table.AddColumn(Fonte, "Regra_43", each if [Valor] >= 4300 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 44', 'Table.AddColumn(Fonte, "Regra_44", each if [Valor] >= 4400 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 45', 'Table.AddColumn(Fonte, "Regra_45", each if [Valor] >= 4500 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 46', 'Table.AddColumn(Fonte, "Regra_46", each if [Valor] >= 4600 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 47', 'Table.AddColumn(Fonte, "Regra_47", each if [Valor] >= 4700 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 48', 'Table.AddColumn(Fonte, "Regra_48", each if [Valor] >= 4800 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 49', 'Table.AddColumn(Fonte, "Regra_49", each if [Valor] >= 4900 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 50', 'Table.AddColumn(Fonte, "Regra_50", each if [Valor] >= 5000 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 51', 'Table.AddColumn(Fonte, "Regra_51", each if [Valor] >= 5100 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 52', 'Table.AddColumn(Fonte, "Regra_52", each if [Valor] >= 5200 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 53', 'Table.AddColumn(Fonte, "Regra_53", each if [Valor] >= 5300 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 54', 'Table.AddColumn(Fonte, "Regra_54", each if [Valor] >= 5400 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 55', 'Table.AddColumn(Fonte, "Regra_55", each if [Valor] >= 5500 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 56', 'Table.AddColumn(Fonte, "Regra_56", each if [Valor] >= 5600 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 57', 'Table.AddColumn(Fonte, "Regra_57", each if [Valor] >= 5700 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 58', 'Table.AddColumn(Fonte, "Regra_58", each if [Valor] >= 5800 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 59', 'Table.AddColumn(Fonte, "Regra_59", each if [Valor] >= 5900 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.'), ('Regra condicional avançada 60', 'Table.AddColumn(Fonte, "Regra_60", each if [Valor] >= 6000 then "OK" else "VALIDAR", type text)', 'Criação de regra auditável.')]
WILDCARDS = [('*', 'Qualquer sequência de caracteres', '=CONT.SE(A:A;"*SKOL*")', 'Text.Contains([Produto], "SKOL")'), ('*', 'Começa com', '=CONT.SE(A:A;"SKOL*")', 'Text.StartsWith([Produto], "SKOL")'), ('*', 'Termina com', '=CONT.SE(A:A;"*350ML")', 'Text.EndsWith([Produto], "350ML")'), ('?', 'Um caractere', '=CONT.SE(A:A;"SKO?")', 'Text.StartsWith([Codigo], "SKO") and Text.Length([Codigo]) = 4'), ('???', 'Três caracteres exatos', '=CONT.SE(A:A;"???")', 'Text.Length([Codigo]) = 3'), ('~*', 'Asterisco literal', '=CONT.SE(A:A;"SKOL~*")', 'Text.Contains([Produto], "SKOL*")'), ('~?', 'Interrogação literal', '=CONT.SE(A:A;"SKOL~?")', 'Text.Contains([Produto], "SKOL?")'), ('~~', 'Til literal', '=CONT.SE(A:A;"SKU~~01")', 'Text.Contains([Produto], "SKU~01")')]
DEPARA_EXAMPLES = [('SKOLL', 'SKOL', 'Grafia duplicada'), ('BRAHMAA', 'BRAHMA', 'Letra excedente'), ('BRHMA', 'BRAHMA', 'Letra faltando'), ('GUARANA ANTARTICA', 'GUARANA ANTARCTICA', 'Grafia comercial'), ('SKOL LATAA', 'SKOL LATA', 'Letra excedente'), ('CERV PILSEN', 'CERVEJA PILSEN', 'Abreviação'), ('LONGNECK', 'LONG NECK', 'Espaçamento'), ('LT', 'LATA', 'Abreviação'), ('CX', 'CAIXA', 'Abreviação'), ('AGUA TONICA 350', 'AGUA TONICA 350ML', 'Volume')]
POWERQUERY_FULL = '\nlet\n    Fonte = Excel.CurrentWorkbook(){[Name="BaseVendas"]}[Content],\n\n    TiposAlterados = Table.TransformColumnTypes(\n        Fonte,\n        {\n            {"Produto", type text},\n            {"Cliente", type text},\n            {"Valor", type any},\n            {"Data", type any}\n        }\n    ),\n\n    fnRemoveAcentos = (texto as nullable text) as nullable text =>\n        let\n            Entrada = if texto = null then null else texto,\n            Substituicoes = {\n                {"Á","A"},{"À","A"},{"Ã","A"},{"Â","A"},{"Ä","A"},\n                {"É","E"},{"È","E"},{"Ê","E"},{"Ë","E"},\n                {"Í","I"},{"Ì","I"},{"Î","I"},{"Ï","I"},\n                {"Ó","O"},{"Ò","O"},{"Õ","O"},{"Ô","O"},{"Ö","O"},\n                {"Ú","U"},{"Ù","U"},{"Û","U"},{"Ü","U"},\n                {"Ç","C"},\n                {"á","a"},{"à","a"},{"ã","a"},{"â","a"},{"ä","a"},\n                {"é","e"},{"è","e"},{"ê","e"},{"ë","e"},\n                {"í","i"},{"ì","i"},{"î","i"},{"ï","i"},\n                {"ó","o"},{"ò","o"},{"õ","o"},{"ô","o"},{"ö","o"},\n                {"ú","u"},{"ù","u"},{"û","u"},{"ü","u"},\n                {"ç","c"}\n            },\n            Resultado =\n                if Entrada = null then null\n                else List.Accumulate(\n                    Substituicoes,\n                    Entrada,\n                    (estado, atual) => Text.Replace(estado, atual{0}, atual{1})\n                )\n        in\n            Resultado,\n\n    fnTextoLimpoMaiusculo = (texto as nullable text) as nullable text =>\n        let\n            Entrada = if texto = null then null else texto,\n            SemAcentos = if Entrada = null then null else fnRemoveAcentos(Entrada),\n            Limpo = if SemAcentos = null then null else Text.Clean(Text.Trim(SemAcentos)),\n            Maiusculo = if Limpo = null then null else Text.Upper(Limpo),\n            SemPontuacao = if Maiusculo = null then null else\n                Text.Remove(Maiusculo, {".", ",", ";", ":", "-", "_", "/", "\\", "(", ")", "[", "]", "{", "}"}),\n            EspacosNormalizados =\n                if SemPontuacao = null then null\n                else Text.Combine(List.Select(Text.Split(SemPontuacao, " "), each _ <> ""), " ")\n        in\n            EspacosNormalizados,\n\n    fnTextoLimpoMinusculo = (texto as nullable text) as nullable text =>\n        let\n            Base = fnTextoLimpoMaiusculo(texto),\n            Resultado = if Base = null then null else Text.Lower(Base)\n        in\n            Resultado,\n\n    ProdutoOriginal = Table.DuplicateColumn(\n        TiposAlterados,\n        "Produto",\n        "Produto_Original"\n    ),\n\n    ProdutoLimpo = Table.AddColumn(\n        ProdutoOriginal,\n        "Produto_Limpo",\n        each fnTextoLimpoMaiusculo([Produto]),\n        type text\n    ),\n\n    ProdutoLimpoMinusculo = Table.AddColumn(\n        ProdutoLimpo,\n        "Produto_Limpo_Min",\n        each fnTextoLimpoMinusculo([Produto]),\n        type text\n    ),\n\n    ClienteLimpo = Table.AddColumn(\n        ProdutoLimpoMinusculo,\n        "Cliente_Limpo",\n        each fnTextoLimpoMaiusculo([Cliente]),\n        type text\n    ),\n\n    SubstituicoesPadrao = {\n        {" LT ", " LATA "},\n        {" LTA ", " LATA "},\n        {" LATAA ", " LATA "},\n        {" CX ", " CAIXA "},\n        {" CERV ", " CERVEJA "},\n        {" LONGNECK ", " LONG NECK "},\n        {" GUARANA ANTARTICA ", " GUARANA ANTARCTICA "},\n        {" BRAHMAA ", " BRAHMA "},\n        {" BRHMA ", " BRAHMA "},\n        {" SKOLL ", " SKOL "}\n    },\n\n    ProdutoPadronizado = Table.AddColumn(\n        ClienteLimpo,\n        "Produto_Padronizado",\n        each\n            let\n                ComEspacos = " " & [Produto_Limpo] & " ",\n                Corrigido = List.Accumulate(\n                    SubstituicoesPadrao,\n                    ComEspacos,\n                    (estado, atual) => Text.Replace(estado, atual{0}, atual{1})\n                ),\n                Final = Text.Trim(Text.Combine(List.Select(Text.Split(Corrigido, " "), each _ <> ""), " "))\n            in\n                Final,\n        type text\n    ),\n\n    MarcaDetectada = Table.AddColumn(\n        ProdutoPadronizado,\n        "Marca_Detectada",\n        each\n            if [Produto_Padronizado] = null then "SEM PRODUTO"\n            else if Text.Contains([Produto_Padronizado], "SKOL") then "SKOL"\n            else if Text.Contains([Produto_Padronizado], "BRAHMA") then "BRAHMA"\n            else if Text.Contains([Produto_Padronizado], "GUARANA") then "GUARANA"\n            else if Text.Contains([Produto_Padronizado], "HEINEKEN") then "HEINEKEN"\n            else "OUTROS",\n        type text\n    ),\n\n    EmbalagemDetectada = Table.AddColumn(\n        MarcaDetectada,\n        "Embalagem_Detectada",\n        each\n            if [Produto_Padronizado] = null then "SEM EMBALAGEM"\n            else if Text.Contains([Produto_Padronizado], "LATA") then "LATA"\n            else if Text.Contains([Produto_Padronizado], "LONG NECK") then "LONG NECK"\n            else if Text.Contains([Produto_Padronizado], "PET") then "PET"\n            else if Text.Contains([Produto_Padronizado], "CAIXA") then "CAIXA"\n            else "OUTROS",\n        type text\n    ),\n\n    VolumeDetectado = Table.AddColumn(\n        EmbalagemDetectada,\n        "Volume_Detectado",\n        each\n            if [Produto_Padronizado] = null then null\n            else if Text.Contains([Produto_Padronizado], "350") then "350ML"\n            else if Text.Contains([Produto_Padronizado], "269") then "269ML"\n            else if Text.Contains([Produto_Padronizado], "600") then "600ML"\n            else if Text.Contains([Produto_Padronizado], "1L") then "1L"\n            else null,\n        type text\n    ),\n\n    DeParaFonte = Excel.CurrentWorkbook(){[Name="DeParaProdutos"]}[Content],\n\n    DeParaTipos = Table.TransformColumnTypes(\n        DeParaFonte,\n        {\n            {"Grafia_Incorreta", type text},\n            {"Produto_Correto", type text}\n        }\n    ),\n\n    DeParaLimpo = Table.TransformColumns(\n        DeParaTipos,\n        {\n            {"Grafia_Incorreta", each fnTextoLimpoMaiusculo(_), type text},\n            {"Produto_Correto", each fnTextoLimpoMaiusculo(_), type text}\n        }\n    ),\n\n    DeParaBuffer = Table.Buffer(DeParaLimpo),\n\n    MergeCorrecoes = Table.NestedJoin(\n        VolumeDetectado,\n        {"Produto_Padronizado"},\n        DeParaBuffer,\n        {"Grafia_Incorreta"},\n        "Correcoes",\n        JoinKind.LeftOuter\n    ),\n\n    Expandido = Table.ExpandTableColumn(\n        MergeCorrecoes,\n        "Correcoes",\n        {"Produto_Correto"},\n        {"Produto_Correto"}\n    ),\n\n    ProdutoFinal = Table.AddColumn(\n        Expandido,\n        "Produto_Final",\n        each if [Produto_Correto] <> null then [Produto_Correto] else [Produto_Padronizado],\n        type text\n    ),\n\n    ValorNumerico = Table.AddColumn(\n        ProdutoFinal,\n        "Valor_Numero",\n        each try Number.From([Valor]) otherwise null,\n        type number\n    ),\n\n    DataConvertida = Table.AddColumn(\n        ValorNumerico,\n        "Data_Convertida",\n        each try Date.From([Data]) otherwise null,\n        type date\n    ),\n\n    StatusCorrecao = Table.AddColumn(\n        DataConvertida,\n        "Status_Correcao",\n        each\n            if [Produto] = null or Text.Trim(Text.From([Produto])) = "" then "ERRO: PRODUTO VAZIO"\n            else if [Produto_Correto] <> null then "CORRIGIDO POR DE/PARA"\n            else if [Produto_Limpo] <> [Produto_Padronizado] then "CORRIGIDO POR REGRA"\n            else "SEM ALTERAÇÃO",\n        type text\n    ),\n\n    StatusDQ = Table.AddColumn(\n        StatusCorrecao,\n        "Status_DQ",\n        each\n            if [Produto_Final] = null or [Produto_Final] = "" then "ERRO PRODUTO"\n            else if [Valor_Numero] = null then "ERRO VALOR"\n            else if [Data_Convertida] = null then "ERRO DATA"\n            else "OK",\n        type text\n    ),\n\n    FaixaValor = Table.AddColumn(\n        StatusDQ,\n        "Faixa_Valor",\n        each\n            if [Valor_Numero] = null then "Sem valor"\n            else if [Valor_Numero] >= 1000 then "Alta"\n            else if [Valor_Numero] >= 500 then "Média"\n            else "Baixa",\n        type text\n    ),\n\n    ChaveAnalitica = Table.AddColumn(\n        FaixaValor,\n        "Chave_Produto_Cliente",\n        each [Produto_Final] & "-" & [Cliente_Limpo],\n        type text\n    ),\n\n    ResultadoFinal = Table.SelectColumns(\n        ChaveAnalitica,\n        {\n            "Produto_Original",\n            "Produto_Limpo",\n            "Produto_Limpo_Min",\n            "Produto_Padronizado",\n            "Produto_Correto",\n            "Produto_Final",\n            "Marca_Detectada",\n            "Embalagem_Detectada",\n            "Volume_Detectado",\n            "Cliente",\n            "Cliente_Limpo",\n            "Valor",\n            "Valor_Numero",\n            "Data",\n            "Data_Convertida",\n            "Faixa_Valor",\n            "Chave_Produto_Cliente",\n            "Status_Correcao",\n            "Status_DQ"\n        },\n        MissingField.Ignore\n    )\nin\n    ResultadoFinal\n'

TOPICS = {
    "Excel Básico": {
        "nivel": "Básico",
        "tag": "EXCEL-BASE",
        "objetivo": "Organizar dados, limpar textos, validar entradas e preparar base tabular.",
        "conceito": "Uma boa análise começa com base tabular: cabeçalhos únicos, sem mesclagem, tipos coerentes e uma linha por registro.",
        "quando": "Controles simples, análises rápidas, bases pequenas e protótipos.",
        "risco": "Automatizar base mal estruturada aumenta erro e retrabalho.",
        "excel": EXCEL_BASICO,
        "pq": PQ_BASICO,
    },
    "Excel Intermediário": {
        "nivel": "Intermediário",
        "tag": "EXCEL-FORMULAS",
        "objetivo": "Combinar fórmulas, criar buscas robustas, aplicar critérios e tratar erros.",
        "conceito": "O intermediário combina funções: SE, E, OU, SEERRO, PROCX, SOMASES, FILTRO, CLASSIFICAR e LET.",
        "quando": "Relatórios recorrentes, conciliações, controles financeiros e análises operacionais.",
        "risco": "Fórmulas longas sem LET ou documentação ficam frágeis.",
        "excel": EXCEL_INTERMEDIARIO,
        "pq": PQ_INTERMEDIARIO,
    },
    "Excel Avançado": {
        "nivel": "Avançado",
        "tag": "EXCEL-ADV",
        "objetivo": "Usar matrizes dinâmicas, LAMBDA, LET, coringas e buscas compostas.",
        "conceito": "O avançado resolve problemas complexos com fórmulas auditáveis e performance consciente.",
        "quando": "Testes técnicos, conciliações complexas e modelos sem VBA.",
        "risco": "Nem toda fórmula avançada é a melhor solução; às vezes Power Query é mais seguro.",
        "excel": EXCEL_AVANCADO,
        "pq": PQ_AVANCADO,
    },
    "Coringas": {
        "nivel": "Intermediário/Avançado",
        "tag": "WILDCARDS",
        "objetivo": "Dominar *, ? e ~ em Excel e equivalentes em Power Query.",
        "conceito": "* significa qualquer sequência, ? significa um caractere e ~ escapa o coringa.",
        "quando": "Produtos, descrições, códigos, SKUs e cadastros textuais.",
        "risco": "Coringas podem gerar falso positivo se a regra for ampla demais.",
        "excel": EXCEL_AVANCADO,
        "pq": PQ_BASICO + PQ_INTERMEDIARIO,
    },
    "Power Query": {
        "nivel": "Básico ao Avançado",
        "tag": "PQ-M",
        "objetivo": "Automatizar limpeza, transformação, junção, agrupamento e auditoria.",
        "conceito": "Power Query usa linguagem M. Ele prepara dados; DAX calcula métricas no modelo.",
        "quando": "Bases recorrentes, múltiplas fontes, arquivos mensais e dados sujos.",
        "risco": "A ordem das etapas e a qualidade da chave determinam a confiabilidade.",
        "excel": EXCEL_INTERMEDIARIO,
        "pq": PQ_BASICO + PQ_INTERMEDIARIO + PQ_AVANCADO,
    },
    "Grafias e De/Para": {
        "nivel": "Avançado",
        "tag": "TEXT-DQ",
        "objetivo": "Corrigir caixa, acentos, abreviações, grafias incorretas e manter rastreabilidade.",
        "conceito": "A arquitetura correta preserva original, cria limpo, aplica regra, aplica De/Para e gera status.",
        "quando": "Produtos, clientes, cidades, fornecedores e descrições de sistemas diferentes.",
        "risco": "Correção manual sem tabela De/Para perde auditoria.",
        "excel": EXCEL_AVANCADO,
        "pq": PQ_AVANCADO,
    },
    "Estatística": {
        "nivel": "Intermediário/Avançado",
        "tag": "STAT-EXCEL",
        "objetivo": "Aplicar média, mediana, moda, dispersão, ranking, percentis, correlação e regressão.",
        "conceito": "Estatística transforma relatório em diagnóstico, comparação, tendência e previsão.",
        "quando": "Vendas, forecast, qualidade, demanda, variação e performance.",
        "risco": "Correlação não implica causalidade; regressão precisa de contexto e validação.",
        "excel": ESTATISTICA,
        "pq": PQ_AVANCADO,
    },
}

np.random.seed(42)
MESES = pd.date_range("2025-01-01", periods=18, freq="MS")
STAT_BASE = pd.DataFrame({
    "Mês": MESES,
    "Investimento": np.linspace(50, 220, 18),
    "Vendas": np.linspace(120, 330, 18) + np.random.normal(0, 18, 18)
})
STAT_BASE["MediaMovel3"] = STAT_BASE["Vendas"].rolling(3).mean()
STAT_BASE["ZScore"] = (STAT_BASE["Vendas"] - STAT_BASE["Vendas"].mean()) / STAT_BASE["Vendas"].std(ddof=1)

def card(title, body, cls="card"):
    st.markdown(f"""
    <div class="{cls}">
        <div class="title-small">{title}</div>
        <div>{body}</div>
    </div>
    """, unsafe_allow_html=True)

def formula_table(rows):
    return pd.DataFrame(rows, columns=["Tema", "Fórmula / Código", "Aplicação"])

def render_formula_cards(rows, language="text"):
    for tema, formula, aplicacao in rows:
        with st.expander(tema, expanded=False):
            st.write(aplicacao)
            st.code(formula, language=language)

def chart_trend():
    fig, ax = plt.subplots(figsize=(9, 4))
    ax.plot(STAT_BASE["Mês"], STAT_BASE["Vendas"], marker="o", label="Vendas")
    ax.plot(STAT_BASE["Mês"], STAT_BASE["MediaMovel3"], marker="o", label="Média móvel 3")
    ax.set_title("Tendência e média móvel")
    ax.set_xlabel("Mês")
    ax.set_ylabel("Vendas")
    ax.legend()
    ax.tick_params(axis="x", rotation=35)
    st.pyplot(fig, clear_figure=True)

def chart_scatter():
    x = STAT_BASE["Investimento"].values
    y = STAT_BASE["Vendas"].values
    coef = np.polyfit(x, y, 1)
    trend = coef[0] * x + coef[1]
    corr = np.corrcoef(x, y)[0, 1]
    r2 = corr ** 2
    fig, ax = plt.subplots(figsize=(9, 4))
    ax.scatter(x, y, label="Observações")
    ax.plot(x, trend, label=f"Regressão | R²={r2:.2f}")
    ax.set_title("Correlação e regressão linear")
    ax.set_xlabel("Investimento")
    ax.set_ylabel("Vendas")
    ax.legend()
    st.pyplot(fig, clear_figure=True)

def chart_hist():
    fig, ax = plt.subplots(figsize=(9, 4))
    ax.hist(STAT_BASE["Vendas"], bins=7)
    ax.set_title("Frequência / Histograma")
    ax.set_xlabel("Faixa de vendas")
    ax.set_ylabel("Frequência")
    st.pyplot(fig, clear_figure=True)

def chart_z():
    fig, ax = plt.subplots(figsize=(9, 4))
    ax.bar(STAT_BASE["Mês"].dt.strftime("%b/%y"), STAT_BASE["ZScore"])
    ax.axhline(0)
    ax.set_title("Padronização — Z-score")
    ax.set_xlabel("Mês")
    ax.set_ylabel("Z-score")
    ax.tick_params(axis="x", rotation=35)
    st.pyplot(fig, clear_figure=True)

with st.sidebar:
    st.markdown("## 📊 Comitê Técnico")
    st.caption("Excel · Power Query · Coringas · Estatística · Data Quality")
    st.divider()
    selected = st.radio("Mapa Mental", list(TOPICS.keys()), index=list(TOPICS.keys()).index("Estatística"))
    st.divider()
    st.metric("Meta", "Nota 9+")
    st.metric("Cobertura", "Básico → Avançado")
    st.info("Critério técnico: não basta calcular; preserve origem, explique regra, valide exceções e entregue leitura executiva.")

active = TOPICS[selected]

st.title("Comitê Técnico — Excel, Power Query e Estatística")
st.caption("Versão v5: biblioteca ampliada, coringas *, ? e ~, Power Query robusto, estatística aplicada e gráficos.")

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
    "Gráficos",
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
    st.dataframe(formula_table(active["excel"]), use_container_width=True, hide_index=True)
    st.subheader("Explicação item a item")
    render_formula_cards(active["excel"], "text")

with tabs[2]:
    st.subheader("Tabela de códigos M")
    st.dataframe(formula_table(active["pq"]), use_container_width=True, hide_index=True)
    st.subheader("Explicação item a item")
    render_formula_cards(active["pq"], "powerquery")

with tabs[3]:
    st.subheader("Coringas do Excel e equivalentes no Power Query")
    df_w = pd.DataFrame(WILDCARDS, columns=["Coringa", "Uso", "Excel", "Power Query"])
    st.dataframe(df_w, use_container_width=True, hide_index=True)
    card("Resumo", "* = qualquer sequência; ? = um caractere; ~ = trata * ou ? como caractere literal.", "card purple")

with tabs[4]:
    st.subheader("Exemplos De/Para")
    df_d = pd.DataFrame(DEPARA_EXAMPLES, columns=["Grafia Incorreta", "Produto Correto", "Motivo"])
    st.dataframe(df_d, use_container_width=True, hide_index=True)
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
    st.dataframe(formula_table(ESTATISTICA), use_container_width=True, hide_index=True)
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
    chart = st.selectbox("Escolha o gráfico", ["Tendência", "Correlação / Regressão", "Histograma", "Z-score"])
    if chart == "Tendência":
        chart_trend()
    elif chart == "Correlação / Regressão":
        chart_scatter()
    elif chart == "Histograma":
        chart_hist()
    else:
        chart_z()

with tabs[7]:
    st.subheader("Data Quality")
    dq_excel = [
        ("Vazios", '=CONTAR.VAZIO(A:A)', "Conta vazios."),
        ("Duplicados", '=SE(CONT.SE(A:A;A2)>1;"Duplicado";"Único")', "Sinaliza duplicidade."),
        ("Número inválido", '=SEERRO(VALOR(A2);"Erro numérico")', "Valida número."),
        ("Status", '=SE(E(A2<>"";B2>0);"OK";"Validar")', "Cria status."),
    ]
    dq_pq = [
        ("Profile", "Table.Profile(Fonte)", "Perfil estatístico da tabela."),
        ("Schema", "Table.Schema(Fonte)", "Estrutura da tabela."),
        ("Exceções", 'Table.SelectRows(Fonte, each [Status_DQ] <> "OK")', "Filtra problemas."),
        ("Duplicados", 'Table.Group(Fonte, Chave, {"Qtd", each Table.RowCount(_), Int64.Type})', "Conta por chave."),
    ]
    st.dataframe(formula_table(dq_excel + dq_pq), use_container_width=True, hide_index=True)

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
        file_name="powerquery_grafias_dataquality_v5.m",
        mime="text/plain"
    )

st.divider()
st.caption("v5 auditada: sem __file__, com bibliotecas ampliadas, coringas completos, estatística e gráficos.")


# Glossário técnico 001: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 002: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 003: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 004: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 005: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 006: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 007: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 008: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 009: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 010: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 011: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 012: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 013: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 014: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 015: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 016: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 017: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 018: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 019: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 020: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 021: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 022: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 023: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 024: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 025: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 026: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 027: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 028: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 029: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 030: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 031: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 032: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 033: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 034: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 035: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 036: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 037: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 038: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 039: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 040: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 041: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 042: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 043: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 044: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 045: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 046: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 047: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 048: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 049: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 050: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 051: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 052: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 053: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 054: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 055: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 056: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 057: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 058: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 059: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 060: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 061: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 062: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 063: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 064: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 065: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 066: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 067: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 068: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 069: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 070: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 071: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 072: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 073: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 074: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 075: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 076: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 077: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 078: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 079: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 080: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 081: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 082: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 083: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 084: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 085: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 086: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 087: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 088: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 089: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 090: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 091: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 092: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 093: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 094: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 095: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 096: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 097: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 098: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 099: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 100: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 101: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 102: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 103: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 104: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 105: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 106: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 107: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 108: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 109: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 110: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 111: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 112: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 113: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 114: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 115: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 116: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 117: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 118: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 119: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 120: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 121: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 122: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 123: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 124: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 125: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 126: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 127: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 128: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 129: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 130: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 131: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 132: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 133: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 134: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 135: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 136: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 137: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 138: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 139: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 140: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 141: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 142: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 143: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 144: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 145: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 146: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 147: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 148: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 149: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 150: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 151: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 152: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 153: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 154: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 155: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 156: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 157: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 158: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 159: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 160: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 161: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 162: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 163: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 164: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 165: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 166: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 167: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 168: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 169: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 170: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 171: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 172: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 173: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 174: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 175: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 176: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 177: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 178: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 179: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 180: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 181: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 182: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 183: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 184: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 185: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 186: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 187: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 188: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 189: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 190: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 191: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 192: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 193: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 194: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 195: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 196: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 197: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 198: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 199: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 200: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 201: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 202: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 203: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 204: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 205: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 206: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 207: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 208: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 209: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 210: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 211: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 212: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 213: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 214: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 215: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 216: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 217: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 218: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 219: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 220: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 221: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 222: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 223: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 224: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 225: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 226: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 227: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 228: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 229: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 230: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 231: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 232: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 233: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 234: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 235: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 236: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 237: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 238: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 239: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 240: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 241: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 242: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 243: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 244: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 245: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 246: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 247: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 248: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 249: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 250: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 251: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 252: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 253: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 254: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 255: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 256: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 257: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 258: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 259: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 260: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 261: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 262: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 263: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 264: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 265: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 266: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 267: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 268: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 269: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 270: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 271: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 272: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 273: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 274: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 275: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 276: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 277: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 278: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 279: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 280: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 281: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 282: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 283: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 284: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 285: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 286: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 287: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 288: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 289: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 290: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 291: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 292: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 293: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 294: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 295: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 296: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 297: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 298: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 299: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 300: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 301: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 302: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 303: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 304: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 305: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 306: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 307: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 308: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 309: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 310: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 311: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 312: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 313: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 314: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 315: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 316: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 317: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 318: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 319: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 320: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 321: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 322: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 323: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 324: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 325: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 326: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 327: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 328: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 329: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 330: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 331: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 332: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 333: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 334: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 335: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 336: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 337: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 338: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 339: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 340: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 341: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 342: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 343: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 344: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 345: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 346: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 347: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 348: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 349: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 350: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 351: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 352: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 353: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 354: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 355: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 356: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 357: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 358: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 359: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 360: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 361: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 362: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 363: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 364: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 365: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 366: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 367: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 368: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 369: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 370: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 371: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 372: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 373: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 374: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 375: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 376: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 377: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 378: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 379: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 380: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 381: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 382: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 383: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 384: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 385: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 386: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 387: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 388: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 389: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 390: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 391: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 392: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 393: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 394: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 395: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 396: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 397: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 398: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 399: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 400: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 401: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 402: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 403: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 404: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 405: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 406: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 407: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 408: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 409: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 410: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 411: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 412: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 413: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 414: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 415: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 416: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 417: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 418: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 419: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 420: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 421: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 422: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 423: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 424: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 425: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 426: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 427: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 428: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 429: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 430: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 431: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 432: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 433: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 434: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 435: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 436: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 437: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 438: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 439: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 440: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 441: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 442: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 443: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 444: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 445: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 446: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 447: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 448: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 449: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 450: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 451: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 452: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 453: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 454: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 455: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 456: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 457: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 458: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 459: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 460: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 461: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 462: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 463: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 464: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 465: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 466: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 467: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 468: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 469: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 470: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 471: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 472: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 473: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 474: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 475: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 476: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 477: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 478: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 479: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 480: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 481: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 482: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 483: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 484: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 485: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 486: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 487: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 488: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 489: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 490: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 491: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 492: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 493: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 494: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 495: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 496: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 497: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 498: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 499: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 500: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 501: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 502: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 503: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 504: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 505: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 506: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 507: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 508: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 509: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 510: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 511: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 512: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 513: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 514: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 515: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 516: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 517: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 518: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 519: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 520: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 521: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 522: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 523: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 524: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 525: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 526: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 527: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 528: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 529: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 530: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 531: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 532: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 533: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 534: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 535: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 536: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 537: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 538: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 539: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 540: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 541: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 542: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 543: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 544: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 545: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 546: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 547: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 548: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 549: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 550: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 551: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 552: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 553: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 554: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 555: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 556: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 557: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 558: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 559: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 560: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 561: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 562: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 563: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 564: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 565: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 566: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 567: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 568: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 569: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 570: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 571: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 572: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 573: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 574: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 575: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 576: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 577: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 578: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 579: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 580: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 581: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 582: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 583: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 584: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 585: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 586: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 587: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 588: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 589: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 590: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 591: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 592: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 593: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 594: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 595: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 596: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 597: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 598: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 599: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 600: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 601: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 602: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 603: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 604: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 605: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 606: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 607: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 608: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 609: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 610: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 611: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 612: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 613: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 614: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 615: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 616: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 617: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 618: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 619: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 620: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 621: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 622: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 623: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 624: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 625: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 626: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 627: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 628: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 629: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 630: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 631: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 632: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 633: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 634: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 635: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 636: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 637: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 638: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 639: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 640: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 641: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 642: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 643: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 644: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 645: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 646: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 647: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 648: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 649: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 650: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 651: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 652: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 653: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 654: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 655: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 656: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 657: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 658: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 659: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 660: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 661: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 662: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 663: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 664: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 665: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 666: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 667: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 668: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 669: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 670: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 671: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 672: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 673: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 674: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 675: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 676: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 677: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 678: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 679: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 680: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 681: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 682: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 683: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 684: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 685: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 686: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 687: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 688: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 689: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 690: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 691: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 692: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 693: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 694: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 695: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 696: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 697: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 698: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 699: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 700: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 701: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 702: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 703: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 704: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 705: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 706: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 707: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 708: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 709: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 710: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 711: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 712: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 713: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 714: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 715: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 716: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 717: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 718: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 719: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 720: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 721: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 722: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 723: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 724: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 725: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 726: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 727: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 728: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 729: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 730: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 731: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 732: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 733: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 734: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 735: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 736: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 737: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 738: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 739: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 740: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 741: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 742: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 743: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 744: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 745: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 746: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 747: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 748: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 749: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 750: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 751: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 752: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 753: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 754: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 755: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 756: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 757: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 758: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 759: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 760: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 761: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 762: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 763: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 764: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 765: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 766: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 767: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 768: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 769: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 770: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 771: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 772: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 773: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 774: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 775: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 776: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 777: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 778: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 779: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 780: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 781: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 782: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 783: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 784: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 785: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 786: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 787: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 788: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 789: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 790: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 791: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 792: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 793: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 794: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 795: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 796: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 797: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 798: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 799: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 800: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 801: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 802: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 803: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 804: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 805: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 806: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 807: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 808: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 809: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 810: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 811: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 812: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 813: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 814: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 815: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 816: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 817: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 818: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 819: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 820: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 821: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 822: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 823: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 824: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 825: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 826: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 827: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 828: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 829: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 830: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 831: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 832: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 833: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 834: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 835: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 836: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 837: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 838: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 839: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 840: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 841: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 842: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 843: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 844: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 845: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 846: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 847: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 848: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 849: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 850: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 851: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 852: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 853: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 854: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 855: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 856: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 857: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 858: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 859: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 860: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 861: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 862: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 863: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 864: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 865: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 866: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 867: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 868: Power Pivot — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 869: coringa asterisco — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 870: coringa interrogação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 871: til de escape — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 872: De Para — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 873: Data Quality — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 874: fuzzy matching — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 875: List.Accumulate — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 876: Table.Group — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 877: Table.NestedJoin — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 878: Table.Profile — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 879: média — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 880: mediana — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 881: moda — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 882: variância — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 883: desvio padrão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 884: z-score — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 885: percentil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 886: quartil — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 887: correlação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 888: regressão — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 889: R quadrado — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 890: inclinação — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 891: intercepção — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 892: frequência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 893: tendência — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 894: PROJ.LIN — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 895: PROJ.LOG — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 896: base tabular — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 897: tabela estruturada — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 898: Power Query M — revisar conceito, sintaxe, risco e aplicação em teste técnico.
# Glossário técnico 899: DAX — revisar conceito, sintaxe, risco e aplicação em teste técnico.
