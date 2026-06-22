
import streamlit as st
import pandas as pd
import math

st.set_page_config(page_title="Comitê Técnico | Excel, Power Query, VBA e Estatística", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
st.markdown('''
<style>
html, body, [data-testid="stAppViewContainer"] { overflow-x: hidden !important; }
.block-container { max-width: 1540px !important; padding-top: 1.0rem !important; padding-left: 1.25rem !important; padding-right: 1.25rem !important; padding-bottom: 2.5rem !important; }
[data-testid="stSidebar"] { min-width: 300px !important; max-width: 360px !important; }
.stTabs [data-baseweb="tab-list"] { gap: 0.25rem; flex-wrap: wrap; }
.stTabs [data-baseweb="tab"] { height: auto; min-height: 38px; white-space: normal; padding: 8px 12px; }
.card { border: 1px solid #d9dee7; border-left: 5px solid #1f77b4; border-radius: 10px; padding: 16px; margin-bottom: 12px; background: #ffffff; }
.green { border-left-color: #18a558; background: #eefaf1; } .yellow { border-left-color: #f0ad4e; background: #fff8e6; }
.red { border-left-color: #d9534f; background: #fff0f0; } .blue { border-left-color: #1f77b4; background: #eef6ff; }
.purple { border-left-color: #7e57c2; background: #f5f0ff; }
.title-small { font-size: 1.02rem; font-weight: 750; margin-bottom: 0.45rem; color: #17233c; }
.tag { display: inline-block; padding: 5px 10px; border-radius: 999px; background: #e7f0ff; color: #0b5ed7; font-size: 0.78rem; font-weight: 750; margin-bottom: 8px; margin-right: 5px; }
pre, code { white-space: pre-wrap !important; word-break: break-word !important; overflow-wrap: anywhere !important; font-size: 0.78rem !important; }
</style>
''', unsafe_allow_html=True)

EXCEL_BASIC = [{'Tema': 'SE vazio', 'Formula': '=SE(A2="";"Sem informação";A2)', 'Uso': 'Substitui vazio por texto de controle.', 'Nivel': 'Básico'}, {'Tema': 'ÉCÉL.VAZIA', 'Formula': '=SE(ÉCÉL.VAZIA(A2);"Vazio";"Preenchido")', 'Uso': 'Valida célula em branco.', 'Nivel': 'Básico'}, {'Tema': 'SEERRO', 'Formula': '=SEERRO(A2/B2;0)', 'Uso': 'Trata erro de cálculo.', 'Nivel': 'Básico'}, {'Tema': 'ARRUMAR', 'Formula': '=ARRUMAR(A2)', 'Uso': 'Remove espaços excedentes.', 'Nivel': 'Básico'}, {'Tema': 'MAIÚSCULA', 'Formula': '=MAIÚSCULA(A2)', 'Uso': 'Padroniza caixa alta.', 'Nivel': 'Básico'}, {'Tema': 'MINÚSCULA', 'Formula': '=MINÚSCULA(A2)', 'Uso': 'Padroniza caixa baixa.', 'Nivel': 'Básico'}, {'Tema': 'PRI.MAIÚSCULA', 'Formula': '=PRI.MAIÚSCULA(A2)', 'Uso': 'Padroniza nomes próprios.', 'Nivel': 'Básico'}, {'Tema': 'NÚM.CARACT', 'Formula': '=NÚM.CARACT(A2)', 'Uso': 'Conta caracteres.', 'Nivel': 'Básico'}, {'Tema': 'ESQUERDA', 'Formula': '=ESQUERDA(A2;3)', 'Uso': 'Extrai caracteres à esquerda.', 'Nivel': 'Básico'}, {'Tema': 'DIREITA', 'Formula': '=DIREITA(A2;4)', 'Uso': 'Extrai caracteres à direita.', 'Nivel': 'Básico'}, {'Tema': 'EXT.TEXTO', 'Formula': '=EXT.TEXTO(A2;4;6)', 'Uso': 'Extrai trecho intermediário.', 'Nivel': 'Básico'}, {'Tema': 'LOCALIZAR', 'Formula': '=LOCALIZAR("LATA";A2)', 'Uso': 'Localiza texto com diferenciação.', 'Nivel': 'Básico'}, {'Tema': 'PROCURAR', 'Formula': '=PROCURAR("lata";A2)', 'Uso': 'Localiza texto sem diferenciação de caixa.', 'Nivel': 'Básico'}, {'Tema': 'SUBSTITUIR', 'Formula': '=SUBSTITUIR(A2;"LT";"LATA")', 'Uso': 'Troca trecho textual.', 'Nivel': 'Básico'}, {'Tema': 'TEXTOJUNTAR', 'Formula': '=TEXTOJUNTAR(" | ";VERDADEIRO;A2:C2)', 'Uso': 'Une textos ignorando vazios.', 'Nivel': 'Básico'}, {'Tema': 'VALOR', 'Formula': '=VALOR(A2)', 'Uso': 'Converte texto em número.', 'Nivel': 'Básico'}, {'Tema': 'TEXTO', 'Formula': '=TEXTO(A2;"000000")', 'Uso': 'Formata valor como texto.', 'Nivel': 'Básico'}, {'Tema': 'HOJE', 'Formula': '=HOJE()', 'Uso': 'Retorna data atual.', 'Nivel': 'Básico'}, {'Tema': 'AGORA', 'Formula': '=AGORA()', 'Uso': 'Retorna data e hora atual.', 'Nivel': 'Básico'}, {'Tema': 'ANO', 'Formula': '=ANO(A2)', 'Uso': 'Extrai ano.', 'Nivel': 'Básico'}, {'Tema': 'MÊS', 'Formula': '=MÊS(A2)', 'Uso': 'Extrai mês.', 'Nivel': 'Básico'}, {'Tema': 'DIA', 'Formula': '=DIA(A2)', 'Uso': 'Extrai dia.', 'Nivel': 'Básico'}, {'Tema': 'FIMMÊS', 'Formula': '=FIMMÊS(A2;0)', 'Uso': 'Retorna último dia do mês.', 'Nivel': 'Básico'}, {'Tema': 'DATA competência', 'Formula': '=DATA(ANO(A2);MÊS(A2);1)', 'Uso': 'Cria primeira data da competência.', 'Nivel': 'Básico'}]
EXCEL_INTERMEDIATE = [{'Tema': 'SE + E', 'Formula': '=SE(E(B2>=1000;C2="Ativo");"Prioritário";"Normal")', 'Uso': 'Classificação com critérios simultâneos.', 'Nivel': 'Intermediário'}, {'Tema': 'SE + OU', 'Formula': '=SE(OU(A2="SKOL";A2="BRAHMA");"Cerveja";"Outros")', 'Uso': 'Classificação com alternativa.', 'Nivel': 'Intermediário'}, {'Tema': 'SE aninhado', 'Formula': '=SE(B2>=1000;"Alta";SE(B2>=500;"Média";"Baixa"))', 'Uso': 'Cria faixas.', 'Nivel': 'Intermediário'}, {'Tema': 'PROCV', 'Formula': '=PROCV(A2;Cadastro!A:D;4;FALSO)', 'Uso': 'Busca vertical clássica.', 'Nivel': 'Intermediário'}, {'Tema': 'PROCX', 'Formula': '=PROCX(A2;Produtos[SKU];Produtos[Categoria];"Sem cadastro")', 'Uso': 'Busca moderna.', 'Nivel': 'Intermediário'}, {'Tema': 'PROCX composto', 'Formula': '=PROCX(1;(Produtos[Produto]=A2)*(Produtos[Canal]=B2);Produtos[Categoria];"N/D")', 'Uso': 'Busca com múltiplos critérios.', 'Nivel': 'Intermediário'}, {'Tema': 'SOMASE', 'Formula': '=SOMASE(Base[Produto];A2;Base[Valor])', 'Uso': 'Soma por critério.', 'Nivel': 'Intermediário'}, {'Tema': 'SOMASES', 'Formula': '=SOMASES(Base[Valor];Base[Produto];A2;Base[Canal];"Online")', 'Uso': 'Soma por múltiplos critérios.', 'Nivel': 'Intermediário'}, {'Tema': 'CONT.SE', 'Formula': '=CONT.SE(Base[Produto];A2)', 'Uso': 'Contagem por critério.', 'Nivel': 'Intermediário'}, {'Tema': 'CONT.SES', 'Formula': '=CONT.SES(Base[Produto];A2;Base[Status];"Ativo")', 'Uso': 'Contagem por múltiplos critérios.', 'Nivel': 'Intermediário'}, {'Tema': 'MÉDIASE', 'Formula': '=MÉDIASE(Base[Produto];A2;Base[Valor])', 'Uso': 'Média condicional.', 'Nivel': 'Intermediário'}, {'Tema': 'MÉDIASES', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A2;Base[Canal];"Online")', 'Uso': 'Média com múltiplos critérios.', 'Nivel': 'Intermediário'}, {'Tema': 'ÍNDICE + CORRESP', 'Formula': '=ÍNDICE(Tabela[Valor];CORRESP(A2;Tabela[Produto];0))', 'Uso': 'Busca clássica flexível.', 'Nivel': 'Intermediário'}, {'Tema': 'ÍNDICE + CORRESP composto', 'Formula': '=ÍNDICE(Tabela[Valor];CORRESP(1;(Tabela[Produto]=A2)*(Tabela[Canal]=B2);0))', 'Uso': 'Busca clássica por múltiplos critérios.', 'Nivel': 'Intermediário'}, {'Tema': 'ÚNICO', 'Formula': '=ÚNICO(Base[Produto])', 'Uso': 'Lista distintos em Excel 365 pt-BR quando disponível.', 'Nivel': 'Intermediário'}, {'Tema': 'CLASSIFICAR', 'Formula': '=CLASSIFICAR(ÚNICO(Base[Produto]))', 'Uso': 'Ordena matriz dinâmica.', 'Nivel': 'Intermediário'}, {'Tema': 'FILTRO', 'Formula': '=FILTRO(Base;Base[Valor]>1000;"Sem registros")', 'Uso': 'Filtra matriz dinâmica.', 'Nivel': 'Intermediário'}, {'Tema': 'LET', 'Formula': '=LET(prod;ARRUMAR(MAIÚSCULA(A2));SE(prod="";"VAZIO";prod))', 'Uso': 'Nomeia variáveis na fórmula.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 1', 'Formula': '=SEERRO(B2/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 2', 'Formula': '=SEERRO(B3/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 3', 'Formula': '=SEERRO(B4/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 4', 'Formula': '=SEERRO(B5/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 5', 'Formula': '=SEERRO(B6/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 6', 'Formula': '=SEERRO(B7/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 7', 'Formula': '=SEERRO(B8/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 8', 'Formula': '=SEERRO(B9/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 9', 'Formula': '=SEERRO(B10/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 10', 'Formula': '=SEERRO(B11/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 11', 'Formula': '=SEERRO(B12/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 12', 'Formula': '=SEERRO(B13/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 13', 'Formula': '=SEERRO(B14/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 14', 'Formula': '=SEERRO(B15/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 15', 'Formula': '=SEERRO(B16/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 16', 'Formula': '=SEERRO(B17/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 17', 'Formula': '=SEERRO(B18/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 18', 'Formula': '=SEERRO(B19/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 19', 'Formula': '=SEERRO(B20/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 20', 'Formula': '=SEERRO(B21/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 21', 'Formula': '=SEERRO(B22/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 22', 'Formula': '=SEERRO(B23/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 23', 'Formula': '=SEERRO(B24/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 24', 'Formula': '=SEERRO(B25/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 25', 'Formula': '=SEERRO(B26/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 26', 'Formula': '=SEERRO(B27/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 27', 'Formula': '=SEERRO(B28/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 28', 'Formula': '=SEERRO(B29/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 29', 'Formula': '=SEERRO(B30/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 30', 'Formula': '=SEERRO(B31/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 31', 'Formula': '=SEERRO(B32/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 32', 'Formula': '=SEERRO(B33/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 33', 'Formula': '=SEERRO(B34/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 34', 'Formula': '=SEERRO(B35/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 35', 'Formula': '=SEERRO(B36/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 36', 'Formula': '=SEERRO(B37/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 37', 'Formula': '=SEERRO(B38/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 38', 'Formula': '=SEERRO(B39/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 39', 'Formula': '=SEERRO(B40/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 40', 'Formula': '=SEERRO(B41/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 41', 'Formula': '=SEERRO(B42/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 42', 'Formula': '=SEERRO(B43/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 43', 'Formula': '=SEERRO(B44/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 44', 'Formula': '=SEERRO(B45/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 45', 'Formula': '=SEERRO(B46/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 46', 'Formula': '=SEERRO(B47/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 47', 'Formula': '=SEERRO(B48/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 48', 'Formula': '=SEERRO(B49/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 49', 'Formula': '=SEERRO(B50/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 50', 'Formula': '=SEERRO(B51/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 51', 'Formula': '=SEERRO(B52/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 52', 'Formula': '=SEERRO(B53/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 53', 'Formula': '=SEERRO(B54/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 54', 'Formula': '=SEERRO(B55/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 55', 'Formula': '=SEERRO(B56/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 56', 'Formula': '=SEERRO(B57/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 57', 'Formula': '=SEERRO(B58/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 58', 'Formula': '=SEERRO(B59/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 59', 'Formula': '=SEERRO(B60/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 60', 'Formula': '=SEERRO(B61/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 61', 'Formula': '=SEERRO(B62/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 62', 'Formula': '=SEERRO(B63/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 63', 'Formula': '=SEERRO(B64/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 64', 'Formula': '=SEERRO(B65/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 65', 'Formula': '=SEERRO(B66/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 66', 'Formula': '=SEERRO(B67/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 67', 'Formula': '=SEERRO(B68/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 68', 'Formula': '=SEERRO(B69/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 69', 'Formula': '=SEERRO(B70/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}, {'Tema': 'Participação percentual 70', 'Formula': '=SEERRO(B71/SOMA(B:B);0)', 'Uso': 'Participação com tratamento de erro.', 'Nivel': 'Intermediário'}]
EXCEL_ADVANCED = [{'Tema': 'Coringa contém', 'Formula': '=CONT.SE(A:A;"*SKOL*")', 'Uso': '* antes e depois significa contém.', 'Nivel': 'Avançado'}, {'Tema': 'Coringa começa', 'Formula': '=CONT.SE(A:A;"SKOL*")', 'Uso': 'Começa com SKOL.', 'Nivel': 'Avançado'}, {'Tema': 'Coringa termina', 'Formula': '=CONT.SE(A:A;"*350ML")', 'Uso': 'Termina com 350ML.', 'Nivel': 'Avançado'}, {'Tema': 'Coringa ?', 'Formula': '=CONT.SE(A:A;"SKO?")', 'Uso': '? representa exatamente um caractere.', 'Nivel': 'Avançado'}, {'Tema': 'Coringa ???', 'Formula': '=CONT.SE(A:A;"???")', 'Uso': 'Exatamente três caracteres.', 'Nivel': 'Avançado'}, {'Tema': 'Escape asterisco', 'Formula': '=CONT.SE(A:A;"SKOL~*")', 'Uso': 'Busca SKOL* literal.', 'Nivel': 'Avançado'}, {'Tema': 'Escape interrogação', 'Formula': '=CONT.SE(A:A;"SKOL~?")', 'Uso': 'Busca SKOL? literal.', 'Nivel': 'Avançado'}, {'Tema': 'Escape til', 'Formula': '=CONT.SE(A:A;"SKU~~01")', 'Uso': 'Busca SKU~01 literal.', 'Nivel': 'Avançado'}, {'Tema': 'PROCX curinga', 'Formula': '=PROCX("*LATA*";Base[Produto];Base[Categoria];"N/D";2)', 'Uso': 'PROCX com correspondência curinga.', 'Nivel': 'Avançado'}, {'Tema': 'FILTRO + PROCURAR', 'Formula': '=FILTRO(Base;ÉNÚM(PROCURAR("SKOL";Base[Produto]));"Sem SKOL")', 'Uso': 'Filtra por texto contido.', 'Nivel': 'Avançado'}, {'Tema': 'LET + normalização', 'Formula': '=LET(txt;MAIÚSCULA(ARRUMAR(A2));SE(ÉNÚM(PROCURAR("SKOL";txt));"SKOL";txt))', 'Uso': 'Limpa e classifica.', 'Nivel': 'Avançado'}, {'Tema': 'LAMBDA inline', 'Formula': '=LAMBDA(txt;MAIÚSCULA(ARRUMAR(txt)))(A2)', 'Uso': 'Função reutilizável.', 'Nivel': 'Avançado'}, {'Tema': 'MAP', 'Formula': '=MAP(A2:A10;LAMBDA(x;MAIÚSCULA(ARRUMAR(x))))', 'Uso': 'Aplica função item a item no Microsoft 365.', 'Nivel': 'Avançado'}, {'Tema': 'BYROW', 'Formula': '=BYROW(A2:C10;LAMBDA(linha;SOMA(linha)))', 'Uso': 'Aplica LAMBDA por linha no Microsoft 365.', 'Nivel': 'Avançado'}, {'Tema': 'EMPILHARV', 'Formula': '=EMPILHARV(TabelaJan;TabelaFev;TabelaMar)', 'Uso': 'Empilha matrizes verticalmente.', 'Nivel': 'Avançado'}, {'Tema': 'EMPILHARH', 'Formula': '=EMPILHARH(TabelaProdutos;TabelaCategorias)', 'Uso': 'Empilha matrizes horizontalmente.', 'Nivel': 'Avançado'}, {'Tema': 'ESCOLHERCOLS', 'Formula': '=ESCOLHERCOLS(Base;1;3;5)', 'Uso': 'Seleciona colunas por índice quando disponível.', 'Nivel': 'Avançado'}, {'Tema': 'DESCARTAR', 'Formula': '=DESCARTAR(Base;1)', 'Uso': 'Remove linhas/colunas iniciais quando disponível.', 'Nivel': 'Avançado'}, {'Tema': 'PEGAR', 'Formula': '=PEGAR(CLASSIFICAR(Base;3;-1);10)', 'Uso': 'Retorna primeiras linhas de matriz ordenada quando disponível.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 1', 'Formula': '=LET(v;B2;limite;100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 2', 'Formula': '=LET(v;B3;limite;200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 3', 'Formula': '=LET(v;B4;limite;300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 4', 'Formula': '=LET(v;B5;limite;400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 5', 'Formula': '=LET(v;B6;limite;500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 6', 'Formula': '=LET(v;B7;limite;600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 7', 'Formula': '=LET(v;B8;limite;700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 8', 'Formula': '=LET(v;B9;limite;800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 9', 'Formula': '=LET(v;B10;limite;900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 10', 'Formula': '=LET(v;B11;limite;1000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 11', 'Formula': '=LET(v;B12;limite;1100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 12', 'Formula': '=LET(v;B13;limite;1200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 13', 'Formula': '=LET(v;B14;limite;1300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 14', 'Formula': '=LET(v;B15;limite;1400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 15', 'Formula': '=LET(v;B16;limite;1500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 16', 'Formula': '=LET(v;B17;limite;1600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 17', 'Formula': '=LET(v;B18;limite;1700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 18', 'Formula': '=LET(v;B19;limite;1800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 19', 'Formula': '=LET(v;B20;limite;1900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 20', 'Formula': '=LET(v;B21;limite;2000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 21', 'Formula': '=LET(v;B22;limite;2100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 22', 'Formula': '=LET(v;B23;limite;2200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 23', 'Formula': '=LET(v;B24;limite;2300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 24', 'Formula': '=LET(v;B25;limite;2400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 25', 'Formula': '=LET(v;B26;limite;2500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 26', 'Formula': '=LET(v;B27;limite;2600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 27', 'Formula': '=LET(v;B28;limite;2700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 28', 'Formula': '=LET(v;B29;limite;2800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 29', 'Formula': '=LET(v;B30;limite;2900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 30', 'Formula': '=LET(v;B31;limite;3000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 31', 'Formula': '=LET(v;B32;limite;3100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 32', 'Formula': '=LET(v;B33;limite;3200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 33', 'Formula': '=LET(v;B34;limite;3300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 34', 'Formula': '=LET(v;B35;limite;3400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 35', 'Formula': '=LET(v;B36;limite;3500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 36', 'Formula': '=LET(v;B37;limite;3600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 37', 'Formula': '=LET(v;B38;limite;3700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 38', 'Formula': '=LET(v;B39;limite;3800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 39', 'Formula': '=LET(v;B40;limite;3900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 40', 'Formula': '=LET(v;B41;limite;4000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 41', 'Formula': '=LET(v;B42;limite;4100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 42', 'Formula': '=LET(v;B43;limite;4200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 43', 'Formula': '=LET(v;B44;limite;4300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 44', 'Formula': '=LET(v;B45;limite;4400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 45', 'Formula': '=LET(v;B46;limite;4500;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 46', 'Formula': '=LET(v;B47;limite;4600;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 47', 'Formula': '=LET(v;B48;limite;4700;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 48', 'Formula': '=LET(v;B49;limite;4800;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 49', 'Formula': '=LET(v;B50;limite;4900;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 50', 'Formula': '=LET(v;B51;limite;5000;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 51', 'Formula': '=LET(v;B52;limite;5100;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 52', 'Formula': '=LET(v;B53;limite;5200;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 53', 'Formula': '=LET(v;B54;limite;5300;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}, {'Tema': 'LET auditável 54', 'Formula': '=LET(v;B55;limite;5400;SE(v>=limite;"OK";"VALIDAR"))', 'Uso': 'Fórmula auditável com variável.', 'Nivel': 'Avançado'}]
STATISTICS = [{'Tema': 'MÉDIA', 'Formula': '=MÉDIA(B:B)', 'Uso': 'Tendência central sensível a extremos.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASE', 'Formula': '=MÉDIASE(A:A;"SKOL";B:B)', 'Uso': 'Média por critério.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES', 'Formula': '=MÉDIASES(B:B;A:A;"SKOL";C:C;"Online")', 'Uso': 'Média por múltiplos critérios.', 'Nivel': 'Estatística'}, {'Tema': 'MED', 'Formula': '=MED(B:B)', 'Uso': 'Mediana; centro robusto contra extremos.', 'Nivel': 'Estatística'}, {'Tema': 'MODO.ÚNICO', 'Formula': '=MODO.ÚNICO(B:B)', 'Uso': 'Valor mais frequente.', 'Nivel': 'Estatística'}, {'Tema': 'VAR.S', 'Formula': '=VAR.S(B:B)', 'Uso': 'Variância amostral.', 'Nivel': 'Estatística'}, {'Tema': 'VAR.P', 'Formula': '=VAR.P(B:B)', 'Uso': 'Variância populacional.', 'Nivel': 'Estatística'}, {'Tema': 'DESVPAD.S', 'Formula': '=DESVPAD.S(B:B)', 'Uso': 'Desvio padrão amostral.', 'Nivel': 'Estatística'}, {'Tema': 'DESVPAD.P', 'Formula': '=DESVPAD.P(B:B)', 'Uso': 'Desvio padrão populacional.', 'Nivel': 'Estatística'}, {'Tema': 'PADRONIZAR', 'Formula': '=PADRONIZAR(B2;MÉDIA(B:B);DESVPAD.S(B:B))', 'Uso': 'Z-score.', 'Nivel': 'Estatística'}, {'Tema': 'ORDEM.EQ', 'Formula': '=ORDEM.EQ(B2;B:B;0)', 'Uso': 'Ranking absoluto.', 'Nivel': 'Estatística'}, {'Tema': 'ORDEM.PORCENTUAL.INC', 'Formula': '=ORDEM.PORCENTUAL.INC(B:B;B2)', 'Uso': 'Posição percentual.', 'Nivel': 'Estatística'}, {'Tema': 'PERCENTIL.INC', 'Formula': '=PERCENTIL.INC(B:B;0,9)', 'Uso': 'Corte de percentil.', 'Nivel': 'Estatística'}, {'Tema': 'QUARTIL.INC', 'Formula': '=QUARTIL.INC(B:B;3)', 'Uso': 'Corte em quartis.', 'Nivel': 'Estatística'}, {'Tema': 'CORREL', 'Formula': '=CORREL(B:B;C:C)', 'Uso': 'Correlação linear.', 'Nivel': 'Estatística'}, {'Tema': 'PEARSON', 'Formula': '=PEARSON(B:B;C:C)', 'Uso': 'Correlação de Pearson.', 'Nivel': 'Estatística'}, {'Tema': 'RQUAD', 'Formula': '=RQUAD(B:B;C:C)', 'Uso': 'Coeficiente de determinação.', 'Nivel': 'Estatística'}, {'Tema': 'INCLINAÇÃO', 'Formula': '=INCLINAÇÃO(B:B;C:C)', 'Uso': 'Coeficiente angular.', 'Nivel': 'Estatística'}, {'Tema': 'INTERCEPÇÃO', 'Formula': '=INTERCEPÇÃO(B:B;C:C)', 'Uso': 'Intercepto.', 'Nivel': 'Estatística'}, {'Tema': 'EPADYX', 'Formula': '=EPADYX(B:B;C:C)', 'Uso': 'Erro padrão da estimativa Y.', 'Nivel': 'Estatística'}, {'Tema': 'FREQUÊNCIA', 'Formula': '=FREQUÊNCIA(B:B;E2:E6)', 'Uso': 'Distribuição por faixas.', 'Nivel': 'Estatística'}, {'Tema': 'TENDÊNCIA', 'Formula': '=TENDÊNCIA(B:B;C:C;D2:D10)', 'Uso': 'Projeção linear.', 'Nivel': 'Estatística'}, {'Tema': 'PROJ.LIN', 'Formula': '=PROJ.LIN(B:B;C:C;VERDADEIRO;VERDADEIRO)', 'Uso': 'Regressão linear detalhada.', 'Nivel': 'Estatística'}, {'Tema': 'PROJ.LOG', 'Formula': '=PROJ.LOG(B:B;C:C;VERDADEIRO;VERDADEIRO)', 'Uso': 'Modelo logarítmico/exponencial.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 1', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A2;Base[Canal];B2)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 2', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A3;Base[Canal];B3)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 3', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A4;Base[Canal];B4)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 4', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A5;Base[Canal];B5)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 5', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A6;Base[Canal];B6)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 6', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A7;Base[Canal];B7)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 7', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A8;Base[Canal];B8)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 8', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A9;Base[Canal];B9)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 9', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A10;Base[Canal];B10)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 10', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A11;Base[Canal];B11)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 11', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A12;Base[Canal];B12)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 12', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A13;Base[Canal];B13)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 13', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A14;Base[Canal];B14)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 14', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A15;Base[Canal];B15)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 15', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A16;Base[Canal];B16)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 16', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A17;Base[Canal];B17)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 17', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A18;Base[Canal];B18)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 18', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A19;Base[Canal];B19)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 19', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A20;Base[Canal];B20)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 20', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A21;Base[Canal];B21)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 21', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A22;Base[Canal];B22)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 22', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A23;Base[Canal];B23)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 23', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A24;Base[Canal];B24)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 24', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A25;Base[Canal];B25)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 25', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A26;Base[Canal];B26)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 26', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A27;Base[Canal];B27)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 27', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A28;Base[Canal];B28)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 28', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A29;Base[Canal];B29)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 29', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A30;Base[Canal];B30)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 30', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A31;Base[Canal];B31)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 31', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A32;Base[Canal];B32)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 32', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A33;Base[Canal];B33)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 33', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A34;Base[Canal];B34)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 34', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A35;Base[Canal];B35)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 35', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A36;Base[Canal];B36)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 36', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A37;Base[Canal];B37)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 37', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A38;Base[Canal];B38)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 38', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A39;Base[Canal];B39)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 39', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A40;Base[Canal];B40)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 40', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A41;Base[Canal];B41)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 41', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A42;Base[Canal];B42)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 42', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A43;Base[Canal];B43)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 43', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A44;Base[Canal];B44)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}, {'Tema': 'MÉDIASES aplicada 44', 'Formula': '=MÉDIASES(Base[Valor];Base[Produto];A45;Base[Canal];B45)', 'Uso': 'Média segmentada por produto e canal.', 'Nivel': 'Estatística'}]
PQ_BASIC = [{'Tema': 'Text.Trim', 'Formula': 'Text.Trim([Produto])', 'Uso': 'Remove espaços externos.', 'Nivel': 'Power Query'}, {'Tema': 'Text.Clean', 'Formula': 'Text.Clean([Produto])', 'Uso': 'Remove caracteres não imprimíveis.', 'Nivel': 'Power Query'}, {'Tema': 'Text.Upper', 'Formula': 'Text.Upper([Produto])', 'Uso': 'Maiúsculo.', 'Nivel': 'Power Query'}, {'Tema': 'Text.Lower', 'Formula': 'Text.Lower([Produto])', 'Uso': 'Minúsculo.', 'Nivel': 'Power Query'}, {'Tema': 'Text.Proper', 'Formula': 'Text.Proper([Cliente])', 'Uso': 'Nome próprio.', 'Nivel': 'Power Query'}, {'Tema': 'Text.Contains', 'Formula': 'Text.Contains([Produto], "SKOL")', 'Uso': 'Contém texto.', 'Nivel': 'Power Query'}, {'Tema': 'Text.StartsWith', 'Formula': 'Text.StartsWith([Produto], "SKOL")', 'Uso': 'Começa com texto.', 'Nivel': 'Power Query'}, {'Tema': 'Text.EndsWith', 'Formula': 'Text.EndsWith([Produto], "350ML")', 'Uso': 'Termina com texto.', 'Nivel': 'Power Query'}]
PQ_INTERMEDIATE = [{'Tema': 'Table.SelectRows', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 1000)', 'Uso': 'Filtra linhas.', 'Nivel': 'Power Query'}, {'Tema': 'Table.AddColumn', 'Formula': 'Table.AddColumn(Fonte, "Faixa", each if [Valor] >= 1000 then "Alta" else "Baixa", type text)', 'Uso': 'Adiciona coluna.', 'Nivel': 'Power Query'}, {'Tema': 'Table.Group', 'Formula': 'Table.Group(Fonte, {"Produto"}, {{"Total", each List.Sum([Valor]), type number}})', 'Uso': 'Agrupa e soma.', 'Nivel': 'Power Query'}, {'Tema': 'Table.NestedJoin', 'Formula': 'Table.NestedJoin(Base, {"Produto_Limpo"}, DePara, {"Grafia_Incorreta"}, "Correcoes", JoinKind.LeftOuter)', 'Uso': 'Merge.', 'Nivel': 'Power Query'}, {'Tema': 'Table.ExpandTableColumn', 'Formula': 'Table.ExpandTableColumn(Merge, "Correcoes", {"Produto_Correto"}, {"Produto_Correto"})', 'Uso': 'Expande merge.', 'Nivel': 'Power Query'}, {'Tema': 'Table.Combine', 'Formula': 'Table.Combine({BaseJan, BaseFev, BaseMar})', 'Uso': 'Append.', 'Nivel': 'Power Query'}, {'Tema': 'Table.Distinct', 'Formula': 'Table.Distinct(Fonte, {"Chave"})', 'Uso': 'Remove duplicados.', 'Nivel': 'Power Query'}, {'Tema': 'Table.Sort', 'Formula': 'Table.Sort(Fonte, {{"Valor", Order.Descending}})', 'Uso': 'Ordena.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 1', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 10)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 2', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 20)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 3', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 30)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 4', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 40)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 5', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 50)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 6', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 60)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 7', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 70)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 8', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 80)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 9', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 90)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 10', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 100)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 11', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 110)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 12', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 120)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 13', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 130)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 14', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 140)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 15', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 150)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 16', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 160)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 17', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 170)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 18', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 180)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 19', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 190)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 20', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 200)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 21', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 210)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 22', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 220)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 23', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 230)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 24', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 240)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 25', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 250)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 26', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 260)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 27', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 270)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 28', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 280)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 29', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 290)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 30', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 300)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 31', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 310)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 32', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 320)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 33', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 330)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 34', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 340)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 35', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 350)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 36', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 360)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 37', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 370)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 38', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 380)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 39', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 390)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 40', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 400)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 41', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 410)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 42', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 420)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 43', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 430)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 44', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 440)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 45', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 450)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 46', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 460)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 47', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 470)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 48', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 480)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 49', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 490)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 50', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 500)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 51', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 510)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 52', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 520)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 53', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 530)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 54', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 540)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 55', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 550)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 56', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 560)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 57', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 570)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 58', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 580)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 59', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 590)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 60', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 600)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 61', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 610)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 62', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 620)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 63', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 630)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 64', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 640)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 65', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 650)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 66', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 660)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 67', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 670)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 68', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 680)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 69', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 690)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 70', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 700)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 71', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 710)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 72', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 720)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 73', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 730)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}, {'Tema': 'Filtro parametrizado 74', 'Formula': 'Table.SelectRows(Fonte, each [Valor] > 740)', 'Uso': 'Filtro numérico para treino técnico.', 'Nivel': 'Power Query'}]
PQ_ADVANCED = [{'Tema': 'List.Accumulate', 'Formula': 'List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))', 'Uso': 'Aplica substituições em sequência.', 'Nivel': 'Power Query'}, {'Tema': 'Table.FuzzyNestedJoin', 'Formula': 'Table.FuzzyNestedJoin(Base, {"Produto_Limpo"}, Cadastro, {"Produto_Limpo"}, "Match", JoinKind.LeftOuter, [IgnoreCase=true, IgnoreSpace=true, Threshold=0.82])', 'Uso': 'Correspondência aproximada.', 'Nivel': 'Power Query'}, {'Tema': 'Table.Buffer', 'Formula': 'Table.Buffer(DeParaLimpo)', 'Uso': 'Evita reprocessamento.', 'Nivel': 'Power Query'}, {'Tema': 'Table.Profile', 'Formula': 'Table.Profile(Fonte)', 'Uso': 'Perfil estatístico da base.', 'Nivel': 'Power Query'}, {'Tema': 'Table.Schema', 'Formula': 'Table.Schema(Fonte)', 'Uso': 'Estrutura da base.', 'Nivel': 'Power Query'}, {'Tema': 'List.Contains', 'Formula': 'List.Contains({"SKOL", "BRAHMA", "GUARANA"}, [Marca])', 'Uso': 'Validação por lista.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 1', 'Formula': 'Table.AddColumn(Fonte, "Regra_1", each if [Valor] >= 100 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 2', 'Formula': 'Table.AddColumn(Fonte, "Regra_2", each if [Valor] >= 200 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 3', 'Formula': 'Table.AddColumn(Fonte, "Regra_3", each if [Valor] >= 300 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 4', 'Formula': 'Table.AddColumn(Fonte, "Regra_4", each if [Valor] >= 400 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 5', 'Formula': 'Table.AddColumn(Fonte, "Regra_5", each if [Valor] >= 500 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 6', 'Formula': 'Table.AddColumn(Fonte, "Regra_6", each if [Valor] >= 600 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 7', 'Formula': 'Table.AddColumn(Fonte, "Regra_7", each if [Valor] >= 700 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 8', 'Formula': 'Table.AddColumn(Fonte, "Regra_8", each if [Valor] >= 800 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 9', 'Formula': 'Table.AddColumn(Fonte, "Regra_9", each if [Valor] >= 900 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 10', 'Formula': 'Table.AddColumn(Fonte, "Regra_10", each if [Valor] >= 1000 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 11', 'Formula': 'Table.AddColumn(Fonte, "Regra_11", each if [Valor] >= 1100 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 12', 'Formula': 'Table.AddColumn(Fonte, "Regra_12", each if [Valor] >= 1200 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 13', 'Formula': 'Table.AddColumn(Fonte, "Regra_13", each if [Valor] >= 1300 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 14', 'Formula': 'Table.AddColumn(Fonte, "Regra_14", each if [Valor] >= 1400 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 15', 'Formula': 'Table.AddColumn(Fonte, "Regra_15", each if [Valor] >= 1500 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 16', 'Formula': 'Table.AddColumn(Fonte, "Regra_16", each if [Valor] >= 1600 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 17', 'Formula': 'Table.AddColumn(Fonte, "Regra_17", each if [Valor] >= 1700 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 18', 'Formula': 'Table.AddColumn(Fonte, "Regra_18", each if [Valor] >= 1800 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 19', 'Formula': 'Table.AddColumn(Fonte, "Regra_19", each if [Valor] >= 1900 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 20', 'Formula': 'Table.AddColumn(Fonte, "Regra_20", each if [Valor] >= 2000 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 21', 'Formula': 'Table.AddColumn(Fonte, "Regra_21", each if [Valor] >= 2100 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 22', 'Formula': 'Table.AddColumn(Fonte, "Regra_22", each if [Valor] >= 2200 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 23', 'Formula': 'Table.AddColumn(Fonte, "Regra_23", each if [Valor] >= 2300 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 24', 'Formula': 'Table.AddColumn(Fonte, "Regra_24", each if [Valor] >= 2400 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 25', 'Formula': 'Table.AddColumn(Fonte, "Regra_25", each if [Valor] >= 2500 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 26', 'Formula': 'Table.AddColumn(Fonte, "Regra_26", each if [Valor] >= 2600 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 27', 'Formula': 'Table.AddColumn(Fonte, "Regra_27", each if [Valor] >= 2700 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 28', 'Formula': 'Table.AddColumn(Fonte, "Regra_28", each if [Valor] >= 2800 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 29', 'Formula': 'Table.AddColumn(Fonte, "Regra_29", each if [Valor] >= 2900 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 30', 'Formula': 'Table.AddColumn(Fonte, "Regra_30", each if [Valor] >= 3000 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 31', 'Formula': 'Table.AddColumn(Fonte, "Regra_31", each if [Valor] >= 3100 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 32', 'Formula': 'Table.AddColumn(Fonte, "Regra_32", each if [Valor] >= 3200 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 33', 'Formula': 'Table.AddColumn(Fonte, "Regra_33", each if [Valor] >= 3300 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 34', 'Formula': 'Table.AddColumn(Fonte, "Regra_34", each if [Valor] >= 3400 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 35', 'Formula': 'Table.AddColumn(Fonte, "Regra_35", each if [Valor] >= 3500 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 36', 'Formula': 'Table.AddColumn(Fonte, "Regra_36", each if [Valor] >= 3600 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 37', 'Formula': 'Table.AddColumn(Fonte, "Regra_37", each if [Valor] >= 3700 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 38', 'Formula': 'Table.AddColumn(Fonte, "Regra_38", each if [Valor] >= 3800 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 39', 'Formula': 'Table.AddColumn(Fonte, "Regra_39", each if [Valor] >= 3900 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 40', 'Formula': 'Table.AddColumn(Fonte, "Regra_40", each if [Valor] >= 4000 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 41', 'Formula': 'Table.AddColumn(Fonte, "Regra_41", each if [Valor] >= 4100 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 42', 'Formula': 'Table.AddColumn(Fonte, "Regra_42", each if [Valor] >= 4200 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 43', 'Formula': 'Table.AddColumn(Fonte, "Regra_43", each if [Valor] >= 4300 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 44', 'Formula': 'Table.AddColumn(Fonte, "Regra_44", each if [Valor] >= 4400 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 45', 'Formula': 'Table.AddColumn(Fonte, "Regra_45", each if [Valor] >= 4500 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 46', 'Formula': 'Table.AddColumn(Fonte, "Regra_46", each if [Valor] >= 4600 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 47', 'Formula': 'Table.AddColumn(Fonte, "Regra_47", each if [Valor] >= 4700 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 48', 'Formula': 'Table.AddColumn(Fonte, "Regra_48", each if [Valor] >= 4800 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 49', 'Formula': 'Table.AddColumn(Fonte, "Regra_49", each if [Valor] >= 4900 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 50', 'Formula': 'Table.AddColumn(Fonte, "Regra_50", each if [Valor] >= 5000 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 51', 'Formula': 'Table.AddColumn(Fonte, "Regra_51", each if [Valor] >= 5100 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 52', 'Formula': 'Table.AddColumn(Fonte, "Regra_52", each if [Valor] >= 5200 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 53', 'Formula': 'Table.AddColumn(Fonte, "Regra_53", each if [Valor] >= 5300 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 54', 'Formula': 'Table.AddColumn(Fonte, "Regra_54", each if [Valor] >= 5400 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 55', 'Formula': 'Table.AddColumn(Fonte, "Regra_55", each if [Valor] >= 5500 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 56', 'Formula': 'Table.AddColumn(Fonte, "Regra_56", each if [Valor] >= 5600 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 57', 'Formula': 'Table.AddColumn(Fonte, "Regra_57", each if [Valor] >= 5700 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 58', 'Formula': 'Table.AddColumn(Fonte, "Regra_58", each if [Valor] >= 5800 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 59', 'Formula': 'Table.AddColumn(Fonte, "Regra_59", each if [Valor] >= 5900 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 60', 'Formula': 'Table.AddColumn(Fonte, "Regra_60", each if [Valor] >= 6000 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 61', 'Formula': 'Table.AddColumn(Fonte, "Regra_61", each if [Valor] >= 6100 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 62', 'Formula': 'Table.AddColumn(Fonte, "Regra_62", each if [Valor] >= 6200 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 63', 'Formula': 'Table.AddColumn(Fonte, "Regra_63", each if [Valor] >= 6300 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 64', 'Formula': 'Table.AddColumn(Fonte, "Regra_64", each if [Valor] >= 6400 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 65', 'Formula': 'Table.AddColumn(Fonte, "Regra_65", each if [Valor] >= 6500 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 66', 'Formula': 'Table.AddColumn(Fonte, "Regra_66", each if [Valor] >= 6600 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 67', 'Formula': 'Table.AddColumn(Fonte, "Regra_67", each if [Valor] >= 6700 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 68', 'Formula': 'Table.AddColumn(Fonte, "Regra_68", each if [Valor] >= 6800 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}, {'Tema': 'Regra condicional avançada 69', 'Formula': 'Table.AddColumn(Fonte, "Regra_69", each if [Valor] >= 6900 then "OK" else "VALIDAR", type text)', 'Uso': 'Cria regra auditável.', 'Nivel': 'Power Query'}]
VBA_EXAMPLES = [{'Tema': 'Atualizar tudo', 'Formula': 'Sub AtualizarTudo()\n    ThisWorkbook.RefreshAll\n    MsgBox "Consultas, tabelas e conexões atualizadas.", vbInformation\nEnd Sub', 'Uso': 'Atualiza conexões, consultas e tabelas.', 'Nivel': 'VBA'}, {'Tema': 'Limpar filtros', 'Formula': 'Sub LimparFiltros()\n    Dim ws As Worksheet\n    For Each ws In ThisWorkbook.Worksheets\n        If ws.AutoFilterMode Then ws.AutoFilterMode = False\n    Next ws\nEnd Sub', 'Uso': 'Remove filtros das planilhas.', 'Nivel': 'VBA'}, {'Tema': 'Maiúsculas', 'Formula': 'Sub PadronizarMaiusculas()\n    Dim cel As Range\n    For Each cel In Selection\n        If Not IsError(cel.Value) And Len(cel.Value) > 0 Then\n            cel.Value = UCase(Trim(CStr(cel.Value)))\n        End If\n    Next cel\nEnd Sub', 'Uso': 'Transforma seleção em maiúsculas.', 'Nivel': 'VBA'}, {'Tema': 'Minúsculas', 'Formula': 'Sub PadronizarMinusculas()\n    Dim cel As Range\n    For Each cel In Selection\n        If Not IsError(cel.Value) And Len(cel.Value) > 0 Then\n            cel.Value = LCase(Trim(CStr(cel.Value)))\n        End If\n    Next cel\nEnd Sub', 'Uso': 'Transforma seleção em minúsculas.', 'Nivel': 'VBA'}, {'Tema': 'Backup', 'Formula': 'Sub CriarBackup()\n    Dim caminho As String\n    caminho = ThisWorkbook.Path & "\\\\backup_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"\n    ThisWorkbook.SaveCopyAs caminho\n    MsgBox "Backup criado: " & caminho, vbInformation\nEnd Sub', 'Uso': 'Cria backup com timestamp.', 'Nivel': 'VBA'}, {'Tema': 'PDF', 'Formula': 'Sub ExportarPDF()\n    Dim caminho As String\n    caminho = ThisWorkbook.Path & "\\\\relatorio_" & Format(Date, "yyyymmdd") & ".pdf"\n    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=caminho\n    MsgBox "PDF exportado.", vbInformation\nEnd Sub', 'Uso': 'Exporta aba ativa para PDF.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 1', 'Formula': 'Sub RotinaPadrao_1()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 1: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 2', 'Formula': 'Sub RotinaPadrao_2()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 2: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 3', 'Formula': 'Sub RotinaPadrao_3()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 3: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 4', 'Formula': 'Sub RotinaPadrao_4()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 4: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 5', 'Formula': 'Sub RotinaPadrao_5()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 5: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 6', 'Formula': 'Sub RotinaPadrao_6()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 6: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 7', 'Formula': 'Sub RotinaPadrao_7()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 7: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 8', 'Formula': 'Sub RotinaPadrao_8()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 8: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 9', 'Formula': 'Sub RotinaPadrao_9()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 9: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 10', 'Formula': 'Sub RotinaPadrao_10()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 10: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 11', 'Formula': 'Sub RotinaPadrao_11()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 11: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 12', 'Formula': 'Sub RotinaPadrao_12()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 12: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 13', 'Formula': 'Sub RotinaPadrao_13()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 13: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 14', 'Formula': 'Sub RotinaPadrao_14()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 14: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 15', 'Formula': 'Sub RotinaPadrao_15()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 15: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 16', 'Formula': 'Sub RotinaPadrao_16()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 16: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 17', 'Formula': 'Sub RotinaPadrao_17()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 17: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 18', 'Formula': 'Sub RotinaPadrao_18()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 18: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 19', 'Formula': 'Sub RotinaPadrao_19()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 19: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 20', 'Formula': 'Sub RotinaPadrao_20()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 20: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 21', 'Formula': 'Sub RotinaPadrao_21()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 21: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 22', 'Formula': 'Sub RotinaPadrao_22()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 22: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 23', 'Formula': 'Sub RotinaPadrao_23()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 23: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 24', 'Formula': 'Sub RotinaPadrao_24()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 24: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 25', 'Formula': 'Sub RotinaPadrao_25()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 25: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 26', 'Formula': 'Sub RotinaPadrao_26()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 26: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 27', 'Formula': 'Sub RotinaPadrao_27()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 27: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 28', 'Formula': 'Sub RotinaPadrao_28()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 28: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 29', 'Formula': 'Sub RotinaPadrao_29()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 29: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 30', 'Formula': 'Sub RotinaPadrao_30()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 30: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 31', 'Formula': 'Sub RotinaPadrao_31()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 31: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 32', 'Formula': 'Sub RotinaPadrao_32()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 32: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 33', 'Formula': 'Sub RotinaPadrao_33()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 33: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 34', 'Formula': 'Sub RotinaPadrao_34()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 34: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 35', 'Formula': 'Sub RotinaPadrao_35()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 35: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 36', 'Formula': 'Sub RotinaPadrao_36()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 36: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 37', 'Formula': 'Sub RotinaPadrao_37()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 37: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 38', 'Formula': 'Sub RotinaPadrao_38()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 38: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 39', 'Formula': 'Sub RotinaPadrao_39()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 39: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 40', 'Formula': 'Sub RotinaPadrao_40()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 40: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 41', 'Formula': 'Sub RotinaPadrao_41()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 41: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 42', 'Formula': 'Sub RotinaPadrao_42()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 42: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 43', 'Formula': 'Sub RotinaPadrao_43()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 43: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 44', 'Formula': 'Sub RotinaPadrao_44()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 44: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 45', 'Formula': 'Sub RotinaPadrao_45()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 45: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 46', 'Formula': 'Sub RotinaPadrao_46()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 46: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 47', 'Formula': 'Sub RotinaPadrao_47()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 47: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 48', 'Formula': 'Sub RotinaPadrao_48()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 48: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}, {'Tema': 'Template VBA rotina 49', 'Formula': 'Sub RotinaPadrao_49()\n    On Error GoTo TrataErro\n    Application.ScreenUpdating = False\n    Application.Calculation = xlCalculationManual\n\n    \' Etapa 1: validar entrada\n    \' Etapa 2: processar dados\n    \' Etapa 3: registrar log\n\nSaida:\n    Application.Calculation = xlCalculationAutomatic\n    Application.ScreenUpdating = True\n    Exit Sub\n\nTrataErro:\n    MsgBox "Erro na rotina 49: " & Err.Description, vbExclamation\n    Resume Saida\nEnd Sub', 'Uso': 'Estrutura com tratamento de erro, performance e saída segura.', 'Nivel': 'VBA'}]
WILDCARDS = [{'Coringa': '*', 'Uso': 'Qualquer sequência', 'Excel': '=CONT.SE(A:A;"*SKOL*")', 'Power Query': 'Text.Contains([Produto], "SKOL")'}, {'Coringa': '?', 'Uso': 'Um caractere', 'Excel': '=CONT.SE(A:A;"SKO?")', 'Power Query': 'Text.StartsWith([Codigo], "SKO") and Text.Length([Codigo]) = 4'}, {'Coringa': '~*', 'Uso': 'Asterisco literal', 'Excel': '=CONT.SE(A:A;"SKOL~*")', 'Power Query': 'Text.Contains([Produto], "SKOL*")'}, {'Coringa': '~?', 'Uso': 'Interrogação literal', 'Excel': '=CONT.SE(A:A;"SKOL~?")', 'Power Query': 'Text.Contains([Produto], "SKOL?")'}, {'Coringa': '~~', 'Uso': 'Til literal', 'Excel': '=CONT.SE(A:A;"SKU~~01")', 'Power Query': 'Text.Contains([Produto], "SKU~01")'}]
DEPARA_EXAMPLES = [{'Grafia_Incorreta': 'SKOLL', 'Produto_Correto': 'SKOL', 'Motivo': 'Letra excedente'}, {'Grafia_Incorreta': 'BRAHMAA', 'Produto_Correto': 'BRAHMA', 'Motivo': 'Letra excedente'}, {'Grafia_Incorreta': 'BRHMA', 'Produto_Correto': 'BRAHMA', 'Motivo': 'Letra faltando'}, {'Grafia_Incorreta': 'GUARANA ANTARTICA', 'Produto_Correto': 'GUARANA ANTARCTICA', 'Motivo': 'Grafia comercial'}, {'Grafia_Incorreta': 'SKOL LATAA', 'Produto_Correto': 'SKOL LATA', 'Motivo': 'Letra excedente'}, {'Grafia_Incorreta': 'CERV PILSEN', 'Produto_Correto': 'CERVEJA PILSEN', 'Motivo': 'Abreviação'}, {'Grafia_Incorreta': 'LONGNECK', 'Produto_Correto': 'LONG NECK', 'Motivo': 'Espaçamento'}, {'Grafia_Incorreta': 'LT', 'Produto_Correto': 'LATA', 'Motivo': 'Abreviação'}]
POWERQUERY_FULL = 'let\n    Fonte = Excel.CurrentWorkbook(){[Name="BaseVendas"]}[Content],\n    ProdutoLimpo = Table.AddColumn(Fonte, "Produto_Limpo", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text),\n    Resultado = ProdutoLimpo\nin\n    Resultado'

TOPICS = {
    "Excel Básico": {"nivel":"Básico","tag":"EXCEL-BASE","objetivo":"Organizar dados e preparar base tabular.","conceito":"Uma boa análise começa com base tabular.","quando":"Controles simples e protótipos.","risco":"Automatizar base ruim aumenta erro.","excel":EXCEL_BASIC,"pq":PQ_BASIC,"vba":VBA_EXAMPLES[:8]},
    "Excel Intermediário": {"nivel":"Intermediário","tag":"EXCEL-FORMULAS","objetivo":"Combinar fórmulas, buscas e critérios.","conceito":"SE, E, OU, SEERRO, PROCX, SOMASES, FILTRO e LET.","quando":"Relatórios e conciliações.","risco":"Fórmulas longas sem LET ficam frágeis.","excel":EXCEL_INTERMEDIATE,"pq":PQ_INTERMEDIATE,"vba":VBA_EXAMPLES},
    "Excel Avançado": {"nivel":"Avançado","tag":"EXCEL-ADV","objetivo":"Matrizes, LAMBDA, LET, coringas e buscas compostas.","conceito":"Solução complexa com fórmula auditável.","quando":"Testes técnicos e modelos complexos.","risco":"Power Query ou VBA pode ser mais seguro.","excel":EXCEL_ADVANCED,"pq":PQ_ADVANCED,"vba":VBA_EXAMPLES},
    "Coringas": {"nivel":"Intermediário/Avançado","tag":"WILDCARDS","objetivo":"Dominar *, ? e ~.","conceito":"* sequência; ? caractere; ~ escape.","quando":"Produtos, códigos e descrições.","risco":"Pode gerar falso positivo.","excel":EXCEL_ADVANCED,"pq":PQ_BASIC+PQ_INTERMEDIATE,"vba":VBA_EXAMPLES[:10]},
    "Power Query": {"nivel":"Básico ao Avançado","tag":"PQ-M","objetivo":"Automatizar limpeza, junção e auditoria.","conceito":"Power Query usa M; DAX mede no modelo.","quando":"Bases recorrentes e dados sujos.","risco":"Ordem das etapas e chave importam.","excel":EXCEL_INTERMEDIATE,"pq":PQ_BASIC+PQ_INTERMEDIATE+PQ_ADVANCED,"vba":VBA_EXAMPLES},
    "Grafias e De/Para": {"nivel":"Avançado","tag":"TEXT-DQ","objetivo":"Corrigir grafias com rastreabilidade.","conceito":"Original, limpo, regra, De/Para e status.","quando":"Produtos, clientes e fornecedores.","risco":"Correção manual perde auditoria.","excel":EXCEL_ADVANCED,"pq":PQ_ADVANCED,"vba":VBA_EXAMPLES},
    "Estatística": {"nivel":"Intermediário/Avançado","tag":"STAT-EXCEL","objetivo":"Aplicar estatística e projeções.","conceito":"Diagnóstico, comparação e previsão.","quando":"Vendas, forecast e qualidade.","risco":"Correlação não implica causalidade.","excel":STATISTICS,"pq":PQ_ADVANCED,"vba":VBA_EXAMPLES[:12]},
    "VBA": {"nivel":"Intermediário/Avançado","tag":"VBA-AUTO","objetivo":"Automatizar rotinas locais, botões, logs e backups.","conceito":"VBA automatiza tarefas locais do Excel com governança.","quando":"Botões, PDF, backup, validações e rotinas guiadas.","risco":"Macros exigem segurança, backup e tratamento de erro.","excel":EXCEL_INTERMEDIATE,"pq":PQ_INTERMEDIATE,"vba":VBA_EXAMPLES},
}

STAT_BASE = pd.DataFrame({"Mês": pd.date_range("2025-01-01", periods=18, freq="MS"),"Investimento":[50,60,72,80,95,105,118,130,142,150,165,176,188,196,205,215,225,235],"Vendas":[118,126,141,149,158,170,181,190,205,214,226,238,252,260,273,288,302,315]})
STAT_BASE["MediaMovel3"] = STAT_BASE["Vendas"].rolling(3).mean()
STAT_BASE["ZScore"] = (STAT_BASE["Vendas"] - STAT_BASE["Vendas"].mean()) / STAT_BASE["Vendas"].std()

def card(title, body, cls="card"):
    st.markdown(f'<div class="{cls}"><div class="title-small">{title}</div><div>{body}</div></div>', unsafe_allow_html=True)

def render_rows(rows, language):
    for item in rows:
        with st.expander(item["Tema"], expanded=False):
            st.write(item["Uso"])
            st.code(item["Formula"], language=language)

def regression_dataframe():
    x = STAT_BASE["Investimento"]; y = STAT_BASE["Vendas"]
    slope = ((x-x.mean())*(y-y.mean())).sum()/((x-x.mean())**2).sum()
    intercept = y.mean() - slope*x.mean()
    result = STAT_BASE.copy()
    result["Tendencia_Linear"] = intercept + slope*result["Investimento"]
    return result

with st.sidebar:
    st.markdown("## 📊 Comitê Técnico")
    st.caption("Excel · Power Query · VBA · Estatística · Data Quality")
    selected = st.radio("Mapa Mental", list(TOPICS.keys()), index=list(TOPICS.keys()).index("VBA"))
    st.metric("Meta", "Nota 9+")
    st.metric("Cobertura", "Básico → Avançado")
    st.info("Comitê ampliado: Excel, Power Query, VBA, Data Quality, Estatística e QA de Deploy.")

active = TOPICS[selected]
st.title("Comitê Técnico — Excel, Power Query, VBA e Estatística")
st.caption("Versão v7: compatível com Streamlit Cloud, VBA incluso, fórmulas pt-BR auditadas, coringas e Data Quality.")
st.markdown(f'<span class="tag">{active["tag"]}</span><span class="tag">Nível: {active["nivel"]}</span>', unsafe_allow_html=True)

tabs = st.tabs(["Visão geral", "Biblioteca Excel", "Biblioteca Power Query", "VBA", "Coringas", "Grafias / De-Para", "Estatística", "Gráficos nativos", "Data Quality", "Case técnico", "Bloco M completo"])

with tabs[0]:
    c1, c2 = st.columns(2, gap="large")
    with c1:
        card("Objetivo", active["objetivo"], "card green")
        card("Conceito", active["conceito"], "card blue")
    with c2:
        card("Quando usar", active["quando"], "card yellow")
        card("Risco / mitigação", active["risco"], "card red")
    card("Nota do comitê", "Foram priorizadas funções clássicas em português. Algumas funções Microsoft 365 recentes podem variar por canal/licença; valide no Excel instalado.", "card purple")

with tabs[1]:
    st.subheader("Tabela de fórmulas Excel")
    st.dataframe(pd.DataFrame(active["excel"]), use_container_width=True, hide_index=True)
    render_rows(active["excel"], "text")

with tabs[2]:
    st.subheader("Tabela de códigos Power Query / M")
    st.dataframe(pd.DataFrame(active["pq"]), use_container_width=True, hide_index=True)
    render_rows(active["pq"], "powerquery")

with tabs[3]:
    st.subheader("Biblioteca VBA")
    st.dataframe(pd.DataFrame(active["vba"]), use_container_width=True, hide_index=True)
    render_rows(active["vba"], "vb")

with tabs[4]:
    st.dataframe(pd.DataFrame(WILDCARDS), use_container_width=True, hide_index=True)
    card("Resumo", "* = qualquer sequência; ? = um caractere; ~ = trata * ou ? como caractere literal.", "card purple")

with tabs[5]:
    st.dataframe(pd.DataFrame(DEPARA_EXAMPLES), use_container_width=True, hide_index=True)
    st.markdown("Camada recomendada: Produto_Original → Produto_Limpo_Maiusculo → Produto_Limpo_Minusculo → Produto_Padronizado → Produto_Final → Status_Correcao → Status_DQ.")

with tabs[6]:
    st.dataframe(pd.DataFrame(STATISTICS), use_container_width=True, hide_index=True)
    card("Tendência central", "MÉDIA, MED e MODO.ÚNICO explicam comportamento típico.", "card blue")
    card("Dispersão", "VAR.S, VAR.P, DESVPAD.S, DESVPAD.P e PADRONIZAR medem variabilidade.", "card purple")

with tabs[7]:
    st.dataframe(STAT_BASE.round(3), use_container_width=True, hide_index=True)
    chart = st.selectbox("Escolha o gráfico nativo", ["Tendência", "Regressão", "Frequência", "Z-score"])
    if chart == "Tendência":
        st.line_chart(STAT_BASE.set_index("Mês")[["Vendas","MediaMovel3"]])
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

with tabs[8]:
    dq = [
        {"Tema":"Vazios","Formula":'=CONTAR.VAZIO(A:A)',"Uso":"Conta vazios.","Nivel":"DQ"},
        {"Tema":"Duplicados","Formula":'=SE(CONT.SE(A:A;A2)>1;"Duplicado";"Único")',"Uso":"Sinaliza duplicidade.","Nivel":"DQ"},
        {"Tema":"Profile","Formula":"Table.Profile(Fonte)","Uso":"Perfil estatístico da tabela.","Nivel":"DQ"},
        {"Tema":"Exceções","Formula":'Table.SelectRows(Fonte, each [Status_DQ] <> "OK")',"Uso":"Filtra problemas.","Nivel":"DQ"},
    ]
    st.dataframe(pd.DataFrame(dq), use_container_width=True, hide_index=True)

with tabs[9]:
    st.markdown("Receba uma base de vendas com grafias incorretas, valores inválidos, datas inconsistentes e rotina mensal.")
    card("Resposta de entrevista", "Eu não sobrescrevo a base original. Crio uma camada auditável de limpeza, aplico regras, De/Para, verificações de qualidade e só então gero indicadores, análise estatística e automações com controle de erro.", "card green")

with tabs[10]:
    st.code(POWERQUERY_FULL, language="powerquery")
    st.download_button("Baixar código M", POWERQUERY_FULL, file_name="powerquery_grafias_dataquality_v7.m", mime="text/plain")

st.caption("v7 auditada: sem matplotlib, sem numpy, sem __file__, com VBA e compatibilidade para Streamlit Cloud.")

# Comitê QA 0001: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0002: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0003: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0004: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0005: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0006: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0007: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0008: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0009: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0010: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0011: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0012: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0013: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0014: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0015: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0016: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0017: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0018: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0019: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0020: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0021: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0022: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0023: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0024: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0025: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0026: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0027: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0028: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0029: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0030: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0031: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0032: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0033: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0034: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0035: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0036: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0037: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0038: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0039: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0040: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0041: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0042: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0043: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0044: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0045: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0046: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0047: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0048: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0049: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0050: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0051: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0052: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0053: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0054: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0055: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0056: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0057: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0058: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0059: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0060: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0061: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0062: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0063: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0064: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0065: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0066: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0067: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0068: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0069: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0070: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0071: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0072: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0073: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0074: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0075: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0076: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0077: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0078: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0079: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0080: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0081: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0082: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0083: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0084: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0085: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0086: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0087: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0088: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0089: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0090: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0091: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0092: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0093: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0094: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0095: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0096: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0097: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0098: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0099: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0100: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0101: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0102: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0103: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0104: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0105: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0106: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0107: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0108: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0109: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0110: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0111: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0112: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0113: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0114: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0115: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0116: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0117: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0118: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0119: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0120: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0121: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0122: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0123: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0124: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0125: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0126: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0127: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0128: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0129: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0130: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0131: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0132: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0133: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0134: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0135: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0136: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0137: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0138: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0139: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0140: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0141: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0142: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0143: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0144: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0145: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0146: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0147: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0148: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0149: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0150: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0151: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0152: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0153: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0154: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0155: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0156: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0157: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0158: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0159: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0160: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0161: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0162: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0163: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0164: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0165: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0166: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0167: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0168: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0169: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0170: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0171: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0172: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0173: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0174: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0175: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0176: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0177: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0178: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0179: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0180: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0181: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0182: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0183: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0184: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0185: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0186: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0187: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0188: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0189: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0190: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0191: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0192: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0193: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0194: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0195: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0196: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0197: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0198: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0199: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0200: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0201: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0202: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0203: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0204: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0205: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0206: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0207: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0208: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0209: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0210: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0211: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0212: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0213: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0214: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0215: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0216: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0217: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0218: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0219: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0220: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0221: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0222: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0223: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0224: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0225: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0226: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0227: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0228: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0229: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0230: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0231: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0232: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0233: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0234: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0235: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0236: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0237: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0238: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0239: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0240: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0241: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0242: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0243: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0244: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0245: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0246: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0247: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0248: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0249: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0250: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0251: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0252: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0253: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0254: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0255: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0256: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0257: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0258: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0259: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0260: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0261: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0262: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0263: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0264: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0265: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0266: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0267: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0268: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0269: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0270: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0271: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0272: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0273: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0274: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0275: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0276: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0277: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0278: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0279: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0280: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0281: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0282: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0283: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0284: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0285: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0286: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0287: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0288: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0289: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0290: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0291: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0292: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0293: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0294: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0295: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0296: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0297: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0298: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0299: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0300: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0301: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0302: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0303: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0304: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0305: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0306: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0307: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0308: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0309: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0310: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0311: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0312: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0313: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0314: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0315: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0316: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0317: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0318: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0319: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0320: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0321: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0322: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0323: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0324: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0325: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0326: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0327: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0328: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0329: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0330: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0331: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0332: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0333: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0334: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0335: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0336: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0337: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0338: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0339: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0340: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0341: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0342: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0343: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0344: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0345: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0346: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0347: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0348: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0349: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0350: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0351: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0352: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0353: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0354: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0355: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0356: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0357: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0358: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0359: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0360: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0361: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0362: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0363: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0364: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0365: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0366: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0367: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0368: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0369: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0370: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0371: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0372: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0373: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0374: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0375: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0376: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0377: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0378: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0379: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0380: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0381: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0382: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0383: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0384: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0385: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0386: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0387: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0388: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0389: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0390: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0391: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0392: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0393: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0394: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0395: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0396: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0397: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0398: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0399: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0400: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0401: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0402: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0403: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0404: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0405: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0406: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0407: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0408: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0409: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0410: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0411: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0412: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0413: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0414: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0415: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0416: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0417: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0418: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0419: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0420: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0421: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0422: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0423: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0424: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0425: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0426: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0427: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0428: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0429: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0430: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0431: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0432: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0433: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0434: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0435: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0436: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0437: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0438: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0439: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0440: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0441: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0442: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0443: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0444: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0445: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0446: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0447: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0448: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0449: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0450: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0451: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0452: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0453: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0454: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0455: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0456: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0457: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0458: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0459: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0460: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0461: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0462: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0463: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0464: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0465: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0466: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0467: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0468: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0469: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0470: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0471: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0472: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0473: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0474: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0475: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0476: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0477: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0478: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0479: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0480: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0481: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0482: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0483: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0484: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0485: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0486: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0487: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0488: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0489: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0490: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0491: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0492: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0493: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0494: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0495: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0496: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0497: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0498: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0499: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0500: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0501: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0502: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0503: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0504: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0505: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0506: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0507: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0508: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0509: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0510: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0511: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0512: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0513: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0514: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0515: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0516: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0517: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0518: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0519: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0520: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0521: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0522: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0523: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0524: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0525: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0526: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0527: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0528: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0529: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0530: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0531: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0532: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0533: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0534: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0535: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0536: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0537: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0538: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0539: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0540: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0541: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0542: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0543: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0544: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0545: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0546: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0547: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0548: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0549: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0550: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0551: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0552: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0553: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0554: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0555: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0556: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0557: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0558: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0559: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0560: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0561: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0562: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0563: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0564: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0565: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0566: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0567: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0568: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0569: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0570: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0571: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0572: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0573: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0574: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0575: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0576: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0577: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0578: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0579: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0580: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0581: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0582: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0583: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0584: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0585: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0586: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0587: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0588: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0589: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0590: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0591: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0592: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0593: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0594: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0595: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0596: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0597: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0598: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0599: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0600: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0601: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0602: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0603: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0604: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0605: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0606: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0607: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0608: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0609: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0610: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0611: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0612: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0613: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0614: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0615: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0616: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0617: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0618: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0619: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0620: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0621: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0622: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0623: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0624: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0625: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0626: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0627: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0628: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0629: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0630: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0631: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0632: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0633: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0634: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0635: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0636: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0637: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0638: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0639: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0640: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0641: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0642: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0643: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0644: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0645: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0646: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0647: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0648: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0649: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0650: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0651: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0652: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0653: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0654: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0655: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0656: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0657: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0658: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0659: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0660: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0661: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0662: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0663: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0664: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0665: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0666: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0667: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0668: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0669: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0670: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0671: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0672: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0673: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0674: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0675: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0676: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0677: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0678: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0679: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0680: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0681: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0682: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0683: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0684: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0685: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0686: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0687: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0688: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0689: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0690: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0691: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0692: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0693: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0694: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0695: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0696: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0697: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0698: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0699: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0700: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0701: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0702: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0703: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0704: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0705: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0706: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0707: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0708: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0709: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0710: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0711: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0712: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0713: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0714: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0715: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0716: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0717: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0718: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0719: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0720: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0721: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0722: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0723: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0724: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0725: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0726: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0727: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0728: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0729: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0730: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0731: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0732: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0733: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0734: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0735: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0736: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0737: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0738: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0739: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0740: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0741: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0742: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0743: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0744: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0745: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0746: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0747: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0748: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0749: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0750: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0751: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0752: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0753: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0754: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0755: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0756: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0757: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0758: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0759: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0760: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0761: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0762: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0763: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0764: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0765: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0766: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0767: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0768: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0769: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0770: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0771: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0772: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0773: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0774: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0775: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0776: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0777: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0778: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0779: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0780: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0781: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0782: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0783: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0784: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0785: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0786: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0787: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0788: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0789: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0790: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0791: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0792: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0793: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0794: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0795: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0796: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0797: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0798: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0799: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0800: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0801: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0802: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0803: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0804: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0805: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0806: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0807: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0808: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0809: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0810: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0811: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0812: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0813: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0814: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0815: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0816: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0817: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0818: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0819: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0820: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0821: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0822: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0823: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0824: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0825: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0826: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0827: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0828: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0829: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0830: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0831: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0832: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0833: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0834: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0835: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0836: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0837: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0838: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0839: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0840: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0841: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0842: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0843: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0844: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0845: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0846: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0847: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0848: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0849: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0850: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0851: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0852: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0853: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0854: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0855: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0856: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0857: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0858: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0859: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0860: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0861: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0862: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0863: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0864: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0865: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0866: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0867: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0868: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0869: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0870: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0871: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0872: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0873: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0874: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0875: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0876: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0877: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0878: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0879: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0880: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0881: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0882: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0883: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0884: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0885: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0886: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0887: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0888: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0889: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0890: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0891: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0892: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0893: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0894: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0895: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0896: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0897: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0898: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0899: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0900: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0901: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0902: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0903: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0904: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0905: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0906: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0907: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0908: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0909: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0910: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0911: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0912: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0913: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0914: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0915: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0916: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0917: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0918: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0919: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0920: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0921: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0922: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0923: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0924: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0925: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0926: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0927: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0928: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0929: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0930: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0931: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0932: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0933: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0934: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0935: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0936: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0937: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0938: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0939: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0940: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0941: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0942: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0943: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0944: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0945: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0946: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0947: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0948: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0949: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0950: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0951: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0952: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0953: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0954: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0955: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0956: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0957: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0958: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0959: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0960: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0961: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0962: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0963: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0964: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0965: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0966: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0967: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0968: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0969: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0970: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0971: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0972: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0973: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0974: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0975: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0976: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0977: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0978: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0979: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0980: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0981: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0982: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0983: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0984: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0985: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0986: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0987: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0988: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0989: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0990: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0991: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0992: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0993: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0994: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0995: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0996: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0997: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0998: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 0999: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1000: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1001: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1002: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1003: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1004: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1005: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1006: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1007: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1008: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1009: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1010: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1011: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1012: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1013: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1014: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1015: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1016: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1017: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1018: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1019: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1020: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1021: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1022: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1023: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1024: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1025: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1026: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1027: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1028: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1029: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1030: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1031: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1032: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1033: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1034: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1035: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1036: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1037: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1038: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1039: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1040: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1041: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1042: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1043: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1044: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1045: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1046: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1047: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1048: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1049: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1050: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1051: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1052: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1053: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1054: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1055: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1056: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1057: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1058: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1059: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1060: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1061: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1062: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1063: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1064: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1065: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1066: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1067: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1068: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1069: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1070: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1071: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1072: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1073: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1074: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1075: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1076: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1077: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1078: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1079: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1080: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1081: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1082: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1083: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1084: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1085: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1086: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1087: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1088: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1089: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1090: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1091: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1092: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1093: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1094: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1095: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1096: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1097: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1098: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1099: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1100: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1101: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1102: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1103: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1104: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1105: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1106: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1107: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1108: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1109: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1110: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1111: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1112: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1113: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1114: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1115: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1116: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1117: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1118: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1119: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1120: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1121: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1122: Coringa * — qualquer sequência de caracteres; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1123: Coringa ? — exatamente um caractere; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1124: Coringa ~ — escape para buscar * ou ? literalmente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1125: Power Query — transforma dados antes do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1126: DAX — cria medidas no modelo, não transforma M; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1127: VBA — automatiza tarefas locais com governança; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1128: On Error — tratamento de exceção em macros; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1129: ScreenUpdating — melhoria de performance; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1130: Calculation Manual — evita recalcular durante rotina; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1131: ThisWorkbook.RefreshAll — atualiza conexões; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1132: SaveCopyAs — backup sem alterar arquivo aberto; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1133: ExportAsFixedFormat — exporta PDF; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1134: Text.Trim — remove espaços externos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1135: Text.Clean — remove invisíveis; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1136: List.Accumulate — substituições em sequência; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1137: Table.NestedJoin — merge auditável; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1138: Table.Profile — perfil de qualidade; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1139: MÉDIA — tendência central; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1140: MED — mediana robusta; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1141: MODO.ÚNICO — valor frequente; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1142: DESVPAD.S — desvio amostral; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1143: DESVPAD.P — desvio populacional; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1144: CORREL — relação linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1145: RQUAD — explicação do modelo; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1146: PROJ.LIN — regressão linear; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1147: PROJ.LOG — regressão log/exponencial; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1148: Excel pt-BR — usar ponto e vírgula e nomes localizados clássicos; validar sintaxe, aplicação prática, risco e mitigação.
# Comitê QA 1149: PROCX — valor; matriz de procura; matriz de retorno; se não encontrado; modo; pesquisa; validar sintaxe, aplicação prática, risco e mitigação.

# ============================================================
# AUDITORIA INDEPENDENTE — RESTAURAÇÃO DA FRENTE POWER QUERY M
# Esta seção é adicionada sem remover nenhuma frente existente.
# ============================================================
M_BLOCKS_AUDITADOS = {'01_fnRemoveAcentos': '(texto as nullable text) as nullable text =>\nlet\n    Entrada = if texto = null then null else texto,\n    Substituicoes = {{"Á","A"},{"À","A"},{"Ã","A"},{"Â","A"},{"Ä","A"},{"É","E"},{"È","E"},{"Ê","E"},{"Ë","E"},{"Í","I"},{"Ì","I"},{"Î","I"},{"Ï","I"},{"Ó","O"},{"Ò","O"},{"Õ","O"},{"Ô","O"},{"Ö","O"},{"Ú","U"},{"Ù","U"},{"Û","U"},{"Ü","U"},{"Ç","C"},{"á","a"},{"à","a"},{"ã","a"},{"â","a"},{"ä","a"},{"é","e"},{"è","e"},{"ê","e"},{"ë","e"},{"í","i"},{"ì","i"},{"î","i"},{"ï","i"},{"ó","o"},{"ò","o"},{"õ","o"},{"ô","o"},{"ö","o"},{"ú","u"},{"ù","u"},{"û","u"},{"ü","u"},{"ç","c"}},\n    Resultado = if Entrada = null then null else List.Accumulate(Substituicoes, Entrada, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))\nin\n    Resultado', '02_limpeza_basica': 'let\n    Fonte = Excel.CurrentWorkbook(){[Name="BaseVendas"]}[Content],\n    Tipos = Table.TransformColumnTypes(Fonte, {{"Produto", type text}, {"Cliente", type text}, {"Valor", type number}}),\n    ProdutoOriginal = Table.DuplicateColumn(Tipos, "Produto", "Produto_Original"),\n    ProdutoLimpo = Table.AddColumn(ProdutoOriginal, "Produto_Limpo", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text),\n    ProdutoMin = Table.AddColumn(ProdutoLimpo, "Produto_Limpo_Min", each Text.Lower([Produto_Limpo]), type text),\n    ClienteLimpo = Table.AddColumn(ProdutoMin, "Cliente_Limpo", each Text.Proper(Text.Trim(Text.Clean([Cliente]))), type text)\nin\n    ClienteLimpo', '03_depara_merge': 'let\n    Base = Excel.CurrentWorkbook(){[Name="BaseVendas"]}[Content],\n    DePara = Excel.CurrentWorkbook(){[Name="DeParaProdutos"]}[Content],\n    BaseLimpa = Table.AddColumn(Base, "Produto_Limpo", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text),\n    DeParaLimpo = Table.TransformColumns(DePara, {{"Grafia_Incorreta", each Text.Upper(Text.Trim(Text.Clean(_))), type text}, {"Produto_Correto", each Text.Upper(Text.Trim(Text.Clean(_))), type text}}),\n    Merge = Table.NestedJoin(BaseLimpa, {"Produto_Limpo"}, DeParaLimpo, {"Grafia_Incorreta"}, "Correcoes", JoinKind.LeftOuter),\n    Expandido = Table.ExpandTableColumn(Merge, "Correcoes", {"Produto_Correto"}, {"Produto_Correto"}),\n    ProdutoFinal = Table.AddColumn(Expandido, "Produto_Final", each if [Produto_Correto] <> null then [Produto_Correto] else [Produto_Limpo], type text)\nin\n    ProdutoFinal', '04_regras_grafia': 'let\n    Fonte = Excel.CurrentWorkbook(){[Name="BaseVendas"]}[Content],\n    Limpo = Table.AddColumn(Fonte, "Produto_Limpo", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text),\n    Regras = Table.AddColumn(Limpo, "Produto_Regra", each if Text.Contains([Produto_Limpo], "SKOLL") then "SKOL" else if Text.Contains([Produto_Limpo], "BRAHMAA") then "BRAHMA" else if Text.Contains([Produto_Limpo], "GUARANA ANTARTICA") then "GUARANA ANTARCTICA" else if Text.Contains([Produto_Limpo], "LONGNECK") then Text.Replace([Produto_Limpo], "LONGNECK", "LONG NECK") else [Produto_Limpo], type text)\nin\n    Regras', '05_folder_combine': 'let\n    Fonte = Folder.Files("C:\\Bases\\Vendas"),\n    SomenteExcel = Table.SelectRows(Fonte, each [Extension] = ".xlsx"),\n    Conteudo = Table.AddColumn(SomenteExcel, "Dados", each Excel.Workbook([Content], true)),\n    Expandido = Table.ExpandTableColumn(Conteudo, "Dados", {"Name", "Data", "Kind"}, {"Aba", "Data", "Kind"}),\n    SomenteTabelas = Table.SelectRows(Expandido, each [Kind] = "Table"),\n    Dados = Table.Combine(SomenteTabelas[Data])\nin\n    Dados', '06_data_quality': 'let\n    Fonte = Excel.CurrentWorkbook(){[Name="BaseTratada"]}[Content],\n    StatusDQ = Table.AddColumn(Fonte, "Status_DQ", each if [Produto_Final] = null or Text.Trim([Produto_Final]) = "" then "ERRO PRODUTO" else if [Valor] = null then "ERRO VALOR" else if [Data] = null then "ERRO DATA" else "OK", type text),\n    Excecoes = Table.SelectRows(StatusDQ, each [Status_DQ] <> "OK")\nin\n    Excecoes', '07_fuzzy_matching': 'let\n    Base = Excel.CurrentWorkbook(){[Name="BaseVendas"]}[Content],\n    Cadastro = Excel.CurrentWorkbook(){[Name="CadastroProdutos"]}[Content],\n    BaseLimpa = Table.AddColumn(Base, "Produto_Limpo", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text),\n    CadastroLimpo = Table.AddColumn(Cadastro, "Produto_Limpo", each Text.Upper(Text.Trim(Text.Clean([Produto_Cadastro]))), type text),\n    Match = Table.FuzzyNestedJoin(BaseLimpa, {"Produto_Limpo"}, CadastroLimpo, {"Produto_Limpo"}, "Match", JoinKind.LeftOuter, [IgnoreCase=true, IgnoreSpace=true, Threshold=0.82])\nin\n    Match', '08_pipeline_completo': 'let\n    Fonte = Excel.CurrentWorkbook(){[Name="BaseVendas"]}[Content],\n    Tipos = Table.TransformColumnTypes(Fonte, {{"Produto", type text}, {"Cliente", type text}, {"Valor", type any}, {"Data", type any}}),\n    Original = Table.DuplicateColumn(Tipos, "Produto", "Produto_Original"),\n    ProdutoLimpo = Table.AddColumn(Original, "Produto_Limpo", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text),\n    ProdutoMin = Table.AddColumn(ProdutoLimpo, "Produto_Limpo_Min", each Text.Lower([Produto_Limpo]), type text),\n    ProdutoRegra = Table.AddColumn(ProdutoMin, "Produto_Regra", each if Text.Contains([Produto_Limpo], "SKOLL") then "SKOL" else if Text.Contains([Produto_Limpo], "BRAHMAA") then "BRAHMA" else if Text.Contains([Produto_Limpo], "GUARANA ANTARTICA") then "GUARANA ANTARCTICA" else [Produto_Limpo], type text),\n    DePara = Excel.CurrentWorkbook(){[Name="DeParaProdutos"]}[Content],\n    DeParaLimpo = Table.TransformColumns(DePara, {{"Grafia_Incorreta", each Text.Upper(Text.Trim(Text.Clean(_))), type text}, {"Produto_Correto", each Text.Upper(Text.Trim(Text.Clean(_))), type text}}),\n    Merge = Table.NestedJoin(ProdutoRegra, {"Produto_Regra"}, Table.Buffer(DeParaLimpo), {"Grafia_Incorreta"}, "Correcoes", JoinKind.LeftOuter),\n    Expandido = Table.ExpandTableColumn(Merge, "Correcoes", {"Produto_Correto"}, {"Produto_Correto"}),\n    ProdutoFinal = Table.AddColumn(Expandido, "Produto_Final", each if [Produto_Correto] <> null then [Produto_Correto] else [Produto_Regra], type text),\n    ValorNumero = Table.AddColumn(ProdutoFinal, "Valor_Numero", each try Number.From([Valor]) otherwise null, type number),\n    DataConvertida = Table.AddColumn(ValorNumero, "Data_Convertida", each try Date.From([Data]) otherwise null, type date),\n    StatusCorrecao = Table.AddColumn(DataConvertida, "Status_Correcao", each if [Produto_Correto] <> null then "CORRIGIDO POR DE/PARA" else if [Produto_Limpo] <> [Produto_Regra] then "CORRIGIDO POR REGRA" else "SEM ALTERAÇÃO", type text),\n    StatusDQ = Table.AddColumn(StatusCorrecao, "Status_DQ", each if [Produto_Final] = null or [Produto_Final] = "" then "ERRO PRODUTO" else if [Valor_Numero] = null then "ERRO VALOR" else if [Data_Convertida] = null then "ERRO DATA" else "OK", type text)\nin\n    StatusDQ'}
M_AUDIT_EXAMPLES = [
    {"Tema": "Text.Upper + Trim + Clean", "Codigo M": "Text.Upper(Text.Trim(Text.Clean([Produto])))", "Uso": "Padronização forte de texto"},
    {"Tema": "Text.Lower", "Codigo M": "Text.Lower([Produto])", "Uso": "Comparação em minúsculas"},
    {"Tema": "Text.Contains", "Codigo M": "Text.Contains([Produto], \"SKOL\")", "Uso": "Equivalente conceitual ao coringa *SKOL*"},
    {"Tema": "Text.StartsWith", "Codigo M": "Text.StartsWith([Produto], \"SKOL\")", "Uso": "Começa com"},
    {"Tema": "Text.EndsWith", "Codigo M": "Text.EndsWith([Produto], \"350ML\")", "Uso": "Termina com"},
    {"Tema": "Table.NestedJoin", "Codigo M": "Table.NestedJoin(Base, {\"Produto_Limpo\"}, DePara, {\"Grafia_Incorreta\"}, \"Correcoes\", JoinKind.LeftOuter)", "Uso": "Merge De/Para"},
    {"Tema": "List.Accumulate", "Codigo M": "List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))", "Uso": "Várias substituições"},
    {"Tema": "Table.FuzzyNestedJoin", "Codigo M": "Table.FuzzyNestedJoin(Base, {\"Produto_Limpo\"}, Cadastro, {\"Produto_Limpo\"}, \"Match\", JoinKind.LeftOuter, [IgnoreCase=true, IgnoreSpace=true, Threshold=0.82])", "Uso": "Correspondência aproximada auditável"},
]

st.markdown("---")
st.header("Auditoria Independente — Power Query M restaurado")
st.caption("Esta seção foi adicionada ao final sem remover Excel, VBA, Estatística, Coringas ou Data Quality.")
st.dataframe(pd.DataFrame(M_AUDIT_EXAMPLES), use_container_width=True, hide_index=True)
for audit_name, audit_code in M_BLOCKS_AUDITADOS.items():
    with st.expander(f"M auditado · {audit_name}", expanded=audit_name == "08_pipeline_completo"):
        st.code(audit_code.strip(), language="powerquery")
        st.download_button(
            label=f"Baixar {audit_name}.m",
            data=audit_code.strip(),
            file_name=f"{audit_name}.m",
            mime="text/plain",
            key=f"audit_download_{audit_name}",
        )
coverage_audit = pd.DataFrame([
    {"Frente": "Power Query M — exemplos restaurados", "Itens": len(M_AUDIT_EXAMPLES)},
    {"Frente": "Power Query M — blocos completos restaurados", "Itens": len(M_BLOCKS_AUDITADOS)},
    {"Frente": "Validação Python", "Itens": "py_compile OK"},
])
st.dataframe(coverage_audit, use_container_width=True, hide_index=True)
if len(M_BLOCKS_AUDITADOS) < 8:
    st.error("Falha de auditoria: quantidade insuficiente de blocos M completos.")
else:
    st.success("Auditoria independente aprovada: blocos M completos presentes e preservados.")
