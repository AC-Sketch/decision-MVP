from __future__ import annotations

# =============================================================================
# BLOCO 08 + BLOCO 09 — VERSÃO COMPLETA INTEGRADA V14
# =============================================================================
# Conteúdo:
# 1. Bloco 08 — Exercícios resolvidos Excel + Power Query V14
# 2. Bloco 09 — Excel Decision Framework V14
#
# Motivo da integração:
# O Bloco 09 depende da lista EXERCICIOS_RESOLVIDOS criada no Bloco 08.
# Portanto, para rodar como bloco completo único, os dois precisam estar juntos.
# =============================================================================


# BLOCO 08 — EXERCÍCIOS RESOLVIDOS EXCEL + POWER QUERY V14
# Objetivo:
# - Criar enunciado e solução completa para cada exercício de Excel e Power Query já existente.
# - Este bloco é autônomo: pode rodar no Colab, localmente ou ser colado no app.py consolidado.
# - Não depende de módulos externos do projeto.
#
# Origem:
# - App único V14 enviado pelo usuário.
# - Exercícios extraídos dos payloads Excel e Power Query.
#
# Como usar:
# 1. Rode este bloco no Colab para validar o conteúdo.
# 2. Para integrar ao Streamlit, use render_exercicios_resolvidos_app() dentro do app principal.
# 3. Para exportar os dados, use export_exercicios_resolvidos_payload().

import json
import datetime as dt
from typing import Any, Dict, List

try:
    import pandas as pd
except ModuleNotFoundError:
    pd = None

try:
    import streamlit as st
    STREAMLIT_AVAILABLE = True
except ModuleNotFoundError:
    st = None
    STREAMLIT_AVAILABLE = False


APP_VERSION = "v14.3.0"
BLOCK_NAME = "bloco_08_exercicios_resolvidos_excel_powerquery"
BLOCK_STATUS = "COMPLETO_CORE_PENDENTE_USUARIO"
BUILD_DATE = "2026-06-22"

MIN_EXCEL_EXERCISES = 100
MIN_POWERQUERY_EXERCISES = 100
MIN_TOTAL_EXERCISES = 200


EXERCICIOS_RESOLVIDOS: List[Dict[str, Any]] = [
    {
        "id": "EXC-001",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Básico",
        "area_negocio": "RH",
        "enunciado": "EXC-001 — Texto (Básico)\n\nContexto de negócio: O RH recebeu nomes com espaços excedentes e precisa padronizar cadastro.\n\nProblema: Limpar nome em A2, removendo espaços extras e colocando em nome próprio.\n\nBase simulada: A2 = '  joão   da   silva  '\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "O dado possui espaços irregulares e caixa inadequada.",
        "raciocinio": "Primeiro remover espaços excedentes com ARRUMAR, depois aplicar PRI.MAIÚSCULA.",
        "solucao_principal": "=PRI.MAIÚSCULA(ARRUMAR(A2))",
        "solucao_excel": "=PRI.MAIÚSCULA(ARRUMAR(A2))",
        "solucao_powerquery": "Text.Proper(Text.Trim(Text.Clean([Nome])))",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "João Da Silva",
        "erros_comuns": "Aplicar PRI.MAIÚSCULA sem ARRUMAR e manter espaços indevidos.",
        "alternativas": "Power Query é melhor quando a limpeza for recorrente.",
        "aplicacao_pratica": "Cadastro de colaboradores, clientes e fornecedores.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-002",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Intermediário",
        "area_negocio": "TI/RH",
        "enunciado": "EXC-002 — Texto (Intermediário)\n\nContexto de negócio: A empresa quer gerar e-mails corporativos com base em nome e sobrenome.\n\nProblema: Criar e-mail nome.sobrenome@empresa.com.br em minúsculas e com limpeza básica.\n\nBase simulada: B2 = José Maria | C2 = Gonçalves\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "A fórmula precisa concatenar, converter caixa e tratar acentos mapeados.",
        "raciocinio": "Usar MINÚSCULA, ARRUMAR, SUBSTITUIR e CONCAT.",
        "solucao_principal": "=CONCAT(SUBSTITUIR(SUBSTITUIR(MINÚSCULA(ARRUMAR(B2));\"é\";\"e\");\"ã\";\"a\");\".\";SUBSTITUIR(SUBSTITUIR(SUBSTITUIR(MINÚSCULA(ARRUMAR(C2));\"é\";\"e\");\"ó\";\"o\");\"ç\";\"c\");\"@empresa.com.br\")",
        "solucao_excel": "=CONCAT(SUBSTITUIR(SUBSTITUIR(MINÚSCULA(ARRUMAR(B2));\"é\";\"e\");\"ã\";\"a\");\".\";SUBSTITUIR(SUBSTITUIR(SUBSTITUIR(MINÚSCULA(ARRUMAR(C2));\"é\";\"e\");\"ó\";\"o\");\"ç\";\"c\");\"@empresa.com.br\")",
        "solucao_powerquery": "Power Query: criar função de remoção de acentos e Text.Combine.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "jose maria.goncalves@empresa.com.br",
        "erros_comuns": "Não remover espaços internos de nome composto quando a regra exigir.",
        "alternativas": "Usar Power Query para padronização de muitos acentos.",
        "aplicacao_pratica": "Onboarding de colaboradores.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-003",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Intermediário",
        "area_negocio": "Comercial",
        "enunciado": "EXC-003 — Texto (Intermediário)\n\nContexto de negócio: Produtos vêm com abreviações diferentes no cadastro.\n\nProblema: Trocar ' LT ' por ' LATA ' na descrição.\n\nBase simulada: A2 = 'SKOL LT 350ML'\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Abreviação precisa virar descrição padronizada.",
        "raciocinio": "SUBSTITUIR corrige o trecho textual.",
        "solucao_principal": "=SUBSTITUIR(A2;\" LT \";\" LATA \")",
        "solucao_excel": "=SUBSTITUIR(A2;\" LT \";\" LATA \")",
        "solucao_powerquery": "Text.Replace([Produto], \" LT \", \" LATA \")",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "SKOL LATA 350ML",
        "erros_comuns": "Trocar LT dentro de outra palavra.",
        "alternativas": "Usar tabela De/Para quando houver muitas abreviações.",
        "aplicacao_pratica": "Padronização de cadastro de produtos.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-004",
        "ferramenta": "Excel",
        "tema": "Datas",
        "nivel": "Básico",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-004 — Datas (Básico)\n\nContexto de negócio: A área financeira precisa agrupar lançamentos por competência.\n\nProblema: Transformar qualquer data no primeiro dia do mês.\n\nBase simulada: A2 = 15/03/2025\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Competência deve ser data real, não texto.",
        "raciocinio": "Usar DATA com ANO e MÊS da data original.",
        "solucao_principal": "=DATA(ANO(A2);MÊS(A2);1)",
        "solucao_excel": "=DATA(ANO(A2);MÊS(A2);1)",
        "solucao_powerquery": "Power Query: Date.StartOfMonth([Data])",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "01/03/2025",
        "erros_comuns": "Usar texto '03/2025' e perder ordenação temporal.",
        "alternativas": "Power Query pode criar coluna de competência no ETL.",
        "aplicacao_pratica": "Fechamento financeiro e contábil.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-005",
        "ferramenta": "Excel",
        "tema": "Condicionais",
        "nivel": "Básico",
        "area_negocio": "Comercial",
        "enunciado": "EXC-005 — Condicionais (Básico)\n\nContexto de negócio: Gestor quer classificar vendas por valor.\n\nProblema: Classificar Alta se >=1000, Média se >=500, Baixa caso contrário.\n\nBase simulada: B2 = 780\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "A regra tem faixas de valor ordenadas.",
        "raciocinio": "Usar SE aninhado começando pela maior faixa.",
        "solucao_principal": "=SE(B2>=1000;\"Alta\";SE(B2>=500;\"Média\";\"Baixa\"))",
        "solucao_excel": "=SE(B2>=1000;\"Alta\";SE(B2>=500;\"Média\";\"Baixa\"))",
        "solucao_powerquery": "if [Valor] >= 1000 then \"Alta\" else if [Valor] >= 500 then \"Média\" else \"Baixa\"",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Média",
        "erros_comuns": "Testar primeiro a faixa menor e classificar tudo errado.",
        "alternativas": "Usar tabela de faixas e PROCX quando faixas mudarem.",
        "aplicacao_pratica": "Segmentação de carteira e priorização.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-006",
        "ferramenta": "Excel",
        "tema": "Buscas",
        "nivel": "Intermediário",
        "area_negocio": "Supply Chain",
        "enunciado": "EXC-006 — Buscas (Intermediário)\n\nContexto de negócio: A base de vendas tem SKU e precisa buscar categoria.\n\nProblema: Retornar categoria do SKU usando PROCX.\n\nBase simulada: A2 = SKU001 | Produtos[SKU] e Produtos[Categoria]\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "A busca deve retornar mensagem quando não houver cadastro.",
        "raciocinio": "PROCX busca SKU na coluna de chaves e retorna categoria.",
        "solucao_principal": "=PROCX(A2;Produtos[SKU];Produtos[Categoria];\"Sem cadastro\")",
        "solucao_excel": "=PROCX(A2;Produtos[SKU];Produtos[Categoria];\"Sem cadastro\")",
        "solucao_powerquery": "Table.NestedJoin(Base, {\"SKU\"}, Produtos, {\"SKU\"}, \"Produto\", JoinKind.LeftOuter)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Categoria correspondente ou Sem cadastro.",
        "erros_comuns": "Não tratar SKU ausente.",
        "alternativas": "PROCV funciona, mas é mais frágil com alteração de colunas.",
        "aplicacao_pratica": "Cadastro de produtos e relatórios de estoque.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-007",
        "ferramenta": "Excel",
        "tema": "Buscas",
        "nivel": "Avançado",
        "area_negocio": "Comercial",
        "enunciado": "EXC-007 — Buscas (Avançado)\n\nContexto de negócio: Preço depende de produto e canal.\n\nProblema: Buscar categoria/preço usando Produto + Canal como critérios.\n\nBase simulada: A2 = Produto | B2 = Canal\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "A chave é composta por duas condições.",
        "raciocinio": "Multiplicar condições gera matriz 1/0 para PROCX.",
        "solucao_principal": "=PROCX(1;(Produtos[Produto]=A2)*(Produtos[Canal]=B2);Produtos[Categoria];\"N/D\")",
        "solucao_excel": "=PROCX(1;(Produtos[Produto]=A2)*(Produtos[Canal]=B2);Produtos[Categoria];\"N/D\")",
        "solucao_powerquery": "Power Query: criar chave composta ou merge por duas colunas.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Categoria do produto naquele canal.",
        "erros_comuns": "Usar apenas produto e ignorar canal.",
        "alternativas": "Criar chave auxiliar Produto&Canal.",
        "aplicacao_pratica": "Tabelas comerciais com preço por canal.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-008",
        "ferramenta": "Excel",
        "tema": "Agregações",
        "nivel": "Intermediário",
        "area_negocio": "Comercial",
        "enunciado": "EXC-008 — Agregações (Intermediário)\n\nContexto de negócio: Diretoria quer total de vendas por produto e canal.\n\nProblema: Somar vendas do produto informado em A2 no canal Online.\n\nBase simulada: Tabela Base com Produto, Canal e Valor.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Há dois critérios: produto e canal.",
        "raciocinio": "SOMASES permite múltiplos critérios.",
        "solucao_principal": "=SOMASES(Base[Valor];Base[Produto];A2;Base[Canal];\"Online\")",
        "solucao_excel": "=SOMASES(Base[Valor];Base[Produto];A2;Base[Canal];\"Online\")",
        "solucao_powerquery": "Table.Group(Table.SelectRows(Fonte, each [Canal] = \"Online\"), {\"Produto\"}, {{\"Total\", each List.Sum([Valor]), type number}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Total de vendas filtrado.",
        "erros_comuns": "Inverter intervalo de soma com intervalo de critério.",
        "alternativas": "Tabela dinâmica ou DAX quando modelo cresce.",
        "aplicacao_pratica": "Relatórios comerciais e metas por canal.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-009",
        "ferramenta": "Excel",
        "tema": "Agregações",
        "nivel": "Intermediário",
        "area_negocio": "Operações",
        "enunciado": "EXC-009 — Agregações (Intermediário)\n\nContexto de negócio: Gestão precisa contar registros ativos por produto.\n\nProblema: Contar linhas do produto A2 com Status = Ativo.\n\nBase simulada: Base[Produto], Base[Status]\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Contagem exige múltiplos critérios.",
        "raciocinio": "CONT.SES conta registros que atendem a ambos.",
        "solucao_principal": "=CONT.SES(Base[Produto];A2;Base[Status];\"Ativo\")",
        "solucao_excel": "=CONT.SES(Base[Produto];A2;Base[Status];\"Ativo\")",
        "solucao_powerquery": "Table.RowCount(Table.SelectRows(Fonte, each [Produto] = ProdutoParametro and [Status] = \"Ativo\"))",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Quantidade de registros ativos.",
        "erros_comuns": "Usar CONT.SE e esquecer status.",
        "alternativas": "DAX é melhor para dashboards.",
        "aplicacao_pratica": "Controle operacional e cadastro.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-010",
        "ferramenta": "Excel",
        "tema": "Matrizes Dinâmicas",
        "nivel": "Avançado",
        "area_negocio": "Comercial",
        "enunciado": "EXC-010 — Matrizes Dinâmicas (Avançado)\n\nContexto de negócio: Analista precisa listar produtos com venda acima de 1000.\n\nProblema: Filtrar tabela Base para Valor > 1000.\n\nBase simulada: Base[Valor]\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Função FILTRO retorna matriz dinâmica.",
        "raciocinio": "Critério deve ter mesmo tamanho das linhas da base.",
        "solucao_principal": "=FILTRO(Base;Base[Valor]>1000;\"Sem registros\")",
        "solucao_excel": "=FILTRO(Base;Base[Valor]>1000;\"Sem registros\")",
        "solucao_powerquery": "Table.SelectRows(Fonte, each [Valor] > 1000)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Linhas com Valor > 1000.",
        "erros_comuns": "Usar intervalo de critério com tamanho diferente.",
        "alternativas": "Power Query é melhor para filtro recorrente.",
        "aplicacao_pratica": "Análise de vendas relevantes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-EXPERT-001",
        "ferramenta": "Excel",
        "tema": "Modelagem",
        "nivel": "Expert",
        "area_negocio": "Comercial",
        "enunciado": "EXC-EXPERT-001 — Modelagem (Expert)\n\nContexto de negócio: A empresa separou FatoVendas e DimProduto para organizar melhor o relatório.\n\nProblema: Trazer Categoria da DimProduto para a FatoVendas usando SKU.\n\nBase simulada: FatoVendas[SKU] e DimProduto[SKU]/DimProduto[Categoria]\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "O problema exige conceito de tabela fato e dimensão.",
        "raciocinio": "Usar PROCX com tabela dimensão e tratar ausência de cadastro.",
        "solucao_principal": "=PROCX([@SKU];DimProduto[SKU];DimProduto[Categoria];\"Sem cadastro\")",
        "solucao_excel": "=PROCX([@SKU];DimProduto[SKU];DimProduto[Categoria];\"Sem cadastro\")",
        "solucao_powerquery": "Table.NestedJoin(FatoVendas, {\"SKU\"}, DimProduto, {\"SKU\"}, \"DimProduto\", JoinKind.LeftOuter)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Categoria preenchida ou Sem cadastro.",
        "erros_comuns": "Copiar categoria manualmente para a fato e perder rastreabilidade.",
        "alternativas": "Power Pivot com relacionamento é alternativa mais escalável.",
        "aplicacao_pratica": "Modelagem analítica em Excel/Power Pivot.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-EXPERT-002",
        "ferramenta": "Excel",
        "tema": "Performance",
        "nivel": "Expert",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-EXPERT-002 — Performance (Expert)\n\nContexto de negócio: A planilha ficou lenta após fórmulas em colunas inteiras.\n\nProblema: Reescrever soma condicional usando tabela estruturada.\n\nBase simulada: Tabela Base com colunas Produto e Valor.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Referências estruturadas reduzem cálculo desnecessário.",
        "raciocinio": "Usar SOMASES com Base[Valor] e Base[Produto].",
        "solucao_principal": "=SOMASES(Base[Valor];Base[Produto];A2)",
        "solucao_excel": "=SOMASES(Base[Valor];Base[Produto];A2)",
        "solucao_powerquery": "Power Query pode pré-agregar a base.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Soma calculada com melhor organização.",
        "erros_comuns": "Usar B:B e A:A em milhares de fórmulas voláteis.",
        "alternativas": "Pré-agregar no Power Query ou Power Pivot.",
        "aplicacao_pratica": "Otimização de relatórios financeiros.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-EXPERT-003",
        "ferramenta": "Excel",
        "tema": "Auditoria",
        "nivel": "Expert",
        "area_negocio": "Governança",
        "enunciado": "EXC-EXPERT-003 — Auditoria (Expert)\n\nContexto de negócio: Auditoria quer rastrear registros incompletos antes da classificação.\n\nProblema: Criar fórmula que sinaliza incompleto antes de aplicar regra de prioridade.\n\nBase simulada: B2 = Valor | C2 = Status\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "A fórmula deve separar validação de regra de negócio.",
        "raciocinio": "LET nomeia valor e status, depois aplica validação e regra.",
        "solucao_principal": "=LET(v;B2;status;C2;SE(OU(v=\"\";status=\"\");\"INCOMPLETO\";SE(E(v>=1000;status=\"Ativo\");\"PRIORITÁRIO\";\"NORMAL\")))",
        "solucao_excel": "=LET(v;B2;status;C2;SE(OU(v=\"\";status=\"\");\"INCOMPLETO\";SE(E(v>=1000;status=\"Ativo\");\"PRIORITÁRIO\";\"NORMAL\")))",
        "solucao_powerquery": "Power Query: criar Status_DQ antes da regra.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "INCOMPLETO, PRIORITÁRIO ou NORMAL.",
        "erros_comuns": "Classificar sem validar incompletos.",
        "alternativas": "Separar validação em coluna auxiliar.",
        "aplicacao_pratica": "Governança de regras de negócio.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-EXPERT-004",
        "ferramenta": "Excel",
        "tema": "Matrizes Dinâmicas",
        "nivel": "Expert",
        "area_negocio": "Diretoria",
        "enunciado": "EXC-EXPERT-004 — Matrizes Dinâmicas (Expert)\n\nContexto de negócio: Diretoria precisa de visão compacta dos 10 maiores produtos.\n\nProblema: Ordenar base por valor, pegar top 10 e exibir apenas colunas relevantes.\n\nBase simulada: Base com Produto, Canal, Valor e demais colunas.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Combinar CLASSIFICAR, PEGAR e ESCOLHERCOLS.",
        "raciocinio": "Ordenar desc pela coluna de valor, limitar top 10 e selecionar colunas.",
        "solucao_principal": "=ESCOLHERCOLS(PEGAR(CLASSIFICAR(Base;3;-1);10);1;2;3)",
        "solucao_excel": "=ESCOLHERCOLS(PEGAR(CLASSIFICAR(Base;3;-1);10);1;2;3)",
        "solucao_powerquery": "Power Query: Table.Sort + Table.FirstN + Table.SelectColumns.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Visão executiva com top 10.",
        "erros_comuns": "Ordenar pela coluna errada ou retornar colunas demais.",
        "alternativas": "Tabela dinâmica com filtro Top 10.",
        "aplicacao_pratica": "Painel gerencial.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-011",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Básico",
        "area_negocio": "RH",
        "enunciado": "EXC-011 — Texto (Básico)\n\nContexto de negócio: Situação real de RH: padronizar cpf sem pontuação.\n\nProblema: Padronizar CPF sem pontuação\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SUBSTITUIR(SUBSTITUIR(SUBSTITUIR(A2;\".\";\"\");\"-\";\"\");\"/\";\"\")",
        "solucao_excel": "=SUBSTITUIR(SUBSTITUIR(SUBSTITUIR(A2;\".\";\"\");\"-\";\"\");\"/\";\"\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de RH e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-012",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Básico",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-012 — Texto (Básico)\n\nContexto de negócio: Situação real de Financeiro: extrair centro de custo de código.\n\nProblema: Extrair centro de custo de código\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=EXT.TEXTO(A2;4;3)",
        "solucao_excel": "=EXT.TEXTO(A2;4;3)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-013",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Intermediário",
        "area_negocio": "Marketing",
        "enunciado": "EXC-013 — Texto (Intermediário)\n\nContexto de negócio: Situação real de Marketing: montar chave de campanha.\n\nProblema: Montar chave de campanha\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=TEXTOJUNTAR(\"-\";VERDADEIRO;A2:C2)",
        "solucao_excel": "=TEXTOJUNTAR(\"-\";VERDADEIRO;A2:C2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Marketing e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-014",
        "ferramenta": "Excel",
        "tema": "Datas",
        "nivel": "Básico",
        "area_negocio": "Logística",
        "enunciado": "EXC-014 — Datas (Básico)\n\nContexto de negócio: Situação real de Logística: calcular sla em dias.\n\nProblema: Calcular SLA em dias\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=DIAS(B2;A2)",
        "solucao_excel": "=DIAS(B2;A2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Logística e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-015",
        "ferramenta": "Excel",
        "tema": "Datas",
        "nivel": "Intermediário",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-015 — Datas (Intermediário)\n\nContexto de negócio: Situação real de Financeiro: identificar vencido.\n\nProblema: Identificar vencido\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(HOJE()>A2;\"Vencido\";\"No prazo\")",
        "solucao_excel": "=SE(HOJE()>A2;\"Vencido\";\"No prazo\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-016",
        "ferramenta": "Excel",
        "tema": "Condicionais",
        "nivel": "Intermediário",
        "area_negocio": "RH",
        "enunciado": "EXC-016 — Condicionais (Intermediário)\n\nContexto de negócio: Situação real de RH: classificar absenteísmo.\n\nProblema: Classificar absenteísmo\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(B2>0,05;\"Crítico\";\"Normal\")",
        "solucao_excel": "=SE(B2>0,05;\"Crítico\";\"Normal\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de RH e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-017",
        "ferramenta": "Excel",
        "tema": "Condicionais",
        "nivel": "Intermediário",
        "area_negocio": "Supply Chain",
        "enunciado": "EXC-017 — Condicionais (Intermediário)\n\nContexto de negócio: Situação real de Supply Chain: status de ruptura.\n\nProblema: Status de ruptura\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(B2=0;\"Ruptura\";\"Com estoque\")",
        "solucao_excel": "=SE(B2=0;\"Ruptura\";\"Com estoque\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Supply Chain e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-018",
        "ferramenta": "Excel",
        "tema": "Buscas",
        "nivel": "Intermediário",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-018 — Buscas (Intermediário)\n\nContexto de negócio: Situação real de Financeiro: buscar conta contábil.\n\nProblema: Buscar conta contábil\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=PROCX(A2;Plano[Codigo];Plano[Conta];\"Não encontrada\")",
        "solucao_excel": "=PROCX(A2;Plano[Codigo];Plano[Conta];\"Não encontrada\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-019",
        "ferramenta": "Excel",
        "tema": "Buscas",
        "nivel": "Avançado",
        "area_negocio": "Comercial",
        "enunciado": "EXC-019 — Buscas (Avançado)\n\nContexto de negócio: Situação real de Comercial: busca parcial com procx.\n\nProblema: Busca parcial com PROCX\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=PROCX(\"*\"&A2&\"*\";Base[Produto];Base[Categoria];\"N/D\";2)",
        "solucao_excel": "=PROCX(\"*\"&A2&\"*\";Base[Produto];Base[Categoria];\"N/D\";2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-020",
        "ferramenta": "Excel",
        "tema": "Agregações",
        "nivel": "Intermediário",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-020 — Agregações (Intermediário)\n\nContexto de negócio: Situação real de Financeiro: somar despesas por centro.\n\nProblema: Somar despesas por centro\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SOMASES(Base[Valor];Base[Centro];A2;Base[Tipo];\"Despesa\")",
        "solucao_excel": "=SOMASES(Base[Valor];Base[Centro];A2;Base[Tipo];\"Despesa\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-021",
        "ferramenta": "Excel",
        "tema": "Agregações",
        "nivel": "Intermediário",
        "area_negocio": "Marketing",
        "enunciado": "EXC-021 — Agregações (Intermediário)\n\nContexto de negócio: Situação real de Marketing: média de roi por canal.\n\nProblema: Média de ROI por canal\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MÉDIASES(Base[ROI];Base[Canal];A2)",
        "solucao_excel": "=MÉDIASES(Base[ROI];Base[Canal];A2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Marketing e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-022",
        "ferramenta": "Excel",
        "tema": "Agregações",
        "nivel": "Intermediário",
        "area_negocio": "RH",
        "enunciado": "EXC-022 — Agregações (Intermediário)\n\nContexto de negócio: Situação real de RH: contar colaboradores ativos.\n\nProblema: Contar colaboradores ativos\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CONT.SES(Base[Area];A2;Base[Status];\"Ativo\")",
        "solucao_excel": "=CONT.SES(Base[Area];A2;Base[Status];\"Ativo\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de RH e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-023",
        "ferramenta": "Excel",
        "tema": "Coringas",
        "nivel": "Avançado",
        "area_negocio": "Comercial",
        "enunciado": "EXC-023 — Coringas (Avançado)\n\nContexto de negócio: Situação real de Comercial: encontrar produtos lata.\n\nProblema: Encontrar produtos lata\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CONT.SE(A:A;\"*LATA*\")",
        "solucao_excel": "=CONT.SE(A:A;\"*LATA*\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-024",
        "ferramenta": "Excel",
        "tema": "Coringas",
        "nivel": "Avançado",
        "area_negocio": "Supply Chain",
        "enunciado": "EXC-024 — Coringas (Avançado)\n\nContexto de negócio: Situação real de Supply Chain: validar código 3 letras.\n\nProblema: Validar código 3 letras\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(CONT.SE(A2;\"???\")>0;\"OK\";\"Validar\")",
        "solucao_excel": "=SE(CONT.SE(A2;\"???\")>0;\"OK\";\"Validar\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Supply Chain e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-025",
        "ferramenta": "Excel",
        "tema": "Coringas",
        "nivel": "Avançado",
        "area_negocio": "TI",
        "enunciado": "EXC-025 — Coringas (Avançado)\n\nContexto de negócio: Situação real de TI: buscar asterisco literal.\n\nProblema: Buscar asterisco literal\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CONT.SE(A:A;\"SKU~*\")",
        "solucao_excel": "=CONT.SE(A:A;\"SKU~*\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de TI e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-026",
        "ferramenta": "Excel",
        "tema": "Matrizes Dinâmicas",
        "nivel": "Intermediário",
        "area_negocio": "Comercial",
        "enunciado": "EXC-026 — Matrizes Dinâmicas (Intermediário)\n\nContexto de negócio: Situação real de Comercial: lista única de vendedores.\n\nProblema: Lista única de vendedores\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CLASSIFICAR(ÚNICO(Base[Vendedor]))",
        "solucao_excel": "=CLASSIFICAR(ÚNICO(Base[Vendedor]))",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-027",
        "ferramenta": "Excel",
        "tema": "Matrizes Dinâmicas",
        "nivel": "Avançado",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-027 — Matrizes Dinâmicas (Avançado)\n\nContexto de negócio: Situação real de Financeiro: top 10 despesas.\n\nProblema: Top 10 despesas\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=PEGAR(CLASSIFICAR(Base;3;-1);10)",
        "solucao_excel": "=PEGAR(CLASSIFICAR(Base;3;-1);10)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-028",
        "ferramenta": "Excel",
        "tema": "Matrizes Dinâmicas",
        "nivel": "Avançado",
        "area_negocio": "Operações",
        "enunciado": "EXC-028 — Matrizes Dinâmicas (Avançado)\n\nContexto de negócio: Situação real de Operações: selecionar colunas principais.\n\nProblema: Selecionar colunas principais\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=ESCOLHERCOLS(Base;1;2;5)",
        "solucao_excel": "=ESCOLHERCOLS(Base;1;2;5)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Operações e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-029",
        "ferramenta": "Excel",
        "tema": "Auditoria",
        "nivel": "Avançado",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-029 — Auditoria (Avançado)\n\nContexto de negócio: Situação real de Financeiro: regra com let para limite.\n\nProblema: Regra com LET para limite\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=LET(v;B2;limite;10000;SE(v>limite;\"Aprovar\";\"OK\"))",
        "solucao_excel": "=LET(v;B2;limite;10000;SE(v>limite;\"Aprovar\";\"OK\"))",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-030",
        "ferramenta": "Excel",
        "tema": "Auditoria",
        "nivel": "Avançado",
        "area_negocio": "Dados",
        "enunciado": "EXC-030 — Auditoria (Avançado)\n\nContexto de negócio: Situação real de Dados: limpar lista com map.\n\nProblema: Limpar lista com MAP\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MAP(A2:A100;LAMBDA(x;MAIÚSCULA(ARRUMAR(x))))",
        "solucao_excel": "=MAP(A2:A100;LAMBDA(x;MAIÚSCULA(ARRUMAR(x))))",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Dados e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-031",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Intermediário",
        "area_negocio": "Comercial",
        "enunciado": "EXC-031 — Estatística no Excel (Intermediário)\n\nContexto de negócio: Situação real de Comercial: média de vendas.\n\nProblema: Média de vendas\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MÉDIA(Base[Valor])",
        "solucao_excel": "=MÉDIA(Base[Valor])",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-032",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Intermediário",
        "area_negocio": "Comercial",
        "enunciado": "EXC-032 — Estatística no Excel (Intermediário)\n\nContexto de negócio: Situação real de Comercial: mediana de vendas.\n\nProblema: Mediana de vendas\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MED(Base[Valor])",
        "solucao_excel": "=MED(Base[Valor])",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-033",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Intermediário",
        "area_negocio": "Operações",
        "enunciado": "EXC-033 — Estatística no Excel (Intermediário)\n\nContexto de negócio: Situação real de Operações: desvio padrão de sla.\n\nProblema: Desvio padrão de SLA\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=DESVPAD.S(Base[SLA])",
        "solucao_excel": "=DESVPAD.S(Base[SLA])",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Operações e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-034",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Avançado",
        "area_negocio": "Marketing",
        "enunciado": "EXC-034 — Estatística no Excel (Avançado)\n\nContexto de negócio: Situação real de Marketing: correlação investimento vendas.\n\nProblema: Correlação investimento vendas\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CORREL(Base[Investimento];Base[Vendas])",
        "solucao_excel": "=CORREL(Base[Investimento];Base[Vendas])",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Marketing e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-035",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Avançado",
        "area_negocio": "Comercial",
        "enunciado": "EXC-035 — Estatística no Excel (Avançado)\n\nContexto de negócio: Situação real de Comercial: projeção linear.\n\nProblema: Projeção linear\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=TENDÊNCIA(Base[Vendas];Base[Mes];D2:D6)",
        "solucao_excel": "=TENDÊNCIA(Base[Vendas];Base[Mes];D2:D6)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-036",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Básico",
        "area_negocio": "RH",
        "enunciado": "EXC-036 — Texto (Básico)\n\nContexto de negócio: Situação real de RH: padronizar cpf sem pontuação.\n\nProblema: Padronizar CPF sem pontuação\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SUBSTITUIR(SUBSTITUIR(SUBSTITUIR(A2;\".\";\"\");\"-\";\"\");\"/\";\"\")",
        "solucao_excel": "=SUBSTITUIR(SUBSTITUIR(SUBSTITUIR(A2;\".\";\"\");\"-\";\"\");\"/\";\"\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de RH e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-037",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Básico",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-037 — Texto (Básico)\n\nContexto de negócio: Situação real de Financeiro: extrair centro de custo de código.\n\nProblema: Extrair centro de custo de código\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=EXT.TEXTO(A2;4;3)",
        "solucao_excel": "=EXT.TEXTO(A2;4;3)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-038",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Intermediário",
        "area_negocio": "Marketing",
        "enunciado": "EXC-038 — Texto (Intermediário)\n\nContexto de negócio: Situação real de Marketing: montar chave de campanha.\n\nProblema: Montar chave de campanha\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=TEXTOJUNTAR(\"-\";VERDADEIRO;A2:C2)",
        "solucao_excel": "=TEXTOJUNTAR(\"-\";VERDADEIRO;A2:C2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Marketing e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-039",
        "ferramenta": "Excel",
        "tema": "Datas",
        "nivel": "Básico",
        "area_negocio": "Logística",
        "enunciado": "EXC-039 — Datas (Básico)\n\nContexto de negócio: Situação real de Logística: calcular sla em dias.\n\nProblema: Calcular SLA em dias\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=DIAS(B2;A2)",
        "solucao_excel": "=DIAS(B2;A2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Logística e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-040",
        "ferramenta": "Excel",
        "tema": "Datas",
        "nivel": "Intermediário",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-040 — Datas (Intermediário)\n\nContexto de negócio: Situação real de Financeiro: identificar vencido.\n\nProblema: Identificar vencido\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(HOJE()>A2;\"Vencido\";\"No prazo\")",
        "solucao_excel": "=SE(HOJE()>A2;\"Vencido\";\"No prazo\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-041",
        "ferramenta": "Excel",
        "tema": "Condicionais",
        "nivel": "Intermediário",
        "area_negocio": "RH",
        "enunciado": "EXC-041 — Condicionais (Intermediário)\n\nContexto de negócio: Situação real de RH: classificar absenteísmo.\n\nProblema: Classificar absenteísmo\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(B2>0,05;\"Crítico\";\"Normal\")",
        "solucao_excel": "=SE(B2>0,05;\"Crítico\";\"Normal\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de RH e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-042",
        "ferramenta": "Excel",
        "tema": "Condicionais",
        "nivel": "Intermediário",
        "area_negocio": "Supply Chain",
        "enunciado": "EXC-042 — Condicionais (Intermediário)\n\nContexto de negócio: Situação real de Supply Chain: status de ruptura.\n\nProblema: Status de ruptura\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(B2=0;\"Ruptura\";\"Com estoque\")",
        "solucao_excel": "=SE(B2=0;\"Ruptura\";\"Com estoque\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Supply Chain e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-043",
        "ferramenta": "Excel",
        "tema": "Buscas",
        "nivel": "Intermediário",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-043 — Buscas (Intermediário)\n\nContexto de negócio: Situação real de Financeiro: buscar conta contábil.\n\nProblema: Buscar conta contábil\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=PROCX(A2;Plano[Codigo];Plano[Conta];\"Não encontrada\")",
        "solucao_excel": "=PROCX(A2;Plano[Codigo];Plano[Conta];\"Não encontrada\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-044",
        "ferramenta": "Excel",
        "tema": "Buscas",
        "nivel": "Avançado",
        "area_negocio": "Comercial",
        "enunciado": "EXC-044 — Buscas (Avançado)\n\nContexto de negócio: Situação real de Comercial: busca parcial com procx.\n\nProblema: Busca parcial com PROCX\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=PROCX(\"*\"&A2&\"*\";Base[Produto];Base[Categoria];\"N/D\";2)",
        "solucao_excel": "=PROCX(\"*\"&A2&\"*\";Base[Produto];Base[Categoria];\"N/D\";2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-045",
        "ferramenta": "Excel",
        "tema": "Agregações",
        "nivel": "Intermediário",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-045 — Agregações (Intermediário)\n\nContexto de negócio: Situação real de Financeiro: somar despesas por centro.\n\nProblema: Somar despesas por centro\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SOMASES(Base[Valor];Base[Centro];A2;Base[Tipo];\"Despesa\")",
        "solucao_excel": "=SOMASES(Base[Valor];Base[Centro];A2;Base[Tipo];\"Despesa\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-046",
        "ferramenta": "Excel",
        "tema": "Agregações",
        "nivel": "Intermediário",
        "area_negocio": "Marketing",
        "enunciado": "EXC-046 — Agregações (Intermediário)\n\nContexto de negócio: Situação real de Marketing: média de roi por canal.\n\nProblema: Média de ROI por canal\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MÉDIASES(Base[ROI];Base[Canal];A2)",
        "solucao_excel": "=MÉDIASES(Base[ROI];Base[Canal];A2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Marketing e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-047",
        "ferramenta": "Excel",
        "tema": "Agregações",
        "nivel": "Intermediário",
        "area_negocio": "RH",
        "enunciado": "EXC-047 — Agregações (Intermediário)\n\nContexto de negócio: Situação real de RH: contar colaboradores ativos.\n\nProblema: Contar colaboradores ativos\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CONT.SES(Base[Area];A2;Base[Status];\"Ativo\")",
        "solucao_excel": "=CONT.SES(Base[Area];A2;Base[Status];\"Ativo\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de RH e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-048",
        "ferramenta": "Excel",
        "tema": "Coringas",
        "nivel": "Avançado",
        "area_negocio": "Comercial",
        "enunciado": "EXC-048 — Coringas (Avançado)\n\nContexto de negócio: Situação real de Comercial: encontrar produtos lata.\n\nProblema: Encontrar produtos lata\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CONT.SE(A:A;\"*LATA*\")",
        "solucao_excel": "=CONT.SE(A:A;\"*LATA*\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-049",
        "ferramenta": "Excel",
        "tema": "Coringas",
        "nivel": "Avançado",
        "area_negocio": "Supply Chain",
        "enunciado": "EXC-049 — Coringas (Avançado)\n\nContexto de negócio: Situação real de Supply Chain: validar código 3 letras.\n\nProblema: Validar código 3 letras\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(CONT.SE(A2;\"???\")>0;\"OK\";\"Validar\")",
        "solucao_excel": "=SE(CONT.SE(A2;\"???\")>0;\"OK\";\"Validar\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Supply Chain e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-050",
        "ferramenta": "Excel",
        "tema": "Coringas",
        "nivel": "Avançado",
        "area_negocio": "TI",
        "enunciado": "EXC-050 — Coringas (Avançado)\n\nContexto de negócio: Situação real de TI: buscar asterisco literal.\n\nProblema: Buscar asterisco literal\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CONT.SE(A:A;\"SKU~*\")",
        "solucao_excel": "=CONT.SE(A:A;\"SKU~*\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de TI e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-051",
        "ferramenta": "Excel",
        "tema": "Matrizes Dinâmicas",
        "nivel": "Intermediário",
        "area_negocio": "Comercial",
        "enunciado": "EXC-051 — Matrizes Dinâmicas (Intermediário)\n\nContexto de negócio: Situação real de Comercial: lista única de vendedores.\n\nProblema: Lista única de vendedores\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CLASSIFICAR(ÚNICO(Base[Vendedor]))",
        "solucao_excel": "=CLASSIFICAR(ÚNICO(Base[Vendedor]))",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-052",
        "ferramenta": "Excel",
        "tema": "Matrizes Dinâmicas",
        "nivel": "Avançado",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-052 — Matrizes Dinâmicas (Avançado)\n\nContexto de negócio: Situação real de Financeiro: top 10 despesas.\n\nProblema: Top 10 despesas\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=PEGAR(CLASSIFICAR(Base;3;-1);10)",
        "solucao_excel": "=PEGAR(CLASSIFICAR(Base;3;-1);10)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-053",
        "ferramenta": "Excel",
        "tema": "Matrizes Dinâmicas",
        "nivel": "Avançado",
        "area_negocio": "Operações",
        "enunciado": "EXC-053 — Matrizes Dinâmicas (Avançado)\n\nContexto de negócio: Situação real de Operações: selecionar colunas principais.\n\nProblema: Selecionar colunas principais\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=ESCOLHERCOLS(Base;1;2;5)",
        "solucao_excel": "=ESCOLHERCOLS(Base;1;2;5)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Operações e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-054",
        "ferramenta": "Excel",
        "tema": "Auditoria",
        "nivel": "Avançado",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-054 — Auditoria (Avançado)\n\nContexto de negócio: Situação real de Financeiro: regra com let para limite.\n\nProblema: Regra com LET para limite\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=LET(v;B2;limite;10000;SE(v>limite;\"Aprovar\";\"OK\"))",
        "solucao_excel": "=LET(v;B2;limite;10000;SE(v>limite;\"Aprovar\";\"OK\"))",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-055",
        "ferramenta": "Excel",
        "tema": "Auditoria",
        "nivel": "Avançado",
        "area_negocio": "Dados",
        "enunciado": "EXC-055 — Auditoria (Avançado)\n\nContexto de negócio: Situação real de Dados: limpar lista com map.\n\nProblema: Limpar lista com MAP\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MAP(A2:A100;LAMBDA(x;MAIÚSCULA(ARRUMAR(x))))",
        "solucao_excel": "=MAP(A2:A100;LAMBDA(x;MAIÚSCULA(ARRUMAR(x))))",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Dados e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-056",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Intermediário",
        "area_negocio": "Comercial",
        "enunciado": "EXC-056 — Estatística no Excel (Intermediário)\n\nContexto de negócio: Situação real de Comercial: média de vendas.\n\nProblema: Média de vendas\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MÉDIA(Base[Valor])",
        "solucao_excel": "=MÉDIA(Base[Valor])",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-057",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Intermediário",
        "area_negocio": "Comercial",
        "enunciado": "EXC-057 — Estatística no Excel (Intermediário)\n\nContexto de negócio: Situação real de Comercial: mediana de vendas.\n\nProblema: Mediana de vendas\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MED(Base[Valor])",
        "solucao_excel": "=MED(Base[Valor])",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-058",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Intermediário",
        "area_negocio": "Operações",
        "enunciado": "EXC-058 — Estatística no Excel (Intermediário)\n\nContexto de negócio: Situação real de Operações: desvio padrão de sla.\n\nProblema: Desvio padrão de SLA\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=DESVPAD.S(Base[SLA])",
        "solucao_excel": "=DESVPAD.S(Base[SLA])",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Operações e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-059",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Avançado",
        "area_negocio": "Marketing",
        "enunciado": "EXC-059 — Estatística no Excel (Avançado)\n\nContexto de negócio: Situação real de Marketing: correlação investimento vendas.\n\nProblema: Correlação investimento vendas\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CORREL(Base[Investimento];Base[Vendas])",
        "solucao_excel": "=CORREL(Base[Investimento];Base[Vendas])",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Marketing e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-060",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Avançado",
        "area_negocio": "Comercial",
        "enunciado": "EXC-060 — Estatística no Excel (Avançado)\n\nContexto de negócio: Situação real de Comercial: projeção linear.\n\nProblema: Projeção linear\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=TENDÊNCIA(Base[Vendas];Base[Mes];D2:D6)",
        "solucao_excel": "=TENDÊNCIA(Base[Vendas];Base[Mes];D2:D6)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-061",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Básico",
        "area_negocio": "RH",
        "enunciado": "EXC-061 — Texto (Básico)\n\nContexto de negócio: Situação real de RH: padronizar cpf sem pontuação.\n\nProblema: Padronizar CPF sem pontuação\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SUBSTITUIR(SUBSTITUIR(SUBSTITUIR(A2;\".\";\"\");\"-\";\"\");\"/\";\"\")",
        "solucao_excel": "=SUBSTITUIR(SUBSTITUIR(SUBSTITUIR(A2;\".\";\"\");\"-\";\"\");\"/\";\"\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de RH e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-062",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Básico",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-062 — Texto (Básico)\n\nContexto de negócio: Situação real de Financeiro: extrair centro de custo de código.\n\nProblema: Extrair centro de custo de código\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=EXT.TEXTO(A2;4;3)",
        "solucao_excel": "=EXT.TEXTO(A2;4;3)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-063",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Intermediário",
        "area_negocio": "Marketing",
        "enunciado": "EXC-063 — Texto (Intermediário)\n\nContexto de negócio: Situação real de Marketing: montar chave de campanha.\n\nProblema: Montar chave de campanha\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=TEXTOJUNTAR(\"-\";VERDADEIRO;A2:C2)",
        "solucao_excel": "=TEXTOJUNTAR(\"-\";VERDADEIRO;A2:C2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Marketing e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-064",
        "ferramenta": "Excel",
        "tema": "Datas",
        "nivel": "Básico",
        "area_negocio": "Logística",
        "enunciado": "EXC-064 — Datas (Básico)\n\nContexto de negócio: Situação real de Logística: calcular sla em dias.\n\nProblema: Calcular SLA em dias\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=DIAS(B2;A2)",
        "solucao_excel": "=DIAS(B2;A2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Logística e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-065",
        "ferramenta": "Excel",
        "tema": "Datas",
        "nivel": "Intermediário",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-065 — Datas (Intermediário)\n\nContexto de negócio: Situação real de Financeiro: identificar vencido.\n\nProblema: Identificar vencido\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(HOJE()>A2;\"Vencido\";\"No prazo\")",
        "solucao_excel": "=SE(HOJE()>A2;\"Vencido\";\"No prazo\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-066",
        "ferramenta": "Excel",
        "tema": "Condicionais",
        "nivel": "Intermediário",
        "area_negocio": "RH",
        "enunciado": "EXC-066 — Condicionais (Intermediário)\n\nContexto de negócio: Situação real de RH: classificar absenteísmo.\n\nProblema: Classificar absenteísmo\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(B2>0,05;\"Crítico\";\"Normal\")",
        "solucao_excel": "=SE(B2>0,05;\"Crítico\";\"Normal\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de RH e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-067",
        "ferramenta": "Excel",
        "tema": "Condicionais",
        "nivel": "Intermediário",
        "area_negocio": "Supply Chain",
        "enunciado": "EXC-067 — Condicionais (Intermediário)\n\nContexto de negócio: Situação real de Supply Chain: status de ruptura.\n\nProblema: Status de ruptura\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(B2=0;\"Ruptura\";\"Com estoque\")",
        "solucao_excel": "=SE(B2=0;\"Ruptura\";\"Com estoque\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Supply Chain e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-068",
        "ferramenta": "Excel",
        "tema": "Buscas",
        "nivel": "Intermediário",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-068 — Buscas (Intermediário)\n\nContexto de negócio: Situação real de Financeiro: buscar conta contábil.\n\nProblema: Buscar conta contábil\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=PROCX(A2;Plano[Codigo];Plano[Conta];\"Não encontrada\")",
        "solucao_excel": "=PROCX(A2;Plano[Codigo];Plano[Conta];\"Não encontrada\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-069",
        "ferramenta": "Excel",
        "tema": "Buscas",
        "nivel": "Avançado",
        "area_negocio": "Comercial",
        "enunciado": "EXC-069 — Buscas (Avançado)\n\nContexto de negócio: Situação real de Comercial: busca parcial com procx.\n\nProblema: Busca parcial com PROCX\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=PROCX(\"*\"&A2&\"*\";Base[Produto];Base[Categoria];\"N/D\";2)",
        "solucao_excel": "=PROCX(\"*\"&A2&\"*\";Base[Produto];Base[Categoria];\"N/D\";2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-070",
        "ferramenta": "Excel",
        "tema": "Agregações",
        "nivel": "Intermediário",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-070 — Agregações (Intermediário)\n\nContexto de negócio: Situação real de Financeiro: somar despesas por centro.\n\nProblema: Somar despesas por centro\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SOMASES(Base[Valor];Base[Centro];A2;Base[Tipo];\"Despesa\")",
        "solucao_excel": "=SOMASES(Base[Valor];Base[Centro];A2;Base[Tipo];\"Despesa\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-071",
        "ferramenta": "Excel",
        "tema": "Agregações",
        "nivel": "Intermediário",
        "area_negocio": "Marketing",
        "enunciado": "EXC-071 — Agregações (Intermediário)\n\nContexto de negócio: Situação real de Marketing: média de roi por canal.\n\nProblema: Média de ROI por canal\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MÉDIASES(Base[ROI];Base[Canal];A2)",
        "solucao_excel": "=MÉDIASES(Base[ROI];Base[Canal];A2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Marketing e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-072",
        "ferramenta": "Excel",
        "tema": "Agregações",
        "nivel": "Intermediário",
        "area_negocio": "RH",
        "enunciado": "EXC-072 — Agregações (Intermediário)\n\nContexto de negócio: Situação real de RH: contar colaboradores ativos.\n\nProblema: Contar colaboradores ativos\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CONT.SES(Base[Area];A2;Base[Status];\"Ativo\")",
        "solucao_excel": "=CONT.SES(Base[Area];A2;Base[Status];\"Ativo\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de RH e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-073",
        "ferramenta": "Excel",
        "tema": "Coringas",
        "nivel": "Avançado",
        "area_negocio": "Comercial",
        "enunciado": "EXC-073 — Coringas (Avançado)\n\nContexto de negócio: Situação real de Comercial: encontrar produtos lata.\n\nProblema: Encontrar produtos lata\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CONT.SE(A:A;\"*LATA*\")",
        "solucao_excel": "=CONT.SE(A:A;\"*LATA*\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-074",
        "ferramenta": "Excel",
        "tema": "Coringas",
        "nivel": "Avançado",
        "area_negocio": "Supply Chain",
        "enunciado": "EXC-074 — Coringas (Avançado)\n\nContexto de negócio: Situação real de Supply Chain: validar código 3 letras.\n\nProblema: Validar código 3 letras\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(CONT.SE(A2;\"???\")>0;\"OK\";\"Validar\")",
        "solucao_excel": "=SE(CONT.SE(A2;\"???\")>0;\"OK\";\"Validar\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Supply Chain e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-075",
        "ferramenta": "Excel",
        "tema": "Coringas",
        "nivel": "Avançado",
        "area_negocio": "TI",
        "enunciado": "EXC-075 — Coringas (Avançado)\n\nContexto de negócio: Situação real de TI: buscar asterisco literal.\n\nProblema: Buscar asterisco literal\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CONT.SE(A:A;\"SKU~*\")",
        "solucao_excel": "=CONT.SE(A:A;\"SKU~*\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de TI e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-076",
        "ferramenta": "Excel",
        "tema": "Matrizes Dinâmicas",
        "nivel": "Intermediário",
        "area_negocio": "Comercial",
        "enunciado": "EXC-076 — Matrizes Dinâmicas (Intermediário)\n\nContexto de negócio: Situação real de Comercial: lista única de vendedores.\n\nProblema: Lista única de vendedores\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CLASSIFICAR(ÚNICO(Base[Vendedor]))",
        "solucao_excel": "=CLASSIFICAR(ÚNICO(Base[Vendedor]))",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-077",
        "ferramenta": "Excel",
        "tema": "Matrizes Dinâmicas",
        "nivel": "Avançado",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-077 — Matrizes Dinâmicas (Avançado)\n\nContexto de negócio: Situação real de Financeiro: top 10 despesas.\n\nProblema: Top 10 despesas\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=PEGAR(CLASSIFICAR(Base;3;-1);10)",
        "solucao_excel": "=PEGAR(CLASSIFICAR(Base;3;-1);10)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-078",
        "ferramenta": "Excel",
        "tema": "Matrizes Dinâmicas",
        "nivel": "Avançado",
        "area_negocio": "Operações",
        "enunciado": "EXC-078 — Matrizes Dinâmicas (Avançado)\n\nContexto de negócio: Situação real de Operações: selecionar colunas principais.\n\nProblema: Selecionar colunas principais\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=ESCOLHERCOLS(Base;1;2;5)",
        "solucao_excel": "=ESCOLHERCOLS(Base;1;2;5)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Operações e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-079",
        "ferramenta": "Excel",
        "tema": "Auditoria",
        "nivel": "Avançado",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-079 — Auditoria (Avançado)\n\nContexto de negócio: Situação real de Financeiro: regra com let para limite.\n\nProblema: Regra com LET para limite\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=LET(v;B2;limite;10000;SE(v>limite;\"Aprovar\";\"OK\"))",
        "solucao_excel": "=LET(v;B2;limite;10000;SE(v>limite;\"Aprovar\";\"OK\"))",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-080",
        "ferramenta": "Excel",
        "tema": "Auditoria",
        "nivel": "Avançado",
        "area_negocio": "Dados",
        "enunciado": "EXC-080 — Auditoria (Avançado)\n\nContexto de negócio: Situação real de Dados: limpar lista com map.\n\nProblema: Limpar lista com MAP\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MAP(A2:A100;LAMBDA(x;MAIÚSCULA(ARRUMAR(x))))",
        "solucao_excel": "=MAP(A2:A100;LAMBDA(x;MAIÚSCULA(ARRUMAR(x))))",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Dados e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-081",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Intermediário",
        "area_negocio": "Comercial",
        "enunciado": "EXC-081 — Estatística no Excel (Intermediário)\n\nContexto de negócio: Situação real de Comercial: média de vendas.\n\nProblema: Média de vendas\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MÉDIA(Base[Valor])",
        "solucao_excel": "=MÉDIA(Base[Valor])",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-082",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Intermediário",
        "area_negocio": "Comercial",
        "enunciado": "EXC-082 — Estatística no Excel (Intermediário)\n\nContexto de negócio: Situação real de Comercial: mediana de vendas.\n\nProblema: Mediana de vendas\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MED(Base[Valor])",
        "solucao_excel": "=MED(Base[Valor])",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-083",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Intermediário",
        "area_negocio": "Operações",
        "enunciado": "EXC-083 — Estatística no Excel (Intermediário)\n\nContexto de negócio: Situação real de Operações: desvio padrão de sla.\n\nProblema: Desvio padrão de SLA\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=DESVPAD.S(Base[SLA])",
        "solucao_excel": "=DESVPAD.S(Base[SLA])",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Operações e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-084",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Avançado",
        "area_negocio": "Marketing",
        "enunciado": "EXC-084 — Estatística no Excel (Avançado)\n\nContexto de negócio: Situação real de Marketing: correlação investimento vendas.\n\nProblema: Correlação investimento vendas\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CORREL(Base[Investimento];Base[Vendas])",
        "solucao_excel": "=CORREL(Base[Investimento];Base[Vendas])",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Marketing e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-085",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Avançado",
        "area_negocio": "Comercial",
        "enunciado": "EXC-085 — Estatística no Excel (Avançado)\n\nContexto de negócio: Situação real de Comercial: projeção linear.\n\nProblema: Projeção linear\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=TENDÊNCIA(Base[Vendas];Base[Mes];D2:D6)",
        "solucao_excel": "=TENDÊNCIA(Base[Vendas];Base[Mes];D2:D6)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-086",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Básico",
        "area_negocio": "RH",
        "enunciado": "EXC-086 — Texto (Básico)\n\nContexto de negócio: Situação real de RH: padronizar cpf sem pontuação.\n\nProblema: Padronizar CPF sem pontuação\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SUBSTITUIR(SUBSTITUIR(SUBSTITUIR(A2;\".\";\"\");\"-\";\"\");\"/\";\"\")",
        "solucao_excel": "=SUBSTITUIR(SUBSTITUIR(SUBSTITUIR(A2;\".\";\"\");\"-\";\"\");\"/\";\"\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de RH e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-087",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Básico",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-087 — Texto (Básico)\n\nContexto de negócio: Situação real de Financeiro: extrair centro de custo de código.\n\nProblema: Extrair centro de custo de código\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=EXT.TEXTO(A2;4;3)",
        "solucao_excel": "=EXT.TEXTO(A2;4;3)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-088",
        "ferramenta": "Excel",
        "tema": "Texto",
        "nivel": "Intermediário",
        "area_negocio": "Marketing",
        "enunciado": "EXC-088 — Texto (Intermediário)\n\nContexto de negócio: Situação real de Marketing: montar chave de campanha.\n\nProblema: Montar chave de campanha\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=TEXTOJUNTAR(\"-\";VERDADEIRO;A2:C2)",
        "solucao_excel": "=TEXTOJUNTAR(\"-\";VERDADEIRO;A2:C2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Marketing e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-089",
        "ferramenta": "Excel",
        "tema": "Datas",
        "nivel": "Básico",
        "area_negocio": "Logística",
        "enunciado": "EXC-089 — Datas (Básico)\n\nContexto de negócio: Situação real de Logística: calcular sla em dias.\n\nProblema: Calcular SLA em dias\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=DIAS(B2;A2)",
        "solucao_excel": "=DIAS(B2;A2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Logística e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-090",
        "ferramenta": "Excel",
        "tema": "Datas",
        "nivel": "Intermediário",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-090 — Datas (Intermediário)\n\nContexto de negócio: Situação real de Financeiro: identificar vencido.\n\nProblema: Identificar vencido\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(HOJE()>A2;\"Vencido\";\"No prazo\")",
        "solucao_excel": "=SE(HOJE()>A2;\"Vencido\";\"No prazo\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-091",
        "ferramenta": "Excel",
        "tema": "Condicionais",
        "nivel": "Intermediário",
        "area_negocio": "RH",
        "enunciado": "EXC-091 — Condicionais (Intermediário)\n\nContexto de negócio: Situação real de RH: classificar absenteísmo.\n\nProblema: Classificar absenteísmo\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(B2>0,05;\"Crítico\";\"Normal\")",
        "solucao_excel": "=SE(B2>0,05;\"Crítico\";\"Normal\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de RH e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-092",
        "ferramenta": "Excel",
        "tema": "Condicionais",
        "nivel": "Intermediário",
        "area_negocio": "Supply Chain",
        "enunciado": "EXC-092 — Condicionais (Intermediário)\n\nContexto de negócio: Situação real de Supply Chain: status de ruptura.\n\nProblema: Status de ruptura\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(B2=0;\"Ruptura\";\"Com estoque\")",
        "solucao_excel": "=SE(B2=0;\"Ruptura\";\"Com estoque\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Supply Chain e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-093",
        "ferramenta": "Excel",
        "tema": "Buscas",
        "nivel": "Intermediário",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-093 — Buscas (Intermediário)\n\nContexto de negócio: Situação real de Financeiro: buscar conta contábil.\n\nProblema: Buscar conta contábil\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=PROCX(A2;Plano[Codigo];Plano[Conta];\"Não encontrada\")",
        "solucao_excel": "=PROCX(A2;Plano[Codigo];Plano[Conta];\"Não encontrada\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-094",
        "ferramenta": "Excel",
        "tema": "Buscas",
        "nivel": "Avançado",
        "area_negocio": "Comercial",
        "enunciado": "EXC-094 — Buscas (Avançado)\n\nContexto de negócio: Situação real de Comercial: busca parcial com procx.\n\nProblema: Busca parcial com PROCX\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=PROCX(\"*\"&A2&\"*\";Base[Produto];Base[Categoria];\"N/D\";2)",
        "solucao_excel": "=PROCX(\"*\"&A2&\"*\";Base[Produto];Base[Categoria];\"N/D\";2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-095",
        "ferramenta": "Excel",
        "tema": "Agregações",
        "nivel": "Intermediário",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-095 — Agregações (Intermediário)\n\nContexto de negócio: Situação real de Financeiro: somar despesas por centro.\n\nProblema: Somar despesas por centro\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SOMASES(Base[Valor];Base[Centro];A2;Base[Tipo];\"Despesa\")",
        "solucao_excel": "=SOMASES(Base[Valor];Base[Centro];A2;Base[Tipo];\"Despesa\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-096",
        "ferramenta": "Excel",
        "tema": "Agregações",
        "nivel": "Intermediário",
        "area_negocio": "Marketing",
        "enunciado": "EXC-096 — Agregações (Intermediário)\n\nContexto de negócio: Situação real de Marketing: média de roi por canal.\n\nProblema: Média de ROI por canal\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MÉDIASES(Base[ROI];Base[Canal];A2)",
        "solucao_excel": "=MÉDIASES(Base[ROI];Base[Canal];A2)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Marketing e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-097",
        "ferramenta": "Excel",
        "tema": "Agregações",
        "nivel": "Intermediário",
        "area_negocio": "RH",
        "enunciado": "EXC-097 — Agregações (Intermediário)\n\nContexto de negócio: Situação real de RH: contar colaboradores ativos.\n\nProblema: Contar colaboradores ativos\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CONT.SES(Base[Area];A2;Base[Status];\"Ativo\")",
        "solucao_excel": "=CONT.SES(Base[Area];A2;Base[Status];\"Ativo\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de RH e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-098",
        "ferramenta": "Excel",
        "tema": "Coringas",
        "nivel": "Avançado",
        "area_negocio": "Comercial",
        "enunciado": "EXC-098 — Coringas (Avançado)\n\nContexto de negócio: Situação real de Comercial: encontrar produtos lata.\n\nProblema: Encontrar produtos lata\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CONT.SE(A:A;\"*LATA*\")",
        "solucao_excel": "=CONT.SE(A:A;\"*LATA*\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-099",
        "ferramenta": "Excel",
        "tema": "Coringas",
        "nivel": "Avançado",
        "area_negocio": "Supply Chain",
        "enunciado": "EXC-099 — Coringas (Avançado)\n\nContexto de negócio: Situação real de Supply Chain: validar código 3 letras.\n\nProblema: Validar código 3 letras\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=SE(CONT.SE(A2;\"???\")>0;\"OK\";\"Validar\")",
        "solucao_excel": "=SE(CONT.SE(A2;\"???\")>0;\"OK\";\"Validar\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Supply Chain e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-100",
        "ferramenta": "Excel",
        "tema": "Coringas",
        "nivel": "Avançado",
        "area_negocio": "TI",
        "enunciado": "EXC-100 — Coringas (Avançado)\n\nContexto de negócio: Situação real de TI: buscar asterisco literal.\n\nProblema: Buscar asterisco literal\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CONT.SE(A:A;\"SKU~*\")",
        "solucao_excel": "=CONT.SE(A:A;\"SKU~*\")",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de TI e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-101",
        "ferramenta": "Excel",
        "tema": "Matrizes Dinâmicas",
        "nivel": "Intermediário",
        "area_negocio": "Comercial",
        "enunciado": "EXC-101 — Matrizes Dinâmicas (Intermediário)\n\nContexto de negócio: Situação real de Comercial: lista única de vendedores.\n\nProblema: Lista única de vendedores\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=CLASSIFICAR(ÚNICO(Base[Vendedor]))",
        "solucao_excel": "=CLASSIFICAR(ÚNICO(Base[Vendedor]))",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-102",
        "ferramenta": "Excel",
        "tema": "Matrizes Dinâmicas",
        "nivel": "Avançado",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-102 — Matrizes Dinâmicas (Avançado)\n\nContexto de negócio: Situação real de Financeiro: top 10 despesas.\n\nProblema: Top 10 despesas\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=PEGAR(CLASSIFICAR(Base;3;-1);10)",
        "solucao_excel": "=PEGAR(CLASSIFICAR(Base;3;-1);10)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-103",
        "ferramenta": "Excel",
        "tema": "Matrizes Dinâmicas",
        "nivel": "Avançado",
        "area_negocio": "Operações",
        "enunciado": "EXC-103 — Matrizes Dinâmicas (Avançado)\n\nContexto de negócio: Situação real de Operações: selecionar colunas principais.\n\nProblema: Selecionar colunas principais\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=ESCOLHERCOLS(Base;1;2;5)",
        "solucao_excel": "=ESCOLHERCOLS(Base;1;2;5)",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Operações e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-104",
        "ferramenta": "Excel",
        "tema": "Auditoria",
        "nivel": "Avançado",
        "area_negocio": "Financeiro",
        "enunciado": "EXC-104 — Auditoria (Avançado)\n\nContexto de negócio: Situação real de Financeiro: regra com let para limite.\n\nProblema: Regra com LET para limite\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=LET(v;B2;limite;10000;SE(v>limite;\"Aprovar\";\"OK\"))",
        "solucao_excel": "=LET(v;B2;limite;10000;SE(v>limite;\"Aprovar\";\"OK\"))",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Financeiro e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-105",
        "ferramenta": "Excel",
        "tema": "Auditoria",
        "nivel": "Avançado",
        "area_negocio": "Dados",
        "enunciado": "EXC-105 — Auditoria (Avançado)\n\nContexto de negócio: Situação real de Dados: limpar lista com map.\n\nProblema: Limpar lista com MAP\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MAP(A2:A100;LAMBDA(x;MAIÚSCULA(ARRUMAR(x))))",
        "solucao_excel": "=MAP(A2:A100;LAMBDA(x;MAIÚSCULA(ARRUMAR(x))))",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Dados e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "EXC-106",
        "ferramenta": "Excel",
        "tema": "Estatística no Excel",
        "nivel": "Intermediário",
        "area_negocio": "Comercial",
        "enunciado": "EXC-106 — Estatística no Excel (Intermediário)\n\nContexto de negócio: Situação real de Comercial: média de vendas.\n\nProblema: Média de vendas\n\nBase simulada: Base simulada com colunas compatíveis com o tema.\n\nTarefa: resolva usando Excel, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar campos, validar formato e aplicar fórmula pt-BR correta.",
        "raciocinio": "Escolher função adequada, tratar erros e preservar rastreabilidade.",
        "solucao_principal": "=MÉDIA(Base[Valor])",
        "solucao_excel": "=MÉDIA(Base[Valor])",
        "solucao_powerquery": "Power Query pode automatizar a transformação quando recorrente.",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Resultado esperado calculado conforme os dados da base simulada.",
        "erros_comuns": "Erro comum: aplicar a fórmula sem limpar dados, sem travar referências ou sem tratar erros.",
        "alternativas": "Alternativa: tabela dinâmica, Power Query, DAX ou coluna auxiliar conforme recorrência.",
        "aplicacao_pratica": "Aplicação prática em rotina de Comercial e preparação para teste técnico.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-001",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "Comercial",
        "enunciado": "PQ-001 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de Comercial: padronizar produto com text.upper/trim/clean.\n\nProblema: Padronizar produto com Text.Upper/Trim/Clean\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Produto_Limpo\", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Produto_Limpo\", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-002",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "RH",
        "enunciado": "PQ-002 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de RH: remover pontuação de documento.\n\nProblema: Remover pontuação de documento\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.TransformColumns(Fonte, {{\"CPF\", each Text.Select(_, {\"0\"..\"9\"}), type text}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.TransformColumns(Fonte, {{\"CPF\", each Text.Select(_, {\"0\"..\"9\"}), type text}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de RH, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-003",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "Marketing",
        "enunciado": "PQ-003 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de Marketing: criar chave de campanha com text.combine.\n\nProblema: Criar chave de campanha com Text.Combine\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Chave_Campanha\", each Text.Combine({[Canal], [Campanha], [Mes]}, \"|\"), type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Chave_Campanha\", each Text.Combine({[Canal], [Campanha], [Mes]}, \"|\"), type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Marketing, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-004",
        "ferramenta": "Power Query",
        "tema": "Tipos",
        "nivel": "ETL",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-004 — Tipos (ETL)\n\nContexto de negócio: Cenário real de Financeiro: converter valor com try.\n\nProblema: Converter valor com try\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Valor_Numero\", each try Number.From([Valor]) otherwise null, type number)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Valor_Numero\", each try Number.From([Valor]) otherwise null, type number)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-005",
        "ferramenta": "Power Query",
        "tema": "Datas",
        "nivel": "Fundamentos",
        "area_negocio": "Logística",
        "enunciado": "PQ-005 — Datas (Fundamentos)\n\nContexto de negócio: Cenário real de Logística: calcular lead time em dias.\n\nProblema: Calcular lead time em dias\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"LeadTime\", each try Duration.Days([Entrega] - [Pedido]) otherwise null, Int64.Type)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"LeadTime\", each try Duration.Days([Entrega] - [Pedido]) otherwise null, Int64.Type)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Logística, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-006",
        "ferramenta": "Power Query",
        "tema": "Datas",
        "nivel": "Fundamentos",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-006 — Datas (Fundamentos)\n\nContexto de negócio: Cenário real de Financeiro: criar competência mensal.\n\nProblema: Criar competência mensal\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Competencia\", each Date.StartOfMonth(Date.From([Data])), type date)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Competencia\", each Date.StartOfMonth(Date.From([Data])), type date)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-007",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-007 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: selecionar colunas obrigatórias.\n\nProblema: Selecionar colunas obrigatórias\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.SelectColumns(Fonte, {\"Produto\", \"Data\", \"Valor\"}, MissingField.UseNull)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.SelectColumns(Fonte, {\"Produto\", \"Data\", \"Valor\"}, MissingField.UseNull)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-008",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-008 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: renomear colunas de sistema.\n\nProblema: Renomear colunas de sistema\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.RenameColumns(Fonte, {{\"Vlr\", \"Valor\"}, {\"Dt\", \"Data\"}}, MissingField.Ignore)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.RenameColumns(Fonte, {{\"Vlr\", \"Valor\"}, {\"Dt\", \"Data\"}}, MissingField.Ignore)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-009",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-009 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: adicionar índice depois de ordenar.\n\nProblema: Adicionar índice depois de ordenar\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddIndexColumn(Table.Sort(Fonte, {{\"Data\", Order.Ascending}}), \"Indice\", 1, 1, Int64.Type)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddIndexColumn(Table.Sort(Fonte, {{\"Data\", Order.Ascending}}), \"Indice\", 1, 1, Int64.Type)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-010",
        "ferramenta": "Power Query",
        "tema": "Agrupamento",
        "nivel": "ETL",
        "area_negocio": "Comercial",
        "enunciado": "PQ-010 — Agrupamento (ETL)\n\nContexto de negócio: Cenário real de Comercial: agrupar vendas por produto.\n\nProblema: Agrupar vendas por produto\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Group(Fonte, {\"Produto\"}, {{\"Total\", each List.Sum([Valor]), type number}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Group(Fonte, {\"Produto\"}, {{\"Total\", each List.Sum([Valor]), type number}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-011",
        "ferramenta": "Power Query",
        "tema": "Agrupamento",
        "nivel": "ETL",
        "area_negocio": "RH",
        "enunciado": "PQ-011 — Agrupamento (ETL)\n\nContexto de negócio: Cenário real de RH: agrupar headcount por área.\n\nProblema: Agrupar headcount por área\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Group(Fonte, {\"Area\"}, {{\"Headcount\", each Table.RowCount(_), Int64.Type}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Group(Fonte, {\"Area\"}, {{\"Headcount\", each Table.RowCount(_), Int64.Type}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de RH, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-012",
        "ferramenta": "Power Query",
        "tema": "Merge",
        "nivel": "ETL",
        "area_negocio": "Supply Chain",
        "enunciado": "PQ-012 — Merge (ETL)\n\nContexto de negócio: Cenário real de Supply Chain: mesclar fato com dimensão.\n\nProblema: Mesclar fato com dimensão\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.NestedJoin(Fato, {\"SKU\"}, DimProduto, {\"SKU\"}, \"DimProduto\", JoinKind.LeftOuter)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.NestedJoin(Fato, {\"SKU\"}, DimProduto, {\"SKU\"}, \"DimProduto\", JoinKind.LeftOuter)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Supply Chain, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-013",
        "ferramenta": "Power Query",
        "tema": "Merge",
        "nivel": "Data Quality",
        "area_negocio": "Supply Chain",
        "enunciado": "PQ-013 — Merge (Data Quality)\n\nContexto de negócio: Cenário real de Supply Chain: identificar produtos sem cadastro após merge.\n\nProblema: Identificar produtos sem cadastro após merge\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.SelectRows(Expandido, each [Categoria] = null)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.SelectRows(Expandido, each [Categoria] = null)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Supply Chain, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-014",
        "ferramenta": "Power Query",
        "tema": "Append",
        "nivel": "ETL",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-014 — Append (ETL)\n\nContexto de negócio: Cenário real de Financeiro: combinar bases mensais.\n\nProblema: Combinar bases mensais\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Combine({BaseJan, BaseFev, BaseMar})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Combine({BaseJan, BaseFev, BaseMar})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-015",
        "ferramenta": "Power Query",
        "tema": "Data Quality",
        "nivel": "Data Quality",
        "area_negocio": "Dados",
        "enunciado": "PQ-015 — Data Quality (Data Quality)\n\nContexto de negócio: Cenário real de Dados: criar status_dq para valor nulo.\n\nProblema: Criar Status_DQ para valor nulo\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Status_DQ\", each if [Valor] = null then \"ERRO VALOR\" else \"OK\", type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Status_DQ\", each if [Valor] = null then \"ERRO VALOR\" else \"OK\", type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-016",
        "ferramenta": "Power Query",
        "tema": "Data Quality",
        "nivel": "Data Quality",
        "area_negocio": "Dados",
        "enunciado": "PQ-016 — Data Quality (Data Quality)\n\nContexto de negócio: Cenário real de Dados: gerar perfil com table.profile.\n\nProblema: Gerar perfil com Table.Profile\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Profile(Fonte)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Profile(Fonte)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-017",
        "ferramenta": "Power Query",
        "tema": "Performance",
        "nivel": "Performance",
        "area_negocio": "Dados",
        "enunciado": "PQ-017 — Performance (Performance)\n\nContexto de negócio: Cenário real de Dados: aplicar table.buffer em dimensão.\n\nProblema: Aplicar Table.Buffer em dimensão\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Buffer(DimProduto)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Buffer(DimProduto)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-018",
        "ferramenta": "Power Query",
        "tema": "Listas",
        "nivel": "Avançado",
        "area_negocio": "Dados",
        "enunciado": "PQ-018 — Listas (Avançado)\n\nContexto de negócio: Cenário real de Dados: aplicar list.accumulate para substituições.\n\nProblema: Aplicar List.Accumulate para substituições\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-019",
        "ferramenta": "Power Query",
        "tema": "Records",
        "nivel": "Avançado",
        "area_negocio": "Dados",
        "enunciado": "PQ-019 — Records (Avançado)\n\nContexto de negócio: Cenário real de Dados: ler campo de record com default.\n\nProblema: Ler campo de record com default\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Record.FieldOrDefault([Registro], \"Campo\", null)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Record.FieldOrDefault([Registro], \"Campo\", null)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-020",
        "ferramenta": "Power Query",
        "tema": "Modelagem",
        "nivel": "Modelagem",
        "area_negocio": "Comercial",
        "enunciado": "PQ-020 — Modelagem (Modelagem)\n\nContexto de negócio: Cenário real de Comercial: normalizar tabela larga com unpivot.\n\nProblema: Normalizar tabela larga com Unpivot\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.UnpivotOtherColumns(Fonte, {\"Produto\"}, \"Indicador\", \"Valor\")",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.UnpivotOtherColumns(Fonte, {\"Produto\"}, \"Indicador\", \"Valor\")",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-021",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "Comercial",
        "enunciado": "PQ-021 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de Comercial: padronizar produto com text.upper/trim/clean.\n\nProblema: Padronizar produto com Text.Upper/Trim/Clean\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Produto_Limpo\", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Produto_Limpo\", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-022",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "RH",
        "enunciado": "PQ-022 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de RH: remover pontuação de documento.\n\nProblema: Remover pontuação de documento\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.TransformColumns(Fonte, {{\"CPF\", each Text.Select(_, {\"0\"..\"9\"}), type text}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.TransformColumns(Fonte, {{\"CPF\", each Text.Select(_, {\"0\"..\"9\"}), type text}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de RH, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-023",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "Marketing",
        "enunciado": "PQ-023 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de Marketing: criar chave de campanha com text.combine.\n\nProblema: Criar chave de campanha com Text.Combine\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Chave_Campanha\", each Text.Combine({[Canal], [Campanha], [Mes]}, \"|\"), type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Chave_Campanha\", each Text.Combine({[Canal], [Campanha], [Mes]}, \"|\"), type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Marketing, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-024",
        "ferramenta": "Power Query",
        "tema": "Tipos",
        "nivel": "ETL",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-024 — Tipos (ETL)\n\nContexto de negócio: Cenário real de Financeiro: converter valor com try.\n\nProblema: Converter valor com try\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Valor_Numero\", each try Number.From([Valor]) otherwise null, type number)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Valor_Numero\", each try Number.From([Valor]) otherwise null, type number)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-025",
        "ferramenta": "Power Query",
        "tema": "Datas",
        "nivel": "Fundamentos",
        "area_negocio": "Logística",
        "enunciado": "PQ-025 — Datas (Fundamentos)\n\nContexto de negócio: Cenário real de Logística: calcular lead time em dias.\n\nProblema: Calcular lead time em dias\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"LeadTime\", each try Duration.Days([Entrega] - [Pedido]) otherwise null, Int64.Type)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"LeadTime\", each try Duration.Days([Entrega] - [Pedido]) otherwise null, Int64.Type)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Logística, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-026",
        "ferramenta": "Power Query",
        "tema": "Datas",
        "nivel": "Fundamentos",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-026 — Datas (Fundamentos)\n\nContexto de negócio: Cenário real de Financeiro: criar competência mensal.\n\nProblema: Criar competência mensal\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Competencia\", each Date.StartOfMonth(Date.From([Data])), type date)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Competencia\", each Date.StartOfMonth(Date.From([Data])), type date)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-027",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-027 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: selecionar colunas obrigatórias.\n\nProblema: Selecionar colunas obrigatórias\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.SelectColumns(Fonte, {\"Produto\", \"Data\", \"Valor\"}, MissingField.UseNull)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.SelectColumns(Fonte, {\"Produto\", \"Data\", \"Valor\"}, MissingField.UseNull)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-028",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-028 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: renomear colunas de sistema.\n\nProblema: Renomear colunas de sistema\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.RenameColumns(Fonte, {{\"Vlr\", \"Valor\"}, {\"Dt\", \"Data\"}}, MissingField.Ignore)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.RenameColumns(Fonte, {{\"Vlr\", \"Valor\"}, {\"Dt\", \"Data\"}}, MissingField.Ignore)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-029",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-029 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: adicionar índice depois de ordenar.\n\nProblema: Adicionar índice depois de ordenar\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddIndexColumn(Table.Sort(Fonte, {{\"Data\", Order.Ascending}}), \"Indice\", 1, 1, Int64.Type)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddIndexColumn(Table.Sort(Fonte, {{\"Data\", Order.Ascending}}), \"Indice\", 1, 1, Int64.Type)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-030",
        "ferramenta": "Power Query",
        "tema": "Agrupamento",
        "nivel": "ETL",
        "area_negocio": "Comercial",
        "enunciado": "PQ-030 — Agrupamento (ETL)\n\nContexto de negócio: Cenário real de Comercial: agrupar vendas por produto.\n\nProblema: Agrupar vendas por produto\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Group(Fonte, {\"Produto\"}, {{\"Total\", each List.Sum([Valor]), type number}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Group(Fonte, {\"Produto\"}, {{\"Total\", each List.Sum([Valor]), type number}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-031",
        "ferramenta": "Power Query",
        "tema": "Agrupamento",
        "nivel": "ETL",
        "area_negocio": "RH",
        "enunciado": "PQ-031 — Agrupamento (ETL)\n\nContexto de negócio: Cenário real de RH: agrupar headcount por área.\n\nProblema: Agrupar headcount por área\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Group(Fonte, {\"Area\"}, {{\"Headcount\", each Table.RowCount(_), Int64.Type}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Group(Fonte, {\"Area\"}, {{\"Headcount\", each Table.RowCount(_), Int64.Type}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de RH, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-032",
        "ferramenta": "Power Query",
        "tema": "Merge",
        "nivel": "ETL",
        "area_negocio": "Supply Chain",
        "enunciado": "PQ-032 — Merge (ETL)\n\nContexto de negócio: Cenário real de Supply Chain: mesclar fato com dimensão.\n\nProblema: Mesclar fato com dimensão\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.NestedJoin(Fato, {\"SKU\"}, DimProduto, {\"SKU\"}, \"DimProduto\", JoinKind.LeftOuter)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.NestedJoin(Fato, {\"SKU\"}, DimProduto, {\"SKU\"}, \"DimProduto\", JoinKind.LeftOuter)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Supply Chain, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-033",
        "ferramenta": "Power Query",
        "tema": "Merge",
        "nivel": "Data Quality",
        "area_negocio": "Supply Chain",
        "enunciado": "PQ-033 — Merge (Data Quality)\n\nContexto de negócio: Cenário real de Supply Chain: identificar produtos sem cadastro após merge.\n\nProblema: Identificar produtos sem cadastro após merge\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.SelectRows(Expandido, each [Categoria] = null)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.SelectRows(Expandido, each [Categoria] = null)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Supply Chain, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-034",
        "ferramenta": "Power Query",
        "tema": "Append",
        "nivel": "ETL",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-034 — Append (ETL)\n\nContexto de negócio: Cenário real de Financeiro: combinar bases mensais.\n\nProblema: Combinar bases mensais\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Combine({BaseJan, BaseFev, BaseMar})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Combine({BaseJan, BaseFev, BaseMar})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-035",
        "ferramenta": "Power Query",
        "tema": "Data Quality",
        "nivel": "Data Quality",
        "area_negocio": "Dados",
        "enunciado": "PQ-035 — Data Quality (Data Quality)\n\nContexto de negócio: Cenário real de Dados: criar status_dq para valor nulo.\n\nProblema: Criar Status_DQ para valor nulo\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Status_DQ\", each if [Valor] = null then \"ERRO VALOR\" else \"OK\", type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Status_DQ\", each if [Valor] = null then \"ERRO VALOR\" else \"OK\", type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-036",
        "ferramenta": "Power Query",
        "tema": "Data Quality",
        "nivel": "Data Quality",
        "area_negocio": "Dados",
        "enunciado": "PQ-036 — Data Quality (Data Quality)\n\nContexto de negócio: Cenário real de Dados: gerar perfil com table.profile.\n\nProblema: Gerar perfil com Table.Profile\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Profile(Fonte)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Profile(Fonte)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-037",
        "ferramenta": "Power Query",
        "tema": "Performance",
        "nivel": "Performance",
        "area_negocio": "Dados",
        "enunciado": "PQ-037 — Performance (Performance)\n\nContexto de negócio: Cenário real de Dados: aplicar table.buffer em dimensão.\n\nProblema: Aplicar Table.Buffer em dimensão\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Buffer(DimProduto)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Buffer(DimProduto)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-038",
        "ferramenta": "Power Query",
        "tema": "Listas",
        "nivel": "Avançado",
        "area_negocio": "Dados",
        "enunciado": "PQ-038 — Listas (Avançado)\n\nContexto de negócio: Cenário real de Dados: aplicar list.accumulate para substituições.\n\nProblema: Aplicar List.Accumulate para substituições\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-039",
        "ferramenta": "Power Query",
        "tema": "Records",
        "nivel": "Avançado",
        "area_negocio": "Dados",
        "enunciado": "PQ-039 — Records (Avançado)\n\nContexto de negócio: Cenário real de Dados: ler campo de record com default.\n\nProblema: Ler campo de record com default\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Record.FieldOrDefault([Registro], \"Campo\", null)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Record.FieldOrDefault([Registro], \"Campo\", null)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-040",
        "ferramenta": "Power Query",
        "tema": "Modelagem",
        "nivel": "Modelagem",
        "area_negocio": "Comercial",
        "enunciado": "PQ-040 — Modelagem (Modelagem)\n\nContexto de negócio: Cenário real de Comercial: normalizar tabela larga com unpivot.\n\nProblema: Normalizar tabela larga com Unpivot\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.UnpivotOtherColumns(Fonte, {\"Produto\"}, \"Indicador\", \"Valor\")",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.UnpivotOtherColumns(Fonte, {\"Produto\"}, \"Indicador\", \"Valor\")",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-041",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "Comercial",
        "enunciado": "PQ-041 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de Comercial: padronizar produto com text.upper/trim/clean.\n\nProblema: Padronizar produto com Text.Upper/Trim/Clean\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Produto_Limpo\", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Produto_Limpo\", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-042",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "RH",
        "enunciado": "PQ-042 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de RH: remover pontuação de documento.\n\nProblema: Remover pontuação de documento\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.TransformColumns(Fonte, {{\"CPF\", each Text.Select(_, {\"0\"..\"9\"}), type text}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.TransformColumns(Fonte, {{\"CPF\", each Text.Select(_, {\"0\"..\"9\"}), type text}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de RH, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-043",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "Marketing",
        "enunciado": "PQ-043 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de Marketing: criar chave de campanha com text.combine.\n\nProblema: Criar chave de campanha com Text.Combine\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Chave_Campanha\", each Text.Combine({[Canal], [Campanha], [Mes]}, \"|\"), type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Chave_Campanha\", each Text.Combine({[Canal], [Campanha], [Mes]}, \"|\"), type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Marketing, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-044",
        "ferramenta": "Power Query",
        "tema": "Tipos",
        "nivel": "ETL",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-044 — Tipos (ETL)\n\nContexto de negócio: Cenário real de Financeiro: converter valor com try.\n\nProblema: Converter valor com try\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Valor_Numero\", each try Number.From([Valor]) otherwise null, type number)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Valor_Numero\", each try Number.From([Valor]) otherwise null, type number)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-045",
        "ferramenta": "Power Query",
        "tema": "Datas",
        "nivel": "Fundamentos",
        "area_negocio": "Logística",
        "enunciado": "PQ-045 — Datas (Fundamentos)\n\nContexto de negócio: Cenário real de Logística: calcular lead time em dias.\n\nProblema: Calcular lead time em dias\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"LeadTime\", each try Duration.Days([Entrega] - [Pedido]) otherwise null, Int64.Type)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"LeadTime\", each try Duration.Days([Entrega] - [Pedido]) otherwise null, Int64.Type)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Logística, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-046",
        "ferramenta": "Power Query",
        "tema": "Datas",
        "nivel": "Fundamentos",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-046 — Datas (Fundamentos)\n\nContexto de negócio: Cenário real de Financeiro: criar competência mensal.\n\nProblema: Criar competência mensal\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Competencia\", each Date.StartOfMonth(Date.From([Data])), type date)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Competencia\", each Date.StartOfMonth(Date.From([Data])), type date)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-047",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-047 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: selecionar colunas obrigatórias.\n\nProblema: Selecionar colunas obrigatórias\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.SelectColumns(Fonte, {\"Produto\", \"Data\", \"Valor\"}, MissingField.UseNull)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.SelectColumns(Fonte, {\"Produto\", \"Data\", \"Valor\"}, MissingField.UseNull)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-048",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-048 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: renomear colunas de sistema.\n\nProblema: Renomear colunas de sistema\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.RenameColumns(Fonte, {{\"Vlr\", \"Valor\"}, {\"Dt\", \"Data\"}}, MissingField.Ignore)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.RenameColumns(Fonte, {{\"Vlr\", \"Valor\"}, {\"Dt\", \"Data\"}}, MissingField.Ignore)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-049",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-049 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: adicionar índice depois de ordenar.\n\nProblema: Adicionar índice depois de ordenar\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddIndexColumn(Table.Sort(Fonte, {{\"Data\", Order.Ascending}}), \"Indice\", 1, 1, Int64.Type)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddIndexColumn(Table.Sort(Fonte, {{\"Data\", Order.Ascending}}), \"Indice\", 1, 1, Int64.Type)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-050",
        "ferramenta": "Power Query",
        "tema": "Agrupamento",
        "nivel": "ETL",
        "area_negocio": "Comercial",
        "enunciado": "PQ-050 — Agrupamento (ETL)\n\nContexto de negócio: Cenário real de Comercial: agrupar vendas por produto.\n\nProblema: Agrupar vendas por produto\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Group(Fonte, {\"Produto\"}, {{\"Total\", each List.Sum([Valor]), type number}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Group(Fonte, {\"Produto\"}, {{\"Total\", each List.Sum([Valor]), type number}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-051",
        "ferramenta": "Power Query",
        "tema": "Agrupamento",
        "nivel": "ETL",
        "area_negocio": "RH",
        "enunciado": "PQ-051 — Agrupamento (ETL)\n\nContexto de negócio: Cenário real de RH: agrupar headcount por área.\n\nProblema: Agrupar headcount por área\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Group(Fonte, {\"Area\"}, {{\"Headcount\", each Table.RowCount(_), Int64.Type}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Group(Fonte, {\"Area\"}, {{\"Headcount\", each Table.RowCount(_), Int64.Type}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de RH, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-052",
        "ferramenta": "Power Query",
        "tema": "Merge",
        "nivel": "ETL",
        "area_negocio": "Supply Chain",
        "enunciado": "PQ-052 — Merge (ETL)\n\nContexto de negócio: Cenário real de Supply Chain: mesclar fato com dimensão.\n\nProblema: Mesclar fato com dimensão\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.NestedJoin(Fato, {\"SKU\"}, DimProduto, {\"SKU\"}, \"DimProduto\", JoinKind.LeftOuter)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.NestedJoin(Fato, {\"SKU\"}, DimProduto, {\"SKU\"}, \"DimProduto\", JoinKind.LeftOuter)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Supply Chain, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-053",
        "ferramenta": "Power Query",
        "tema": "Merge",
        "nivel": "Data Quality",
        "area_negocio": "Supply Chain",
        "enunciado": "PQ-053 — Merge (Data Quality)\n\nContexto de negócio: Cenário real de Supply Chain: identificar produtos sem cadastro após merge.\n\nProblema: Identificar produtos sem cadastro após merge\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.SelectRows(Expandido, each [Categoria] = null)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.SelectRows(Expandido, each [Categoria] = null)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Supply Chain, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-054",
        "ferramenta": "Power Query",
        "tema": "Append",
        "nivel": "ETL",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-054 — Append (ETL)\n\nContexto de negócio: Cenário real de Financeiro: combinar bases mensais.\n\nProblema: Combinar bases mensais\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Combine({BaseJan, BaseFev, BaseMar})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Combine({BaseJan, BaseFev, BaseMar})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-055",
        "ferramenta": "Power Query",
        "tema": "Data Quality",
        "nivel": "Data Quality",
        "area_negocio": "Dados",
        "enunciado": "PQ-055 — Data Quality (Data Quality)\n\nContexto de negócio: Cenário real de Dados: criar status_dq para valor nulo.\n\nProblema: Criar Status_DQ para valor nulo\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Status_DQ\", each if [Valor] = null then \"ERRO VALOR\" else \"OK\", type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Status_DQ\", each if [Valor] = null then \"ERRO VALOR\" else \"OK\", type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-056",
        "ferramenta": "Power Query",
        "tema": "Data Quality",
        "nivel": "Data Quality",
        "area_negocio": "Dados",
        "enunciado": "PQ-056 — Data Quality (Data Quality)\n\nContexto de negócio: Cenário real de Dados: gerar perfil com table.profile.\n\nProblema: Gerar perfil com Table.Profile\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Profile(Fonte)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Profile(Fonte)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-057",
        "ferramenta": "Power Query",
        "tema": "Performance",
        "nivel": "Performance",
        "area_negocio": "Dados",
        "enunciado": "PQ-057 — Performance (Performance)\n\nContexto de negócio: Cenário real de Dados: aplicar table.buffer em dimensão.\n\nProblema: Aplicar Table.Buffer em dimensão\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Buffer(DimProduto)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Buffer(DimProduto)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-058",
        "ferramenta": "Power Query",
        "tema": "Listas",
        "nivel": "Avançado",
        "area_negocio": "Dados",
        "enunciado": "PQ-058 — Listas (Avançado)\n\nContexto de negócio: Cenário real de Dados: aplicar list.accumulate para substituições.\n\nProblema: Aplicar List.Accumulate para substituições\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-059",
        "ferramenta": "Power Query",
        "tema": "Records",
        "nivel": "Avançado",
        "area_negocio": "Dados",
        "enunciado": "PQ-059 — Records (Avançado)\n\nContexto de negócio: Cenário real de Dados: ler campo de record com default.\n\nProblema: Ler campo de record com default\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Record.FieldOrDefault([Registro], \"Campo\", null)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Record.FieldOrDefault([Registro], \"Campo\", null)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-060",
        "ferramenta": "Power Query",
        "tema": "Modelagem",
        "nivel": "Modelagem",
        "area_negocio": "Comercial",
        "enunciado": "PQ-060 — Modelagem (Modelagem)\n\nContexto de negócio: Cenário real de Comercial: normalizar tabela larga com unpivot.\n\nProblema: Normalizar tabela larga com Unpivot\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.UnpivotOtherColumns(Fonte, {\"Produto\"}, \"Indicador\", \"Valor\")",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.UnpivotOtherColumns(Fonte, {\"Produto\"}, \"Indicador\", \"Valor\")",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-061",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "Comercial",
        "enunciado": "PQ-061 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de Comercial: padronizar produto com text.upper/trim/clean.\n\nProblema: Padronizar produto com Text.Upper/Trim/Clean\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Produto_Limpo\", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Produto_Limpo\", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-062",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "RH",
        "enunciado": "PQ-062 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de RH: remover pontuação de documento.\n\nProblema: Remover pontuação de documento\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.TransformColumns(Fonte, {{\"CPF\", each Text.Select(_, {\"0\"..\"9\"}), type text}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.TransformColumns(Fonte, {{\"CPF\", each Text.Select(_, {\"0\"..\"9\"}), type text}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de RH, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-063",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "Marketing",
        "enunciado": "PQ-063 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de Marketing: criar chave de campanha com text.combine.\n\nProblema: Criar chave de campanha com Text.Combine\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Chave_Campanha\", each Text.Combine({[Canal], [Campanha], [Mes]}, \"|\"), type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Chave_Campanha\", each Text.Combine({[Canal], [Campanha], [Mes]}, \"|\"), type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Marketing, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-064",
        "ferramenta": "Power Query",
        "tema": "Tipos",
        "nivel": "ETL",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-064 — Tipos (ETL)\n\nContexto de negócio: Cenário real de Financeiro: converter valor com try.\n\nProblema: Converter valor com try\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Valor_Numero\", each try Number.From([Valor]) otherwise null, type number)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Valor_Numero\", each try Number.From([Valor]) otherwise null, type number)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-065",
        "ferramenta": "Power Query",
        "tema": "Datas",
        "nivel": "Fundamentos",
        "area_negocio": "Logística",
        "enunciado": "PQ-065 — Datas (Fundamentos)\n\nContexto de negócio: Cenário real de Logística: calcular lead time em dias.\n\nProblema: Calcular lead time em dias\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"LeadTime\", each try Duration.Days([Entrega] - [Pedido]) otherwise null, Int64.Type)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"LeadTime\", each try Duration.Days([Entrega] - [Pedido]) otherwise null, Int64.Type)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Logística, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-066",
        "ferramenta": "Power Query",
        "tema": "Datas",
        "nivel": "Fundamentos",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-066 — Datas (Fundamentos)\n\nContexto de negócio: Cenário real de Financeiro: criar competência mensal.\n\nProblema: Criar competência mensal\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Competencia\", each Date.StartOfMonth(Date.From([Data])), type date)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Competencia\", each Date.StartOfMonth(Date.From([Data])), type date)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-067",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-067 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: selecionar colunas obrigatórias.\n\nProblema: Selecionar colunas obrigatórias\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.SelectColumns(Fonte, {\"Produto\", \"Data\", \"Valor\"}, MissingField.UseNull)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.SelectColumns(Fonte, {\"Produto\", \"Data\", \"Valor\"}, MissingField.UseNull)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-068",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-068 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: renomear colunas de sistema.\n\nProblema: Renomear colunas de sistema\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.RenameColumns(Fonte, {{\"Vlr\", \"Valor\"}, {\"Dt\", \"Data\"}}, MissingField.Ignore)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.RenameColumns(Fonte, {{\"Vlr\", \"Valor\"}, {\"Dt\", \"Data\"}}, MissingField.Ignore)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-069",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-069 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: adicionar índice depois de ordenar.\n\nProblema: Adicionar índice depois de ordenar\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddIndexColumn(Table.Sort(Fonte, {{\"Data\", Order.Ascending}}), \"Indice\", 1, 1, Int64.Type)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddIndexColumn(Table.Sort(Fonte, {{\"Data\", Order.Ascending}}), \"Indice\", 1, 1, Int64.Type)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-070",
        "ferramenta": "Power Query",
        "tema": "Agrupamento",
        "nivel": "ETL",
        "area_negocio": "Comercial",
        "enunciado": "PQ-070 — Agrupamento (ETL)\n\nContexto de negócio: Cenário real de Comercial: agrupar vendas por produto.\n\nProblema: Agrupar vendas por produto\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Group(Fonte, {\"Produto\"}, {{\"Total\", each List.Sum([Valor]), type number}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Group(Fonte, {\"Produto\"}, {{\"Total\", each List.Sum([Valor]), type number}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-071",
        "ferramenta": "Power Query",
        "tema": "Agrupamento",
        "nivel": "ETL",
        "area_negocio": "RH",
        "enunciado": "PQ-071 — Agrupamento (ETL)\n\nContexto de negócio: Cenário real de RH: agrupar headcount por área.\n\nProblema: Agrupar headcount por área\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Group(Fonte, {\"Area\"}, {{\"Headcount\", each Table.RowCount(_), Int64.Type}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Group(Fonte, {\"Area\"}, {{\"Headcount\", each Table.RowCount(_), Int64.Type}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de RH, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-072",
        "ferramenta": "Power Query",
        "tema": "Merge",
        "nivel": "ETL",
        "area_negocio": "Supply Chain",
        "enunciado": "PQ-072 — Merge (ETL)\n\nContexto de negócio: Cenário real de Supply Chain: mesclar fato com dimensão.\n\nProblema: Mesclar fato com dimensão\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.NestedJoin(Fato, {\"SKU\"}, DimProduto, {\"SKU\"}, \"DimProduto\", JoinKind.LeftOuter)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.NestedJoin(Fato, {\"SKU\"}, DimProduto, {\"SKU\"}, \"DimProduto\", JoinKind.LeftOuter)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Supply Chain, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-073",
        "ferramenta": "Power Query",
        "tema": "Merge",
        "nivel": "Data Quality",
        "area_negocio": "Supply Chain",
        "enunciado": "PQ-073 — Merge (Data Quality)\n\nContexto de negócio: Cenário real de Supply Chain: identificar produtos sem cadastro após merge.\n\nProblema: Identificar produtos sem cadastro após merge\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.SelectRows(Expandido, each [Categoria] = null)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.SelectRows(Expandido, each [Categoria] = null)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Supply Chain, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-074",
        "ferramenta": "Power Query",
        "tema": "Append",
        "nivel": "ETL",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-074 — Append (ETL)\n\nContexto de negócio: Cenário real de Financeiro: combinar bases mensais.\n\nProblema: Combinar bases mensais\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Combine({BaseJan, BaseFev, BaseMar})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Combine({BaseJan, BaseFev, BaseMar})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-075",
        "ferramenta": "Power Query",
        "tema": "Data Quality",
        "nivel": "Data Quality",
        "area_negocio": "Dados",
        "enunciado": "PQ-075 — Data Quality (Data Quality)\n\nContexto de negócio: Cenário real de Dados: criar status_dq para valor nulo.\n\nProblema: Criar Status_DQ para valor nulo\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Status_DQ\", each if [Valor] = null then \"ERRO VALOR\" else \"OK\", type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Status_DQ\", each if [Valor] = null then \"ERRO VALOR\" else \"OK\", type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-076",
        "ferramenta": "Power Query",
        "tema": "Data Quality",
        "nivel": "Data Quality",
        "area_negocio": "Dados",
        "enunciado": "PQ-076 — Data Quality (Data Quality)\n\nContexto de negócio: Cenário real de Dados: gerar perfil com table.profile.\n\nProblema: Gerar perfil com Table.Profile\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Profile(Fonte)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Profile(Fonte)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-077",
        "ferramenta": "Power Query",
        "tema": "Performance",
        "nivel": "Performance",
        "area_negocio": "Dados",
        "enunciado": "PQ-077 — Performance (Performance)\n\nContexto de negócio: Cenário real de Dados: aplicar table.buffer em dimensão.\n\nProblema: Aplicar Table.Buffer em dimensão\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Buffer(DimProduto)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Buffer(DimProduto)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-078",
        "ferramenta": "Power Query",
        "tema": "Listas",
        "nivel": "Avançado",
        "area_negocio": "Dados",
        "enunciado": "PQ-078 — Listas (Avançado)\n\nContexto de negócio: Cenário real de Dados: aplicar list.accumulate para substituições.\n\nProblema: Aplicar List.Accumulate para substituições\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-079",
        "ferramenta": "Power Query",
        "tema": "Records",
        "nivel": "Avançado",
        "area_negocio": "Dados",
        "enunciado": "PQ-079 — Records (Avançado)\n\nContexto de negócio: Cenário real de Dados: ler campo de record com default.\n\nProblema: Ler campo de record com default\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Record.FieldOrDefault([Registro], \"Campo\", null)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Record.FieldOrDefault([Registro], \"Campo\", null)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-080",
        "ferramenta": "Power Query",
        "tema": "Modelagem",
        "nivel": "Modelagem",
        "area_negocio": "Comercial",
        "enunciado": "PQ-080 — Modelagem (Modelagem)\n\nContexto de negócio: Cenário real de Comercial: normalizar tabela larga com unpivot.\n\nProblema: Normalizar tabela larga com Unpivot\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.UnpivotOtherColumns(Fonte, {\"Produto\"}, \"Indicador\", \"Valor\")",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.UnpivotOtherColumns(Fonte, {\"Produto\"}, \"Indicador\", \"Valor\")",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-081",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "Comercial",
        "enunciado": "PQ-081 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de Comercial: padronizar produto com text.upper/trim/clean.\n\nProblema: Padronizar produto com Text.Upper/Trim/Clean\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Produto_Limpo\", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Produto_Limpo\", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-082",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "RH",
        "enunciado": "PQ-082 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de RH: remover pontuação de documento.\n\nProblema: Remover pontuação de documento\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.TransformColumns(Fonte, {{\"CPF\", each Text.Select(_, {\"0\"..\"9\"}), type text}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.TransformColumns(Fonte, {{\"CPF\", each Text.Select(_, {\"0\"..\"9\"}), type text}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de RH, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-083",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "Marketing",
        "enunciado": "PQ-083 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de Marketing: criar chave de campanha com text.combine.\n\nProblema: Criar chave de campanha com Text.Combine\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Chave_Campanha\", each Text.Combine({[Canal], [Campanha], [Mes]}, \"|\"), type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Chave_Campanha\", each Text.Combine({[Canal], [Campanha], [Mes]}, \"|\"), type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Marketing, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-084",
        "ferramenta": "Power Query",
        "tema": "Tipos",
        "nivel": "ETL",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-084 — Tipos (ETL)\n\nContexto de negócio: Cenário real de Financeiro: converter valor com try.\n\nProblema: Converter valor com try\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Valor_Numero\", each try Number.From([Valor]) otherwise null, type number)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Valor_Numero\", each try Number.From([Valor]) otherwise null, type number)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-085",
        "ferramenta": "Power Query",
        "tema": "Datas",
        "nivel": "Fundamentos",
        "area_negocio": "Logística",
        "enunciado": "PQ-085 — Datas (Fundamentos)\n\nContexto de negócio: Cenário real de Logística: calcular lead time em dias.\n\nProblema: Calcular lead time em dias\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"LeadTime\", each try Duration.Days([Entrega] - [Pedido]) otherwise null, Int64.Type)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"LeadTime\", each try Duration.Days([Entrega] - [Pedido]) otherwise null, Int64.Type)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Logística, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-086",
        "ferramenta": "Power Query",
        "tema": "Datas",
        "nivel": "Fundamentos",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-086 — Datas (Fundamentos)\n\nContexto de negócio: Cenário real de Financeiro: criar competência mensal.\n\nProblema: Criar competência mensal\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Competencia\", each Date.StartOfMonth(Date.From([Data])), type date)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Competencia\", each Date.StartOfMonth(Date.From([Data])), type date)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-087",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-087 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: selecionar colunas obrigatórias.\n\nProblema: Selecionar colunas obrigatórias\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.SelectColumns(Fonte, {\"Produto\", \"Data\", \"Valor\"}, MissingField.UseNull)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.SelectColumns(Fonte, {\"Produto\", \"Data\", \"Valor\"}, MissingField.UseNull)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-088",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-088 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: renomear colunas de sistema.\n\nProblema: Renomear colunas de sistema\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.RenameColumns(Fonte, {{\"Vlr\", \"Valor\"}, {\"Dt\", \"Data\"}}, MissingField.Ignore)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.RenameColumns(Fonte, {{\"Vlr\", \"Valor\"}, {\"Dt\", \"Data\"}}, MissingField.Ignore)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-089",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-089 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: adicionar índice depois de ordenar.\n\nProblema: Adicionar índice depois de ordenar\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddIndexColumn(Table.Sort(Fonte, {{\"Data\", Order.Ascending}}), \"Indice\", 1, 1, Int64.Type)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddIndexColumn(Table.Sort(Fonte, {{\"Data\", Order.Ascending}}), \"Indice\", 1, 1, Int64.Type)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-090",
        "ferramenta": "Power Query",
        "tema": "Agrupamento",
        "nivel": "ETL",
        "area_negocio": "Comercial",
        "enunciado": "PQ-090 — Agrupamento (ETL)\n\nContexto de negócio: Cenário real de Comercial: agrupar vendas por produto.\n\nProblema: Agrupar vendas por produto\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Group(Fonte, {\"Produto\"}, {{\"Total\", each List.Sum([Valor]), type number}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Group(Fonte, {\"Produto\"}, {{\"Total\", each List.Sum([Valor]), type number}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-091",
        "ferramenta": "Power Query",
        "tema": "Agrupamento",
        "nivel": "ETL",
        "area_negocio": "RH",
        "enunciado": "PQ-091 — Agrupamento (ETL)\n\nContexto de negócio: Cenário real de RH: agrupar headcount por área.\n\nProblema: Agrupar headcount por área\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Group(Fonte, {\"Area\"}, {{\"Headcount\", each Table.RowCount(_), Int64.Type}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Group(Fonte, {\"Area\"}, {{\"Headcount\", each Table.RowCount(_), Int64.Type}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de RH, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-092",
        "ferramenta": "Power Query",
        "tema": "Merge",
        "nivel": "ETL",
        "area_negocio": "Supply Chain",
        "enunciado": "PQ-092 — Merge (ETL)\n\nContexto de negócio: Cenário real de Supply Chain: mesclar fato com dimensão.\n\nProblema: Mesclar fato com dimensão\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.NestedJoin(Fato, {\"SKU\"}, DimProduto, {\"SKU\"}, \"DimProduto\", JoinKind.LeftOuter)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.NestedJoin(Fato, {\"SKU\"}, DimProduto, {\"SKU\"}, \"DimProduto\", JoinKind.LeftOuter)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Supply Chain, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-093",
        "ferramenta": "Power Query",
        "tema": "Merge",
        "nivel": "Data Quality",
        "area_negocio": "Supply Chain",
        "enunciado": "PQ-093 — Merge (Data Quality)\n\nContexto de negócio: Cenário real de Supply Chain: identificar produtos sem cadastro após merge.\n\nProblema: Identificar produtos sem cadastro após merge\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.SelectRows(Expandido, each [Categoria] = null)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.SelectRows(Expandido, each [Categoria] = null)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Supply Chain, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-094",
        "ferramenta": "Power Query",
        "tema": "Append",
        "nivel": "ETL",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-094 — Append (ETL)\n\nContexto de negócio: Cenário real de Financeiro: combinar bases mensais.\n\nProblema: Combinar bases mensais\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Combine({BaseJan, BaseFev, BaseMar})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Combine({BaseJan, BaseFev, BaseMar})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-095",
        "ferramenta": "Power Query",
        "tema": "Data Quality",
        "nivel": "Data Quality",
        "area_negocio": "Dados",
        "enunciado": "PQ-095 — Data Quality (Data Quality)\n\nContexto de negócio: Cenário real de Dados: criar status_dq para valor nulo.\n\nProblema: Criar Status_DQ para valor nulo\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Status_DQ\", each if [Valor] = null then \"ERRO VALOR\" else \"OK\", type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Status_DQ\", each if [Valor] = null then \"ERRO VALOR\" else \"OK\", type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-096",
        "ferramenta": "Power Query",
        "tema": "Data Quality",
        "nivel": "Data Quality",
        "area_negocio": "Dados",
        "enunciado": "PQ-096 — Data Quality (Data Quality)\n\nContexto de negócio: Cenário real de Dados: gerar perfil com table.profile.\n\nProblema: Gerar perfil com Table.Profile\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Profile(Fonte)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Profile(Fonte)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-097",
        "ferramenta": "Power Query",
        "tema": "Performance",
        "nivel": "Performance",
        "area_negocio": "Dados",
        "enunciado": "PQ-097 — Performance (Performance)\n\nContexto de negócio: Cenário real de Dados: aplicar table.buffer em dimensão.\n\nProblema: Aplicar Table.Buffer em dimensão\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Buffer(DimProduto)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Buffer(DimProduto)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-098",
        "ferramenta": "Power Query",
        "tema": "Listas",
        "nivel": "Avançado",
        "area_negocio": "Dados",
        "enunciado": "PQ-098 — Listas (Avançado)\n\nContexto de negócio: Cenário real de Dados: aplicar list.accumulate para substituições.\n\nProblema: Aplicar List.Accumulate para substituições\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "List.Accumulate(Substituicoes, TextoInicial, (estado, atual) => Text.Replace(estado, atual{0}, atual{1}))",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-099",
        "ferramenta": "Power Query",
        "tema": "Records",
        "nivel": "Avançado",
        "area_negocio": "Dados",
        "enunciado": "PQ-099 — Records (Avançado)\n\nContexto de negócio: Cenário real de Dados: ler campo de record com default.\n\nProblema: Ler campo de record com default\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Record.FieldOrDefault([Registro], \"Campo\", null)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Record.FieldOrDefault([Registro], \"Campo\", null)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-100",
        "ferramenta": "Power Query",
        "tema": "Modelagem",
        "nivel": "Modelagem",
        "area_negocio": "Comercial",
        "enunciado": "PQ-100 — Modelagem (Modelagem)\n\nContexto de negócio: Cenário real de Comercial: normalizar tabela larga com unpivot.\n\nProblema: Normalizar tabela larga com Unpivot\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.UnpivotOtherColumns(Fonte, {\"Produto\"}, \"Indicador\", \"Valor\")",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.UnpivotOtherColumns(Fonte, {\"Produto\"}, \"Indicador\", \"Valor\")",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-101",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "Comercial",
        "enunciado": "PQ-101 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de Comercial: padronizar produto com text.upper/trim/clean.\n\nProblema: Padronizar produto com Text.Upper/Trim/Clean\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Produto_Limpo\", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Produto_Limpo\", each Text.Upper(Text.Trim(Text.Clean([Produto]))), type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-102",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "RH",
        "enunciado": "PQ-102 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de RH: remover pontuação de documento.\n\nProblema: Remover pontuação de documento\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.TransformColumns(Fonte, {{\"CPF\", each Text.Select(_, {\"0\"..\"9\"}), type text}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.TransformColumns(Fonte, {{\"CPF\", each Text.Select(_, {\"0\"..\"9\"}), type text}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de RH, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-103",
        "ferramenta": "Power Query",
        "tema": "Texto",
        "nivel": "Fundamentos",
        "area_negocio": "Marketing",
        "enunciado": "PQ-103 — Texto (Fundamentos)\n\nContexto de negócio: Cenário real de Marketing: criar chave de campanha com text.combine.\n\nProblema: Criar chave de campanha com Text.Combine\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Chave_Campanha\", each Text.Combine({[Canal], [Campanha], [Mes]}, \"|\"), type text)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Chave_Campanha\", each Text.Combine({[Canal], [Campanha], [Mes]}, \"|\"), type text)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Marketing, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-104",
        "ferramenta": "Power Query",
        "tema": "Tipos",
        "nivel": "ETL",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-104 — Tipos (ETL)\n\nContexto de negócio: Cenário real de Financeiro: converter valor com try.\n\nProblema: Converter valor com try\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Valor_Numero\", each try Number.From([Valor]) otherwise null, type number)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Valor_Numero\", each try Number.From([Valor]) otherwise null, type number)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-105",
        "ferramenta": "Power Query",
        "tema": "Datas",
        "nivel": "Fundamentos",
        "area_negocio": "Logística",
        "enunciado": "PQ-105 — Datas (Fundamentos)\n\nContexto de negócio: Cenário real de Logística: calcular lead time em dias.\n\nProblema: Calcular lead time em dias\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"LeadTime\", each try Duration.Days([Entrega] - [Pedido]) otherwise null, Int64.Type)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"LeadTime\", each try Duration.Days([Entrega] - [Pedido]) otherwise null, Int64.Type)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Logística, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-106",
        "ferramenta": "Power Query",
        "tema": "Datas",
        "nivel": "Fundamentos",
        "area_negocio": "Financeiro",
        "enunciado": "PQ-106 — Datas (Fundamentos)\n\nContexto de negócio: Cenário real de Financeiro: criar competência mensal.\n\nProblema: Criar competência mensal\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddColumn(Fonte, \"Competencia\", each Date.StartOfMonth(Date.From([Data])), type date)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddColumn(Fonte, \"Competencia\", each Date.StartOfMonth(Date.From([Data])), type date)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Financeiro, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-107",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-107 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: selecionar colunas obrigatórias.\n\nProblema: Selecionar colunas obrigatórias\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.SelectColumns(Fonte, {\"Produto\", \"Data\", \"Valor\"}, MissingField.UseNull)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.SelectColumns(Fonte, {\"Produto\", \"Data\", \"Valor\"}, MissingField.UseNull)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-108",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-108 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: renomear colunas de sistema.\n\nProblema: Renomear colunas de sistema\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.RenameColumns(Fonte, {{\"Vlr\", \"Valor\"}, {\"Dt\", \"Data\"}}, MissingField.Ignore)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.RenameColumns(Fonte, {{\"Vlr\", \"Valor\"}, {\"Dt\", \"Data\"}}, MissingField.Ignore)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-109",
        "ferramenta": "Power Query",
        "tema": "Tabelas",
        "nivel": "ETL",
        "area_negocio": "Dados",
        "enunciado": "PQ-109 — Tabelas (ETL)\n\nContexto de negócio: Cenário real de Dados: adicionar índice depois de ordenar.\n\nProblema: Adicionar índice depois de ordenar\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.AddIndexColumn(Table.Sort(Fonte, {{\"Data\", Order.Ascending}}), \"Indice\", 1, 1, Int64.Type)",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.AddIndexColumn(Table.Sort(Fonte, {{\"Data\", Order.Ascending}}), \"Indice\", 1, 1, Int64.Type)",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Dados, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    },
    {
        "id": "PQ-110",
        "ferramenta": "Power Query",
        "tema": "Agrupamento",
        "nivel": "ETL",
        "area_negocio": "Comercial",
        "enunciado": "PQ-110 — Agrupamento (ETL)\n\nContexto de negócio: Cenário real de Comercial: agrupar vendas por produto.\n\nProblema: Agrupar vendas por produto\n\nBase simulada: Base simulada com colunas compatíveis com o tema, contendo valores válidos, nulos, grafias divergentes e registros para auditoria.\n\nTarefa: resolva usando Power Query, documentando o raciocínio, a regra aplicada e o resultado esperado.",
        "analise": "Identificar origem do dado, tipo esperado, regra de transformação, risco de perda de rastreabilidade e impacto no modelo final.",
        "raciocinio": "Aplicar etapa M auditável, preservar dado original quando houver limpeza, registrar Status_DQ quando houver exceção e evitar regra invisível.",
        "solucao_principal": "Table.Group(Fonte, {\"Produto\"}, {{\"Total\", each List.Sum([Valor]), type number}})",
        "solucao_excel": "Pode ser resolvido no Excel se for pontual; se for recorrente, Power Query é preferível.",
        "solucao_powerquery": "Table.Group(Fonte, {\"Produto\"}, {{\"Total\", each List.Sum([Valor]), type number}})",
        "passo_a_passo": [
            "1. Entender o contexto de negócio e identificar a granularidade do dado.",
            "2. Validar campos de entrada, valores vazios, tipos de dados e possíveis inconsistências.",
            "3. Aplicar a solução principal proposta para o exercício.",
            "4. Conferir o resultado em uma amostra pequena antes de aplicar ao conjunto completo.",
            "5. Documentar erro comum, alternativa técnica e aplicação prática."
        ],
        "resultado_esperado": "Tabela transformada conforme regra, com colunas limpas, tipos corretos, exceções rastreáveis e sem quebra de consulta.",
        "erros_comuns": "Erro comum: remover registros sem criar tabela de exceções, converter tipo sem try/otherwise ou sobrescrever coluna original.",
        "alternativas": "Alternativas: tabela De/Para, merge, Table.Profile, Table.Schema, Excel pontual, DAX no modelo ou VBA para automação.",
        "aplicacao_pratica": "Rotina de Comercial, preparação para teste técnico e automação de bases recorrentes.",
        "criterio_correcao": "A resposta deve conter fórmula/código correto, explicação do raciocínio, validação do resultado e observação de risco ou erro comum."
    }
]


def get_exercicios_resolvidos() -> List[Dict[str, Any]]:
    return EXERCICIOS_RESOLVIDOS


def get_exercicios_por_ferramenta(ferramenta: str) -> List[Dict[str, Any]]:
    return [item for item in EXERCICIOS_RESOLVIDOS if item.get("ferramenta") == ferramenta]


def export_exercicios_resolvidos_payload() -> Dict[str, Any]:
    return {
        "block_name": BLOCK_NAME,
        "version": APP_VERSION,
        "status": BLOCK_STATUS,
        "build_date": BUILD_DATE,
        "streamlit_available": STREAMLIT_AVAILABLE,
        "exercicios_resolvidos": EXERCICIOS_RESOLVIDOS,
    }


def validate_exercicios_resolvidos() -> List[Dict[str, Any]]:
    excel = get_exercicios_por_ferramenta("Excel")
    powerquery = get_exercicios_por_ferramenta("Power Query")

    checks = [
        {
            "regra": "Versão V14 definida",
            "status": "OK" if APP_VERSION.startswith("v14") else "FALHA",
            "evidencia": APP_VERSION,
            "criticidade": "Crítica",
        },
        {
            "regra": "Exercícios Excel >= mínimo",
            "status": "OK" if len(excel) >= MIN_EXCEL_EXERCISES else "FALHA",
            "evidencia": str(len(excel)),
            "criticidade": "Crítica",
        },
        {
            "regra": "Exercícios Power Query >= mínimo",
            "status": "OK" if len(powerquery) >= MIN_POWERQUERY_EXERCISES else "FALHA",
            "evidencia": str(len(powerquery)),
            "criticidade": "Crítica",
        },
        {
            "regra": "Total de exercícios >= mínimo",
            "status": "OK" if len(EXERCICIOS_RESOLVIDOS) >= MIN_TOTAL_EXERCISES else "FALHA",
            "evidencia": str(len(EXERCICIOS_RESOLVIDOS)),
            "criticidade": "Crítica",
        },
        {
            "regra": "Todos possuem enunciado",
            "status": "OK" if all(item.get("enunciado") for item in EXERCICIOS_RESOLVIDOS) else "FALHA",
            "evidencia": "enunciado preenchido",
            "criticidade": "Alta",
        },
        {
            "regra": "Todos possuem solução principal",
            "status": "OK" if all(item.get("solucao_principal") for item in EXERCICIOS_RESOLVIDOS) else "FALHA",
            "evidencia": "solucao_principal preenchida",
            "criticidade": "Alta",
        },
        {
            "regra": "Todos possuem passo a passo",
            "status": "OK" if all(item.get("passo_a_passo") for item in EXERCICIOS_RESOLVIDOS) else "FALHA",
            "evidencia": "passo_a_passo preenchido",
            "criticidade": "Alta",
        },
        {
            "regra": "Todos possuem erros comuns",
            "status": "OK" if all(item.get("erros_comuns") for item in EXERCICIOS_RESOLVIDOS) else "FALHA",
            "evidencia": "erros_comuns preenchido",
            "criticidade": "Média",
        },
        {
            "regra": "Streamlit opcional",
            "status": "OK",
            "evidencia": "Import tratado com ModuleNotFoundError",
            "criticidade": "Crítica",
        },
    ]
    return checks


def print_validation_report() -> None:
    checks = validate_exercicios_resolvidos()
    falhas = sum(1 for item in checks if item["status"] == "FALHA")
    excel = get_exercicios_por_ferramenta("Excel")
    powerquery = get_exercicios_por_ferramenta("Power Query")
    temas = sorted(set(item.get("tema", "") for item in EXERCICIOS_RESOLVIDOS if item.get("tema")))
    niveis = sorted(set(item.get("nivel", "") for item in EXERCICIOS_RESOLVIDOS if item.get("nivel")))

    print("=" * 78)
    print("VALIDAÇÃO DO BLOCO 08 — EXERCÍCIOS RESOLVIDOS EXCEL + POWER QUERY V14")
    print("=" * 78)
    print(f"Bloco: {BLOCK_NAME}")
    print(f"Versão: {APP_VERSION}")
    print(f"Status: {BLOCK_STATUS}")
    print(f"Streamlit disponível: {STREAMLIT_AVAILABLE}")
    print(f"Exercícios Excel: {len(excel)}")
    print(f"Exercícios Power Query: {len(powerquery)}")
    print(f"Total exercícios resolvidos: {len(EXERCICIOS_RESOLVIDOS)}")
    print(f"Temas: {len(temas)}")
    print(f"Níveis: {len(niveis)}")
    print(f"Falhas core: {falhas}")
    print(f"Aprovado core: {falhas == 0}")
    print("-" * 78)
    print("CHECKS CORE:")
    for item in checks:
        print(f"[{item['status']}] {item['regra']} — {item['evidencia']}")
    print("-" * 78)
    if falhas == 0:
        print("RESULTADO: BLOCO 08 APROVADO PARA ANÁLISE DO USUÁRIO.")
        print("Este bloco já contém enunciado e solução para cada exercício de Excel e Power Query.")
    else:
        print("RESULTADO: BLOCO 08 REPROVADO. CORRIGIR ANTES DE AVANÇAR.")
    print("=" * 78)


def _to_dataframe(rows: List[Dict[str, Any]]):
    if pd is None:
        return rows
    return pd.DataFrame(rows)


def render_exercicios_resolvidos_app() -> None:
    if not STREAMLIT_AVAILABLE:
        raise RuntimeError("Streamlit não instalado. Use print_validation_report() para validar no Colab.")

    st.set_page_config(page_title="Exercícios Resolvidos V14", page_icon="🧠", layout="wide")
    st.title("Bloco 08 — Exercícios Resolvidos Excel + Power Query V14")
    st.caption("Enunciado · Solução · Passo a passo · Erros comuns · Aplicação prática")

    with st.sidebar:
        ferramenta = st.selectbox("Ferramenta", ["Todas", "Excel", "Power Query"])
        nivel = st.selectbox("Nível", ["Todos"] + sorted(set(item.get("nivel", "") for item in EXERCICIOS_RESOLVIDOS)))
        tema = st.selectbox("Tema", ["Todos"] + sorted(set(item.get("tema", "") for item in EXERCICIOS_RESOLVIDOS)))
        busca = st.text_input("Buscar", "")

    dados = EXERCICIOS_RESOLVIDOS
    if ferramenta != "Todas":
        dados = [item for item in dados if item.get("ferramenta") == ferramenta]
    if nivel != "Todos":
        dados = [item for item in dados if item.get("nivel") == nivel]
    if tema != "Todos":
        dados = [item for item in dados if item.get("tema") == tema]
    if busca:
        termo = busca.lower()
        dados = [item for item in dados if termo in json.dumps(item, ensure_ascii=False).lower()]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total filtrado", len(dados))
    c2.metric("Excel total", len(get_exercicios_por_ferramenta("Excel")))
    c3.metric("Power Query total", len(get_exercicios_por_ferramenta("Power Query")))
    c4.metric("Falhas auditoria", sum(1 for item in validate_exercicios_resolvidos() if item["status"] == "FALHA"))

    st.divider()

    if pd is not None:
        resumo = pd.DataFrame(dados)
        cols = [c for c in ["id", "ferramenta", "tema", "nivel", "area_negocio", "resultado_esperado"] if c in resumo.columns]
        st.dataframe(resumo[cols], use_container_width=True, hide_index=True, height=320)
    else:
        st.write(dados[:20])

    st.divider()
    st.subheader("Cards dos exercícios")

    for item in dados[:80]:
        titulo = f"{item.get('id')} · {item.get('ferramenta')} · {item.get('tema')} · {item.get('nivel')}"
        with st.expander(titulo, expanded=False):
            st.markdown("### Enunciado")
            st.text(item.get("enunciado", ""))

            st.markdown("### Análise")
            st.write(item.get("analise", ""))

            st.markdown("### Raciocínio")
            st.write(item.get("raciocinio", ""))

            st.markdown("### Solução principal")
            language = "powerquery" if item.get("ferramenta") == "Power Query" else "text"
            st.code(item.get("solucao_principal", ""), language=language)

            st.markdown("### Passo a passo")
            for step in item.get("passo_a_passo", []):
                st.write(step)

            st.markdown("### Resultado esperado")
            st.write(item.get("resultado_esperado", ""))

            st.markdown("### Erros comuns")
            st.write(item.get("erros_comuns", ""))

            st.markdown("### Alternativas")
            st.write(item.get("alternativas", ""))

            st.markdown("### Aplicação prática")
            st.write(item.get("aplicacao_pratica", ""))

    if len(dados) > 80:
        st.info("Exibindo os 80 primeiros cards filtrados para preservar performance. Use filtros ou busca para refinar.")


if __name__ == "__main__":
    if STREAMLIT_AVAILABLE:
        try:
            render_exercicios_resolvidos_app()
        except Exception as exc:
            print_validation_report()
            print(f"AVISO: UI Streamlit não renderizada. Detalhe: {exc}")
    else:
        print_validation_report()



# =============================================================================
# INÍCIO DO BLOCO 09 — EXCEL DECISION FRAMEWORK V14
# =============================================================================


# BLOCO 09 — EXCEL DECISION FRAMEWORK V14
# Objetivo:
# - Considerar apenas Excel.
# - Incluir mapa mental de tomada de decisão.
# - Relacionar exercícios Excel com tipos de problema do mapa mental.
# - Trocar a visão Power Query por estratégias Excel.
# - Ensinar o caminho desde a leitura do enunciado até a solução.
#
# Como usar:
# 1. Cole este bloco depois do Bloco 08, onde já existe EXERCICIOS_RESOLVIDOS.
# 2. Rode este bloco.
# 3. Para Streamlit, chame render_excel_decision_framework_app().
#
# Observação:
# Este bloco não remove conteúdo anterior. Ele cria uma camada pedagógica nova.

import json
import datetime as dt
from typing import Any, Dict, List

try:
    import pandas as pd
except ModuleNotFoundError:
    pd = None

try:
    import streamlit as st
    STREAMLIT_AVAILABLE = True
except ModuleNotFoundError:
    st = None
    STREAMLIT_AVAILABLE = False


APP_VERSION = "v14.4.0"
BLOCK_NAME = "bloco_09_excel_decision_framework"
BLOCK_STATUS = "COMPLETO_CORE_PENDENTE_USUARIO"
BUILD_DATE = dt.date.today().isoformat()

MIN_PROBLEM_TYPES = 10
MIN_EXCEL_EXERCISES = 100
MIN_EXERCISES_WITH_MAPPING = 100


# =============================================================================
# 1. COMITÊ EXPANDIDO PARA ESTA FRENTE
# =============================================================================

COMITE_EXCEL_DECISAO = [
    {
        "nucleo": "Técnico",
        "papel": "Excel Architect",
        "responsabilidade": "Garantir que as fórmulas estejam corretas em português e aplicáveis a testes técnicos.",
    },
    {
        "nucleo": "Pedagógico",
        "papel": "Especialista em Testes Técnicos de Dados",
        "responsabilidade": "Mapear os tipos de exercícios mais frequentes em processos seletivos.",
    },
    {
        "nucleo": "Pedagógico",
        "papel": "Instrutor Excel Corporativo",
        "responsabilidade": "Traduzir fórmulas em raciocínio prático e evitar complexidade desnecessária.",
    },
    {
        "nucleo": "Pedagógico",
        "papel": "Designer Instrucional",
        "responsabilidade": "Criar mapa mental, árvore de decisão e progressão de aprendizado.",
    },
    {
        "nucleo": "Qualidade",
        "papel": "Auditor Independente",
        "responsabilidade": "Validar que o bloco não remove conteúdo e que todos os exercícios Excel ganham estratégia.",
    },
]


# =============================================================================
# 2. TIPOS DE PROBLEMA — MAPA MENTAL
# =============================================================================

PROBLEM_TYPES: List[Dict[str, Any]] = [
    {
        "id": "TIPO-01",
        "tipo_problema": "Limpeza e padronização de texto",
        "pergunta_diagnostico": "O enunciado pede corrigir, limpar, padronizar, normalizar ou comparar textos?",
        "palavras_chave": ["limpar", "padronizar", "corrigir", "normalizar", "cadastro", "texto", "nome", "produto", "descrição", "acentos", "espaços", "cpf", "e-mail", "email"],
        "quando_usar": "Quando a entrada é texto inconsistente: espaços, caixa, pontuação, acentos, abreviações ou padrões divergentes.",
        "quando_nao_usar": "Quando a questão exige cálculo, agregação ou busca em tabela de apoio.",
        "formulas_principais": ["ARRUMAR", "MAIÚSCULA", "MINÚSCULA", "PRI.MAIÚSCULA", "SUBSTITUIR", "EXT.TEXTO", "ESQUERDA", "DIREITA", "TEXTOJUNTAR"],
        "combinacoes_estrategicas": [
            '=MAIÚSCULA(ARRUMAR(A2))',
            '=SUBSTITUIR(SUBSTITUIR(A2;".";"");"-";"")',
            '=TEXTOJUNTAR("-";VERDADEIRO;A2:C2)',
            '=SE(MAIÚSCULA(ARRUMAR(A2))=MAIÚSCULA(ARRUMAR(B2));"OK";"DIVERGENTE")',
        ],
        "estrategia_solucao": [
            "Identificar se o problema é de texto e não de cálculo.",
            "Verificar espaços, caixa, pontuação, acentos ou padrões misturados.",
            "Limpar antes de comparar ou concatenar.",
            "Usar coluna auxiliar se a fórmula ficar muito longa.",
            "Validar com vazio, acento, espaço e texto já correto.",
        ],
        "erro_comum": "Comparar A2=B2 sem padronizar os dois lados.",
        "nivel_cobranca": "Muito frequente em testes júnior e pleno.",
        "peso_teste": 5,
    },
    {
        "id": "TIPO-02",
        "tipo_problema": "Busca e enriquecimento de cadastro",
        "pergunta_diagnostico": "O enunciado pede retornar uma informação a partir de uma chave?",
        "palavras_chave": ["buscar", "retornar", "procurar", "localizar", "encontrar", "consultar", "categoria", "cadastro", "sku", "cliente", "conta", "tabela de apoio", "procv", "procx"],
        "quando_usar": "Quando existe uma base principal e uma tabela auxiliar de cadastro.",
        "quando_nao_usar": "Quando a saída é soma, contagem, média ou lista filtrada de várias linhas.",
        "formulas_principais": ["PROCX", "PROCV", "ÍNDICE", "CORRESP", "SEERRO"],
        "combinacoes_estrategicas": [
            '=PROCX(A2;Cadastro[Chave];Cadastro[Retorno];"Sem cadastro")',
            '=SEERRO(PROCV(A2;Cadastro!A:D;4;FALSO);"Sem cadastro")',
            '=ÍNDICE(Tabela[Valor];CORRESP(A2;Tabela[Chave];0))',
            '=PROCX(1;(Tabela[Produto]=A2)*(Tabela[Canal]=B2);Tabela[Preço];"N/D")',
        ],
        "estrategia_solucao": [
            "Identificar a chave de busca.",
            "Identificar a tabela onde a chave existe.",
            "Identificar a coluna de retorno.",
            "Prever o que acontece quando a chave não existir.",
            "Validar duplicidade de chave antes de confiar no primeiro retorno.",
        ],
        "erro_comum": "Usar PROCV sem tratar item não encontrado ou depender de coluna fixa.",
        "nivel_cobranca": "Muito frequente em testes de analista de dados.",
        "peso_teste": 5,
    },
    {
        "id": "TIPO-03",
        "tipo_problema": "Agregação por critérios",
        "pergunta_diagnostico": "O enunciado pede total, quantidade, média, receita, gasto, volume ou KPI por condição?",
        "palavras_chave": ["total", "somar", "soma", "quantidade", "contar", "média", "media", "ticket", "receita", "despesa", "valor", "por produto", "por área", "por canal", "por mês", "somases", "cont.ses", "médiases"],
        "quando_usar": "Quando a resposta é um número consolidado com um ou mais critérios.",
        "quando_nao_usar": "Quando a resposta precisa listar linhas completas ou retornar dado cadastral.",
        "formulas_principais": ["SOMASE", "SOMASES", "CONT.SE", "CONT.SES", "MÉDIASE", "MÉDIASES", "SUBTOTAL", "AGREGAR"],
        "combinacoes_estrategicas": [
            '=SOMASES(Base[Valor];Base[Produto];A2;Base[Canal];B2)',
            '=CONT.SES(Base[Área];A2;Base[Status];"Ativo")',
            '=MÉDIASES(Base[ROI];Base[Canal];A2)',
            '=SUBTOTAL(9;Base[Valor])',
        ],
        "estrategia_solucao": [
            "Separar campo calculado dos campos de critério.",
            "Identificar se a pergunta é soma, contagem ou média.",
            "Mapear todos os critérios do enunciado.",
            "Conferir se intervalos têm o mesmo tamanho.",
            "Validar com uma amostra filtrada manualmente.",
        ],
        "erro_comum": "Inverter intervalo de soma com intervalo de critério em SOMASES.",
        "nivel_cobranca": "O tipo mais recorrente em testes de Excel para analista.",
        "peso_teste": 5,
    },
    {
        "id": "TIPO-04",
        "tipo_problema": "Classificação, status e regra de negócio",
        "pergunta_diagnostico": "O enunciado pede criar status, faixa, categoria, prioridade ou classificação?",
        "palavras_chave": ["status", "classificar", "faixa", "categoria", "prioridade", "crítico", "critico", "normal", "alto", "baixo", "se ", "regra", "ses", "ativo", "vencido"],
        "quando_usar": "Quando a resposta é um rótulo baseado em condições.",
        "quando_nao_usar": "Quando a regra depende de tabela de faixas externa; nesse caso, considere PROCX por aproximação.",
        "formulas_principais": ["SE", "SES", "E", "OU", "SEERRO", "LET"],
        "combinacoes_estrategicas": [
            '=SE(B2>=1000;"Alta";SE(B2>=500;"Média";"Baixa"))',
            '=SES(B2>=1000;"Alta";B2>=500;"Média";VERDADEIRO;"Baixa")',
            '=SE(E(B2>=1000;C2="Ativo");"Prioritário";"Normal")',
            '=LET(valor;B2;limite;1000;SE(valor>=limite;"OK";"Validar"))',
        ],
        "estrategia_solucao": [
            "Listar todas as regras antes da fórmula.",
            "Ordenar faixas do maior para o menor quando houver limites.",
            "Definir saída padrão para casos não previstos.",
            "Usar LET quando houver muitas variáveis.",
            "Testar valores na fronteira: 500, 1000, vazio e erro.",
        ],
        "erro_comum": "Começar pela menor faixa e classificar tudo errado.",
        "nivel_cobranca": "Muito frequente em júnior, pleno e testes de negócio.",
        "peso_teste": 5,
    },
    {
        "id": "TIPO-05",
        "tipo_problema": "Listas, filtros e matrizes dinâmicas",
        "pergunta_diagnostico": "O enunciado pede listar registros, criar lista única, ordenar, top N ou retornar uma tabela filtrada?",
        "palavras_chave": ["listar", "filtrar", "mostrar", "único", "unico", "distinto", "classificar", "ordenar", "top", "ranking", "retornar registros", "filtro"],
        "quando_usar": "Quando a saída esperada é uma matriz ou lista dinâmica, não uma única célula.",
        "quando_nao_usar": "Quando a resposta é apenas um indicador agregado ou uma busca pontual.",
        "formulas_principais": ["FILTRO", "ÚNICO", "CLASSIFICAR", "PEGAR", "ESCOLHERCOLS", "EMPILHARV", "EMPILHARH"],
        "combinacoes_estrategicas": [
            '=FILTRO(Base;Base[Valor]>1000;"Sem registros")',
            '=CLASSIFICAR(ÚNICO(Base[Vendedor]))',
            '=PEGAR(CLASSIFICAR(Base;3;-1);10)',
            '=ESCOLHERCOLS(PEGAR(CLASSIFICAR(Base;3;-1);10);1;2;3)',
        ],
        "estrategia_solucao": [
            "Confirmar se a entrega precisa derramar várias linhas.",
            "Definir critério do filtro ou coluna de ordenação.",
            "Usar mensagem para resultado vazio.",
            "Reduzir colunas se a saída ficar poluída.",
            "Testar se há espaço livre para a matriz derramar.",
        ],
        "erro_comum": "Usar PROCX quando a pergunta pede várias linhas.",
        "nivel_cobranca": "Frequente em Excel 365 e testes atuais.",
        "peso_teste": 4,
    },
    {
        "id": "TIPO-06",
        "tipo_problema": "Datas, prazos e competência",
        "pergunta_diagnostico": "O enunciado pede calcular prazo, atraso, competência, vencimento ou período?",
        "palavras_chave": ["data", "prazo", "vencimento", "competência", "competencia", "mês", "mes", "ano", "dias", "sla", "atraso", "fechamento", "hoje"],
        "quando_usar": "Quando a regra depende de datas reais e períodos.",
        "quando_nao_usar": "Quando a data está como texto não tratado; primeiro converta ou corrija o tipo.",
        "formulas_principais": ["HOJE", "AGORA", "DATA", "ANO", "MÊS", "DIAS", "FIMMÊS", "DIA.DA.SEMANA"],
        "combinacoes_estrategicas": [
            '=DATA(ANO(A2);MÊS(A2);1)',
            '=DIAS(B2;A2)',
            '=SE(HOJE()>A2;"Vencido";"No prazo")',
            '=FIMMÊS(A2;0)',
        ],
        "estrategia_solucao": [
            "Verificar se a data é data real ou texto.",
            "Identificar se é diferença, status ou agrupamento mensal.",
            "Evitar transformar competência em texto quando precisar ordenar.",
            "Tratar datas vazias.",
            "Validar com datas antes, depois e no mesmo dia.",
        ],
        "erro_comum": "Criar competência como texto e perder ordenação temporal.",
        "nivel_cobranca": "Frequente em financeiro, logística e RH.",
        "peso_teste": 4,
    },
    {
        "id": "TIPO-07",
        "tipo_problema": "Coringas e busca parcial",
        "pergunta_diagnostico": "O enunciado pede encontrar texto parcial, padrão flexível ou caractere literal como *, ? ou ~?",
        "palavras_chave": ["contém", "contem", "começa", "termina", "parcial", "coringa", "asterisco", "interrogação", "literal", "padrão", "*", "?", "~"],
        "quando_usar": "Quando a correspondência não é exata ou o texto segue padrão variável.",
        "quando_nao_usar": "Quando a chave precisa ser exata para evitar falso positivo.",
        "formulas_principais": ["CONT.SE", "CONT.SES", "PROCX", "PROCURAR", "ÉNÚM", "FILTRO"],
        "combinacoes_estrategicas": [
            '=CONT.SE(A:A;"*LATA*")',
            '=CONT.SE(A:A;"???")',
            '=CONT.SE(A:A;"SKU~*")',
            '=PROCX("*"&A2&"*";Base[Produto];Base[Categoria];"N/D";2)',
        ],
        "estrategia_solucao": [
            "Definir se o padrão é contém, começa, termina ou tamanho fixo.",
            "Usar * para vários caracteres.",
            "Usar ? para um caractere.",
            "Usar ~ para buscar * ou ? literalmente.",
            "Validar falso positivo em descrições semelhantes.",
        ],
        "erro_comum": "Usar * quando deveria exigir correspondência exata.",
        "nivel_cobranca": "Médio, mas diferencia candidatos atentos a detalhe.",
        "peso_teste": 3,
    },
    {
        "id": "TIPO-08",
        "tipo_problema": "Auditoria, validação e tratamento de erro",
        "pergunta_diagnostico": "O enunciado pede validar inconsistência, divergência, incompleto, erro ou rastreabilidade?",
        "palavras_chave": ["auditoria", "validar", "divergente", "erro", "incompleto", "conferir", "rastrear", "controle", "governança", "let", "seerro"],
        "quando_usar": "Quando a solução precisa ser explicável, segura e resistente a erro.",
        "quando_nao_usar": "Quando a fórmula é simples e não há risco relevante.",
        "formulas_principais": ["LET", "SEERRO", "ÉERROS", "ÉCÉL.VAZIA", "SE", "E", "OU"],
        "combinacoes_estrategicas": [
            '=SEERRO(A2/B2;0)',
            '=SE(ARRUMAR(A2)="";"Vazio";"Preenchido")',
            '=LET(v;B2;status;C2;SE(OU(v="";status="");"INCOMPLETO";"OK"))',
            '=SE(ÉERROS(PROCX(A2;Dim[SKU];Dim[SKU]));"Sem cadastro";"OK")',
        ],
        "estrategia_solucao": [
            "Separar validação de dados da regra de negócio.",
            "Nomear variáveis com LET quando houver muitas etapas.",
            "Não esconder erro sem entender a causa.",
            "Criar status de qualidade quando houver risco de entrada inválida.",
            "Explicar a fórmula como faria para auditor ou gestor.",
        ],
        "erro_comum": "Usar SEERRO para mascarar erro de cadastro ou erro lógico.",
        "nivel_cobranca": "Muito relevante em testes pleno/sênior.",
        "peso_teste": 4,
    },
    {
        "id": "TIPO-09",
        "tipo_problema": "Estatística e análise exploratória no Excel",
        "pergunta_diagnostico": "O enunciado pede média, mediana, dispersão, correlação, tendência ou previsão?",
        "palavras_chave": ["média", "media", "mediana", "moda", "desvio", "variância", "variancia", "correlação", "correlacao", "rquad", "tendência", "forecast", "percentil", "quartil"],
        "quando_usar": "Quando a questão pede interpretar comportamento dos dados, não apenas calcular total.",
        "quando_nao_usar": "Quando a amostra é pequena demais ou quando a relação não é linear e o enunciado não pede modelo.",
        "formulas_principais": ["MÉDIA", "MED", "MODO.ÚNICO", "DESVPAD.S", "DESVPAD.P", "CORREL", "RQUAD", "PROJ.LIN", "TENDÊNCIA", "QUARTIL.INC", "PERCENTIL.INC"],
        "combinacoes_estrategicas": [
            '=MÉDIA(Base[Valor])',
            '=MED(Base[Valor])',
            '=DESVPAD.S(Base[Valor])',
            '=CORREL(Base[Investimento];Base[Vendas])',
            '=RQUAD(Base[Vendas];Base[Investimento])',
        ],
        "estrategia_solucao": [
            "Identificar se o pedido é descritivo, comparação, dispersão ou relação entre variáveis.",
            "Comparar média e mediana quando houver outliers.",
            "Usar desvio padrão para dispersão.",
            "Não confundir correlação com causalidade.",
            "Sempre escrever interpretação em linguagem de negócio.",
        ],
        "erro_comum": "Entregar só o número sem interpretação.",
        "nivel_cobranca": "Médio em Excel, alto em analista de dados.",
        "peso_teste": 3,
    },
    {
        "id": "TIPO-10",
        "tipo_problema": "Performance, modelagem e estrutura de planilha",
        "pergunta_diagnostico": "O problema fala de lentidão, base grande, tabela estruturada, modelo ou manutenção?",
        "palavras_chave": ["lenta", "performance", "tabela estruturada", "modelo", "manutenção", "base grande", "fato", "dimensão", "referência", "estruturada"],
        "quando_usar": "Quando o enunciado avalia maturidade de modelagem e não apenas sintaxe.",
        "quando_nao_usar": "Quando é um exercício simples e isolado.",
        "formulas_principais": ["Tabelas Estruturadas", "LET", "PROCX", "SOMASES", "SUBTOTAL", "FILTRO"],
        "combinacoes_estrategicas": [
            '=SOMASES(Base[Valor];Base[Produto];A2)',
            '=PROCX([@SKU];DimProduto[SKU];DimProduto[Categoria];"Sem cadastro")',
            '=SUBTOTAL(9;Base[Valor])',
            '=LET(chave;A2&"|"&B2;SE(chave="|";"SEM CHAVE";chave))',
        ],
        "estrategia_solucao": [
            "Preferir tabelas estruturadas a colunas inteiras.",
            "Separar fato e dimensão quando houver cadastro auxiliar.",
            "Evitar fórmulas voláteis em excesso.",
            "Usar LET para evitar repetição de cálculo.",
            "Explicar por que sua solução escala melhor.",
        ],
        "erro_comum": "Resolver com fórmula que funciona em 10 linhas, mas trava em 100 mil.",
        "nivel_cobranca": "Diferencial em testes pleno/sênior.",
        "peso_teste": 4,
    },
]


THEME_TO_TYPE = {
    "Texto": "TIPO-01",
    "Buscas": "TIPO-02",
    "Agregações": "TIPO-03",
    "Condicionais": "TIPO-04",
    "Matrizes Dinâmicas": "TIPO-05",
    "Datas": "TIPO-06",
    "Coringas": "TIPO-07",
    "Auditoria": "TIPO-08",
    "Estatística no Excel": "TIPO-09",
    "Estatística": "TIPO-09",
    "Performance": "TIPO-10",
    "Modelagem": "TIPO-10",
}

TYPE_LOOKUP = {item["id"]: item for item in PROBLEM_TYPES}


# =============================================================================
# 3. FUNÇÕES DE ENRIQUECIMENTO DOS EXERCÍCIOS EXCEL
# =============================================================================

def get_base_exercises() -> List[Dict[str, Any]]:
    base = globals().get("EXERCICIOS_RESOLVIDOS")
    if base is None:
        raise RuntimeError(
            "EXERCICIOS_RESOLVIDOS não encontrado. Cole este Bloco 09 depois do Bloco 08 ou antes carregue a lista de exercícios."
        )
    return base


def get_excel_base_exercises() -> List[Dict[str, Any]]:
    return [dict(item) for item in get_base_exercises() if item.get("ferramenta") == "Excel"]


def classify_exercise(exercise: Dict[str, Any]) -> str:
    tema = exercise.get("tema", "")
    if tema in THEME_TO_TYPE:
        return THEME_TO_TYPE[tema]

    text = " ".join(
        str(exercise.get(k, ""))
        for k in ["enunciado", "problema", "analise", "raciocinio", "solucao_excel", "solucao_principal"]
    ).lower()

    scored = []
    for item in PROBLEM_TYPES:
        score = sum(1 for word in item["palavras_chave"] if word.lower() in text)
        scored.append((score, item["id"]))

    scored.sort(reverse=True)
    return scored[0][1] if scored and scored[0][0] > 0 else "TIPO-08"


def difficulty_score(exercise: Dict[str, Any]) -> int:
    nivel = str(exercise.get("nivel", "")).lower()
    formula = str(exercise.get("solucao_excel", exercise.get("solucao_principal", ""))).upper()

    score = 1
    if "intermedi" in nivel:
        score = 2
    if "avanç" in nivel or "avanc" in nivel:
        score = 3
    if "expert" in nivel:
        score = 4

    advanced_tokens = ["LET", "LAMBDA", "MAP", "BYROW", "PROCX(1", "ESCOLHERCOLS", "PEGAR", "FILTRO", "ÚNICO", "CLASSIFICAR"]
    if any(token in formula for token in advanced_tokens):
        score = max(score, 4)

    if sum(token in formula for token in ["FILTRO", "ÚNICO", "CLASSIFICAR", "PROCX", "SOMASES", "CONT.SES", "LET"]) >= 3:
        score = 5

    return min(score, 5)


def make_related(exercises: List[Dict[str, Any]], current: Dict[str, Any], problem_type_id: str, limit: int = 6) -> List[str]:
    same_type = []
    same_theme = []

    for item in exercises:
        if item.get("id") == current.get("id"):
            continue
        item_type = classify_exercise(item)
        if item_type == problem_type_id:
            same_type.append(item.get("id"))
        elif item.get("tema") == current.get("tema"):
            same_theme.append(item.get("id"))

    return (same_type + same_theme)[:limit]


def strategy_for_exercise(exercise: Dict[str, Any], problem_type: Dict[str, Any]) -> Dict[str, Any]:
    formula = exercise.get("solucao_excel") or exercise.get("solucao_principal", "")
    return {
        "diagnostico_do_enunciado": problem_type["pergunta_diagnostico"],
        "tipo_problema": problem_type["tipo_problema"],
        "caminho_recomendado": [
            "Ler verbo principal do enunciado",
            "Identificar entrada e saída esperada",
            problem_type["tipo_problema"],
            "Selecionar família de fórmulas",
            "Aplicar solução",
            "Validar resultado e erro comum",
        ],
        "familia_formulas": problem_type["formulas_principais"],
        "formula_principal": formula,
        "estrategia_excel_1_formula_direta": "Usar fórmula direta quando a tarefa é pontual, clara e de baixa recorrência.",
        "estrategia_excel_2_coluna_auxiliar": "Usar coluna auxiliar quando houver limpeza, validação, regra intermediária ou necessidade de auditoria.",
        "estrategia_excel_3_tabela_dinamica": "Usar Tabela Dinâmica quando a pergunta pede exploração rápida, resumo por dimensões e filtros.",
        "estrategia_excel_4_tabela_estruturada": "Usar Tabela Estruturada para manter referências legíveis e expansão automática.",
        "estrategia_excel_5_let_auditoria": "Usar LET quando a fórmula ficar longa, repetir cálculo ou precisar ser explicada em entrevista.",
        "estrategia_entrevista": (
            "Explique primeiro o tipo de problema, depois a família de fórmulas, depois a fórmula. "
            "Finalize dizendo como validaria e qual erro comum evitaria."
        ),
        "quando_simplificar": "Se a solução couber em uma fórmula curta e auditável, não complique com fórmula matricial avançada.",
        "quando_escalar": "Se virar processo recorrente, base grande ou regra corporativa, evoluir para tabela estruturada, tabela dinâmica ou modelo de dados.",
    }


def build_enriched_excel_exercises() -> List[Dict[str, Any]]:
    exercises = get_excel_base_exercises()
    enriched = []

    for item in exercises:
        ex = dict(item)
        ptype_id = classify_exercise(ex)
        ptype = TYPE_LOOKUP[ptype_id]

        # Remove visão Power Query do exercício Excel enriquecido.
        ex.pop("solucao_powerquery", None)

        ex["tipo_problema_id"] = ptype_id
        ex["tipo_problema"] = ptype["tipo_problema"]
        ex["pergunta_de_diagnostico"] = ptype["pergunta_diagnostico"]
        ex["dificuldade_logica"] = difficulty_score(ex)
        ex["formulas_estrategicas_relacionadas"] = ptype["formulas_principais"]
        ex["combinacoes_estrategicas_relacionadas"] = ptype["combinacoes_estrategicas"]
        ex["exercicios_correspondentes_no_mapa"] = make_related(exercises, ex, ptype_id)
        ex["estrategias_excel"] = strategy_for_exercise(ex, ptype)
        ex["armadilha_do_enunciado"] = ptype["erro_comum"]
        ex["decisao_final_recomendada"] = (
            f"Classificar como '{ptype['tipo_problema']}', aplicar "
            f"{', '.join(ptype['formulas_principais'][:3])} e validar com amostra."
        )

        enriched.append(ex)

    return enriched


def build_problem_types_with_exercises() -> List[Dict[str, Any]]:
    enriched = build_enriched_excel_exercises()
    problem_types = [dict(item) for item in PROBLEM_TYPES]

    for ptype in problem_types:
        associados = [ex["id"] for ex in enriched if ex["tipo_problema_id"] == ptype["id"]]
        ptype["exercicios_associados"] = associados
        ptype["quantidade_exercicios_associados"] = len(associados)

    return problem_types


def export_excel_decision_framework_payload() -> Dict[str, Any]:
    return {
        "app_version": APP_VERSION,
        "block_name": BLOCK_NAME,
        "status": BLOCK_STATUS,
        "build_date": BUILD_DATE,
        "committee": COMITE_EXCEL_DECISAO,
        "problem_types": build_problem_types_with_exercises(),
        "excel_exercises_enriched": build_enriched_excel_exercises(),
    }


# =============================================================================
# 4. MAPA MENTAL E DIAGNÓSTICO
# =============================================================================

def build_decision_tree_text() -> str:
    return """ENUNCIADO
│
├─ 1. O que o verbo principal pede?
│   ├─ Limpar / padronizar / comparar texto → Limpeza e padronização
│   ├─ Buscar / retornar / consultar → Busca e cadastro
│   ├─ Somar / contar / calcular média → Agregação por critérios
│   ├─ Classificar / criar status / faixa → Condicionais e regra
│   ├─ Listar / filtrar / top N → Matrizes dinâmicas
│   ├─ Prazo / vencimento / competência → Datas
│   ├─ Contém / começa / termina / padrão → Coringas
│   ├─ Validar / divergente / incompleto → Auditoria
│   ├─ Média / correlação / desvio → Estatística
│   └─ Lentidão / modelo / manutenção → Performance e modelagem
│
├─ 2. Qual é a saída esperada?
│   ├─ Uma célula → fórmula direta
│   ├─ Uma coluna → fórmula em tabela estruturada
│   ├─ Uma tabela/lista → FILTRO / ÚNICO / CLASSIFICAR
│   ├─ Um resumo → SOMASES / CONT.SES / Tabela Dinâmica
│   └─ Um diagnóstico → LET / SEERRO / validação
│
├─ 3. Qual é o risco?
│   ├─ Texto sujo → limpar antes
│   ├─ Chave inexistente → tratar erro
│   ├─ Duplicidade → validar chave
│   ├─ Base grande → evitar colunas inteiras
│   └─ Regra longa → usar LET
│
└─ 4. Como entregar?
    ├─ Fórmula
    ├─ Raciocínio
    ├─ Validação
    ├─ Erro comum
    └─ Alternativa Excel
"""


def recommend_path_from_statement(statement: str) -> Dict[str, Any]:
    text = statement.lower()
    scores = []

    for ptype in PROBLEM_TYPES:
        score = sum(1 for word in ptype.get("palavras_chave", []) if word.lower() in text)
        scores.append((score, ptype))

    scores.sort(key=lambda item: item[0], reverse=True)
    best_score, best_type = scores[0]

    if best_score == 0:
        best_type = TYPE_LOOKUP["TIPO-08"]

    problem_types_with_exercises = build_problem_types_with_exercises()
    full_type = next(pt for pt in problem_types_with_exercises if pt["id"] == best_type["id"])

    return {
        "tipo_recomendado": full_type["tipo_problema"],
        "pergunta_diagnostico": full_type["pergunta_diagnostico"],
        "formulas_principais": full_type["formulas_principais"],
        "combinacoes_estrategicas": full_type["combinacoes_estrategicas"],
        "estrategia_solucao": full_type["estrategia_solucao"],
        "erro_comum": full_type["erro_comum"],
        "exercicios_associados": full_type.get("exercicios_associados", []),
    }


# =============================================================================
# 5. VALIDAÇÃO
# =============================================================================

def validate_payload() -> List[Dict[str, Any]]:
    checks = []

    try:
        enriched = build_enriched_excel_exercises()
        problem_types = build_problem_types_with_exercises()
        source_ok = True
        source_error = ""
    except Exception as exc:
        enriched = []
        problem_types = []
        source_ok = False
        source_error = str(exc)

    checks.append({
        "regra": "Fonte EXERCICIOS_RESOLVIDOS carregada",
        "status": "OK" if source_ok else "FALHA",
        "evidencia": "Fonte encontrada" if source_ok else source_error,
        "criticidade": "Crítica",
    })
    checks.append({
        "regra": "Tipos de problema >= mínimo",
        "status": "OK" if len(PROBLEM_TYPES) >= MIN_PROBLEM_TYPES else "FALHA",
        "evidencia": str(len(PROBLEM_TYPES)),
        "criticidade": "Crítica",
    })
    checks.append({
        "regra": "Exercícios Excel >= mínimo",
        "status": "OK" if len(enriched) >= MIN_EXCEL_EXERCISES else "FALHA",
        "evidencia": str(len(enriched)),
        "criticidade": "Crítica",
    })
    checks.append({
        "regra": "Exercícios com mapa e estratégia >= mínimo",
        "status": "OK" if sum(1 for ex in enriched if ex.get("tipo_problema_id") and ex.get("estrategias_excel")) >= MIN_EXERCISES_WITH_MAPPING else "FALHA",
        "evidencia": str(sum(1 for ex in enriched if ex.get("tipo_problema_id") and ex.get("estrategias_excel"))),
        "criticidade": "Crítica",
    })
    checks.append({
        "regra": "Power Query removido da visão Excel",
        "status": "OK" if all("solucao_powerquery" not in ex for ex in enriched) else "FALHA",
        "evidencia": "Campo solucao_powerquery removido dos exercícios Excel enriquecidos",
        "criticidade": "Alta",
    })
    checks.append({
        "regra": "Todos os tipos principais possuem estrutura pedagógica",
        "status": "OK" if all(pt.get("formulas_principais") and pt.get("estrategia_solucao") for pt in PROBLEM_TYPES) else "FALHA",
        "evidencia": "Tipos possuem fórmulas e estratégia",
        "criticidade": "Alta",
    })
    checks.append({
        "regra": "Streamlit opcional",
        "status": "OK",
        "evidencia": "Import tratado com ModuleNotFoundError",
        "criticidade": "Crítica",
    })

    return checks


def print_validation_report() -> None:
    checks = validate_payload()
    falhas = sum(1 for item in checks if item["status"] == "FALHA")

    try:
        enriched_count = len(build_enriched_excel_exercises())
        problem_type_count = len(build_problem_types_with_exercises())
        mapped_count = sum(1 for ex in build_enriched_excel_exercises() if ex.get("estrategias_excel"))
    except Exception:
        enriched_count = 0
        problem_type_count = len(PROBLEM_TYPES)
        mapped_count = 0

    print("=" * 78)
    print("VALIDAÇÃO DO BLOCO 09 — EXCEL DECISION FRAMEWORK V14")
    print("=" * 78)
    print(f"Bloco: {BLOCK_NAME}")
    print(f"Versão: {APP_VERSION}")
    print(f"Status: {BLOCK_STATUS}")
    print(f"Streamlit disponível: {STREAMLIT_AVAILABLE}")
    print(f"Tipos de problema: {problem_type_count}")
    print(f"Exercícios Excel enriquecidos: {enriched_count}")
    print(f"Exercícios com estratégia Excel: {mapped_count}")
    print(f"Falhas core: {falhas}")
    print(f"Aprovado core: {falhas == 0}")
    print("-" * 78)
    print("CHECKS CORE:")
    for item in checks:
        print(f"[{item['status']}] {item['regra']} — {item['evidencia']}")
    print("-" * 78)
    print("MAPA MENTAL RESUMIDO:")
    print(build_decision_tree_text())
    print("=" * 78)


# =============================================================================
# 6. UI STREAMLIT OPCIONAL
# =============================================================================

CUSTOM_CSS = """
<style>
.block-container { max-width: 1600px !important; padding-top: 1rem !important; padding-left: 1.25rem !important; padding-right: 1.25rem !important; }
.stTabs [data-baseweb="tab-list"] { gap: 0.25rem; flex-wrap: wrap; }
.stTabs [data-baseweb="tab"] { height: auto; min-height: 38px; white-space: normal; padding: 8px 12px; }
pre, code { white-space: pre-wrap !important; word-break: break-word !important; overflow-wrap: anywhere !important; font-size: 0.78rem !important; }
</style>
"""


def render_table(title: str, rows: List[Dict[str, Any]], height: int = 420) -> None:
    st.subheader(title)
    if pd is None:
        st.write(rows)
        return
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True, height=height)


def filter_exercises(rows: List[Dict[str, Any]], search: str = "", tipo: str = "Todos", nivel: str = "Todos") -> List[Dict[str, Any]]:
    result = []
    term = search.lower().strip()

    for row in rows:
        if tipo != "Todos" and row.get("tipo_problema") != tipo:
            continue
        if nivel != "Todos" and row.get("nivel") != nivel:
            continue
        if term and term not in json.dumps(row, ensure_ascii=False).lower():
            continue
        result.append(row)

    return result


def render_excel_decision_framework_app() -> None:
    if not STREAMLIT_AVAILABLE:
        raise RuntimeError("Streamlit não instalado. Use print_validation_report().")

    st.set_page_config(page_title="Excel Decision Framework V14", page_icon="🧠", layout="wide")
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

    enriched = build_enriched_excel_exercises()
    problem_types = build_problem_types_with_exercises()

    st.title("Bloco 09 — Excel Decision Framework V14")
    st.caption("Mapa mental · Árvore de decisão · Exercícios relacionados · Estratégias Excel")

    tipos = ["Todos"] + sorted(set(pt["tipo_problema"] for pt in problem_types))
    niveis = ["Todos"] + sorted(set(ex.get("nivel", "") for ex in enriched))

    with st.sidebar:
        page = st.radio(
            "Seção",
            ["Dashboard", "Mapa Mental", "Tipos de Problema", "Exercícios Relacionados", "Diagnóstico por Enunciado", "Auditoria"],
            index=0,
        )
        tipo = st.selectbox("Tipo de problema", tipos)
        nivel = st.selectbox("Nível", niveis)
        search = st.text_input("Buscar", "")

    if page == "Dashboard":
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Tipos de problema", len(problem_types))
        c2.metric("Exercícios Excel", len(enriched))
        c3.metric("Com estratégias", sum(1 for ex in enriched if ex.get("estrategias_excel")))
        c4.metric("Falhas core", sum(1 for item in validate_payload() if item["status"] == "FALHA"))

        render_table("Distribuição por tipo", [
            {
                "tipo": pt["tipo_problema"],
                "peso_teste": pt["peso_teste"],
                "nível_cobrança": pt["nivel_cobranca"],
                "exercícios": pt["quantidade_exercicios_associados"],
            }
            for pt in problem_types
        ])

    elif page == "Mapa Mental":
        st.code(build_decision_tree_text(), language="text")
        st.info("Regra de ouro: não comece pela fórmula. Comece classificando o tipo de problema, a saída esperada e o risco do enunciado.")

    elif page == "Tipos de Problema":
        rows = problem_types
        if tipo != "Todos":
            rows = [row for row in rows if row["tipo_problema"] == tipo]
        if search:
            rows = [row for row in rows if search.lower() in json.dumps(row, ensure_ascii=False).lower()]

        render_table("Tipos de problema", rows, height=500)

        for pt in rows:
            with st.expander(f"{pt['id']} · {pt['tipo_problema']} · exercícios: {pt['quantidade_exercicios_associados']}", expanded=False):
                st.markdown(f"**Pergunta de diagnóstico:** {pt['pergunta_diagnostico']}")
                st.markdown(f"**Quando usar:** {pt['quando_usar']}")
                st.markdown(f"**Quando NÃO usar:** {pt['quando_nao_usar']}")
                st.markdown("**Fórmulas principais:**")
                st.write(pt["formulas_principais"])
                st.markdown("**Combinações estratégicas:**")
                for formula in pt["combinacoes_estrategicas"]:
                    st.code(formula, language="text")
                st.markdown("**Exercícios associados:**")
                st.write(pt["exercicios_associados"])

    elif page == "Exercícios Relacionados":
        rows = filter_exercises(enriched, search=search, tipo=tipo, nivel=nivel)
        summary_rows = [
            {
                "id": ex.get("id"),
                "tema": ex.get("tema"),
                "nivel": ex.get("nivel"),
                "tipo_problema": ex.get("tipo_problema"),
                "dificuldade_logica": ex.get("dificuldade_logica"),
                "formula": ex.get("solucao_excel") or ex.get("solucao_principal"),
                "relacionados": ", ".join(ex.get("exercicios_correspondentes_no_mapa", [])),
            }
            for ex in rows
        ]

        render_table("Exercícios Excel mapeados", summary_rows, height=500)

        for ex in rows[:50]:
            with st.expander(f"{ex.get('id')} · {ex.get('tipo_problema')} · {ex.get('nivel')}", expanded=False):
                st.markdown("**Enunciado**")
                st.write(ex.get("enunciado", ""))
                st.markdown("**Diagnóstico**")
                st.write(ex.get("pergunta_de_diagnostico", ""))
                st.markdown("**Solução Excel**")
                st.code(ex.get("solucao_excel") or ex.get("solucao_principal", ""), language="text")
                st.markdown("**Estratégias Excel**")
                st.json(ex.get("estrategias_excel", {}))
                st.markdown("**Exercícios correspondentes no mapa**")
                st.write(ex.get("exercicios_correspondentes_no_mapa", []))

    elif page == "Diagnóstico por Enunciado":
        st.markdown("Cole um enunciado e o framework sugerirá o tipo de problema e as fórmulas prováveis.")
        statement = st.text_area("Enunciado", height=160, placeholder="Exemplo: Somar o valor de vendas por produto e canal...")
        if statement:
            result = recommend_path_from_statement(statement)
            st.subheader("Recomendação")
            st.json(result)

    elif page == "Auditoria":
        render_table("Checklist", validate_payload(), height=420)


if __name__ == "__main__":
    print_validation_report()

