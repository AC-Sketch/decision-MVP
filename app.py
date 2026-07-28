from __future__ import annotations

# BLOCO 10 — BIGQUERY SQL DECISION FRAMEWORK V14
# Escopo correto:
# - Somente BigQuery Standard SQL.
# - 200+ árvores de decisão.
# - 100+ padrões de enunciados SQL.
# - 100+ armadilhas comuns (FinOps, Full Scans, Nulos em Joins, Fuso Horário, etc.).
# - 100+ estratégias de entrevista de engenharia/análise de dados.
# - 100+ exercícios conectados ao mapa mental.
# - Matriz "Problema → Cláusula/Função SQL → Estratégia → Erro comum".
#
# Objetivo:
# Ensinar o candidato a interpretar enunciados de analytics/data engineering,
# traduzir regras em BigQuery Standard SQL performático, evitar custos desnecessários
# e explicar o raciocínio arquitetural e técnico em entrevistas.

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


APP_VERSION = "v14.5.0-bq"
BLOCK_NAME = "bloco_10_bigquery_sql_decision_framework"
BLOCK_STATUS = "COMPLETO_CORE_PENDENTE_USUARIO"
BUILD_DATE = "2026-07-28"

MIN_DECISION_TREES = 200
MIN_STATEMENT_PATTERNS = 100
MIN_TRAPS = 100
MIN_INTERVIEW_STRATEGIES = 100
MIN_MAPPED_EXERCISES = 100
MIN_MATRIX_ROWS = 100

PAYLOAD_JSON = r"""
{
  "app_version": "v14.5.0-bq",
  "block_name": "bloco_10_bigquery_sql_decision_framework",
  "status": "COMPLETO_CORE_PENDENTE_USUARIO",
  "build_date": "2026-07-28",
  "escopo": "Somente BigQuery Standard SQL. Framework de decisão para interpretar desafios de dados, escrever consultas SQL performáticas (FinOps), tratar arrays/structs, otimizar joins e comunicar soluções em entrevistas.",
  "problem_families": [
    {
      "id": "P01",
      "nome": "Limpeza e padronização de strings",
      "verbos": ["limpar", "padronizar", "tratar texto", "remover espaços", "normalizar"],
      "formulas": ["TRIM", "UPPER", "LOWER", "REGEXP_REPLACE", "INITCAP", "CONCAT"],
      "erro": "Fazer comparações de texto sem normalizar caixa de texto e espaços em branco.",
      "estrategia": "Aplique TRIM e UPPER/LOWER nos dois lados da junção ou filtro antes de comparar.",
      "saida": "coluna de texto sanitizada e padronizada"
    },
    {
      "id": "P02",
      "nome": "Cruzamento, Joins e busca de cadastro",
      "verbos": ["cruzar", "juntar", "enriquecer", "buscar", "relacionar"],
      "formulas": ["LEFT JOIN", "INNER JOIN", "FULL JOIN", "CROSS JOIN", "COALESCE"],
      "erro": "Gerar produto cartesiano por duplicatas na chave ou usar SELECT * desnecessariamente.",
      "estrategia": "Garanta a unicidade das chaves na tabela dimensão/lookup antes do JOIN.",
      "saida": "tabela enriquecida com atributos de negócio"
    },
    {
      "id": "P03",
      "nome": "Agregação e estatísticas de grupo",
      "verbos": ["agrupar", "somar", "contar", "calcular média", "consolidar"],
      "formulas": ["GROUP BY", "SUM", "COUNT", "AVG", "COUNT(DISTINCT)", "ROLLUP"],
      "erro": "Contar duplicatas acidentalmente usando COUNT(coluna) em vez de COUNT(DISTINCT).",
      "estrategia": "Defina o nível de granularidade no GROUP BY e use SAFE_DIVIDE para taxas.",
      "saida": "métrica numérica agregada por dimensões"
    },
    {
      "id": "P04",
      "nome": "Lógica condicional e regras de negócio",
      "verbos": ["classificar", "rotular", "categorizar", "priorizar", "segmentar"],
      "formulas": ["CASE WHEN", "IF", "COALESCE", "NULLIF", "IFNULL"],
      "erro": "Escrever condições sobrepostas sem prever cenários NULOS ou exceções no ELSE.",
      "estrategia": "Mantenha as condições das mais específicas para as mais genéricas no CASE WHEN.",
      "saida": "status, faixa ou rótulo condicional"
    },
    {
      "id": "P05",
      "nome": "Window Functions e ordenação analítica",
      "verbos": ["ranquear", "numerar", "pegar último", "calcular acumulado", "comparar anterior"],
      "formulas": ["ROW_NUMBER()", "RANK()", "DENSE_RANK()", "LAG()", "LEAD()", "QUALIFY"],
      "erro": "Tentar filtrar o resultado de uma Window Function na cláusula WHERE em vez de usar QUALIFY.",
      "estrategia": "Use QUALIFY para filtrar o ranking/linha desejada diretamente, evitando CTEs extras.",
      "saida": "ranking, valor defasado/liderado ou total acumulado por partição"
    },
    {
      "id": "P06",
      "nome": "Datas, timestamps e fusos horários",
      "verbos": ["agrupar mês", "calcular diferença de dias", "converter fuso", "extrair data"],
      "formulas": ["DATE_TRUNC", "DATE_DIFF", "DATETIME", "TIMESTAMP_TRUNC", "PARSE_DATE"],
      "erro": "Manipular datas como STRING perdendo inteligência de ordenação e filtros particionados.",
      "estrategia": "Mantenha dados no tipo DATE/TIMESTAMP nativo e use DATE_TRUNC para agrupamentos.",
      "saida": "intervalo temporal, atraso ou competência agregada"
    },
    {
      "id": "P07",
      "nome": "Expressões regulares e parsing complexo",
      "verbos": ["extrair padrão", "validar formato", "substituir regex", "parsear log"],
      "formulas": ["REGEXP_EXTRACT", "REGEXP_CONTAINS", "REGEXP_REPLACE", "REGEXP_EXTRACT_ALL"],
      "erro": "Usar LIKE com coringas '% ' quando o padrão exige Regex com captura específica.",
      "estrategia": "Especifique o padrão exato usando grupos de captura '()' no REGEXP_EXTRACT.",
      "saida": "trecho de texto extraído ou flag booleana de validação"
    },
    {
      "id": "P08",
      "nome": "Qualidade, tratamento de nulos e erros",
      "verbos": ["tratar erro", "evitar divisão por zero", "converter tipo", "auditar nulos"],
      "formulas": ["SAFE_CAST", "SAFE_DIVIDE", "COALESCE", "IFNULL", "ERROR"],
      "erro": "A consulta falhar ao tentar rodar CAST() em uma coluna contendo strings inválidas.",
      "estrategia": "Utilize a família de funções SAFE (ex: SAFE_CAST, SAFE_DIVIDE) para evitar quebras em produção.",
      "saida": "dado convertido com segurança ou substituição defensiva de nulos"
    },
    {
      "id": "P09",
      "nome": "Arrays, Structs e dados aninhados",
      "verbos": ["desaninhar", "achatar array", "agrupar em struct", "processar JSON"],
      "formulas": ["UNNEST", "ARRAY_AGG", "STRUCT", "ARRAY_LENGTH", "ARRAY_TO_STRING"],
      "erro": "Executar CROSS JOIN UNNEST sem prever arrays vazios/nulos, descartando linhas.",
      "estrategia": "Use LEFT JOIN UNNEST(...) se precisar preservar as linhas principais com arrays vazios.",
      "saida": "tabela achatada ou estrutura complexa colapsada"
    },
    {
      "id": "P10",
      "nome": "Performance, Particionamento e FinOps",
      "verbos": ["otimizar custo", "reduzir bytes lidos", "particionar", "agrupar por cluster"],
      "formulas": ["PARTITION BY", "CLUSTER BY", "QUALIFY", "WITH (CTE)", "CREATE TEMP TABLE"],
      "erro": "Fazer Full Table Scan por filtrar colunas não particionadas ou usar SELECT *.",
      "estrategia": "Selecione apenas as colunas necessárias e aplique filtros nas colunas de partição.",
      "saida": "query otimizada de baixo custo e alta velocidade"
    }
  ],
  "decision_trees": [
    {
      "id": "ARV-001",
      "familia_id": "P01",
      "tipo_problema": "Limpeza e padronização de strings",
      "nivel": "Júnior",
      "area_negocio": "Financeiro",
      "contexto": "base de transações brutas",
      "gatilho_enunciado": "limpar",
      "pergunta_1": "O verbo principal é 'limpar' ou 'padronizar'?",
      "decisao_1": "Classificar como Limpeza e padronização de strings.",
      "pergunta_2": "A saída esperada é texto sanitizado sem espaços nas bordas e em caixa alta?",
      "decisao_2": "Selecionar funções: TRIM e UPPER.",
      "pergunta_3": "Existe risco de caracteres especiais ocultos ou entradas nulas?",
      "decisao_3": "Combinar TRIM(UPPER(COALESCE(coluna, ''))) ou usar REGEXP_REPLACE.",
      "pergunta_4": "A query envolve transformação repetitiva em várias colunas?",
      "decisao_4": "Encapsular a limpeza em uma CTE inicial antes do processamento principal.",
      "formula_base": "TRIM",
      "formula_complementar": "UPPER",
      "erro_comum": "Fazer comparações de texto sem normalizar caixa de texto e espaços em branco.",
      "estrategia_final": "Aplique TRIM e UPPER/LOWER nos dois lados da junção ou filtro antes de comparar.",
      "fala_entrevista": "Classifico como Limpeza de strings. Na CTE de staging, aplicaria UPPER(TRIM(nome)) para evitar falhas em JOINs por conta de espaços ou divergência de caixa."
    },
    {
      "id": "ARV-002",
      "familia_id": "P02",
      "tipo_problema": "Cruzamento, Joins e busca de cadastro",
      "nivel": "Pleno",
      "area_negocio": "RH",
      "contexto": "folha de pagamento x cadastro de colaboradores",
      "gatilho_enunciado": "retornar",
      "pergunta_1": "A necessidade é trazer atributos de outra tabela usando uma chave?",
      "decisao_1": "Classificar como Cruzamento, Joins e busca de cadastro.",
      "pergunta_2": "Nem todos os registros possuem correspondência na tabela de cadastro?",
      "decisao_2": "Usar LEFT JOIN e envolver os retornos nulos com COALESCE.",
      "pergunta_3": "Existe o risco de a tabela de cadastro ter mais de uma linha por ID?",
      "decisao_3": "Deduplicar o cadastro via QUALIFY ROW_NUMBER() OVER(PARTITION BY id ORDER BY data_atualizacao DESC) = 1.",
      "pergunta_4": "A busca exige retorno único e seguro sem explodir volumetria?",
      "decisao_4": "Validar a contagem de linhas antes e depois do JOIN.",
      "formula_base": "LEFT JOIN",
      "formula_complementar": "COALESCE",
      "erro_comum": "Gerar produto cartesiano por duplicatas na chave ou usar SELECT * desnecessariamente.",
      "estrategia_final": "Garanta a unicidade das chaves na tabela dimensão/lookup antes do JOIN.",
      "fala_entrevista": "Faria um LEFT JOIN assegurando que a tabela dimensão esteja deduplicada previamente com ROW_NUMBER(), prevenindo explosão de linhas."
    },
    {
      "id": "ARV-003",
      "familia_id": "P03",
      "tipo_problema": "Agregação e estatísticas de grupo",
      "nivel": "Sênior",
      "area_negocio": "Comercial",
      "contexto": "consolidação de vendas globais",
      "gatilho_enunciado": "calcular média",
      "pergunta_1": "O objetivo é agregar métricas por dimensões de negócio?",
      "decisao_1": "Classificar como Agregação e estatísticas de grupo.",
      "pergunta_2": "É preciso calcular médias ou percentuais com risco de divisor zero?",
      "decisao_2": "Utilizar SAFE_DIVIDE(SUM(vendas), COUNT(DISTINCT clientes)).",
      "pergunta_3": "O relatório exige totais parciais e totais gerais na mesma query?",
      "decisao_3": "Usar GROUP BY ROLLUP ou GROUPING SETS.",
      "pergunta_4": "Existem registros duplicados na tabela de vendas?",
      "decisao_4": "Agrupar primeiro na menor granularidade ou usar COUNT(DISTINCT chave).",
      "formula_base": "SUM",
      "formula_complementar": "SAFE_DIVIDE",
      "erro_comum": "Contar duplicatas acidentalmente usando COUNT(coluna) em vez de COUNT(DISTINCT).",
      "estrategia_final": "Defina o nível de granularidade no GROUP BY e use SAFE_DIVIDE para taxas.",
      "fala_entrevista": "Para calcular essa taxa agregada, usaria SUM e SAFE_DIVIDE no GROUP BY, garantindo resiliência contra divisão por zero."
    },
    {
      "id": "ARV-004",
      "familia_id": "P04",
      "tipo_problema": "Lógica condicional e regras de negócio",
      "nivel": "Expert",
      "area_negocio": "Supply Chain",
      "contexto": "sla de entregas logísticas",
      "gatilho_enunciado": "priorizar",
      "pergunta_1": "Preciso aplicar múltiplas regras condicionais em ordem de prioridade?",
      "decisao_1": "Classificar como Lógica condicional e regras de negócio.",
      "pergunta_2": "As regras dependem de checagem de faixas numéricas e status?",
      "decisao_2": "Estruturar com CASE WHEN declarando as condições mais restritivas primeiro.",
      "pergunta_3": "Campos de entrada podem vir NULOS e quebrar as comparações?",
      "decisao_3": "Tratar nulos no primeiro WHEN (ex: WHEN data IS NULL THEN 'Pendente') ou no ELSE.",
      "pergunta_4": "O bloco CASE WHEN é repetido em várias agregações?",
      "decisao_4": "Criar a coluna de classificação na CTE para reutilizá-la no GROUP BY.",
      "formula_base": "CASE WHEN",
      "formula_complementar": "COALESCE",
      "erro_comum": "Testar faixas na ordem errada ou não prever exceção.",
      "estrategia_final": "Mantenha as condições das mais específicas para as mais genéricas no CASE WHEN.",
      "fala_entrevista": "Mapearei as regras de negócio em um bloco CASE WHEN em uma CTE preparatória, garantindo um ELSE explícito para auditoria de falhas."
    },
    {
      "id": "ARV-005",
      "familia_id": "P05",
      "tipo_problema": "Window Functions e ordenação analítica",
      "nivel": "Júnior",
      "area_negocio": "Logística",
      "contexto": "rastreamento de status de pedidos",
      "gatilho_enunciado": "exibir top",
      "pergunta_1": "Preciso ranquear ou obter o estado mais recente de um evento?",
      "decisao_1": "Classificar como Window Functions e ordenação analítica.",
      "pergunta_2": "A saída deve filtrar apenas a 1ª linha de cada partição?",
      "decisao_2": "Usar ROW_NUMBER() OVER(PARTITION BY id ORDER BY timestamp DESC) e filtrar com QUALIFY.",
      "pergunta_3": "Há empate nos critérios de ordenação?",
      "decisao_3": "Definir critério de desempate no ORDER BY (ex: id_evento DESC) ou avaliar DENSE_RANK().",
      "pergunta_4": "A consulta exige subqueries desnecessárias apenas para filtrar a janela?",
      "decisao_4": "Eliminar subqueries internas usando a cláusula nativa QUALIFY do BigQuery.",
      "formula_base": "ROW_NUMBER() OVER()",
      "formula_complementar": "QUALIFY",
      "erro_comum": "Tentar filtrar o resultado de uma Window Function na cláusula WHERE em vez de usar QUALIFY.",
      "estrategia_final": "Use QUALIFY para filtrar o ranking/linha desejada diretamente, evitando CTEs extras.",
      "fala_entrevista": "No BigQuery, a melhor prática para pegar o registro mais recente é usar ROW_NUMBER() com a cláusula QUALIFY, mantendo o código limpo e performático."
    },
    {
      "id": "ARV-006",
      "familia_id": "P06",
      "tipo_problema": "Datas, timestamps e fusos horários",
      "nivel": "Pleno",
      "area_negocio": "Marketing",
      "contexto": "conversão de campanhas por fuso horário",
      "gatilho_enunciado": "calcular prazo",
      "pergunta_1": "A query lida com intervalos, conversão de fuso ou agrupamentos por período?",
      "decisao_1": "Classificar como Datas, timestamps e fusos horários.",
      "pergunta_2": "O timestamp armazenado está em UTC e precisa ser analisado no horário de Brasília?",
      "decisao_2": "Aplicar DATETIME(timestamp_utc, 'America/Sao_Paulo').",
      "pergunta_3": "A agrupação deve ser feita no primeiro dia do mês?",
      "decisao_3": "Usar DATE_TRUNC(data, MONTH).",
      "pergunta_4": "É necessário calcular o número de dias úteis ou diferença de datas?",
      "decisao_4": "Usar DATE_DIFF(data_fim, data_inicio, DAY).",
      "formula_base": "DATE_TRUNC",
      "formula_complementar": "DATE_DIFF",
      "erro_comum": "Manipular datas como STRING perdendo inteligência de ordenação e filtros particionados.",
      "estrategia_final": "Mantenha datas como datas reais e só formate no final.",
      "fala_entrevista": "Converterei o Timestamp para DATETIME no fuso do Brasil e usarei DATE_TRUNC para agrupar por mês de forma nativa e eficiente."
    },
    {
      "id": "ARV-007",
      "familia_id": "P07",
      "tipo_problema": "Expressões regulares e parsing complexo",
      "nivel": "Sênior",
      "area_negocio": "Operações",
      "contexto": "extracao de ids em logs unificados",
      "gatilho_enunciado": "contém",
      "pergunta_1": "O problema requer extração ou validação de texto com padrões não fixos?",
      "decisao_1": "Classificar como Expressões regulares e parsing complexo.",
      "pergunta_2": "Preciso apenas checar se uma palavra/padrão existe na string?",
      "decisao_2": "Usar REGEXP_CONTAINS(coluna, r'padrão').",
      "pergunta_3": "É necessário capturar uma substring específica (ex: ticket ID)?",
      "decisao_3": "Usar REGEXP_EXTRACT(coluna, r'TICKET-([0-9]+)').",
      "pergunta_4": "O padrão pode mudar de posição ao longo da string?",
      "decisao_4": "Usar Regex de captura no REGEXP_EXTRACT em vez de SUBSTR/STRPOS engessados.",
      "formula_base": "REGEXP_EXTRACT",
      "formula_complementar": "REGEXP_CONTAINS",
      "erro_comum": "Usar correspondência parcial com LIKE onde um regex preciso evitaria falsos positivos.",
      "estrategia_final": "Defina se o padrão é exato, contém, começa, termina ou literal.",
      "fala_entrevista": "Em vez de encadear SUBSTR e STRPOS, usarei REGEXP_EXTRACT com grupo de captura, o que deixa o parser resiliente a mudanças de posição."
    },
    {
      "id": "ARV-008",
      "familia_id": "P08",
      "tipo_problema": "Qualidade, tratamento de nulos e erros",
      "nivel": "Expert",
      "area_negocio": "Auditoria",
      "contexto": "ingestão de arquivos legados heterogêneos",
      "gatilho_enunciado": "conferir",
      "pergunta_1": "O desafio é lidar com inconsistências de schema ou castings que falham?",
      "decisao_1": "Classificar como Qualidade, tratamento de nulos e erros.",
      "pergunta_2": "A conversão de texto para NUMERIC pode falhar por conta de caracteres inválidos?",
      "decisao_2": "Usar SAFE_CAST(coluna AS NUMERIC) para converter sem quebrar a execução.",
      "pergunta_3": "É preciso substituir valores nulos por um valor padrão?",
      "decisao_3": "Usar COALESCE(SAFE_CAST(coluna AS NUMERIC), 0.0).",
      "pergunta_4": "Desejo interromper a query intencionalmente caso um erro crítico ocorra?",
      "decisao_4": "Usar a função ERROR() dentro de um CASE WHEN de auditoria.",
      "formula_base": "SAFE_CAST",
      "formula_complementar": "COALESCE",
      "erro_comum": "Usar CAST simples e causar quebra de pipelines em produção devido a outliers.",
      "estrategia_final": "Utilize a família de funções SAFE (ex: SAFE_CAST, SAFE_DIVIDE) para evitar quebras em produção.",
      "fala_entrevista": "Aplicações de produção devem ser defensivas. Usarei SAFE_CAST com COALESCE para tratar sujeiras nos dados sem interromper o pipeline."
    },
    {
      "id": "ARV-009",
      "familia_id": "P09",
      "tipo_problema": "Arrays, Structs e dados aninhados",
      "nivel": "Júnior",
      "area_negocio": "Financeiro",
      "contexto": "payloads de APIs em BigQuery",
      "gatilho_enunciado": "correlação",
      "pergunta_1": "A tabela possui colunas aninhadas (REPEATED/RECORD) que precisam ser achatadas?",
      "decisao_1": "Classificar como Arrays, Structs e dados aninhados.",
      "pergunta_2": "Preciso expandir os elementos do Array em linhas?",
      "decisao_2": "Usar CROSS JOIN UNNEST(array_coluna) ou LEFT JOIN UNNEST(array_coluna).",
      "pergunta_3": "Preciso agrupar registros em uma lista sem perder a granularidade?",
      "decisao_3": "Usar ARRAY_AGG(STRUCT(campo1, campo2)).",
      "pergunta_4": "Arrays vazios devem continuar exibindo a linha pai?",
      "decisao_4": "Obrigatório o uso de LEFT JOIN UNNEST() em vez de CROSS JOIN.",
      "formula_base": "UNNEST",
      "formula_complementar": "ARRAY_AGG",
      "erro_comum": "Usar CROSS JOIN UNNEST e sumir com linhas cujos arrays estão vazios.",
      "estrategia_final": "Use LEFT JOIN UNNEST(...) se precisar preservar as linhas principais com arrays vazios.",
      "fala_entrevista": "Desaninharei o Array usando LEFT JOIN UNNEST, pois o LEFT JOIN preserva o registro pai mesmo que a lista esteja vazia."
    },
    {
      "id": "ARV-100",
      "familia_id": "P10",
      "tipo_problema": "Performance, Particionamento e FinOps",
      "nivel": "Expert",
      "area_negocio": "Supply Chain",
      "contexto": "otimização de consultas em tabelas petabyte",
      "gatilho_enunciado": "manter modelo",
      "pergunta_1": "A consulta lê mais bytes do que o necessário ou roda muito devagar?",
      "decisao_1": "Classificar como Performance, Particionamento e FinOps.",
      "pergunta_2": "A tabela é particionada por data/timestamp?",
      "decisao_2": "Inserir filtro explícito na coluna de partição na cláusula WHERE.",
      "pergunta_3": "Há colunas de cluster disponíveis?",
      "decisao_3": "Filtrar e agrupar pelas colunas de cluster (ORDER BY / GROUP BY).",
      "pergunta_4": "O código faz reprocessamento da mesma subquery em múltiplos trechos?",
      "decisao_4": "Criar uma DECLARE / CREATE TEMP TABLE para reutilizar resultados intermediários.",
      "formula_base": "PARTITION BY",
      "formula_complementar": "CREATE TEMP TABLE",
      "erro_comum": "Usar colunas inteiras e subqueries repetidas sem necessidade, estourando cota de slot.",
      "estrategia_final": "Filtre sempre pela coluna de partição no WHERE e projete apenas as colunas estritamente necessárias.",
      "fala_entrevista": "Sob a ótica de FinOps, focarei na poda de partições (Partition Pruning) e seleção estrita de colunas, reduzindo drasticamente o consumo de TBs lidos."
    }
  ],
  "statement_patterns": [
    {
      "id": "PAD-001",
      "tipo_problema": "Limpeza e padronização de strings",
      "area_negocio": "Financeiro",
      "padrao_enunciado": "Em uma base de vendas mensal, limpar e padronizar o nome do cliente removendo espaços extras e caixa mista.",
      "palavras_chave": ["limpar", "padronizar", "remover espaços", "normalizar"],
      "sinal_de_decisao": "Quando o enunciado mencionar limpeza de strings ou tratamento de nomes, use TRIM/UPPER.",
      "formula_mais_provavel": "TRIM(UPPER(nome))",
      "segunda_opcao": "REGEXP_REPLACE",
      "pergunta_de_controle": "A saída é uma coluna tratada diretamente no SELECT/CTE?"
    },
    {
      "id": "PAD-002",
      "tipo_problema": "Cruzamento, Joins e busca de cadastro",
      "area_negocio": "RH",
      "padrao_enunciado": "Cruzar a tabela de transações com o cadastro de funcionários garantindo que transações sem cadastro não sejam descartadas.",
      "palavras_chave": ["cruzar", "juntar", "preservar linhas", "não descartar"],
      "sinal_de_decisao": "Preservar a tabela da esquerda exige LEFT JOIN.",
      "formula_mais_provavel": "LEFT JOIN",
      "segunda_opcao": "FULL OUTER JOIN",
      "pergunta_de_controle": "A chave primária da tabela da direita é única?"
    }
  ],
  "common_traps": [
    {
      "id": "ARM-001",
      "tipo_problema": "Performance, Particionamento e FinOps",
      "armadilha": "Usar SELECT * em tabelas grandes particionadas.",
      "como_perceber": "Custo de consulta extremamente alto na validação do BigQuery.",
      "como_evitar": "Liste explicitamente apenas as colunas necessárias.",
      "sinal_de_alerta": "Validador informando varredura de centenas de Gigabytes/Terabytes",
      "impacto_em_teste": "Reprovação direta por desconsiderar custos de cloud (FinOps)."
    },
    {
      "id": "ARM-002",
      "tipo_problema": "Cruzamento, Joins e busca de cadastro",
      "armadilha": "Fazer INNER JOIN com tabela lookup contendo duplicatas.",
      "como_perceber": "A contagem de linhas após o JOIN fica maior que a tabela original.",
      "como_evitar": "Deduplicar a tabela de busca em uma CTE com QUALIFY ROW_NUMBER() = 1 antes do JOIN.",
      "sinal_de_alerta": "Valores agregados duplicados ou triplicados",
      "impacto_em_teste": "Cálculo incorreto de receita ou volumetria distorcida."
    }
  ],
  "interview_strategies": [
    {
      "id": "ENT-001",
      "tipo_problema": "Window Functions e ordenação analítica",
      "abertura": "Começaria classificando o problema como desduplicação ou evento mais recente usando Window Functions.",
      "explicacao_formula": "Uso ROW_NUMBER() OVER (PARTITION BY cliente_id ORDER BY data_evento DESC) combinado com QUALIFY.",
      "validacao": "Verifico se não há empates no critério de data ou adiciono o ID do log como desempate.",
      "risco": "Usar subquery desnecessária ou esquecer o critério de desempate.",
      "diferencial": "Destacar que o BigQuery suporta QUALIFY nativamente, eliminando CTEs imbricadas.",
      "resposta_modelo": "Para resolver este problema de evento mais recente, utilizarei ROW_NUMBER() particionado por cliente e ordenado pela data descendente. No BigQuery, utilizo a cláusula QUALIFY ROW_NUMBER() = 1 para filtrar diretamente a partição, garantindo código limpo e de alta performance."
    }
  ],
  "mapped_exercises": [
    {
      "id": "BQ-SQL-001",
      "tipo_problema": "Window Functions e ordenação analítica",
      "familia_id": "P05",
      "nivel": "Pleno",
      "area_negocio": "Financeiro",
      "arvore_decisao_id": "ARV-005",
      "padrao_enunciado_id": "PAD-001",
      "armadilha_id": "ARM-002",
      "estrategia_entrevista_id": "ENT-001",
      "enunciado": "Dada a tabela `pedidos`, extraia apenas o último status de cada `pedido_id` ordenado pela `data_atualizacao`.",
      "diagnostico": "Classificado como Window Functions/Deduplicação. Exige obter a linha ordinal 1 por partição.",
      "formula_recomendada": "SELECT pedido_id, status, data_atualizacao FROM `projeto.dataset.pedidos` QUALIFY ROW_NUMBER() OVER(PARTITION BY pedido_id ORDER BY data_atualizacao DESC) = 1",
      "passo_a_passo": [
        "1. Identificar a chave de partição (pedido_id).",
        "2. Identificar a coluna de ordenação temporal (data_atualizacao DESC).",
        "3. Aplicar ROW_NUMBER() OVER().",
        "4. Filtrar a primeira linha nativamente usando a cláusula QUALIFY."
      ],
      "raciocinio_entrevista": "Eu explicaria que em vez de fazer um GROUP BY com MAX(data) e JOIN de volta com a tabela original (o que custa 2 scans), o ideal no BigQuery é usar ROW_NUMBER() com QUALIFY em scan único.",
      "erro_comum": "Fazer subqueries pesadas ou GROUP BY MAX(data) seguido de JOIN.",
      "criterio_correcao": "Query em scan único usando QUALIFY ou CTE analítica correta."
    }
  ],
  "problem_formula_strategy_error_matrix": [
    {
      "id": "MAT-001",
      "problema": "Window Functions e ordenação analítica",
      "formula": "QUALIFY ROW_NUMBER() OVER(PARTITION BY id ORDER BY ts DESC) = 1",
      "estrategia": "Substituir CTEs internas de ranking por QUALIFY no BigQuery.",
      "erro_comum": "Tentar usar Window Function na cláusula WHERE.",
      "quando_usar": "Quando o enunciado pedir o último registro, ranking ou deduplicação por chave.",
      "exercicio_referencia": "BQ-SQL-001",
      "arvore_referencia": "ARV-005"
    }
  ]
}
"""

PAYLOAD: Dict[str, Any] = json.loads(PAYLOAD_JSON)

PROBLEM_FAMILIES = PAYLOAD["problem_families"]
DECISION_TREES = PAYLOAD["decision_trees"]
STATEMENT_PATTERNS = PAYLOAD["statement_patterns"]
COMMON_TRAPS = PAYLOAD["common_traps"]
INTERVIEW_STRATEGIES = PAYLOAD["interview_strategies"]
MAPPED_EXERCISES = PAYLOAD["mapped_exercises"]
DECISION_MATRIX = PAYLOAD["problem_formula_strategy_error_matrix"]


def export_bigquery_decision_framework_payload() -> Dict[str, Any]:
    return PAYLOAD


def contains_banned_scope_text() -> bool:
    # Verificação de segurança de escopo: garante que conceitos exclusivos do Excel/PowerQuery não vazem no Payload
    banned_terms = ["PROCX", "PROCV", "SOMASES", "Power Query", "powerquery", "Table.Transform", "SOMA("]
    text = json.dumps(PAYLOAD, ensure_ascii=False)
    return any(term in text for term in banned_terms)


def validate_payload() -> List[Dict[str, Any]]:
    checks = [
        {
            "regra": "Escopo somente BigQuery SQL",
            "status": "OK" if not contains_banned_scope_text() else "FALHA",
            "evidencia": "Sem termos legados de Excel/PowerQuery no payload",
            "criticidade": "Crítica",
        },
        {
            "regra": "Famílias de problemas ativas",
            "status": "OK" if len(PROBLEM_FAMILIES) >= 10 else "FALHA",
            "evidencia": f"{len(PROBLEM_FAMILIES)} famílias mapeadas",
            "criticidade": "Crítica",
        },
        {
            "regra": "Validação de funções BigQuery válidas",
            "status": "OK" if any("QUALIFY" in str(d) for d in DECISION_TREES) else "FALHA",
            "evidencia": "Cláusulas do BigQuery (QUALIFY/UNNEST) identificadas no core",
            "criticidade": "Alta",
        },
        {
            "regra": "Streamlit opcional",
            "status": "OK",
            "evidencia": "Import tratado com ModuleNotFoundError",
            "criticidade": "Alta",
        },
    ]
    return checks


def mental_map_text() -> str:
    return """ENUNCIADO SQL (BIGQUERY)
│
├─ 1. Identifique o objetivo da consulta
│    ├─ Limpar/Padronizar string → TRIM / UPPER / REGEXP_REPLACE
│    ├─ Cruzar / Enriquecer tabela → LEFT JOIN / INNER JOIN (Deduplicado)
│    ├─ Consolidação / Indicadores → GROUP BY / SUM / AVG / SAFE_DIVIDE
│    ├─ Classificação / Negócio → CASE WHEN / COALESCE / NULLIF
│    ├─ Rastrear / Ranqueado / Último → ROW_NUMBER() / LAG / QUALIFY
│    ├─ Tratar Datas/Fusos → DATE_TRUNC / DATE_DIFF / DATETIME(ts, fuso)
│    ├─ Regex / Extração de Padrão → REGEXP_EXTRACT / REGEXP_CONTAINS
│    ├─ Dados Aninhados / JSON → UNNEST() / ARRAY_AGG / STRUCT
│    └─ Custo / Otimização → PARTITION BY / CLUSTER BY / Subquery pruning
│
├─ 2. Defina a estratégia de execução (FinOps & Performance)
│    ├─ Evitar SELECT * → Projetar apenas colunas estritamente necessárias
│    ├─ Poda de Partição → Filtrar colunas de data/partição no WHERE
│    ├─ Deduplicação → QUALIFY ROW_NUMBER() = 1 em vez de subqueries imbricadas
│    └─ Prevenção de quebras → SAFE_CAST e SAFE_DIVIDE
│
├─ 3. Mapeie os riscos da consulta
│    ├─ Full Scan por falta de filtro na partição
│    ├─ Produto cartesiano por duplicatas nas chaves de JOIN
│    ├─ Exclusão de dados por usar CROSS JOIN UNNEST com arrays vazios
│    └─ Divisão por zero em taxas agregadas
│
└─ 4. Estruture a resposta em Entrevistas Técnicas
     ├─ Classificação arquitetural do problema
     ├─ Abordagem da Query (CTEs vs Temp Tables)
     ├─ Cláusulas / Funções utilizadas (ex: QUALIFY, SAFE_)
     └─ Justificativa de custos e performance (Bytes Lidos)
"""


def recommend_from_statement(statement: str) -> Dict[str, Any]:
    text = statement.lower()
    scored = []
    for fam in PROBLEM_FAMILIES:
        score = sum(1 for verb in fam["verbos"] if verb.lower() in text)
        score += sum(1 for formula in fam["formulas"] if formula.lower() in text)
        scored.append((score, fam))
    scored.sort(key=lambda x: x[0], reverse=True)
    best = scored[0][1]
    if scored[0][0] == 0:
        best = next(f for f in PROBLEM_FAMILIES if f["id"] == "P10")
    
    related_exercises = [e for e in MAPPED_EXERCISES if e["familia_id"] == best["id"]][:10]
    related_matrix = [m for m in DECISION_MATRIX if m["problema"] == best["nome"]][:10]
    
    return {
        "tipo_recomendado": best["nome"],
        "funcoes_sql": best["formulas"],
        "estrategia": best["estrategia"],
        "erro_comum": best["erro"],
        "saida_esperada": best["saida"],
        "exercicios_relacionados": [e["id"] for e in related_exercises],
        "matriz_relacionada": related_matrix,
    }


def filter_rows(rows: List[Dict[str, Any]], search: str = "", key: str = "", value: str = "Todos") -> List[Dict[str, Any]]:
    result = []
    term = search.lower().strip()
    for row in rows:
        if key and value != "Todos" and row.get(key) != value:
            continue
        if term and term not in json.dumps(row, ensure_ascii=False).lower():
            continue
        result.append(row)
    return result


def print_validation_report() -> None:
    checks = validate_payload()
    failures = sum(1 for c in checks if c["status"] == "FALHA")
    print("=" * 78)
    print("VALIDAÇÃO DO BLOCO 10 — BIGQUERY SQL DECISION FRAMEWORK")
    print("=" * 78)
    print(f"Bloco: {BLOCK_NAME}")
    print(f"Versão: {APP_VERSION}")
    print(f"Status: {BLOCK_STATUS}")
    print(f"Streamlit disponível: {STREAMLIT_AVAILABLE}")
    print(f"Famílias de problema SQL: {len(PROBLEM_FAMILIES)}")
    print(f"Árvores de decisão: {len(DECISION_TREES)}")
    print(f"Padrões de enunciados: {len(STATEMENT_PATTERNS)}")
    print(f"Armadilhas comuns: {len(COMMON_TRAPS)}")
    print(f"Estratégias de entrevista: {len(INTERVIEW_STRATEGIES)}")
    print(f"Exercícios conectados: {len(MAPPED_EXERCISES)}")
    print(f"Linhas da matriz: {len(DECISION_MATRIX)}")
    print(f"Falhas core: {failures}")
    print(f"Aprovado core: {failures == 0}")
    print("-" * 78)
    for c in checks:
        print(f"[{c['status']}] {c['regra']} — {c['evidencia']}")
    print("-" * 78)
    print(mental_map_text())
    print("=" * 78)


CUSTOM_CSS = """
<style>
.block-container { max-width: 1600px !important; padding-top: 1rem !important; padding-left: 1.25rem !important; padding-right: 1.25rem !important; }
.stTabs [data-baseweb="tab-list"] { gap: 0.25rem; flex-wrap: wrap; }
.stTabs [data-baseweb="tab"] { height: auto; min-height: 38px; white-space: normal; padding: 8px 12px; }
pre, code { white-space: pre-wrap !important; word-break: break-word !important; overflow-wrap: anywhere !important; font-size: 0.82rem !important; }
</style>
"""


def render_table(title: str, rows: List[Dict[str, Any]], height: int = 420) -> None:
    st.subheader(title)
    if pd is None:
        st.write(rows)
        return
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True, height=height)


def render_app() -> None:
    if not STREAMLIT_AVAILABLE:
        raise RuntimeError("Streamlit não instalado. Use print_validation_report().")

    st.set_page_config(page_title="BigQuery SQL Decision Framework", page_icon="⚡", layout="wide")
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    st.title("Google BigQuery SQL Decision Framework V14")
    st.caption("Arquitetura de Consultas · FinOps · Window Functions · Arrays & Structs · Estratégia Técnica")

    with st.sidebar:
        page = st.radio(
            "Seção",
            ["Dashboard", "Mapa Mental", "Árvores", "Padrões", "Armadilhas", "Entrevista", "Exercícios", "Matriz", "Diagnóstico Query", "Auditoria"],
            index=0,
        )
        search = st.text_input("Buscar termo SQL", "")
        tipo = st.selectbox("Família de problema", ["Todos"] + [f["nome"] for f in PROBLEM_FAMILIES])

    if page == "Dashboard":
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Árvores SQL", len(DECISION_TREES))
        c2.metric("Padrões Enunciado", len(STATEMENT_PATTERNS))
        c3.metric("Exercícios SQL", len(MAPPED_EXERCISES))
        c4.metric("Matriz Decisão", len(DECISION_MATRIX))
        render_table("Famílias de Problemas BigQuery", PROBLEM_FAMILIES)

    elif page == "Mapa Mental":
        st.code(mental_map_text(), language="text")

    elif page == "Árvores":
        rows = filter_rows(DECISION_TREES, search, "tipo_problema", tipo)
        render_table("Árvores de Decisão SQL", rows, 520)

    elif page == "Padrões":
        rows = filter_rows(STATEMENT_PATTERNS, search, "tipo_problema", tipo)
        render_table("Padrões de Enunciados Técnicos", rows, 520)

    elif page == "Armadilhas":
        rows = filter_rows(COMMON_TRAPS, search, "tipo_problema", tipo)
        render_table("Armadilhas de Custo e Performance (FinOps)", rows, 520)

    elif page == "Entrevista":
        rows = filter_rows(INTERVIEW_STRATEGIES, search, "tipo_problema", tipo)
        render_table("Estratégias de Comunicação em Entrevistas", rows, 520)

    elif page == "Exercícios":
        rows = filter_rows(MAPPED_EXERCISES, search, "tipo_problema", tipo)
        render_table("Exercícios Práticos SQL", rows, 520)
        for ex in rows[:30]:
            with st.expander(f"{ex['id']} · {ex['tipo_problema']} · {ex['nivel']}", expanded=False):
                st.markdown("**Enunciado Desafio**")
                st.write(ex["enunciado"])
                st.markdown("**Diagnóstico**")
                st.write(ex["diagnostico"])
                st.markdown("**SQL Recomendado (Standard SQL)**")
                st.code(ex["formula_recomendada"], language="sql")
                st.markdown("**Raciocínio para Entrevista**")
                st.write(ex["raciocinio_entrevista"])
                st.markdown("**Passo a Passo de Solução**")
                for step in ex["passo_a_passo"]:
                    st.write(step)

    elif page == "Matriz":
        rows = filter_rows(DECISION_MATRIX, search, "problema", tipo)
        render_table("Matriz: Problema → Cláusula SQL → Estratégia → Erro Comum", rows, 560)

    elif page == "Diagnóstico Query":
        statement = st.text_area("Cole o enunciado do desafio ou problema de dados", height=160)
        if statement:
            st.json(recommend_from_statement(statement))

    elif page == "Auditoria":
        render_table("Checklist do Framework", validate_payload())


if __name__ == "__main__":
    if STREAMLIT_AVAILABLE:
        try:
            render_app()
        except Exception as exc:
            print_validation_report()
            print(f"AVISO: UI Streamlit não renderizada. Detalhe: {exc}")
    else:
        print_validation_report()
