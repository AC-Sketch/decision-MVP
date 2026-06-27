from __future__ import annotations

# BLOCO INGLÊS — DECISION FRAMEWORK V14 (DATA ANALYTICS)
# Escopo:
# - 200 árvores de decisão.
# - 100 padrões de enunciados.
# - 100 armadilhas comuns.
# - Matriz "Problema → Estratégia → Fala de Entrevista".
#
# Objetivo:
# Ensinar o candidato a articular experiências reais em inglês para entrevistas.

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

APP_VERSION = "v14.5.0"
BLOCK_NAME = "bloco_ingles_decision_framework"
BLOCK_STATUS = "COMPLETO_CORE_PENDENTE_USUARIO"
BUILD_DATE = "2026-06-27"

PAYLOAD_JSON = r"""
{
  "app_version": "v14.5.0",
  "block_name": "bloco_ingles_entrevista_analytics",
  "status": "BUILDING_PHASE",
  "build_date": "2026-06-27",
  "problem_families": [
    {
      "id": "PF-01",
      "nome": "Small Talk & Context",
      "descricao": "Quebra-gelo e adequação cultural EMEA"
    },
    {
      "id": "PF-02",
      "nome": "Método STAR (Comportamental)",
      "descricao": "Relatos de cases reais (Itaú, Fretebras, etc)"
    },
    {
      "id": "PF-03",
      "nome": "Technical Deep Dive",
      "descricao": "Tradução de conceitos de Engenharia/Dados para o inglês"
    }
  ],
  "decision_trees": [
    {
      "id": "DT-01",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 1 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-02",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 2 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-03",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 3 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-04",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 4 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-05",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 5 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-06",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 6 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-07",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 7 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-08",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 8 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-09",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 9 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-010",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 10 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-011",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 11 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-012",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 12 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-013",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 13 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-014",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 14 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-015",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 15 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-016",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 16 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-017",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 17 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-018",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 18 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-019",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 19 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-020",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 20 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-021",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 21 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-022",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 22 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-023",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 23 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-024",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 24 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-025",
      "tipo_problema": "Small Talk",
      "nivel": "Basic",
      "enunciado": "Cenário de aquecimento 25 - Entrevistador pergunta sobre clima ou rotina.",
      "diagnostico": "Identificar se o candidato usa 'It' para clima e tempos verbais corretos.",
      "formula_recomendada": "Subject + Verb + Context",
      "raciocinio_entrevista": "O objetivo é mostrar naturalidade. Evite traduções literais do português.",
      "passo_a_passo": [
        "Escuta ativa",
        "Uso de 'It' como sujeito",
        "Adição de um detalhe pessoal"
      ]
    },
    {
      "id": "DT-026",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 1 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-027",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 2 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-028",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 3 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-029",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 4 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-030",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 5 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-031",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 6 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-032",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 7 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-033",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 8 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-034",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 9 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-035",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 10 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-036",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 11 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-037",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 12 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-038",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 13 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-039",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 14 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-040",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 15 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-041",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 16 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-042",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 17 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-043",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 18 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-044",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 19 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-045",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 20 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-046",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 21 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-047",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 22 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-048",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 23 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-049",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 24 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-050",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Pergunta comportamental 25 sobre conflito ou priorização.",
      "diagnostico": "Avaliar a estruturação da resposta conforme o framework STAR.",
      "formula_recomendada": "Situation + Task + Action + Result",
      "raciocinio_entrevista": "Focar nos resultados quantitativos (ex: economia de custos/tempo).",
      "passo_a_passo": [
        "Definir o contexto",
        "Detalhar a ação técnica",
        "Quantificar o impacto"
      ]
    },
    {
      "id": "DT-051",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 51: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-052",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 52: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-053",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 53: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-054",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 54: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-055",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 55: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-056",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 56: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-057",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 57: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-058",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 58: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-059",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 59: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-060",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 60: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-061",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 61: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-062",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 62: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-063",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 63: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-064",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 64: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-065",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 65: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-066",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 66: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-067",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 67: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-068",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 68: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-069",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 69: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-070",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 70: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-071",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 71: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-072",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 72: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-073",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 73: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-074",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 74: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-075",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 75: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-076",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 76: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-077",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 77: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-078",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 78: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-079",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 79: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-080",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 80: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-081",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 81: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-082",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 82: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-083",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 83: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-084",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 84: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-085",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 85: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-086",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 86: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-087",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 87: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-088",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 88: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-089",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 89: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-090",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 90: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-091",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 91: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-092",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 92: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-093",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 93: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-094",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 94: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-095",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 95: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-096",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 96: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-097",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 97: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-098",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 98: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-099",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 99: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-100",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 100: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-101",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 101: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-102",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 102: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-103",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 103: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-104",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 104: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-105",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 105: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-106",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 106: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-107",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 107: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-108",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 108: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-109",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 109: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-110",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 110: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-111",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 111: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-112",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 112: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-113",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 113: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-114",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 114: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-115",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 115: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-116",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 116: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-117",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 117: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-118",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 118: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-119",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 119: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-120",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 120: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-121",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 121: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-122",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 122: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-123",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 123: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-124",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 124: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-125",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 125: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-126",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 126: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-127",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 127: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-128",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 128: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-129",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 129: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-130",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 130: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-131",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 131: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-132",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 132: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-133",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 133: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-134",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 134: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-135",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 135: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-136",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 136: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-137",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 137: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-138",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 138: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-139",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 139: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-140",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 140: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-141",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 141: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-142",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 142: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-143",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 143: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-144",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 144: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-145",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 145: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-146",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 146: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-147",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 147: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-148",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 148: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-149",
      "tipo_problema": "STAR_Behavioral",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 149: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-150",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 150: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-151",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 151: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-152",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 152: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-153",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 153: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-154",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 154: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-155",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 155: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-156",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 156: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-157",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 157: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-158",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 158: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-159",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 159: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-160",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 160: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-161",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 161: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-162",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 162: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-163",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 163: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-164",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 164: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-165",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 165: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-166",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 166: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-167",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 167: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-168",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 168: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-169",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 169: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-170",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 170: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-171",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 171: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-172",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 172: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-173",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 173: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-174",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 174: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-175",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 175: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-176",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 176: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-177",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 177: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-178",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 178: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-179",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 179: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-180",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 180: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-181",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 181: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-182",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 182: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-183",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 183: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-184",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 184: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-185",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 185: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-186",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 186: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-187",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 187: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-188",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 188: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-189",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 189: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-190",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 190: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-191",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 191: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-192",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 192: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-193",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 193: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-194",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 194: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-195",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 195: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-196",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 196: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-197",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 197: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-198",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 198: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-199",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 199: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    },
    {
      "id": "DT-200",
      "tipo_problema": "Technical_DeepDive",
      "nivel": "Senior",
      "enunciado": "Cenário de entrevista 200: Simulação de case de Dados.",
      "diagnostico": "Análise de proficiência técnica e comportamental.",
      "formula_recomendada": "STAR/Technical_Framework",
      "raciocinio_entrevista": "Clareza na explicação do impacto de negócio.",
      "passo_a_passo": [
        "Identificar intenção",
        "Estruturar resposta",
        "Validar com métrica"
      ]
    }
  ],
  "statement_patterns": [
    {
      "id": "SP-001",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-002",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-003",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-004",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-005",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-006",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-007",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-008",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-009",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-010",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-011",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-012",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-013",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-014",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-015",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-016",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-017",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-018",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-019",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-020",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-021",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-022",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-023",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-024",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-025",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-026",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-027",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-028",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-029",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-030",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-031",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-032",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-033",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-034",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-035",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-036",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-037",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-038",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-039",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-040",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-041",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-042",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-043",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-044",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-045",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-046",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-047",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-048",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-049",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-050",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-051",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-052",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-053",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-054",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-055",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-056",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-057",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-058",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-059",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-060",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-061",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-062",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-063",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-064",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-065",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-066",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-067",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-068",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-069",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-070",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-071",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-072",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-073",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-074",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-075",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-076",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-077",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-078",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-079",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-080",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-081",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-082",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-083",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-084",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-085",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-086",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-087",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-088",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-089",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-090",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-091",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-092",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-093",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-094",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-095",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-096",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-097",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-098",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-099",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    },
    {
      "id": "SP-100",
      "trigger": "Tell me about",
      "pattern": "Behavioral_Prompt",
      "acao_sugerida": "Aplicar Método STAR"
    }
  ],
  "interview_strategies": [],
  "traps": [
    {
      "id": "TRP-001",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-002",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-003",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-004",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-005",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-006",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-007",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-008",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-009",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-010",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-011",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-012",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-013",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-014",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-015",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-016",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-017",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-018",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-019",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-020",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-021",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-022",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-023",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-024",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-025",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-026",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-027",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-028",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-029",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-030",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-031",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-032",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-033",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-034",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-035",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-036",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-037",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-038",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-039",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-040",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-041",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-042",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-043",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-044",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-045",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-046",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-047",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-048",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-049",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-050",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-051",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-052",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-053",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-054",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-055",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-056",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-057",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-058",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-059",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-060",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-061",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-062",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-063",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-064",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-065",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-066",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-067",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-068",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-069",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-070",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-071",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-072",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-073",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-074",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-075",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-076",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-077",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-078",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-079",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-080",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-081",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-082",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-083",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-084",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-085",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-086",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-087",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-088",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-089",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-090",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-091",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-092",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-093",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-094",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-095",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-096",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-097",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-098",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-099",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    },
    {
      "id": "TRP-100",
      "desc": "Armadilha de tradução literal do português",
      "correcao": "Utilizar termos técnicos de mercado (ex: stockouts, throughput)"
    }
  ]
}
"""
