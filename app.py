
import json
import streamlit as st

# --- CAMADA DE DADOS ---
PAYLOAD_JSON = r"""{
  "app_version": "v14.5.0",
  "decision_trees": [
    {
      "id": "DT-001",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 1",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-002",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 2",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-003",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 3",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-004",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 4",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-005",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 5",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-006",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 6",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-007",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 7",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-008",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 8",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-009",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 9",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-010",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 10",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-011",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 11",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-012",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 12",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-013",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 13",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-014",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 14",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-015",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 15",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-016",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 16",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-017",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 17",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-018",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 18",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-019",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 19",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-020",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 20",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-021",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 21",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-022",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 22",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-023",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 23",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-024",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 24",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-025",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 25",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-026",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 26",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-027",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 27",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-028",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 28",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-029",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 29",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-030",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 30",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-031",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 31",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-032",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 32",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-033",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 33",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-034",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 34",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-035",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 35",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-036",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 36",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-037",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 37",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-038",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 38",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-039",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 39",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-040",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 40",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-041",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 41",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-042",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 42",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-043",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 43",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-044",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 44",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-045",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 45",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-046",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 46",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-047",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 47",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-048",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 48",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-049",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 49",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-050",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 50",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-051",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 51",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-052",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 52",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-053",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 53",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-054",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 54",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-055",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 55",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-056",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 56",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-057",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 57",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-058",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 58",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-059",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 59",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-060",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 60",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-061",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 61",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-062",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 62",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-063",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 63",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-064",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 64",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-065",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 65",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-066",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 66",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-067",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 67",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-068",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 68",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-069",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 69",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-070",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 70",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-071",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 71",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-072",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 72",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-073",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 73",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-074",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 74",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-075",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 75",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-076",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 76",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-077",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 77",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-078",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 78",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-079",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 79",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-080",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 80",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-081",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 81",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-082",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 82",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-083",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 83",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-084",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 84",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-085",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 85",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-086",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 86",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-087",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 87",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-088",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 88",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-089",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 89",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-090",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 90",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-091",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 91",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-092",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 92",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-093",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 93",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-094",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 94",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-095",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 95",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-096",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 96",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-097",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 97",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-098",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 98",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-099",
      "tipo_problema": "STAR",
      "nivel": "Senior",
      "enunciado": "Cenário 99",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-100",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 100",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-101",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 101",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-102",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 102",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-103",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 103",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-104",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 104",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-105",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 105",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-106",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 106",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-107",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 107",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-108",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 108",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-109",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 109",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-110",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 110",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-111",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 111",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-112",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 112",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-113",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 113",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-114",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 114",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-115",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 115",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-116",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 116",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-117",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 117",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-118",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 118",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-119",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 119",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-120",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 120",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-121",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 121",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-122",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 122",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-123",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 123",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-124",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 124",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-125",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 125",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-126",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 126",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-127",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 127",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-128",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 128",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-129",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 129",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-130",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 130",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-131",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 131",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-132",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 132",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-133",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 133",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-134",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 134",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-135",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 135",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-136",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 136",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-137",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 137",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-138",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 138",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-139",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 139",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-140",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 140",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-141",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 141",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-142",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 142",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-143",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 143",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-144",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 144",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-145",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 145",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-146",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 146",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-147",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 147",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-148",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 148",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-149",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 149",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-150",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 150",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-151",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 151",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-152",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 152",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-153",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 153",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-154",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 154",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-155",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 155",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-156",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 156",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-157",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 157",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-158",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 158",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-159",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 159",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-160",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 160",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-161",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 161",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-162",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 162",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-163",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 163",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-164",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 164",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-165",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 165",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-166",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 166",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-167",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 167",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-168",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 168",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-169",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 169",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-170",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 170",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-171",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 171",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-172",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 172",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-173",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 173",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-174",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 174",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-175",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 175",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-176",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 176",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-177",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 177",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-178",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 178",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-179",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 179",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-180",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 180",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-181",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 181",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-182",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 182",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-183",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 183",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-184",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 184",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-185",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 185",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-186",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 186",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-187",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 187",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-188",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 188",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-189",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 189",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-190",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 190",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-191",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 191",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-192",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 192",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-193",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 193",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-194",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 194",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-195",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 195",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-196",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 196",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-197",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 197",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-198",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 198",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-199",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 199",
      "formula_recomendada": "STAR/Tech"
    },
    {
      "id": "DT-200",
      "tipo_problema": "Technical",
      "nivel": "Senior",
      "enunciado": "Cenário 200",
      "formula_recomendada": "STAR/Tech"
    }
  ],
  "statement_patterns": [
    {
      "id": "SP-001",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-002",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-003",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-004",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-005",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-006",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-007",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-008",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-009",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-010",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-011",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-012",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-013",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-014",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-015",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-016",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-017",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-018",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-019",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-020",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-021",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-022",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-023",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-024",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-025",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-026",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-027",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-028",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-029",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-030",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-031",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-032",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-033",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-034",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-035",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-036",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-037",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-038",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-039",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-040",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-041",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-042",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-043",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-044",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-045",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-046",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-047",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-048",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-049",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-050",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-051",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-052",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-053",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-054",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-055",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-056",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-057",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-058",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-059",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-060",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-061",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-062",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-063",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-064",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-065",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-066",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-067",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-068",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-069",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-070",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-071",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-072",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-073",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-074",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-075",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-076",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-077",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-078",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-079",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-080",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-081",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-082",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-083",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-084",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-085",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-086",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-087",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-088",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-089",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-090",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-091",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-092",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-093",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-094",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-095",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-096",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-097",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-098",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-099",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    },
    {
      "id": "SP-100",
      "trigger": "Tell me",
      "pattern": "Behavioral"
    }
  ],
  "traps": [
    {
      "id": "TRP-001",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-002",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-003",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-004",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-005",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-006",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-007",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-008",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-009",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-010",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-011",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-012",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-013",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-014",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-015",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-016",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-017",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-018",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-019",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-020",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-021",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-022",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-023",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-024",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-025",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-026",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-027",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-028",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-029",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-030",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-031",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-032",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-033",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-034",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-035",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-036",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-037",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-038",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-039",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-040",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-041",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-042",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-043",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-044",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-045",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-046",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-047",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-048",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-049",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-050",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-051",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-052",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-053",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-054",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-055",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-056",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-057",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-058",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-059",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-060",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-061",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-062",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-063",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-064",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-065",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-066",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-067",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-068",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-069",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-070",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-071",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-072",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-073",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-074",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-075",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-076",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-077",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-078",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-079",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-080",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-081",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-082",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-083",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-084",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-085",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-086",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-087",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-088",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-089",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-090",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-091",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-092",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-093",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-094",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-095",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-096",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-097",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-098",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-099",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    },
    {
      "id": "TRP-100",
      "desc": "Tradução literal",
      "correcao": "Termo técnico"
    }
  ]
}"""

# --- CAMADA DE LÓGICA (ENGINE) ---
class InterviewEngine:
    def __init__(self, data):
        self.data = json.loads(data)
    
    def get_data(self, key):
        return self.data.get(key, [])

# --- CAMADA DE INTERFACE (VIEW) ---
def render_app():
    st.set_page_config(page_title="Data Analytics Interview", layout="wide")
    st.title("Framework de Decisão: Entrevistas (EMEA)")
    
    engine = InterviewEngine(PAYLOAD_JSON)
    
    page = st.sidebar.selectbox("Selecione o Módulo", ["Dashboard", "Árvores", "Auditoria"])
    
    if page == "Dashboard":
        st.subheader("Resumo do Framework")
        st.metric("Total de Árvores", len(engine.get_data("decision_trees")))
        st.metric("Padrões Mapeados", len(engine.get_data("statement_patterns")))
        
    elif page == "Árvores":
        tipo = st.selectbox("Filtrar por Tipo", ["STAR", "Technical"])
        for tree in [t for t in engine.get_data("decision_trees") if t["tipo_problema"] == tipo]:
            with st.expander(f"{tree['id']} - {tree['enunciado']}"):
                st.write(f"Fórmula: {tree['formula_recomendada']}")
                
    elif page == "Auditoria":
        st.json(engine.data['app_version'])

if __name__ == "__main__":
    render_app()
