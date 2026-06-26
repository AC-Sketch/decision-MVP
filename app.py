# -*- coding: utf-8 -*-
import streamlit as st
import random
import json

# ==========================================
# 1. DATABASE COMPLETA EMBUTIDA (JSON)
# ==========================================
DATABASE_DB = {
    "modulos": [
        {
            "id": "MOD-01",
            "titulo": "Introdução à Matéria e Suas Transformações",
            "descricao": "Estudo das propriedades da matéria, estados físicos, curvas de aquecimento e sistemas químicos.",
            "topicos": [
                {
                    "id": "TOP-011",
                    "nome": "Propriedades Gerais e Específicas",
                    "importancia_enem": "Média",
                    "importancia_fuvest": "Alta"
                },
                {
                    "id": "TOP-012",
                    "nome": "Estados Físicos e Fenômenos",
                    "importancia_enem": "Alta",
                    "importancia_fuvest": "Média"
                },
                {
                    "id": "TOP-013",
                    "nome": "Gráficos de Aquecimento (Puras, Eutéticas, Azeotrópicas)",
                    "importancia_enem": "Alta",
                    "importancia_fuvest": "Alta"
                },
                {
                    "id": "TOP-014",
                    "nome": "Métodos de Separação de Misturas",
                    "importancia_enem": "Máxima",
                    "importancia_fuvest": "Alta"
                }
            ]
        },
        {
            "id": "MOD-02",
            "titulo": "Atomística e Estrutura Atômica",
            "descricao": "Evolução dos modelos atômicos, partículas subatômicas, semelhanças atômicas e distribuição eletrônica.",
            "topicos": [
                {
                    "id": "TOP-021",
                    "nome": "Evolução dos Modelos Atômicos",
                    "importancia_enem": "Média",
                    "importancia_fuvest": "Alta"
                },
                {
                    "id": "TOP-022",
                    "nome": "Estrutura Atômica e Íons",
                    "importancia_enem": "Média",
                    "importancia_fuvest": "Média"
                },
                {
                    "id": "TOP-023",
                    "nome": "Semelhanças Atômicas (Isótopos, Isóbaros, Isótonos, Isoeletrônicos)",
                    "importancia_enem": "Baixa",
                    "importancia_fuvest": "Média"
                },
                {
                    "id": "TOP-024",
                    "nome": "Distribuição Eletrônica (Linus Pauling) e Números Quânticos",
                    "importancia_enem": "Média",
                    "importancia_fuvest": "Alta"
                }
            ]
        },
        {
            "id": "MOD-03",
            "titulo": "Classificação Periódica dos Elementos",
            "descricao": "Organização da Tabela Periódica moderna e propriedades periódicas dos elementos químicos.",
            "topicos": [
                {
                    "id": "TOP-031",
                    "nome": "Organização e Grupos da Tabela Periódica",
                    "importancia_enem": "Baixa",
                    "importancia_fuvest": "Média"
                },
                {
                    "id": "TOP-032",
                    "nome": "Propriedades Periódicas (Raio, Ionização, Eletronegatividade)",
                    "importancia_enem": "Alta",
                    "importancia_fuvest": "Alta"
                }
            ]
        },
        {
            "id": "MOD-04",
            "titulo": "Ligações Químicas e Forças Intermoleculares",
            "descricao": "Estudo de como os átomos se unem para formar substâncias e suas interações espaciais.",
            "topicos": [
                {
                    "id": "TOP-041",
                    "nome": "Ligação Iônica e Metálica",
                    "importancia_enem": "Média",
                    "importancia_fuvest": "Média"
                },
                {
                    "id": "TOP-042",
                    "nome": "Ligação Covalente e Geometria Molecular",
                    "importancia_enem": "Média",
                    "importancia_fuvest": "Alta"
                },
                {
                    "id": "TOP-043",
                    "nome": "Polaridade Molecular e Forças Intermoleculares",
                    "importancia_enem": "Máxima",
                    "importancia_fuvest": "Máxima"
                }
            ]
        },
        {
            "id": "MOD-05",
            "titulo": "Funções Inorgânicas",
            "descricao": "Estudo de Ácidos, Bases, Sais e Óxidos segundo a teoria de Arrhenius.",
            "topicos": [
                {
                    "id": "TOP-051",
                    "nome": "Ácidos e Bases (Arrhenius, Força e Nomenclatura)",
                    "importancia_enem": "Alta",
                    "importancia_fuvest": "Alta"
                },
                {
                    "id": "TOP-052",
                    "nome": "Sais e Óxidos (Reações de Neutralização e Classificação)",
                    "importancia_enem": "Alta",
                    "importancia_fuvest": "Alta"
                }
            ]
        },
        {
            "id": "MOD-06",
            "titulo": "Reações Químicas e Introdução à Estequiometria",
            "descricao": "Leis ponderais, balanceamento e cálculos estequiométricos fundamentais.",
            "topicos": [
                {
                    "id": "TOP-061",
                    "nome": "Classificação de Reações e Balanceamento",
                    "importancia_enem": "Média",
                    "importancia_fuvest": "Média"
                },
                {
                    "id": "TOP-062",
                    "nome": "Leis Ponderais e Estequiometria Simples",
                    "importancia_enem": "Máxima",
                    "importancia_fuvest": "Máxima"
                }
            ]
        }
    ],
    "flashcards": [
        {
            "id": "FC-01",
            "topico_id": "TOP-011",
            "frente": "O que diferencia uma propriedade física de uma propriedade química?",
            "verso": "A propriedade física pode ser medida sem alterar a identidade da substância (ex: ponto de fusão). A propriedade química descreve a capacidade de uma substância de sofrer uma transformação química (ex: combustibilidade)."
        },
        {
            "id": "FC-02",
            "topico_id": "TOP-013",
            "frente": "Como se comporta a temperatura durante a fusão de uma mistura azeotrópica?",
            "verso": "Ela varia durante a fusão, mas permanece constante durante a ebulição (comportamento de substância pura na ebulição)."
        },
        {
            "id": "FC-03",
            "topico_id": "TOP-014",
            "frente": "Qual método separa água e acetona baseado nos pontos de ebulição?",
            "verso": "Destilação fracionada."
        },
        {
            "id": "FC-04",
            "topico_id": "TOP-021",
            "frente": "Qual a principal característica do modelo atômico de Thomson?",
            "verso": "O átomo é uma esfera maciça de carga positiva com elétrons incrustados, conhecido como 'pudim de passas'."
        },
        {
            "id": "FC-05",
            "topico_id": "TOP-024",
            "frente": "O que estabelece o Princípio da Exclusão de Pauli?",
            "verso": "Dois elétrons em um mesmo átomo não podem ter os mesmos quatro números quânticos."
        },
        {
            "id": "FC-06",
            "topico_id": "TOP-032",
            "frente": "Qual é a tendência do raio atômico em um período da Tabela Periódica?",
            "verso": "O raio atômico diminui da esquerda para a direita devido ao aumento da carga nuclear efetiva puxando os elétrons mais fortemente."
        },
        {
            "id": "FC-07",
            "topico_id": "TOP-043",
            "frente": "Que tipo de força intermolecular ocorre entre moléculas de água?",
            "verso": "Ligação de Hidrogênio (interação dipolo-dipolo extremamente forte)."
        },
        {
            "id": "FC-08",
            "topico_id": "TOP-051",
            "frente": "Segundo Arrhenius, o que é um ácido?",
            "verso": "É uma substância molecular que, em solução aquosa, sofre ionização, liberando exclusivamente o cátion H+ (ou H3O+)."
        }
    ]
}
FORMULAS_DB = {
    "formulas": [
        {
            "id": "FOR-01",
            "nome": "Densidade Absoluta",
            "expressao": "d = m / V",
            "variaveis": {
                "d": "Densidade (g/cm³ ou g/L)",
                "m": "Massa (g)",
                "V": "Volume (cm³ ou L)"
            },
            "aplicacao": "Relação de massa por unidade de volume de um sistema. Essencial em problemas de flutuação e pureza de materiais.",
            "unidades_alerta": "Cuidado com conversões: 1 cm³ = 1 mL; 1 dm³ = 1 L; 1 m³ = 1000 L."
        },
        {
            "id": "FOR-02",
            "nome": "Relações Fundamentais do Átomo",
            "expressao": "A = Z + n",
            "variaveis": {
                "A": "Número de Massa",
                "Z": "Número Atômico (prótons)",
                "n": "Número de Nêutrons"
            },
            "aplicacao": "Determinação de estrutura de partículas atômicas e identificação de isótopos/isóbaros/isótonos.",
            "unidades_alerta": "Massa e número atômico são adimensionais (números inteiros)."
        },
        {
            "id": "FOR-03",
            "nome": "Cálculo de Mol e Massa Molar",
            "expressao": "n = m / MM",
            "variaveis": {
                "n": "Quantidade de matéria (mol)",
                "m": "Massa da amostra (g)",
                "MM": "Massa Molar (g/mol)"
            },
            "aplicacao": "Conversão de massa macroscópica em contagem de entidades elementares para estequiometria.",
            "unidades_alerta": "A massa 'm' deve sempre estar em gramas para cancelar com a unidade g/mol."
        },
        {
            "id": "FOR-04",
            "nome": "Volume Molar nas CNTP",
            "expressao": "V = n * 22,4",
            "variaveis": {
                "V": "Volume ocupado por gás ideal (L)",
                "n": "Quantidade de mols (mol)",
                "22.4": "Constante de volume molar nas CNTP (L/mol)"
            },
            "aplicacao": "Cálculo de volume de gases produzidos ou consumidos em reações químicas sob condições normais de temperatura e pressão.",
            "unidades_alerta": "Válido apenas nas CNTP (0 °C e 1 atm)."
        }
    ]
}
TREES_DB = {
    "arvores": [
        {
            "id": "TREE-01",
            "titulo": "Classificação de Sistemas e Misturas",
            "raiz": {
                "pergunta": "O sistema apresenta o mesmo aspecto visual em toda a sua extensão sob ultramicroscópio?",
                "sim": {
                    "pergunta": "O sistema é formado por apenas uma substância química (composta ou simples)?",
                    "sim": "Substância Pura Homogênea (Ex: Água destilada)",
                    "nao": "Mistura Homogênea ou Solução (Ex: Água salgada filtrada)"
                },
                "nao": {
                    "pergunta": "A heterogeneidade é visível a olho nu?",
                    "sim": "Mistura Heterogênea Macroscópica (Ex: Água e óleo, granito)",
                    "nao": "Mistura Heterogênea Coloidal (Ex: Leite, Sangue, Maionese)"
                }
            }
        },
        {
            "id": "TREE-02",
            "titulo": "Escolha do Método de Separação de Misturas",
            "raiz": {
                "pergunta": "O sistema é Homogêneo ou Heterogêneo?",
                "homogeneo": {
                    "pergunta": "Quais são os estados físicos dos componentes?",
                    "solido-liquido": {
                        "pergunta": "Deseja-se recuperar o líquido?",
                        "sim": "Destilação Simples (Ex: Água e sal)",
                        "nao": "Evaporação (Ex: Salinas)"
                    },
                    "liquido-liquido": "Destilação Fracionada (Ex: Água e Acetona, Petróleo)",
                    "gas-gas": "Liquefação Fracionada seguida de Destilação Fracionada"
                },
                "heterogeneo": {
                    "pergunta": "Quais são os estados físicos dos componentes?",
                    "solido-liquido": {
                        "pergunta": "A separação precisa ser rápida ou por gravidade lenta?",
                        "rapida": "Centrifugação ou Filtração a Vácuo",
                        "lenta": "Filtração Comum ou Decantação"
                    },
                    "liquido-liquido": "Decantação com Funil de Separação (Funil de Bromo / Decantação)",
                    "solido-solido": {
                        "pergunta": "Existe diferença de densidade tratável por líquido intermediário?",
                        "sim": "Flotação",
                        "nao": "Catação, Ventilação, Separação Magnética ou Tamisação"
                    }
                }
            }
        },
        {
            "id": "TREE-03",
            "titulo": "Identificação de Forças Intermoleculares",
            "raiz": {
                "pergunta": "A molécula é Polar ou Apolar?",
                "apolar": {
                    "pergunta": "Envolve gases nobres ou hidrocarbonetos puros?",
                    "sim": "Dispersões de London / Dipolo Induzido-Dipolo Induzido",
                    "nao": "Dipolo Induzido-Dipolo Induzido"
                },
                "polar": {
                    "pergunta": "Possui átomo de Hidrogênio ligado diretamente a F, O ou N?",
                    "sim": "Ligação de Hidrogênio (Antiga Ponte de Hidrogênio)",
                    "nao": {
                        "pergunta": "Ocorre na presença de íons livres em solução?",
                        "sim": "Interação Íon-Dipolo",
                        "nao": "Dipolo Permanente-Dipolo Permanente (Dipolo-Dipolo)"
                    }
                }
            }
        }
    ]
}
MISTAKES_DB = {
    "armadilhas": [
        {
            "id": "TRAP-01",
            "topico_id": "TOP-013",
            "nome": "Confundir Ponto de Ebulição Constante com Substância Pura",
            "descricao": "Achar que se o gráfico tem um patamar fixo, o sistema é obrigatoriamente puro.",
            "mecanismo": "Misturas azeotrópicas mimetizam substâncias puras durante a ebulição, mantendo a temperatura constante.",
            "como_evitar": "Verifique o comportamento do gráfico na Fusão. Se a Fusão variar e a Ebulição for constante, é uma mistura azeotrópica."
        },
        {
            "id": "TRAP-02",
            "topico_id": "TOP-024",
            "nome": "Distribuição de Íons de Elementos de Transição",
            "descricao": "Fazer a distribuição eletrônica retirando elétrons do subnível mais energético em vez da camada de valência.",
            "mecanismo": "Para o Ferro (Z=26): [Ar] 4s² 3d⁶. Ao formar Fe²⁺, o estudante remove do 3d (ficando 4s² 3d⁴), o que é incorreto.",
            "como_evitar": "Sempre faça a distribuição do átomo neutro primeiro. Identifique a camada de valência (maior nível 'n') e remova dali os elétrons (correto para Fe²⁺: [Ar] 3d⁶)."
        },
        {
            "id": "TRAP-03",
            "topico_id": "TOP-043",
            "nome": "CO₂ possui ligações polares, logo a molécula é polar",
            "descricao": "Associar polaridade da ligação diretamente com a polaridade da molécula de forma automática.",
            "mecanismo": "O CO₂ possui ligações C=O polares, porém sua geometria é linear. Os vetores de momento dipolar se anulam, resultando em momento dipolar resultante nulo (μ = 0), tornando a molécula apolar.",
            "como_evitar": "Determine a geometria da molécula pelo modelo de repulsão dos pares eletrônicos antes de bater o martelo sobre a polaridade molecular."
        }
    ]
}
PROMPTS_DB = {
    "padroes": [
        {
            "id": "PAT-01",
            "topico_id": "TOP-014",
            "contexto": "Instalações industriais, tratamento de água ou laboratórios rudimentares.",
            "gatilhos": [
                "sequência de operações",
                "fluxograma de separação",
                "insumos insolúveis",
                "fases distintas"
            ],
            "estrategia_leitura": "Focar diretamente no estado físico de cada componente e em qual propriedade física (densidade, solubilidade, tamanho) difere entre eles a cada etapa."
        },
        {
            "id": "PAT-02",
            "topico_id": "TOP-043",
            "contexto": "Explicação de ponto de ebulição, solubilidade de poluentes ou comportamento de polímeros.",
            "gatilhos": [
                "série homóloga",
                "temperatura de ebulição",
                "solúvel em água",
                "interação de compostos"
            ],
            "estrategia_leitura": "Identificar se a questão compara tamanhos de cadeias moleculares (forças de London maiores em moléculas maiores) ou tipos de forças distintas (London vs Dipolo vs Ligação de H)."
        }
    ]
}
VESTIBULAR_DB = {
    "bancas": [
        {
            "nome": "ENEM",
            "perfil": "Altamente contextualizado, foco extremo em sustentabilidade, separação de misturas no cotidiano, forças intermoleculares aplicadas a materiais e biologia.",
            "pesos": {
                "TOP-014": 10,
                "TOP-043": 9,
                "TOP-062": 8,
                "TOP-021": 3
            },
            "dicas_mestre": "Não busque decoreba de fórmulas. Busque entender o fenômeno físico-químico aplicado a cenários ambientais ou industriais."
        },
        {
            "nome": "FUVEST",
            "perfil": "Técnico, preciso, exige domínio analítico profundo, clareza em conceitos abstratos como números quânticos, geometria molecular e estequiometria rigorosa com pureza/rendimento.",
            "pesos": {
                "TOP-014": 7,
                "TOP-043": 9,
                "TOP-062": 10,
                "TOP-024": 6
            },
            "dicas_mestre": "Cuidado com a precisão dos cálculos e a escrita de fórmulas estruturais completas na segunda fase."
        },
        {
            "nome": "UNICAMP",
            "perfil": "Interdisciplinar, focado em experimentos contemporâneos, leitura de gráficos complexos, tabelas informativas densas e dedução lógica a partir do enunciado.",
            "pesos": {
                "TOP-013": 8,
                "TOP-014": 8,
                "TOP-043": 8,
                "TOP-062": 9
            },
            "dicas_mestre": "Use os dados fornecidos na própria prova; a UNICAMP adora testar sua capacidade de extrair informação estruturada de textos científicos."
        }
    ]
}
EXERCISES_DB = {
    "exercicios": [
        {
            "id": "EXE-01",
            "topico_id": "TOP-013",
            "enunciado": "Um estudante aquece uma amostra sólida contida em um cadinho e monitora sua variação térmica. Observa que a fusão inicia-se a 80 °C e termina a 85 °C. Contudo, ao atingir o ponto de ebulição a 120 °C, a temperatura estabiliza-se completamente até a evaporação total. Esse comportamento indica que a amostra original constitui uma:",
            "alternativas": {
                "A": "Substância pura composta.",
                "B": "Mistura homogênea eutética.",
                "C": "Mistura homogênea azeotrópica.",
                "D": "Substância pura simples.",
                "E": "Mistura heterogênea coloidal."
            },
            "resposta_correta": "C",
            "dificuldade": "Média",
            "banca_origem": "FUVEST (Adaptado)",
            "resolucao": "As misturas azeotrópicas comportam-se como substâncias puras durante a ebulição (patamar constante), mas exibem faixa de temperatura variável durante o processo de fusão.",
            "link_matrix": "MAT-01"
        },
        {
            "id": "EXE-02",
            "topico_id": "TOP-024",
            "enunciado": "O elemento ferro (Z = 26) é essencial para o transporte de oxigênio no sangue humano por meio da hemoglobina, onde se encontra na forma de cátion Fe²⁺. A configuração eletrônica correta da camada de valência desse cátion no estado fundamental é:",
            "alternativas": {
                "A": "4s² 3d⁴",
                "B": "3d⁶",
                "C": "4s¹ 3d⁵",
                "D": "4s² 3d⁶",
                "E": "3d⁴"
            },
            "resposta_correta": "B",
            "dificuldade": "Difícil",
            "banca_origem": "UNICAMP",
            "resolucao": "A configuração do Ferro neutro é [Ar] 4s² 3d⁶. Sendo a camada de valência o quarto nível (4s²), a ionização para formar Fe²⁺ remove prioritariamente esses dois elétrons periféricos, restando a subcamada [Ar] 3d⁶.",
            "link_matrix": "MAT-02"
        },
        {
            "id": "EXE-03",
            "topico_id": "TOP-043",
            "enunciado": "A água apresenta ponto de ebulição anomalamente elevado (100 °C) quando comparada com o sulfeto de hidrogênio (H₂S, -60 °C), embora o enxofre esteja logo abaixo do oxigênio no mesmo grupo da Tabela Periódica. Essa disparidade de comportamento físico decorre prioritariamente do fato de que:",
            "alternativas": {
                "A": "O H₂S possui ligações covalentes apolares.",
                "B": "As moléculas de H₂O unem-se por forças de dipolo induzido.",
                "C": "A água estabelece ligações de hidrogênio intermoleculares intensas.",
                "D": "O raio atômico do enxofre confere maior geometria linear ao H₂S.",
                "E": "A água é um composto iônico em condições ambientes."
            },
            "resposta_correta": "C",
            "dificuldade": "Fácil",
            "banca_origem": "ENEM",
            "resolucao": "O oxigênio é altamente eletronegativo. A ligação H-O gera dipolos marcantes que interagem via ligações de hidrogênio, demandando expressiva energia térmica para transição de fase líquida para gasosa.",
            "link_matrix": "MAT-03"
        }
    ]
}
ROADMAP_DB = {
    "mapa_mental": {
        "nodo_central": "Química Geral 1° Ano (Objetivo)",
        "conexoes": [
            {
                "origem": "Matéria Inicial",
                "destino": "Sistemas Químicos",
                "tipo": "Pré-requisito",
                "detalhes": "Compreender estados físicos antes de mapear técnicas de fracionamento."
            },
            {
                "origem": "Modelos Atômicos",
                "destino": "Configuração Eletrônica",
                "tipo": "Evolução Conceitual",
                "detalhes": "O modelo de Bohr e Schrödinger serve de alicerce para a distribuição eletrônica em subníveis."
            },
            {
                "origem": "Configuração Eletrônica",
                "destino": "Tabela Periódica",
                "tipo": "Mapeamento Estático",
                "detalhes": "A posição do elemento (período/grupo) é condicionada pelo seu subnível eletrônico terminal."
            },
            {
                "origem": "Tabela Periódica",
                "destino": "Ligações Químicas",
                "tipo": "Aplicações de Afinidade",
                "detalhes": "Diferenças de eletronegatividade definem se a ligação será iônica, covalente ou metálica."
            }
        ]
    }
}
MATRIX_DB = {
    "matriz": [
        {
            "id": "MAT-01",
            "problema": "Análise de curvas térmicas e comportamento de fusão/ebulição",
            "formula_id": null,
            "estrategia_id": "PAT-01",
            "erro_id": "TRAP-01"
        },
        {
            "id": "MAT-02",
            "problema": "Configuração eletrônica de íons metálicos estáveis",
            "formula_id": "FOR-02",
            "estrategia_id": null,
            "erro_id": "TRAP-02"
        },
        {
            "id": "MAT-03",
            "problema": "Comparação de volatilidade e temperaturas de ebulição",
            "formula_id": null,
            "estrategia_id": "PAT-02",
            "erro_id": "TRAP-03"
        }
    ]
}

# ==========================================
# 2. MOTOR DE DECISÃO INTERNO (ENGINE)
# ==========================================
class ChemistryEngine:
    def __init__(self):
        self.db = DATABASE_DB
        self.formulas = FORMULAS_DB
        self.trees = TREES_DB
        self.mistakes = MISTAKES_DB
        self.prompts = PROMPTS_DB
        self.vestibular = VESTIBULAR_DB
        self.exercises = EXERCISES_DB
        self.roadmap = ROADMAP_DB
        self.matrix = MATRIX_DB

    def smart_search(self, query):
        if not query:
            return []
        query = query.lower()
        results = []
        for mod in self.db.get("modulos", []):
            if query in mod["titulo"].lower() or query in mod["descricao"].lower():
                results.append({"tipo": "Módulo", "nome": mod["titulo"], "contexto": mod["descricao"]})
            for top in mod.get("topicos", []):
                if query in top["nome"].lower():
                    results.append({"tipo": "Tópico", "nome": top["nome"], "contexto": "Dentro do módulo: " + mod["titulo"]})
        
        for exe in self.exercises.get("exercicios", []):
            if query in exe["enunciado"].lower():
                results.append({"tipo": "Exercício", "nome": exe["banca_origem"], "contexto": exe["enunciado"][:100] + "..."})
        return results

    def get_adaptive_test(self, performance_profile, num_questions=3):
        all_questions = self.exercises.get("exercicios", [])
        if performance_profile == "Iniciante":
            filtered = [q for q in all_questions if q["dificuldade"] in ["Fácil", "Média"]]
        else:
            filtered = [q for q in all_questions if q["dificuldade"] in ["Média", "Difícil"]]
        
        if not filtered:
            filtered = all_questions
        return random.sample(filtered, min(len(filtered), num_questions))

    def resolve_matrix(self, matrix_id):
        match = next((item for item in self.matrix.get("matriz", []) if item["id"] == matrix_id), None)
        if not match:
            return None
        
        formula = next((f for f in self.formulas.get("formulas", []) if f["id"] == match["formula_id"]), None)
        strategy = next((p for p in self.prompts.get("padroes", []) if p["id"] == match["estrategia_id"]), None)
        trap = next((m for m in self.mistakes.get("armadilhas", []) if m["id"] == match["erro_id"]), None)
        
        return {
            "problema": match["problema"],
            "formula": formula,
            "estrategia": strategy,
            "armadilha": trap
        }

# ==========================================
# 3. INTERFACE DE USUÁRIO (STREAMLIT APP)
# ==========================================
st.set_page_config(page_title="Química Objetivo - 1° Ano", layout="wide", initial_sidebar_state="expanded")

@st.cache_resource
def get_engine():
    return ChemistryEngine()

engine = get_engine()

st.sidebar.title("⚛️ Química - Objetivo 1° Ano")
st.sidebar.markdown("Plataforma Unificada e Inteligente de Aprendizagem de Química Geral.")

menu = st.sidebar.radio(
    "Navegação Estruturada",
    ["Painel Geral & Roadmap", "Árvores de Decisão", "Matriz Problema-Solução", "Simulado Adaptativo", "Flashcards Interativos", "Busca Unificada"]
)

if menu == "Painel Geral & Roadmap":
    st.title("🗺️ Mapa de Bordo & Grade Curricular")
    st.subheader("Conteúdo Programático Oficial - Padrão Colégio Objetivo")
    
    for mod in engine.db.get("modulos", []):
        with st.expander(f"📦 {mod['id']} - {mod['titulo']}", expanded=True):
            st.markdown(f"*{mod['descricao']}*")
            for top in mod.get("topicos", []):
                col1, col2, col3 = st.columns([2, 1, 1])
                col1.write(f"🔹 **{top['nome']}**")
                col2.write(f"🎯 ENEM: `{top['importancia_enem']}`")
                col3.write(f"🏫 FUVEST: `{top['importancia_fuvest']}`")

    st.markdown("---")
    st.subheader("🔗 Conexões e Dependências Cognitivas (Roadmap)")
    for conn in engine.roadmap.get("mapa_mental", {}).get("conexoes", []):
        st.info(f"**{conn['origem']}** ➔ **{conn['destino']}** | *{conn['tipo']}*: {conn['detalhes']}")

elif menu == "Árvores de Decisão":
    st.title("🌳 Algoritmos e Árvores de Decisão para Resolução")
    st.markdown("Utilize os diagramas lógicos estruturados para resolver desafios conceituais comuns.")
    
    selected_tree = st.selectbox("Escolha uma árvore de diagnóstico:", [t["titulo"] for t in engine.trees.get("arvores", [])])
    tree_obj = next(t for t in engine.trees.get("arvores", []) if t["titulo"] == selected_tree)
    
    st.write("### Estrutura de Decisão:")
    st.json(tree_obj["raiz"])

elif menu == "Matriz Problema-Solução":
    st.title("🧮 Matriz de Resolução Inteligente (Problema ➔ Fórmula ➔ Estratégia)")
    st.markdown("Navegue pela correlação analítica multidimensional para mitigar erros recorrentes nos grandes vestibulares.")
    
    for item in engine.matrix.get("matriz", []):
        resolved = engine.resolve_matrix(item["id"])
        with st.container():
            st.markdown(f"### 🎯 Desafio: {resolved['problema']}")
            c1, c2, c3 = st.columns(3)
            
            with c1:
                st.warning("⚠️ Armadilha / Erro Comum")
                if resolved["armadilha"]:
                    st.markdown(f"**{resolved['armadilha']['nome']}**\n\n{resolved['armadilha']['descricao']}")
                else:
                    st.write("Sem armadilhas críticas mapeadas.")
                    
            with c2:
                st.success("📐 Fórmula / Relação")
                if resolved["formula"]:
                    st.markdown(f"**{resolved['formula']['nome']}**\n\n`{resolved['formula']['expressao']}`\n\n*{resolved['formula']['aplicacao']}*")
                else:
                    st.write("Conceitual - Não requer fórmula.")
                    
            with c3:
                st.info("💡 Estratégia de Leitura")
                if resolved["estrategia"]:
                    st.markdown(f"**Contexto:** {resolved['estrategia']['contexto']}\n\n**Gatilhos:** {', '.join(resolved['estrategia']['gatilhos'])}")
                else:
                    st.write("Geral aplicada.")
            st.markdown("---")

elif menu == "Simulado Adaptativo":
    st.title("🧠 Simulado Adaptativo Computacional")
    perfil = st.select_slider("Selecione seu nível de proficiência atual:", options=["Iniciante", "Avançado"])
    
    if st.button("Gerar Bloco de Questões Adaptativas"):
        st.session_state.questions = engine.get_adaptive_test(perfil)
        st.session_state.answers = {}
        st.session_state.submitted = False

    if "questions" in st.session_state:
        for idx, q in enumerate(st.session_state.questions):
            st.markdown(f"#### Questão {idx+1} ({q['banca_origem']} - Nível {q['dificuldade']})")
            st.markdown(q["enunciado"])
            opts = [f"{k}) {v}" for k, v in q["alternativas"].items()]
            ans = st.radio(f"Selecione a alternativa para a questão {idx+1}:", opts, key=f"q_{idx}")
            st.session_state.answers[idx] = ans[0]
            st.markdown("---")
            
        if st.button("Submeter Respostas e Ver Análise"):
            st.session_state.submitted = True
            
        if st.session_state.get("submitted"):
            st.subheader("📊 Relatório de Desempenho Analítico")
            for idx, q in enumerate(st.session_state.questions):
                user_ans = st.session_state.answers.get(idx)
                correct = q["resposta_correta"]
                if user_ans == correct:
                    st.success(f"Questão {idx+1}: Resposta Correta ({user_ans})")
                else:
                    st.error(f"Questão {idx+1}: Você marcou {user_ans}. O gabarito é {correct}.")
                st.markdown(f"**Resolução:** {q['resolucao']}")
                st.markdown("---")

elif menu == "Flashcards Interativos":
    st.title("🗂️ Sistema de Revisão Espaçada (Flashcards)")
    cards = engine.db.get("flashcards", [])
    
    if "card_idx" not in st.session_state:
        st.session_state.card_idx = 0
        st.session_state.reveal = False
        
    if st.session_state.card_idx < len(cards):
        card = cards[st.session_state.card_idx]
        st.info(f"### Pergunta:\n{card['frente']}")
        
        if st.button("Revelar Verso (Resposta)"):
            st.session_state.reveal = True
            
        if st.session_state.reveal:
            st.success(f"### Resposta:\n{card['verso']}")
            if st.button("Próximo Card"):
                st.session_state.card_idx += 1
                st.session_state.reveal = False
                st.rerun()
    else:
        st.balloons()
        st.success("Você concluiu todos os flashcards do módulo!")
        if st.button("Reiniciar"):
            st.session_state.card_idx = 0
            st.session_state.reveal = False
            st.rerun()

elif menu == "Busca Unificada":
    st.title("🔍 Sistema de Busca Inteligente Indexada")
    query = st.text_input("Digite o termo, conceito ou palavra-chave (ex: 'mistura', 'ferro', 'ebulição'):")
    if query:
        res = engine.smart_search(query)
        st.write(f"Encontrados {len(res)} resultados correspondentes:")
        for r in res:
            with st.chat_message("assistant"):
                st.markdown(f"**[{r['tipo']}]** {r['nome']}")
                st.markdown(f"Contexto: {r['contexto']}")
