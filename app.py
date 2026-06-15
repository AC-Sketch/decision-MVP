import streamlit as st
import streamlit.components.v1 as components
import html
from statistics import mean

# ==============================================================================
# 1. CONFIGURAÇÃO DE INTERFACE E DESIGN SYSTEM PREMIUM (UI/UX)
# ==============================================================================
st.set_page_config(
    page_title="War Room Executivo - Recrutamento Sênior",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Injeção de CSS Executivo baseado no seu modelo estável
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');

html, body, [data-testid="stAppViewContainer"] {
    font-family: 'Inter', sans-serif;
    background-color: #F8FAFC;
    color: #1E293B;
}

.block-container {
    padding-top: 1.5rem;
    padding-bottom: 1.5rem;
    max-width: 1600px;
}

.main-title {
    font-size: 2.4rem;
    font-weight: 800;
    color: #0F172A;
    letter-spacing: -0.04em;
    margin-bottom: 0.2rem;
}

.subtitle {
    font-size: 1.1rem;
    color: #64748B;
    margin-bottom: 1.8rem;
    font-weight: 400;
}

.avatar-container {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 12px;
    background: #FFFFFF;
    border-radius: 10px;
    margin-bottom: 8px;
    border: 1px solid #E2E8F0;
    box-shadow: 0 2px 4px rgba(0,0,0,0.01);
}

.avatar-img {
    font-size: 22px;
    background: #F1F5F9;
    padding: 6px;
    border-radius: 50%;
    width: 38px;
    height: 38px;
    display: flex;
    align-items: center;
    justify-content: center;
}

.response-box {
    background-color: #e8f8f5;
    border-left: 4px solid #18bc9c;
    padding: 14px !important;
    border-radius: 8px;
    margin-top: 0.4rem;
    margin-bottom: 0.4rem;
}

.logs-box {
    background-color: #0F172A;
    color: #94A3B8;
    font-family: 'Courier New', Courier, monospace;
    padding: 14px;
    border-radius: 10px;
    max-height: 220px;
    overflow-y: auto;
    font-size: 11.5px;
    border-left: 4px solid #3B82F6;
}

.log-entry {
    margin-bottom: 5px;
    border-bottom: 1px solid #1E293B;
    padding-bottom: 4px;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 10px;
    background-color: #E2E8F0;
    padding: 6px;
    border-radius: 10px;
}

.stTabs [data-baseweb="tab"] {
    background-color: transparent;
    border-radius: 6px;
    padding: 10px 20px;
    font-weight: 600;
    color: #475569;
}

.stTabs [data-baseweb="tab"][aria-selected="true"] {
    background-color: #FFFFFF;
    color: #0F172A !important;
    box-shadow: 0 2px 4px rgba(0,0,0,0.05);
}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. MAPEAMENTO DOS TÓPICOS DA ENTREVISTA (INTERFACE EM BLOCOS)
# ==============================================================================
TOPICOS_PAINEL = {
    0: {"titulo": "Fale sobre você (Senioridade)", "icon": "👤", "id_ref": 1},
    1: {"titulo": "Por que a IURD?", "icon": "🏛️", "id_ref": 2},
    2: {"titulo": "Visão sobre ETL e Migrações", "icon": "🔄", "id_ref": 3},
    3: {"titulo": "Sua Proposta de Valor", "icon": "💎", "id_ref": 4},
    4: {"titulo": "Expectativa Salarial / Contrato", "icon": "💰", "id_ref": 5},
    5: {"titulo": "Vivência com Oracle Fusion Cloud", "icon": "☁️", "id_ref": 6},
    6: {"titulo": "Domínio Avançado de SQL / PL-SQL", "icon": "💾", "id_ref": 7},
    7: {"titulo": "Você é superqualificado?", "icon": "🚀", "id_ref": 8},
    8: {"titulo": "Explicação sobre Ciclos Curtos", "icon": "⏱️", "id_ref": 9},
    9: {"titulo": "Evolução Profissional e Engenharia", "icon": "📐", "id_ref": 10},
    10: {"titulo": "Case ASICS (Otimização e FinOps)", "icon": "👟", "id_ref": 11},
    11: {"titulo": "Case Itaú / NTT DATA (Volumetria)", "icon": "🏦", "id_ref": 12},
    12: {"titulo": "Case Afinz (Automação de ETL)", "icon": "⚡", "id_ref": 14},
    13: {"titulo": "Case Heineken (Normalização)", "icon": "🍺", "id_ref": 13},
    14: {"titulo": "Case Burity (Capacidade Investigativa)", "icon": "🔍", "id_ref": 15},
    15: {"titulo": "Tratamento de Inconsistências Críticas", "icon": "🚨", "id_ref": 16},
    16: {"titulo": "Demandas Não Mapeadas sob Pressão", "icon": "🔥", "id_ref": 17},
    17: {"titulo": "Alinhamento com Áreas de Negócio", "icon": "🤝", "id_ref": 18},
}

# ==============================================================================
# 3. BASE ESTRUTURADA DE RESPOSTAS SÊNIOR (CONTEÚDO DA VAGA)
# ==============================================================================
DATA_MAPPING = {
    1: {
        "title": "Fale sobre você (Senioridade)",
        "category": "WHAT - Capabilities & Profile",
        "tag": "PERFIL",
        "bridge": "Sou Engenheiro de Produção com MBA pela FGV e mais de 7 anos de experiência sólida na área de dados, especializado em transformar grandes volumes de dados brutos e dispersos em pipelines de ETL altamente performáticos e fontes de verdade unificadas.",
        "followup": "Minha trajetória une o raciocínio lógico da engenharia com a visão de governança corporativa. Sou especialista em otimização de queries complexas in SQL, automação de fluxos com Python e sustentação de ecossistemas híbridos envolvendo desde bancos tradicionais como Oracle e PostgreSQL até NoSQL como MongoDB.",
        "case": "Engenharia + MBA + Liderança Técnica de Dados (Heineken, Itaú, ASICS, NTT DATA).",
        "bullets": [
            "Lidero a arquitetura e otimização de dados em ambientes corporativos de alta complexidade regulatória e volumetria (Itaú, Heineken).",
            "Tenho facilidade para navegar entre o desenvolvimento de código robusto (Python/SQL) e a entrega visual para a diretoria (Dashboards).",
            "Atuo ativamente com foco em governança de dados, documentação técnica estruturada e garantia de qualidade (Data Quality)."
        ],
        "qa_responses": [
            {"q": "Qual é o diferencial do seu perfil para uma posição sênior?", "a": "Eu não sou apenas um construtor de queries. Eu entendo o impacto do dado no negócio. Garanto que o dado na ponta — seja no Apache Superset ou no Power BI — seja idêntico ao payload bruto do banco, desenhando arquiteturas que não geram gargalos de infraestrutura."},
            {"q": "Como você lida com ambientes dinâmicos e pilhas tecnológicas variadas?", "a": "Com flexibilidade e mentalidade investigativa. Se o fluxo exige extrair dados de uma API REST via Python, tratar no banco Oracle via PL/SQL e disponibilizar no MongoDB, eu estruturo a esteira focando sempre em performance e manutenibilidade."}
        ]
    },
    2: {
        "title": "Por que a IURD?",
        "category": "WHY - Intent & Fit",
        "tag": "ESTRATÉGIA",
        "bridge": "Uma instituição desse porte gerencia diariamente um volume massivo e heterogêneo de dados operacionais, financeiros e de sistemas institucionais, o que exige uma governança de dados impecável.",
        "followup": "Quero aplicar meu toolkit de engenharia de dados e otimização de infraestrutura exatamente onde o dado dita o ritmo da eficiência das operações internas e do atendimento às áreas finalísticas.",
        "case": "Cenários de Alta Volumetria e Ecossistemas Híbridos.",
        "bullets": [
            "Identifico-me com culturas que valorizam o sentimento de dono, a resiliência e a busca por melhoria contínua.",
            "O desafio de atuar com múltiplos bancos de dados (Oracle, Postgres, MariaDB, Mongo) e ferramentas de BI variadas me motiva intelectualmente.",
            "Vejo uma oportunidade clara de gerar impacto imediato na otimização de fluxos legados e na eficiência do processamento de dados."
        ],
        "qa_responses": [
            {"q": "O que mais te atrai nesta oportunidade?", "a": "A oportunidade de atuar de ponta a ponta: desde a extração via APIs ou Apache NiFi, passando pela transformação pesada em banco com PL/SQL, até a camada final de entrega no Oracle Fusion Cloud e Superset. É o cenário perfeito para um profissional sênior."}
        ]
    },
    3: {
        "title": "Visão sobre ETL e Migrações",
        "category": "WHY - Intent & Fit",
        "tag": "ARQUITETURA",
        "bridge": "Enxergo os processos de ETL/ELT como a espinha dorsal de qualquer tomada de decisão corporativa; eles precisam ser invisíveis, rápidos e fáceis de manter.",
        "followup": "Sustentar fluxos existentes (como no Apache NiFi) exige respeito ao legado, mas migrar gradualmente para pipelines em Python é o caminho ideal para garantir escalabilidade e testes automatizados.",
        "case": "Migração e Automação de Pipelines (Case Afinz/Stalse).",
        "bullets": [
            "Tenho experiência prática na reengenharia de processos manuais ou legados para rotinas automatizadas em Python.",
            "Utilizo bibliotecas consagradas como pandas, sqlalchemy, requests (APIs) e openpyxl para construir conexões seguras e limpas.",
            "Meu foco em ETL é garantir a integridade total do dado (traceabilidade) antes que ele atinja as camadas de visualização."
        ],
        "qa_responses": [
            {"q": "Qual o seu plano para apoiar nossa migração de ETL para Python?", "a": "O primeiro passo é documentar e auditar os fluxos atuais no Apache NiFi para mapear dependências. Em seguida, traduzo as transformações lógicas para scripts Python modulares e performáticos, utilizando sqlalchemy para a escrita eficiente nos bancos de destino, garantindo zero downtime."}
        ]
    },
    4: {
        "title": "Sua Proposta de Valor",
        "category": "WHAT - Capabilities & Profile",
        "tag": "VALOR",
        "bridge": "Ofereço o domínio técnico avançado em SQL/Python necessário para resolver gargalos de performance e a maturidade analítica para dialogar diretamente com as áreas de negócio.",
        "followup": "Elimino o abismo que costuma existir entre o que a equipe de TI desenvolve e o que os diretores e usuários de negócio realmente precisam enxergar nos relatórios.",
        "case": "FinOps, Performance Tuning e Governança de Dados.",
        "bullets": [
            "Consolido dados dispersos para gerar relatórios corporativos com precisão matemática centesimal.",
            "Otimizo consultas SQL lentas através de análise de planos de execução (Performance Tuning), reduzindo consumo de memória e CPU.",
            "Entrego documentação técnica clara de dados e processos, mitigando o risco de perda de conhecimento."
        ],
        "qa_responses": [
            {"q": "O que você consegue entregar nos primeiros 30 dias se for contratado?", "a": "Focar na imersão dos ambientes Oracle e PostgreSQL, assumir a sustentação das cargas diárias do Apache NiFi e mapear as principais dores dos usuários nos relatórios do OTBI / BI Publisher, garantindo estabilidade imediata na operação."}
        ]
    },
    5: {
        "title": "Expectativa Salarial / Contrato",
        "category": "WHY - Intent & Fit",
        "tag": "ANCORAGEM",
        "bridge": "Minha pretensão salarial está fundamentada na minha senioridade de mais de 7 anos e no valor imediato que posso gerar na otimização da infraestrutura de dados e relatórios da instituição.",
        "followup": "Busco uma remuneração justa para um Especialista que atua no modelo Prestador de Serviços (PJ) com total autonomia e prontidão técnica.",
        "case": "Contratação PJ Especialista - Período Integral.",
        "bullets": [
            "Minha pretensão salarial para o formato de prestação de serviços (PJ) está na faixa de 12.000 a 15.000 Reais mensais.",
            "Estou totalmente pronto para iniciar o modelo híbrido em São Paulo - SP com total disponibilidade.",
            "Possuo empresa aberta (CNPJ) regularizada com emissão de nota fiscal imediata."
        ],
        "qa_responses": [
            {"q": "Este valor é negociável?", "a": "Estou aberto a entender o pacote completo oferecido pela instituição, os desafios de longo prazo e as possibilidades de evolução dentro do ecossistema técnico. Havendo sinergia, o valor pode ser ajustado."}
        ]
    },
    6: {
        "title": "Vivência com Oracle Fusion Cloud",
        "category": "WHAT - Capabilities & Profile",
        "tag": "ORACLE-FUSION",
        "bridge": "Considero o ecossistema Oracle Fusion Cloud, especificamente OTBI e BI Publisher, uma camada estratégica vital, onde o SQL robusto é a chave para destravar relatórios corporativos complexos.",
        "followup": "Muitos analistas focam apenas em ferramentas de visualização modernas, mas eu domino a extração nativa dentro do ambiente corporativo Oracle, manipulando data models complexos e views customizadas.",
        "case": "Criação, Ajuste e Manutenção de Relatórios OTBI / BI Publisher.",
        "bullets": [
            "Compreendo a estrutura de dados interna dos módulos Oracle ERP/Fusion para localizar tabelas e campos rapidamente.",
            "Utilizo SQL avançado para customizar os Data Models que alimentam os layouts do BI Publisher.",
            "Sei apoiar as áreas de negócio na criação de análises em tempo real utilizando as áreas de assunto (Subject Areas) do OTBI."
        ],
        "qa_responses": [
            {"q": "O que você faz quando um relatório do BI Publisher apresenta lentidão?", "a": "O problema quase sempre está no Data Model subjacente. Extraio a query SQL contida nele, analiso o plano de execução dentro do ambiente Oracle e faço o refactoring da consulta."}
        ]
    },
    7: {
        "title": "Domínio Avançado de SQL / PL-SQL",
        "category": "WHAT - Capabilities & Profile",
        "tag": "SQL-TUNING",
        "bridge": "Para mim, SQL não é apenas escrever SELECTs básicos; é dominar joins complexos, CTEs, subqueries correlacionadas e funções analíticas para processar dados com máxima performance.",
        "followup": "Em ambientes com grande volume de dados, o design da query dita o custo e a velocidade da entrega. Escrevo códigos limpos, estruturados e fáceis de auditar.",
        "case": "Performance Tuning e Otimização de Consultas Complexas.",
        "bullets": [
            "Utilizo CTEs (Common Table Expressions) extensivamente para modularizar queries longas e torná-las legíveis.",
            "Domino funções analíticas (ROW_NUMBER, RANK, LEAD, LAG) para evitar subqueries pesadas.",
            "Possuo conhecimento em PL/SQL para criar triggers, procedures e funções que automatizam lógicas direto no Oracle."
        ],
        "qa_responses": [
            {"q": "Qual a sua experiência com bancos NoSQL como MongoDB?", "a": "Utilizo o MongoDB para cenários onde a estrutura do dado é altamente fluida ou semiestruturada (JSONs). Sei trabalhar com o framework de agregação do Mongo e utilizo Python para normalizá-los se necessário."}
        ]
    },
    8: {
        "title": "Você é superqualificado?",
        "category": "WHY - Intent & Fit",
        "tag": "RETENÇÃO",
        "bridge": "Acredito que o termo correto não é superqualificado, mas sim plenamente preparado para os desafios de alta complexidade técnica que a instituição possui.",
        "followup": "Um profissional sênior não busca apenas tarefas complexas; busca estabilidade, processos organizados e a oportunidade de construir pipelines eficientes que resolvam problemas reais.",
        "case": "Maturidade Profissional e Foco em Soluções Duradouras.",
        "bullets": [
            "Tenho real motivação em atuar na sustentação técnica e melhoria contínua de ecossistemas maduros.",
            "Para mim, o desafio intelectual está em otimizar rotinas que hoje demoram horas para rodar em poucos minutos.",
            "Busco um vínculo de longo prazo onde minha senioridade técnica possa apoiar o crescimento da equipe."
        ],
        "qa_responses": [
            {"q": "Sendo sênior, você aceitaria realizar tarefas mais operacionais?", "a": "Com certeza. Minha senioridade me dá a maturidade de entender que o suporte a relatórios corporativos existentes é fundamental para a governança da instituição. Encaro isso com total senso de dono."}
        ]
    },
    9: {
        "title": "Explicação sobre Ciclos Curtos",
        "category": "WHAT - Capabilities & Profile",
        "tag": "PROJETOS",
        "bridge": "Minhas passagens recentes por empresas como Stalse (atendendo ASICS) e NTT DATA (atendendo Itaú) foram alocações estratégicas focadas em projetos com escopo fechado e entregas de tiro curto.",
        "followup": "Atuei como um acelerador de soluções, resolvendo gargalos de arquitetura específicos e realizando migrações críticas que demandavam força técnica sênior imediata.",
        "case": "Sprints Ágeis e Consultoria Técnica de Dados.",
        "bullets": [
            "Na *Stalse*, unifiquei dados financeiros internacionais da *ASICS* e apliquei conceitos de FinOps.",
            "Na *NTT DATA*, trabalhei imerso no ambiente AWS Cloud do *Itaú*, lidando com volumetria massiva.",
            "Agora, meu objetivo estratégico é fixar minhas habilidades em uma posição estável de longo prazo."
        ],
        "qa_responses": [
            {"q": "Por que seus últimos contratos duraram cerca de 4 meses?", "a": "Ambos foram contratos de consultoria por escopo. Concluídas as automações de pipelines, lançamento de dashboards e documentações completas, o ciclo do projeto encerrou com sucesso."}
        ]
    },
    10: {
        "title": "Evolução Profissional e Engenharia",
        "category": "WHY - Intent & Fit",
        "tag": "EVOLUÇÃO",
        "bridge": "Minha transição da análise de gestão tradicional para a engenharia e análise avançada de dados reflete o movimento natural do mercado corporativo moderno.",
        "followup": "Minha formação em Engenharia de Produção me deu a mentalidade de processos e eficiência, enquanto o MBA na FGV consolidou a visão executiva de indicadores. Unir isso ao desenvolvimento em Python e SQL foi o passo lógico.",
        "case": "Combinação de Gestão Estruturada com Hard Skills de TI.",
        "bullets": [
            "Minha formação pelo CREA-SP garante meu foco em métodos ágeis e eliminação de desperdícios.",
            "Desenvolvi a capacidade investigativa ponta a ponta através de anos auditando processos corporativos.",
            "Enxergo os dashboards no Apache Superset ou Power BI como ferramentas de gestão ativa."
        ],
        "qa_responses": [
            {"q": "Por que focar em ferramentas como Apache Superset?", "a": "O Apache Superset é uma ferramenta fantástica, open-source, extremamente performática para grandes volumes. Tenho facilidade em plugá-lo sobre bancos PostgreSQL ou DWs para democratizar o acesso sem custos abusivos."}
        ]
    },
    11: {
        "title": "Case ASICS (Otimização e FinOps)",
        "category": "HOW - Case Methodology (STAR)",
        "tag": "CASE-ETL",
        "bridge": "Na Stalse, liderei a reengenharia de uma arquitetura de dados financeiros e operacionais para a operação Latam da ASICS (Brasil, Chile e Colômbia), focando em performance de queries e redução drástica de custos.",
        "followup": "O cenário era de fragmentação de dados e consultas extremamente pesadas que geravam um consumo financeiro excessivo nas plataformas de nuvem (BigQuery).",
        "case": "Redução de Custos e Unificação de Dados Financeiros.",
        "bullets": [
            "Situação: Queries ineficientes consumiam Gigabytes de processamento desnecessário, gerando lentidão e custos elevados.",
            "Action: Refatorei as consultas SQL aplicando boas práticas de particionamento, CTEs e eliminando joins redundantes.",
            "Result: Reduzi o consumo de dados de escala de Gigabytes para Megabytes, otimizando o tempo de resposta significativamente."
        ],
        "qa_responses": [
            {"q": "Como você aplicou a otimização de dados na prática?", "a": "Substituí subqueries pesadas dentro de loops por tabelas temporárias bem indexadas e queries utilizando funções analíticas. Isso acelerou drasticamente a atualização."}
        ]
    },
    12: {
        "title": "Case Itaú / NTT DATA (Volumetria AWS)",
        "category": "HOW - Case Methodology (STAR)",
        "tag": "CASE-SCALE",
        "bridge": "Alocado na NTT DATA para atender o Itaú, atuei imerso em ambiente de computação em nuvem (AWS), construindo visões SQL complexas e robustas sobre tabelas com bilhões de registros.",
        "followup": "O projeto exigia precisão absoluta e velocidade de processamento para atender métricas de desempenho corporativo consultadas por mais de 5.000 executivos e gestores.",
        "case": "Construção de Views Complexas em Amazon Athena e AWS S3.",
        "bullets": [
            "Situação: Havia a necessidade de calcular métricas gerenciais complexas cruzando tabelas gigantescas de logs.",
            "Action: Desenvolvi visões SQL estruturadas utilizando Amazon Athena, aplicando estratégias de filtros eficientes e parametrização dinâmica.",
            "Result: Entreguei painéis automatizados estáveis, garantindo consistência total do dado."
        ],
        "qa_responses": [
            {"q": "Como manter as queries rápidas trabalhando com bilhões de linhas?", "a": "A chave foi trabalhar em sintonia com a estrutura de particionamento dos dados no S3. Garantindo o partition pruning correto, o Athena lia apenas o necessário, reduzindo o tempo para segundos."}
        ]
    },
    14: {
        "title": "Case Afinz (Automação de ETL 1h30 para 15m)",
        "category": "HOW - Case Methodology (STAR)",
        "tag": "CASE-AUTOMATION",
        "bridge": "Como Analista de MIS na Afinz/Sorocred, identifiquei um fluxo crítico de relatórios diários que era executado de forma manual, gerando atrasos crônicos na disponibilização de dados operacionais.",
        "followup": "A rotina dependia de operadores compilando dados manualmente em Excel e disparando queries fragmentadas, o que aumentava drasticamente o risco de erros humanos.",
        "case": "Redução do Tempo de Processamento de ETL de 1h30 para 15 minutos.",
        "bullets": [
            "Situação: Rotinas manuais consumiam 1 hora e meia diária da equipe técnica e atrasavam as análises.",
            "Action: Desenvolvi pipelines de ETL automatizados de ponta a ponta utilizando Python e consultas SQL diretas.",
            "Result: Reduzi o tempo total de execução para apenas 15 minutos, eliminando as inconsistências das cargas."
        ],
        "qa_responses": [
            {"q": "Como essa mentalidade se aplica à migração de fluxos do Apache NiFi?", "a": "A lógica é análoga. No NiFi, você tem os processadores visuais. Para migrar para Python de forma performática, eu analiso o comportamento lógico de cada processador e reescrevo nativamente de forma modular."}
        ]
    },
    13: {
        "title": "Case Heineken (Normalização de Dados)",
        "category": "HOW - Case Methodology (STAR)",
        "tag": "CASE-BI",
        "bridge": "Na Heineken, atuando na divisão digital de e-commerce e eRetail, fui responsável por consolidar e normalizar grandes conjuntos de dados provenientes de múltiplos clientes e parceiros externos.",
        "followup": "Os dados chegavam em formatos totalmente heterogêneos (APIs variadas, arquivos de texto, planilhas bagunçadas), o que impedia um acompanhamento unificado das campanhas comerciais.",
        "case": "Modelagem Star Schema e Integração de Múltiplas Fontes de Clientes.",
        "bullets": [
            "Situação: A falta de padronização nas fontes externas gerava falhas de integração frequentes.",
            "Action: Desenhei um modelo relacional robusto no padrão Star Schema (Fatos e Dimensões) e apliquei validações estritas.",
            "Result: Montei painéis corporativos estáveis no Power BI que batiam com os indicadores financeiros até a última casa."
        ],
        "qa_responses": [
            {"q": "Qual a vantagem de usar a modelagem Star Schema nesse cenário?", "a": "O Star Schema simplifica os Joins na hora de construir as métricas. Ao separar dados transacionais dos cadastrais, as queries rodam infinitamente mais rápido, melhorando a performance geral."}
        ]
    },
    15: {
        "title": "Case Burity (Capacidade Investigativa Legal)",
        "category": "HOW - Case Methodology (STAR)",
        "tag": "CASE-AUDIT",
        "bridge": "Durante minha longa e sólida trajetória na Burity Empresarial, atuei na gestão de ativos e conformidade regulatória como Procurador Legal, desarrollando uma capacidade investigativa ponta a ponta e um olhar cirúrgico para mitigação de riscos operacionais.",
        "followup": "Fui responsável por auditar processos administrativos complexos, contratos corporativos de alto valor e plantas técnicas de engenharia, retificando erros históricos e eliminando passivos com órgãos governamentais.",
        "case": "Retificação de Registros, Auditoria de Processos e Risco Zero.",
        "bullets": [
            "Situação: Divergências descritivas em cadastros históricos geravam riscos de multas severas e travavam ativos.",
            "Action: Liderei auditorias documentais profundas, cruzei dados técnicos e coordenei times multidisciplinares.",
            "Result: Garanti a conformidade jurídica e patrimonial dos ativos de forma 100% administrativa, com zero litígios."
        ],
        "qa_responses": [
            {"q": "Como essa experiência jurídica se aplica a uma vaga estritamente de dados?", "a": "A essência da auditoria e investigação é rigorosamente a mesma. Investigar um registro omisso exige o mesmo nível de atenção, ceticismo e busca por evidências do que encontrar um bug oculto em um log de banco de dados."}
        ]
    },
    16: {
        "title": "Tratamento de Inconsistências Críticas",
        "category": "WHEN - Extreme Scenarios & Crisis",
        "tag": "CRISE-DADO",
        "bridge": "Se um fluxo de integração quebrar na madrugada ou os dados do dashboard amanhecerem duplicados, minha postura sênior imediata é conter o impacto e isolar a falha com total transparência.",
        "followup": "Não busco culpados; busco logs. Identifico o range de dados afetado, executo um script de reversão limpo e aplico a correção definitiva no pipeline para que o erro nunca mais se repita.",
        "case": "Gestão de Incidentes em Ambientes de Produção.",
        "bullets": [
            "Mantenho a calma, desligo gatilhos automáticos problemáticos e analiso os logs de erro do Python ou do Apache NiFi.",
            "Desenvolvo scripts rápidos de remediação para limpar duplicidades respeitando chaves primárias e constraints do banco.",
            "Registro um post-mortem técnico detalhando o ocorrido e a solução aplicada para alimentar a base de conhecimento."
        ],
        "qa_responses": [
            {"q": "Você tem autonomia para debugar problemas em servidores Linux?", "a": "Sim. Tenho excelente familiaridade com comandos de terminal Linux para navegar em servidores, checar uso de memória de processos Python, olhar logs de Docker e verificar agendamentos de tarefas (cronjobs)."}
        ]
    },
    17: {
        "title": "Demandas Não Mapeadas sob Pressão",
        "category": "WHEN - Extreme Scenarios & Crisis",
        "tag": "FLEXIBILIDADE",
        "bridge": "Diante de solicitações urgentes de indicadores feitas pela diretoria em cenários caóticos, utilizo o pensamento estruturado para focar no MVP (Mínimo Produto Viável) do dado.",
        "followup": "Isolo o ruído emocional, extraio uma amostra confiável diretamente via SQL dos bancos relacionais e apresento um panorama claro do risco ou oportunidade com os dados disponíveis no momento.",
        "case": "Extrações Rápidas de Emergência para Apoio de Decisão.",
        "bullets": [
            "Alinho com a liderança as prioridades para entender o núcleo real da necessidade do negócio.",
            "Escrevo consultas SQL otimizadas usando indexação adequada para não derrubar a performance do banco operacional.",
            "Entrego o resultado de forma limpa, apontando as premissas adotadas e eventuais limitações técnicas."
        ],
        "qa_responses": [
            {"q": "O que você faz se precisar utilizar uma tecnologia que nunca viu na vida?", "a": "Abordo com curiosidade técnica e proatividade investigativa. Como sênior, compreendo os fundamentos de lógica e engenharia. Aprender a sintaxe de uma ferramenta nova é apenas questão de ler a documentação."}
        ]
    },
    18: {
        "title": "Alinhamento com Áreas de Negócio",
        "category": "WHEN - Extreme Scenarios & Crisis",
        "tag": "COMUNICAÇÃO",
        "bridge": "Não discuto dados com base em achismos ou opiniões subjetivas; trago fatos, volumetria e métricas de qualidade para a mesa para alinhar equipes multidisciplinares.",
        "followup": "Geralmente, as áreas de negócio criam atritos com a TI porque sentem falta de agilidade ou não entendem as restrições técnicas. Atuo como o tradutor ideal entre esses dois mundos.",
        "case": "Construção de Pontes Técnicas entre TI e Usuários de Negócios.",
        "bullets": [
            "Escuto ativamente as necessidades dos usuários para entender quais dores de negócio eles tentam sanar.",
            "Apresento protótipos rápidos de dashboards no Apache Superset para validar o layout antes de fechar o código backend.",
            "Explico limitações de infraestrutura de forma simples, mostrando como uma query otimizada protege a velocidade deles."
        ],
        "qa_responses": [
            {"q": "Como você reage quando um gestor exige um indicador inviável na estrutura atual?", "a": "Apresento o mapeamento atual nos bancos e explico o gap. Proponho uma alternativa provisória e me coloco à disposição para apoiar a evolução do sistema de origem para passarmos a capturar esse dado."}
        ]
    }
}

# ==============================================================================
# 4. CONTROLADOR DE ESTADOS (ESTRUTURA ANTI-BUG DE ATUALIZAÇÃO)
# ==============================================================================
if "posicao_ativa" not in st.session_state:
    st.session_state.posicao_ativa = 0
if "historico_eventos" not in st.session_state:
    st.session_state.historico_eventos = []

if "matriz_dinamica" not in st.session_state:
    st.session_state.matriz_dinamica = {
        "Cenario_1": [5, 4, 3, 4, 2],
        "Cenario_2": [4, 5, 4, 3, 5],
        "Cenario_3": [3, 4, 5, 5, 4],
        "Cenario_4": [5, 5, 5, 4, 5],
    }

ROTULOS_CENARIOS = {
    "Cenario_1": "Eixo 1: Arquitetura de Pipelines e Engenharia (Python/NiFi)",
    "Cenario_2": "Eixo 2: Otimização de Consultas Pesadas (Performance Tuning SQL)",
    "Cenario_3": "Eixo 3: Relatórios Corporativos e Cloud (Fusion Cloud/Superset)",
    "Cenario_4": "Eixo 4: Governança, LGPD e Interface com Negócio",
}

CRITERIOS_AVALIAÇÃO = [
    "Retorno sobre Eficiência de Infraestrutura (FinOps de Nuvem)",
    "Segurança e Rastreabilidade Lógica do Dado de Origem à Ponta",
    "Capacidade Investigativa e Resolução de Inconsistências Críticas",
    "Aderência Prática aos Requisitos Técnicos da Vaga da IURD",
    "Habilidade de Tradução Técnica de Dados para Nível Executivo"
]

def registrar_evento(texto):
    st.session_state.historico_eventos.insert(0, f"⏱️ Log | {texto}")

# ==============================================================================
# 5. GERADOR DO RELATÓRIO EXECUTIVO (HTML COMPONENT / A4 PAISAGEM)
# ==============================================================================
def gerar_html_boardgame(titulo, objective, contexto, matriz_dados, logs_jogo):
    medias = {k: round(mean(v), 2) for k, v in matriz_dados.items()}
    ranking = sorted(medias.items(), key=lambda x: x[1], reverse=True)
    
    header_alternativas = "".join(f"<th>{html.escape(ROTULOS_CENARIOS[k])}</th>" for k in matriz_dados)
    
    linhas_criterios = ""
    for i, criterio in enumerate(CRITERIOS_AVALIAÇÃO):
        celulas = ""
        for k, notes in matriz_dados.items():
            nota = notes[i]
            cor_fundo = "#EF4444" if nota == 1 else "#FCA5A5" if nota == 2 else "#FEF08A" if nota == 3 else "#86EFAC" if nota == 4 else "#22C55E"
            cor_texto = "white" if nota in [1, 5] else "#1E293B"
            celulas += f'<td style="background-color: {cor_fundo}; color: {cor_texto}; font-weight: bold; text-align: center;">{html.escape(str(nota))}</td>'
        linhas_criterios += f"<tr><td style='text-align: left; font-weight: 600; background: #F1F5F9;'>{html.escape(criterio)}</td>{celulas}</tr>"

    linha_medias = "".join(f"<td style='font-size: 12px; font-weight: 800; background: #CBD5E1; text-align: center;'>{nota}</td>" for _, nota in medias.items())
    ranking_html = "".join(f"<li style='margin-bottom:4px;'><strong>{idx}º Eixo:</strong> {html.escape(ROTULOS_CENARIOS[k])} — <span style='color: #2563EB; font-weight: bold;'>Maturidade: {nota}</span></li>" for idx, (k, nota) in enumerate(ranking, 1))
    
    logs_renderizados = "".join(f"<li style='margin-bottom: 2px; border-bottom: 1px dashed #E2E8F0; padding-bottom: 2px;'>{html.escape(log)}</li>" for log in logs_jogo[:4])
        
    return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
@page {{ size: A4 landscape; margin: 10mm; }}
body {{ font-family: 'Segoe UI', Arial, sans-serif; color: #1E293B; background: #FFF; margin: 0; padding: 0; font-size: 11.5px; line-height: 1.4; }}
.wrapper {{ width: 100%; max-width: 1150px; margin: 0 auto; border: 2px solid #0F172A; padding: 20px; border-radius: 10px; box-shadow: 0 4px 10px rgba(0,0,0,0.03); }}
header {{ border-bottom: 3px solid #0F172A; padding-bottom: 10px; margin-bottom: 15px; }}
.title {{ font-size: 18px; font-weight: 800; color: #0F172A; text-transform: uppercase; letter-spacing: -0.5px; }}
.grid {{ display: grid; grid-template-columns: 1.4fr 1fr; gap: 20px; }}
.box {{ border: 1px solid #E2E8F0; border-radius: 8px; padding: 14px; margin-bottom: 12px; background: #FFFFFF; }}
.box-title {{ font-size: 12px; font-weight: 700; color: #1E3A8A; margin-top: 0; margin-bottom: 10px; border-left: 4px solid #2563EB; padding-left: 8px; text-transform: uppercase; }}
table {{ width: 100%; border-collapse: collapse; font-size: 10.5px; }}
th {{ background: #0F172A; color: white; padding: 10px; text-transform: uppercase; font-size: 9px; border: 1px solid #0F172A; }}
td {{ padding: 8px; border: 1px solid #E2E8F0; }}
.badge-winner {{ background: #EFF6FF; border: 1px solid #BFDBFE; color: #1E40AF; padding: 8px; border-radius: 6px; font-weight: bold; margin-bottom: 10px; font-size: 11.5px; }}
</style>
</head>
<body>
<div class="wrapper">
    <header>
        <div class="title">{html.escape(titulo)}</div>
        <div style="color: #64748B; font-weight: 500; margin-top: 4px;">Escopo Operacional: {html.escape(objective)}</div>
    </header>
    <div class="grid">
        <div>
            <div class="box">
                <div class="box-title">Matriz Verde de Validação de Competências Sênior</div>
                <table>
                    <thead><tr><th>Critérios Analíticos de TI</th>{header_alternativas}</tr></thead>
                    <tbody>{linhas_criterios}<tr style="background: #E2E8F0; font-weight: bold;"><td style="background: #CBD5E1; font-weight:800;">MÉDIA GERAL SÊNIOR</td>{linha_medias}</tr></tbody>
                </table>
            </div>
            <div class="box">
                <div class="box-title">Ordem de Força do Alinhamento Técnico</div>
                <div class="badge-winner">Maior Grau de Domínio: {html.escape(ROTULOS_CENARIOS[ranking[0][0]])}</div>
                <ol style="margin: 0; padding-left: 18px;">{ranking_html}</ol>
            </div>
        </div>
        <div>
            <div class="box">
                <div class="box-title">Parecer Narrativo Técnico de Campo</div>
                <p style="text-align: justify; margin: 0 0 12px 0; color: #334155;">{html.escape(contexto)}</p>
                <div style="background: #F8FAFC; padding: 10px; border-radius: 6px; border: 1px dashed #CBD5E1;">
                    <span style="font-weight: 700; color: #1E3A8A; display: block; margin-bottom: 6px;">Atividades Registradas no Painel:</span>
                    <ul style="margin: 0; padding-left: 16px; color: #475569; font-size: 10.5px;">{logs_renderizados}</ul>
                </div>
            </div>
        </div>
    </div>
</div>
</body>
</html>
"""

# ==============================================================================
# 6. SIDEBAR: MONITOR DE PERFIL DE DADOS
# ==============================================================================
with st.sidebar:
    st.markdown("### 🛠️ Monitor do Cockpit")
    st.markdown("### Status do Candidato")
    st.markdown(f"""
    <div class="avatar-container" style="border-left: 5px solid #3B82F6;">
        <div class="avatar-img">👨‍💻</div>
        <div style="flex-grow: 1;">
            <div style="font-weight: 700; color: #0F172A; font-size: 13px;">André Carvalho</div>
            <div style="font-size: 11px; color: #4A5568;">Analista Sênior</div>
            <div style="font-size: 10px; color: #64748B; font-style: italic;">Oracle, Python, SQL, NiFi</div>
        </div>
        <div style="text-align: right; font-weight: 800; color: #2563EB; font-size: 13px;">
            94%<br>
            <span style='font-size: 10px; color: #94A3B8; font-weight: 500;'>Match</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.caption("**Target:** IURD · Analista de Dados Sênior")

# ==============================================================================
# 7. CORPO CENTRAL E RENDERIZAÇÃO DAS ABAS COMPLETA
# ==============================================================================
st.markdown("<div class='main-title'>Senior Data Analytics Executive Board</div>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>Painel de Simulação de Casos Críticos, Respostas Diretas de Impacto e Métricas de Engenharia</div>", unsafe_allow_html=True)

tab_blocos, tab_vaga, tab_framework = st.tabs([
    "🎯 Painel de Tópicos Interativos",
    "📜 Detalhamento do Perfil Técnico da Vaga",
    "📈 Framework de Arquitetura e Impacto"
])

# --- ABA 1: RENDERIZAÇÃO COM GRID DE MESA SEM PERMITIR VAZAMENTO ---
with tab_blocos:
    st.markdown("#### 🗺️ Mapa Temático (Selecione o bloco para chavear as respostas e queries)")
    
    # Grid de 3 linhas x 6 colunas usando estritamente colunas nativas do Streamlit para evitar quebra estrutural
    for row_idx in range(3):
        cols = st.columns(6)
        for col_idx in range(6):
            n_casa = row_idx * 6 + col_idx
            block_info = TOPICOS_PAINEL[n_casa]
            
            is_active_block = (st.session_state.posicao_ativa == n_casa)
            
            with cols[col_idx]:
                if is_active_block:
                    st.info(f"📍 **{block_info['icon']} Bloco #{n_casa}**\n\n**{block_info['titulo']}**\n\n🟢 *Exibido*")
                else:
                    if st.button(f"{block_info['icon']} #{n_casa}\n{block_info['titulo']}", key=f"btn_nav_{n_casa}", use_container_width=True):
                        st.session_state.posicao_ativa = n_casa
                        registrar_evento(f"Foco do painel alterado para o Bloco #{n_casa}: {block_info['titulo']}")
                        st.rerun()
                        
    st.divider()
    
    # Divisão Operacional 45% (Texto e Estratégia) / 55% (Relatório e HTML Interativo)
    col_mecanica, col_auditoria = st.columns([0.45, 0.55])
    
    current_topic = TOPICOS_PAINEL[st.session_state.posicao_ativa]
    mapped_id = current_topic["id_ref"]
    active_data = DATA_MAPPING.get(mapped_id, DATA_MAPPING[1])
    
    with col_mecanica:
        st.markdown(f"#### 🎯 Tópico Ativo: **{active_data['title']}**")
        st.caption(f"Categoria: `{active_data['category']}` | Tag de Posicionamento: `{active_data['tag']}`")
        
        st.markdown(
            f"""
            <div class="response-box">
                <span style="color:#117a65; font-size:11px; font-weight:bold; text-transform:uppercase;">The Golden Bridge (Resposta Direta Executiva):</span><br>
                <p style="font-size:13.5px; color:#2c3e50; line-height:1.4; margin-top:4px; font-weight:600;">"{active_data['bridge']}"</p>
            </div>
            """, 
            unsafe_allow_html=True
        )
        
        st.markdown(
            f"""
            <div style="background-color: #f1f5f9; padding: 12px; border-radius: 8px; border: 1px solid #e2e8f0; margin-top: 0.4rem;">
                <span style="color:#1e3a8a; font-size:10px; font-weight:bold; text-transform:uppercase;">Aprofundamento Técnico do Caso:</span>
                <p style="font-size:12px; color:#334155; line-height:1.4; margin-top:4px;">{active_data['followup']}</p>
            </div>
            """,
            unsafe_allow_html=True
        )
        
        st.markdown("<p style='font-weight:bold; font-size:12px; color:#1a202c; margin-top:0.6rem; margin-bottom:0.4rem;'>Argumentos de Suporte e Evidências:</p>", unsafe_allow_html=True)
        for bullet in active_data["bullets"]:
            st.markdown(f"<p style='font-size:12px; margin-bottom:4px !important; color: #2d3748;'>• {bullet}</p>", unsafe_allow_html=True)
            
        st.markdown("<p style='font-weight:bold; font-size:11px; color:#2b6cb0; margin-top:0.8rem; margin-bottom:0.4rem; text-transform:uppercase;'>⚡ Simulação de Perguntas Hard Core (C-Level):</p>", unsafe_allow_html=True)
        for qa in active_data["qa_responses"]:
            st.markdown(f"""
            <div style="margin-bottom: 6px; background: white; padding: 10px; border-radius: 6px; border: 1px solid #e2e8f0;">
                <strong style="font-size:11.5px; color:#2c5282; display:block;">Q: {qa['q']}</strong>
                <span style="font-size:11.5px; color:#4a5568; display:block; margin-top:1px;"><strong>A:</strong> {qa['a']}</span>
            </div>
            """, unsafe_allow_html=True)
            
        st.markdown("##### 📟 Histórico Recente de Interações")
        log_str = "".join(f"<div class='log-entry'>{l}</div>" for l in st.session_state.historico_eventos)
        st.markdown(f"<div class='logs-box'>{log_str}</div>", unsafe_allow_html=True)
        
    with col_auditoria:
        st.markdown("#### 📄 Visualizador Corporativo A4 (HTML Renderizado)")
        tx_tit = st.text_input("Título do Parecer de Avaliação", value="Parecer Executivo de Maturidade Técnica - Candidato Sênior")
        tx_obj = st.text_input("Escopo Alvo das Responsabilidades", value="Migração NiFi para Python, Tunings de Queries no Oracle e PostgreSQL, e dashboards corporativos.")
        tx_ctx = st.text_area("Narrativa de Fatos e Bagagem Comprovada", value=f"O profissional demonstrou domínio avançado e senioridade no tópico '{active_data['title']}' utilizando como base técnica o caso prático: {active_data['case']}.", height=80)
        
        # Invocação limpa do componente HTML nativo com scroll funcional
        html_a4 = gerar_html_boardgame(tx_tit, tx_obj, tx_ctx, st.session_state.matriz_dinamica, st.session_state.historico_eventos)
        components.html(html_a4, height=480, scrolling=True)
        
        st.download_button("💾 Exportar Documento Completo em HTML", data=html_a4, file_name="parecer_senior_data.html", mime="text/html", use_container_width=True)

# --- ABA 2: DETALHAMENTO DE REQUISITOS (IURD) ---
with tab_vaga:
    st.subheader("📜 Mapeamento de Atividades Obrigatórias da Vaga")
    cl1, cl2 = st.columns(2)
    with cl1:
        st.markdown("""
        #### 🏗️ Responsabilidades e Atuação Diária
        * **Extração de Dados Multi-Origens:** Extrair e tratar dados de bancos relacionais (Oracle, PostgreSQL, MariaDB), NoSQL (MongoDB), APIs REST e arquivos estruturados.
        * **Migração Progressiva:** Sustentar fluxos existentes em **Apache NiFi** e apoiar a migração gradual desses pipelines legados para arquiteturas eficientes em **Python nativo**.
        * **Suporte Corporativo a Relatórios:** Criar, ajustar e manter visões de alta relevância no **Oracle Fusion Cloud (OTBI e BI Publisher)** usando fortemente queries customizadas.
        """)
    with cl2:
        st.markdown("""
        #### 🎯 Perfil Técnico Sênior Esperado
        * **SQL e PL/SQL Avançado:** Domínio absoluto de joins complexos, CTEs, funções analíticas e desenvolvimento de triggers, procedures e tuning de performance.
        * **Bibliotecas Python Aplicadas:** Prática no ecossistema de engenharia (`pandas`, `sqlalchemy`, `requests`, `openpyxl`).
        * **Governança de Dados:** Padronização de dicionários de dados, documentação de pipelines e conformidade com a LGPD.
        """)

# --- ABA 3: FRAMEWORK DE ARQUITETURA E MÉTRICAS ---
with tab_framework:
    st.subheader("📈 Métricas Estatísticas de Impacto e ROI de Engenharia")
    ct1, ct2 = st.columns(2)
    with ct1:
        st.markdown("""
        #### 🧪 Arquitetura de Validação de Dados Sênior
        * **Camada de Staging:** Scripts em Python isolam registros mal formatados ou IDs duplicados em tabelas de quarentena, evitando a poluição da base operacional de produção.
        * **Performance Tuning Dinâmico:** Análise do plano de execução do banco Oracle (Explain Plan), criando estratégias de indexação e partition pruning que evitam gargalos de CPU no servidor.
        """)
    with ct2:
        st.markdown("#### 📊 Retorno sobre Investimento Técnico (KPIs de TI)")
        st.table({
            "Métrica de Eficiência de Infraestrutura": [
                "Redução de tempo de processamento de cargas diárias manuais via Python", 
                "Melhoria na velocidade de carregamento de relatórios e views SQL após Tuning", 
                "Consumo de dados na nuvem após reengenharia de consultas (FinOps)",
                "Acurácia exigida para indicadores e relatórios executivos de alta gestão"
            ],
            "Impacto Estatístico Comprovado": ["De 1h30 para apenas 15 minutos", "Até 80% mais ágil", "Escala de Gigabytes reduzida para Megabytes", "100% (Precisão centesimal matemática)"]
        })
