import streamlit as st

st.set_page_config(
    page_title="War Room",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
.block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 0.5rem !important;
    padding-left: 1rem !important;
    padding-right: 1rem !important;
    max-width: 100% !important;
}

div[data-testid="stVerticalBlock"] {
    gap: 0.3rem !important;
}

[data-testid="stSidebarUserContent"] {
    padding-top: 1rem !important;
}

h3, p, div {
    margin-top: 0rem !important;
    margin-bottom: 0rem !important;
}

div[data-testid="stMetric"] {
    background-color: #f8f9fa;
    padding: 6px !important;
    border-radius: 4px;
    border: 1px solid #e9ecef;
    text-align: center;
}

.response-box {
    background-color: #e8f8f5;
    border-left: 4px solid #18bc9c;
    padding: 10px !important;
    border-radius: 4px;
    margin-top: 0.4rem;
    margin-bottom: 0.4rem;
}

/* Correção do seletor para garantir aplicação no botão nativo do Streamlit */
div.trigger-btn > div[data-testid="stButton"] > button {
    width: 100% !important;
    white-space: nowrap !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
    font-size: 11px !important;
    padding: 0.2rem 0.4rem !important;
    height: 28px !important;
    margin-bottom: 4px !important;
}

.category-header {
    font-size: 11px !important;
    font-weight: bold !important;
    color: #2c3e50;
    border-bottom: 1px solid #e9ecef;
    padding-bottom: 2px;
    margin-top: 0.1rem !important;
    margin-bottom: 0.4rem !important;
}
</style>
""", unsafe_allow_html=True)


# 20 Strategic Framework Database Items - Adaptado para o cenário de Analista de Dados Sênior (IURD)
DATA_MAPPING = {
    1: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Fale sobre você (Senioridade)",
        "tag": "PERFIL",
        "bridge": "Sou Engenheiro de Produção com MBA pela FGV e mais de 7 anos de experiência sólida na área de dados, especializado em transformar grandes volumes de dados brutos e dispersos em pipelines de ETL altamente performáticos e fontes de verdade unificadas.",
        "followup": "Minha trajetória une o raciocínio lógico da engenharia com a visão de governança corporativa. Sou especialista em otimização de queries complexas em SQL, automação de fluxos com Python e sustentação de ecossistemas híbridos envolvendo desde bancos tradicionais como Oracle e PostgreSQL até NoSQL como MongoDB.",
        "match": "Estabelece autoridade sênior imediata, alinhando-se exatamente aos requisitos de +7 anos e formação técnica da vaga.",
        "growth": "A contratante precisa de alguém com perfil resolutivo para assumir sistemas legados e apoiar as áreas de negócio com indicadores precisos desde o primeiro dia.",
        "case": "Engenharia + MBA + Liderança Técnica de Dados (*Heineken*, *Itaú*, *ASICS*, *NTT DATA*).",
        "bullets": [
            "Lidero a arquitetura e otimização de dados em ambientes corporativos de alta complexidade regulatória e volumetria (*Itaú*, *Heineken*).",
            "Tenho facilidade para navegar entre o desenvolvimento de código robusto (Python/SQL) e a entrega visual para a diretoria (Dashboards).",
            "Atuo ativamente com foco em governança de dados, documentação técnica estruturada e garantia de qualidade (Data Quality)."
        ],
        "qa_responses": [
            {"q": "Qual é o diferencial do seu perfil para uma posição sênior?", "a": "Eu não sou apenas um construtor de queries. Eu entendo o impacto do dado no negócio. Garanto que o dado na ponta — seja no Apache Superset ou no Power BI — seja idêntico ao payload bruto do banco, desenhando arquiteturas que não geram gargalos de infraestrutura."},
            {"q": "Como você lida com ambientes dinâmicos e pilhas tecnológicas variadas?", "a": "Com flexibilidade e mentalidade investigativa. Para mim, a tecnologia é um meio. Se o fluxo exige extrair dados de uma API REST via Python, tratar no banco Oracle via PL/SQL e disponibilizar no MongoDB, eu estruturo a esteira focando sempre em performance e manutenibilidade."},
            {"q": "Você tem experiência apoiando equipes e áreas de negócio?", "a": "Sim. Minha função sênior sempre incluiu traduzir requisitos de negócio vagos in especificações técnicas de dados, além de orientar e apoiar tecnicamente os membros mais juniores da equipe para garantir boas práticas de desenvolvimento."}
        ]
    },
    2: {
        "category": "WHY - Intent & Fit",
        "title": "Por que a IURD?",
        "tag": "ESTRATÉGIA",
        "bridge": "Uma instituição desse porte gerencia diariamente um volume massivo e heterogêneo de dados operacionais, financeiros e de sistemas institucionais, o que exige uma governança de dados impecável.",
        "followup": "Quero aplicar meu toolkit de engenharia de dados e otimização de infraestrutura exatamente onde o dado dita o ritmo da eficiência das operações internas e do atendimento às áreas finalísticas.",
        "match": "Prova que compreende a escala gigantesca e a necessidade de governança rigorosa da instituição.",
        "growth": "A operação não busca um analista júnior de relatórios; precisa de um especialista focado em sustentação, performance tuning e migração de tecnologia.",
        "case": "Cenários de Alta Volumetria e Ecossistemas Híbridos.",
        "bullets": [
            "Identifico-me com culturas que valorizam o sentimento de dono, a resiliência e a busca por melhoria contínua.",
            "O desafio de atuar com múltiplos bancos de dados (Oracle, Postgres, MariaDB, Mongo) e ferramentas de BI variadas me motiva intelectualmente.",
            "Vejo uma oportunidade clara de gerar impacto imediato na otimização de fluxos legados e na eficiência do processamento de dados."
        ],
        "qa_responses": [
            {"q": "O que mais te atrai nesta oportunidade?", "a": "A oportunidade de atuar de ponta a ponta: desde a extração via APIs ou Apache NiFi, passando pela transformação pesada em banco com PL/SQL, até a camada final de entrega no Oracle Fusion Cloud e Superset. É o cenário perfeito para um profissional sênior."},
            {"q": "Como sua experiência corporativa se conecta com os nossos desafios?", "a": "Passei por grandes corporações como Itaú e Heineken. Sei o que significa trabalhar em ambientes onde uma falha na carga ou uma query mal otimizada impacta decisões executivas. Trago essa bagagem de responsabilidade e resiliência."},
            {"q": "Você está confortável em trabalhar no modelo PJ / Híbrido?", "a": "Totalmente. Meu setup e minha rotina profissional são altamente organizados. O modelo híbrido permite o melhor dos dois mundos: foco absoluto na entrega técnica no remoto e alinhamento estratégico presencial com o time."}
        ]
    },
    3: {
        "category": "WHY - Intent & Fit",
        "title": "Visão sobre ETL e Migrações",
        "tag": "ARQUITETURA",
        "bridge": "Enxergo os processos de ETL/ELT como a espinha dorsal de qualquer tomada de decision corporativa; eles precisam ser invisíveis, rápidos e fáceis de manter.",
        "followup": "Sustentar fluxos existentes (como no Apache NiFi) exige respeito ao legado, mas migrar gradualmente para pipelines em Python é o caminho ideal para garantir escalabilidade e testes automatizados.",
        "match": "Posiciona o candidato como o executor perfeito para o plano de migração explícito na vaga.",
        "growth": "A transição de ferramentas visuais de ETL para código (Python/pandas/sqlalchemy) reduz custos de infraestrutura e aumenta o controle sobre o dado.",
        "case": "Migração e Automação de Pipelines (Case Afinz/Stalse).",
        "bullets": [
            "Tenho experiência prática na reengenharia de processos manuais ou legados para rotinas automatizadas em Python.",
            "Utilizo bibliotecas consagradas como pandas, sqlalchemy, requests (APIs) e openpyxl para construir conexões seguras e limpas.",
            "Meu foco em ETL é garantir a integridade total do dado (traceabilidade) antes que ele atinja as camadas de visualização."
        ],
        "qa_responses": [
            {"q": "Qual o seu plano para apoiar nossa migração de ETL para Python?", "a": "O primeiro passo é documentar e auditar os fluxos atuais no Apache NiFi para mapear dependências. Em seguida, traduzo as transformações lógicas para scripts Python modulares e performáticos, utilizando sqlalchemy para a escrita eficiente nos bancos de destino, garantindo zero downtime."},
            {"q": "Como você garante que uma migração de pipeline não corrompa os dados históricos?", "a": "Implementando validações em paralelo. Durante um período, o fluxo antigo e o novo rodam juntos. Utilizo scripts de reconciliação de dados para checar se os hashes e somatórios batem em 100% antes de desligar o legado."},
            {"q": "Como gerenciar múltiplos conectores (APIs, Arquivos, Bancos Relacionais)?", "a": "Centralizando e padronizando as credenciais via variáveis de ambiente seguras e criando classes utilitárias em Python para reaproveitamento de código na extração de APIs REST ou leitura de arquivos textuais e planilhas."}
        ]
    },
    4: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Sua Proposta de Valor",
        "tag": "VALOR",
        "bridge": "Ofereço o domínio técnico avançado em SQL/Python necessário para resolver gargalos de performance e a maturidade analítica para dialogar diretamente com as áreas de negócio.",
        "followup": "Elimino o abismo que costuma existir entre o que a equipe de TI desenvolve e o que os diretores e usuários de negócio realmente precisam enxergar nos relatórios.",
        "match": "Alinea-se à exigência de criar visões de dados claras utilizando ferramentas modernas (Apache Superset) e corporativas (Oracle Fusion).",
        "growth": "A capacidade investigativa de ponta a ponta poupa horas de reuniões de crise quando acontecem falhas de integração.",
        "case": "FinOps, Performance Tuning e Governança de Dados.",
        "bullets": [
            "Consolido dados dispersos para gerar relatórios corporativos com precisão matemática centesimal.",
            "Otimizo consultas SQL lentas através de análise de planos de execução (Performance Tuning), reduzindo consumo de memória e CPU.",
            "Entrego documentação técnica clara de dados e processos, mitigando o risco de perda de conhecimento."
        ],
        "qa_responses": [
            {"q": "O que você consegue entregar nos primeiros 30 dias se for contratado?", "a": "Focar na imersão dos ambientes Oracle e PostgreSQL, assumir a sustentação das cargas diárias do Apache NiFi e mapear as principais dores dos usuários nos relatórios do OTBI / BI Publisher, garantindo estabilidade imediata na operação."},
            {"q": "Qual é a sua abordagem em relação à padronização e governança de dados?", "a": "O dado só tem valor se for confiável e documentado. Sigo padrões rígidos de nomenclatura, mapeamento de metadados e linhagem do dado, alinhado com as diretrizes da LGPD para garantir que informações sensíveis sejam tratadas com as permissões corretas."},
            {"q": "Como você lida com a pressão de relatórios executivos urgentes?", "a": "Com organização e foco em solução. Entendo que relatórios corporativos para tomada de decisão não podem esperar. Minha senioridade me permite focar na extração rápida da métrica correta, sem perder a qualidade técnica na query."}
        ]
    },
    5: {
        "category": "WHY - Intent & Fit",
        "title": "Expectativa Salarial / Contrato",
        "tag": "ANCORAGEM",
        "bridge": "Minha pretensão salarial está fundamentada na minha senioridade de mais de 7 anos e no valor imediato que posso gerar na otimização da infraestrutura de dados e relatórios da instituição.",
        "followup": "Busco uma remuneração justa para um Especialista que atua no modelo Prestador de Serviços (PJ) com total autonomia e prontidão técnica.",
        "match": "Demonstra clareza profissional, maturidade de negociação e posicionamento como especialista de mercado.",
        "growth": "O custo de um especialista sênior se paga rapidamente através do ganho de eficiência em processos de dados e no tuning de infraestrutura que reduz o desperdício de nuvem/servidores.",
        "case": "Contratação PJ Especialista - Período Integral.",
        "bullets": [
            "Minha pretensão salarial para o formato de prestação de serviços (PJ) está na faixa de 12.000 a 15.000 Reais mensais, dependendo da complexidade exata e benefícios.",
            "Estou totalmente pronto para iniciar o modelo híbrido em São Paulo - SP com total disponibilidade de horários.",
            "Possuo empresa aberta (CNPJ) regularizada com emissão de nota fiscal imediata para processos de faturamento sem burocracia."
        ],
        "qa_responses": [
            {"q": "Este valor é negociável?", "a": "Estou aberto a entender o pacote completo oferecido pela instituição, os desafios de longo prazo e as possibilidades de evolução dentro do ecossistema técnico. Havendo sinergia, o valor pode ser ajustado de forma justa para ambos."},
            {"q": "Por que deveríamos investir esse orçamento no seu perfil?", "a": "Porque trago autonomia completa. Não preciso de treinamento para escrever PL/SQL avançado, corrigir fluxos em NiFi ou criar dashboards no Superset. Minha contratação minimiza o tempo de onboarding e elimina os erros comuns cometidos por profissionais menos experientes."},
            {"q": "Você tem flexibilidade para atuar com diferentes tecnologias de bancos?", "a": "Sim, essa flexibilidade é uma das minhas forças. Minha experiência me permite transitar de um tuning de query em Oracle para uma agregação complexa em MongoDB (NoSQL) sem atritos, garantindo soluções homogêneas."}
        ]
    },
    6: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Vivência com Oracle Fusion Cloud",
        "tag": "ORACLE-FUSION",
        "bridge": "Considero o ecossistema Oracle Fusion Cloud, especificamente OTBI e BI Publisher, uma camada estratégica vital, onde o SQL robusto é a chave para destravar relatórios corporativos complexos.",
        "followup": "Muitos analistas focam apenas em ferramentas de visualização modernas, mas eu domino a extração nativa dentro do ambiente corporativo Oracle, manipulando data models complexos e views customizadas.",
        "match": "Atende diretamente o requisito desejável da vaga, transformando-o em um diferencial competitivo matador.",
        "growth": "A contratante utiliza o Oracle Fusion para sua gestão corporativa; ter um profissional focado em ajustar e manter esses relatórios sem dependência de consultorias externas é um enorme ganho de agilidade.",
        "case": "Criação, Ajuste e Manutenção de Relatórios OTBI / BI Publisher.",
        "bullets": [
            "Compreendo a estrutura de dados interna dos módulos Oracle ERP/Fusion para localizar tabelas e campos rapidamente.",
            "Utilizo SQL avançado para customizar os Data Models que alimentam os layouts do BI Publisher.",
            "Sei apoiar as áreas de negócio na criação de análises em tempo real utilizando as áreas de assunto (Subject Areas) do OTBI."
        ],
        "qa_responses": [
            {"q": "Como você avalia sua experiência com OTBI and BI Publisher?", "a": "Minha experiência foca em resolver o que as interfaces padrões não entregam. Utilizo o OTBI para relatórios analíticos rápidos e arrastáveis para o usuário, e recorro ao BI Publisher quando a diretoria exige relatórios altamente formatados, com layouts complexos, alimentados por queries SQL customizadas."},
            {"q": "O que você faz quando um relatório do BI Publisher apresenta lentidão?", "a": "O problema quase sempre está no Data Model subjacente. Extraio a query SQL contida nele, analiso o plano de execução dentro do ambiente Oracle, verifico se há joins desnecessários ou falta de índices adequados e faço o refactoring da consulta."},
            {"q": "Você sabe integrar os dados do Oracle Fusion com outras bases (como Postgres)?", "a": "Sim. Podemos construir rotinas em Python que consomem as APIs REST nativas do Oracle Fusion ou extraem relatórios agendados do BI Publisher em formatos estruturados (CSV/XML) para alimentar um DW centralizado ou tabelas no PostgreSQL/MongoDB."}
        ]
    },
    7: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Domínio Avançado de SQL / PL-SQL",
        "tag": "SQL-TUNING",
        "bridge": "Para mim, SQL não é apenas escrever SELECTs básicos; é dominar joins complexos, CTEs, subqueries correlacionadas e funções analíticas para processar dados com máxima performance.",
        "followup": "Em ambientes com grande volume de dados, o design da query dita o custo e a velocidade da entrega. Escrevo códigos limpos, estruturados e fáceis de auditar.",
        "match": "Preenche perfeitamente o principal pilar técnico da vaga (Domínio avançado em SQL e PL/SQL).",
        "growth": "Garante que os sistemas gerenciados no Oracle, PostgreSQL e MariaDB rodem de forma otimizada, evitando travamentos em bancos de produção.",
        "case": "Performance Tuning e Otimização de Consultas Complexas.",
        "bullets": [
            "Utilizo CTEs (Common Table Expressions) extensivamente para modularizar queries longas e torná-las legíveis.",
            "Domino funções analíticas (ROW_NUMBER, RANK, LEAD, LAG, SUM OVER) para evitar subqueries pesadas e desnecessárias.",
            "Possuo conhecimento em PL/SQL para criar triggers, procedures e funções que automatizam lógicas de tratamento direto no banco de dados Oracle."
        ],
        "qa_responses": [
            {"q": "Como você abordagem o refactoring de uma query SQL de 500 linhas que está travando o banco?", "a": "Com método. Primeiro isolo as partes utilizando CTEs para entender o fluxo de dados. Analiso o Explain Plan para identificar Table Scans desnecessários. Verifico se os Joins estão usando chaves indexadas e se há filtros que podem ser antecipados para diminuir a volumetria processada logo no início da consulta."},
            {"q": "Qual a sua experiência com bancos NoSQL como MongoDB?", "a": "Utilizo o MongoDB para cenários onde a estrutura do dado é altamente fluida ou semiestruturada (JSONs). Sei trabalhar com o framework de agregação (Aggregation Pipeline) do Mongo e utilizo Python para ler esses documentos e normalizá-los caso precisem ser inseridos em bancos relacionais."},
            {"q": "O que você prioriza na criação de uma View?", "a": "Priorizo a consistência lógica e o impacto na performance. Se a view for consultada com muita frequência sobre milhões de linhas, avalio junto ao DBA a criação de visões materializadas ou a persistência dos dados tratados em tabelas intermediárias via processos de ETL noturnos."}
        ]
    },
    8: {
        "category": "WHY - Intent & Fit",
        "title": "Você é superqualificado?",
        "tag": "RETENÇÃO",
        "bridge": "Acredito que o termo correto não é superqualificado, mas sim plenamente preparado para os desafios de alta complexidade técnica que a instituição possui.",
        "followup": "Um profissional sênior não busca apenas tarefas complexas; busca estabilidade, processos organizados e a oportunidade de construir pipelines eficientes que resolvam problemas reais.",
        "match": "Elimina o receio do recrutador de que o profissional ache o trabalho monótono ou saia rapidamente.",
        "growth": "A escala de dados da contratante exige alguém experiente para evitar retrabalho estrutural e débitos técnicos nas rotinas de dados.",
        "case": "Maturidade Profissional e Foco em Soluções Duradouras.",
        "bullets": [
            "Tenho real motivação em atuar na sustentação técnica e melhoria contínua de ecossistemas maduros.",
            "Para mim, o desafio intelectual está em otimizar rotinas que hoje demoram horas para rodar em poucos minutos.",
            "Busco um vínculo de longo prazo onde minha senioridade técnica possa apoiar o crescimento da equipe de dados."
        ],
        "qa_responses": [
            {"q": "Sendo sênior, você aceitaria realizar tarefas mais operacionais ou suporte a relatórios?", "a": "Com certeza. Minha senioridade me dá a maturidade de entender que o suporte a relatórios corporativos existentes é fundamental para a governança da instituição. Se a área de negócio precisa de um ajuste rápido em um relatório do Power BI ou OTBI, encaro isso com o mesmo senso de dono de quando estou projetando um pipeline do zero."},
            {"q": "O que te mantém motivado e engajado em um projeto?", "a": "A autonomia para propor melhorias de performance e a oportunidade de sanar inconsistências de dados ponta a ponta. Ver um fluxo de ETL migrado para Python rodando de forma lisa e sem erros é o que me traz satisfação profissional."},
            {"q": "Como sua senioridade ajuda a economizar recursos da empresa?", "a": "Evitando o retrabalho. Um desenvolvedor menos experiente pode criar soluções paliativas em planilhas que quebram no mês seguinte. Eu desenho pipelines estruturados, documentados em ferramentas corporativas, que suportam o crescimento do volume de transações por anos sem necessidade de reengenharia."}
        ]
    },
    9: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Explicação sobre Ciclos Curtos",
        "tag": "PROJETOS",
        "bridge": "Minhas passagens recentes por empresas como Stalse (atendendo ASICS) e NTT DATA (atendendo Itaú) foram alocações estratégicas focadas em projetos com escopo fechado e entregas de tiro curto.",
        "followup": "Atuei como um acelerador de soluções, resolvendo gargalos de arquitetura específicos e realizando migrações críticas que demandavam força técnica sênior imediata.",
        "match": "Transforma uma possível objeção de instabilidade em uma prova de agilidade, adaptabilidade e entrega rápida de resultados.",
        "growth": "Demonstra que o candidato tem facilidade de onboarding instantâneo e consegue gerar valor desde a primeira semana.",
        "case": "Sprints Ágeis e Consultoria Técnica de Dados.",
        "bullets": [
            "Na *Stalse*, unifiquei dados financeiros internacionais da *ASICS* e apliquei conceitos de FinOps para reduzir o consumo de nuvem.",
            "Na *NTT DATA*, trabalhei imerso no ambiente AWS Cloud do *Itaú*, lidando com volumetria massiva através de queries complexas em Amazon Athena.",
            "Agora, meu objetivo estratégico é fixar minhas habilidades em uma posição estável de longo prazo."
        ],
        "qa_responses": [
            {"q": "Por que seus últimos contratos duraram cerca de 4 meses?", "a": "Ambos foram contratos de consultoria por escopo. No Itaú (NTT DATA), o objetivo era estruturar transformações de views complexas sob metodologia ágil em um período determinado. Na ASICS (Stalse), a meta era otimizar o consumo de dados de Gigabytes para Megabytes. Finalizadas as entregas com sucesso e a documentação concluída, o ciclo do projeto encerrou naturalmente."},
            {"q": "Como você se adapta tão rápido a novas culturas de equipe?", "a": "Minha bagagem de mais de 2.000 análises estruturadas desenvolvidas ao longo da carreira me deu uma adaptabilidade técnica e comportamental fora da curva. Eu entro no projeto focado em ouvir as dores do time, entender o mapeamento das tabelas e começar a codificar as soluções sem gerar ruído."},
            {"q": "Você está buscando estabilidade no momento?", "a": "Sim, exatamente. Escolhi participar deste processo seletivo porque vejo na IURD um ambiente robusto, com um fluxo de dados contínuo e volumoso, ideal para eu me consolidar como um pilar técnico de dados por muitos anos."}
        ]
    },
    10: {
        "category": "WHY - Intent & Fit",
        "title": "Evolução Profissional e Engenharia",
        "tag": "EVOLUÇÃO",
        "bridge": "Minha transição da análise de gestão de negócios tradicional para a engenharia e análise avançada de dados reflete o movimento natural do mercado corporativo moderno.",
        "followup": "Minha formação em Engenharia de Produção me deu a mentalidade de processos e eficiência, enquanto o MBA na FGV consolidou a visão executiva de indicadores. Unir isso ao desenvolvimento em Python e SQL foi o passo lógico para me tornar um profissional completo.",
        "match": "Posiciona o candidato como um perfil híbrido raro: possui a casca técnica de codificação e a maturidade de negócios para conversar com diretores.",
        "growth": "A vaga pede alguém que ajude as áreas de negócio na definição e evolução de indicadores. Técnicos puros costumam falhar nessa comunicação.",
        "case": "Combinação de Gestão Estruturada com Hard Skills de TI.",
        "bullets": [
            "Minha formação pelo CREA-SP garante meu foco em métodos ágeis, eliminação de desperdícios e otimização de fluxos.",
            "Desenvolvi a capacidade investigativa ponta a ponta através de anos auditando processos corporativos e registros complexos.",
            "Enxergo os dashboards no Apache Superset ou Power BI como ferramentas de gestão, e não apenas visuais bonitos."
        ],
        "qa_responses": [
            {"q": "Como sua formação em Engenharia de Produção agrega valor no dia a dia de dados?", "a": "A engenharia foca em gargalos. Quando olho para um pipeline de ETL que demora para rodar, ou uma integração de API que cai constantemente, não vejo só um erro de código; vejo um desperdício de processo que atrasa a tomada de decisão. Aplico conceitos de fluxo contínuo e qualidade total na esteira de dados."},
            {"q": "Por que focar em ferramentas como Apache Superset?", "a": "O Apache Superset é uma ferramenta fantástica, open-source, extremamente performática para grandes volumes e que consome dados direto via SQL de forma muito limpa. Tenho facilidade em plugá-lo sobre bancos PostgreSQL ou de Data Warehouse para democratizar o acesso aos dados sem custos abusivos de licenciamento comercial."},
            {"q": "Qual a sua visão sobre a LGPD no tratamento de dados?", "a": "A conformidade com a LGPD é inegociável. Um Analista Sênior deve garantir que dados sensíveis de usuários ou clientes sejam devidamente mascarados ou criptografados nas camadas de banco e que o acesso aos dashboards do Superset/Power BI respeite estritamente os perfis e permissões de cada cargo."}
        ]
    },
    11: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "Case ASICS (Otimização e FinOps)",
        "tag": "CASE-ETL",
        "bridge": "Na Stalse, liderei a reengenharia de uma arquitetura de dados financeiros e operacionais para a operação Latam da ASICS (Brasil, Chile e Colômbia), focando em performance de queries e redução drástica de custos.",
        "followup": "O cenário era de fragmentação de dados e consultas extremamente pesadas que geravam um consumo financeiro excessivo nas plataformas de nuvem (BigQuery).",
        "match": "Prova capacidade real de lidar com volumetria e aplicar otimização de queries (performance tuning), um dos diferenciais pedidos pela vaga.",
        "growth": "Demonstra proatividade e senso de dono ao tratar os recursos computacionais da contratante como se fossem seus.",
        "case": "Redução de Custos e Unificação de Dados Financeiros.",
        "bullets": [
            "Situação: Queries ineficientes consumiam Gigabytes de processamento desnecessário, gerando lentidão nos painéis de indicadores e custos elevados.",
            "Action: Refatorei as consultas SQL aplicando boas práticas de particionamento, CTEs e eliminando joins redundantes.",
            "Result: Slashed o consumo de dados de escala de Gigabytes para Megabytes, otimizando o tempo de resposta das visões gerenciais."
        ],
        "qa_responses": [
            {"q": "Como você aplicou a otimização de dados na prática para a ASICS?", "a": "Identifiquei que muitas subqueries eram executadas repetidamente dentro de loops lógicos. Substituí essa estrutura por tabelas temporárias bem indexadas e queries utilizando funções analíticas. Isso reduziu o volume de dados escaneados e acelerou a atualização dos relatórios."},
            {"q": "Esse projeto envolveu dados de múltiplos países. Como lidou com as diferenças de regras?", "a": "Criamos uma camada de staging em Python que padronizava os tipos de dados e os formatos de entrada antes de consolidá-los na base relacional principal. Isso garantiu que as regras de negócio específicas de cada região fossem aplicadas uniformemente."},
            {"q": "Como você documentou essa nova estrutura?", "a": "Criei um dicionário de dados detalhado e um diagrama de linhagem (data lineage). Assim, qualquer novo analista que entrasse no time conseguiria entender exatamente de qual tabela original saía cada métrica do dashboard."}
        ]
    },
    12: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "Case Itaú / NTT DATA (Volumetria AWS)",
        "tag": "CASE-SCALE",
        "bridge": "Alocado na NTT DATA para atender o Itaú, atuei imerso em ambiente de computação em nuvem (AWS), construindo visões SQL complexas e robustas sobre tabelas com bilhões de registros.",
        "followup": "O projeto exigia precisão absoluta e velocidade de processamento para atender métricas de desempenho corporativo consultadas por mais de 5.000 executivos e gestores.",
        "match": "Valida a experiência em ambientes com alto volume de dados, queries complexas e arquitetura moderna.",
        "growth": "A IURD exige experiência sólida em SQL e capacidade de rodar transformações eficientes. Esse case elimina qualquer dúvida sobre a capacidade técnica do candidato.",
        "case": "Construção de Views Complexas em Amazon Athena e AWS S3.",
        "bullets": [
            "Situação: Havia a necessidade de calcular métricas gerenciais complexas cruzando tabelas gigantescas de logs e dados cadastrais que mudavam constantemente.",
            "Action: Desenvolvi visões SQL estruturadas utilizando Amazon Athena, aplicando estratégias de filtros eficientes e parametrização dinâmica.",
            "Result: Entreguei painéis automatizados estáveis, garantindo consistência total do dado e tempo de carregamento otimizado."
        ],
        "qa_responses": [
            {"q": "Como manter as queries rápidas trabalhando com bilhões de linhas no Itaú?", "a": "A chave foi trabalhar em sintonia com a estrutura de particionamento dos dados no S3. Garantindo que as cláusulas WHERE da query SQL fizessem o 'partition pruning' correto, o Athena lia apenas a fração necessária de arquivos, reduzindo o tempo de execução de minutos para segundos."},
            {"q": "Como você lidava com as mudanças constantes nas regras de negócio desse projeto?", "a": "Evitava hardcodar regras diretamente nas queries principais. Criava tabelas de parametrização de indicadores. Quando uma regra de negócio mudava, apenas atualizávamos o registro de configuração no banco e a view SQL interpretava a nova regra dinamicamente, sem necessidade de deploy de código."},
            {"q": "Qual foi a importância da metodologia ágil nesse ambiente?", "a": "Essencial. Trabalhávamos em sprints curtas com entregas semanais. Isso garantia alinhamento total com as unidades de negócio do banco e permitia corrigir desvios de escopo de forma imediata."}
        ]
    },
    14: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "Case Afinz (Automação de ETL 1h30 para 15m)",
        "tag": "CASE-AUTOMATION",
        "bridge": "Como Analista de MIS na Afinz/Sorocred, identifiquei um fluxo crítico de relatórios diários que era executado de forma manual, gerando atrasos crônicos na disponibilização de dados operacionais.",
        "followup": "A rotina dependia de operadores compilando dados manualmente em Excel e disparando queries fragmentadas, o que aumentava drasticamente o risco de erros humanos.",
        "match": "Atende perfeitamente o requisito da vaga de 'desenvolver, manter e otimizar processos de transformação de dados e rotinas em Python'.",
        "growth": "Prova capacidade de automação de processos, liberando tempo técnico da equipe para análises mais profundas e estratégicas.",
        "case": "Redução do Tempo de Processamento de ETL de 1h30 para 15 minutos.",
        "bullets": [
            "Situação: Rotinas manuais consumiam 1 hora e meia diária da equipe técnica e atrasavam o início das análises de negócio gerenciais.",
            "Action: Desenvolvi pipelines de ETL automatizados de ponta a ponta utilizando Python, integrando o agendamento de scripts com consultas SQL diretas no banco.",
            "Result: Reduzi o tempo total de execução para apenas 15 minutos, fortalecendo a governança e eliminando as inconsistências das cargas."
        ],
        "qa_responses": [
            {"q": "Qual foi o principal desafio na automação do processo na Afinz?", "a": "O principal desafio foi mapear todas as exceções e 'gambiarras' que os operadores faziam manualmente nas planilhas e traduzir isso em regras lógicas condicionais dentro do código Python usando a biblioteca pandas. Uma vez mapeado, o script passou a tratar os dados de forma idêntica e sem falhas."},
            {"q": "Como você estruturou a governança de dados nesse projeto?", "a": "Aproveitei a automação para criar uma tabela de log de execução no banco via AWS Glue/S3. Cada vez que o script rodava, ele registrava a hora de início, fim, quantidade de linhas inseridas e se houve algum aviso de erro. Isso deu total transparência para auditorias internas."},
            {"q": "Como essa mentalidade se aplica à migração de fluxos do Apache NiFi que a nossa vaga pede?", "a": "A lógica é análoga. No NiFi, você tem os processadores visuais. Para migrar para Python de forma performática, eu analiso o que cada processador faz (um SplitText, um EvaluateJsonPath, um PutSQL), reescrevo isso de forma nativa e otimizada em um script Python modular e garanto que o consumo de memória do servidor seja o menor possível."}
        ]
    },
    13: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "Case Heineken (Normalização de Dados)",
        "tag": "CASE-BI",
        "bridge": "Na Heineken, atuando na divisão digital de e-commerce e eRetail, fui responsável por consolidar e normalizar grandes conjuntos de dados provenientes de múltiplos clientes e parceiros externos.",
        "followup": "Os dados chegavam em formatos totalmente heterogêneos (APIs variadas, arquivos de texto, planilhas bagunçadas), o que impedia um acompanhamento unificado das campanhas comerciais.",
        "match": "Demonstra domínio prático na extração de dados de diversas fontes e modelagem estruturada (DW / Star Schema).",
        "growth": "A IURD busca alguém capaz de unificar múltiplas origens de dados. Esse case comprova essa aptidão de alto nível corporativo.",
        "case": "Modelagem Star Schema e Integration de Múltiplas Fontes de Clientes.",
        "bullets": [
            "Situação: A falta de padronização nas fontes externas gerava falhas de integração frequentes e dados duplicados nos relatórios gerenciais.",
            "Action: Desenhei um modelo relacional robusto no padrão Star Schema (Fatos e Dimensões) e apliquei rotinas estritas de validação de qualidade de dados.",
            "Result: Montei painéis corporativos estáveis no Power BI que batiam com os indicadores financeiros até a última casa decimal."
        ],
        "qa_responses": [
            {"q": "Como você lidou com dados bagunçados de terceiros na Heineken?", "a": "Criei uma camada rígida de higienização de dados antes de inseri-los nas tabelas de dimensão. Se um parceiro enviava um nome de produto incorreto ou um ID fora do padrão, o script Python isolava essa linha em uma tabela de quarentena e notificava o time, impedindo a poluição do Data Warehouse principal."},
            {"q": "Qual a vantagem de usar a modelagem Star Schema nesse cenário?", "a": "O Star Schema simplifica os Joins na hora de construir os dashboards. Ao separar os dados de transações (Fato) dos dados cadastrais (Dimensões), as queries SQL rodam infinitamente mais rápido, seja no Power BI ou no Apache Superset, melhorando a experiência do usuário final."},
            {"q": "Você tem experiência com ferramentas open-source de BI como Apache Superset?", "a": "Sim. O Superset brilha justamente quando o backend está bem modelado em Star Schema. Como ele faz consultas diretas e eficientes no banco SQL, ter tabelas de dimensões limpas permite que os gestores montem seus próprios gráficos via drag-and-drop sem pesar a infraestrutura."}
        ]
    },
    15: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "Case Burity (Capacidade Investigativa Legal)",
        "tag": "CASE-AUDIT",
        "bridge": "Durante minha longa e sólida trajetória na Burity Empresarial, atuei na gestão de ativos e conformidade regulatória como Procurador Legal, desenvolvendo uma capacidade investigativa ponta a ponta e um olhar cirúrgico para mitigação de riscos operacionais.",
        "followup": "Fui responsável por auditar processos administrativos complexos, contratos corporativos de alto valor e plantas técnicas de engenharia, retificando erros históricos e eliminando passivos com órgãos governamentais.",
        "match": "Evidencia as competências comportamentais críticas descritas no perfil: capacidade investigativa ponta a ponta, atenção a detalhes, ética e resiliência.",
        "growth": "A governança de dados e o respeito a normas internas e à LGPD exigem a mentalidade de auditor rígido que esse histórico comprova.",
        "case": "Retificação de Registros, Auditoria de Processos e Risco Zero.",
        "bullets": [
            "Situação: Divergências descritivas em cadastros históricos geravam riscos de multas severas e travavam a expansão de ativos de grande escala.",
            "Action: Liderei auditorias documentais profundas, cruzei dados técnicos e coordenei times multidisciplinares para sanar erros de registro.",
            "Result: Garanti a conformidade jurídica e patrimonial dos ativos de forma 100% administrativa, com zero litígios ou penalidades."
        ],
        "qa_responses": [
            {"q": "Como essa experiência jurídica e de gestão se aplica a uma vaga estritamente de dados?", "a": "A essência da auditoria é rigorosamente a mesma. Investigar uma inconsistência em um registro imobiliário na Burity exige o mesmo nível de atenção, ceticismo e busca por evidências que investigar uma falha de carga em uma tabela Oracle ou uma inconsistência de dados entre o OTBI e a base PostgreSQL. Eu não desisto até encontrar a causa raiz do problema."},
            {"q": "Você se considera uma pessoa resiliente para lidar com sistemas legados complicados?", "a": "Sem dúvidas. Minha trajetória gerindo crises patrimoniais e lidando com burocracias pesadas me ensinou a manter a calma sob extrema pressão. Sei navegar por sistemas legados confusos com paciência, extraindo a lógica de negócio necessária para documentar e organizar os fluxos de dados de forma transparente."},
            {"q": "Como era sua comunicação com os diferentes níveis hierárquicos nesse período?", "a": "Sempre foi clara e precisa. Eu precisava traduzir termos técnicos de engenharia civil e cláusulas jurídicas complexas para diretores corporativos tomarem decisões financeiras. Na área de dados faço o mesmo: traduzo queries complexas e arquiteturas NoSQL em impactos de custo e eficiência de negócios."}
        ]
    },
    16: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "Tratamento de Inconsistências Críticas",
        "tag": "CRISE-DADO",
        "bridge": "Se um fluxo de integração quebrar na madrugada ou os dados do dashboard amanhecerem duplicados, minha postura sênior imediata é conter o impacto e isolar a falha com total transparência.",
        "followup": "Não busco culpados; busco logs. Identifico o range de dados afetado, executo um script de reversão limpo e aplico a correção definitiva no pipeline para que o erro nunca mais se repita.",
        "match": "Alinha-se com a responsabilidade de 'investigar inconsistências de dados, falhas de integração e problemas de carga'.",
        "growth": "A maturidade em momentos de crise evita decisões precipitadas que possam corromper dados de produção históricos.",
        "case": "Gestão de Incidentes em Ambientes de Produção.",
        "bullets": [
            "Mantenho a calma, desligo gatilhos automáticos problemáticos e analiso os logs de erro do Python ou do Apache NiFi.",
            "Desenvolvo scripts rápidos de remediação para limpar duplicidades respeitando chaves primárias e constraints do banco.",
            "Registro um post-mortem técnico detalhando o ocorrido e a solução aplicada para alimentar a base de conhecimento da TI."
        ],
        "qa_responses": [
            {"q": "O que você faz quando a área de negócios aponta que o relatório financeiro do início do mês está incorreto?", "a": "Primeiro, valido a reclamação cruzando o dado visualizado com a base transacional bruta através de uma query de validação SQL. Se a inconsistência for real, investigo a esteira de ETL para ver se alguma carga falhou ou se uma regra de negócio mudou na origem e não foi atualizada no nosso pipeline. Corrijo, processo a carga retroativa e aviso os stakeholders com clareza."},
            {"q": "Como garantir que um script de correção em Python não piore a situação do banco Oracle?", "a": "Nunca rodo scripts de correção direto em produção sem antes testar rigorosamente em ambiente de homologação (Staging). Utilizo transações controladas no banco de dados (BEGIN TRANSACTION / COMMIT / ROLLBACK) para garantir que, se algo falhar no meio do processo de correção, o banco volte ao estado anterior com segurança."},
            {"q": "Você tem autonomia para debugar problemas em servidores Linux?", "a": "Sim. Tenho excelente familiaridade com comandos de terminal Linux para navegar em servidores, checar uso de memória de processos Python, olhar logs de Docker e verificar agendamentos de tarefas (cronjobs) que possam estar travando queries ou integrações."}
        ]
    },
    17: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "Demandas Não Mapeadas sob Pressão",
        "tag": "FLEXIBILIDADE",
        "bridge": "Diante de solicitações urgentes de indicadores feitas pela diretoria em cenários caóticos, utilizo o pensamento estruturado para focar no MVP (Mínimo Produto Viável) do dado.",
        "followup": "Isolo o ruído emocional, extraio uma amostra confiável diretamente via SQL dos bancos relacionais e apresento um panorama claro do risco ou oportunidade com os dados disponíveis no momento.",
        "match": "Comprova as competências comportamentais de proatividade, senso de dono e flexibilidade para lidar com imprevistos.",
        "growth": "Instituições dinâmicas enfrentam mudanças de cenário rápidas; um sênior precisa ser o porto seguro analítico nesses momentos.",
        "case": "Extrações Rápidas de Emergência para Apoio de Decisão.",
        "bullets": [
            "Alinho com a liderança as prioridades para entender o núcleo real da necessidade do negócio.",
            "Escrevo consultas SQL otimizadas usando indexação adequada para não derrubar a performance do banco operacional.",
            "Entrego o resultado de forma limpa, apontando as premissas adotadas e eventuais limitações técnicas dos dados extraídos."
        ],
        "qa_responses": [
            {"q": "Como você lida quando duas áreas de negócio pedem relatórios urgentes e conflitantes ao mesmo tempo?", "a": "Uso minha comunicação clara e diplomática. Converso com os gestores das duas áreas para entender o impacto financeiro ou operacional de cada demanda. Caso não haja um consenso óbvio de prioridade, escalo o cenário para a liderança de TI apresentar o cronograma técnico, garantindo que o senso de dono guie a melhor escolha para a instituição."},
            {"q": "Você se sente confortável trabalhando em ambientes que mudam de prioridade rapidamente?", "a": "Sim. Minha experiência em engenharia me ensinou a construir estruturas modulares. Se eu crio um pipeline em Python ou um relatório no BI Publisher bem estruturado e documentado, uma mudança de prioridade não joga meu trabalho no lixo; eu apenas reconfiguro os módulos para atender o novo cenário de forma ágil."},
            {"q": "O que você faz se precisar utilizar uma tecnologia que nunca viu na vida?", "a": "Abordo com curiosidade técnica e proatividade investigativa. Como sênior, compreendo os fundamentos de lógica, bancos de dados e engenharia. Aprender a sintaxe de uma ferramenta nova — seja uma ferramenta de BI ou um orquestrador diferente — é apenas questão de ler a documentação oficial e aplicar boas práticas desde o primeiro dia."}
        ]
    },
    18: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "Alinhamento com Áreas de Negócio",
        "tag": "COMUNICAÇÃO",
        "bridge": "Não discuto dados com base em achismos ou opiniões subjetivas; trago fatos, volumetria e métricas de qualidade para a mesa para alinhar equipes multidisciplinares.",
        "followup": "Geralmente, as áreas de negócio criam atritos com a TI porque sentem falta de agilidade ou não entendem as restrições técnicas. Atuo como o tradutor ideal entre esses dois mundos.",
        "match": "Cumpre a responsabilidade de 'apoiar áreas de negócio na definição e evolução de indicadores e análises' de forma colaborativa e assertiva.",
        "growth": "Evita o isolamento da equipe de dados, integrando-a estrategicamente à rotina de crescimento da contratante.",
        "case": "Construção de Pontes Técnicas entre TI e Usuários de Negócios.",
        "bullets": [
            "Escuto ativamente as necessidades dos usuários para entender quais dores de negócio eles tentam sanar com os relatórios.",
            "Apresento protótipos rápidos de dashboards no Apache Superset para validar o layout e os indicadores antes de fechar o código backend.",
            "Explico limitações de infraestrutura de forma simples, mostrando como uma query otimizada protege a integridade e a velocidade do dado deles."
        ],
        "qa_responses": [
            {"q": "Como você reage quando um gestor de negócios exige um indicador que você sabe que a estrutura de dados atual não consegue calcular?", "a": "Não digo apenas 'não é possível'. Apresento o mapeamento atual dos dados nos nossos bancos Oracle/Postgres e explico de forma didática o gap existente (por exemplo, a falta de uma flag na API de origem). Proponho uma solução alternativa provisória e me coloco à disposição para apoiar a evolução do sistema de origem para passarmos a capturar esse dado no futuro."},
            {"q": "Como você conduz reuniões de definição de indicadores corporativos?", "a": "Foco na padronização. Garanto que todos na sala concordem com a mesma regra de cálculo. Se a área A calcula faturamento de um jeito e a área B de outro, meu papel como sênior de dados é provocar essa unificação de conceitos para que o relatório corporativo exiba uma única versão da verdade (Single Source of Truth)."},
            {"q": "Você tem experiência em treinar ou apoiar usuários na utilização de ferramentas de self-service BI?", "a": "Sim. Gosto de capacitar os usuários-chave (Key Users) nas ferramentas como OTBI ou Power BI. Criando visões e views bem limpas no banco, dou autonomia para eles criarem seus relatórios básicos, o que desafoga a equipe de TI sênior para focar em pipelines estruturais e performance tuning."}
        ]
    },
    19: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "Tradução Técnica para Executivos",
        "tag": "EXECUTIVO",
        "bridge": "Ao me reportar à diretoria ou a stakeholders não técnicos, elimino jargões pesados de programação e foco puramente em eficiência operacional, acurácia financeira e mitigação de riscos.",
        "followup": "Diretores não precisam saber as linhas de comando do pandas ou os nós do Apache NiFi; eles precisam ter certeza de que o indicador na tela está correto e que o processamento está seguro.",
        "match": "Garante uma comunicação executiva clara, direta e orientada a resultados, condizente com um profissional sênior.",
        "growth": "Fornece relatórios gerenciais e dashboards que transmitem segurança imediata para decisões de alto escalão da instituição.",
        "case": "Apresentação de Indicadores e Relatórios Gerenciais de Alta Relevância.",
        "bullets": [
            "Substituo explicações sobre queries SQL complexas por métricas de horas de trabalho economizadas com automações.",
            "Uso recursos visuais claros e objetivos no Apache Superset/Power BI que permitem o entendimento do cenário de negócios em um vislumbre de 5 segundos.",
            "Foco em demonstrar a consistência e integridade dos dados (Governança), provando que o número apresentado é auditável."
        ],
        "qa_responses": [
            {"q": "Como você apresenta um relatório técnico de volumetria de dados para a alta liderança?", "a": "Não falo de bytes ou partições. Traduzo isso em capacidade de atendimento e velocidade. Apresento gráficos de tendência de crescimento de registros nos bancos MariaDB/Oracle, mostrando como a otimização que fiz nos processos de ETL evitou o travamento de relatórios e garantiu que a diretoria recebesse as informações matinais impreterivelmente no horário correto."},
            {"q": "Qual a importância da documentação técnica para o nível executivo?", "a": "A documentação técnica garante a continuidade do negócio. Mostro para a gerência que todos os nossos fluxos de ETL em Python, relatórios do BI Publisher e visões do PostgreSQL estão totalmente catalogados. Se amanhã a equipe mudar, a instituição não perde o controle sobre suas regras de dados. Isso é governança corporativa de verdade."},
            {"q": "Como você lida com feedbacks negativos de diretores sobre a usabilidade de um painel?", "a": "Com total flexibilidade e foco na solução. Entendo que o dashboard precisa ser útil para quem toma a decisão. Se um executivo acha um painel do Superset confuso, reavalio a disposição dos filtros, simplifico os gráficos e busco uma abordagem mais limpa e minimalista, mantendo a robustez técnica nos bastidores da query SQL."}
        ]
    },
    20: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "Abordagem Investigativa e Fechamento",
        "tag": "FECHAMENTO",
        "bridge": "Para encerrar, gostaria de destacar que meu perfil técnico sênior e minha bagagem de negócios estão perfeitamente calibrados para assumir a sustentação e evolução do ecossistema de dados da IURD.",
        "followup": "Estou totalmente motivado para assumir o desafio de otimizar seus relatórios em Oracle Fusion Cloud (OTBI/BI Publisher), estabilizar e migrar fluxos do Apache NiFi para Python e elevar o nível de governança técnica do time.",
        "match": "Aplica um fechamento de altíssimo impacto (estilo McKinsey/Harvard), demonstrando proatividade, maturidade profissional e prontidão para o início imediato.",
        "growth": "Mostra que o candidato não encara o processo seletivo como um teste passivo, mas sim como o início de uma parceria estratégica resolutiva.",
        "case": "Integração Técnica Imediata e Geração de Valor PJ Híbrido.",
        "bullets": [
            "💼 [FECHAMENTO DE VALOR]: 'Estou pronto para integrar minhas habilidades avançadas em SQL e Python ao time, trazendo senso de dono, atenção a detalhes e resiliência para blindar seus pipelines de dados contra falhas e inconsistências.'",
            "⏳ [CONEXÃO E PRÓXIMOS PASSOS]: 'Agradeço pela excelente conversa e pela visão clara dos desafios da vaga. Ficou evidente que vocês precisam de um especialista resolutivo e investigativo, o que se alinha perfeitamente à minha trajetória de sucesso.'",
            "🌅 [SAUDAÇÕES PROFISSIONAIS]: Escolha conforme o momento: 1) 'Desejo uma semana excelente e produtiva.' | 2) 'Tenha um ótimo final de semana.' | 3) 'Foi um grande prazer compartilhar experiências hoje. Vamos nos falando.'"
        ],
        "qa_responses": [
            {"q": "Você teria facilidade em passar por um teste técnico prático de SQL avançado ou Python agora?", "a": "Seria um prazer enorme. Sinto total segurança no meu conhecimento prático em joins complexos, CTEs, funções analíticas em ambientes Oracle/Postgres e manipulação de bibliotecas como pandas/sqlalchemy em Python. Estou pronto para demonstrar minha velocidade de código, lógica limpa e visão de performance in qualquer desafio que vocês propuserem."},
            {"q": "Como você avalia sua capacidade de trabalhar em equipe compartilhando conhecimento?", "a": "Considero fundamental. Minha senioridade me ensinou que o sucesso de um projeto de dados nunca é individual. Apoiar tecnicamente membros mais juniores da equipe, revisar códigos (Code Review) visando melhoria contínua e manter o time alinhado sobre as inconsistências investigadas é a minha forma padrão de trabalhar."},
            {"q": "Qual é a sua disponibilidade para início?", "a": "Como possuo estrutura PJ totalmente regularizada e ativa, minha disponibilidade para início no modelo híbrido em São Paulo - SP é imediata, respeitando os trâmites normais de contratação da instituição. Estou pronto para focar 100% e assumir as responsabilidades da posição."}
        ]
    }
}


if "active_id" not in st.session_state:
    st.session_state.active_id = 1

with st.sidebar:
    st.markdown("### Workspace Input")
    cv_file = st.file_uploader("CV (PDF/TXT)", type=["txt", "pdf"], label_visibility="collapsed")
    jd_file = st.file_uploader("Job Description", type=["txt", "pdf"], label_visibility="collapsed")
    
    st.markdown("### Match Analytics")
    st.metric(label="Adherence Score", value="94%", delta="Elite Sênior Match")
        
    st.caption("**Target:** IURD · Analista de Dados Sênior")

# CORREÇÃO CRÍTICA: Ajustado para mapear exatamente as chaves de categoria presentes no dicionário
categories_list = [
    "WHAT - Capabilities & Profile", 
    "WHY - Intent & Fit", 
    "HOW - Case Methodology (STAR)", 
    "WHEN - Extreme Scenarios & Crisis"
]
cols = st.columns(4)

for idx, cat_name in enumerate(categories_list):
    with cols[idx]:
        # Exibe apenas o prefixo limpo na interface (ex: WHAT, WHY, HOW, WHEN)
        clean_header = cat_name.split(" - ")[0]
        st.markdown(f'<div class="category-header">{clean_header}</div>', unsafe_allow_html=True)
        
        cat_items = {k: v for k, v in DATA_MAPPING.items() if v["category"] == cat_name}
        
        for item_id, item_data in cat_items.items():
            is_active = (st.session_state.active_id == item_id)
            tag_token = f"[{item_data.get('tag', 'CONTEXT')}] "
            btn_label = f"▸ {tag_token}{item_data['title']}" if is_active else f"{tag_token}{item_data['title']}"
            
            st.markdown('<div class="trigger-btn">', unsafe_allow_html=True)
            if st.button(btn_label, key=f"btn_{item_id}"):
                st.session_state.active_id = item_id
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

st.markdown("<div style='margin-top: 0.6rem; border-top: 1px solid #e9ecef; margin-bottom: 0.4rem;'></div>", unsafe_allow_html=True)

active_data = DATA_MAPPING[st.session_state.active_id]

col_out1, col_out2 = st.columns([0.45, 0.55])

with col_out1:
    st.markdown(
        f"""
        <div class="response-box">
            <span style="color:#117a65; font-size:11px; font-weight:bold; text-transform:uppercase;">The Golden Bridge (Resposta Direta):</span><br>
            <p style="font-size:13px; color:#2c3e50; line-height:1.3; margin-top:4px;">"{active_data['bridge']}"</p>
        </div>
        """, 
        unsafe_allow_html=True
    )
    
    st.markdown(
        f"""
        <div style="background-color: #f8f9fa; padding: 10px; border-radius: 4px; border: 1px solid #e9ecef; margin-top: 0.4rem;">
            <span style="color:#2c3e50; font-size:10px; font-weight:bold; text-transform:uppercase;">Aprofundamento Estratégico:</span>
            <p style="font-size:12px; color:#4a5568; line-height:1.35; margin-top:4px;">{active_data['followup']}</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    st.markdown(f"<p style='font-size:11px; margin-top:0.4rem; color: #718096;'><strong>Âncora de Experiência:</strong> {active_data['case']}</p>", unsafe_allow_html=True)

with col_out2:
    st.markdown("<p style='font-weight:bold; font-size:12px; color:#1a202c; margin-bottom:0.4rem;'>Argumentos de Suporte (Lógica Base):</p>", unsafe_allow_html=True)
    for bullet in active_data["bullets"]:
        st.markdown(f"<p style='font-size:12px; margin-bottom:6px !important; color: #2d3748;'>• {bullet}</p>", unsafe_allow_html=True)
        
    st.markdown("<div style='margin-top:0.8rem; border-top: 1px dashed #e2e8f0; padding-top:0.4rem;'></div>", unsafe_allow_html=True)
    st.markdown("<p style='font-weight:bold; font-size:11px; color:#2b6cb0; margin-bottom:0.4rem; text-transform:uppercase;'>⚡ Simulação de Q&A Crítico de TI:</p>", unsafe_allow_html=True)
    
    for qa in active_data["qa_responses"]:
        st.markdown(f"""
        <div style="margin-bottom: 6px;">
            <strong style="font-size:11.5px; color:#2c5282; display:block;">Q: {qa['q']}</strong>
            <span style="font-size:11.5px; color:#4a5568; display:block; margin-top:1px;"><strong>A:</strong> {qa['a']}</span>
        </div>
        """, unsafe_allow_html=True)
