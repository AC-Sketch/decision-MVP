import streamlit as st

st.set_page_config(
    page_title="War Room - Brastemp Data & Analytics",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ADVANCED EXECUTIVE UX INJECTION ---
st.markdown("""
<style>
.block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 1.0rem !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
    max-width: 100% !important;
}
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
.response-box { background-color: #fff7ed; border-left: 5px solid #f97316; padding: 12px 16px !important; border-radius: 6px; margin-bottom: 0.75rem; }
.followup-box { background-color: #f8fafc; border-left: 5px solid #475569; padding: 12px 16px !important; border-radius: 6px; margin-bottom: 0.75rem; }
.growth-box { background-color: #fffbeb; border-left: 5px solid #d97706; padding: 12px 16px !important; border-radius: 6px; margin-bottom: 0.75rem; }
.match-box { background-color: #f0f9ff; border-left: 5px solid #0284c7; padding: 12px 16px !important; border-radius: 6px; margin-bottom: 0.75rem; }
.bullet-container-box { background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 6px; padding: 14px 18px !important; margin-bottom: 0.75rem; box-shadow: 0 1px 3px rgba(0,0,0,0.02); }
.qa-container-box { background-color: #f1f5f9; border: 1px solid #cbd5e1; border-left: 5px solid #0f172a; border-radius: 6px; padding: 14px 18px !important; }
.qa-item { margin-bottom: 8px !important; padding-bottom: 8px; border-bottom: 1px dashed #cbd5e1; }
.qa-item:last-child { border-bottom: none; margin-bottom: 0px !important; padding-bottom: 0px; }
[data-testid="stSidebarUserContent"] { padding-top: 1.5rem !important; }
.doc-container { background:#ffffff; border:1px solid #e2e8f0; border-radius:10px; padding:20px; }
.doc-title { font-size:22px; font-weight:800; color:#0f172a; margin-bottom:12px; }
.doc-subtitle { font-size:16px; font-weight:700; color:#f97316; margin-top:14px; margin-bottom:6px; }
.doc-section { font-size:14px; color:#334155; line-height:1.55; margin-bottom:10px; }
</style>
""", unsafe_allow_html=True)


def item(category, title, tag, bridge, followup, match, growth, case, bullets, qa):
    return {
        "category": category,
        "title": title,
        "tag": tag,
        "bridge": bridge,
        "followup": followup,
        "match": match,
        "growth": growth,
        "case": case,
        "bullets": bullets,
        "qa_responses": qa,
    }

DATA_MAPPING = {
    1: item("WHAT - Perfil & Stack", "Tell me about yourself", "PROFILE",
        "Sou um profissional sênior de dados com base em Engenharia de Produção, MBA e forte atuação em Analytics, BI, Governança, Cloud Analytics, Automação e FinOps.",
        "Trago experiência em empresas líderes como Heineken, Itaú, ASICS Latam, Gerdau, Sabesp, Fretebras e Afinz, sempre conectando dados complexos a decisões executivas e eficiência operacional.",
        "Para Brastemp/Whirlpool, isso se traduz em capacidade de transformar dados de vendas, produto, estoque, e-commerce, assistência técnica e consumidor em indicadores confiáveis.",
        "Posiciona o perfil como parceiro de negócio para digital, consumer experience, trade, supply chain e manufatura inteligente.",
        "Engenharia + MBA + Python/SQL + cases Heineken, Itaú, ASICS.",
        ["Automação de rotinas analíticas, reduzindo esforço operacional em mais de 80%.", "Modelagem de bases complexas com múltiplas nomenclaturas, canais e clientes.", "Otimização de ambientes cloud, reduzindo processamento de GB/TB para MB.", "Dashboards executivos para indicadores comerciais, financeiros, operacionais, marketing, e-commerce e eRetail."],
        [{"q":"Como sua experiência se conecta à Brastemp?", "a":"Brastemp combina marca forte, portfólio amplo, canais digitais, varejo, indústria e atendimento ao consumidor. Meu diferencial é integrar esses dados em uma visão confiável para acelerar decisões."}]
    ),
    2: item("WHY - Intenção & Fit", "Why Brastemp / Whirlpool?", "BRASTEMP-FIT",
        "Brastemp é uma marca icônica no Brasil e faz parte da Whirlpool, uma organização global de eletrodomésticos com enorme potencial de dados em produto, consumidor, indústria e canais digitais.",
        "A oportunidade é aplicar analytics em uma empresa que conecta manufatura, varejo, e-commerce, CX, logística e inovação para melhorar a vida das pessoas dentro de casa.",
        "Mostra alinhamento com uma empresa de bens de consumo duráveis onde dados precisam gerar impacto real: melhor previsão, menos ruptura, melhor experiência e mais eficiência.",
        "Conecta carreira em dados ao desafio de escala da Whirlpool no Brasil e América Latina.",
        "Experiência multi-indústria aplicada a consumer goods e varejo.",
        ["Forte aderência com dados de consumidor, produto, vendas e pós-venda.", "Capacidade de apoiar decisões em canais D2C, marketplaces, varejo e trade.", "Visão de eficiência operacional, FinOps e automação compatível com escala industrial."],
        [{"q":"Por que sair de um contexto puramente técnico para Brastemp?", "a":"Porque Brastemp permite transformar modelos, pipelines e dashboards em impacto visível no consumidor, no varejo e na operação industrial."}]
    ),
    3: item("WHY - Intenção & Fit", "Obsession with Data Quality", "DATA-QUALITY",
        "Em consumer goods, pequenas falhas em cadastro, SKU, preço, canal ou estoque distorcem toda a leitura comercial. Minha metodologia parte de qualidade, rastreabilidade e validação.",
        "Tenho experiência em normalizar bases fragmentadas, eliminar ambiguidades e criar camadas confiáveis de indicadores.",
        "Aderência direta a ambientes com grande portfólio de eletrodomésticos, peças, acessórios, campanhas e múltiplos canais.",
        "Reduz retrabalho, melhora forecast e aumenta confiança dos stakeholders no dado.",
        "Normalização Heineken + governança Afinz.",
        ["Validação por schema, granularidade e consistência histórica.", "Auditoria de nomenclaturas divergentes entre canais e sistemas.", "Camadas analíticas documentadas para tomada de decisão segura."],
        [{"q":"O que significa qualidade de dados para uma marca como Brastemp?", "a":"Significa garantir que SKU, produto, canal, preço, estoque, venda e atendimento conversem entre si sem ruído, permitindo decisões comerciais e operacionais confiáveis."}]
    ),
    4: item("WHAT - Perfil & Stack", "Value Proposition", "VALUE-PROP",
        "Entrego a ponte entre negócio e tecnologia: entendo o problema, modelo os dados, automatizo pipelines e comunico o impacto de forma executiva.",
        "Meu perfil combina Python, SQL, BI, cloud, FinOps, storytelling e domínio de indicadores de negócio.",
        "Ideal para squads que precisam transformar dados dispersos em produtos analíticos estáveis.",
        "Acelera rotinas de digital analytics, performance comercial, CX e supply chain analytics.",
        "Projetos end-to-end em ASICS, Itaú e Heineken.",
        ["Python, SQL, Power BI, Streamlit, AWS, GCP e BigQuery.", "Desenho de ETL/ELT e modelos dimensionais.", "Comunicação com executivos, áreas técnicas e áreas de negócio."],
        [{"q":"Qual é sua principal entrega?", "a":"Transformar bases caóticas em indicadores confiáveis, automatizados e acionáveis para reduzir custo, tempo e incerteza."}]
    ),
    5: item("WHY - Intenção & Fit", "Salary Expectations", "COMPENSATION",
        "Minha expectativa é compatível com um profissional sênior de dados capaz de gerar impacto em eficiência, automação e decisões estratégicas.",
        "Estou aberto a discutir o pacote considerando escopo, senioridade, modelo de trabalho, impacto esperado e trilha de crescimento.",
        "Mantém maturidade profissional e flexibilidade para negociação estruturada.",
        "Mostra foco em valor entregue, não apenas remuneração fixa.",
        "Alinhamento com mercado sênior de Data/Analytics.",
        ["Negociação baseada em impacto e responsabilidade.", "Abertura para CLT, PJ ou formato corporativo adequado.", "Foco em retorno operacional mensurável."],
        [{"q":"Você é flexível quanto ao modelo de contratação?", "a":"Sim, desde que haja clareza de escopo, responsabilidades, metas e possibilidades de crescimento."}]
    ),
    6: item("WHAT - Perfil & Stack", "Data Manipulation Stack", "TECH-STACK",
        "Uso Python e SQL como ferramentas centrais para limpar, transformar, validar e analisar grandes volumes de dados.",
        "Tenho experiência com AWS Athena, Glue, S3, BigQuery, Power BI, Pandas, NumPy, Scikit-Learn e Streamlit.",
        "Atende demandas de analytics em e-commerce, vendas, produto, CX, trade, logística e manufatura.",
        "Permite criar protótipos rápidos e pipelines escaláveis para áreas de negócio.",
        "AWS Itaú + GCP ASICS + Streamlit.",
        ["Análise exploratória e preparação de dados em Python.", "SQL avançado para views, particionamento e modelagem.", "Dashboards e aplicações analíticas interativas."],
        [{"q":"Você consegue atuar fora de BI tradicional?", "a":"Sim. Tenho experiência em pipeline, cloud, automação, modelagem e aplicações analíticas, não apenas visualização."}]
    ),
    7: item("WHAT - Perfil & Stack", "Agile & Product Delivery", "AGILE",
        "Atuo bem em squads ágeis, tratando dados como produto: backlog, escopo, documentação, validação e entrega incremental.",
        "Minha experiência com times de banco, varejo e bens de consumo me permite dialogar com negócio, engenharia e produto.",
        "Aderente a times digitais e cross-funcionais da Whirlpool.",
        "Reduz ruído de comunicação e aumenta previsibilidade de entrega.",
        "Squads NTT DATA/Itaú e projetos de BI corporativo.",
        ["User stories e critérios de aceite claros.", "Documentação em Confluence e governança de indicadores.", "MVPs estáveis com evolução incremental."],
        [{"q":"Como lida com mudança de escopo?", "a":"Reavalio dependências, risco de qualidade e impacto no negócio antes de adaptar a entrega."}]
    ),
    8: item("WHY - Intenção & Fit", "Short Tenures as Consulting Sprints", "SPRINTS",
        "Algumas passagens foram projetos intensivos com objetivo claro: resolver gargalos de dados, cloud ou automação rapidamente.",
        "Em ASICS e Itaú, entrei em ambientes complexos, entendi as estruturas e entreguei melhorias em curto prazo.",
        "Mostra curva de aprendizagem rápida e capacidade de gerar valor em ambientes já em movimento.",
        "Para Brastemp, isso significa capacidade de aterrissar em frentes críticas sem longo onboarding.",
        "Projetos 2025-2026 em cloud analytics.",
        ["Diagnóstico rápido de schema, gargalos e regras de negócio.", "Entrega documentada para continuidade interna.", "Foco em impacto mensurável e transferência de conhecimento."],
        [{"q":"Como entrega valor rápido?", "a":"Mapeio dados, regras e gargalos primeiro; depois priorizo automações e modelos que removem atrito operacional imediatamente."}]
    ),
    9: item("WHAT - Perfil & Stack", "Product & Consumer Analytics", "CONSUMER",
        "Conecto dados de produto, canal e comportamento do consumidor para explicar adoção, conversão, retenção e experiência.",
        "Brastemp tem um ecossistema rico: site oficial, linhas de produto, peças, acessórios, canais corporativos, varejo e pós-venda.",
        "Forte aderência a digital, D2C, e-commerce, CX e product analytics.",
        "Ajuda a entender jornada do consumidor e oportunidades de receita.",
        "Product analytics + digital channel tracking.",
        ["Métricas de conversão, funil, mix, ticket e disponibilidade.", "Análise de comportamento por categoria: cozinhar, gelar, lavar e beber.", "Dashboards para adoção, ruptura, preço e performance de campanha."],
        [{"q":"Como dados melhoram a experiência do consumidor?", "a":"Ao identificar gargalos de jornada, produtos com atrito, regiões com ruptura e padrões de atendimento que indicam oportunidade de melhoria."}]
    ),
    10: item("WHY - Intenção & Fit", "Innovation & AI Alignment", "AI-DATA",
        "A Whirlpool/Brastemp tem espaço para usar dados, automação e IA em previsão de demanda, CX, manutenção, personalização e manufatura inteligente.",
        "Minha trajetória em analytics e modelos estatísticos permite conectar IA a problemas práticos, sem perder governança.",
        "Alinha o perfil a inovação aplicada, não apenas experimentação técnica.",
        "Cria base para projetos de IA confiáveis e mensuráveis.",
        "Forecast, clustering e automação analítica.",
        ["Forecast de vendas, demanda e disponibilidade.", "Clusterização de consumidores, canais e produtos.", "Automação de análises recorrentes e alertas de anomalia."],
        [{"q":"Como vê IA em uma empresa de eletrodomésticos?", "a":"Como uma camada de inteligência para prever demanda, melhorar atendimento, otimizar estoque, personalizar ofertas e apoiar decisões industriais."}]
    ),
    11: item("HOW - Cases STAR", "ASICS Latam - Cloud & FinOps", "CLOUD",
        "Reestruturei um modelo financeiro multi-país em BigQuery para criar uma fonte única e otimizada de verdade analítica.",
        "O desafio envolvia dados do Brasil, Chile e Colômbia com layouts e regras regionais diferentes.",
        "Prova capacidade de lidar com operações multi-mercado e dados complexos.",
        "Aplicável à integração de dados de canais, regiões, fábricas e categorias da Whirlpool.",
        "Stalse / ASICS Latam - GCP.",
        ["Situação: dados fragmentados e custo elevado de processamento.", "Ação: redesign de views/tabelas com particionamento e cargas incrementais.", "Resultado: redução de processamento de GB/TB para MB."],
        [{"q":"Como isso se aplica à Brastemp?", "a":"Ambientes com muitos SKUs, canais e regiões exigem modelos enxutos, rastreáveis e baratos de processar."}]
    ),
    12: item("HOW - Cases STAR", "NTT DATA / Itaú - Massive Scale", "MASSIVE-SCALE",
        "Desenvolvi views e consultas em AWS Athena para apoiar relatórios executivos em ambiente de alto volume.",
        "O projeto exigia consistência em bases grandes e regras de negócio em evolução.",
        "Demonstra robustez técnica para dados corporativos críticos.",
        "Útil para escalar indicadores de vendas, logística, assistência e operação.",
        "AWS Athena, Glue e S3.",
        ["Situação: alto volume e regras complexas.", "Ação: views otimizadas e lógicas parametrizadas.", "Resultado: relatórios automatizados e consistentes."],
        [{"q":"Como mantém integridade em volume alto?", "a":"Com particionamento, schemas bem definidos, validações e rastreabilidade de regras."}]
    ),
    13: item("HOW - Cases STAR", "Heineken - Data Normalization", "NORMALIZATION",
        "Estruturei modelo dimensional para unificar dados caóticos de e-commerce e canais digitais.",
        "O desafio incluía mais de 10 mil variações de nomenclatura para um mesmo universo de produtos.",
        "Altamente relevante para portfólios com muitos SKUs, categorias, linhas e canais.",
        "Ajuda Brastemp a consolidar visão por produto, canal, campanha e cliente.",
        "Heineken Digital Analytics.",
        ["Situação: taxonomia inconsistente entre clientes e sistemas.", "Ação: star schema, mapeamento e rotinas de limpeza.", "Resultado: indicadores conciliados e auditáveis."],
        [{"q":"Qual paralelo com Brastemp?", "a":"Produtos, peças, acessórios, combos e canais precisam de taxonomia confiável para evitar erro em vendas, margem, estoque e campanha."}]
    ),
    14: item("HOW - Cases STAR", "Afinz - Automation & Governance", "AUTOMATION",
        "Automatizei rotinas manuais com Python e SQL, reduzindo o ciclo diário de atualização de 1h30 para 15 minutos.",
        "Também organizei metadados e documentação para fortalecer governança.",
        "Mostra capacidade de eliminar gargalos operacionais recorrentes.",
        "Aplicável a rotinas comerciais, financeiras, operacionais e de CX.",
        "Afinz/Sorocred MIS Analytics.",
        ["Situação: processos manuais e frágeis.", "Ação: pipelines automatizados e validações.", "Resultado: mais velocidade, confiabilidade e rastreabilidade."],
        [{"q":"Por que documentação é crítica?", "a":"Porque indicador sem definição e linhagem vira ruído; documentação garante continuidade e confiança."}]
    ),
    15: item("HOW - Cases STAR", "Burity - Risk, Logic & Standards", "RISK",
        "Atuei em processos regulatórios, contratos e ativos de alto valor, desenvolvendo precisão, leitura crítica e gestão de risco.",
        "Essa experiência fortalece minha capacidade de auditar regras, inconsistências e dependências.",
        "Útil em governança de dados, compliance, LGPD, contratos e processos corporativos.",
        "Aumenta confiabilidade em decisões baseadas em dados sensíveis.",
        "Gestão patrimonial, contratos e risco.",
        ["Interpretação de normas e documentação complexa.", "Mitigação de risco com evidência e rastreabilidade.", "Pensamento lógico aplicado a governança."],
        [{"q":"Como isso entra em dados?", "a":"Governança de dados exige a mesma disciplina: regra clara, evidência, versionamento e responsabilidade."}]
    ),
    16: item("WHEN - Crise & Fechamento", "Handling a Major Data Error", "CONTAINMENT",
        "Quando uma regra quebra ou um dado se corrompe, minha prioridade é conter, medir impacto, comunicar e corrigir a causa raiz.",
        "Eu isolo janela afetada, verifico logs, valido a exposição e implemento proteção permanente.",
        "Mostra maturidade para lidar com incidentes em ambientes críticos.",
        "Protege decisões comerciais, operacionais e financeiras.",
        "Root-cause analysis e data quality filters.",
        ["Isolamento da camada afetada.", "Análise de causa raiz em logs e transformações.", "Nova validação automatizada para evitar recorrência."],
        [{"q":"Como comunica erro crítico?", "a":"Com transparência: causa, período afetado, impacto mensurado, correção e prevenção."}]
    ),
    17: item("WHEN - Crise & Fechamento", "High-Pressure Pivot", "PIVOT",
        "Em demandas urgentes, organizo o problema por entradas, regras, restrições e decisão esperada.",
        "Evito improviso: uso amostras confiáveis, análises rápidas e entregáveis mínimos seguros.",
        "Aderente a ambientes com sazonalidade, lançamentos e pressão de mercado.",
        "Ajuda em campanhas, Black Friday, ruptura, recall analítico ou mudanças de meta.",
        "Exploratory analysis e MVP analytics.",
        ["Separar ruído de causa real.", "Construir baseline confiável.", "Entregar visão acionável sem sacrificar qualidade."],
        [{"q":"Como reage a pedido urgente indefinido?", "a":"Primeiro estruturo a pergunta de negócio, depois monto uma análise mínima confiável para orientar decisão."}]
    ),
    18: item("WHEN - Crise & Fechamento", "Conflict with Stakeholders", "STAKEHOLDERS",
        "Reduzo conflito trazendo evidência, linhagem e impacto dos cenários para a conversa.",
        "Stakeholders discordam menos quando enxergam regra, origem e consequência do dado.",
        "Importante para áreas como comercial, trade, digital, supply e financeiro.",
        "Cria consenso e velocidade em decisões cross-funcionais.",
        "Data-driven consensus frameworks.",
        ["Escuta ativa para mapear o problema real.", "Comparação objetiva de cenários.", "Decisão guiada por evidência e impacto."],
        [{"q":"Como lida com pressão por número enviesado?", "a":"Mostro tecnicamente o risco da distorção e proponho uma alternativa validada e auditável."}]
    ),
    19: item("WHEN - Crise & Fechamento", "Technical Storytelling", "STORYTELLING",
        "Transformo queries, modelos e estatísticas em impacto de negócio: receita, margem, tempo, custo, ruptura, conversão e satisfação.",
        "Executivos não precisam ver toda a sintaxe; precisam entender estabilidade, risco e caminho de ação.",
        "Fortalece apresentações para lideranças e áreas não técnicas.",
        "Aumenta adoção dos produtos analíticos.",
        "Executive analytics storytelling.",
        ["Começar pelo insight e impacto.", "Explicar limites e confiabilidade do modelo.", "Fechar com decisão recomendada e próximos passos."],
        [{"q":"Como apresenta dados para não técnicos?", "a":"Uso narrativa top-down: conclusão, evidência, impacto, risco e ação."}]
    ),
    20: item("WHEN - Crise & Fechamento", "Cultural Fit & Closing", "GLOBAL-MINDSET",
        "Brastemp exige rigor técnico, sensibilidade de consumidor e colaboração entre indústria, digital, varejo e marca. Esse é exatamente o tipo de ambiente em que consigo gerar valor.",
        "Tenho inglês avançado, vivência internacional e histórico de adaptação a diferentes setores e culturas corporativas.",
        "Combina autonomia, comunicação e foco em entrega.",
        "Fecha a entrevista com posicionamento de longo prazo e mentalidade de dono.",
        "Experiência internacional + disciplina + adaptabilidade.",
        ["Curiosidade para entender produtos, consumidores e operações.", "Capacidade de transitar entre técnico e executivo.", "Foco em qualidade, documentação e impacto mensurável."],
        [{"q":"Por que você seria um bom fit cultural?", "a":"Porque combino rigor analítico, comunicação clara, adaptabilidade e vontade de gerar impacto real em uma marca brasileira de enorme relevância."}]
    ),
}

if "view_mode" not in st.session_state:
    st.session_state.view_mode = "Main Interview Board"

with st.sidebar:
    st.markdown("### Workspace Mode")
    if st.button("📊 Main Interview Board", use_container_width=True):
        st.session_state.view_mode = "Main Interview Board"
        st.rerun()
    if st.button("📄 View: André Carvalho Resume", use_container_width=True):
        st.session_state.view_mode = "CV Doc"
        st.rerun()
    if st.button("📘 View: Brastemp Target Matrix", use_container_width=True):
        st.session_state.view_mode = "Guide Doc"
        st.rerun()

    st.markdown("---")
    st.markdown("### Strategic Framework")
    st.info("**WHY:** Intenção & Fit\n\n**WHAT:** Perfil & Stack\n\n**HOW:** Cases STAR\n\n**WHEN:** Crise & Fechamento")

    st.markdown("### Alignment Analytics")
    st.metric(label="Brastemp / Whirlpool Job Fit", value="96%", delta="Consumer Analytics + Cloud Match")
    st.caption("**Target:** Brastemp · Whirlpool Brasil · Data / Analytics")

if st.session_state.view_mode == "CV Doc":
    st.markdown('<div class="doc-container"><div class="doc-title">Document View: André Carvalho — Resume Adaptado para Brastemp</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="doc-section">
        <strong>ENG. ANDRÉ CARVALHO, MBA</strong><br>
        Senior Data Professional | Advanced Analytics | Cloud Data | SQL & Python | BI | Consumer & Business Analytics
    </div>
    <div class="doc-subtitle">Professional Summary</div>
    <div class="doc-section">
        Profissional sênior de dados com formação em Engenharia de Produção e MBA pela FGV. Experiência em modelagem, automação, governança, dashboards executivos, cloud analytics e otimização de pipelines para empresas líderes. Perfil altamente aderente a ambientes de bens de consumo, varejo, e-commerce, supply chain, CX e manufatura orientada por dados.
    </div>
    <div class="doc-subtitle">Core Corporate Milestones</div>
    <div class="doc-section">
        • <strong>ASICS Latam:</strong> redesenho de views em BigQuery e redução de processamento de GB/TB para MB.<br>
        • <strong>Itaú / NTT DATA:</strong> consultas complexas em AWS Athena, Glue e S3 para ambientes de alto volume.<br>
        • <strong>Heineken:</strong> normalização de dados digitais e taxonomia com mais de 10.000 variações de nomenclatura.<br>
        • <strong>Afinz:</strong> automação com Python e SQL, reduzindo rotinas de 1h30 para 15 minutos.
    </div>
    </div>
    """, unsafe_allow_html=True)

elif st.session_state.view_mode == "Guide Doc":
    st.markdown('<div class="doc-container"><div class="doc-title">Document View: Brastemp / Whirlpool Target Matrix</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="doc-section"><strong>Brastemp / Whirlpool Alignment Paradigm</strong></div>
    <div class="doc-subtitle">Como o perfil responde aos desafios de dados da empresa:</div>
    <div class="doc-section">
        1. <strong>Consumer & Product Analytics:</strong> conectar comportamento do consumidor, produto, canal, vendas e pós-venda.<br>
        2. <strong>Data Quality:</strong> criar taxonomias confiáveis para SKUs, categorias, peças, campanhas e canais.<br>
        3. <strong>Cloud & FinOps:</strong> otimizar pipelines e reduzir custo computacional em ambientes analíticos.<br>
        4. <strong>Automation:</strong> eliminar rotinas manuais em relatórios, forecast, trade, supply e performance digital.<br>
        5. <strong>Storytelling Executivo:</strong> transformar análises técnicas em decisões claras para negócio, produto e operação.
    </div>
    </div>
    """, unsafe_allow_html=True)

else:
    st.title("🟠 Brastemp Data & Analytics Strategy Board")
    st.caption("Strategic execution dashboard for interview, analytics alignment and business impact storytelling.")

    tab_categories = [
        "📋 WHAT - Perfil & Stack",
        "🎯 WHY - Intenção & Fit",
        "🚀 HOW - Cases STAR",
        "⚡ WHEN - Crise & Fechamento"
    ]

    tab_objs = st.tabs(tab_categories)
    category_mapping = {
        "WHAT - Perfil & Stack": tab_objs[0],
        "WHY - Intenção & Fit": tab_objs[1],
        "HOW - Cases STAR": tab_objs[2],
        "WHEN - Crise & Fechamento": tab_objs[3]
    }

    for cat_name, tab_obj in category_mapping.items():
        with tab_obj:
            cat_items = {k: v for k, v in DATA_MAPPING.items() if v["category"] == cat_name}
            for item_id, item_data in cat_items.items():
                with st.expander(f"🔶 [{item_data['tag']}] — {item_data['title']}", expanded=(item_id == 1)):
                    col_out1, col_out2 = st.columns([0.50, 0.50])
                    with col_out1:
                        st.markdown(f"""
                        <div class="response-box">
                            <span style="color:#f97316; font-size:10px; font-weight:bold; text-transform:uppercase;">The Golden Bridge:</span><br>
                            <strong style="font-size:13px; color:#1e293b; line-height:1.2;">"{item_data['bridge']}"</strong>
                        </div>
                        <div class="followup-box">
                            <span style="color:#475569; font-size:10px; font-weight:bold; text-transform:uppercase;">Context / Elaboration:</span><br>
                            <p style="font-size:12.5px; color:#334155; line-height:1.3;">{item_data['followup']}</p>
                        </div>
                        <div class="growth-box">
                            <span style="color:#d97706; font-size:10px; font-weight:bold; text-transform:uppercase;">The Brastemp Strategic Fit:</span><br>
                            <p style="font-size:12.5px; color:#78350f; line-height:1.3;">{item_data['growth']}</p>
                        </div>
                        <div class="match-box">
                            <span style="color:#0284c7; font-size:10px; font-weight:bold; text-transform:uppercase;">Core Concept Objective:</span><br>
                            <p style="font-size:12.5px; color:#0369a1; line-height:1.3;">{item_data['match']}</p>
                        </div>
                        """, unsafe_allow_html=True)

                    with col_out2:
                        bullets_html = "".join(f'<li style="font-size:12.5px; color:#334155; line-height:1.35; margin-bottom:4px;">{b}</li>' for b in item_data['bullets'])
                        st.markdown(f"""
                        <div class="bullet-container-box">
                            <span style="color:#1e293b; font-size:10.5px; font-weight:bold; text-transform:uppercase; display:block; margin-bottom:6px;">Supporting Architectural Arguments:</span>
                            <ul style="margin-top:0px; padding-left:18px;">{bullets_html}</ul>
                            <p style='font-size:10.5px; color:#64748b; margin-top:6px;'><strong>Baseline Track:</strong> {item_data['case']}</p>
                        </div>
                        """, unsafe_allow_html=True)

                        st.markdown("""<div class="qa-container-box"><span style="color:#0f172a; font-size:10.5px; font-weight:bold; text-transform:uppercase; display:block; margin-bottom:6px;">⚡ TOUGHEST BRAS­TEMP / WHIRLPOOL Q&A SIMULATOR:</span>""", unsafe_allow_html=True)
                        for qa in item_data['qa_responses']:
                            st.markdown(f"""
                            <div class="qa-item">
                                <strong style="font-size:12px; color:#0f172a; display:block; line-height:1.2;">Q: {qa['q']}</strong>
                                <p style="font-size:12px; color:#1e293b; line-height:1.35; margin-top:2px !important;"><strong>A:</strong> {qa['a']}</p>
                            </div>
                            """, unsafe_allow_html=True)
                        st.markdown("""</div>""", unsafe_allow_html=True)
