import streamlit as st

st.set_page_config(
    page_title="War Room - Ambev BA/Data",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Fluid & Non-Overlapping UX Injection
st.markdown("""
<style>
/* Reset main padding limits to prevent overlapping headers */
.block-container {
    padding-top: 1.0rem !important;
    padding-bottom: 0.1rem !important;
    padding-left: 1rem !important;
    padding-right: 1rem !important;
    max-width: 100% !important;
}

/* Global scrollbar behavior control for clean hardware feel */
::-webkit-scrollbar {
    display: none !important;
}

div[data-testid="stVerticalBlock"] {
    gap: 0.15rem !important;
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

/* Force column button wrappers to have exact uniform heights */
div.stButton > button {
    width: 100% !important;
    height: 44px !important; 
    white-space: normal !important; 
    word-break: keep-all !important;
    overflow: hidden !important;
    font-size: 10px !important;
    line-height: 1.15 !important;
    padding: 0.3rem 0.3rem !important;
    text-align: center !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    border-radius: 5px !important;
    margin-bottom: 4px !important;
}

.category-header {
    font-size: 11px !important;
    font-weight: bold !important;
    color: #2c3e50;
    border-bottom: 2px solid #e9ecef;
    padding-bottom: 3px;
    margin-bottom: 0.4rem !important;
    text-transform: uppercase;
    letter-spacing: 0.3px;
}

/* Output Display Blocks - Compressed & Symmetrical heights */
.response-box {
    background-color: #e8f8f5;
    border-left: 4px solid #18bc9c;
    padding: 6px 10px !important;
    border-radius: 4px;
    margin-bottom: 0.3rem;
    min-height: 48px;
}

.followup-box {
    background-color: #f4f6f7;
    border-left: 4px solid #34495e;
    padding: 6px 10px !important;
    border-radius: 4px;
    margin-bottom: 0.3rem;
    min-height: 48px;
}

.growth-box {
    background-color: #fef9e7;
    border-left: 4px solid #f39c12;
    padding: 5px 10px !important;
    border-radius: 4px;
    margin-bottom: 0.3rem;
    min-height: 40px;
}

.match-box {
    background-color: #ebf5fb;
    border-left: 4px solid #3498db;
    padding: 5px 10px !important;
    border-radius: 4px;
    margin-bottom: 0.3rem;
    min-height: 40px;
}

.bullet-container-box {
    background-color: #ffffff;
    border: 1px solid #e9ecef;
    border-radius: 4px;
    padding: 6px 10px !important;
    min-height: 120px;
}

.qa-container-box {
    background-color: #f2f4f4;
    border: 1px solid #d5dbdb;
    border-left: 4px solid #1b4f72;
    border-radius: 4px;
    padding: 6px 10px !important;
    margin-top: 3px;
    min-height: 180px;
}

.qa-item {
    margin-bottom: 4px !important;
    padding-bottom: 3px;
    border-bottom: 1px dashed #d5dbdb;
}
.qa-item:last-child {
    border-bottom: none;
    margin-bottom: 0px !important;
}

/* Embedded Document Viewer Styles with Header Protection Gap */
.doc-container {
    background-color: #ffffff;
    border: 1px solid #d5dbdb;
    border-radius: 6px;
    padding: 24px !important;
    padding-top: 25px !important;
    box-shadow: 0 4px 6px rgba(0,0,0,0.02);
    max-height: 78vh;
    overflow-y: auto !important;
}
.doc-title {
    color: #1b4f72;
    font-size: 18px;
    font-weight: bold;
    border-bottom: 3px solid #1b4f72;
    padding-bottom: 8px;
    margin-bottom: 18px !important;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
.doc-section {
    font-size: 13px;
    color: #2c3e50;
    margin-bottom: 12px !important;
    line-height: 1.5;
}
.doc-subtitle {
    font-size: 14px;
    color: #154360;
    font-weight: bold;
    margin-top: 16px !important;
    margin-bottom: 6px !important;
    border-bottom: 1px solid #eaeded;
    padding-bottom: 2px;
}
.commentary-box {
    background-color: #ebf5fb;
    border-left: 4px solid #2980b9;
    padding: 10px !important;
    margin-top: 6px !important;
    margin-bottom: 12px !important;
    border-radius: 4px;
    font-size: 12.5px;
}
</style>
""", unsafe_allow_html=True)

# 20 Strategic Framework Database Items - Syntactically Audited & Corrected
DATA_MAPPING = {
    1: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Tell me about yourself",
        "tag": "PROFILE",
        "bridge": "Analista de Dados Sênior com formação em Engenharia de Produção, MBA em Administração e experiência em Analytics, Business Intelligence, Governança de Dados, Cloud Analytics e FinOps.",
        "followup": "Atuação em projetos estratégicos para empresas como Heineken, Itaú, ASICS Latam, Gerdau, Sabesp, Fretebras, Afinz e outras, apoiando a tomada de decisão por meio da transformação de dados em insights acionáveis e indicadores de negócio.",
        "match": "Experiência em modelagem de dados, desenvolvimento de pipelines ETL/ELT, automação de processos, consolidação de KPIs, dashboards executivos e otimização de ambientes analíticos em nuvem.",
        "growth": "Foco total na agenda de Data Products e eficiência operacional que suportam o crescimento acelerado do ecossistema Ambev.",
        "case": "Engenharia de Produção + MBA + Heineken, Itaú, ASICS Latam.",
        "bullets": [
            "Automação de processos analíticos e fluxos de atualização de relatórios, reduzindo em mais de 80% o esforço operacional. Diminuição do tempo de atualização de dashboards de 1h30 para 15 minutos por meio da implementação de pipelines de dados estruturados, automações e padronização de processos.",
            "Consolidação e modelagem de múltiplas bases de dados complexas para grandes operações de varejo e bens de consumo, incluindo cenários em que um único produto possuía mais de 10 mil denominações distintas entre clientes, canais e sistemas. Estruturação de modelos analíticos capazes de unificar e correlacionar informações dispersas para generation de indicadores confiáveis e suporte à tomada de decisão.",
            "Reestruturação de arquiteturas analíticas em ambiente cloud, redesenhando múltiplas tabelas e views para uma arquitetura escalável e otimizada. Implementação de particionamento, processamento incremental e práticas de FinOps, reduzindo o consumo de processamento de dados de GB/TB para MB.",
            "Desenvolvimento de dashboards executivos e soluções de Analytics para monitoramento de indicadores comerciais, financeiros, operacionais, marketing, eCommerce e eRetail, incluindo operações LATAM (Brasil, Chile e Colômbia). Aplicação de Machine Learning e Analytics Avançado para projeção de receita e identificação de padrões de negócio utilizando modelos como ARIMA+ e K-Means."
        ],
        "qa_responses": [
            {"q": "Como sua formação apoia a interface entre negócio e tecnologia?", "a": "A Engenharia de Produção e o MBA me dão a visão de processo e P&L necessários para entender as dores dos diretores de negócio da Ambev, enquanto meu domínio técnico em SQL e arquiteturas cloud me permite traduzir essas necessidades em especificações exatas para os engenheiros de dados."},
            {"q": "Qual o seu diferencial para atuar em um ambiente global como a Ambev?", "a": "Tenho vivência internacional no Reino Unido, França e Espanha, e inglês avançado testado corporativamente. Estou acostumado a interagir com stakeholders de diferentes culturas e alinhar requisitos complexos de forma clara e concisa."},
            {"q": "Como você lida com dados caóticos e descentralizados?", "a": "Na Heineken, unifiquei bases onde o mesmo produto tinha mais de 10 mil denominações. Trato a bagunça de dados não como um problema intransponível, mas como uma oportunidade de implementar Data Quality e criar uma 'single source of truth' auditável."}
        ]
    },
    2: {
        "category": "WHY - Intent & Fit",
        "title": "Why Ambev?",
        "tag": "AMBEV-FIT",
        "bridge": "Ambev is the benchmark for execution, meritocracy, and data-driven culture, and I want to scale my analytical framework inside an ecosystem where optimization directly shifts massive market results.",
        "followup": "Your scale requires high-performance execution. Having integrated digital channel data at *Heineken* and multicountry cloud models at *ASICS*, I understand FMCG dynamics and the exact technical bottlenecks of large-scale digital distribution.",
        "match": "Proves deep alignment with Ambev’s fast-paced, owner-mindset culture and digital products expansion (BEES/Ze Delivery).",
        "growth": "As Ambev consolidates its tech-first operations, my profile guarantees a seamless translation of retail complexity into robust data governance structures.",
        "case": "FMCG Analytics Architecture (*Heineken*) + Cloud Scalability (*ASICS* / *Itaú*).",
        "bullets": [
            "Ambev’s massive transaction volume demands structured Data Products, not fragmented spreadsheet management.",
            "My background allows me to communicate fluently with commercial stakeholders and tech squads with the same velocity.",
            "I thrive in meritocratic cultures where process optimization is backed by indisputable database evidence."
        ],
        "qa_responses": [
            {"q": "Why transition from pure consulting/agencies to Ambev?", "a": "Consulting gives you cross-industry velocity, but Ambev offers the ultimate scale. I want to deploy my data stack where the optimization of a pipeline directly affects the commercial efficiency of thousands of point-of-sales."},
            {"q": "How does your experience fit our digital evolution (Data Products)?", "a": "A true Data Product must be robust, documented, and have trusted Data Quality. At *Afinz*, I built metadata repositories and slashed processing time by 80%, proving I build data solutions meant to scale safely."},
            {"q": "Are you comfortable with our fast and sometimes high-pressure environment?", "a": "Absolutely. My background is in Production Engineering. I see pressure as a parameter to optimize, and chaotic scenarios as the perfect landscape to build agile automation."}
        ]
    },
    3: {
        "category": "WHY - Intent & Fit",
        "title": "Role: Business & Data Liaison",
        "tag": "LIAISON",
        "bridge": "I act as a strategic translator, ensuring that corporate growth goals do not break down due to hidden engineering logic or poor data governance rules.",
        "followup": "Technical teams focus on code syntax, while business teams focus on ROI and market share. I bridge this gap by documenting clean business definitions and writing the SQL validation rules that guarantee technical execution.",
        "match": "Positions you as the ideal Business Analyst who can code and enforce data governance without losing business speed.",
        "growth": "Ambev requires data assets that think like owners and execute with engineering precision; I provide that exact balance.",
        "case": "Requirement Gathering + Data Governance Layers (*Afinz* / *Itaú*).",
        "bullets": [
            "I translate commercial KPIs into technical data mapping schemas for Databricks or BigQuery environments.",
            "I use agile methodologies to ensure data product development sprints match evolving business realities.",
            "I believe documentation in *Confluence* is just as critical as writing optimized code."
        ],
        "qa_responses": [
            {"q": "How do you ensure non-technical stakeholders understand technical limitations?", "a": "I don't explain the database structure; I explain the metric's reliability. If a pipeline lag impacts a commercial dashboard, I present the business exposure and the automated correction framework in clear financial terms."},
            {"q": "What is your methodology for gathering requirements?", "a": "I run a structured process: first, mapping the stakeholder's decision-making goal; second, tracing the raw data availability; and third, detailing the Data Quality rules required before any dashboard or data product is deployed."},
            {"q": "How do you ensure agile squads deliver what business actually requested?", "a": "By writing bulletproof user stories with clear technical acceptance criteria, including verified SQL queries that act as the benchmark for data validation during the pipeline development."}
        ]
    },
    4: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Your Value Proposition",
        "tag": "VALUE",
        "bridge": "I bring process engineering discipline, high-volume cloud data analytics, and a rigorous data quality lens to Ambev’s business analytics framework.",
        "followup": "I don't just consume dashboards; I audit the underlying data architecture to ensure data products are scalable, optimized for FinOps, and 100% reliable for executive decision-making.",
        "match": "Directly links your analytical depth to Ambev's need for lean, automated, and hyper-reliable data products.",
        "growth": "Slashes processing bottlenecks and operational manual lag, unlocking immediate stakeholder data trust.",
        "case": "FinOps + Database Transformation (*ASICS* / *Heineken*).",
        "bullets": [
            "I turn manual spreadsheet chaos into automated cloud pipelines, saving hundreds of operational hours.",
            "I implement strict Data Quality frameworks to eliminate reporting discrepancies before they reach C-level rooms.",
            "I optimize query spending (FinOps), scaling processing capability while keeping infrastructure lean."
        ],
        "qa_responses": [
            {"q": "What is your immediate 30-day plan if hired?", "a": "Map out your current business reporting pain points, audit the data streams feeding your key dashboards, and identify manual dependencies to optimize them into automated, documented data workflows."},
            {"q": "How do you define a successful Data Product?", "a": "A successful Data Product must have three pillars: it must solve a recurring business question, it must have automated Data Quality checks embedded, and its metadata must be fully documented for any stakeholder to use seamlessly."},
            {"q": "How do you measure your own impact as an analyst?", "a": "Through two hard metrics: operational hours saved by automation and data infrastructure cost reduction via proper query optimization and FinOps."}
        ]
    },
    5: {
        "category": "WHY - Intent & Fit",
        "title": "Salary Expectations",
        "tag": "ANCHOR",
        "bridge": "My salary expectations are aligned with the market standard for a senior data analytics profile who brings immediate execution value.",
        "followup": "I am open to Ambev's compensation architecture, especially considering the professional growth roadmap and the impact incentives of the ecosystem.",
        "match": "Maintains professional transparency and flexibility while establishing solid seniority benchmarks.",
        "growth": "Demonstrates owner mindset, anchoring the compensation discussion around delivered efficiency.",
        "case": "Standard Corporate Range Alignment.",
        "bullets": [
            "My baseline benchmark is targeted around standard Senior Analyst/Business Consultant structures.",
            "I prioritize total compensation fit, including corporate benefits, bonus alignment, and remote flexibility.",
            "My delivery in process automation and data cloud optimization naturally offsets resource overhead."
        ],
        "qa_responses": [
            {"q": "What is your specific current expectation number?", "a": "I am looking to align within the corporate benchmark for Senior Business/Data Analytics roles, and I am completely flexible to evaluate your standard offer package including your variable performance incentives."},
            {"q": "Are you open to hybrid models if required?", "a": "Yes, while my setup is highly optimized for agile remote work across global teams, I am completely comfortable aligning with Ambev's internal collaboration framework."}
        ]
    },
    6: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Tech Stack: Databricks & Azure",
        "tag": "CLOUD-STACK",
        "bridge": "I treat cloud data platforms as logical ecosystems where the core architectural principles of optimization, partition, and governance remain identical.",
        "followup": "While my recent fast sprints deployed massive workloads in Google BigQuery (*ASICS*) and AWS Athena (*Itaú*), my engineering background allows me to master Azure and Databricks with zero onboarding lag.",
        "match": "Reframes tech stack transitions as a highly transferable data engineering capability.",
        "growth": "Ambev uses modern distributed frameworks; my deep SQL/Python optimization background fits Databricks logic perfectly.",
        "case": "Cross-Cloud Adaptability (*AWS Athena* to *GCP BigQuery* in under a year).",
        "bullets": [
            "Writing optimized SQL views and executing incremental loads follows the same architectural best practices in Azure Data Factory as in AWS Glue.",
            "My strong Python foundation enables me to interface cleanly with Spark notebooks inside Databricks ecosystems.",
            "I have a track record of slashing data consumption from Gigabytes to Megabytes by mastering database engines rapidly."
        ],
        "qa_responses": [
            {"q": "We use Databricks heavily. How long will it take for you to be independent?", "a": "Day one. I already use advanced Python, SQL, and query optimization. Transitioning from Athena or BigQuery to Databricks is just a syntax calibration. The underlying logic of data distribution and partitioning is my daily routine."},
            {"q": "Have you worked with distributed data frameworks like Spark?", "a": "In my cloud data projects, I handle high-volume data structures (billions of rows at *Itaú* via AWS). I write my Python scripts with optimization and scaling in mind, ensuring seamless integration into Spark architectures."},
            {"q": "How do you approach learning a new data tool required by the squad?", "a": "With high speed. My career shows rapid adaptation: I deployed data solutions across AWS and GCP back-to-back, leveraging tools like Streamlit, Looker, and QuickSight by focusing on fundamental data engineering rules."}
        ]
    },
    7: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Data Governance & Quality",
        "tag": "GOVERNANCE",
        "bridge": "I believe that an analytics dashboard without built-in Data Quality controls is just an automated way of spreading wrong business assumptions.",
        "followup": "At *Heineken* and *Afinz*, I built data validation loops and metadata catalogs to ensure that when an executive looked at a sales indicator, it matched the absolute financial truth to the exact penny.",
        "match": "Directly satisfies Ambev's desire for an analyst who actively guards information reliability and governance.",
        "growth": "Protects decision speed by building structural trust inside every developed data product.",
        "case": "Star Schema Modeling (*Heineken*) + Metadata Repositories (*Afinz*).",
        "bullets": [
            "I design clean dimensional models (Star Schema) that force messy, fragmented inputs into standardized, auditable layouts.",
            "I automate data validation scripts to flag schema drift, duplicate keys, or volume anomalies before the business notices them.",
            "I map out data lineages so any internal user can easily verify the origin and calculation logic of every metric."
        ],
        "qa_responses": [
            {"q": "How do you implement Data Quality inside a raw, messy ingestion pipeline?", "a": "By establishing staging validation gates. Before raw rows are transformed into a Data Product layer, I deploy script-based checks to audit data types, null counts, and business rule consistency, isolating anomalies instantly."},
            {"q": "Can you give an example of a Data Governance optimization you delivered?", "a": "At *Afinz*, I documented and structured metadata organization layers using *Confluence* and SQL repositories, transforming a siloed, manual 1.5-hour reporting struggle into a transparent, governed 15-minute pipeline."},
            {"q": "How do you convince business users to respect Data Governance rules?", "a": "By showing them how governance prevents errors that damage their targets. When a commercial team realizes that structured metadata ensures their campaign bonuses are calculated accurately and without lag, they become your biggest allies."}
        ]
    },
    8: {
        "category": "WHY - Intent & Fit",
        "title": "Short Tenures (Stalse/NTT)",
        "tag": "AGILITY",
        "bridge": "These roles were structured, high-intensity contract projects brought in as tactical consulting sprints to solve specific infrastructure blocks.",
        "followup": "At *Stalse*, my mission was to unify Latin American revenue models for *ASICS*. At *NTT Data*, I was brought in to optimize complex SQL views for *Itaú*. Once the architecture was optimized and documented, the goals were met.",
        "match": "Frames short tenures as high-impact, intentional consulting deliveries rather than professional instability.",
        "growth": "Proves that you have an extreme learning curve and can deliver immediate business results without onboarding friction.",
        "case": "High-Impact Temporary Sprints (2025 - 2026).",
        "bullets": [
            "I am highly adaptable; I entered complex corporate cloud environments and delivered automated tools in less than 4 months.",
            "Every project I touch is left fully documented, ensuring seamless handover and no technical debt for the internal squads.",
            "Now, I am looking to invest this cross-industry cloud arsenal into a permanent, long-term journey inside Ambev."
        ],
        "qa_responses": [
            {"q": "Why should we trust you will stay at Ambev long-term?", "a": "Because Ambev has the scale and systemic complexity that keeps a senior analytical mind challenged. Consulting sprints are great for learning, but I want to own a product long-term, watching my data architectures scale alongside your global business lines."},
            {"q": "How did you manage to deliver results so fast at ASICS and Itaú?", "a": "By utilizing structured thinking and a plug-and-play technical mindset. I didn't wait for weeks of onboarding; I immediately analyzed the database schemas, identified the bottlenecks, and started writing optimization code from week two."},
            {"q": "Did these short projects involve stakeholder interaction?", "a": "Yes, constantly. In both roles, I acted as the liaison, translating evolving corporate business definitions into exact automated technical views within agile squads."}
        ]
    },
    9: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Product Analytics Focus",
        "tag": "PRODUCT",
        "bridge": "I treat data tables as living products that require user-centric design, clear feature scoping, and continuous optimization.",
        "followup": "My resume explicitly highlights my technical skills in Product Analytics frameworks (*Pendo*, *Gainsight PX*, *WalkMe*), ensuring I know how to track user digital adoption and behavior logs.",
        "match": "Perfect alignment for Ambev's modern Data Products squads aiming to maximize internal or external platform engagement.",
        "growth": "Ensures that data solutions developed for Ambev act as intuitive tools that drive user adoption, not complex tech burdens.",
        "case": "Product Analytics Stack Mastery + Digital Channel tracking.",
        "bullets": [
            "I connect backend database performance metrics to front-end user behavioral events cleanly.",
            "I help product managers understand feature drop-offs and funnel blockages using hard analytics evidence.",
            "I design data product metrics that measure active usage, data reliability, and business impact."
        ],
        "qa_responses": [
            {"q": "What is the difference between a traditional report and a Data Product?", "a": "A traditional report is a static look at past data. A Data Product is an interactive, scalable, and documented asset engineered to solve a continuous operational problem, built with user adoption metrics and strict SLA tracking."},
            {"q": "How would you track user engagement inside an Ambev internal dashboard?", "a": "I would map out user journey sessions using product analytics tracking, analyzing which filters or sections have the highest drop-off rate to continuously refine the data architecture and UI clarity."},
            {"q": "How does your background in T&D/Management support Product Analytics?", "a": "At *Integral*, I conducted over 2,000 structured analyses focused on corporate learning and user engagement gaps for market leaders like Gerdau and Sabesp. I know exactly how to read human behavior logs and translate them into process improvements."}
        ]
    },
    10: {
        "category": "WHY - Intent & Fit",
        "title": "Why Business Analyst role?",
        "tag": "EVOLUTION",
        "bridge": "The best Business Analysts are those who do not depend on someone else to pull data from the database to answer a critical commercial question.",
        "followup": "My profile combines an MBA and an engineering background with deep technical execution in SQL and Python. I am not just going to write down requirements; I am going to validate if they are technically viable from day one.",
        "match": "Positions you as a self-sufficient, elite analyst who drives true speed inside agile data operations.",
        "growth": "Eliminates communication delays between commercial leaders and database engineering squads.",
        "case": "Production Engineering + MBA + Advanced Tech Arsenal.",
        "bullets": [
            "I bring an owner's mindset to requirement gathering, always questioning the core financial impact behind a technical request.",
            "My data background ensures that my documentation is mathematically precise and structurally aligned with cloud best practices.",
            "I am moving into this role because the future of business strategy belongs to those who build robust data products."
        ],
        "qa_responses": [
            {"q": "Are you more technical or more business-oriented?", "a": "I am a hybrid asset. My brain is structured by process engineering and corporate administration, but my hands write clean SQL, Python, and cloud view optimizations. I sit exactly in the middle."},
            {"q": "How do you handle a business stakeholder who demands a feature that will corrupt data quality?", "a": "I use data storytelling. I model the negative impact that poor schema data will have on their ultimate commercial KPIs, guiding them to an alternative requirement that protects both their targets and data governance standards."},
            {"q": "Why is your profile better than a candidate with a pure business background?", "a": "Because I eliminate tech-translation friction. A pure business candidate takes weeks to understand why an architecture cannot support a specific query speed; I can look at the BigQuery or Databricks views and map the solution instantly."}
        ]
    },
    11: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "ASICS (2026 - FinOps & Cloud)",
        "tag": "CLOUD-OPTIMIZATION",
        "bridge": "I re-engineered the multicountry financial data architecture for a retail leader to ensure a single, automated source of revenue truth.",
        "followup": "The project involved integrating fragmented transactional metrics across Brazil, Chile, and Colombia into an optimized cloud structure.",
        "match": "Proves high-level capability to handle complex financial logic and drive massive cloud performance gains.",
        "growth": "Directly mirrors the multinational, multi-channel distribution complexity that Ambev manages daily.",
        "case": "Stalse Project for *ASICS Latam*.",
        "bullets": [
            "Situation: Fragmented regional inputs created severe tracking discrepancies and massive data processing costs for management.",
            "Action: Redesigned 16+ tables and views in Google BigQuery using advanced SQL, implementing partitioning and incremental loads.",
            "Result: Slashed data consumption from Gigabytes/Terabytes down to Megabytes, implementing true FinOps and automating the entire dashboard."
        ],
        "qa_responses": [
            {"q": "How did you manage currency conversions across different Latin American regions?", "a": "I built automated SQL layer transformations that ingested regional currency parameters dynamically, translating distinct transactional local streams into a single standardized corporate view without manual input lag."},
            {"q": "What does reducing data consumption from GB to MB mean for business stakeholders?", "a": "It means speed and cost-efficiency. Dashboards that used to take minutes to load now refresh in seconds, enabling executive decisions to happen instantly while heavily reducing cloud platform infrastructure costs."},
            {"q": "How would you apply this ASICS logic at Ambev?", "a": "I will apply this exact structural optimization to Ambev’s commercial data product lines, ensuring that data transformations remain fast, scalable, and cost-effective as volume expands."}
        ]
    },
    12: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "NTT DATA / Itaú (2025)",
        "tag": "SCALE-DATA",
        "bridge": "I engineered optimized cloud data views within an AWS environment to support complex corporate indicator analysis for a tier-one bank.",
        "followup": "Operating inside an agile squad, the role demanded absolute data precision and the ability to turn complex business definitions into high-performing code.",
        "match": "Proves fluid execution inside elite tech consulting squads and massive database environments.",
        "growth": "Guarantees that your data logic can withstand the massive, high-velocity transactional scales of Ambev.",
        "case": "Data Analyst for *NTT Data* allocated at *Itaú*.",
        "bullets": [
            "Situation: The corporate sector needed to pull actionable metrics from tables containing billions of rows of complex records.",
            "Action: Executed optimized SQL queries and complex view transformations using Amazon Athena, S3, and AWS Glue frameworks.",
            "Result: Delivered automated analytical views and reports via QuickSight, ensuring absolute consistency under massive data scales."
        ],
        "qa_responses": [
            {"q": "How did you ensure data integrity when business rules changed rapidly during the squad sprint?", "a": "Instead of hardcoding rules into static scripts, I leveraged parameter-driven view logic. When indicators evolved, we updated the master configuration parameters, and the historical cloud structure adapted instantly without data downtime."},
            {"q": "What was your dynamic with the agile squad product owner?", "a": "Continuous alignment. I acted as the technical translator, helping the PO understand database schema boundaries while finding alternative SQL engineering paths to deliver the commercial metrics required for the sprint milestone."},
            {"q": "How does your AWS experience translate to Ambev’s data stack?", "a": "The cloud fundamentals are identical. Whether you run Athena on AWS or Spark on Azure/Databricks, the strategic focus on table optimization, indexing, and efficient partitioning remains the absolute benchmark for data success."}
        ]
    },
    14: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "Afinz (2022 - 2023)",
        "tag": "PROCESS-AUTOMATION",
        "bridge": "I applied process engineering logic to slash reporting manual lag by over 80%, transforming slow routines into governed data assets.",
        "followup": "By mapping metadata structures and building automated Python and SQL workflows, we eliminated operational vulnerabilities and gave hours back to the business squads.",
        "match": "Solid proof of an owner-mindset execution that targets waste and drives operational efficiency.",
        "growth": "Fits perfectly with Ambev's core culture of continuous optimization, lean operations, and process discipline.",
        "case": "MIS Analyst at *Afinz* / *Sorocred*.",
        "bullets": [
            "Situation: Daily update routines were entirely manual, taking 1.5 hours every morning and causing massive reporting delays.",
            "Action: Engineered automated ETL data pipelines utilizing Python, SQL, and AWS infrastructure (Glue, S3).",
            "Result: Reduced the entire daily process down to just 15 minutes, while building robust metadata repositories in Confluence."
        ],
        "qa_responses": [
            {"q": "Why was documenting metadata in Confluence a priority for this project?", "a": "Because automation without documentation creates technical debt. By mapping the metadata, I ensured that any business or technical stakeholder could audit the data lineage independently, securing data governance long-term."},
            {"q": "How did you handle data quality errors during the automated ingestion?", "a": "I coded automated warning alerts into the Python script. If an incoming file had structural anomalies or missing attributes, the system flagged the discrepancy immediately, preventing corrupted metrics from hitting executive dashboards."},
            {"q": "How will you apply this automation drive at Ambev?", "a": "I will systematically identify any manual spreadsheet workflows or repetitive extraction routines within your business analytics lines and transform them into lean, automated data products."}
        ]
    },
    13: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "Heineken (2023 - 2024)",
        "tag": "FMCG-DIGITAL",
        "bridge": "I unified and normalized chaotic e-commerce data streams across hundreds of distinct digital channels for a major beverage brand.",
        "followup": "Operating within the Digital Area (eRetail/eCommerce), my focus was resolving massive naming discrepancies to build a reliable one-page executive performance view.",
        "match": "Gives you a near-zero learning curve regarding the beverage industry, digital sales channels, and FMCG performance metrics.",
        "growth": "Directly matches Ambev’s strategic push into digital commercial ecosystems and advanced eRetail analytics tracking.",
        "case": "BI Analyst at Sxpel allocated at *Heineken* (Digital Area).",
        "bullets": [
            "Situation: A single product possessed more than 10,000 distinct denominations across fragmented client files, blinding campaign tracking.",
            "Action: Designed a rigid dimensional Star Schema data model and built a unified, automated performance pipeline inside Power BI.",
            "Result: Delivered the brand's first stable one-page performance dashboard, matching commercial indicators to the exact penny."
        ],
        "qa_responses": [
            {"q": "Looking back at your Heineken tenure, how do you see the scalability of digital distribution?", "a": "FMCG data products require a solid relational mapping layer. In a fast-moving market, you cannot have key accounts or regional distributors naming products under localized guidelines. Normalization ensures that commercial indicators scale cleanly as transaction logs increase."},
            {"q": "What specific FMCG metrics did you master during this tenure?", "a": "I worked closely with stock ruptures, promotional campaign conversions, client segments, eRetail channel growth, and multi-brand market share correlations."},
            {"q": "Why is this Heineken experience a massive advantage for Ambev?", "a": "Because I already speak your industry's language. I understand the tension between commercial campaigns and digital channel stock availability. I can step into Ambev's data environment and generate impact from week one without needing industry training."}
        ]
    },
    15: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "Burity (Risk & Compliance)",
        "tag": "RISK-MITIGATION",
        "bridge": "I spent years acting as a legal proxy, managing high-value corporate agreements and structural risk with zero liabilities.",
        "followup": "This experience built my clinical attention to detail and taught me how to analyze strict regulatory text, blueprints, and corporate contracts to safeguard assets.",
        "match": "Validates your deep foundational maturity, professional ethics, and mastery over corporate risk management.",
        "growth": "Ambev values professionals who guard company health and compliance; this tenure proves your long-term trustworthiness.",
        "case": "Asset & Property Manager / Legal Proxy at *Burity Empresarial*.",
        "bullets": [
            "Situation: Historical descriptive errors and contract ambiguities created multi-million dollar liabilities with government registries.",
            "Action: Audited complex legal frameworks, coordinated cross-functional squads, and executed data correction steps administratively.",
            "Result: Secured critical corporate infrastructure expansions smoothly, mitigating 100% of risk with zero lawsuits."
        ],
        "qa_responses": [
            {"q": "How does legal proxy experience translate into a Data Business Analyst role?", "a": "Compliance and analysis require the exact same mental discipline: reading strict rules, verifying alignment, and detecting gaps. Auditing land registries trained my eye to ensure that data flows strictly adhere to institutional governance frameworks."},
            {"q": "How do you handle negotiations with tough corporate stakeholders?", "a": "With absolute data preparation. My long experience managing institutional risks taught me that the best way to resolve conflict is to lay indisputable, well-documented evidence on the table, removing emotional biases completely."},
            {"q": "What did this long tenure teach you about business operations?", "a": "It taught me to think like an owner. Every process, whether it is a physical asset or a cloud database view, must be managed with absolute accountability, clear documentation, and a focus on long-term corporate security."}
        ]
    },
    16: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "Handling a Major Mistake",
        "tag": "CONTAINMENT",
        "bridge": "I operate with absolute accountability. If an ingestion pipeline breaks or a requirement gap creates an reporting bug, my immediate response is instant containment, not finger-pointing.",
        "followup": "I isolate the affected data range, notify leadership with transparent exposure numbers, and immediately deploy a permanent architectural patch.",
        "match": "Showcases C-level executive maturity, treating operational failures as objective engineering parameters to solve.",
        "growth": "In fast-paced tech environments, bugs happen; an owner-minded analyst contains the leak calmly and safeguards the data product.",
        "case": "Root-Cause Analysis & Append-Only Database Integrity.",
        "bullets": [
            "I suppress emotional noise, pull the raw logs, and isolate the exact parameters that triggered the system failure.",
            "I apply temporary data containment filters to preserve dashboard visibility while engineering the permanent fix.",
            "I translate that specific failure into a new automated Data Quality constraint rule so the bug can never repeat."
        ],
        "qa_responses": [
            {"q": "Can you describe a time you uncovered a critical reporting error?", "a": "In a previous routine, an unmapped source parameter shifted, distorting a regional KPI view. I instantly contained the visualization layer, tracked the schema drift via logs, corrected the SQL logic, and deployed an automated validation alert to proactively catch future parameter variations."},
            {"q": "How do you break bad data news to a corporate director?", "a": "With data precision and immediate solutions. I never present an open-ended issue. I present a concise summary explaining the root cause, the exact quantified impact on the current report, and the precise timeline for the resolution patch."},
            {"q": "How do you ensure data corrections do not mess up historical data auditability?", "a": "By enforcing strict data engineering best practices. Any correction is processed as a transparent, documented ledger adjustment with full metadata traceability, keeping the master tables clean and fully auditable."}
        ]
    },
    17: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "Unmapped High-Pressure Task",
        "tag": "PRESSURE-AGILE",
        "bridge": "Under extreme pressure or when facing completely unmapped business issues, I rely on structured agile frameworks, not emotional guessing.",
        "followup": "When sudden corporate shifts occur—like a new commercial channel launch or a sudden tax policy update—you isolate the variables, pull quick data samples, and roll out a calculated MVP solution.",
        "match": "Demonstrates an organized, analytical, and highly stable profile inside chaotic corporate environments.",
        "growth": "Ambev operates with velocity; my background ensures a steady, data-backed analytical filter during urgent strategic pivots.",
        "case": "Agile Problem Diagnosis & Pattern Analysis.",
        "bullets": [
            "Step 1: I isolate the core variable to understand which commercial KPI or data pipeline is directly threatened.",
            "Step 2: I extract an exploratory database sample via optimized SQL to support our choices with immediate evidence.",
            "Step 3: I roll out a lean MVP data product view and monitor its performance behavior metrics in real-time."
        ],
        "qa_responses": [
            {"q": "How do you start mapping a business requirement for a department you know nothing about?", "a": "I apply structured thinking. I treat their department as a system: I map their inputs, their manual operational friction points, and their ultimate target output. This logic allows me to understand their requirements fast, regardless of the sector."},
            {"q": "How do you maintain delivery speed when requirements shift mid-sprint?", "a": "By leveraging agile adaptability. I work closely with the squad stakeholders to re-prioritize the backlog, focusing purely on the core MVP features that preserve data quality and answer the immediate business crisis first."},
            {"q": "What happens if you lack the immediate data to make an analytical recommendation?", "a": "I do not guess. I rapidly query our data lakes to perform an accelerated pattern analysis, presenting a data-backed risk assessment to leadership so we can execute a calculated corporate decision."}
        ]
    },
    18: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "Conflict with Stakeholders",
        "tag": "STAKEHOLDERS",
        "bridge": "I do not fight subjective opinions with more opinions; I take emotional bias out of the room by laying hard data on the table.",
        "followup": "Stakeholders usually push back or challenge requirements because of underlying business anxieties regarding their targets. Once you show them transparent data models, the noise stops.",
        "match": "Validates elite communication, stakeholder management, and non-combative alignment techniques.",
        "growth": "Ensures smooth collaboration bridges between aggressive commercial targets and strict IT/engineering structures.",
        "case": "Data-Driven Consensus & Strategic Alignment.",
        "bullets": [
            "I start by practicing active listening to truly understand the stakeholder's operational or commercial fear.",
            "I present clear, comparative data performance options, mapping the exact impact of each requirement path.",
            "Once the transitional numbers and Data Quality realities are visual, teams naturally converge on the same logical execution."
        ],
        "qa_responses": [
            {"q": "How do you handle a sales manager who demands a dashboard metric that you know is mathematically incorrect?", "a": "I don't just refuse. I schedule a quick sync, pull up the raw database query lineage, and demonstrate visually how that specific formula logic will distort their eventual campaign calculation, showing them the correct path that secures their bonus accuracy."},
            {"q": "What is your strategy for managing demanding, high-level corporate stakeholders?", "a": "Complete data transparency and structured updates. I establish clear communication cadences, use concise data storytelling, and ensure that every developed data feature is backed by fully documented requirements."},
            {"q": "How do you foster collaboration between engineering squads and business units?", "a": "By acting as the bilingual asset. I leverage my MBA to talk strategy with the directors, and I use my engineering and SQL background to write clear technical stories for the developers, eliminating misunderstanding entirely."}
        ]
    },
    19: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "Tech to Non-Tech",
        "tag": "DATA-STORYTELLING",
        "bridge": "I translate complex cloud data engineering structures into clear commercial impact metrics and corporate cost optimization summaries.",
        "followup": "Vice presidents and directors do not need to hear about database joins, partition syntax, or Spark configurations; they need to know if the data product is reliable enough to unlock market share.",
        "match": "Perfect alignment for a Business Analyst role where high-level corporate reporting clarity is a mandatory daily rule.",
        "growth": "Empowers Ambev’s leadership with intuitive, digestible dashboards that support rapid strategic execution.",
        "case": "Executive Reporting Layers & One-Page Performance Dashboards (*Heineken* / *Looker*).",
        "bullets": [
            "I never showcase code logic to business units; I showcase operational hours saved or revenue leaks eliminated.",
            "I leverage standard corporate metrics that leadership cares about, such as ROI, CAC optimization, and processing velocity.",
            "I make data quality status visually obvious through intuitive indicator layers inside clean dashboards."
        ],
        "qa_responses": [
            {"q": "How do you design a dashboard meant for C-level presentation?", "a": "I follow a strict top-down structure: the top layer displays the ultimate macro KPIs; the middle layer breaks it down by channels or segments; and the bottom layer provides an automated anomaly log. It must tell a business story at a single glance."},
            {"q": "What do you do if an executive challenges the validity of your report?", "a": "I confidently walk them back to our documented data governance layer. I show them the automated data quality checkpoints and the lineage trace, proving that the metric is a verified reflection of our raw cloud architecture data."},
            {"q": "How did you manage executive communication at Heineken and Itaú?", "a": "By leading structured data alignment alignments. I didn't just hand over a dashboard link; I delivered clear executive analytical summaries, highlighting the main business trends and operational gaps revealed by the data."}
        ]
    },
    20: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "Cultural Fit, Travels & Closing",
        "tag": "GLOBAL-STORY",
        "bridge": "To wrap up, a global company like Ambev requires more than technical lines; it requires cultural adaptability, a keen eye for details, and a high capacity for learning from diverse environments.",
        "followup": "My advanced English and international perspective were built during nearly 6 months in England. I treat discovering new places with the same investigative mindset I use to audit database anomalies.",
        "match": "Demonstrates an elite profile with global mindset, rich life experiences, sports discipline, and local accountability in São Paulo.",
        "growth": "Perfect interpersonal alignment for senior stakeholder interaction, showcasing an open, disciplined, and highly curious professional.",
        "case": "International Adaptation (UK, France, Spain) + Multi-Sport Discipline + SP Local Execution.",
        "bullets": [
            "During my 6 months in London, I deep-dived into structural details—from exploring iconic spots like Baker Street and the British Museum to auditing history inside the library of the Royal Institution of Chartered Surveyors (RICS) and attending professional congresses at the Merchant House.",
            "I leverage dynamic solutions in daily life: buying a folding bicycle allowed me to explore every corner of Hyde Park and Trafalgar Square with agility, a practice of efficiency that maps directly to how I build plug-and-play code structures.",
            "My travel adaptability stretches across high-velocity shifts: navigating Paris during an intense, piercing winter after arriving on the high-speed bullet train, analyzing engineering details like rubber-tired subway networks, or exploring architecture and logistics like the massive moving walkways at Madrid Airport.",
            "I maintain strong personal discipline through sports like capoeira (which I've played in historical grounds like Salvador's Mercado Modelo) and swimming, combined with an active weekend routine walking along Avenida Paulista, exploring Parque Vila-Lobos e Ibirapuera com minha esposa (que é concursada e com quem divido a rotina em SP há 3 anos), além de viagens por destinos como a engenharia de Itaipu em Foz do Iguaçu, Belém do Pará, Búzios, Ilhabela e a Serra da Mantiqueira.",
            "Tenho um filho de 15 anos que está em plena fase de adolescência, compartilha do meu gosto por esportes e andar de bike, e está se preparando ativamente para um intercâmbio internacional—um passo natural em nossa família, dado que ele tem uma irmã mais velha por parte de mãe que é casada e reside nos EUA há mais de 10 anos."
        ],
        "qa_responses": [
            {"q": "How do your travel experiences and family background help you as a Data Business Analyst?", "a": "Traveling teaches you pattern recognition and adaptability. For instance, notice how I spotted the unique design of rubber-tired trains in Paris, or the engineering scale of the walkways in Madrid and the turbines at Itaipu. Furthermore, guiding my 15-year-old son through his international exchange preparation while tracking family roots in the US shows the strategic planning and long-term vision I apply to life and data products."},
            {"q": "We noticed you visited the RICS library in London. What drove you there?", "a": "I am deeply curious about institutional structures, standards, and regulatory frameworks. Stepping inside a century-old symbol of governance like the Merchant House or studying at RICS highlights my appreciation for established compliance, data quality, and documentation benchmarks."},
            {"q": "How do your sports background and personal routine relate to Ambev's owner mindset?", "a": "Esporte como a capoeira exige flexibilidade, respeito à tradição e reflexo rápido; a natação exige consistência e fôlego de longo prazo. Em São Paulo, onde moro há três anos com minha esposa, mantive essa rotina ativa explorando parques e teatros. Eu trago essa mesma energia e disciplina para os projetos da Ambev, encarando desafios complexos com resiliência e foco absoluto em entregar resultados consistentes para o time."}
        ]
    }
}

# Navigation State Initialization
if "view_mode" not in st.session_state:
    st.session_state.view_mode = "Main Interview Board"
if "active_id" not in st.session_state:
    st.session_state.active_id = 1

with st.sidebar:
    st.markdown("### Select Workspace View")
    # Native Streamlit navigation buttons inside the sidebar to toggle views seamlessly
    if st.button("📊 Main Interview Board", use_container_width=True):
        st.session_state.view_mode = "Main Interview Board"
        st.rerun()
        
    if st.button("📄 View: André Carvalho ENG_2.pdf", use_container_width=True):
        st.session_state.view_mode = "CV Doc"
        st.rerun()
        
    if st.button("📘 View: Ambev Target Requirements.pdf", use_container_width=True):
        st.session_state.view_mode = "Guide Doc"
        st.rerun()
    
    st.markdown("---")
    st.markdown("### Strategic Framework")
    st.info("**WHY:** Motivation & Fit\n\n**WHAT:** Scope & Profile\n\n**HOW:** STAR Actions\n\n**WHEN:** Crisis & Investigative Closing")
    
    st.markdown("### Match Analytics")
    st.metric(label="Ambev Adherence Score", value="99%", delta="Elite Stack Match")
    st.caption("**Target:** Ambev · Business/Data Analyst")

# --- VIEW ROUTING ENGINE ---

if st.session_state.view_mode == "CV Doc":
    st.markdown('<div class="doc-container"><div class="doc-title">Document View: André Carvalho — Resume</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="doc-section">
        <strong>ENG. ANDRÉ CARVALHO, MBA</strong><br>
        Analista de Dados Sênior | Data Analytics | Business Intelligence | SQL | Python
    </div>
    <div class="doc-subtitle">Professional Summary</div>
    <div class="doc-section">
        Analista de Dados e Business Intelligence com formação em Engenharia, MBA e inglês avançado. Ampla experiência em Analytics, Modelagem de Dados, Governança de Dados, FinOps e Cloud Analytics, entregando soluções de dados de alto impacto para grandes empresas como Heineken, Itaú, ASICS Latam, Gerdau, Sabesp, Fretebras, Afinz e outras. Experiência na construção de pipelines de dados, modelagem analítica, automação de processos, consolidação de KPIs e desenvolvimento de dashboards executivos. Foco na transformação de dados em insights acionáveis que suportam tomadas de decisão de alto nível, otimização operacional e redução de custos. Traz vivência internacional (Reino Unido, França e Espanha) e perspectiva global para equipes multidisciplinares.
    </div>
    <div class="doc-subtitle">Professional Experience</div>
    <div class="doc-section">
        <strong>Analista de Business Intelligence :: Stalse | Jan 2026 - Apr 2026</strong><br>
        • Atuação em projetos estratégicos para ASICS Latam, Fretebras e outras empresas com foco em Analytics, Revenue e Business Intelligence.<br>
        • Redesenho de mais de 16 tabelas e views para uma arquitetura escalável e otimizada, reduzindo o volume de processamento de dados de escala de GB/TB para MB por meio de particionamento, cargas incrementais e boas práticas de modelagem (FinOps).<br>
        • Criação de visão de receita unificada para operações Latam (Brasil, Chile e Colômbia) com conversão cambial integrada.
    </div>
    <div class="doc-section">
        <strong>Data Analyst :: NTT DATA / Itaú | Feb 2025 - May 2025</strong><br>
        • Desenvolvimento e manutenção de views SQL complexas utilizando Amazon Athena em ambiente cloud da AWS para o Itaú.<br>
        • Participação ativa em squads ágeis para entrega contínua de soluções orientadas a dados e tradução de regras de negócio em visões analíticas complexas.
    </div>
    <div class="doc-section">
        <strong>Analista de Business Intelligence :: Sxpel / Heineken | Nov 2023 - Out 2024</strong><br>
        • Alocado na área digital da Heineken (eCommerce, eRetail), consolidando e modelando múltiplas bases de dados complexas de diversos clientes.<br>
        • Estruturação de modelo dimensional (Star Schema) para unificar cenários em que o mesmo produto possuía mais de 10 mil denominações distintas entre clientes e sistemas, garantindo Data Quality e rastreabilidade.
    </div>
    <div class="doc-section">
        <strong>MIS Analyst :: Afinz | Sep 2022 - Jun 2023</strong><br>
        • Otimização de rotinas de atualização de relatórios, reduzindo o tempo de processamento de 1h30 para apenas 15 minutos via automação ETL com Python e SQL.<br>
        • Estruturação de repositórios de metadados e implementação de práticas sólidas de Data Governance e Data Quality (Confluence).
    </div>
    <div class="doc-subtitle">Education & Certifications</div>
    <div class="doc-section">
        • <strong>MBA em Administração de Empresas</strong> — Fundação Getulio Vargas (FGV), 2006.<br>
        • <strong>Bacharelado em Engenharia de Produção</strong> — CREA-SP, 2021.<br>
        • <strong>Idiomas:</strong> Português (Nativo) | Inglês Avançado (Formação internacional pela St Giles International, Londres).
    </div>
    """, unsafe_allow_html=True)

elif st.session_state.view_mode == "Guide Doc":
    st.markdown('<div class="doc-container"><div class="doc-title">Document View: Ambev Alignment Matrix</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="doc-section">
        <strong>Análise Estratégica Ambev: Business/Data Analyst</strong>
    </div>
    <div class="doc-subtitle">Conexão do Perfil à Vaga</div>
    <div class="doc-section">
        • <strong>A Vaga:</strong> Busca um profissional híbrido (Liaison), capaz de fazer o levantamento de requisitos com as áreas de negócio e traduzir esses cenários para os times técnicos de Engenharia de Dados. Foco em Data Products, Data Quality e metodologias ágeis em ambiente global (inglês obrigatório).<br>
        • <strong>Os Diferenciais da Vaga:</strong> Domínio de Python, conceitos de Data Governance, e ferramentas de Big Data como <strong>Databricks, Azure Data Factory ou Spark</strong>.<br>
        • <strong>Gatilhos do Seu Currículo para a Ambev:</strong><br>
        &nbsp;&nbsp;&nbsp;&nbsp;- <em>Heineken (FMCG/Bebidas):</em> Você já domina a dinâmica de eRetail, canais de distribuição e reconciliação de dados de produtos do setor de bebidas. Viés imediato para entender iniciativas como o BEES.<br>
        &nbsp;&nbsp;&nbsp;&nbsp;- <em>Itaú (Squads Ágeis & Escala):</em> Prova sua facilidade em trabalhar em squads multidisciplinares utilizando metodologias ágeis sob rígidos critérios técnicos e altíssimo volume de logs (bilhões de linhas).<br>
        &nbsp;&nbsp;&nbsp;&nbsp;- <em>ASICS (Cloud FinOps):</em> Mostra sua mentalidade de eficiência de custos (cultura de dono da Ambev), reduzindo processamento de GBs para MBs.<br>
        &nbsp;&nbsp;&nbsp;&nbsp;- <em>Afinz (Data Governance & Quality):</em> Entrega exata do diferencial solicitado: documentação de requisitos em Confluence, governança de metadados e automação de ETL de 1h30 para 15 min.
    </div>
    <div class="commentary-box">
        <strong>Dica Executiva C-Level:</strong> Nas respostas em inglês, evite focar apenas na sintaxe do código. Enfatize como você estrutura as definições de negócio para que o produto de dados (Data Product) nasça com qualidade e atenda perfeitamente o ROI comercial esperado pela liderança.
    </div>
    """, unsafe_allow_html=True)

else:
    # --- DEFAULT MAIN INTERVIEW BOARD VIEW ---
    categories_list = [
        "WHAT - Capabilities & Profile", 
        "WHY - Intent & Fit", 
        "HOW - Case Methodology (STAR)", 
        "WHEN - Extreme Scenarios & Crisis"
    ]

    # Symmetrical 4-Column Layout Grid for Category Navigation
    cols = st.columns(len(categories_list))

    for idx, cat_name in enumerate(categories_list):
        with cols[idx]:
            st.markdown(f'<div class="category-header">{cat_name.split(" - ")[0]}</div>', unsafe_allow_html=True)
            cat_items = {k: v for k, v in DATA_MAPPING.items() if v["category"] == cat_name}
            
            for item_id, item_data in cat_items.items():
                is_active = (st.session_state.active_id == item_id)
                tag_token = f"[{item_data.get('tag', 'CONTEXT')}] "
                clean_title = item_data.get('title', 'Untitled')
                btn_label = f"▸ {tag_token}{clean_title}" if is_active else f"{tag_token}{clean_title}"
                
                if st.button(btn_label, key=f"btn_{item_id}"):
                    st.session_state.active_id = item_id
                    st.rerun()

    st.markdown("<div style='margin-top: 0.15rem; border-top: 1px solid #e9ecef; margin-bottom: 0.25rem;'></div>", unsafe_allow_html=True)

    active_data = DATA_MAPPING.get(st.session_state.active_id, DATA_MAPPING[1])

    # Responsive 50-50 Split View below navigation matrix
    col_out1, col_out2 = st.columns([0.50, 0.50])

    with col_out1:
        st.markdown(
            f"""
            <div class="response-box">
                <span style="color:#117a65; font-size:9.5px; font-weight:bold; text-transform:uppercase;">The Golden Bridge (Natural phrasing):</span><br>
                <strong style="font-size:12.5px; color:#2c3e50; line-height:1.2;">"{active_data.get('bridge', '')}"</strong>
            </div>
            """, 
            unsafe_allow_html=True
        )
        
        st.markdown(
            f"""
            <div class="followup-box">
                <span style="color:#2c3e50; font-size:9.5px; font-weight:bold; text-transform:uppercase;">Deep Dive / Context:</span><br>
                <p style="font-size:12px; color:#34495e; line-height:1.25;">{active_data.get('followup', '')}</p>
            </div>
            """, 
            unsafe_allow_html=True
        )
        
        st.markdown(
            f"""
            <div class="growth-box">
                <strong style="color:#d35400; text-transform:uppercase; font-size:9px;">The Ambev Alignment Link (The Strategic Approach):</strong><br>
                <p style="color:#ba4a00; font-size:11.5px; line-height:1.25; margin-top:1px;">{active_data.get('growth', '')}</p>
            </div>
            """, 
            unsafe_allow_html=True
        )
        
        st.markdown(
            f"""
            <div class="match-box">
                <strong style="color:#2980b9; text-transform:uppercase; font-size:9px;">The Match Concept / Objective:</strong><br>
                <p style="color:#1f618d; font-size:11.5px; line-height:1.25; margin-top:1px;">{active_data.get('match', '')}</p>
            </div>
            """, 
            unsafe_allow_html=True
        )

    with col_out2:
        # ULTRA-COMPACT CONTAINER: Reduced margins, font size and inline layout for the reference text
        bullets_list = active_data.get("bullets", [])
        bullets_html = "".join(f'<p style="font-size:12px; color:#2c3e50; line-height:1.25; margin-bottom:2px !important;">• {b}</p>' for b in bullets_list)
        
        st.markdown(
            f"""
            <div class="bullet-container-box">
                <span style="color:#2c3e50; font-size:10.5px; font-weight:bold; text-transform:uppercase; display:block; margin-bottom:4px;">Supporting Core Arguments:</span>
                {bullets_html}
                <p style='font-size:10px; color:#7f8c8d; margin-top:2px !important;'><strong>Baseline Case:</strong> {active_data.get('case', '')}</p>
            </div>
            """,
            unsafe_allow_html=True
        )
        
        # REFACTORED UX BOX: Dark Blue Accent & Symmetrical layout spacing with fallback control
        qa_list = active_data.get("qa_responses", [])
        
        qa_html_items = ""
        for qa in qa_list:
            qa_html_items += f"""
            <div class="qa-item">
                <strong style="font-size:11.5px; color:#1b4f72; display:block; line-height:1.2;">Q: {qa.get('q', '')}</strong>
                <p style="font-size:11.5px; color:#154360; line-height:1.25; margin-top:1px !important;"><strong>A:</strong> {qa.get('a', '')}</p>
            </div>
            """
            
        st.markdown(
            f"""
            <div class="qa-container-box">
                <span style="color:#154360; font-size:10.5px; font-weight:bold; text-transform:uppercase; display:block; margin-bottom:4px;">⚡ TOUGHEST C-LEVEL Q&A SIMULATOR:</span>
                {qa_html_items}
            </div>
            """,
            unsafe_allow_html=True
        )
