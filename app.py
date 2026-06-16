import streamlit as st

st.set_page_config(
    page_title="War Room - micro1 AI Data Science",
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

# 20 Strategic Framework Database Items - Fully Tailored for micro1 AI Data Scientist Role
DATA_MAPPING = {
    1: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Tell me about yourself",
        "tag": "PROFILE",
        "bridge": "Senior Data Scientist with a background in Production Engineering, an MBA in Business Administration, and extensive expertise in Analytics, Business Intelligence, Data Governance, Cloud Analytics, and FinOps.",
        "followup": "Proven track record in strategic data projects for market-leading corporations like Heineken, Itaú, ASICS Latam, Gerdau, Sabesp, Fretebras, Afinz, and others, supporting high-level decision-making by transforming raw data into actionable insights and business indicators.",
        "match": "Expertise in data modeling, developing ETL/ELT pipelines, process automation, KPI consolidation, executive dashboards, and cloud analytics environment optimization.",
        "growth": "Fully prepared to apply my engineering and mathematical rigor to train and evaluate next-generation LLMs and AI agents inside the micro1 ecosystem.",
        "case": "Production Engineering + MBA + Advanced Python & SQL Modeling (*Heineken*, *Itaú*, *ASICS*).",
        "bullets": [
            "Automation of analytical processes and report update workflows, reducing operational effort by over 80%. Slashed dashboard refresh times from 1h30m to just 15 minutes by implementing structured data pipelines, automations, and process standardization.",
            "Consolidation and modeling of multiple complex databases for major retail and consumer goods operations, including scenarios where a single product had over 10,000 distinct denominations across different clients, channels, and systems. Structured analytical models capable of unifying and correlating dispersed data to generate reliable indicators and support strategic decision-making.",
            "Re-engineering of cloud analytics architectures, redesigning multiple tables and views into a scalable, optimized structure. Implemented partitioning, incremental processing, and FinOps best practices, slashing data processing consumption from GB/TB scale down to MB.",
            "Development of executive dashboards and Analytics solutions to monitor commercial, financial, operational, marketing, eCommerce, and eRetail indicators, including LATAM operations (Brazil, Chile, and Colombia). Applied Machine Learning and Advanced Analytics for revenue forecasting and business pattern identification utilizing models such as ARIMA+ and K-Means."
        ],
        "qa_responses": [
            {"q": "How does your engineering background enhance your data science execution?", "a": "Production Engineering focuses on systemic efficiency and optimization. Combined with an MBA, I don't just build statistical models; I design robust data validation, cleaning, and transformation processes that ensure real-world parameters map flawlessly into analytical frameworks."},
            {"q": "What is your experience with statistical and predictive modeling?", "a": "I routinely deploy advanced statistical analysis, time-series forecasting (such as ARIMA+), and clustering algorithms (like K-Means) to isolate business patterns, predict revenues, and segment massive multi-country datasets."},
            {"q": "How do you ensure data quality across highly fragmented pipelines?", "a": "At *Heineken*, I normalized datasets containing over 10,000 conflicting naming variations into a single cohesive Star Schema. I apply strict Data Quality and governance checks at every stage—ingestion, cleaning, and modeling—ensuring the outputs are 100% reliable."}
        ]
    },
    2: {
        "category": "WHY - Intent & Fit",
        "title": "Why micro1?",
        "tag": "MICRO1-FIT",
        "bridge": "micro1 is building the essential human intelligence layer for frontier AI, and I want to leverage my advanced analytical toolkit to structure the high-quality data training loops that shape how foundation models reason.",
        "followup": "Your mission to enable 1 billion people to do meaningful work by applying their expertise to AI resonates with my background. Having optimized workflows across banking, retail, and FMCG, I understand how to turn complex, domain-specific logic into clean, structured data inputs.",
        "match": "Highlights a senior professional who understands that the primary bottleneck in frontier AI is no longer just compute, but the absolute meticulous quality of training and evaluation data.",
        "growth": "Aligns your corporate scaling experience with micro1’s massive global expert network expansion.",
        "case": "Data Quality Scaling (*Heineken*) + Cloud Pipeline Automation (*Afinz* / *ASICS*).",
        "bullets": [
            "Frontier models require gold-standard data inputs to prevent hallucinations and ensure logical reasoning.",
            "My experience in highly regulated and complex fields (*Itaú*, *Heineken*) allows me to act as an elite domain expert and data evaluator.",
            "I am highly driven by micro1's remote, global-first laboratory model where data engineering meets AI alignment."
        ],
        "qa_responses": [
            {"q": "Why apply to an AI training lab instead of a traditional corporate data role?", "a": "Traditional corporate roles focus on visualizing the past. micro1 focuses on building the future of intelligence. Applying my data science and engineering rigor to benchmark, clean, and evaluate datasets for LLMs is the highest-leverage application of my skills."},
            {"q": "How do you view micro1's role in the market?", "a": "As models scale, raw internet scraping is exhausted. The winner of the AI race will be the company that curates the best specialized human feedback and agent evaluation pipelines. micro1 is exactly that layer."},
            {"q": "Are you comfortable working as a remote contractor in a global ecosystem?", "a": "Completely. I have advanced English skills, international academic experience in London, and a personal setup optimized for agile, cross-functional delivery across global timezones."}
        ]
    },
    3: {
        "category": "WHY - Intent & Fit",
        "title": "The Data Quality Obsession",
        "tag": "AI-QUALITY",
        "bridge": "In data science and AI training, garbage in is garbage out. My methodology is explicitly engineered to enforce meticulous attention to detail and zero data corruption.",
        "followup": "Whether preparing features for a machine learning model or structuring feedback datasets for an AI agent, data readiness requires rigorous pre-processing, schema validation, and outlier isolation.",
        "match": "Directly hits micro1's core need for specialists who can guarantee data integrity for frontier AI training.",
        "growth": "Positions you as an analyst who avoids shortcuts, ensuring that every token or data row meets high-quality benchmarks.",
        "case": "Meticulous Pre-processing and Metadata Repositories (*Afinz* / *Burity*).",
        "bullets": [
            "I am accustomed to auditing complex, unstructured records to rectify anomalies before they hit production environments.",
            "At *Afinz*, I built metadata frameworks that secured data consistency while speeding up processing lines by 80%.",
            "I treat data quality as an algorithmic discipline, setting up automated validation gates to catch drift instantly."
        ],
        "qa_responses": [
            {"q": "What does 'meticulous attention to detail' mean to you in a data context?", "a": "It means validating assumptions at the lowest granularity. It means checking null behavior, analyzing feature distributions, ensuring balance in datasets, and documenting every step so that the pipeline is completely auditable and reproducible."},
            {"q": "How would you evaluate if an expert's input dataset is ready for AI model training?", "a": "I would check for three core dimensions: statistical consistency, lack of demographic or formatting bias, and strict adherence to the target instruction schema. Any deviation must be isolated at the ingestion stage."},
            {"q": "How do you handle messy data when under a tight deadline?", "a": "You automate the cleaning logic but never automate the validation. I write Python scripts to handle the heavy pre-processing, but I always run rigorous descriptive audits on the output sample to guarantee its structural health."}
        ]
    },
    4: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Your Value Proposition",
        "tag": "VALUE-PROP",
        "bridge": "I bring a rigorous mix of process engineering logic, robust Python/SQL model building, and proven multi-industry domain knowledge to micro1.",
        "followup": "I bridge the gap between high-level conceptual questions and hard cloud data engineering pipelines, allowing micro1 to scale its evaluation and data pipelines cleanly.",
        "match": "Showcases a self-sufficient Data Scientist who can handle end-to-end projects from ideation to delivery.",
        "growth": "Slashes technical debt and manual pipeline bottlenecks, accelerating the feedback cycle for AI model training.",
        "case": "End-to-End Data Science Lifecycle (*ASICS* / *Itaú* / *Heineken*).",
        "bullets": [
            "I write optimized, clean Python and SQL code that scales smoothly across AWS and GCP cloud infrastructures.",
            "I translate chaotic, disparate client formats into rigid, standardized data assets.",
            "I communicate complex statistical findings with absolute clarity to both technical and C-level stakeholders."
        ],
        "qa_responses": [
            {"q": "What is your immediate 30-day value add?", "a": "I will audit your existing data curation or agent evaluation pipelines, identify manual extraction lags, and implement automated Python/SQL scripts to optimize processing times and data quality verification gates."},
            {"q": "How do you communicate data science models to non-technical teams?", "a": "By using clear data storytelling. I never showcase raw code lines or deep mathematical proofs to business units. I present the model's reliability boundaries, its business impact, and its visual trends through intuitive dashboards."},
            {"q": "What distinguishes your approach to programming?", "a": "I design code for production. I prioritize modularity, efficient indexing, proper partitioning, and thorough documentation (using GitHub and repositories) to ensure that any developer can scale my scripts."}
        ]
    },
    6: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Data Manipulation Stack",
        "tag": "TECH-STACK",
        "bridge": "I treat Python and SQL as native languages to manipulate, clean, and model complex, high-volume datasets.",
        "followup": "My cloud experience spans AWS Athena, Glue, and S3 (*Itaú*) and Google BigQuery (*ASICS*), giving me complete fluidity to execute exploratory data analysis across distributed environments.",
        "match": "Directly satisfies the technical requirement for programming, data cleaning, and data visualization tools.",
        "growth": "Ensures that you can jump into micro1’s platform and interface with any cloud structure with zero onboarding lag.",
        "case": "High-Volume Query Optimization + Streamlit Application Building.",
        "bullets": [
            "I leverage libraries like Pandas, NumPy, and Scikit-Learn to build, clean, and validate predictive structures.",
            "I am highly comfortable designing interactive visual environments using Streamlit to prototype analytical data views fast.",
            "I apply FinOps to data engineering, refactoring views to slash resource consumption from Terabytes to Megabytes."
        ],
        "qa_responses": [
            {"q": "Are you comfortable working outside a standard SQL environment?", "a": "Yes. My Python foundation allows me to manipulate unstructured data, parse JSON log payloads from user-agent interactions, and handle diverse API outputs natively."},
            {"q": "How do you approach visualizing multidimensional data?", "a": "I start with clear dimensional reduction or aggregation depending on the goal, then utilize high-performance visual tools like Power BI, Looker, or Python visualization libraries to deliver clean, highly scannable insights."},
            {"q": "How fast can you adapt to a custom proprietary data tool?", "a": "Instantly. Because I master the underlying data primitives—relational algebra, script automation, and schema configurations—adapting to a new interface is just a matter of hours."}
        ]
    },
    11: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "ASICS (2026 - FinOps & Cloud)",
        "tag": "CLOUD-MODELING",
        "bridge": "I re-engineered the multi-country financial data model for a global retail leader to ensure a single, optimized source of analytical truth.",
        "followup": "The project demanded consolidating disparate transactional records across Brazil, Chile, and Colombia into Google BigQuery.",
        "match": "Proves high-level capability to handle complex multi-source data streams and implement extreme cloud performance optimizations.",
        "growth": "Demonstrates the technical discipline needed to build micro1's massive global expert data tracking pipelines.",
        "case": "Stalse Project for *ASICS Latam* (GCP Environment).",
        "bullets": [
            "Situation: Fragmented regional data layouts created severe tracking anomalies and massive, unsustainable query processing costs.",
            "Action: Redesigned 16+ core tables and views in BigQuery using advanced SQL, partitioning strategies, and incremental data loads.",
            "Result: Slashed processing data volume from GB/TB scales down to MB, heavily optimizing cloud infrastructure costs and dashboard speed."
        ],
        "qa_responses": [
            {"q": "How did you handle the cross-border currency and tax variance in this model?", "a": "I engineered dynamic SQL lookup layers that automatically processed regional currency variations and standardized tax parameters into a single corporate reporting layer, eliminating manual conversion errors entirely."},
            {"q": "What was the direct benefit of optimization from Terabytes to Megabytes?", "a": "It transformed the data lifecycle. Query executions that used to cause system lag and high data costs became near-instantaneous and cost-efficient, allowing stakeholders to run real-time exploratory analysis safely."},
            {"q": "How does this map to micro1’s data requirements?", "a": "It proves I can handle high-volume, multi-source ingestion challenges and structure data architectures that remain lean, fast, and cost-effective as scaling expands."}
        ]
    },
    12: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "NTT DATA / Itaú (2025)",
        "tag": "MASSIVE-SCALE",
        "bridge": "I engineered optimized cloud data views within an AWS environment to handle complex statistical indicators for a tier-one bank.",
        "followup": "Operating inside an agile squad, the role required absolute consistency while querying high-volume tables with billions of rows of messy records.",
        "match": "Proves technical capability to handle massive datasets without compromising data integrity or system performance.",
        "growth": "Guarantees that your data manipulation scripts can handle the vast data flows generated by micro1's global AI recruitment platform.",
        "case": "Data Analyst at *NTT Data* for *Itaú* (AWS Cloud Environment).",
        "bullets": [
            "Situation: The corporate unit needed to calculate executive performance indicators by joining giant tables with billions of unstructured records.",
            "Action: Developed optimized SQL views and structured views using Amazon Athena, Amazon S3, and AWS Glue.",
            "Result: Delivered automated analytical views with absolute data consistency, adapting smoothly to evolving business rules."
        ],
        "qa_responses": [
            {"q": "How did you maintain data integrity when handling tables with billions of rows?", "a": "By leveraging partition strategies in Amazon S3, building optimized query schemas in Amazon Athena, and avoiding heavy, non-indexed full table scans. This kept our statistical runs stable and free from corruption."},
            {"q": "How did you handle shifting indicators mid-project?", "a": "I built parameter-driven view logic instead of hardcoding business constraints. When rules evolved, we simply updated the configuration metadata tables, and the entire historical database adapted instantly without downtime."},
            {"q": "Why is your AWS cloud experience valuable for micro1?", "a": "Because modern AI development happens on the cloud. Knowing how to efficiently read, partition, and transform raw logs stored in object storage (like S3) means I can easily build and optimize micro1's data readiness layers."}
        ]
    },
    14: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "Afinz (2022 - 2023)",
        "tag": "AUTOMATION-ETL",
        "bridge": "I applied process engineering and Python automation to eliminate over 80% of manual reporting lag, strengthening our data governance layer.",
        "followup": "By auditing manual, brittle routines and rewriting them into structured pipelines, we turned human-dependent tasks into governed data assets.",
        "match": "Solid proof of capability to enhance analytical methodologies and optimize end-to-end data workflows.",
        "growth": "Directly supports micro1's goal of processing high-quality expert contributions at a massive scale without manual bottlenecks.",
        "case": "MIS Analyst at *Afinz* / *Sorocred*.",
        "bullets": [
            "Situation: Daily data aggregation and verification workflows were entirely manual, taking 1.5 hours every morning and creating operational lag.",
            "Action: Developed automated ETL data pipelines utilizing Python, SQL scripts, and AWS infrastructure (Glue, S3).",
            "Result: Drastically reduced processing time down to just 15 minutes, while establishing rigorous metadata organization in Confluence."
        ],
        "qa_responses": [
            {"q": "Why was documenting metadata in Confluence just as important as writing the Python code?", "a": "Because an undocumented pipeline is a black box. Mapping the metadata ensured full data governance, allowing any team member to audit the data lineage and verify metric definitions independently."},
            {"q": "How did you approach data cleaning during this automation?", "a": "I programmed automated validation checks into the script. If incoming source payloads had missing variables or schema drift, the system isolated the record and flagged an alert, securing data quality before the analytics layer."},
            {"q": "How will you apply this automation focus at micro1?", "a": "I will systematically identify any repetitive manual validation or file formatting routines in your AI training data loops and replace them with automated Python workflows, giving hours of analysis time back to the squad."}
        ]
    },
    13: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "Heineken (2023 - 2024)",
        "tag": "DATA-NORMALIZATION",
        "bridge": "I engineered a dimensional normalization model to unify and clean chaotic e-commerce data streams across hundreds of distinct clients for a top brand.",
        "followup": "Operating within the Digital Area, the main challenge was mapping completely fragmented inputs into a transparent, fully auditable performance pipeline.",
        "match": "Demonstrates elite skills in data cleaning, handling complex taxonomy problems, and building unified single sources of truth.",
        "growth": "Perfect capability to normalize highly diverse, unstructured human text and evaluation outputs into uniform training matrices for LLMs.",
        "case": "BI Analyst at Sxpel allocated at *Heineken* (Digital Channel Analytics).",
        "bullets": [
            "Situation: A single product line possessed over 10,000 distinct naming conventions across different customer sheets, blinding campaign tracking.",
            "Action: Designed a rigid Star Schema dimensional model and built a unified, automated performance pipeline inside Power BI.",
            "Result: Successfully resolved the data friction, matching complex commercial indicators across channels to the exact penny."
        ],
        "qa_responses": [
            {"q": "How did you statistically align 10,000 text variations to a single master list?", "a": "By mapping out a clear relational mapping index and deploying Python-based text cleaning routines to strip noise, forcing all incoming disparate source data into a standardized relational architecture."},
            {"q": "What does 'matching indicators to the exact penny' mean for data quality?", "a": "It means achieving absolute auditability. If a data pipeline allows micro-discrepancies to slip through, it corrupts the downstream model's trust. I build verification loops that ensure 100% data fidelity."},
            {"q": "How does this experience help you prepare data for frontier AI models?", "a": "AI models learn from consistency. The exact same discipline required to align chaotic commercial product names is what is needed to clean, parse, and structure multi-domain human expert feedback into high-quality training tokens."}
        ]
    },
    15: {
        "category": "HOW - Case Methodology (STAR)",
        "title": "Burity (Risk, Logic & Standards)",
        "tag": "LOGICAL-RIGOR",
        "bridge": "I spent years acting as a legal proxy, managing high-value regulatory processes, corporate agreements, and asset risks with zero liabilities.",
        "followup": "This extensive tenure developed my analytical attention to detail and trained my eye to digest, interpret, and audit complex legal text, structural blueprints, and institutional rules.",
        "match": "Validates extreme precision, logical reasoning, and a deep understanding of corporate compliance and normative logic.",
        "growth": "Provides micro1 with the sophisticated reasoning capability needed to evaluate AI models in advanced domains like law, governance, and business STEM.",
        "case": "Asset & Property Manager / Legal Proxy at *Burity Empresarial*.",
        "bullets": [
            "Situation: Historical descriptive errors and structural ambiguities in legal records created multi-million dollar liabilities with registries.",
            "Action: Audited complex regulatory frameworks, aligned cross-functional teams, and executed administrative data corrections.",
            "Result: Secured critical corporate infrastructure expansions smoothly, mitigating 100% of institutional risk with zero lawsuits."
        ],
        "qa_responses": [
            {"q": "How does your legal proxy background translate to advanced AI Data Science?", "a": "Advanced AI alignment (like training a model to reason about compliance or finance) requires deep textual precision and strict adherence to logic rules. Auditing contracts and blueprints trained my brain to identify contradictions, logical gaps, and semantic flaws with expert-level accuracy."},
            {"q": "How do you handle complex interactions with strict regulatory frameworks?", "a": "With absolute data preparation and zero speculation. Regulators and frontier systems both respond to indisputable, fully documented evidence. I apply that exact same precision to evaluating data streams and model outputs."},
            {"q": "What did this long tenure teach you about data governance?", "a": "It taught me to think like an owner. Every system boundary and data schema must be treated with total accountability. A data asset is only valuable if its governance guarantees its accuracy and compliance over time."}
        ]
    },
    16: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "Handling a Major Data Leak/Error",
        "tag": "CONTAINMENT",
        "bridge": "I operate with absolute accountability. If a pipeline rule breaks or a data corruption bug slips through, my immediate priority is instant containment to minimize downside risk.",
        "followup": "I suppress emotional noise, isolate the affected data range via query logs, notify leadership with transparent exposure analytics, and deploy a permanent architectural patch.",
        "match": "Highlights corporate maturity under pressure, treating system errors as objective optimization parameters.",
        "growth": "In a high-velocity startup environment like micro1, bugs can appear; a senior data scientist contains the leak calmly and secures the training layers.",
        "case": "Root-Cause Analysis, Data Quality Filters & Append-Only Log Security.",
        "bullets": [
            "I pull the raw query logs, isolate the precise timestamp window, and apply temporary containment filters to protect active layers.",
            "I run a thorough root-cause analysis to find out why the ingestion rule or validation constraint failed to catch the drift.",
            "I write a new automated validation check directly into the script syntax so that specific vulnerability can never repeat."
        ],
        "qa_responses": [
            {"q": "Can you describe a time a data anomaly disrupted an operation?", "a": "In a previous pipeline, an unmapped source attribute changed its data type, causing an aggregation view to distort. I immediately isolated the visualization layer, parsed the backend logs to find the type mismatch, rewrote the SQL validation cast, and added an automated schema check to proactively block future datatype shifts."},
            {"q": "How do you communicate a critical pipeline error to non-technical stakeholders?", "a": "With complete data precision and immediate solutions. I never present an open-ended problem. I deliver a concise summary explaining the root cause, the exact quantified impact on the active dataset, and the specific resolution path."},
            {"q": "How do you ensure data corrections do not corrupt historical training logs?", "a": "By enforcing strict append-only database frameworks. Any corrective dataset adjustment is processed as a separate, documented ledger entry with clear metadata traceability, keeping the master historical layers fully auditable."}
        ]
    },
    20: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "Cultural Fit, Travels & Closing",
        "tag": "GLOBAL-MINDSET",
        "bridge": "To wrap up, a global, pioneering AI laboratory like micro1 requires more than just technical lines of code; it demands cultural adaptability, a keen eye for real-world details, and an unyielding curiosity.",
        "followup": "My advanced English skills and international perspective were built during nearly 6 months in London. I treat exploring new global infrastructures with the exact same investigative mindset I use to analyze data pipeline anomalies.",
        "match": "Combines a highly stable family structure in São Paulo with an elite global mindset, multi-sport discipline, and deep cultural curiosity.",
        "growth": "Perfect interpersonal and values alignment for a remote-first, high-velocity global contractor setup.",
        "case": "International Adaptation (UK, France, Spain) + Multi-Sport Discipline + SP Local Execution.",
        "bullets": [
            "During my 6 months in London (with academic experience at St Giles International), I deep-dived into structural details—from exploring Baker Street and the British Museum to auditing history inside the library of the Royal Institution of Chartered Surveyors (RICS) and attending professional congresses at the Merchant House.",
            "I leverage dynamic solutions in daily life: buying a folding bicycle allowed me to explore every corner of Hyde Park and Trafalgar Square with agility, a practice of efficiency that maps directly to how I build plug-and-play code structures.",
            "My travel adaptability stretches across high-velocity shifts: navigating Paris during an intense, piercing winter after arriving on the high-speed bullet train, analyzing engineering details like rubber-tired subway networks, or exploring architecture and logistics like the massive moving walkways at Madrid Airport.",
            "I maintain strong personal discipline through sports like capoeira (which I've played in historical grounds like Salvador's Mercado Modelo) and swimming, combined with an active weekend routine walking along Avenida Paulista, exploring Parque Vila-Lobos and Ibirapuera with my wife (who is a public servant and with whom I have shared a stable routine in SP for the past 3 years), alongside travels exploring the engineering marvel of Itaipu in Foz do Iguaçu, Belém do Pará, Búzios, Ilhabela, and the Mantiqueira mountains.",
            "I have a 15-year-old son who is in the midst of his teenage years; he shares my passion for sports and cycling, and is actively preparing for an international exchange program—a natural step in our family, given that he has an older maternal half-sister who is married and has been living in the US for over 10 years."
        ],
        "qa_responses": [
            {"q": "How do your travel experiences and family background help you as a Data Scientist at micro1?", "a": "Traveling teaches you pattern recognition and adaptability. For instance, notice how I spotted the unique design of rubber-tired trains in Paris, or the engineering scale of the walkways in Madrid and the turbines at Itaipu. Furthermore, guiding my 15-year-old son through his international exchange preparation while tracking family roots in the US shows the strategic planning and long-term vision I apply to life and data products."},
            {"q": "We noticed you visited the RICS library in London. What drove you there?", "a": "I am deeply curious about institutional structures, standards, and regulatory frameworks. Stepping inside a century-old symbol of governance like the Merchant House or studying at RICS highlights my appreciation for established compliance, data quality, and documentation benchmarks."},
            {"q": "How do your sports background and personal routine relate to micro1's autonomous remote model?", "a": "Sports like capoeira demand flexibility, respect for tradition, and quick reflexes, while swimming requires consistency and long-term stamina. In São Paulo, I maintain this active routine by exploring parks and theaters. I bring this exact same energy and discipline to micro1's remote ecosystem, tackling complex data challenges with resilience and an absolute focus on delivering consistent, high-quality results for the team."}
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
        
    if st.button("📘 View: micro1 Target Requirements.pdf", use_container_width=True):
        st.session_state.view_mode = "Guide Doc"
        st.rerun()
    
    st.markdown("---")
    st.markdown("### Strategic Framework")
    st.info("**WHY:** Motivation & Fit\n\n**WHAT:** Scope & Profile\n\n**HOW:** STAR Actions\n\n**WHEN:** Crisis & Investigative Closing")
    
    st.markdown("### Match Analytics")
    st.metric(label="micro1 Adherence Score", value="98%", delta="Elite AI Lab Match")
    st.caption("**Target:** micro1 · Data Scientist")

# --- VIEW ROUTING ENGINE ---

if st.session_state.view_mode == "CV Doc":
    st.markdown('<div class="doc-container"><div class="doc-title">Document View: André Carvalho — Resume</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="doc-section">
        <strong>ENG. ANDRÉ CARVALHO, MBA</strong><br>
        Senior Data Analyst | Data Analytics | Business Intelligence | SQL | Python
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
    st.markdown('<div class="doc-container"><div class="doc-title">Document View: micro1 Alignment Matrix</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="doc-section">
        <strong>Strategic Blueprint: micro1 AI Data Scientist Role</strong>
    </div>
    <div class="doc-subtitle">Core Job Alignment Matrix</div>
    <div class="doc-section">
        • <strong>The Role Paradigm:</strong> micro1 is not a traditional corporate spreadsheet shop. It is a cutting-edge data laboratory training next-generation foundation models. Your role as a Data Scientist means ensuring strict data readiness—collecting, cleaning, preprocessing, validating statistical consistency, and deploying automation workflows.<br>
        • <strong>The Tech Stack Bridge:</strong> The job requires Python proficiency, statistical data modeling, data collection, and visualization. Your real-world data runs (K-Means clustering at Fretebras/Stalse, time-series concepts, and heavy Python pipeline automation at Afinz) perfectly fulfill these qualifiers.<br>
        • <strong>Key Triggers to Highlight:</strong><br>
        &nbsp;&nbsp;&nbsp;&nbsp;- <em>Data Cleaning & Preprocessing:</em> Use the Heineken case (sorting out 10,000 product text variations into a strict schema) as absolute gold proof of your meticulous attention to data quality.<br>
        &nbsp;&nbsp;&nbsp;&nbsp;- <em>Statistical Modeling:</em> Showcase your Production Engineering logic and MBA background to prove you treat data data validation as a rigorous scientific discipline.<br>
        &nbsp;&nbsp;&nbsp;&nbsp;- <em>Agile & Remote Autonomy:</em> Your remote contractor experiences across multi-cloud infrastructure environments (AWS at Itaú/Afinz, GCP at ASICS) show you are ready to deliver end-to-end data products independently.
    </div>
    <div class="commentary-box">
        <strong>C-Level Interview Tip:</strong> Always frame your answers around <strong>Data Integrity</strong>. Emphasize that you don't just dump raw files into a script—you audit schemas, clear out anomalies, map metadata, and build automated verification frameworks that guarantee top-tier training inputs.
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
                <strong style="color:#d35400; text-transform:uppercase; font-size:9px;">The micro1 Alignment Link (The Strategic Approach):</strong><br>
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
