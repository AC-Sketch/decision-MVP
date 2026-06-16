import streamlit as st

st.set_page_config(
    page_title="War Room - micro1 AI Data Science",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ADVANCED EXECUTIVE UX INJECTION ---
st.markdown("""
<style>
/* Reset main padding limits for cleaner structural proportions */
.block-container {
    padding-top: 1.5rem !important;
    padding-bottom: 1.0rem !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
    max-width: 100% !important;
}

/* Custom Scrollbar and Clean CSS Reset */
::-webkit-scrollbar {
    width: 6px;
    height: 6px;
}
::-webkit-scrollbar-thumb {
    background: #cbd5e1;
    border-radius: 4px;
}

/* Strategic Presentation Container Boxes */
.response-box {
    background-color: #f0fdf4;
    border-left: 5px solid #16a34a;
    padding: 12px 16px !important;
    border-radius: 6px;
    margin-bottom: 0.75rem;
}

.followup-box {
    background-color: #f8fafc;
    border-left: 5px solid #475569;
    padding: 12px 16px !important;
    border-radius: 6px;
    margin-bottom: 0.75rem;
}

.growth-box {
    background-color: #fffbeb;
    border-left: 5px solid #d97706;
    padding: 12px 16px !important;
    border-radius: 6px;
    margin-bottom: 0.75rem;
}

.match-box {
    background-color: #f0f9ff;
    border-left: 5px solid #0284c7;
    padding: 12px 16px !important;
    border-radius: 6px;
    margin-bottom: 0.75rem;
}

.bullet-container-box {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    padding: 14px 18px !important;
    margin-bottom: 0.75rem;
    box-shadow: 0 1px 3px rgba(0,0,0,0.02);
}

.qa-container-box {
    background-color: #f1f5f9;
    border: 1px solid #cbd5e1;
    border-left: 5px solid #0f172a;
    border-radius: 6px;
    padding: 14px 18px !important;
}

.qa-item {
    margin-bottom: 8px !important;
    padding-bottom: 8px;
    border-bottom: 1px dashed #cbd5e1;
}
.qa-item:last-child {
    border-bottom: none;
    margin-bottom: 0px !important;
    padding-bottom: 0px;
}

/* Sidebar Custom Look */
[data-testid="stSidebarUserContent"] {
    padding-top: 1.5rem !important;
}
</style>
""", unsafe_allow_html=True)


# --- 20 STRATEGIC FRAMEWORK ITEMS FOR MICRO1 AI DATA SCIENCE ---
DATA_MAPPING = {
    1: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Tell me about yourself",
        "tag": "PROFILE",
        "bridge": "Senior Data Scientist with a background in Production Engineering, an MBA in Business Administration, and extensive expertise in Analytics, Business Intelligence, Data Governance, Cloud Analytics, and FinOps.",
        "followup": "Proven track record in strategic data projects for market-leading corporations like Heineken, Itaú, ASICS Latam, Gerdau, Sabesp, Fretebras, Afinz, and others, supporting high-level decision-making by transforming raw data into actionable insights and business indicators.",
        "match": "Expertise in data modeling, developing ETL/ELT pipelines, process automation, KPI consolidation, executive dashboards, and cloud analytics environment optimization.",
        "growth": "Fully prepared to apply my engineering and mathematical rigor to train and evaluate next-generation LLMs and AI agents inside the micro1 ecosystem.",
        "case": "Production Engineering + MBA + Advanced Python & SQL Modeling (Heineken, Itaú, ASICS).",
        "bullets": [
            "Automation of analytical processes and report update workflows, reducing operational effort by over 80%. Slashed dashboard refresh times from 1h30m to just 15 minutes by implementing structured data pipelines, automations, and process standardization.",
            "Consolidation and modeling of multiple complex databases for major retail and consumer goods operations, including scenarios where a single product had over 10,000 distinct denominations across different clients, channels, and systems. Structured analytical models capable of unifying and correlating dispersed data to generate reliable indicators and support strategic decision-making.",
            "Re-engineering of cloud analytics architectures, redesigning multiple tables and views into a scalable, optimized structure. Implemented partitioning, incremental processing, and FinOps best practices, slashing data processing consumption from GB/TB scale down to MB.",
            "Development of executive dashboards and Analytics solutions to monitor commercial, financial, operational, marketing, eCommerce, and eRetail indicators, including LATAM operations (Brazil, Chile, and Colombia). Applied Machine Learning and Advanced Analytics for revenue forecasting and business pattern identification utilizing models such as ARIMA+ and K-Means."
        ],
        "qa_responses": [
            {"q": "How does your engineering background enhance your data science execution?", "a": "Production Engineering focuses on systemic efficiency and optimization. Combined with an MBA, I don't just build statistical models; I design robust data validation, cleaning, and transformation processes that ensure real-world parameters map flawlessly into analytical frameworks."},
            {"q": "What is your experience with statistical and predictive modeling?", "a": "I routinely deploy advanced statistical analysis, time-series forecasting (such as ARIMA+), and clustering algorithms (like K-Means) to isolate business patterns, predict revenues, and segment massive multi-country datasets."},
            {"q": "How do you ensure data quality across highly fragmented pipelines?", "a": "At Heineken, I normalized datasets containing over 10,000 conflicting naming variations into a single cohesive Star Schema. I apply strict Data Quality and governance checks at every stage—ingestion, cleaning, and modeling—ensuring the outputs are 100% reliable."}
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
        "case": "Data Quality Scaling (Heineken) + Cloud Pipeline Automation (Afinz / ASICS).",
        "bullets": [
            "Frontier models require gold-standard data inputs to prevent hallucinations and ensure logical reasoning.",
            "My experience in highly regulated and complex fields (Itaú, Heineken) allows me to act as an elite domain expert and data evaluator.",
            "I am highly driven by micro1's remote, global-first laboratory model where data engineering meets AI alignment."
        ],
        "qa_responses": [
            {"q": "Why apply to an AI training lab instead of a traditional corporate data role?", "a": "Traditional corporate roles focus on visualizing the past. micro1 focuses on building the future of intelligence. Applying my data science and engineering rigor to benchmark, clean, and evaluate datasets for LLMs is the highest-leverage application of my skills."},
            {"q": "How do you view micro1's role in the market?", "a": "As models scale, raw internet scraping is exhausted. The winner of the AI race will be the company that curates the best specialized human feedback and agent evaluation pipelines. micro1 is exactly that layer."}
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
        "case": "Meticulous Pre-processing and Metadata Repositories (Afinz / Burity).",
        "bullets": [
            "I am accustomed to auditing complex, unstructured records to rectify anomalies before they hit production environments.",
            "At Afinz, I built metadata frameworks that secured data consistency while speeding up processing lines by 80%.",
            "I treat data quality as an algorithmic discipline, setting up automated validation gates to catch drift instantly."
        ],
        "qa_responses": [
            {"q": "What does 'meticulous attention to detail' mean to you in a data context?", "a": "It means validating assumptions at the lowest granularity. It means checking null behavior, analyzing feature distributions, ensuring balance in datasets, and documenting every step so that the pipeline is completely auditable and reproducible."}
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
        "case": "End-to-End Data Science Lifecycle (ASICS / Itaú / Heineken).",
        "bullets": [
            "I write optimized, clean Python and SQL code that scales smoothly across AWS and GCP cloud infrastructures.",
            "I translate chaotic, disparate client formats into rigid, standardized data assets.",
            "I communicate complex statistical findings with absolute clarity to both technical and C-level stakeholders."
        ],
        "qa_responses": [
            {"q": "How do you communicate data science models to non-technical teams?", "a": "By using clear data storytelling. I never showcase raw code lines or deep mathematical proofs to business units. I present the model's reliability boundaries, its business impact, and its visual trends through intuitive dashboards."}
        ]
    },
    5: {
        "category": "WHY - Intent & Fit",
        "title": "Salary Expectations",
        "tag": "ANCHOR",
        "bridge": "My salary expectations are aligned with the international market standard for a senior data professional operating in a high-leverage contractor capacity.",
        "followup": "I am open to micro1's standardized contract framework, taking into account the long-term impact metrics and growth roadmap of the ecosystem.",
        "match": "Maintains absolute professional transparency while establishing deep corporate seniority.",
        "growth": "Demonstrates an owner mindset, grounding the rate discussion strictly on delivered efficiency.",
        "case": "Standard Remote Contractor Range Alignment.",
        "bullets": [
            "I target standard senior developer/data scientist benchmarks for global remote talent allocation.",
            "I am completely comfortable with B2B international contractor arrangements.",
            "My infrastructure optimizations naturally offset resource overhead from day one."
        ],
        "qa_responses": [
            {"q": "Are you flexible regarding the contract structure?", "a": "Yes, I am fully open to discussing the overall alignment package, especially considering the project's complexity and remote collaboration framework."}
        ]
    },
    6: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Data Manipulation Stack",
        "tag": "TECH-STACK",
        "bridge": "I treat Python and SQL as native languages to manipulate, clean, and model complex, high-volume datasets.",
        "followup": "My cloud experience spans AWS Athena, Glue, and S3 (Itaú) and Google BigQuery (ASICS), giving me complete fluidity to execute exploratory data analysis across distributed environments.",
        "match": "Directly satisfies the technical requirement for programming, data cleaning, and data visualization tools.",
        "growth": "Ensures that you can jump into micro1’s platform and interface with any cloud structure with zero onboarding lag.",
        "case": "High-Volume Query Optimization + Streamlit Application Building.",
        "bullets": [
            "I leverage libraries like Pandas, NumPy, and Scikit-Learn to build, clean, and validate predictive structures.",
            "I am highly comfortable designing interactive visual environments using Streamlit to prototype analytical data views fast.",
            "I apply FinOps to data engineering, refactoring views to slash resource consumption from Terabytes to Megabytes."
        ],
        "qa_responses": [
            {"q": "Are you comfortable working outside a standard SQL environment?", "a": "Yes. My Python foundation allows me to manipulate unstructured data, parse JSON log payloads from user-agent interactions, and handle diverse API outputs natively."}
        ]
    },
    7: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Agile & Product Delivery",
        "tag": "AGILE",
        "bridge": "I execute within modern agile framework rules, guaranteeing that data workflows are treated as evolving products with clear feature scoping.",
        "followup": "My background in squads ensures I maintain fluid communication lines with product managers, developers, and data engineers synchronously.",
        "match": "Confirms your absolute adaptability to modern, fast-paced startup workflows.",
        "growth": "Eliminates operational latency when collaborating across distributed tech engineering lines.",
        "case": "Agile Squad Interaction (NTT Data / Itaú).",
        "bullets": [
            "I run data projects from ideation to delivery using clear user stories and definition criteria.",
            "I keep documentation up-to-date in Confluence, preventing code or process ambiguity.",
            "I focus on delivering stable MVPs to maintain organizational delivery momentum."
        ],
        "qa_responses": [
            {"q": "How do you handle scope changes mid-sprint?", "a": "By applying structured prioritization. I evaluate the data dependency of the new feature and assess if it threats data quality boundaries before adapting the workflow."}
        ]
    },
    8: {
        "category": "WHY - Intent & Fit",
        "title": "Short Tenures (Stalse/NTT)",
        "tag": "AGILITY-SPRINT",
        "bridge": "These roles were structured, high-intensity contract projects brought in as tactical consulting sprints to solve specific engineering blocks.",
        "followup": "At Stalse, my mission was to unify Latin American revenue models for ASICS. At NTT Data, I optimized complex Athena views for Itaú. Once the pipelines were automated, the goals were met.",
        "match": "Frames short tenures as high-impact, intentional consulting deliveries rather than professional instability.",
        "growth": "Proves that you have an extreme learning curve and can deliver immediate value without onboarding friction.",
        "case": "High-Impact Temporary Sprints (2025 - 2026).",
        "bullets": [
            "I entered complex corporate cloud environments and delivered automated tools in less than 4 months.",
            "Every project I touch is left fully documented, ensuring zero technical debt for the internal squads.",
            "Now, I am looking to invest this cross-industry cloud arsenal into micro1's long-term AI scaling."
        ],
        "qa_responses": [
            {"q": "How did you manage to deliver results so fast at ASICS and Itaú?", "a": "By utilizing structured thinking and a plug-and-play technical mindset. I immediately analyzed the database schemas, identified the bottlenecks, and started writing optimization code from week two."}
        ]
    },
    9: {
        "category": "WHAT - Capabilities & Profile",
        "title": "Product Analytics Focus",
        "tag": "PRODUCT-ANALYTICS",
        "bridge": "I connect backend dataset metrics to front-end user behavior logs cleanly, treating data as an interactive asset.",
        "followup": "My profile explicitly highlights my technical skills in Product Analytics frameworks (Pendo, Gainsight PX, WalkMe), ensuring I know how to track user digital adoption.",
        "match": "Perfect alignment for micro1's evaluation of AI agents and user-agent interaction tracking.",
        "growth": "Ensures that data solutions developed for micro1 act as intuitive tools that drive user adoption.",
        "case": "Product Analytics Stack Mastery + Digital Channel tracking.",
        "bullets": [
            "I help engineering squads understand user drop-offs using hard database log evidence.",
            "I design data validation criteria that measure active usage, data reliability, and script behavior.",
            "At Integral, I conducted over 2,000 structured analyses focused on human learning and engagement gaps."
        ],
        "qa_responses": [
            {"q": "What is the difference between a traditional report and a Data Product?", "a": "A traditional report is a static look at past data. A Data Product is an interactive, scalable, and documented asset engineered to solve a continuous operational problem, built with user adoption metrics and strict SLA tracking."}
        ]
    },
    10: {
        "category": "WHY - Intent & Fit",
        "title": "AI Frontier Alignment",
        "tag": "AI-VISION",
        "bridge": "The absolute limit of AI model scaling is no longer just processing power, but the acquisition of pristine data repositories curated by human experts.",
        "followup": "I am shifting my senior analytical drive to micro1 because your data laboratory sits exactly at the center of this paradigm change.",
        "match": "Positions you as a forward-thinking technologist who understands the macroeconomic realities of generative AI.",
        "growth": "Secures micro1's data advantage by embedding a rigorous engineer into its alignment pipelines.",
        "case": "Transition from Analytics to AI Training Ecosystems.",
        "bullets": [
            "I want to apply my cloud data arsenal to clean and optimize the data flows that train foundation models.",
            "My background allows me to spot statistical flaws in large text/log extractions rapidly.",
            "I thrive in environments where data engineering meets advanced cognitive reasoning models."
        ],
        "qa_responses": [
            {"q": "How do you see the evolution of data data science inside AI training loops?", "a": "We are moving away from brute-force scraping. The future depends on high-fidelity, syntactically perfect evaluation datasets—exactly what micro1 scales globally."}
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
            {"q": "How will you apply this automation focus at micro1?", "a": "I will systematically identify any manual validation or file formatting routines in your AI training data loops and replace them with automated Python workflows, giving hours of analysis time back to the squad."}
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
    17: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "High-Pressure Pivot",
        "tag": "CRISIS-PIVOT",
        "bridge": "When sudden model evaluation blockers or system drift occur, I anchor my strategy around empirical database tracking rather than emotional guessing.",
        "followup": "I isolate the baseline features, run rapid exploratory analysis, and deploy robust validation patches safely.",
        "match": "Confirms an organized, highly stable analytical lens during unpredictable execution sprints.",
        "growth": "Protects micro1's project schedules when tracking dynamic model training events globally.",
        "case": "Exploratory Data Scripting & Quick Diagnostics.",
        "bullets": [
            "I instantly isolate the metrics to evaluate what specific data silo is failing.",
            "I query our backend clusters to extract a reliable validation baseline sample.",
            "I build a clean MVP verification query to monitor pipeline parameters in real-time."
        ],
        "qa_responses": [
            {"q": "How do you handle urgent, undefined data requests?", "a": "I apply structured thinking. I treat the problem as an engineering system: mapping out the required inputs, the extraction limits, and the absolute mathematical logic needed to back the business choice."}
        ]
    },
    18: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "Conflict with Stakeholders",
        "tag": "STAKEHOLDERS",
        "bridge": "I remove personal biases and subjective disagreements from the conversation by laying hard statistical proof on the table.",
        "followup": "Teams usually argue because of underlying friction regarding target delivery alignment. Once the data lineage is visual, convergence happens naturally.",
        "match": "Validates advanced communication, client-facing experience, and stakeholder prioritization.",
        "growth": "Ensures seamless bridges between micro1’s platform managers and external domain specialists.",
        "case": "Data-Driven Consensus Frameworks.",
        "bullets": [
            "I start by practicing active listening to truly map the stakeholder's technical or operational bottleneck.",
            "I present clear, comparative data options, detailing the specific technical tradeoff of each selection.",
            "I let database evidence and model accuracy metrics drive the team's engineering resolution."
        ],
        "qa_responses": [
            {"q": "How do you handle a team member who insists on using a biased dataset?", "a": "I don't just point out the mistake. I compile a descriptive baseline report showing how the data skew will explicitly damage downstream model accuracy, mapping a transparent optimization path."}
        ]
    },
    19: {
        "category": "WHEN - Extreme Scenarios & Crisis",
        "title": "Technical Storytelling",
        "tag": "DATA-STORY",
        "bridge": "I translate complex database queries and statistical algorithms into clear commercial impact metrics and operational cost savings.",
        "followup": "Non-technical cross-functional leaders do not need to hear about database execution plans or syntax formatting; they need to know if the asset is stable.",
        "match": "Fulfills the requirement for excellent verbal and written communication with an emphasis on clarity.",
        "growth": "Empowers micro1's management team with intuitive, highly actionable summaries.",
        "case": "Executive Analytical Syntheses & Operational hour tracking.",
        "bullets": [
            "I focus on presenting time saved, cost reductions (FinOps), or data accuracy guarantees instead of raw logic syntax.",
            "I leverage clean dashboards to make validation and dataset health statuses instantly visual.",
            "I write crisp documentation summaries, making sure technical definitions are accessible to anyone."
        ],
        "qa_responses": [
            {"q": "How do you structure an analytical presentation for non-technical teams?", "a": "I use a top-down model: starting with the absolute core takeaway and business trajectory, followed by the supporting statistical data, and ending with a clear, automated log framework."}
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


# --- WORKSPACE APP ROUTING ---
if "view_mode" not in st.session_state:
    st.session_state.view_mode = "Main Interview Board"

# Sidebar Architecture
with st.sidebar:
    st.markdown("### Workspace Mode")
    if st.button("📊 Main Interview Board", use_container_width=True):
        st.session_state.view_mode = "Main Interview Board"
        st.rerun()
    if st.button("📄 View: André Carvalho Resume", use_container_width=True):
        st.session_state.view_mode = "CV Doc"
        st.rerun()
    if st.button("📘 View: micro1 Target Matrix", use_container_width=True):
        st.session_state.view_mode = "Guide Doc"
        st.rerun()
        
    st.markdown("---")
    st.markdown("### Strategic Framework")
    st.info("**WHY:** Intent & Fit\n\n**WHAT:** Core Profile\n\n**HOW:** STAR Deliveries\n\n**WHEN:** Crisis Containment")
    
    st.markdown("### Alignment Analytics")
    st.metric(label="micro1 Job Fit Match", value="98%", delta="Elite AI Stack Match")
    st.caption("**Target:** micro1 · Data Scientist (Remote)")

# --- RENDERER ENGINE ---

if st.session_state.view_mode == "CV Doc":
    st.markdown('<div class="doc-container"><div class="doc-title">Document View: André Carvalho — English Resume</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="doc-section">
        <strong>ENG. ANDRÉ CARVALHO, MBA</strong><br>
        Senior Data Scientist | Advanced Analytics | Distributed Cloud Frameworks | SQL & Python Expert
    </div>
    <div class="doc-subtitle">Professional Summary</div>
    <div class="doc-section">
        Highly analytical Senior Data professional combining a Production Engineering background with an FGV MBA. 
        Extensive execution history in Data Manipulation, Cloud Architecture Optimization, FinOps, Data Governance, and Machine Learning models. 
        Proven capability to clean, model, and automate end-to-end distributed pipelines for high-volume market leaders including Heineken, Itaú, ASICS Latam, Gerdau, Sabesp, Fretebras, and Afinz. 
        Deeply detail-oriented and optimized for high-velocity remote frameworks across international ecosystems.
    </div>
    <div class="doc-subtitle">Core Corporate Milestones</div>
    <div class="doc-section">
        • <strong>Stalse / ASICS Latam (2026):</strong> Redesigned multi-country data views in BigQuery, slashing data ingestion loads from GB/TB scale down to MB using strict partitioning and incremental structures (FinOps logic).<br>
        • <strong>NTT DATA / Itaú (2025):</strong> Engineered complex database queries via Amazon Athena to filter and process billions of transaction logs inside agile data squads.<br>
        • <strong>Sxpel / Heineken (2023 - 2024):</strong> Managed and cleaned chaotic digital channel data, resolving a massive taxonomy challenge where a single item carried over 10,000 conflicting denominations.<br>
        • <strong>Afinz (2022 - 2023):</strong> Automated manual MIS procedures via Python and SQL, cutting update loops from 1h30m down to 15 minutes, while building metadata catalogs in Confluence.
    </div>
    """, unsafe_allow_html=True)

elif st.session_state.view_mode == "Guide Doc":
    st.markdown('<div class="doc-container"><div class="doc-title">Document View: micro1 Target Matrix</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="doc-section">
        <strong>micro1 Laboratory Alignment Paradigm</strong>
    </div>
    <div class="doc-subtitle">How your profile answers their specific AI challenges:</div>
    <div class="doc-section">
        1. <strong>Data Cleaning & Processing:</strong> micro1 needs analysts obsessed with data precision to train foundation models. Your Heineken normalization architecture (sorting 10,000 text anomalies) is the ultimate proof of this skill.<br>
        2. <strong>Model Building & Statistics:</strong> Your Production Engineering degree, FGV MBA, and history deploying K-Means clustering and ARIMA+ time-series models map exactly to their requirement for structural data analysis.<br>
        3. <strong>Programming & Automation:</strong> Your deep mastery over Python and optimization scripts across AWS Athena and GCP BigQuery ensures you code for production, avoiding manual spreadsheet tech debt.<br>
        4. <strong>Communication:</strong> Your international background, advanced English, and structured business presentation style ensure you can synthesize technical findings for diverse cross-functional teams.
    </div>
    """, unsafe_allow_html=True)

else:
    # --- DEFAULT MAIN INTERVIEW BOARD VIEW WITH TABS ---
    st.title("🛡️ micro1 Data Science Strategy Board")
    st.caption("Strategic execution dashboard for peer-to-peer technical alignment.")
    
    # Organizing the 20 cards inside clean horizontal category tabs
    tab_categories = [
        "📋 WHAT - Profile & Stack",
        "🎯 WHY - Intent & Fit",
        "🚀 HOW - STAR Cases",
        "⚡ WHEN - Crisis & Closing"
    ]
    
    tab_objs = st.tabs(tab_categories)
    
    # Mapping cards to respective categories to ensure ALL 20 cards are represented cleanly
    category_mapping = {
        "WHAT - Capabilities & Profile": tab_objs[0],
        "WHY - Intent & Fit": tab_objs[1],
        "HOW - Case Methodology (STAR)": tab_objs[2],
        "WHEN - Extreme Scenarios & Crisis": tab_objs[3]
    }
    
    for cat_name, tab_obj in category_mapping.items():
        with tab_obj:
            cat_items = {k: v for k, v in DATA_MAPPING.items() if v["category"] == cat_name}
            
            if not cat_items:
                st.info("Additional alignment modules in this tier are calibrated dynamically.")
            
            for item_id, item_data in cat_items.items():
                # Clean, scannable expandable containers for each individual card to avoid screen clutter
                with st.expander(f"🔮 [{item_data['tag']}] — {item_data['title']}", expanded=(item_id == 1)):
                    col_out1, col_out2 = st.columns([0.50, 0.50])
                    
                    with col_out1:
                        st.markdown(f"""
                        <div class="response-box">
                            <span style="color:#16a34a; font-size:10px; font-weight:bold; text-transform:uppercase;">The Golden Bridge:</span><br>
                            <strong style="font-size:13px; color:#1e293b; line-height:1.3;">"{item_data['bridge']}"</strong>
                        </div>
                        <div class="followup-box">
                            <span style="color:#475569; font-size:10px; font-weight:bold; text-transform:uppercase;">Context / Elaboration:</span><br>
                            <p style="font-size:12.5px; color:#334155; line-height:1.3;">{item_data['followup']}</p>
                        </div>
                        <div class="growth-box">
                            <span style="color:#d97706; font-size:10px; font-weight:bold; text-transform:uppercase;">The micro1 Strategic Fit:</span><br>
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
                        
                        qa_html_items = ""
                        for qa in item_data['qa_responses']:
                            qa_html_items += f"""
                            <div class="qa-item">
                                <strong style="font-size:12px; color:#0f172a; display:block; line-height:1.2;">Q: {qa['q']}</strong>
                                <p style="font-size:12px; color:#1e293b; line-height:1.35; margin-top:2px !important;"><strong>A:</strong> {qa['a']}</p>
                            </div>
                            """
                        st.markdown(f"""
                        <div class="qa-container-box">
                            <span style="color:#0f172a; font-size:10.5px; font-weight:bold; text-transform:uppercase; display:block; margin-bottom:6px;">⚡ TOUGHEST FRONTIER AI Q&A SIMULATOR:</span>
                            {qa_html_items}
                        </div>
                        """, unsafe_allow_html=True)
