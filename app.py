import streamlit as st

st.set_page_config(
    page_title="War Room - Georgia IT | Senior Data Engineer GCP",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# WAR ROOM - SENIOR DATA ENGINEER GCP REMOTE
# Context: Georgia IT recruiter outreach | Remote Brazil/Mexico
# Goal: Interview preparation, strategic positioning, and Q&A
# ============================================================

st.markdown(
    """
<style>
.block-container {
    padding-top: 1rem !important;
    padding-bottom: 0.3rem !important;
    padding-left: 1rem !important;
    padding-right: 1rem !important;
    max-width: 100% !important;
}
::-webkit-scrollbar { display: none !important; }
div[data-testid="stVerticalBlock"] { gap: 0.2rem !important; }
[data-testid="stSidebarUserContent"] { padding-top: 1rem !important; }
h1, h2, h3, p, div { margin-top: 0rem !important; margin-bottom: 0rem !important; }
div[data-testid="stMetric"] {
    background-color: #f8f9fa;
    padding: 7px !important;
    border-radius: 6px;
    border: 1px solid #e9ecef;
    text-align: center;
}
div.stButton > button {
    width: 100% !important;
    min-height: 46px !important;
    white-space: normal !important;
    word-break: keep-all !important;
    overflow: hidden !important;
    font-size: 10.5px !important;
    line-height: 1.15 !important;
    padding: 0.35rem 0.35rem !important;
    text-align: center !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;
    border-radius: 6px !important;
    margin-bottom: 4px !important;
}
.category-header {
    font-size: 11px !important;
    font-weight: bold !important;
    color: #1b4f72;
    border-bottom: 2px solid #d6eaf8;
    padding-bottom: 4px;
    margin-bottom: 0.45rem !important;
    text-transform: uppercase;
    letter-spacing: 0.3px;
}
.response-box, .followup-box, .growth-box, .match-box,
.bullet-container-box, .qa-container-box, .doc-container, .commentary-box {
    border-radius: 6px;
}
.response-box {
    background-color: #e8f8f5;
    border-left: 4px solid #18bc9c;
    padding: 8px 11px !important;
    margin-bottom: 0.35rem;
    min-height: 58px;
}
.followup-box {
    background-color: #f4f6f7;
    border-left: 4px solid #34495e;
    padding: 8px 11px !important;
    margin-bottom: 0.35rem;
    min-height: 58px;
}
.growth-box {
    background-color: #fef9e7;
    border-left: 4px solid #f39c12;
    padding: 7px 11px !important;
    margin-bottom: 0.35rem;
    min-height: 50px;
}
.match-box {
    background-color: #ebf5fb;
    border-left: 4px solid #3498db;
    padding: 7px 11px !important;
    margin-bottom: 0.35rem;
    min-height: 50px;
}
.bullet-container-box {
    background-color: #ffffff;
    border: 1px solid #e9ecef;
    padding: 8px 11px !important;
    min-height: 170px;
}
.qa-container-box {
    background-color: #f2f4f4;
    border: 1px solid #d5dbdb;
    border-left: 4px solid #1b4f72;
    padding: 8px 11px !important;
    margin-top: 5px;
    min-height: 220px;
}
.qa-item {
    margin-bottom: 7px !important;
    padding-bottom: 6px;
    border-bottom: 1px dashed #d5dbdb;
}
.qa-item:last-child { border-bottom: none; margin-bottom: 0px !important; }
.doc-container {
    background-color: #ffffff;
    border: 1px solid #d5dbdb;
    padding: 24px !important;
    box-shadow: 0 4px 6px rgba(0,0,0,0.03);
    max-height: 82vh;
    overflow-y: auto !important;
}
.doc-title {
    color: #1b4f72;
    font-size: 18px;
    font-weight: bold;
    border-bottom: 3px solid #1b4f72;
    padding-bottom: 8px;
    margin-bottom: 16px !important;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
.doc-section {
    font-size: 13px;
    color: #2c3e50;
    margin-bottom: 13px !important;
    line-height: 1.55;
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
    padding: 11px !important;
    margin-top: 8px !important;
    margin-bottom: 12px !important;
    font-size: 12.5px;
}
.small-label {
    font-size: 10px;
    font-weight: bold;
    text-transform: uppercase;
    letter-spacing: 0.2px;
}
</style>
""",
    unsafe_allow_html=True,
)

DATA_MAPPING = {
    1: {
        "category": "WHO - Positioning & Fit",
        "title": "Tell me about yourself",
        "tag": "PROFILE",
        "bridge": "I am a data professional with over 10 years of experience across Data Analytics, BI, Cloud Analytics and Data Engineering, with recent hands-on work in GCP, BigQuery, SQL, Python, ETL/ELT and Data Modeling.",
        "followup": "My recent work includes restructuring BigQuery tables and views, implementing scalable data models, applying incremental processing logic, reducing cloud consumption, and building analytics solutions for LATAM operations.",
        "match": "Strong fit for a GCP Data Engineer role requiring BigQuery, SQL, Python, cloud analytics, data modeling, data quality and pipeline development.",
        "growth": "I am especially interested in expanding my experience in international data engineering environments and global remote teams.",
        "case": "ASICS LATAM + Itaú + Heineken + Afinz",
        "bullets": [
            "Over 10 years of experience combining analytics, BI, cloud data, governance and data engineering responsibilities.",
            "Recent hands-on GCP and BigQuery work: 16+ tables and views restructured into a more scalable architecture.",
            "Strong SQL and Python foundation applied to ETL/ELT, data pipelines, data modeling and automation.",
            "Business-facing profile: able to translate stakeholder requirements into reliable technical data assets.",
        ],
        "qa_responses": [
            {"q": "How would you summarize your profile in two minutes?", "a": "I combine analytics engineering, BI, cloud data engineering and business understanding. I have worked with GCP, BigQuery, SQL, Python, ETL/ELT, data modeling and executive analytics, and I am looking to apply this experience in a remote international data engineering environment."},
            {"q": "What makes you relevant for this role?", "a": "The role asks for GCP, BigQuery, Python, SQL, ETL/ELT and data warehousing. These are directly aligned with my recent ASICS LATAM work and with my broader experience in AWS, BI, data modeling and automation."},
            {"q": "Are you more BI, analytics engineering or data engineering?", "a": "My foundation is BI and analytics, but my recent projects moved strongly into analytics engineering and cloud data engineering: BigQuery modeling, pipeline logic, SQL/Python automation, incremental processing and cloud cost optimization."},
        ],
    },
    2: {
        "category": "WHO - Positioning & Fit",
        "title": "Why this opportunity?",
        "tag": "MOTIVATION",
        "bridge": "This opportunity is attractive because it is a remote international GCP data engineering role, which matches both my current technical stack and my professional growth direction.",
        "followup": "I want to work in a global data environment where BigQuery, Python, SQL, ETL/ELT, data modeling and cloud optimization are used to build reliable data products at scale.",
        "match": "Shows motivation without sounding desperate or junior; the message is growth-oriented and aligned with the role.",
        "growth": "The correct positioning is: strong current experience, open to learning the client’s standards, and motivated to grow in international delivery.",
        "case": "International remote role + GCP specialization",
        "bullets": [
            "Do not say: I can be junior.",
            "Say: I am open to discussing scope depending on client expectations.",
            "Emphasize international growth, not lack of seniority.",
            "Keep the conversation focused on fit, responsibilities and value delivered.",
        ],
        "qa_responses": [
            {"q": "Why are you interested in this position?", "a": "Because it connects directly with my recent work in GCP and BigQuery and allows me to expand my experience in international data engineering teams. I see it as an opportunity to contribute immediately while continuing to grow in a global environment."},
            {"q": "Why Georgia IT?", "a": "Georgia IT appears to connect technology professionals with international clients, and this opportunity is aligned with the type of remote cloud data engineering role I am looking for. I would like to better understand the client, scope and team structure."},
            {"q": "What are you looking for in your next role?", "a": "A data engineering or analytics engineering role where I can work with cloud platforms, BigQuery, SQL, Python, data pipelines, data quality and scalable data models, ideally in an international remote team."},
        ],
    },
    3: {
        "category": "WHO - Positioning & Fit",
        "title": "Senior vs Mid-Level positioning",
        "tag": "LEVEL",
        "bridge": "My profile is strong in analytics engineering, BI and cloud data engineering; I am comfortable discussing Senior or Mid-Level scope depending on the client’s expectations.",
        "followup": "The key is not to down-level myself proactively. I should present evidence, ask about scope, and let the interview clarify whether the role is hands-on BigQuery/pipelines or deeper platform engineering.",
        "match": "Protects professional value while showing humility and flexibility.",
        "growth": "Ideal wording: strong experience, transparent learning mindset, and interest in international data engineering maturity.",
        "case": "GCP/BigQuery hands-on + international growth",
        "bullets": [
            "Avoid calling yourself junior because the CV shows 10+ years in data and recent cloud engineering work.",
            "Use Mid-Level only if the scope is more advanced than your current platform engineering experience.",
            "Frame it as scope alignment, not insecurity.",
            "Mention international growth only as a motivation, not as a weakness.",
        ],
        "qa_responses": [
            {"q": "Are you senior enough for this role?", "a": "I believe my experience is highly relevant, especially in GCP, BigQuery, SQL, Python, data modeling, ETL/ELT and cloud optimization. My fit depends on the exact scope. If the role is focused on BigQuery, data models, pipelines and business-facing data engineering, I can contribute strongly."},
            {"q": "Would you consider a Mid-Level role?", "a": "Yes, I am open to discussing scope. My priority is to join a strong international data engineering environment where I can contribute with my current cloud and analytics engineering experience while growing with the team."},
            {"q": "Why not apply as junior?", "a": "I would not position myself as junior because I already have strong experience in data, cloud analytics, modeling and automation. The most accurate positioning is Senior or Mid-Level depending on the role’s technical depth and expectations."},
        ],
    },
    4: {
        "category": "TECH - GCP Data Engineering",
        "title": "GCP and BigQuery experience",
        "tag": "GCP",
        "bridge": "My strongest recent cloud experience is with GCP and BigQuery, especially designing analytical structures, optimizing SQL views, implementing incremental logic and reducing processing consumption.",
        "followup": "At ASICS LATAM, I restructured more than 16 BigQuery tables and views, consolidated multicountry analytics, and applied FinOps practices such as partitioning and optimized processing.",
        "match": "Direct match with the role’s GCP, BigQuery, SQL, Python and data warehousing requirements.",
        "growth": "Position this as real production experience, not only dashboard consumption.",
        "case": "ASICS LATAM - BigQuery optimization and LATAM analytics",
        "bullets": [
            "BigQuery table and view restructuring.",
            "Partitioning, incremental processing and cost optimization.",
            "Looker, BQML and Vertex AI exposure for advanced analytics.",
            "Multicountry data consolidation for Brazil, Chile and Colombia.",
        ],
        "qa_responses": [
            {"q": "What is your experience with BigQuery?", "a": "I used BigQuery to restructure analytical tables and views, build scalable models and optimize query consumption. In one project, we reduced processing from GB/TB scale to MB in specific workloads by using better modeling, partitioning and incremental logic."},
            {"q": "How do you optimize BigQuery cost and performance?", "a": "I start by analyzing query patterns, table size, partitioning opportunities, unnecessary scans, repeated transformations and materialization strategy. Then I redesign the model to reduce scanned data, improve joins, and use incremental processing whenever possible."},
            {"q": "What GCP services have you used?", "a": "My recent experience includes GCP with BigQuery, Looker, BQML and Vertex AI. I also understand the broader GCP data ecosystem and can adapt to Cloud Storage, Cloud Composer and related pipeline orchestration components."},
        ],
    },
    5: {
        "category": "TECH - GCP Data Engineering",
        "title": "Python and SQL depth",
        "tag": "SQL-PYTHON",
        "bridge": "SQL and Python are my main technical tools for extracting, transforming, validating and automating data workflows.",
        "followup": "I use SQL for modeling, transformations, validations, joins, analytical views and performance tuning; Python supports automation, ETL routines, data checks and workflow standardization.",
        "match": "The role explicitly requires Python and SQL; this is one of the safest areas to emphasize.",
        "growth": "Speak in terms of business reliability, not only syntax.",
        "case": "BigQuery SQL + Athena SQL + Python automation",
        "bullets": [
            "Complex SQL queries and views in BigQuery and Amazon Athena.",
            "Python automation for reporting, data routines and ETL support.",
            "SQL validation logic for data quality and consistency checks.",
            "Performance mindset: fewer scans, cleaner joins, better models.",
        ],
        "qa_responses": [
            {"q": "How strong is your SQL?", "a": "SQL is one of my core strengths. I have used it in BigQuery and Athena for complex transformations, analytical views, data quality validations, joins, aggregations and performance improvements in cloud environments."},
            {"q": "How do you use Python in data engineering?", "a": "I use Python mainly for automation, data transformation support, validation routines, pipeline logic and reducing manual operational work. My focus is to make recurring data processes more reliable, repeatable and governed."},
            {"q": "How do you validate a transformation?", "a": "I compare source totals, row counts, null rates, duplicate keys, business rules and metric reconciliation. I also prefer creating clear validation queries so the logic is auditable and repeatable."},
        ],
    },
    6: {
        "category": "TECH - GCP Data Engineering",
        "title": "Airflow / Cloud Composer gap",
        "tag": "AIRFLOW",
        "bridge": "I understand orchestration logic well through ETL/ELT, automation and cloud workflows, even if Airflow or Cloud Composer is not my deepest hands-on tool yet.",
        "followup": "The right answer is honest: dependency management, scheduling, retries, monitoring and incremental loads are familiar concepts; Airflow/Composer would be a tool adaptation, not a conceptual gap.",
        "match": "Prepares a credible answer for one of the likely recruiter or technical screening questions.",
        "growth": "Do not overclaim. Show confidence based on Python, SQL and pipeline experience.",
        "case": "Python/SQL pipelines + AWS Glue + BigQuery workflows",
        "bullets": [
            "Do not say you are an Airflow expert if you are not.",
            "Connect Airflow to orchestration concepts you already know.",
            "Mention DAG logic: dependencies, retries, scheduling and monitoring.",
            "Show fast learning based on strong Python/SQL/cloud foundation.",
        ],
        "qa_responses": [
            {"q": "Have you worked with Airflow or Cloud Composer?", "a": "I have worked with ETL/ELT, automation and cloud-based data processes using Python, SQL, BigQuery and AWS Glue. Airflow/Composer is not my deepest tool, but I understand the orchestration logic: DAG dependencies, scheduled jobs, retries, monitoring and incremental loads. Given my Python and cloud data background, I can adapt quickly."},
            {"q": "How would you design an Airflow DAG for a daily pipeline?", "a": "I would separate extraction, staging validation, transformation, load, data quality checks and notification tasks. I would use retries, clear dependencies, parameterized execution dates and logs so failures are traceable and recoverable."},
            {"q": "What would you monitor in a pipeline?", "a": "Runtime, row counts, failed tasks, schema drift, null rates, duplicate keys, cost anomalies and business metric reconciliation. A pipeline is not finished when it runs; it is finished when it is observable and trusted."},
        ],
    },
    7: {
        "category": "TECH - GCP Data Engineering",
        "title": "ETL/ELT and data pipelines",
        "tag": "PIPELINES",
        "bridge": "I have built and improved ETL/ELT routines, data pipelines and automated workflows using SQL, Python and cloud data platforms.",
        "followup": "My pipeline mindset is based on clear stages: raw ingestion, staging validation, transformation, data quality, analytical model and consumption layer.",
        "match": "Directly supports the job requirement for ETL/ELT development and data pipeline work.",
        "growth": "Use structured vocabulary: staging, transformation, validation, monitoring, lineage.",
        "case": "Afinz automation + ASICS BigQuery + Itaú Athena",
        "bullets": [
            "Reduced reporting update effort by more than 80% at Afinz.",
            "Reduced dashboard update time from 1h30 to 15 minutes.",
            "Built data pipelines using Python, SQL, Power Query and cloud tools.",
            "Implemented data quality and governance practices to support reliability.",
        ],
        "qa_responses": [
            {"q": "How do you approach building a data pipeline?", "a": "I start with the business output, map the source data, define staging rules, apply transformations, validate quality, and only then expose the data to dashboards or analytical layers. I also document lineage and ownership."},
            {"q": "ETL or ELT?", "a": "It depends on the platform and volume. In modern cloud warehouses like BigQuery, ELT is often efficient because we can load data and transform it inside the warehouse using optimized SQL. But validation and governance still need to be designed carefully."},
            {"q": "What is a good pipeline?", "a": "A good pipeline is reliable, observable, cost-efficient, documented and aligned with business definitions. It should not only move data; it should deliver trusted data products."},
        ],
    },
    8: {
        "category": "TECH - GCP Data Engineering",
        "title": "Data warehousing and modeling",
        "tag": "MODELING",
        "bridge": "Data modeling is one of my strongest areas, especially dimensional structures, analytical views, KPI consolidation and business-oriented data layers.",
        "followup": "I have worked with complex product, commercial, financial, marketing and eCommerce data, building models that allow executives and operational teams to trust the same metrics.",
        "match": "The role asks for data warehousing and data modeling; this is a key differentiator from pure pipeline candidates.",
        "growth": "Position modeling as reliability, performance and business alignment.",
        "case": "Heineken product consolidation + ASICS LATAM data model",
        "bullets": [
            "Dimensional modeling and KPI consolidation.",
            "Consolidation of 10,000+ product denominations into a unified analytical model.",
            "BigQuery modeling for multicountry analytics.",
            "Data quality and traceability embedded in the model design.",
        ],
        "qa_responses": [
            {"q": "How do you design a data warehouse model?", "a": "I start from business processes and core metrics, then define facts, dimensions, grains, keys and transformation rules. I also consider query performance, data quality and how users will consume the model."},
            {"q": "What is your experience with dimensional modeling?", "a": "At Heineken, I structured analytical models to consolidate fragmented product and channel data, including scenarios with more than 10,000 different product denominations. The objective was to create a reliable and traceable performance view."},
            {"q": "How do you avoid metric disagreement?", "a": "By defining metric ownership, source-of-truth rules, transformation logic, reconciliation queries and documentation. Most metric conflicts are governance problems before they are technical problems."},
        ],
    },
    9: {
        "category": "CASES - STAR Evidence",
        "title": "ASICS LATAM - GCP/BigQuery",
        "tag": "ASICS",
        "bridge": "At ASICS LATAM, I helped modernize the analytics architecture using GCP and BigQuery for multicountry commercial, financial and marketing indicators.",
        "followup": "The work involved restructuring 16+ tables and views, consolidating data from digital platforms, applying BigQuery optimization and supporting advanced analytics with BQML and Vertex AI.",
        "match": "This is the strongest case for the Georgia IT GCP role.",
        "growth": "Use this case early in the interview because it maps directly to BigQuery, GCP, ETL/ELT, data modeling and FinOps.",
        "case": "Situation-Action-Result: GCP analytics modernization",
        "bullets": [
            "Situation: LATAM analytics needed more scalable and reliable structures.",
            "Action: Redesigned BigQuery tables/views, improved modeling, applied partitioning and incremental logic.",
            "Result: Reduced cloud processing consumption in specific workloads from GB/TB scale to MB.",
            "Business impact: faster, more reliable analytics for Brazil, Chile and Colombia.",
        ],
        "qa_responses": [
            {"q": "Tell me about a GCP project you delivered.", "a": "At ASICS LATAM, I worked on a GCP/BigQuery analytics modernization project. I restructured more than 16 tables and views, improved the data model, applied partitioning and incremental processing logic, and reduced processing consumption significantly in specific workloads."},
            {"q": "What was the business value?", "a": "The business gained faster, more reliable dashboards and lower cloud processing cost. The solution supported multicountry indicators across Brazil, Chile and Colombia, helping stakeholders make decisions using a more consistent data layer."},
            {"q": "What would you improve if you did it again?", "a": "I would add even more automated data quality checks and pipeline observability from the beginning, so every transformation stage had clear SLA, anomaly detection and lineage documentation."},
        ],
    },
    10: {
        "category": "CASES - STAR Evidence",
        "title": "Itaú / NTT DATA - AWS scale",
        "tag": "ITAU",
        "bridge": "At NTT DATA allocated to Itaú, I worked in an AWS analytics environment developing and optimizing SQL queries and views using Athena.",
        "followup": "This experience is useful because it proves adaptability across cloud environments and work inside agile squads for a large enterprise client.",
        "match": "Even though the role is GCP, AWS experience reinforces cloud fundamentals and enterprise maturity.",
        "growth": "Frame AWS-to-GCP as transferable cloud data engineering logic, not a mismatch.",
        "case": "Athena + QuickSight + Glue + S3 + Agile squads",
        "bullets": [
            "Developed and optimized complex SQL queries in Amazon Athena.",
            "Worked with QuickSight, Glue, S3 and agile squads.",
            "Translated business rules into analytical views and dashboards.",
            "Built enterprise experience in a high-governance financial environment.",
        ],
        "qa_responses": [
            {"q": "How does your AWS experience help in a GCP role?", "a": "The cloud platform changes, but the principles remain similar: understand storage, compute, query optimization, data modeling, governance and cost. Moving from Athena to BigQuery requires syntax and platform adaptation, but the engineering mindset is transferable."},
            {"q": "What did you do at Itaú?", "a": "I developed and optimized SQL queries in Athena, supported corporate analytics, worked with QuickSight dashboards and participated in agile squads to deliver data solutions aligned with business rules."},
            {"q": "What did you learn from banking data environments?", "a": "Precision, governance and traceability. In banking, data definitions and consistency matter a lot, so I learned to be careful with business rules, validation and documentation."},
        ],
    },
    11: {
        "category": "CASES - STAR Evidence",
        "title": "Heineken - Data modeling",
        "tag": "HEINEKEN",
        "bridge": "At Heineken, I worked with digital and eCommerce analytics, consolidating complex product and commercial data into reliable analytical models.",
        "followup": "A major challenge was normalizing more than 10,000 product denominations across clients, channels and systems to build consistent KPI reporting.",
        "match": "Shows strong data modeling, data quality, stakeholder alignment and business analytics capability.",
        "growth": "Use this case when asked about messy data, modeling or business-facing analytics.",
        "case": "Product master data consolidation and KPI dashboards",
        "bullets": [
            "Consolidated 10,000+ product naming variations.",
            "Built executive dashboards in Power BI.",
            "Implemented data quality and traceability practices.",
            "Modeled commercial, marketing, eCommerce and revenue analytics indicators.",
        ],
        "qa_responses": [
            {"q": "Tell me about a messy data challenge.", "a": "At Heineken, the same product appeared with thousands of different names across sources. I helped consolidate those variations into a unified analytical model, improving data quality, traceability and KPI reliability."},
            {"q": "What does this show about your data engineering ability?", "a": "It shows that I understand the importance of data modeling, standardization and governance before analytics consumption. Good pipelines are not useful if the underlying business entities are inconsistent."},
            {"q": "How do you work with business stakeholders?", "a": "I translate their KPI needs into data definitions, source mappings and validation rules. I avoid technical jargon when speaking to business users and focus on reliability, decision impact and transparency."},
        ],
    },
    12: {
        "category": "CASES - STAR Evidence",
        "title": "Afinz - Automation and governance",
        "tag": "AFINZ",
        "bridge": "At Afinz, I automated reporting and data routines, reducing operational effort by more than 80% and cutting dashboard refresh time from 1h30 to 15 minutes.",
        "followup": "The project combined Python, SQL, Power BI, AWS Glue, ETL, data governance, metadata organization and data quality practices.",
        "match": "Strong proof of automation mindset, process improvement and governed data delivery.",
        "growth": "Use this when asked about impact, ownership, automation or measurable results.",
        "case": "Manual reporting transformed into governed automated workflows",
        "bullets": [
            "Reduced operational reporting effort by more than 80%.",
            "Reduced dashboard update time from 1h30 to 15 minutes.",
            "Developed data pipelines using Python, SQL and Power Query.",
            "Implemented Data Governance, Data Quality and metadata practices.",
        ],
        "qa_responses": [
            {"q": "Give me an example of measurable impact.", "a": "At Afinz, I reduced a reporting update process from 1h30 to 15 minutes and reduced operational effort by more than 80% through automation, ETL routines and process standardization."},
            {"q": "How did governance fit into the automation?", "a": "Automation without governance creates hidden risk. I helped structure metadata repositories and data quality practices so the automated process was traceable and maintainable."},
            {"q": "How would this apply to the new role?", "a": "In a GCP data engineering role, I would look for repetitive manual transformations, unstable reporting flows or expensive queries and convert them into governed, optimized and observable pipelines."},
        ],
    },
    13: {
        "category": "RISK - Gaps & Strategic Answers",
        "title": "Short tenures",
        "tag": "TENURE",
        "bridge": "Some recent roles were temporary or project-based engagements, focused on specific deliveries in cloud analytics, BI and data optimization.",
        "followup": "The correct framing is delivery-based consulting, not instability. Then pivot to the desire for a longer-term international data engineering environment.",
        "match": "Handles a likely recruiter concern proactively and professionally.",
        "growth": "Be concise; do not overexplain. Show maturity and forward direction.",
        "case": "STALSE/ASICS and NTT/Itaú temporary projects",
        "bullets": [
            "State clearly that the projects were temporary or project-based.",
            "Emphasize delivered outcomes and documentation.",
            "Say you are now looking for a stable international data engineering path.",
            "Avoid sounding defensive.",
        ],
        "qa_responses": [
            {"q": "Why were your recent roles short?", "a": "They were project-based engagements with defined scopes. At ASICS LATAM, the focus was analytics modernization and BigQuery optimization. At Itaú, the work was cloud analytics support through NTT DATA. I delivered specific outcomes and am now looking for a longer-term international data engineering opportunity."},
            {"q": "Will you stay long-term?", "a": "Yes. My current goal is to consolidate my experience in an international data engineering environment. A remote global role with GCP and BigQuery is exactly aligned with that direction."},
            {"q": "Were the projects completed?", "a": "Yes, they were structured temporary projects. The important point is that I delivered measurable improvements such as optimized BigQuery structures, cloud consumption reduction and automated analytics processes."},
        ],
    },
    14: {
        "category": "RISK - Gaps & Strategic Answers",
        "title": "English and global teams",
        "tag": "GLOBAL",
        "bridge": "I have advanced English and international exposure, including training in London, and I am comfortable communicating with global teams in a clear, structured way.",
        "followup": "The best strategy is to show practical communication: concise updates, written documentation, asynchronous collaboration and clarification of requirements.",
        "match": "The role is remote and international, so communication is as important as technical skill.",
        "growth": "Emphasize clarity, documentation and accountability.",
        "case": "Advanced English + St Giles London + LATAM analytics",
        "bullets": [
            "Advanced English with international training.",
            "Comfortable with remote collaboration and written communication.",
            "Experience with LATAM multicountry stakeholders.",
            "Uses documentation to reduce ambiguity.",
        ],
        "qa_responses": [
            {"q": "Are you comfortable working in English?", "a": "Yes. I have advanced English and international training in London. I am comfortable with interviews, technical discussions, documentation and remote collaboration in English."},
            {"q": "How do you work remotely with international teams?", "a": "I focus on clear written communication, structured updates, documented decisions and proactive clarification. In data work, ambiguity is risky, so I prefer to confirm definitions and acceptance criteria early."},
            {"q": "How do you handle time zone differences?", "a": "I am flexible and used to planning work around priorities. I use asynchronous documentation and clear handoffs to keep progress moving even when schedules do not fully overlap."},
        ],
    },
    15: {
        "category": "RISK - Gaps & Strategic Answers",
        "title": "Compensation and rate",
        "tag": "RATE",
        "bridge": "I am open to discussing compensation based on the client, contract model, responsibilities, expected seniority level and duration of the project.",
        "followup": "Do not give a number too early unless required. First ask for client, scope, duration, seniority expectations and budget range.",
        "match": "Professional, flexible and protects negotiation power.",
        "growth": "If forced to give a range, provide it only after understanding contract details.",
        "case": "Remote contractor Brazil - DOE compensation",
        "bullets": [
            "Ask for compensation range first.",
            "Clarify contract model: Brazil contractor/PJ, payment currency, duration and benefits.",
            "Avoid undervaluing yourself due to international inexperience.",
            "Connect rate to scope, autonomy and technical expectations.",
        ],
        "qa_responses": [
            {"q": "What is your expected rate?", "a": "I am open to discussing compensation based on the client, responsibilities, contract model and expected seniority level. I would prefer to understand the project scope, duration, team structure and budget range before confirming a specific number."},
            {"q": "Can you share a number now?", "a": "I can be flexible, but I would like to avoid misalignment before understanding whether the role is primarily BigQuery/SQL pipeline work, full platform engineering, or broader senior ownership. Could you share the expected range for Brazil-based contractors?"},
            {"q": "Are you open to contract?", "a": "Yes, I am open to a contract model for Brazil, depending on payment terms, duration, scope and overall conditions."},
        ],
    },
    16: {
        "category": "RISK - Gaps & Strategic Answers",
        "title": "Technical screening strategy",
        "tag": "SCREEN",
        "bridge": "For technical screening, the safest strategy is to anchor answers in BigQuery, SQL optimization, data modeling, pipeline stages, data quality and cost control.",
        "followup": "Avoid overclaiming tools. When a tool is not your deepest experience, connect it to underlying concepts and show how you would operate professionally.",
        "match": "Keeps answers credible and senior enough without creating risk in follow-up questions.",
        "growth": "Technical interviews reward precision more than big claims.",
        "case": "GCP + SQL + data modeling + orchestration concepts",
        "bullets": [
            "Strong claims: BigQuery, SQL, data modeling, ETL/ELT, BI/analytics engineering, optimization.",
            "Careful claims: Airflow/Cloud Composer, Spark/Databricks if asked.",
            "Bridge gaps with concepts: orchestration, dependencies, distributed processing, observability.",
            "Always bring answers back to reliability, scalability and cost.",
        ],
        "qa_responses": [
            {"q": "What should I emphasize most?", "a": "GCP, BigQuery, SQL, Python, ETL/ELT, data modeling, data quality, FinOps and measurable business impact. Your ASICS case should be your main technical proof."},
            {"q": "What should I avoid?", "a": "Avoid saying you are junior, avoid claiming deep Airflow or Spark expertise if challenged, and avoid sounding like you only build dashboards. Position yourself as analytics engineering and cloud data engineering."},
            {"q": "What is the closing message?", "a": "I believe my GCP, BigQuery, SQL, Python and data modeling experience can contribute to this role, and I am motivated to grow in an international data engineering team while delivering reliable, optimized data products."},
        ],
    },
    17: {
        "category": "RISK - Gaps & Strategic Answers",
        "title": "Questions to ask them",
        "tag": "ASK",
        "bridge": "Strong candidates ask questions that reveal scope, maturity and expectations without sounding suspicious or defensive.",
        "followup": "Ask about client, project objective, data stack, orchestration, team structure, contract duration, timezone and rate range.",
        "match": "Shows professionalism and protects you from poor-fit opportunities.",
        "growth": "Use questions to discover whether the role is truly Senior, Mid-Level, or flexible.",
        "case": "Recruiter screen and hiring manager interview",
        "bullets": [
            "Could you share more about the client and project scope?",
            "Is the main workload BigQuery modeling, pipeline development, migration or optimization?",
            "Which orchestration tools are used: Cloud Composer, Airflow, Dataflow or others?",
            "What is the expected seniority level and autonomy for this position?",
        ],
        "qa_responses": [
            {"q": "What should I ask the recruiter?", "a": "Could you share the client, project duration, contract model, compensation range, timezone expectations and interview process?"},
            {"q": "What should I ask the technical interviewer?", "a": "What are the main data sources, current pain points, orchestration tools, BigQuery volume, data quality challenges and expected first 90-day outcomes?"},
            {"q": "How do I ask about flexibility in level?", "a": "Based on the role scope, is the client strictly looking for a Senior Data Engineer, or are they open to strong Mid-Level/Analytics Engineering profiles with hands-on GCP and BigQuery experience?"},
        ],
    },
    18: {
        "category": "RISK - Gaps & Strategic Answers",
        "title": "Closing pitch",
        "tag": "CLOSE",
        "bridge": "My strongest value is the combination of GCP/BigQuery hands-on experience, SQL/Python execution, data modeling, optimization mindset and business-facing communication.",
        "followup": "Close with confidence, not apology. Reinforce that you can contribute and grow in the client’s international data engineering environment.",
        "match": "A clean closing helps the interviewer remember your value proposition.",
        "growth": "End interviews by connecting skills, motivation and availability.",
        "case": "Final 60-second pitch",
        "bullets": [
            "Thank them for the conversation.",
            "Repeat fit: GCP, BigQuery, SQL, Python, ETL/ELT, modeling.",
            "Mention measurable impact: optimization, automation, reliability.",
            "Express interest in next steps.",
        ],
        "qa_responses": [
            {"q": "How should I close the interview?", "a": "Thank you for the conversation. I believe my experience with GCP, BigQuery, SQL, Python, data modeling, ETL/ELT and cloud optimization is well aligned with the role. I am especially interested in contributing to an international data engineering environment where reliability, scalability and cost efficiency are important. I would be glad to continue to the next step."},
            {"q": "What if I feel the interview went only average?", "a": "Close calmly and reinforce your interest. A strong final summary can recover confidence and remind them that your core experience matches the role."},
            {"q": "What is the key message they should remember?", "a": "I am a GCP/BigQuery-oriented analytics engineering and data engineering professional who can build reliable data models, optimize cloud processing and communicate clearly with business and technical teams."},
        ],
    },
}

if "view_mode" not in st.session_state:
    st.session_state.view_mode = "Main Interview Board"
if "active_id" not in st.session_state:
    st.session_state.active_id = 1

with st.sidebar:
    st.markdown("### Select Workspace View")
    if st.button("📊 Main Interview Board", use_container_width=True):
        st.session_state.view_mode = "Main Interview Board"
        st.rerun()
    if st.button("📄 View: CV Match Summary", use_container_width=True):
        st.session_state.view_mode = "CV Doc"
        st.rerun()
    if st.button("📘 View: Interview Strategy", use_container_width=True):
        st.session_state.view_mode = "Guide Doc"
        st.rerun()

    st.markdown("---")
    st.markdown("### Strategic Framework")
    st.info(
        "**WHO:** Positioning & Fit\n\n"
        "**TECH:** GCP, BigQuery, SQL, Python\n\n"
        "**CASES:** STAR evidence\n\n"
        "**RISK:** gaps, rate, closing"
    )
    st.markdown("### Match Analytics")
    st.metric(label="Georgia IT / GCP Match", value="88%", delta="Strong BigQuery + GCP fit")
    st.caption("**Target:** Senior Data Engineer – GCP · Remote Brazil/Mexico")
    st.caption("Tone: confident, honest, technical, international.")

if st.session_state.view_mode == "CV Doc":
    st.markdown(
        """
        <div class="doc-container">
            <div class="doc-title">CV Match Summary — Senior Data Engineer GCP</div>
            <div class="doc-section">
                <strong>Target Role:</strong> Senior Data Engineer – GCP, Remote Brazil/Mexico.<br>
                <strong>Core Requirements:</strong> GCP, BigQuery, Python, SQL, Airflow/Cloud Composer, ETL/ELT, Data Warehousing and Data Modeling.
            </div>
            <div class="doc-subtitle">Strongest Fit Points</div>
            <div class="doc-section">
                • Recent hands-on GCP and BigQuery experience in ASICS LATAM.<br>
                • Restructuring of 16+ BigQuery tables and views.<br>
                • SQL, Python, ETL/ELT, data modeling, data pipelines and cloud analytics.<br>
                • Cloud optimization and FinOps: reduced processing from GB/TB scale to MB in specific workloads.<br>
                • LATAM analytics for Brazil, Chile and Colombia.<br>
                • Business-facing experience with executive dashboards, KPI consolidation and stakeholder communication.
            </div>
            <div class="doc-subtitle">Main Risk Areas</div>
            <div class="doc-section">
                • Airflow / Cloud Composer may be tested; answer honestly using orchestration concepts.<br>
                • Senior title may imply autonomy; position yourself as strong in GCP/BigQuery and open to scope alignment.<br>
                • Recent short tenures should be framed as project-based consulting engagements.
            </div>
            <div class="commentary-box">
                <strong>Positioning Rule:</strong> Do not say you are junior. Say you are open to Senior or Mid-Level scope depending on expectations, while emphasizing hands-on GCP, BigQuery, SQL, Python, data modeling and ETL/ELT experience.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
elif st.session_state.view_mode == "Guide Doc":
    st.markdown(
        """
        <div class="doc-container">
            <div class="doc-title">Interview Strategy — Georgia IT / GCP Remote</div>
            <div class="doc-subtitle">Main Objective</div>
            <div class="doc-section">
                Pass recruiter screening, show technical credibility, protect compensation negotiation and position André as a strong GCP/BigQuery data engineering candidate for an international remote team.
            </div>
            <div class="doc-subtitle">Core Narrative</div>
            <div class="doc-section">
                “I am a data professional with strong analytics engineering and cloud data engineering experience. My recent work with GCP and BigQuery involved data modeling, table/view restructuring, ETL/ELT logic, SQL/Python automation and cloud cost optimization. I am now looking to expand my experience in international data engineering environments.”
            </div>
            <div class="doc-subtitle">Best Evidence to Use</div>
            <div class="doc-section">
                1. <strong>ASICS LATAM:</strong> GCP, BigQuery, 16+ tables/views, FinOps, LATAM analytics.<br>
                2. <strong>Itaú / NTT DATA:</strong> AWS, Athena, SQL, enterprise cloud analytics and agile squads.<br>
                3. <strong>Heineken:</strong> data modeling, product normalization, data quality and KPI reliability.<br>
                4. <strong>Afinz:</strong> 80%+ operational effort reduction and dashboard refresh from 1h30 to 15 minutes.
            </div>
            <div class="doc-subtitle">Negotiation Guidance</div>
            <div class="doc-section">
                Ask for scope before rate. Clarify client, duration, contract model, payment currency, timezone, orchestration tools and expected autonomy.
            </div>
            <div class="commentary-box">
                <strong>Final Interview Message:</strong> “I believe my experience with GCP, BigQuery, SQL, Python, data modeling, ETL/ELT and cloud optimization is well aligned with the role. I am especially interested in contributing to an international data engineering team focused on reliable, scalable and cost-efficient data products.”
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    categories_list = [
        "WHO - Positioning & Fit",
        "TECH - GCP Data Engineering",
        "CASES - STAR Evidence",
        "RISK - Gaps & Strategic Answers",
    ]

    cols = st.columns(len(categories_list))
    for idx, cat_name in enumerate(categories_list):
        with cols[idx]:
            st.markdown(f'<div class="category-header">{cat_name.split(" - ")[0]}</div>', unsafe_allow_html=True)
            cat_items = {k: v for k, v in DATA_MAPPING.items() if v["category"] == cat_name}
            for item_id, item_data in cat_items.items():
                is_active = st.session_state.active_id == item_id
                tag_token = f"[{item_data.get('tag', 'CONTEXT')}] "
                clean_title = item_data.get("title", "Untitled")
                btn_label = f"▸ {tag_token}{clean_title}" if is_active else f"{tag_token}{clean_title}"
                if st.button(btn_label, key=f"btn_{item_id}"):
                    st.session_state.active_id = item_id
                    st.rerun()

    st.markdown("<div style='margin-top:0.2rem; border-top:1px solid #e9ecef; margin-bottom:0.35rem;'></div>", unsafe_allow_html=True)
    active_data = DATA_MAPPING.get(st.session_state.active_id, DATA_MAPPING[1])

    col_out1, col_out2 = st.columns([0.5, 0.5])
    with col_out1:
        st.markdown(
            f"""
            <div class="response-box">
                <span class="small-label" style="color:#117a65;">Golden Bridge — natural answer:</span><br>
                <strong style="font-size:12.8px; color:#2c3e50; line-height:1.25;">“{active_data.get('bridge', '')}”</strong>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown(
            f"""
            <div class="followup-box">
                <span class="small-label" style="color:#2c3e50;">Context / Deep Dive:</span><br>
                <p style="font-size:12.2px; color:#34495e; line-height:1.35;">{active_data.get('followup', '')}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown(
            f"""
            <div class="growth-box">
                <strong style="color:#d35400; text-transform:uppercase; font-size:9.5px;">Strategic Direction:</strong><br>
                <p style="color:#ba4a00; font-size:11.8px; line-height:1.35; margin-top:2px;">{active_data.get('growth', '')}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown(
            f"""
            <div class="match-box">
                <strong style="color:#2980b9; text-transform:uppercase; font-size:9.5px;">Match Objective:</strong><br>
                <p style="color:#1f618d; font-size:11.8px; line-height:1.35; margin-top:2px;">{active_data.get('match', '')}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_out2:
        bullets_html = "".join(
            f'<p style="font-size:12px; color:#2c3e50; line-height:1.35; margin-bottom:4px !important;">• {b}</p>'
            for b in active_data.get("bullets", [])
        )
        st.markdown(
            f"""
            <div class="bullet-container-box">
                <span style="color:#2c3e50; font-size:10.5px; font-weight:bold; text-transform:uppercase; display:block; margin-bottom:6px;">Supporting Core Arguments:</span>
                {bullets_html}
                <p style='font-size:10.5px; color:#7f8c8d; margin-top:5px !important;'><strong>Baseline Case:</strong> {active_data.get('case', '')}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        qa_html_items = ""
        for qa in active_data.get("qa_responses", []):
            qa_html_items += f"""
            <div class="qa-item">
                <strong style="font-size:11.7px; color:#1b4f72; display:block; line-height:1.3;">Q: {qa.get('q', '')}</strong>
                <p style="font-size:11.7px; color:#154360; line-height:1.35; margin-top:2px !important;"><strong>A:</strong> {qa.get('a', '')}</p>
            </div>
            """
        st.markdown(
            f"""
            <div class="qa-container-box">
                <span style="color:#154360; font-size:10.5px; font-weight:bold; text-transform:uppercase; display:block; margin-bottom:6px;">⚡ Tough Interview Q&A Simulator:</span>
                {qa_html_items}
            </div>
            """,
            unsafe_allow_html=True,
        )

