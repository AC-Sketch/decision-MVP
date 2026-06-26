import streamlit as st

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA E DESIGN SYSTEM (UX/UI AVANÇADA)
# ==============================================================================
st.set_page_config(
    page_title="Plataforma Objetivo Premium - Química 1º Ano",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Design System robusto simulando o padrão de apostilas físicas premium
CUSTOM_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;500&family=Inter:wght@300;400;500;600;700&display=swap');
    
    html, body, [data-testid="stSidebar"] {
        font-family: 'Inter', sans-serif;
        color: #1E293B;
        background-color: #F8FAFC;
    }
    
    /* Box Textual de Alta Densidade */
    .apostila-bloco {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-left: 6px solid #1E3A8A;
        padding: 30px;
        border-radius: 12px;
        margin-bottom: 25px;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.02);
    }
    
    .alerta-academico {
        background-color: #FFFBEB;
        border-left: 6px solid #D97706;
        padding: 20px;
        border-radius: 8px;
        margin: 20px 0;
        font-size: 0.95em;
    }
    
    .titulo-capitulo {
        color: #1E3A8A;
        font-weight: 700;
        border-bottom: 2px solid #E2E8F0;
        padding-bottom: 10px;
        margin-top: 25px;
        margin-bottom: 20px;
    }
    
    .subtitulo-tecnico {
        color: #0D9488;
        font-weight: 600;
        margin-top: 15px;
        margin-bottom: 10px;
    }

    /* Estilização de Tabelas de Dados */
    .tabela-quimica {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
    }
    .tabela-quimica th {
        background-color: #1E3A8A;
        color: white;
        padding: 12px;
        text-align: left;
        border: 1px solid #CBD5E1;
    }
    .tabela-quimica td {
        padding: 12px;
        border: 1px solid #CBD5E1;
    }
    .tabela-quimica tr:nth-child(even) {
        background-color: #F8FAFC;
    }
    
    /* Elementos do Organograma Cognitivo */
    .node-master {
        background: linear-gradient(135deg, #1E3A8A 0%, #0F172A 100%);
        color: white; padding: 22px; border-radius: 12px; text-align: center; font-weight: 700; box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .node-branch {
        background: linear-gradient(135deg, #0D9488 0%, #115E59 100%);
        color: white; padding: 15px; border-radius: 10px; text-align: center; font-weight: 600; margin-top: 15px;
    }
    .node-leaf {
        background-color: #FFFFFF; color: #334155; padding: 12px; border: 1px solid #E2E8F0; border-left: 4px solid #7C3AED; border-radius: 8px; margin-top: 8px; font-size: 0.88em;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ==============================================================================
# PROTOTIPAGEM DE ESTADOS PERSISTENTES (SESSION STATE)
# ==============================================================================
if 'simulado_finalizado' not in st.session_state:
    st.session_state.simulado_finalizado = False

# ==============================================================================
# ARQUITETURA DA SIDEBAR (CONTEÚDO PROGRAMÁTICO INTEGRAL)
# ==============================================================================
st.sidebar.markdown("# 👀 **Plataforma Objetivo**\n*Sistema de Ensino Avançado — 2026*")
st.sidebar.markdown("---")

capitulo = st.sidebar.radio(
    "Sumário da Apostila Digital:",
    [
        "📖 Cap 1: Matéria, Sistemas e Fenômenos",
        "🌡️ Cap 2: Estados e Curvas de Transição",
        "⚛️ Cap 3: Atomística e Estrutura Atômica",
        "🧠 Módulo 4: Organogramas de Revisão",
        "📊 Módulo 5: Simulador Laboratorial",
        "📝 Módulo 6: Caderno de Exercícios (OBR)",
        "📚 Módulo 7: Fontes Acadêmicas & Vestibulares"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption("Suporte Pedagógico: Área de Ciências da Natureza.")

# ==============================================================================
# CAPÍTULO 1: MATÉRIA, SISTEMAS E FENÔMENOS
# ==============================================================================
if capitulo == "📖 Cap 1: Matéria, Sistemas e Fenômenos":
    st.title("🔬 Capítulo 1: Fundamentos da Matéria e Classificação de Sistemas")
    st.markdown("---")
    
    st.markdown("""
    <div class="apostila-bloco">
        <h3 class="titulo-capitulo">1. Conceituação Científica da Matéria</h3>
        <p>A Química analisa a matéria em nível molecular e atômico. Definimos <b>Matéria</b> como tudo o que possui massa inercial, ocupa lugar geométrico no espaço (volume) e é passível de ser quantificado.</p>
        
        <h4 class="subtitulo-tecnico">1.1 Matéria, Corpo e Objeto</h4>
        <ul>
            <li><b>Matéria:</b> O conceito bruto. Exemplo: O ouro, a madeira, o ferro.</li>
            <li><b>Corpo:</b> Uma porção limitada e definida da matéria. Exemplo: Uma barra de ouro, um tronco de madeira.</li>
            <li><b>Objeto:</b> Um corpo trabalhado e manufaturado para desempenhar uma utilidade prática específica ao ser humano. Exemplo: Um anel de ouro, uma cadeira de madeira.</li>
        </ul>
        
        <h4 class="subtitulo-tecnico">1.2 Propriedades Gerais da Matéria (Comuns a todos os corpos)</h4>
        <p>Essas propriedades não servem para caracterizar isoladamente uma substância pura, pois manifestam-se em qualquer tipo de amostra:</p>
        <ol>
            <li><b>Massa:</b> Medida quantitativa da inércia do corpo.</li>
            <li><b>Extensão (Volume):</b> Espaço físico tridimensional ocupado pelo corpo.</li>
            <li><b>Impenetrabilidade:</b> Princípio mecânico onde dois corpos não ocupam o mesmo espaço tridimensional ao mesmo tempo.</li>
            <li><b>Divisibilidade:</b> Capacidade da matéria de ser fracionada em porções progressivamente menores sem perder suas propriedades originais (até o limite molecular).</li>
            <li><b>Porosidade:</b> Presença de espaços vazios intermoleculares, mesmo nas estruturas aparentemente mais maciças.</li>
        </ol>
    </div>
    
    <div class="apostila-bloco">
        <h3 class="titulo-capitulo">2. Diferenciação Analítica de Sistemas Químicos</h3>
        <p>Um sistema é a porção do universo físico isolada para estudo ou observação experimental. Classificamos os sistemas com base em seu aspecto visual microscópico e macroscópico:</p>
        
        <h4 class="subtitulo-tecnico">2.1 Sistemas Homogêneos e Heterogêneos</h4>
        <ul>
            <li><b>Sistemas Homogêneos:</b> Apresentam aspecto visual perfeitamente uniforme e contínuo ao longo de toda a sua extensão, possuindo uma única fase. Exemplo: Água destilada, mistura de água e sal totalmente solubilizado, ar atmosférico filtrado.</li>
            <li><b>Sistemas Heterogêneos:</b> Apresentam descontinuidade visual, exibindo duas ou mais fases nitidamente separáveis por métodos físicos. Exemplo: Água e óleo, água líquida e cubos de gelo, granito.</li>
        </ul>
        
        <div class="alerta-academico">
            <b>⚠️ Pegadinha de Vestibular Recorrente:</b> O sangue e o leite parecem homogêneos a olho nu, mas ao microscópio revelam-se sistemas heterogêneos (o leite possui gotículas de gordura suspensas e o sangue possui hemácias e plaquetas em suspensão plasmática).
        </div>
    </div>
    """, unsafe_allow_html=True)

# ==============================================================================
# CAPÍTULO 2: ESTADOS E CURVAS DE TRANSIÇÃO
# ==============================================================================
elif capitulo == "🌡️ Cap 2: Estados e Curvas de Transição":
    st.title("🌡️ Capítulo 2: Estados de Agregação da Matéria e Análise Térmica")
    st.markdown("---")
    
    st.markdown("""
    <div class="apostila-bloco">
        <h3 class="titulo-capitulo">1. Teoria Cinética dos Gases e Estados Físicos</h3>
        <p>A matéria se organiza dinamicamente em três estados principais, governados pelo balanço fino entre a energia térmica vibracional (Repulsão) e as forças de atração intermoleculares (Coesão).</p>
        
        <table class="tabela-quimica">
            <thead>
                <tr>
                    <th>Estado Físico</th>
                    <th>Forma Geométrica</th>
                    <th>Volume Espacial</th>
                    <th>Arranjo Microscópico</th>
                </tr>
            </thead>
            <tbody>
                <tr>
                    <td><b>Sólido</b></td>
                    <td>Constante / Rígida</td>
                    <td>Constante / Definido</td>
                    <td>Retículo cristalino altamente ordenado. Forças de coesão predominam.</td>
                </tr>
                <tr>
                    <td><b>Líquido</b></td>
                    <td>Variável (Adota o vaso)</td>
                    <td>Constante</td>
                    <td>Partículas com liberdade de translação e rotação moderada. Coesão equilibrada.</td>
                </tr>
                <tr>
                    <td><b>Gasoso</b></td>
                    <td>Variável (Expansível)</td>
                    <td>Variável (Ocupa todo o recipiente)</td>
                    <td>Movimento molecular caótico, colisões elásticas e forças de atração nulas.</td>
                </tr>
            </tbody>
        </table>
    </div>

    <div class="apostila-bloco">
        <h3 class="titulo-capitulo">2. Gráficos Térmicos Especiais: Misturas Azeotrópicas e Eutéticas</h3>
        <p>Enquanto uma <b>Substância Pura</b> possui tanto o Ponto de Fusão (PF) quanto o Ponto de Ebulição (PE) constantes durante a mudança de fase, as misturas comportam-se de formas específicas:</p>
        <ul>
            <li><b>Mistura Comum:</b> Apresenta variação de temperatura tanto na fusão quanto na ebulição (os patamares do gráfico são inclinados).</li>
            <li><b>Mistura Eutética:</b> Comporta-se como substância pura estritamente durante a <b>Fusão</b> (temperatura constante), mas varia na ebulição. Exemplo: Liga metálica de Solda (Estanho + Chumbo).</li>
            <li><b>Mistura Azeotrópica:</b> Comporta-se como substância pura estritamente durante a <b>Ebulição</b> (temperatura constante), mas varia na fusão. Exemplo: Álcool hidratado (96% Etanol + 4% Água).</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

# ==============================================================================
# CAPÍTULO 3: ATOMÍSTICA E ESTRUTURA ATÔMICA
# ==============================================================================
elif capitulo == "⚛️ Cap 3: Atomística e Estrutura Atômica":
    st.title("⚛️ Capítulo 3: Evolução Cronológica dos Modelos Atômicos Clássicos")
    st.markdown("---")
    
    st.markdown("""
    <div class="apostila-bloco">
        <h3 class="titulo-capitulo">1. John Dalton (1808) — Consolidação Ponderal</h3>
        <p>Dalton resgatou o termo grego 'átomo' aplicando formalismo matemático atrelado às Leis de Lavoisier (Conservação de Massas) e Proust (Proporções Definidas).</p>
        <p><b>Postulados Centrais:</b> Os átomos são esferas maciças, totalmente neutras, indivisíveis e indestrutíveis. Elementos diferentes possuem massas diferentes. Reações químicas representam apenas reorganizações geométricas desses componentes estáveis.</p>
        
        <h3 class="titulo-capitulo">2. J.J. Thomson (1897) — A Descoberta da Natureza Elétrica</h3>
        <p>Ao submeter gases a altas voltagens dentro de tubos de vidro sob vácuo (Ampola de Crookes), Thomson percebeu um fluxo de radiação partindo do polo negativo (cátodo) ao positivo (ânodo). Ele batizou essas partículas de <b>elétrons</b>.</p>
        <p><b>O Modelo:</b> O átomo não é indestrutível. Ele consiste em um fluido contínuo positivo dotado de elétrons incrustados de forma homogênea para garantir a neutralidade elétrica global (Modelo do Pudim de Passas).</p>
    </div>

    <div class="apostila-bloco">
        <h3 class="titulo-capitulo">3. Ernest Rutherford (1911) — O Modelo Nucleado</h3>
        <p>Rutherford disparou partículas Alpha (núcleos de Hélio dotados de carga positiva) contra uma lâmina ultrafina de ouro Puro. O resultado contradisse o modelo de Thomson:</p>
        <ul>
            <li>99% das partículas atravessaram a folha sem sofrer nenhuma alteração vetorial.</li>
            <li>Uma fração ínfima sofreu desvios angulares severos ou foi refletida de volta.</li>
        </ul>
        <p><b>Conclusão Científica:</b> O átomo é composto por um imenso espaço vazio chamado <b>Eletrosfera</b>, onde orbitam os elétrons, e uma densa e minúscula região central positiva chamada <b>Núcleo</b>, que concentra quase toda a massa atômica.</p>
    </div>
    """, unsafe_allow_html=True)

# ==============================================================================
# MÓDULO 4: ORGANOGRAMAS DE REVISÃO (MAPA MENTAL)
# ==============================================================================
elif capitulo == "🧠 Módulo 4: Organogramas de Revisão":
    st.title("🧠 Módulo Cognitivo Interativo — Mapas Conceituais")
    st.write("Estude conectando os nós hierárquicos montados em blocos estruturados.")
    st.markdown("---")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown('<div class="node-master">PROPRIEDADES QUÍMICAS</div>', unsafe_allow_html=True)
        st.markdown('<div class="node-branch">Combustão</div>', unsafe_allow_html=True)
        st.markdown('<div class="node-leaf">Capacidade de reagir com O₂ gerando calor.</div>', unsafe_allow_html=True)
        st.markdown('<div class="node-branch">Oxidação</div>', unsafe_allow_html=True)
        st.markdown('<div class="node-leaf">Perda de elétrons para um agente oxidante.</div>', unsafe_allow_html=True)
        
    with col2:
        st.markdown('<div class="node-master" style="background: linear-gradient(135deg, #0D9488 0%, #115E59 100%);">MISTURAS ESPECIAIS</div>', unsafe_allow_html=True)
        st.markdown('<div class="node-branch" style="background-color: #1E3A8A;">Eutética</div>', unsafe_allow_html=True)
        st.markdown('<div class="node-leaf">Ponto de Fusão fixo e constante.</div>', unsafe_allow_html=True)
        st.markdown('<div class="node-branch" style="background-color: #1E3A8A;">Azeotrópica</div>', unsafe_allow_html=True)
        st.markdown('<div class="node-leaf">Ponto de Ebulição fixo e constante.</div>', unsafe_allow_html=True)
        
    with col3:
        st.markdown('<div class="node-master" style="background: linear-gradient(135deg, #7C3AED 0%, #4C1D95 100%);">MODELOS HISTÓRICOS</div>', unsafe_allow_html=True)
        st.markdown('<div class="node-branch" style="background-color: #6D28D9;">Dalton</div>', unsafe_allow_html=True)
        st.markdown('<div class="node-leaf">Maciço, indestrutível, bola de bilhar.</div>', unsafe_allow_html=True)
        st.markdown('<div class="node-branch" style="background-color: #6D28D9;">Thomson</div>', unsafe_allow_html=True)
        st.markdown('<div class="node-leaf">Divisível, elétrons, pudim de passas.</div>', unsafe_allow_html=True)

# ==============================================================================
# MÓDULO 5: SIMULADOR LABORATORIAL NATIVO (À PROVA DE FALHAS)
# ==============================================================================
elif capitulo == "📊 Módulo 5: Simulador Laboratorial":
    st.title("📊 Simulador Clínico de Mudanças Físicas")
    st.write("Ajuste as variáveis e analise os gráficos de aquecimento e resfriamento gerados nativamente pelo mecanismo interno.")
    st.markdown("---")
    
    col_l, col_r = st.columns([1, 2])
    
    with col_l:
        st.subheader("Painel de Controle Térmico")
        fusao = st.slider("Ponto de Fusão Alvo (°C)", min_value=-10, max_value=30, value=0, step=1)
        ebulicao = st.slider("Ponto de Ebulição Alvo (°C)", min_value=65, max_value=115, value=100, step=1)
        
        if fusao >= ebulicao:
            st.error("Inconsistência Térmica: O ponto de fusão deve ser menor que o ponto de ebulição!")
            fusao, ebulicao = 0, 100
            
        st.markdown("""
        ### Diagnóstico da Amostra:
        * **Patamar Inferior:** Coexistência Sólido + Líquido.
        * **Patamar Superior:** Coexistência Líquido + Gasoso.
        """)
        
    with col_r:
        # Gráfico de Aquecimento Nativo
        dados_aquecimento = {
            "Temperatura de Aquecimento (°C)": [fusao - 15, fusao, fusao, (fusao+ebulicao)/2, ebulicao, ebulicao, ebulicao + 15]
        }
        st.write("#### Curva de Aquecimento Experimental")
        st.line_chart(dados_aquecimento)
        
        # Gráfico de Resfriamento Nativo
        dados_resfriamento = {
            "Temperatura de Resfriamento (°C)": [ebulicao + 15, ebulicao, ebulicao, (fusao+ebulicao)/2, fusao, fusao, fusao - 15]
        }
        st.write("#### Curva de Resfriamento Experimental")
        st.line_chart(dados_resfriamento)

# ==============================================================================
# MÓDULO 6: CADERNO DE EXERCÍCIOS (OBR)
# ==============================================================================
elif capitulo == "📝 Módulo 6: Caderno de Exercícios (OBR)":
    st.title("📝 Caderno de Exercícios Avaliativos — Foco Vestibular")
    st.write("Selecione as alternativas corretas e clique no botão de validação pedagógica.")
    st.markdown("---")
    
    st.markdown("##### 1. (Objetivo) Um sistema fechado contendo água salgada, três pedras de granito e vapor d'água retido apresenta quantas fases e quantos componentes químicos no total?")
    ex1 = st.radio("Selecione:", ["3 fases e 2 componentes", "4 fases e 3 componentes", "5 fases e 4 componentes", "5 fases e 5 componentes"], key="ex_1")
    
    st.markdown("---")
    st.markdown("##### 2. (Fuvest) Qual dos experimentos abaixo serviu como base empírica imediata para a derrubada do modelo atômico de Dalton, comprovando a divisibilidade da matéria?")
    ex2 = st.radio("Selecione:", ["O bombardeio de raios alfa em folhas de ouro.", "A análise de espectros de emissão de hidrogênio.", "O desvio de raios catódicos por campos elétricos.", "A organização dos elementos em oitavas por Newlands."], key="ex_2")
    
    st.markdown("---")
    st.markdown("##### 3. (Unicamp) Uma liga metálica funde-se de modo progressivo iniciando a liquefação em 210°C e terminando em 225°C. Contudo, seu ponto de ebulição estabiliza-se rigorosamente em 890°C. Esse material se trata de uma:")
    ex3 = st.radio("Selecione:", ["Substância composta pura", "Mistura Eutética", "Mistura Azeotrópica", "Mistura Comum heterogênea"], key="ex_3")
    
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚀 Submeter Respostas ao Servidor"):
        st.session_state.simulado_finalizado = True
        
    if st.session_state.simulado_finalizado:
        st.markdown("### 📊 Relatório Estatístico do Aluno")
        pontos = 0
        
        # Correção Q1 (Granito tem 3 fases, água salgada 1 fase, vapor 1 fase = 5 fases)
        if ex1 == "5 fases e 4 componentes":
            pontos += 1
            st.success("✓ Questão 1: Correta! (Granito possui 3 minerais/fases, água salgada é 1 fase e vapor é outra fase = 5 fases).")
        else:
            st.error("✕ Questão 1: Incorreta. O sistema contém 5 fases distintas.")
            
        # Correção Q2
        if ex2 == "O desvio de raios catódicos por campos elétricos.":
            pontos += 1
            st.success("✓ Questão 2: Correta! Thomson utilizou raios catódicos provando a natureza do elétron.")
        else:
            st.error("✕ Questão 2: Incorreta. Os raios catódicos derrubaram a indivisibilidade de Dalton.")
            
        # Correção Q3
        if ex3 == "Mistura Azeotrópica":
            pontos += 1
            st.success("✓ Questão 3: Correta! Ponto de ebulição constante e fusão variável define misturas azeotrópicas.")
        else:
            st.error("✕ Questão 3: Incorreta. O comportamento descrito é o de uma Mistura Azeotrópica.")
            
        nota_final = (pontos / 3) * 100
        st.metric("Aproveitamento Acadêmico", f"{nota_final:.1f}%", f"{pontos} de 3 questões corretas")
        st.progress(nota_final / 100)

# ==============================================================================
# MÓDULO 7: FONTES ACADÊMICAS & VESTIBULARES
# ==============================================================================
elif capitulo == "📚 Módulo 7: Fontes Acadêmicas & Vestibulares":
    st.title("📚 Referências Bibliográficas e Direcionamento de Estudos")
    st.markdown("---")
    
    st.markdown("""
    <div class="apostila-section">
        <h4 style="color: #1E3A8A; font-weight:bold;">📖 Bibliografia Recomendada para Vestibulares</h4>
        <ul>
            <li><b>FELTRE, Ricardo.</b> <i>Química Geral - Volume 1.</i> Editora Moderna, 2014. (Abordagem analítica clássica).</li>
            <li><b>ATKINS, Peter; JONES, Loretta.</b> <i>Princípios de Química: Questionando a Vida Moderna.</i> Bookman.</li>
            <li><b>USBERCO, João; SALVADOR, Edgard.</b> <i>Química Geral.</i> Editora Saraiva, 2018.</li>
        </ul>
    </div>
    
    <div class="apostila-section" style="border-left-color: #0D9488;">
        <h4 style="color: #0D9488; font-weight:bold;">🌐 Links de Extensão e Exercícios Online</h4>
        <ul>
            <li><b>Portal Objetivo Oficial:</b> <a href="https://www.objetivo.br" target="_blank">Acesse a Área do Aluno</a> para resoluções comentadas.</li>
            <li><b>Khan Academy em Português:</b> Módulos interativos de estrutura atômica avançada e mecânica quântica primitiva.</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
