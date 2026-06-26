import streamlit as st
import plotly.graph_objects as go
import numpy as np

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA & CONSTANTES DE DESIGN
# ==============================================================================
st.set_page_config(
    page_title="Apostila Virtual de Química - 1º Ano",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Paleta de cores: Azul Científico (#1E3A8A), Verde-Água (#0D9488), Grafite (#374151)
CUSTOM_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');
    
    html, body, [data-testid="stSidebar"] {
        font-family: 'Inter', sans-serif;
    }
    
    /* Customização de Cards */
    .chem-card {
        background-color: #F8FAFC;
        border-left: 5px solid #0D9488;
        padding: 20px;
        border-radius: 8px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    
    .chem-header {
        color: #1E3A8A;
        font-weight: 700;
        margin-bottom: 10px;
    }
    
    /* Elementos visuais do Mapa Mental */
    .mindmap-node {
        background-color: #1E3A8A;
        color: white;
        padding: 15px;
        border-radius: 12px;
        text-align: center;
        font-weight: 600;
        margin: 10px 0;
    }
    .mindmap-subnode {
        background-color: #0D9488;
        color: white;
        padding: 10px;
        border-radius: 8px;
        text-align: center;
        font-size: 14px;
        margin: 5px 0;
    }
    .mindmap-connection {
        text-align: center;
        color: #94A3B8;
        font-size: 20px;
        line-height: 1;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ==============================================================================
# INICIALIZAÇÃO DO ESTADO GLOBAL (SESSION STATE)
# ==============================================================================
if 'quiz_submitted' not in st.session_state:
    st.session_state.quiz_submitted = False
if 'user_answers' not in st.session_state:
    st.session_state.user_answers = {}

# ==============================================================================
# ESTRUTURA DE NAVEGAÇÃO DA SIDEBAR
# ==============================================================================
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3022/3022271.png", width=80)
st.sidebar.title("Objetivo Virtual")
st.sidebar.subtitle("Química - 1º Ano Ensino Médio")

menu = st.sidebar.radio(
    "Navegue pelos Módulos:",
    [
        "📖 Conteúdo & Explicações",
        "🧠 Mapa Mental",
        "📊 Gráficos Interativos",
        "📝 Teste de Nivelamento",
        "📚 Referências & Fontes"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption("Desenvolvido para alta performance acadêmica. © 2026")

# ==============================================================================
# MÓDULO 1: CONTEÚDO & EXPLICAÇÕES
# ==============================================================================
if menu == "📖 Conteúdo & Explicações":
    st.title("🧪 Módulo 1: Introdução à Química e Atomística")
    st.markdown("---")
    
    tabs = st.tabs(["1. Matéria e Energia", "2. Estados Físicos", "3. Atomística Básica"])
    
    with tabs[0]:
        st.subheader("O que é Matéria?")
        st.markdown("""
        <div class="chem-card">
            <h4 class="chem-header">Definição Fundamental</h4>
            <p><b>Matéria</b> é tudo aquilo que tem massa e ocupa lugar no espaço. Ela é constituída por pequenas partículas chamadas átomos. A energia, por sua vez, é a capacidade de realizar trabalho ou transferir calor.</p>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            st.info("**Propriedades Gerais:** Massa, volume, impenetrabilidade, divisibilidade.")
        with col2:
            st.success("**Propriedades Específicas:** Ponto de fusão, ponto de ebulição, densidade, solubilidade.")

    with tabs[1]:
        st.subheader("Estados Físicos e Mudanças de Fase")
        st.write("A matéria se apresenta principalmente em três estados físicos fundamentais:")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("### 🧊 Sólido\nForma e volume constantes. Partículas altamente organizadas.")
        with col2:
            st.markdown("### 💧 Líquido\nForma variável (adota o recipiente) e volume constante. Forças de atração moderadas.")
        with col3:
            st.markdown("### 💨 Gasoso\nForma e volume variáveis. Alta desorganização e grande liberdade de movimento.")
            
        st.markdown("""
        <div class="chem-card" style="border-left-color: #1E3A8A;">
            <h4 class="chem-header">💡 Fique Atento às Transições!</h4>
            <ul>
                <li><b>Fusão:</b> Sólido ➔ Líquido</li>
                <li><b>Vaporização:</b> Líquido ➔ Gasoso (Evaporação, Ebulição ou Calefação)</li>
                <li><b>Sublimação:</b> Sólido ➔ Gasoso (direto, ex: Naftalina)</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    with tabs[2]:
        st.subheader("Evolução dos Modelos Atômicos")
        st.markdown("A compreensão do átomo evoluiu ao longo do tempo. Veja os principais marcos cobrados no Colégio Objetivo:")
        
        st.markdown("""
        <div class="chem-card">
            <h4 class="chem-header">⚫ 1. Dalton (1808) - Modelo da Bola de Bilhar</h4>
            <p>O átomo é uma esfera maciça, indivisível, indestrutível e neutra.</p>
        </div>
        <div class="chem-card">
            <h4 class="chem-header">🍇 2. Thomson (1898) - Modelo do Pudim de Passas</h4>
            <p>Descoberta do elétron. O átomo é uma esfera de carga positiva recheada de elétrons de carga negativa.</p>
        </div>
        <div class="chem-card">
            <h4 class="chem-header">🪐 3. Rutherford (1911) - Modelo Planetário</h4>
            <p>O átomo possui um núcleo central denso e positivo, e uma eletrosfera imensa onde orbitam os elétrons.</p>
        </div>
        """, unsafe_allow_html=True)

# ==============================================================================
# MÓDULO 2: MAPA MENTAL
# ==============================================================================
elif menu == "🧠 Mapa Mental":
    st.title("🧠 Mapa Mental Dinâmico")
    st.write("Utilize esta estrutura visual para revisar rapidamente os conceitos fundamentais antes dos exercícios.")
    st.markdown("---")
    
    # Renderização de uma estrutura de árvore usando blocos HTML/CSS construídos lado a lado
    col_center = st.columns([1, 2, 1])[1]
    
    with col_center:
        st.markdown('<div class="mindmap-node">PROPRIEDADES DA MATÉRIA</div>', unsafe_allow_html=True)
        st.markdown('<div class="mindmap-connection">▼</div>', unsafe_allow_html=True)
        
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="mindmap-node">Gerais</div>', unsafe_allow_html=True)
        st.markdown('<div class="mindmap-subnode">Massa</div>', unsafe_allow_html=True)
        st.markdown('<div class="mindmap-subnode">Volume</div>', unsafe_allow_html=True)
        st.markdown('<div class="mindmap-subnode">Impenetrabilidade</div>', unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="mindmap-node">Específicas</div>', unsafe_allow_html=True)
        st.markdown('<div class="mindmap-subnode">Químicas (Combustão)</div>', unsafe_allow_html=True)
        st.markdown('<div class="mindmap-subnode">Físicas (Ponto de Fusão/Ebulição)</div>', unsafe_allow_html=True)
        st.markdown('<div class="mindmap-subnode">Organolépticas (Sabor, Odor)</div>', unsafe_allow_html=True)

# ==============================================================================
# MÓDULO 3: GRÁFICOS INTERATIVOS
# ==============================================================================
elif menu == "📊 Gráficos Interativos":
    st.title("📊 Simulador Científico: Curva de Aquecimento")
    st.write("Modifique as propriedades físicas abaixo e observe como o gráfico de mudança de fase da substância pura se comporta.")
    st.markdown("---")
    
    col_input, col_graph = st.columns([1, 2])
    
    with col_input:
        st.subheader("Configurações Térmicas")
        pf = st.slider("Ponto de Fusão (°C)", min_value=-50, max_value=50, value=0, step=5)
        pe = st.slider("Ponto de Ebulição (°C)", min_value=60, max_value=150, value=100, step=5)
        
        if pf >= pe:
            st.error("Erro: O Ponto de Fusão não pode ser maior ou igual ao Ponto de Ebulição!")
            pf = 0
            pe = 100
            
        st.info("""
        **Entendendo o Gráfico:**
        * Nos patamares retos, a temperatura não muda pois a energia é usada para romper as ligações intermoleculares (Calor Latente).
        * Nas subidas, a energia aumenta a agitação cinética das moléculas (Calor Sensível).
        """)

    with col_graph:
        # Geração dos pontos da curva de aquecimento
        tempo = np.array([0, 10, 20, 30, 40, 50, 60])
        temp = np.array([pf - 20, pf, pf, (pf+pe)/2, pe, pe, pe + 20])
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=tempo, y=temp,
            mode='lines+markers',
            line=dict(color='#0D9488', width=4),
            marker=dict(size=8, color='#1E3A8A'),
            name='Curva de Aquecimento'
        ))
        
        fig.add_annotation(x=15, y=pf, text="Fusão (Sólido + Líquido)", showarrow=True, arrowhead=1, yshift=10)
        fig.add_annotation(x=45, y=pe, text="Ebulição (Líquido + Vapor)", showarrow=True, arrowhead=1, yshift=10)
        
        fig.update_layout(
            title="Curva de Aquecimento da Substância Pureza",
            xaxis_title="Tempo de Aquecimento (min)",
            yaxis_title="Temperatura (°C)",
            plot_bgcolor="rgba(248,250,252,1)",
            margin=dict(l=20, r=20, t=40, b=20)
        )
        st.plotly_chart(fig, use_container_width=True)

# ==============================================================================
# MÓDULO 4: TESTE DE NIVELAMENTO & EXERCÍCIOS
# ==============================================================================
elif menu == "📝 Teste de Nivelamento":
    st.title("📝 Teste de Nivelamento OBR (Padrão Objetivo)")
    st.write("Responda às questões e clique em 'Enviar Respostas' no final para calcular seu rendimento.")
    st.markdown("---")
    
    # Banco de Questões Estático
    questions = [
        {
            "id": "q1",
            "text": "1. Um sistema formado por água líquida, cubos de gelo e óleo apresenta quantas fases?",
            "options": ["1 fase", "2 fases", "3 fases", "4 fases"],
            "correct": "3 fases"
        },
        {
            "id": "q2",
            "text": "2. Quem propôs o primeiro modelo atômico baseado em fatos experimentais, conhecido como 'bola de bilhar'?",
            "options": ["Thomson", "Rutherford", "Dalton", "Bohr"],
            "correct": "Dalton"
        },
        {
            "id": "q3",
            "text": "3. A passagem direta do estado sólido para o estado gasoso sem passar pelo estado líquido é chamada de:",
            "options": ["Sublimação", "Fusão", "Condensação", "Evaporação"],
            "correct": "Sublimação"
        }
    ]
    
    # Coletando respostas de forma segura sem resetar o estado
    for q in questions:
        st.markdown(f"#### {q['text']}")
        current_selection = st.radio(
            "Escolha a alternativa correta:", 
            q["options"], 
            key=f"radio_{q['id']}"
        )
        st.session_state.user_answers[q["id"]] = current_selection
        st.write("")

    if st.button("🚀 Enviar Respostas"):
        st.session_state.quiz_submitted = True
        
    if st.session_state.quiz_submitted:
        score = 0
        st.markdown("### 📊 Resultado do Seu Desempenho")
        
        for q in questions:
            user_ans = st.session_state.user_answers.get(q["id"])
            if user_ans == q["correct"]:
                score += 1
                st.success(f"✓ {q['text']} - Você acertou! Resposta: {q['correct']}")
            else:
                st.error(f"✕ {q['text']} - Errado. Você escolheu: {user_ans}. Correta: {q['correct']}")
                
        pct = (score / len(questions)) * 100
        st.metric(label="Sua Nota Final", value=f"{pct:.1f}%", delta=f"{score} de {len(questions)} corretas")
        st.progress(pct / 100)

# ==============================================================================
# MÓDULO 5: REFERÊNCIAS & FONTES
# ==============================================================================
elif menu == "📚 Referências & Fontes":
    st.title("📚 Fontes de Estudo Complementares")
    st.write("Aprofunde seus conhecimentos utilizando materiais externos de alta qualidade recomendados pela coordenação:")
    st.markdown("---")
    
    st.markdown("""
    <div class="chem-card">
        <h4 class="chem-header">🔗 Portais Educacionais & Exercícios</h4>
        <ul>
            <li><b>Portal Objetivo:</b> <a href="https://www.objetivo.br" target="_blank">Acesse a área do aluno oficial</a> para listas extras.</li>
            <li><b>Brasil Escola - Química Geral:</b> Artigos aprofundados sobre atomística e tabela periódica.</li>
        </ul>
    </div>
    <div class="chem-card" style="border-left-color: #1E3A8A;">
        <h4 class="chem-header">📺 Canais Recomendados no YouTube</h4>
        <ul>
            <li><b>Química em Ação (Prof. Paulo Valim):</b> Playlists completas com foco em vestibulares e base forte.</li>
            <li><b>Professor Gabriel Cabral:</b> Didática descontraída, ideal para fixação rápida do 1º ano.</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
