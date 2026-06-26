import streamlit as st
import plotly.graph_objects as go

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA E DESIGN SYSTEM
# ==============================================================================
st.set_page_config(
    page_title="Plataforma Objetivo - Química 1º Ano",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização em CSS para emular uma apostila digital de alta performance
CUSTOM_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    
    html, body, [data-testid="stSidebar"] {
        font-family: 'Inter', sans-serif;
        color: #1E293B;
    }
    
    .apostila-section {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-left: 6px solid #1E3A8A;
        padding: 25px;
        border-radius: 12px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.02);
    }
    
    .titulo-quimica {
        color: #1E3A8A;
        font-weight: 700;
        border-bottom: 2px solid #E2E8F0;
        padding-bottom: 8px;
        margin-top: 10px;
        margin-bottom: 15px;
    }
    
    .mm-box {
        background: linear-gradient(135deg, #1E3A8A 0%, #0F172A 100%);
        color: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        font-weight: bold;
    }
    
    .mm-sub {
        background-color: #0D9488;
        color: white;
        padding: 12px;
        border-radius: 8px;
        text-align: center;
        margin-top: 10px;
        font-size: 0.9em;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ==============================================================================
# GERENCIAMENTO DE ESTADO SEGURO (SESSION STATE)
# ==============================================================================
if 'quiz_completado' not in st.session_state:
    st.session_state.quiz_completado = False

# ==============================================================================
# MENU LATERAL DE NAVEGAÇÃO
# ==============================================================================
st.sidebar.markdown("# **Objetivo Virtual**\n*Material Didático de Química*")
st.sidebar.markdown("---")

modulo = st.sidebar.radio(
    "Selecione o Capítulo da Apostila:",
    [
        "📖 1. Teoria Aprofundada",
        "🧠 2. Mapa Mental Interativo",
        "📊 3. Simulador de Mudança de Fase",
        "📝 4. Teste de Nivelamento",
        "📚 5. Referências e Fontes"
    ]
)

st.sidebar.markdown("---")
st.sidebar.info("**Assunto:** Matéria, Estados Físicos e Modelos Atômicos Clássicos.\n\n**Nível:** 1º Ano Ensino Médio.")

# ==============================================================================
# MÓDULO 1: TEORIA APROFUNDADA
# ==============================================================================
if modulo == "📖 1. Teoria Aprofundada":
    st.title("📖 Material Didático de Química — Módulo Completo")
    st.markdown("---")
    
    sub_tab1, sub_tab2, sub_tab3 = st.tabs([
        "🔬 Matéria, Energia e Suas Propriedades", 
        "🌡️ Estados Físicos da Matéria", 
        "⚛️ Evolução dos Modelos Atômicos"
    ])
    
    with sub_tab1:
        st.markdown("""
        <div class="apostila-section">
            <h3 class="titulo-quimica">1. Introdução à Ciência Central</h3>
            <p>A Química estuda a <b>matéria</b>, suas transformações e as variações energéticas associadas. Define-se matéria como tudo o que apresenta <i>massa inercial</i> e ocupa lugar no espaço (<i>volume</i>).</p>
            
            <h4 style="color:#0D9488;">1.1 Propriedades Gerais vs. Específicas</h4>
            <ul>
                <li><b>Propriedades Gerais:</b> Comuns a qualquer corpo, não servem para identificar a substância pura. Exemplos: Massa, Volume, Impenetratividade (dois corpos não ocupam o mesmo lugar ao mesmo tempo), Divisibilidade e Compressibilidade.</li>
                <li><b>Propriedades Específicas:</b> Características exclusivas de uma substância pura, permitindo sua identificação. Subdividem-se em:
                    <ul>
                        <li><b>Físicas:</b> Constantes termodinâmicas como Ponto de Fusão (PF), Ponto de Ebulição (PE) e Densidade (massa/volume).</li>
                        <li><b>Químicas:</b> Reatividade química da matéria (ex: combustão, oxidação).</li>
                        <li><b>Organolépticas:</b> Perceptíveis pelos sentidos (cor, sabor, odor, brilho).</li>
                    </ul>
                </li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
    with sub_tab2:
        st.markdown("""
        <div class="apostila-section">
            <h3 class="titulo-quimica">2. Estados de Agregação da Matéria</h3>
            <p>O estado físico depende do balanço entre as <b>Forças de Coesão (atração)</b> e as <b>Forças de Repulsão (agitação térmica)</b> das partículas.</p>
            <ul>
                <li><b>Estado Sólido:</b> Forma e volume constantes. Partículas altamente organizadas com energia cinética mínima (apenas vibração local).</li>
                <li><b>Estado Líquido:</b> Forma variável e volume constante. Forças de atração e repulsão moderadas, permitindo fluidez.</li>
                <li><b>Estado Gasoso:</b> Forma e volume variáveis. Movimento caótico, alta desorganização e grande liberdade molecular.</li>
            </ul>
            <h4 style="color:#0D9488;">2.1 Classificação das Mudanças de Estado</h4>
            <ul>
                <li><b>Processos Endotérmicos (Absorvem calor):</b> Fusão (Sólido ➔ Líquido), Vaporização (Líquido ➔ Gás) e Sublimação (Sólido ➔ Gás).</li>
                <li><b>Processos Exotérmicos (Liberam calor):</b> Solidificação (Líquido ➔ Sólido), Condensação (Gás ➔ Líquido) e Resublimação (Gás ➔ Sólido).</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
    with sub_tab3:
        st.markdown("""
        <div class="apostila-section">
            <h3 class="titulo-quimica">3. Evolução dos Modelos Atômicos</h3>
            <p>O conceito de átomo evoluiu de ideias filosóficas para modelos validados por experimentos a partir do século XIX:</p>
            
            <h4 style="color:#1E3A8A;">⚫ 3.1 John Dalton (1808) — Bola de Bilhar</h4>
            <p>Baseado nas leis ponderais das reações químicas, postulou que a matéria é composta por partículas maciças, esféricas, indivisíveis e indestrutíveis.</p>
            
            <h4 style="color:#1E3A8A; margin-top:15px;">🍇 3.2 J.J. Thomson (1897) — Pudim de Passas</h4>
            <p>Utilizando a ampola de raios catódicos, provou a existência de partículas subatômicas negativas: os <b>elétrons</b>. O átomo passa a ser divisível, consistindo em uma esfera positiva recheada de cargas negativas.</p>
            
            <h4 style="color:#1E3A8A; margin-top:15px;">🪐 3.3 Ernest Rutherford (1911) — Modelo Planetário</h4>
            <p>Bombardeando uma folha de ouro com partículas alfa, descobriu que o átomo possui imensos espaços vazios (eletrosfera) e uma região central extremamente pequena e densa carregada positivamente chamada de <b>núcleo</b>.</p>
        </div>
        """, unsafe_allow_html=True)

# ==============================================================================
# MÓDULO 2: MAPA MENTAL
# ==============================================================================
elif modulo == "🧠 2. Mapa Mental Interativo":
    st.title("🧠 Estrutura Cognitiva Base — Química Geral")
    st.markdown("---")
    
    col_a, col_b, col_c = st.columns(3)
    
    with col_a:
        st.markdown('<div class="mm-box">MATÉRIA</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub"><b>Massa:</b> Quantidade de matéria medida em kg.</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub"><b>Volume:</b> Extensão ocupada no espaço (L ou m³).</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub"><b>Estados:</b> Sólido, Líquido e Gasoso.</div>', unsafe_allow_html=True)
        
    with col_b:
        st.markdown('<div class="mm-box" style="background: linear-gradient(135deg, #0D9488 0%, #115E59 100%);">MUDANÇAS DE FASE</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub" style="background-color:#1E3A8A;"><b>Fusão:</b> Absorção térmica Sólido ➔ Líquido.</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub" style="background-color:#1E3A8A;"><b>Sublimação:</b> Transição Sólido ➔ Gás direta.</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub" style="background-color:#1E3A8A;"><b>Condensação:</b> Perda de calor Gás ➔ Líquido.</div>', unsafe_allow_html=True)
        
    with col_c:
        st.markdown('<div class="mm-box" style="background: linear-gradient(135deg, #7C3AED 0%, #4C1D95 100%);">MODELOS ATÔMICOS</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub" style="background-color:#6D28D9;"><b>Dalton:</b> Esfera maciça, neutra e indivisível.</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub" style="background-color:#6D28D9;"><b>Thomson:</b> Cargas negativas fixas numa massa positiva.</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub" style="background-color:#6D28D9;"><b>Rutherford:</b> Núcleo central com eletrosfera vazia.</div>', unsafe_allow_html=True)

# ==============================================================================
# MÓDULO 3: SIMULADOR DE MUDANÇA DE FASE (BLINDADO CONTRA ERROS)
# ==============================================================================
elif modulo == "📊 3. Simulador de Mudança de Fase":
    st.title("📊 Simulador de Curva de Aquecimento")
    st.write("Ajuste as temperaturas de transição nos controles laterais.")
    st.markdown("---")
    
    col_ctrl, col_plot = st.columns([1, 2])
    
    with col_ctrl:
        st.subheader("Controle de Temperatura")
        pf = st.slider("Ponto de Fusão (°C)", min_value=-20, max_value=20, value=0, step=1)
        pe = st.slider("Ponto de Ebulição (°C)", min_value=60, max_value=120, value=100, step=1)
        
        st.info("""
        **Comportamento Gráfico:**
        Os patamares retos horizontais representam a coexistência de estados físicos enquanto a energia térmica rompe interações moleculares (Calor Latente).
        """)
        
    with col_plot:
        # Lógica de pontos estáticos lineares simples para eliminar qualquer erro de array
        tempo = [0, 5, 15, 25, 35, 45, 55]
        temperatura = [pf - 10, pf, pf, (pf + pe) / 2, pe, pe, pe + 10]
        
        figura = go.Figure()
        figura.add_trace(go.Scatter(
            x=tempo, 
            y=temperatura,
            mode='lines+markers',
            line=dict(color='#1E3A8A', width=4),
            marker=dict(size=8, color='#0D9488'),
            name='Substância Pura'
        ))
        
        figura.update_layout(
            title="Diagrama Termodinâmico de Aquecimento",
            xaxis_title="Tempo Decorrido (min)",
            yaxis_title="Temperatura (°C)",
            plot_bgcolor="#F8FAFC",
            margin=dict(l=20, r=20, t=40, b=20)
        )
        st.plotly_chart(figura, use_container_width=True)

# ==============================================================================
# MÓDULO 4: TESTE DE NIVELAMENTO (SEGURO CONTRA RESETS)
# ==============================================================================
elif modulo == "📝 4. Teste de Nivelamento":
    st.title("📝 Teste de Nivelamento — Padrão Objetivo")
    st.markdown("---")
    
    # Questões fixadas diretamente para evitar loops complexos no session_state
    st.markdown("##### 1. Um sistema composto por água líquida, gelo picado e óleo de cozinha apresenta quantas fases visíveis?")
    q1 = st.radio("Escolha uma opção:", ["1 fase", "2 fases", "3 fases", "4 fases"], key="quest_1")
    
    st.markdown("##### 2. Qual cientista propôs que o átomo se assemelhava a um 'pudim de passas' devido às suas descobertas elétricas?")
    q2 = st.radio("Escolha uma opção:", ["Dalton", "Thomson", "Rutherford", "Bohr"], key="quest_2")
    
    st.markdown("##### 3. A transição direta do estado gasoso para o estado sólido é denominada corretamente como:")
    q3 = st.radio("Escolha uma opção:", ["Fusão", "Vaporização", "Condensação", "Sublimação ou Resublimação"], key="quest_3")
    
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🏁 Corrigir Simulado"):
        st.session_state.quiz_completado = True
        
    if st.session_state.quiz_completado:
        st.markdown("### 📊 Análise do Seu Desempenho")
        acertos = 0
        
        if q1 == "3 fases":
            acertos += 1
            st.success("✓ Questão 1: Correta! (Água, gelo e óleo formam 3 fases distintas).")
        else:
            st.error("✕ Questão 1: Incorreta. O correto são 3 fases.")
            
        if q2 == "Thomson":
            acertos += 1
            st.success("✓ Questão 2: Correta! Thomson determinou a natureza elétrica do elétron.")
        else:
            st.error("✕ Questão 2: Incorreta. O correto é Thomson.")
            
        if q3 == "Sublimação ou Resublimação":
            acertos += 1
            st.success("✓ Questão 3: Correta! É a mudança direta de fase sem passar pelo estado líquido.")
        else:
            st.error("✕ Questão 3: Incorreta. O correto é Sublimação ou Resublimação.")
            
        nota = (acertos / 3) * 100
        st.metric("Aproveitamento Líquido", f"{nota:.1f}%", f"{acertos} de 3 acertos")
        st.progress(nota / 100)

# ==============================================================================
# MÓDULO 5: REFERÊNCIAS E FONTES
# ==============================================================================
elif modulo == "📚 5. Referências e Fontes":
    st.title("📚 Referências Bibliográficas Complementares")
    st.markdown("---")
    
    st.markdown("""
    <div class="apostila-section">
        <h4 style="color: #1E3A8A; font-weight:bold;">📖 Indicações de Leitura Teórica</h4>
        <ul>
            <li><b>FELTRE, Ricardo.</b> <i>Química Geral - Volume 1.</i> Editora Moderna, 2014.</li>
            <li><b>ATKINS, Peter.</b> <i>Princípios de Química.</i> Bookman Editora.</li>
        </ul>
    </div>
    
    <div class="apostila-section" style="border-left-color: #0D9488;">
        <h4 style="color: #0D9488; font-weight:bold;">🌐 Links Externos Oficiais</h4>
        <ul>
            <li><b>Portal Objetivo:</b> <a href="https://www.objetivo.br" target="_blank">Acesse o Material do Aluno</a></li>
            <li><b>Sociedade Brasileira de Química (SBQ):</b> Conteúdos e artigos acadêmicos complementares.</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
