import streamlit as st
import plotly.graph_objects as go
import numpy as np

# ==============================================================================
# CONFIGURAÇÃO DA PÁGINA E DESIGN SYSTEM (UX/UI ALTERNATIVA)
# ==============================================================================
st.set_page_config(
    page_title="Plataforma Objetivo - Química 1º Ano",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilização profissional injetada para simular um ambiente de apostila física/virtual
CUSTOM_CSS = """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Fira+Code:wght@400;500&family=Inter:wght@300;400;500;600;700&display=swap');
    
    html, body, [data-testid="stSidebar"] {
        font-family: 'Inter', sans-serif;
        color: #1E293B;
    }
    
    /* Box de Conteúdo Profundo (Apostila) */
    .apostila-section {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-left: 6px solid #1E3A8A;
        padding: 28px;
        border-radius: 12px;
        margin-bottom: 25px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.03);
    }
    
    .apostila-alerta {
        background-color: #FFFBEB;
        border-left: 6px solid #D97706;
        padding: 18px;
        border-radius: 8px;
        margin: 15px 0;
    }
    
    .titulo-quimica {
        color: #1E3A8A;
        font-weight: 700;
        border-bottom: 2px solid #E2E8F0;
        padding-bottom: 8px;
        margin-top: 20px;
    }
    
    /* Nós Avançados do Mapa Mental */
    .mm-box {
        background: linear-gradient(135deg, #1E3A8A 0%, #0F172A 100%);
        color: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
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
# GERENCIAMENTO DE ESTADO ROBUSTO (SESSION STATE)
# ==============================================================================
if 'quiz_completado' not in st.session_state:
    st.session_state.quiz_completado = False
if 'respostas_salvas' not in st.session_state:
    st.session_state.respostas_salvas = {}

# ==============================================================================
# MENU LATERAL DE NAVEGAÇÃO
# ==============================================================================
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/3022/3022271.png", width=70)
st.sidebar.markdown("# **Objetivo Virtual**\n*Material de Apoio de Alta Performance*")
st.sidebar.markdown("---")

modulo = st.sidebar.radio(
    "Selecione o Capítulo da Apostila:",
    [
        "📖 1. Teoria Aprofundada",
        "🧠 2. Mapa Mental Interativo",
        "📊 3. Simulador de Mudança de Fase",
        "📝 4. Teste de Nivelamento (Foco Vestibular)",
        "📚 5. Referências e Direcionamento"
    ]
)

st.sidebar.markdown("---")
st.sidebar.info("**Assunto:** Introdução à Química Geral & Modelos Atômicos antigos.\n\n**Público:** Ensino Médio / Pré-Vestibular.")

# ==============================================================================
# MÓDULO 1: TEORIA APROFUNDADA (MÁXIMO CONTEÚDO TÉCNICO)
# ==============================================================================
if modulo == "📖 1. Teoria Aprofundada":
    st.title("📖 Material Didático de Química — Módulo Completo")
    st.write("Abaixo encontra-se a transcrição integral dos conceitos analíticos de Matéria, Estados Físicos e evolução da Atomística Clássica.")
    st.markdown("---")
    
    sub_tab1, sub_tab2, sub_tab3 = st.tabs([
        "🔬 Matéria, Energia e Suas Propriedades", 
        "🌡️ Termodinâmica Fenomenológica e Estados Físicos", 
        "⚛️ Atomística e a Evolução dos Modelos Clássicos"
    ])
    
    with sub_tab1:
        st.markdown("""
        <div class="apostila-section">
            <h3 class="titulo-quimica">1. Introdução à Ciência Central</h3>
            <p>A Química estuda a <b>matéria</b>, suas transformações e as variações energéticas associadas a esses processos. Diferente do senso comum, cientificamente define-se matéria como tudo o que apresenta <i>massa inercial</i> e ocupa uma porção definida no espaço-tempo (<i>volume</i>).</p>
            
            <h4 style="color:#0D9488;">1.1 Propriedades Gerais vs. Específicas</h4>
            <p>As propriedades da matéria dividem-se em dois grandes blocos analíticos indispensáveis para exames de seleção:</p>
            <ul>
                <li><b>Propriedades Gerais:</b> Comuns a qualquer tipo de corpo, independente da substância que o compõe. Não permitem identificar a matéria. Exemplos: Massa, Extensão (Volume), Impenetratividade (dois corpos não ocupam o mesmo espaço simultaneamente), Divisibilidade e Compressibilidade.</li>
                <li><b>Propriedades Específicas:</b> Características exclusivas de um determinado grupo ou substância pura, permitindo sua identificação molecular ou atômica. Subdividem-se em:
                    <ul>
                        <li><b>Físicas:</b> Constantes termodinâmicas como Ponto de Fusão (PF), Ponto de Ebulição (PE) e Densidade ($\rho = m/V$).</li>
                        <li><b>Químicas:</b> Capacidade de reagir quimicamente e alterar a identidade estrutural (ex: inflamabilidade, oxidabilidade).</li>
                        <li><b>Organolépticas:</b> Perceptíveis pelos sentidos humanos (cor, sabor, odor, brilho).</li>
                    </ul>
                </li>
            </ul>
        </div>
        
        <div class="apostila-alerta">
            <b>⚠️ Nota de Rodapé do Vestibular:</b> Densidade é uma propriedade intensiva (não depende da quantidade de massa da amostra), enquanto massa e volume isolados são propriedades extensivas.
        </div>
        """, unsafe_allow_html=True)
        
    with sub_tab2:
        st.markdown("""
        <div class="apostila-section">
            <h3 class="titulo-quimica">2. Estados de Agregação da Matéria</h3>
            <p>O estado físico de uma substância depende diretamente do balanço cinético entre duas forças opostas no nível microscópico: as <b>Forças de Coesão (atração)</b> e as <b>Forças de Repulsão (agitação térmica)</b>.</p>
            
            <table style="width:100%; border-collapse: collapse; margin-top: 15px;">
                <tr style="background-color: #1E3A8A; color: white;">
                    <th style="padding: 10px; border: 1px solid #CBD5E1;">Característica</th>
                    <th style="padding: 10px; border: 1px solid #CBD5E1;">Estado Sólido</th>
                    <th style="padding: 10px; border: 1px solid #CBD5E1;">Estado Líquido</th>
                    <th style="padding: 10px; border: 1px solid #CBD5E1;">Estado Gasoso</th>
                </tr>
                <tr>
                    <td style="padding: 10px; border: 1px solid #CBD5E1; font-weight:bold;">Forma</td>
                    <td style="padding: 10px; border: 1px solid #CBD5E1;">Fixa e Constante</td>
                    <td style="padding: 10px; border: 1px solid #CBD5E1;">Variável (toma a forma do recipiente)</td>
                    <td style="padding: 10px; border: 1px solid #CBD5E1;">Variável (ocupa todo o espaço)</td>
                </tr>
                <tr style="background-color: #F8FAFC;">
                    <td style="padding: 10px; border: 1px solid #CBD5E1; font-weight:bold;">Volume</td>
                    <td style="padding: 10px; border: 1px solid #CBD5E1;">Fixo e Definido</td>
                    <td style="padding: 10px; border: 1px solid #CBD5E1;">Fixo e Constante</td>
                    <td style="padding: 10px; border: 1px solid #CBD5E1;">Variável (altamente expansível)</td>
                </tr>
                <tr>
                    <td style="padding: 10px; border: 1px solid #CBD5E1; font-weight:bold;">Energia Cinética</td>
                    <td style="padding: 10px; border: 1px solid #CBD5E1;">Mínima (Apenas vibração local)</td>
                    <td style="padding: 10px; border: 1px solid #CBD5E1;">Intermediária (Rotação e translação livre)</td>
                    <td style="padding: 10px; border: 1px solid #CBD5E1;">Máxima (Movimento caótico e desordenado)</td>
                </tr>
            </table>
            
            <h4 style="color:#0D9488; margin-top:20px;">2.1 Classificação das Mudanças de Estado</h4>
            <p>As transições de fase ocorrem por transferência calórica e são divididas em:</p>
            <ul>
                <li><b>Processos Endotérmicos (Absorvem calor):</b> Fusão (Sólido ➔ Líquido), Vaporização (Líquido ➔ Gás) e Sublimação (Sólido ➔ Gás).</li>
                <li><b>Processos Exotérmicos (Liberam calor):</b> Solidificação (Líquido ➔ Sólido), Condensação/Liquefação (Gás ➔ Líquido) e Resublimação (Gás ➔ Sólido).</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
    with sub_tab3:
        st.markdown("""
        <div class="apostila-section">
            <h3 class="titulo-quimica">3. Evolução dos Modelos Atômicos</h3>
            <p>O conceito de átomo migrou de uma abordagem filosófica na Grécia Antiga (Demócrito e Leucipo) para modelos científicos validados experimentalmente a partir do século XIX.</p>
            
            <h4 style="color:#1E3A8A;">⚫ 3.1 John Dalton (1808) — A Esfera Indivisível</h4>
            <p>Baseado nas leis ponderais das reações químicas (Lei de Lavoisier e Lei de Proust), Dalton postulou que:</p>
            <ol>
                <li>A matéria é composta por partículas maciças, esféricas e indivisíveis denominadas átomos.</li>
                <li>Átomos de um mesmo elemento químico possuem massas e propriedades idênticas.</li>
                <li>Nas reações químicas, os átomos não são criados nem destruídos, apenas rearranjados para formar novas substâncias.</li>
            </ol>
            
            <h4 style="color:#1E3A8A; margin-top:15px;">🍇 3.2 J.J. Thomson (1897) — A Natureza Elétrica</h4>
            <p>Utilizando a famosa ampola de Crookes (Experimento com Raios Catódicos), Thomson demonstrou a existência de partículas subatômicas de carga negativa desviadas por campos magnéticos: os <b>elétrons</b>.</p>
            <ul>
                <li><b>Derrubada do modelo anterior:</b> O átomo passa a ser <b>divisível</b> e possui carga elétrica líquida neutra.</li>
                <li><b>Analogia visual:</b> Uma massa fluida e contínua de carga positiva uniformemente distribuída salpicada por cargas negativas encrustadas (Modelo do Pudim de Passas).</li>
            </ul>
            
            <h4 style="color:#1E3A8A; margin-top:15px;">🪐 3.3 Ernest Rutherford (1911) — O Átomo Vazio</h4>
            <p>Bombardeando uma finíssima folha de ouro com partículas alfa ($\alpha$, de carga positiva), Rutherford observou que a maioria das partículas atravessava a chapa sem desvios, enquanto pouquíssimas sofriam desvios colossais ou rebatiam.</p>
            <ul>
                <li><b>Conclusões Disruptivas:</b> O átomo não é maciço; possui imensos espaços vazios (eletrosfera) e uma região central extremamente pequena, densa e carregada positivamente chamada de <b>núcleo</b>.</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

# ==============================================================================
# MÓDULO 2: MAPA MENTAL AVANÇADO
# ==============================================================================
elif modulo == "🧠 2. Mapa Mental Interativo":
    st.title("🧠 Estrutura Cognitiva Base — Química Geral")
    st.write("Mapa conceitual estruturado hierarquicamente para fixação de termos técnicos.")
    st.markdown("---")
    
    col_a, col_b, col_c = st.columns(3)
    
    with col_a:
        st.markdown('<div class="mm-box">MATÉRIA</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub"><b>Massa:</b> Quantidade de inércia medida em kg.</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub"><b>Volume:</b> Extensão tridimensional em m³ ou Litros.</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub"><b>Estados:</b> Sólido, Líquido e Gasoso.</div>', unsafe_allow_html=True)
        
    with col_b:
        st.markdown('<div class="mm-box" style="background: linear-gradient(135deg, #0D9488 0%, #115E59 100%);">TRANSFORMAÇÕES</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub" style="background-color:#1E3A8A;"><b>Fusão:</b> Absorção energética S➔L.</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub" style="background-color:#1E3A8A;"><b>Sublimação:</b> Direto do Sólido para Gás.</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub" style="background-color:#1E3A8A;"><b>Condensação:</b> Perda de calor G➔L.</div>', unsafe_allow_html=True)
        
    with col_c:
        st.markdown('<div class="mm-box" style="background: linear-gradient(135deg, #7C3AED 0%, #4C1D95 100%);">MODELOS ATÔMICOS</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub" style="background-color:#6D28D9;"><b>Dalton:</b> Esfera maciça, neutra e indestrutível.</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub" style="background-color:#6D28D9;"><b>Thomson:</b> Pudim de passas com cargas elétricas.</div>', unsafe_allow_html=True)
        st.markdown('<div class="mm-sub" style="background-color:#6D28D9;"><b>Rutherford:</b> Sistema planetário (Núcleo + Eletrosfera).</div>', unsafe_allow_html=True)

# ==============================================================================
# MÓDULO 3: SIMULADOR DE MUDANÇA DE FASE (SEM ERRO)
# ==============================================================================
elif modulo == "📊 3. Simulador de Mudança de Fase":
    st.title("📊 Simulador Científico de Curvas Termodinâmicas")
    st.write("Ajuste as propriedades termodinâmicas da substância fictícia para alterar a dinâmica dos patamares de transição de fase.")
    st.markdown("---")
    
    col_ctrl, col_plot = st.columns([1, 2])
    
    with col_ctrl:
        st.subheader("Parâmetros do Sistema")
        p_fusao = st.slider("Temperatura de Fusão (°C)", min_value=-40, max_value=40, value=0, step=2)
        p_ebulicao = st.slider("Temperatura de Ebulição (°C)", min_value=50, max_value=140, value=100, step=2)
        
        # Validação matemática lógica para evitar anomalias no gráfico
        if p_fusao >= p_ebulicao:
            st.error("Erro Crítico: O Ponto de Fusão deve ser estritamente inferior ao Ponto de Ebulição!")
            p_fusao, p_ebulicao = 0, 100
            
        st.markdown("""
        ### Análise do Gráfico:
        1. **Segmentos Inclinados:** Representam o aquecimento de uma única fase física (Aumento da Energia Cinética Média - Calor Sensível).
        2. **Patamares Horizontais:** Coexistência de fases durante a transição estrutural (Aumento da Energia Potencial - Calor Latente).
        """)
        
    with col_plot:
        # Criação de dados consistentes garantindo tamanho igual de arrays para x e y
        eixo_tempo = np.array([0, 5, 15, 20, 30, 35, 45])
        eixo_temp = np.array([p_fusao - 15, p_fusao, p_fusao, (p_fusao + p_ebulicao)/2, p_ebulicao, p_ebulicao, p_ebulicao + 15])
        
        figura = go.Figure()
        figura.add_trace(go.Scatter(
            x=eixo_tempo, 
            y=eixo_temp,
            mode='lines+markers',
            line=dict(color='#1E3A8A', width=4),
            marker=dict(size=10, color='#0D9488'),
            name='Substância Pura'
        ))
        
        figura.add_annotation(x=10, y=p_fusao, text="Patamar de Fusão (S + L)", showarrow=True, arrowhead=2, yshift=15, bgcolor="white")
        figura.add_annotation(x=32.5, y=p_ebulicao, text="Patamar de Ebulição (L + G)", showarrow=True, arrowhead=2, yshift=15, bgcolor="white")
        
        figura.update_layout(
            title="Diagrama de Aquecimento de uma Substância Pura",
            xaxis_title="Tempo de Exposição ao Calor (minutos)",
            yaxis_title="Temperatura Operacional (°C)",
            plot_bgcolor="#F8FAFC",
            grid=dict(rows=1, columns=1),
            margin=dict(l=20, r=20, t=40, b=20)
        )
        st.plotly_chart(figura, use_container_width=True)

# ==============================================================================
# MÓDULO 4: TESTE DE NIVELAMENTO AVANÇADO (PERSISTENTE)
# ==============================================================================
elif modulo == "📝 4. Teste de Nivelamento (Foco Vestibular)":
    st.title("📝 Simulado Analítico de Química Geral")
    st.write("Responda às questões complexas baseadas no estilo Fuvest/Unicamp adaptadas para o Colégio Objetivo.")
    st.markdown("---")
    
    # Base de dados estruturada do Quiz
    banco_questoes = [
        {
            "id": "item1",
            "pergunta": "1. (Objetivo-Adaptado) Um cientista analisa uma amostra que se funde totalmente a uma temperatura constante de 55°C, mas cuja temperatura de ebulição varia entre 110°C e 118°C. Essa amostra pode ser classificada como:",
            "opcoes": ["Substância pura simples", "Substância pura composta", "Mistura comum", "Mistura Eutética", "Mistura Azeotrópica"],
            "gabarito": "Mistura Eutética"
        },
        {
            "id": "item2",
            "pergunta": "2. (Fuvest) Thomson determinou, pela primeira vez, a relação entre a massa e a carga do elétron, o que pode ser considerado como a descoberta do elétron. O modelo atômico proposto por ele ficou conhecido como:",
            "opcoes": ["Modelo planetário", "Bola de bilhar", "Pudim de passas", "Átomo quântico", "Modelo orbital"],
            "gabarito": "Pudim de passas"
        },
        {
            "id": "item3",
            "pergunta": "3. No experimento de Rutherford, o fato de algumas poucas partículas alfa sofrerem grandes desvios ao atravessarem a lâmina de ouro permitiu concluir que:",
            "opcoes": ["O átomo é totalmente maciço.", "Os elétrons estão concentrados no núcleo.", "O núcleo é muito pequeno, denso e possui carga positiva.", "O átomo tem carga líquida negativa.", "A eletrosfera abriga a maior parte da massa do átomo."],
            "gabarito": "O núcleo é muito pequeno, denso e possui carga positiva."
        }
    ]
    
    # Loop de renderização das perguntas coletando no dicionário de estados
    for q in banco_questoes:
        st.markdown(f"##### {q['pergunta']}")
        resposta_usuario = st.radio(
            "Selecione sua alternativa:",
            q["opcoes"],
            key=f"interacao_{q['id']}"
        )
        st.session_state.respostas_salvas[q["id"]] = resposta_usuario
        st.markdown("<br>", unsafe_allow_html=True)
        
    if st.button("🏁 Corrigir e Calcular Desempenho Acadêmico"):
        st.session_state.quiz_completado = True
        
    if st.session_state.quiz_completado:
        st.markdown("### 📊 Relatório Estatístico do Aluno")
        acertos = 0
        
        for q in banco_questoes:
            resp = st.session_state.respostas_salvas.get(q["id"])
            if resp == q["gabarito"]:
                acertos += 1
                st.success(f"🎯 **Questão {q['id'][-1]}: Correta!** Você marcou '{resp}'.")
            else:
                st.error(f"❌ **Questão {q['id'][-1]}: Incorreta.** Você escolheu '{resp}'. O gabarito oficial é '{q['gabarito']}'.")
                
        porcentagem = (acertos / len(banco_questoes)) * 100
        st.metric("Aproveitamento Líquido", f"{porcentagem:.1f}%", f"{acertos} acertos de {len(banco_questoes)}")
        st.progress(porcentagem / 100)

# ==============================================================================
# MÓDULO 5: REFERÊNCIAS E DIRECIONAMENTO EXTERNO
# ==============================================================================
elif modulo == "📚 5. Referências e Direcionamento":
    st.title("📚 Fontes Bibliográficas e Extensões Didáticas")
    st.write("Consulte os materiais científicos de referência nacional utilizados na elaboração deste ecossistema virtual.")
    st.markdown("---")
    
    st.markdown("""
    <div class="apostila-section">
        <h4 style="color: #1E3A8A; font-weight:bold;">📖 Bibliografia Recomendada (Livros-Base)</h4>
        <ul>
            <li><b>FELTRE, Ricardo.</b> <i>Química Geral - Volume 1.</i> Editora Moderna, 2014. (Referência absoluta para exames nacionais).</li>
            <li><b>ATKINS, Peter; JONES, Loretta.</b> <i>Princípios de Química: Questionando a Vida Moderna e o Meio Ambiente.</i> Bookman Editora.</li>
        </ul>
    </div>
    
    <div class="apostila-section" style="border-left-color: #0D9488;">
        <h4 style="color: #0D9488; font-weight:bold;">🌐 Portais Oficiais de Pesquisa e Exercícios</h4>
        <ul>
            <li><b>Portal Oficial Objetivo:</b> <a href="https://www.objetivo.br" target="_blank">Acesso ao Aluno</a> para downloads de cadernos de exercícios e resoluções comentadas.</li>
            <li><b>Sociedade Brasileira de Química (SBQ):</b> <a href="http://www.sbq.org.br" target="_blank">Acesse o portal da SBQ</a> para artigos acadêmicos sobre história da ciência.</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
