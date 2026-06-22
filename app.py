from __future__ import annotations

import importlib
from typing import Any, Dict, List

try:
    import pandas as pd
except ModuleNotFoundError:
    pd = None

try:
    import streamlit as st
    STREAMLIT_AVAILABLE = True
except ModuleNotFoundError:
    st = None
    STREAMLIT_AVAILABLE = False

APP_VERSION = "v14.0.11"
APP_NAME = "Comitê Técnico V14"

MODULE_REGISTRY = [
    ("Governança", "modules.governance"),
    ("Excel Master", "modules.excel_master"),
    ("Power Query Master", "modules.powerquery_master"),
    ("Casos Reais", "modules.casos_reais"),
    ("UX Educacional", "modules.ux"),
    ("Auditoria Integrada", "modules.auditoria"),
]


def load_module_payload(module_path: str) -> Dict[str, Any]:
    module = importlib.import_module(module_path)
    return module.export_payload() if hasattr(module, "export_payload") else {}


def load_all_payloads() -> List[Dict[str, Any]]:
    payloads = []
    for label, module_path in MODULE_REGISTRY:
        payload = load_module_payload(module_path)
        payload["rotulo"] = label
        payloads.append(payload)
    return payloads


def validate_app() -> List[Dict[str, Any]]:
    payloads = load_all_payloads()
    return [
        {"regra": "Módulos registrados", "status": "OK" if len(MODULE_REGISTRY) >= 6 else "FALHA", "evidencia": str(len(MODULE_REGISTRY))},
        {"regra": "Payloads carregados", "status": "OK" if len(payloads) >= 6 else "FALHA", "evidencia": str(len(payloads))},
        {"regra": "Versão V14", "status": "OK" if APP_VERSION.startswith("v14") else "FALHA", "evidencia": APP_VERSION},
        {"regra": "Blocos sem falha core", "status": "OK" if all(p.get("metricas", {}).get("falhas_core", 0) == 0 for p in payloads) else "FALHA", "evidencia": "falhas_core == 0"},
    ]


def print_validation_report() -> None:
    checks = validate_app()
    falhas = sum(1 for c in checks if c["status"] == "FALHA")
    print("=" * 78)
    print("VALIDAÇÃO DO APP FINAL — BUILD V14")
    print("=" * 78)
    print(f"App: {APP_NAME}")
    print(f"Versão: {APP_VERSION}")
    print(f"Streamlit disponível: {STREAMLIT_AVAILABLE}")
    print(f"Módulos registrados: {len(MODULE_REGISTRY)}")
    print(f"Falhas core: {falhas}")
    print(f"Aprovado core: {falhas == 0}")
    print("-" * 78)
    for item in checks:
        print(f"[{item['status']}] {item['regra']} — {item['evidencia']}")
    print("=" * 78)


def render_app() -> None:
    if not STREAMLIT_AVAILABLE:
        raise RuntimeError("Streamlit não está instalado. Use print_validation_report().")
    st.set_page_config(page_title="Comitê Técnico V14", page_icon="📊", layout="wide")
    st.title(APP_NAME)
    st.caption("Build Final V14 · Excel · Power Query · Casos Reais · UX · Auditoria")
    payloads = load_all_payloads()
    with st.sidebar:
        page = st.radio("Navegação", ["Dashboard", "Módulos", "Auditoria", "Sobre"], index=0)
    if page == "Dashboard":
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Módulos", len(payloads))
        c2.metric("Falhas Core", sum(p.get("metricas", {}).get("falhas_core", 0) for p in payloads))
        c3.metric("Versão", APP_VERSION)
        c4.metric("Status", "V14")
        rows = []
        for p in payloads:
            row = {"Módulo": p.get("rotulo"), "Versão": p.get("versao"), "Status": p.get("status")}
            row.update(p.get("metricas", {}))
            rows.append(row)
        if pd is not None:
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
        else:
            st.write(rows)
    elif page == "Módulos":
        labels = [p.get("rotulo", p.get("nome", "Módulo")) for p in payloads]
        selected = st.selectbox("Selecione o módulo", labels)
        st.json(payloads[labels.index(selected)])
    elif page == "Auditoria":
        checks = validate_app()
        if pd is not None:
            st.dataframe(pd.DataFrame(checks), use_container_width=True, hide_index=True)
        else:
            st.write(checks)
    elif page == "Sobre":
        st.markdown("""
### Comitê Técnico V14

Plataforma educacional para preparação técnica em Excel, Power Query, Casos Reais, UX Educacional e Auditoria.

Módulos futuros previstos: DAX Master, Estatística Master, Simulados Técnicos e VBA Master dedicado.
""")


if __name__ == "__main__":
    if STREAMLIT_AVAILABLE:
        try:
            render_app()
        except Exception as exc:
            print_validation_report()
            print(f"AVISO: UI Streamlit não renderizada. Detalhe: {exc}")
    else:
        print_validation_report()
