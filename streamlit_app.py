# ────────────────────────────────────────────────────────────────────────────────
# app.py – Selecionar amostras no Google Sheets, marcar “Retorno 1241”,
# informar Ordem de Serviço (coluna AH) e baixar Excel
# Execute com:  streamlit run app.py
# ────────────────────────────────────────────────────────────────────────────────
from __future__ import annotations

import io, os, json
from datetime import datetime
from typing import List

import pandas as pd
import streamlit as st
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

# ╭────────────────────────── CONFIGURAÇÕES ───────────────────────────╮
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "1VLDQUCO3Aw4ClAvhjkUsnBxG44BTjz-MjHK04OqPxYM"
SHEET_NAME     = "Geral"

STATUS_COL = "AF"                     # Status
DATE_COL   = "AG"                     # Data
OS_COL     = "AH"                     # Ordem de Serviço

STATUS_VAL = "Retorno 1241"
DATE_FMT   = "%d/%m/%Y"
# ╰────────────────────────────────────────────────────────────────────╯

# ─────────────────────── Google Sheets helpers ───────────────────────
@st.cache_resource
def _authorize_google() -> Credentials:
    token_path = "token.json"
    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            try:
                client_config = json.loads(st.secrets["GOOGLE_CLIENT_SECRET"])
            except Exception:
                st.error("❌ Credenciais do Google não encontradas.")
                st.stop()
            flow  = InstalledAppFlow.from_client_config(client_config, SCOPES)
            creds = flow.run_console()
        with open(token_path, "w", encoding="utf-8") as fp:
            fp.write(creds.to_json())
    return creds


@st.cache_resource
def _get_service():
    return build("sheets", "v4", credentials=_authorize_google(), cache_discovery=False)


def _col_to_index(col: str) -> int:
    """Converte 'A' → 0, 'B' → 1, …"""
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c.upper()) - 64)
    return idx - 1


def fetch_sheet() -> List[List[str]]:
    """Lê a aba inteira como texto."""
    res = (
        _get_service()
        .spreadsheets()
        .values()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            range=SHEET_NAME,
            valueRenderOption="FORMATTED_VALUE",
        )
        .execute()
    )
    return res.get("values", [])


def update_rows(rows_idx: List[int], today: str, os_val: str) -> None:
    """Atualiza Status, Data e Ordem de Serviço nas linhas indicadas (1-based)."""
    svc = _get_service()

    ranges = {
        STATUS_COL: {"values": [[STATUS_VAL]] * len(rows_idx)},
        DATE_COL  : {"values": [[today]]       * len(rows_idx)},
        OS_COL    : {"values": [[os_val]]      * len(rows_idx)},
    }

    try:
        for col, body in ranges.items():
            rng = f"{SHEET_NAME}!{col}{rows_idx[0]}:{col}{rows_idx[-1]}"
            svc.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=rng,
                valueInputOption="RAW",
                body=body,
            ).execute()
    except HttpError:
        st.error("❌ Falha ao atualizar dados no Google Sheets.")
        st.stop()

# ───────────────────────────── UI Streamlit ────────────────────────────
st.set_page_config(page_title="Selecionar Amostras", page_icon="🛢️", layout="centered")
st.title("Selecionar Amostras 🛢️")

if "samples" not in st.session_state:
    st.session_state["samples"] = []
if "current_input" not in st.session_state:
    st.session_state["current_input"] = ""
if "order_service" not in st.session_state:
    st.session_state["order_service"] = ""

def _add_sample():
    code = st.session_state["current_input"].strip()
    if code and code not in st.session_state["samples"]:
        st.session_state["samples"].append(code)
    st.session_state["current_input"] = ""   # limpa campo após Enter

st.text_input(
    "📷 Escaneie o código de barras e pressione Enter",
    key="current_input",
    on_change=_add_sample,
)

# NOVO: campo Ordem de Serviço
st.text_input(
    "🔧 Ordem de Serviço (para todas as amostras selecionadas)",
    key="order_service",
)

st.write("### Amostras pré-selecionadas")
st.write(", ".join(st.session_state["samples"]) or "Nenhuma.")

col1, col2 = st.columns(2)
with col1:
    if st.button("🗑️ Limpar lista"):
        st.session_state["samples"].clear()
with col2:
    gerar = st.button("📥 Gerar planilha")

# ────────────────── Seleção, atualização e exportação ──────────────────
if gerar and st.session_state["samples"]:
    os_val = st.session_state["order_service"].strip()
    if not os_val:
        st.warning("Informe a Ordem de Serviço antes de gerar a planilha.")
        st.stop()

    with st.spinner("Consultando Google Sheets…"):
        sheet_rows = fetch_sheet()
        if not sheet_rows:
            st.error("Aba vazia ou não encontrada.")
            st.stop()

        header, *data = sheet_rows
        g_idx   = _col_to_index("G")
        os_idx  = _col_to_index(OS_COL)

        selected_rows, lines_idx = [], []
        samples_set = {s.strip() for s in st.session_state["samples"]}

        for i, row in enumerate(data, start=2):          # 1-based (linha 1 = header)
            sample_text = str(row[g_idx]).strip() if g_idx < len(row) else ""
            if sample_text in samples_set:
                selected_rows.append(row)
                lines_idx.append(i)

        if not selected_rows:
            st.warning("Nenhuma amostra encontrada na planilha.")
            st.stop()

    today = datetime.now().strftime(DATE_FMT)
    with st.spinner("Atualizando Google Sheets…"):
        update_rows(lines_idx, today, os_val)

    with st.spinner("Gerando arquivo Excel…"):
        # Normaliza tamanho das linhas e insere OS
        norm_rows = []
        for r in selected_rows:
            r += [""] * (len(header) - len(r))           # completa células vazias
            r[os_idx] = os_val                           # preenche coluna AH
            norm_rows.append(r)

        df  = pd.DataFrame(norm_rows, columns=header)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
            df.to_excel(xw, index=False, sheet_name="Amostras")
        buf.seek(0)

    st.success(f"✔️ {len(df)} amostra(s) exportada(s).")
    st.download_button(
        "⬇️ Baixar Excel",
        data=buf,
        file_name=f"amostras_{today}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

elif gerar:
    st.error("📋 A lista de amostras está vazia.")
