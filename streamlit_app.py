# ────────────────────────────────────────────────────────────────────────────────
# app.py – fluxo código → OS → Enter (uma OS por amostra), atualiza Sheets
# e exporta Excel
# ────────────────────────────────────────────────────────────────────────────────
from __future__ import annotations
import io, os, json
from datetime import datetime
from typing import List, Dict

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

STATUS_COL = "AF"
DATE_COL   = "AG"
OS_COL     = "AH"

STATUS_VAL = "Retorno 1241"
DATE_FMT   = "%d/%m/%Y"
# ╰─────────────────────────────────────────────────────────────────────╯


# ───────────────────── Google Sheets helpers ──────────────────────────
@st.cache_resource
def _authorize_google() -> Credentials:
    token_path = "token.json"
    creds = Credentials.from_authorized_user_file(token_path, SCOPES) if os.path.exists(token_path) else None
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            client_config = json.loads(st.secrets["GOOGLE_CLIENT_SECRET"])
            flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
            creds = flow.run_console()
        with open(token_path, "w", encoding="utf-8") as fp:
            fp.write(creds.to_json())
    return creds


@st.cache_resource
def _svc():
    return build("sheets", "v4", credentials=_authorize_google(), cache_discovery=False)


def _col_to_idx(col: str) -> int:
    idx = 0
    for c in col: idx = idx * 26 + ord(c.upper()) - 64
    return idx - 1


def fetch_sheet() -> List[List[str]]:
    res = (
        _svc()
        .spreadsheets()
        .values()
        .get(spreadsheetId=SPREADSHEET_ID, range=SHEET_NAME, valueRenderOption="FORMATTED_VALUE")
        .execute()
    )
    return res.get("values", [])


def update_rows(rows: List[int], today: str, os_vals: List[str]) -> None:
    svc = _svc()
    data = []
    for idx, os_val in zip(rows, os_vals):
        data.extend([
            {"range": f"{SHEET_NAME}!{STATUS_COL}{idx}", "values": [[STATUS_VAL]]},
            {"range": f"{SHEET_NAME}!{DATE_COL}{idx}",   "values": [[today]]},
            {"range": f"{SHEET_NAME}!{OS_COL}{idx}",     "values": [[os_val]]},
        ])
    try:
        svc.spreadsheets().values().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"valueInputOption": "RAW", "data": data},
        ).execute()
    except HttpError as e:
        st.error("❌ Falha ao gravar no Google Sheets.")
        st.stop()


# ───────────────────────────── UI ─────────────────────────────────────
st.set_page_config("Selecionar Amostras", "🛢️", "centered")
st.title("Selecionar Amostras 🛢️")

# estado
if "pendente_codigo" not in st.session_state: st.session_state["pendente_codigo"] = ""
if "lista"            not in st.session_state: st.session_state["lista"] = {}      # {codigo: os}

def _add_codigo():
    cod = st.session_state["in_codigo"].strip()
    if cod:
        st.session_state["pendente_codigo"] = cod      # aguarda OS
    st.session_state["in_codigo"] = ""                 # limpa campo

def _add_os():
    os_val = st.session_state["in_os"].strip()
    cod = st.session_state["pendente_codigo"]
    if cod and os_val:
        st.session_state["lista"][cod] = os_val
        st.session_state["pendente_codigo"] = ""       # zera pendência
    st.session_state["in_os"] = ""

# entrada de código
st.text_input("📷 Código da amostra", key="in_codigo", on_change=_add_codigo)

# se há código pendente, pede OS
if st.session_state["pendente_codigo"]:
    st.text_input(f"🔧 OS p/ amostra {st.session_state['pendente_codigo']}", key="in_os", on_change=_add_os)

# mostra resumo
if st.session_state["lista"]:
    st.write("### Amostras lançadas")
    st.dataframe(pd.DataFrame(
        [{"Amostra": c, "OS": o} for c, o in st.session_state["lista"].items()]
    ), hide_index=True)
else:
    st.info("Nenhuma amostra adicionada.")

col1, col2 = st.columns(2)
with col1:
    if st.button("🗑️ Limpar lista"):
        st.session_state["lista"].clear()
        st.session_state["pendente_codigo"] = ""
with col2:
    gerar = st.button("📥 Gerar planilha")

# ───────────────────── GERAÇÃO / EXPORTAÇÃO ───────────────────────────
if gerar:
    if not st.session_state["lista"]:
        st.error("📋 A lista está vazia.")
        st.stop()

    with st.spinner("Consultando planilha…"):
        sheet = fetch_sheet()
        if not sheet: st.error("Aba vazia."); st.stop()
        header, *data = sheet
        g_idx  = _col_to_idx("G")
        os_idx = _col_to_idx(OS_COL)

        rows_idx, os_vals, rows_data = [], [], []
        for i, row in enumerate(data, start=2):                        # 1-based
            code = str(row[g_idx]).strip() if g_idx < len(row) else ""
            if code in st.session_state["lista"]:
                rows_idx.append(i)
                os_vals.append(st.session_state["lista"][code])
                rows_data.append(row)

        if not rows_idx:
            st.warning("Nenhuma das amostras está na planilha.")
            st.stop()

    today = datetime.now().strftime(DATE_FMT)
    with st.spinner("Atualizando planilha…"):
        update_rows(rows_idx, today, os_vals)

    with st.spinner("Gerando Excel…"):
        # monta DataFrame das amostras selecionadas
        norm = []
        for r, os_val in zip(rows_data, os_vals):
            r += [""] * (len(header) - len(r))
            r[os_idx] = os_val
            norm.append(r)
        df = pd.DataFrame(norm, columns=header)

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
            df.to_excel(xw, index=False, sheet_name="Amostras")
        buf.seek(0)

    st.success(f"✔️ {len(df)} amostra(s) exportada(s).")
    st.download_button(
        "⬇️ Baixar Excel",
        buf,
        file_name=f"amostras_{today}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
