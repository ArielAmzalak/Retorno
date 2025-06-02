# ────────────────────────────────────────────────────────────────────────────────
# app.py – Selecionar amostras, informar OS individual, marcar “Retorno 1241”
# e baixar Excel com as mesmas colunas gravadas no Sheets
# ────────────────────────────────────────────────────────────────────────────────
from __future__ import annotations
import io, os, json
from datetime import datetime
from typing import Dict, List

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

STATUS_COL  = "AF"                    # Status
DATE_COL    = "AG"                    # Data
OS_COL      = "AH"                    # Ordem de Serviço
SAMPLE_COL  = "G"                     # código da amostra

STATUS_VAL  = "Retorno 1241"
DATE_FMT    = "%d/%m/%Y"
# ╰─────────────────────────────────────────────────────────────────────╯


# ─────────────────────── Helpers Google Sheets ────────────────────────
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
    for c in col: idx = idx * 26 + (ord(c.upper()) - 64)
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


def update_rows(rows_idx: List[int], today: str, os_vals: List[str]) -> None:
    """Escreve STATUS, DATA e OS em cada linha indicada (1-based)."""
    if len(rows_idx) != len(os_vals):
        st.error("Inconsistência interna: linhas e OS não batem.")
        st.stop()

    svc, data = _svc(), []
    for idx, os_val in zip(rows_idx, os_vals):
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
        st.error(f"❌ Falha ao gravar no Google Sheets: {e}")
        st.stop()


# ─────────────────────────── Interface Streamlit ──────────────────────
st.set_page_config(page_title="Selecionar Amostras", page_icon="🛢️", layout="centered")
st.title("Selecionar Amostras 🛢️")

# ---------- estado ----------
st.session_state.setdefault("lista", {})     # {código: OS}
st.session_state.setdefault("in_codigo", "")
st.session_state.setdefault("in_os", "")
st.session_state.setdefault("msg", "")

# ---------- callback ----------
def add_item() -> None:
    cod, osv = st.session_state.in_codigo.strip(), st.session_state.in_os.strip()
    if not cod or not osv:
        st.session_state.msg = "Preencha **ambos** os campos."
        return
    if cod in st.session_state.lista:
        st.session_state.msg = f"Amostra {cod} já lançada."
        return
    st.session_state.lista[cod] = osv
    st.session_state.in_codigo = ""
    st.session_state.in_os = ""
    st.session_state.msg = ""

# ---------- formulário ----------
c1, c2, c3 = st.columns([3, 3, 1])
with c1: st.text_input("📷 Código da amostra", key="in_codigo")
with c2: st.text_input("🔧 Ordem de Serviço (OS)", key="in_os")
with c3: st.button("➕ Adicionar", on_click=add_item)

if st.session_state.msg: st.warning(st.session_state.msg)

# ---------- tabela resumo ----------
if st.session_state.lista:
    st.dataframe(pd.DataFrame(
        [{"Amostra": c, "OS": o} for c, o in st.session_state.lista.items()]
    ), hide_index=True)
else:
    st.info("Nenhuma amostra adicionada.")

# ---------- botões ----------
col1, col2 = st.columns(2)
with col1:
    if st.button("🗑️ Limpar lista"):
        st.session_state.lista.clear()
        st.session_state.msg = ""
with col2:
    gerar = st.button("📥 Gerar planilha")

# ──────────────────── Seleção, gravação e exportação ───────────────────
if gerar:
    if not st.session_state.lista:
        st.error("📋 A lista está vazia.")
        st.stop()

    # 1. Localiza as amostras na planilha
    with st.spinner("Consultando planilha…"):
        sheet = fetch_sheet()
        if not sheet: st.error("Aba vazia."); st.stop()

        header, *data = sheet
        sample_idx, os_idx = _col_to_idx(SAMPLE_COL), _col_to_idx(OS_COL)

        rows_idx, os_vals, rows_data = [], [], []
        for i, row in enumerate(data, start=2):
            code = str(row[sample_idx]).strip() if sample_idx < len(row) else ""
            if code in st.session_state.lista:
                rows_idx.append(i)
                os_vals.append(st.session_state.lista[code])
                rows_data.append(row)

        if not rows_idx:
            st.warning("Nenhuma das amostras digitadas está na planilha.")
            st.stop()

    # 2. Atualiza AF / AG / AH
    today = datetime.now().strftime(DATE_FMT)
    with st.spinner("Gravando no Google Sheets…"):
        update_rows(rows_idx, today, os_vals)

    # 3. Prepara DataFrame já com STATUS & DATA
    status_idx, date_idx = _col_to_idx(STATUS_COL), _col_to_idx(DATE_COL)
    norm_rows = []
    for row, os_val in zip(rows_data, os_vals):
        row += [""] * (len(header) - len(row))        # completa largura
        row[status_idx] = STATUS_VAL
        row[date_idx]   = today
        row[os_idx]     = os_val
        norm_rows.append(row)

    df = pd.DataFrame(norm_rows, columns=header)

    # 4. Exporta Excel
    with st.spinner("Gerando Excel…"):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as xw:
            df.to_excel(xw, index=False, sheet_name="Amostras")
        buf.seek(0)

    st.success(f"✔️ {len(df)} amostra(s) atualizada(s) e exportada(s).")
    st.download_button(
        "⬇️ Baixar Excel",
        data=buf,
        file_name=f"amostras_{today}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
