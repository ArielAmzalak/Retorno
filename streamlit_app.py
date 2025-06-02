# ────────────────────────────────────────────────────────────────────────────────
# app.py – Selecionar amostras, atribuir OS individual, marcar “Retorno 1241”
# e baixar Excel
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

# ╭──────────────────────── CONFIGURAÇÕES GERAIS ────────────────────────╮
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]
SPREADSHEET_ID = "1VLDQUCO3Aw4ClAvhjkUsnBxG44BTjz-MjHK04OqPxYM"
SHEET_NAME     = "Geral"

STATUS_COL = "AF"                     # Status
DATE_COL   = "AG"                     # Data
OS_COL     = "AH"                     # Ordem de Serviço

STATUS_VAL = "Retorno 1241"
DATE_FMT   = "%d/%m/%Y"
# ╰──────────────────────────────────────────────────────────────────────╯


# ─────────────────────────── Google helpers ────────────────────────────
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
            client_config = json.loads(st.secrets["GOOGLE_CLIENT_SECRET"])
            flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
            creds = flow.run_console()
        with open(token_path, "w", encoding="utf-8") as fp:
            fp.write(creds.to_json())
    return creds


@st.cache_resource
def _get_service():
    return build("sheets", "v4", credentials=_authorize_google(), cache_discovery=False)


def _col_to_index(col: str) -> int:
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c.upper()) - 64)
    return idx - 1


def fetch_sheet() -> List[List[str]]:
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


def update_rows(rows_idx: List[int], today: str, os_vals: List[str]) -> None:
    """
    Atualiza cada linha (mesmo não contígua) com:
      • STATUS_COL  = STATUS_VAL
      • DATE_COL    = today
      • OS_COL      = os_vals[i]
    """
    svc  = _get_service()
    data = []

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
    except HttpError:
        st.error("❌ Falha ao atualizar dados no Google Sheets.")
        st.stop()


# ───────────────────────────── UI Streamlit ────────────────────────────
st.set_page_config(page_title="Selecionar Amostras", page_icon="🛢️", layout="centered")
st.title("Selecionar Amostras 🛢️")

# ----- estado -----
if "samples"    not in st.session_state: st.session_state["samples"] = []      # lista de códigos
if "sample_os"  not in st.session_state: st.session_state["sample_os"] = {}    # dict código ➜ OS
if "current_in" not in st.session_state: st.session_state["current_in"] = ""


def _add_sample():
    code = st.session_state["current_in"].strip()
    if code and code not in st.session_state["samples"]:
        st.session_state["samples"].append(code)
        st.session_state["sample_os"][code] = ""       # OS em branco
    st.session_state["current_in"] = ""                # limpa input


st.text_input(
    "📷 Escaneie o código de barras e pressione Enter",
    key="current_in",
    on_change=_add_sample,
)

# ----- editor de OS -----
if st.session_state["samples"]:
    df_display = pd.DataFrame(
        [{"Amostra": s, "OS": st.session_state["sample_os"][s]}
         for s in st.session_state["samples"]]
    )
    edited_df = st.data_editor(
        df_display,
        column_config={"Amostra": st.column_config.TextColumn(read_only=True)},
        num_rows="dynamic",
        use_container_width=True,
        key="data_editor",
    )

    # salva mudanças no estado
    for _, row in edited_df.iterrows():
        st.session_state["sample_os"][row["Amostra"]] = str(row["OS"]).strip()

else:
    st.write("### Nenhuma amostra adicionada.")

# ----- botões -----
col1, col2 = st.columns(2)
with col1:
    if st.button("🗑️ Limpar lista"):
        st.session_state["samples"].clear()
        st.session_state["sample_os"].clear()
with col2:
    gerar = st.button("📥 Gerar planilha")


# ────────────────── Seleção, atualização e exportação ──────────────────
if gerar and st.session_state["samples"]:
    # Verifica se todas as OS foram preenchidas
    os_vazias = [s for s in st.session_state["samples"]
                 if not st.session_state["sample_os"][s]]
    if os_vazias:
        st.warning(f"Preencha a OS das amostras em branco: {', '.join(os_vazias)}.")
        st.stop()

    with st.spinner("Consultando Google Sheets…"):
        sheet_rows = fetch_sheet()
        if not sheet_rows:
            st.error("Aba vazia ou não encontrada.")
            st.stop()

        header, *data = sheet_rows
        g_idx  = _col_to_index("G")
        os_idx = _col_to_index(OS_COL)

        selected_rows, lines_idx, os_vals = [], [], []
        samples_set = set(st.session_state["samples"])

        for i, row in enumerate(data, start=2):              # 1-based
            sample_text = str(row[g_idx]).strip() if g_idx < len(row) else ""
            if sample_text in samples_set:
                selected_rows.append(row)
                lines_idx.append(i)
                os_vals.append(st.session_state["sample_os"][sample_text])

        if not selected_rows:
            st.warning("Nenhuma amostra encontrada na planilha.")
            st.stop()

    today = datetime.now().strftime(DATE_FMT)
    with st.spinner("Atualizando Google Sheets…"):
        update_rows(lines_idx, today, os_vals)

    with st.spinner("Gerando arquivo Excel…"):
        # Garante tamanho da linha e injeta OS respectiva
        norm_rows = []
        for row, os_val in zip(selected_rows, os_vals):
            row += [""] * (len(header) - len(row))
            row[os_idx] = os_val
            norm_rows.append(row)

        df = pd.DataFrame(norm_rows, columns=header)
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
