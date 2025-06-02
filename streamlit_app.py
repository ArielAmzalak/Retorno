# ────────────────────────────────────────────────────────────────────────────────
# app.py – Selecionar amostras, informar OS individual, marcar “Retorno 1241”
# e baixar Excel com as linhas correspondentes
#
# Uso:
#   1. Digite/escaneie o código da amostra.
#   2. Digite a Ordem de Serviço (OS) correspondente.
#   3. Clique “Adicionar” – a dupla Amostra/OS aparece na tabela.
#   4. Repita (1–3) quantas vezes precisar.
#   5. Clique “Gerar planilha” para:
#        • AF ← “Retorno 1241”
#        • AG ← data de hoje (dd/mm/aaaa)
#        • AH ← OS de cada amostra
#        • Baixar Excel contendo apenas essas amostras.
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
SHEET_NAME     = "Geral"               # aba onde estão as amostras

STATUS_COL = "AF"                     # Status
DATE_COL   = "AG"                     # Data
OS_COL     = "AH"                     # Ordem de Serviço
SAMPLE_COL = "G"                      # onde está o código da amostra

STATUS_VAL = "Retorno 1241"
DATE_FMT   = "%d/%m/%Y"
# ╰─────────────────────────────────────────────────────────────────────╯


# ──────────────────────── Helpers Google Sheets ───────────────────────
@st.cache_resource
def _authorize_google() -> Credentials:
    """Faz login (token.json) ou inicia fluxo OAuth para obter credenciais."""
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
    """Instância de serviço Sheets (cacheada)."""
    return build("sheets", "v4", credentials=_authorize_google(), cache_discovery=False)


def _col_to_idx(col: str) -> int:
    """Converte 'A'→0, 'B'→1 … para usar como índice de lista."""
    idx = 0
    for c in col:
        idx = idx * 26 + (ord(c.upper()) - 64)
    return idx - 1


def fetch_sheet() -> List[List[str]]:
    """Lê toda a aba como valores de texto."""
    res = (
        _svc()
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
    Atualiza cada linha indicada (1-based) com:
      • STATUS_COL = STATUS_VAL
      • DATE_COL   = today
      • OS_COL     = os_vals[i]
    """
    svc  = _svc()
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
        st.error("❌ Falha ao gravar no Google Sheets.")
        st.stop()


# ─────────────────────────── Interface Streamlit ──────────────────────
st.set_page_config(page_title="Selecionar Amostras", page_icon="🛢️", layout="centered")
st.title("Selecionar Amostras 🛢️")

# ---------- Estado ----------
if "lista" not in st.session_state:          # {codigo: os}
    st.session_state["lista"]: Dict[str, str] = {}
if "in_codigo" not in st.session_state:
    st.session_state["in_codigo"] = ""
if "in_os" not in st.session_state:
    st.session_state["in_os"] = ""

# ---------- Formulário de entrada ----------
c1, c2, c3 = st.columns([3, 3, 1])
with c1:
    st.text_input("📷 Código da amostra", key="in_codigo")
with c2:
    st.text_input("🔧 Ordem de Serviço (OS)", key="in_os")
with c3:
    if st.button("➕ Adicionar"):
        cod = st.session_state["in_codigo"].strip()
        osv = st.session_state["in_os"].strip()
        if not cod or not osv:
            st.warning("Preencha **ambos** os campos antes de adicionar.")
        elif cod in st.session_state["lista"]:
            st.warning(f"Amostra {cod} já lançada; edite ou remova antes.")
        else:
            st.session_state["lista"][cod] = osv
            st.session_state["in_codigo"] = ""
            st.session_state["in_os"] = ""

# ---------- Tabela resumo ----------
if st.session_state["lista"]:
    st.write("### Amostras lançadas")
    st.dataframe(
        pd.DataFrame(
            [{"Amostra": c, "OS": o} for c, o in st.session_state["lista"].items()]
        ),
        hide_index=True,
    )
else:
    st.info("Nenhuma amostra adicionada.")

# ---------- Botões de ação ----------
col1, col2 = st.columns(2)
with col1:
    if st.button("🗑️ Limpar lista"):
        st.session_state["lista"].clear()
with col2:
    gerar = st.button("📥 Gerar planilha")

# ──────────────────── Seleção, gravação e exportação ───────────────────
if gerar:
    if not st.session_state["lista"]:
        st.error("📋 A lista está vazia.")
        st.stop()

    # 1. Carrega planilha e encontra as linhas das amostras
    with st.spinner("Consultando planilha…"):
        sheet = fetch_sheet()
        if not sheet:
            st.error("Aba vazia ou não encontrada.")
            st.stop()
        header, *data = sheet
        sample_idx  = _col_to_idx(SAMPLE_COL)
        os_col_idx  = _col_to_idx(OS_COL)

        rows_idx, os_vals, rows_data = [], [], []
        for i, row in enumerate(data, start=2):  # 1-based
            code = str(row[sample_idx]).strip() if sample_idx < len(row) else ""
            if code in st.session_state["lista"]:
                rows_idx.append(i)
                os_vals.append(st.session_state["lista"][code])
                rows_data.append(row)

        if not rows_idx:
            st.warning("Nenhuma das amostras está presente no Sheets.")
            st.stop()

    # 2. Atualiza AF / AG / AH
    today = datetime.now().strftime(DATE_FMT)
    with st.spinner("Atualizando Google Sheets…"):
        update_rows(rows_idx, today, os_vals)

    # 3. Gera Excel com as linhas selecionadas
    with st.spinner("Gerando Excel…"):
        norm_rows = []
        for row, os_val in zip(rows_data, os_vals):
            row += [""] * (len(header) - len(row))   # completa tamanho
            row[os_col_idx] = os_val                # garante OS correta
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
