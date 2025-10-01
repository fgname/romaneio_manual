import io, os, re, zipfile, unicodedata, base64
from datetime import date, datetime
from urllib.parse import urlparse, quote

import pandas as pd
import requests
import streamlit as st

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import ParagraphStyle

# ==========================
# CONFIG
# ==========================
APP_TITLE       = "Romaneio Manual — OneDrive (FINALIZADOS)"
ONEDRIVE_LINK   = "https://tecadi-my.sharepoint.com/:x:/g/personal/rafael_alves_tecadi_com_br/EaJshSFavb5Pv8z_dpW3ZWwB8cAXsuPSrGyYoAB7ye11Aw"

SHEET_NAME      = "PROCESSOS S.LEITURA"
CLIENTE_FIXO    = "SPRINGER CARRIER LTDA (MIDEA)."
CONTEUDO_FIXO   = "CHEIO"
PRODUTO_TITULO  = "FAST CIF/FOB"
FORM_COD        = "FM 108"
FORM_REV        = "00"

BG_IMAGE_PATH   = "fundoapp.jpg"      # fundo do APP
LOGO_IMAGE_PATH = "logo_tecadi.png"   # logo no PDF (fundo do PDF é branco)

REQUIRED_FOR_PDF = ["ARMADOR","TECADI","SKU","QTD","LISTA","DEMANDA","TRANSPORTADORA","NOME","PLACA"]

# ==========================
# HELPERS (planilha)
# ==========================
def normalize(s: str) -> str:
    s = str(s)
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    s = " ".join(s.strip().split())
    return s.upper()

COLMAP = {
    "DEMANDA":"DEMANDA","HORARIO":"HORARIO","ARMADOR":"ARMADOR","TRANSPORTADORA":"TRANSPORTADORA",
    "DATA PROGRAMACAO":"DATA PROGRAMAÇÃO","DATA PROGRAMAÇÃO":"DATA PROGRAMAÇÃO","ROMANEIO":"ROMANEIO",
    "SKU":"SKU","QTD":"QTD","M3":"M3","STATUS":"STATUS","NOME":"NOME","MOTORISTA":"NOME",
    "PLACA":"PLACA","TECADI":"TECADI","LISTA":"LISTA"
}

def find_header_row(df_raw: pd.DataFrame) -> int:
    for i in range(min(20, len(df_raw))):
        joined = " ".join([str(x) for x in list(df_raw.iloc[i].values)])
        if "DEMANDA" in joined.upper() and "ARMADOR" in joined.upper():
            return i
    return 0

def rename_columns(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns={c: COLMAP.get(normalize(c), normalize(c)) for c in df.columns}) \
             .loc[:, lambda d: ~d.columns.duplicated()].copy()

def load_excel_from_bytes(xls_bytes: bytes) -> pd.DataFrame:
    df_raw = pd.read_excel(io.BytesIO(xls_bytes), sheet_name=SHEET_NAME, header=None, engine="openpyxl")
    header_row = find_header_row(df_raw)
    headers = df_raw.iloc[header_row].astype(str).tolist()
    df = df_raw.iloc[header_row+1:].copy()
    df.columns = headers
    df = df.iloc[:, :15]  # A:O
    df = df.dropna(how="all")
    df = rename_columns(df)

    if "QTD" in df.columns:
        df["QTD"] = (df["QTD"].astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
        df["QTD"] = pd.to_numeric(df["QTD"], errors="coerce")

    if "STATUS" in df.columns:
        df["STATUS"] = df["STATUS"].astype(str).str.strip().str.upper()

    return df.reset_index(drop=True)

# ==========================
# ONEDRIVE / SHAREPOINT
# ==========================
def onedrive_direct_download(shared_url: str) -> bytes:
    u = urlparse(shared_url); base = f"{u.scheme}://{u.netloc}"
    url1 = f"{base}/_layouts/15/download.aspx?SourceUrl={quote(shared_url, safe='')}"
    r1 = requests.get(url1, allow_redirects=True, timeout=60)
    if r1.ok and r1.content and len(r1.content) > 1000:
        return r1.content
    sep = "&" if "?" in shared_url else "?"
    url2 = f"{shared_url}{sep}download=1"
    r2 = requests.get(url2, allow_redirects=True, timeout=60)
    if r2.ok and r2.content and len(r2.content) > 1000:
        return r2.content
    raise RuntimeError("Não foi possível baixar a planilha. Verifique o compartilhamento do link no OneDrive.")

# ==========================
# PDF (layout)
# ==========================
MARGIN_L = 1.6*cm
MARGIN_R = 1.6*cm
COL_LABEL_X = MARGIN_L + 0.9*cm
COL_VALUE_X = MARGIN_L + 6.3*cm
STEP_Y     = 0.95*cm
ROUND_R    = 8

# RESPIROS/ESPACOS
GAP_AFTER_INFO_TITLE = 0.60*cm    # espaço entre “Informações:” e DATA
GAP_AFTER_FAST_TITLE = 0.60*cm    # espaço entre “FAST CIF/FOB” e CNTR(ORIGEM)
SIG_DELTA_DOWN       = 1.50*cm    # assinaturas mais para baixo
TITLE_SHIFT_LEFT     = 0.50*cm    # título Romaneio Manual 0,5 cm para a esquerda

def format_qtd(val) -> str:
    try:
        f = float(val)
        return str(int(round(f))) if abs(f - round(f)) < 1e-9 else str(f)
    except Exception:
        return str(val)

def str_or_default(row, col, default=""):
    return str(row[col]).strip() if col in row and pd.notna(row[col]) else default

def draw_header(c: canvas.Canvas, data_cabecalho: datetime, page_w, page_h):
    c.setFillColor(colors.black); c.setStrokeColor(colors.black)

    # Logo
    if os.path.exists(LOGO_IMAGE_PATH):
        try:
            c.drawImage(LOGO_IMAGE_PATH, MARGIN_L, page_h - 2.35*cm,
                        width=3.9*cm, height=1.25*cm, preserveAspectRatio=True, mask='auto')
        except Exception:
            pass

    # Título CENTRAL deslocado 0,5 cm para a ESQUERDA
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(page_w/2 - TITLE_SHIFT_LEFT, page_h - 2.7*cm, "Romaneio Manual")

    # Caixa meta (direita)
    box_w, box_h = 5.4*cm, 3.0*cm
    box_x = page_w - MARGIN_R - box_w
    box_y = page_h - 2.0*cm - box_h
    c.setLineWidth(1.5)
    c.roundRect(box_x, box_y, box_w, box_h, 6, stroke=1, fill=0)

    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(box_x + box_w/2, box_y + box_h - 0.9*cm, "Romaneio")

    c.setFont("Helvetica", 10.5)
    c.drawRightString(box_x + box_w - 0.3*cm, box_y + box_h - 1.55*cm, f"Cód.: {FORM_COD}")
    c.drawRightString(box_x + box_w - 0.3*cm, box_y + box_h - 2.10*cm, f"Rev.: {FORM_REV}")
    c.drawRightString(box_x + box_w - 0.3*cm, box_y + 0.55*cm,               f"Data: {data_cabecalho.strftime('%d/%m/%Y')}")

def draw_info_section(c, info, page_w, page_h):
    title_h = 1.1*cm
    content_h = 6*STEP_Y
    pad_top, pad_bottom = 0.5*cm, 0.6*cm
    sec_h = title_h + pad_top + GAP_AFTER_INFO_TITLE + content_h + pad_bottom

    top_y = page_h - 6.1*cm
    bottom_y = top_y - sec_h

    c.setLineWidth(1.5)
    c.roundRect(MARGIN_L, bottom_y, page_w - MARGIN_L - MARGIN_R, sec_h, ROUND_R, stroke=1, fill=0)

    c.setFont("Helvetica-Bold", 20)
    c.drawString(MARGIN_L + 0.7*cm, top_y - 0.85*cm, "Informações:")

    y = top_y - title_h - GAP_AFTER_INFO_TITLE
    c.setFont("Helvetica-Bold", 12)
    rows = [
        ("DATA:",           info.get("DATA","")),
        ("CONTEUDO:",       info.get("CONTEUDO","")),
        ("CLIENTE:",        info.get("CLIENTE","")),
        ("TRANSPORTADORA:", info.get("TRANSPORTADORA","")),
        ("MOTORISTA:",      info.get("MOTORISTA","")),
        ("PLACA:",          info.get("PLACA","")),
    ]
    for rot, val in rows:
        c.drawString(COL_LABEL_X, y, rot)
        c.setFont("Helvetica", 12); c.drawString(COL_VALUE_X, y, str(val)); c.setFont("Helvetica-Bold", 12)
        y -= STEP_Y

    return bottom_y - 0.9*cm

def draw_products_section(c, prod_info, start_y, page_w):
    title_h = 1.0*cm
    fields_h = 4*STEP_Y
    lista_h  = 5.0*cm
    pad_top, pad_bottom = 0.5*cm, 0.6*cm
    gap_fields_lista = 0.35*cm

    sec_h = title_h + pad_top + GAP_AFTER_FAST_TITLE + fields_h + gap_fields_lista + lista_h + pad_bottom
    bottom_y = start_y - sec_h

    c.setLineWidth(1.5)
    c.roundRect(MARGIN_L, bottom_y, page_w - MARGIN_L - MARGIN_R, sec_h, ROUND_R, stroke=1, fill=0)

    c.setFont("Helvetica-Bold", 18)
    c.drawString(MARGIN_L + 0.7*cm, start_y - 0.8*cm, PRODUTO_TITULO)

    y = start_y - title_h - GAP_AFTER_FAST_TITLE
    c.setFont("Helvetica-Bold", 12)
    fields = [
        ("CNTR(ORIGEM):", prod_info.get("CNTR_ORIGEM","")),
        ("CNTR(TECADI):", prod_info.get("CNTR_TECADI","")),
        ("SKU:",          prod_info.get("SKU","")),
        ("QTD:",          prod_info.get("QTD","")),
    ]
    for rot, val in fields:
        c.drawString(COL_LABEL_X, y, rot)
        c.setFont("Helvetica", 12); c.drawString(COL_VALUE_X, y, str(val)); c.setFont("Helvetica-Bold", 12)
        y -= STEP_Y

    y -= gap_fields_lista
    c.drawString(COL_LABEL_X, y, "LISTA:")
    y -= 0.3*cm

    lista_texto = str(prod_info.get("LISTA","")).replace(" ;", ";").replace("  ", " ")
    style = ParagraphStyle(name="lista", fontName="Helvetica", fontSize=12, leading=14, textColor=colors.black)
    para = Paragraph(lista_texto, style=style)

    frame_x = COL_LABEL_X
    frame_y = bottom_y + pad_bottom + 0.2*cm
    frame_w = page_w - MARGIN_R - frame_x
    frame_h = lista_h
    Frame(frame_x, frame_y, frame_w, frame_h, showBoundary=0).addFromList([para], c)

    # Assinaturas (mais para baixo)
    sig_y = 2.6*cm - SIG_DELTA_DOWN
    c.setStrokeColor(colors.black); c.setLineWidth(1)
    c.line(2*cm, sig_y, 9*cm, sig_y); c.line(11*cm, sig_y, 18*cm, sig_y)
    c.setFont("Helvetica", 10)
    c.drawCentredString(5.5*cm, sig_y - 0.45*cm, "ASS: MOTORISTA")
    c.drawCentredString(14.5*cm, sig_y - 0.45*cm, "ASS: CONFERENTE")

def gerar_pdf_row(row: pd.Series, data_cabecalho: datetime) -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    page_w, page_h = A4

    draw_header(c, data_cabecalho, page_w, page_h)

    info = {
        "DATA": data_cabecalho.strftime("%d/%m/%Y"),
        "CONTEUDO": CONTEUDO_FIXO,
        "CLIENTE": CLIENTE_FIXO,
        "TRANSPORTADORA": str_or_default(row, "TRANSPORTADORA"),
        "MOTORISTA":      str_or_default(row, "NOME"),
        "PLACA":          str_or_default(row, "PLACA"),
    }
    next_y = draw_info_section(c, info, page_w, page_h)

    prod = {
        "CNTR_ORIGEM": str_or_default(row, "ARMADOR"),
        "CNTR_TECADI": str_or_default(row, "TECADI"),
        "SKU":         str_or_default(row, "SKU"),
        "QTD":         format_qtd(str_or_default(row, "QTD")),
        "LISTA":       str_or_default(row, "LISTA"),
    }
    draw_products_section(c, prod, next_y, page_w)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf.read()

# ==========================
# UI — PRIMEIRA CHAMADA
# ==========================
st.set_page_config(page_title=APP_TITLE, layout="centered")

# Fundo do APP + legibilidade
def set_app_background(image_path: str):
    if not os.path.exists(image_path): return
    with open(image_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    st.markdown(f"""
        <style>
        header[data-testid="stHeader"] {{ display:none; }}
        div[data-testid="stToolbar"] {{ display:none; }}
        #MainMenu {{ visibility:hidden; }}
        footer {{ visibility:hidden; }}
        .stApp {{
            background-image: url("data:image/jpg;base64,{b64}");
            background-size: cover; background-position: center center; background-attachment: fixed;
        }}
        .block-container {{
            background: rgba(0,0,0,0.35);
            border-radius: 12px;
            padding: 24px 28px;
            color: #fff !important;
        }}
        .stTextInput>div>div>input {{ color:#000 !important; background:#fff !important; }}
        .stButton>button, .stDownloadButton>button {{
            color:#000 !important; background:#fff !important;
            border:1px solid #ccc !important; border-radius:10px; padding:8px 16px; font-weight:700;
        }}
        </style>
    """, unsafe_allow_html=True)

set_app_background(BG_IMAGE_PATH)

st.title("🧾 Tecadi: Romaneio Manual — FG V1")
st.caption("Gerador de PDFs dos FASTFOBS/CIFS STATUS = FINALIZADO.")

# Data PT-BR
default_data = date.today().strftime("%d/%m/%Y")
data_str = st.text_input("Data do cabeçalho do PDF (dd/mm/aaaa)", value=default_data)
def parse_data_ptbr(s: str) -> datetime:
    return datetime.strptime(s.strip(), "%d/%m/%Y")
try:
    data_pdf = parse_data_ptbr(data_str)
    st.caption(f"Data selecionada: {data_pdf.strftime('%d/%m/%Y')}")
except Exception:
    st.error("Informe a data no formato dd/mm/aaaa.")
    st.stop()

col_a, col_b = st.columns([1,1])
with col_a:
    bt_atualizar = st.button("🔄 Atualizar informações do OneDrive", use_container_width=True)
with col_b:
    bt_gerar = st.button("🧾 Gerar PDFs (FINALIZADO)", use_container_width=True)

if "df_cache" not in st.session_state: st.session_state.df_cache = None
if "fetch_error" not in st.session_state: st.session_state.fetch_error = None

if bt_atualizar:
    try:
        xls_bytes = onedrive_direct_download(ONEDRIVE_LINK)
        st.session_state.df_cache = load_excel_from_bytes(xls_bytes)
        st.session_state.fetch_error = None
        # opcional: ao atualizar, limpa seleção anterior
        # st.session_state["selected_keys"] = []
    except Exception as e:
        st.session_state.df_cache = None
        st.session_state.fetch_error = str(e)

if st.session_state.fetch_error:
    st.error("Falha ao baixar do OneDrive: " + st.session_state.fetch_error)

# Resumo enxuto em tela
if st.session_state.df_cache is not None:
    df = st.session_state.df_cache
    df_fin = df[df["STATUS"].astype(str).str.upper().eq("FINALIZADO")].copy()

    st.markdown(f"### ✅ FAST FOB finalizados: **{len(df_fin)}**")

    cols_view = ["DEMANDA","TRANSPORTADORA","NOME","PLACA","LISTA"]
    cols_existing = [c for c in cols_view if c in df_fin.columns]
    if cols_existing:
        df_view = df_fin[cols_existing].copy()
        for c in df_view.columns:
            df_view[c] = df_view[c].astype("string")
        st.dataframe(df_view.reset_index(drop=True), use_container_width=True, hide_index=True)
    else:
        st.warning("Não encontrei as colunas esperadas para exibição (DEMANDA, TRANSPORTADORA, NOME, PLACA, LISTA).")

    # ==========================
    # SELEÇÃO MANUAL (NOVO)
    # ==========================
    def _make_key(row: pd.Series) -> str:
        demanda_raw = str(row.get("DEMANDA", ""))
        m = re.search(r"(\d+)", demanda_raw)
        demanda_num = m.group(1) if m else demanda_raw
        # se quiser diferenciar por LISTA também, inclua: + '|' + str(row.get("LISTA",""))
        return f"{demanda_num}|{str(row.get('ARMADOR',''))}"

    df_fin = df_fin.copy()
    df_fin["_key"] = df_fin.apply(_make_key, axis=1)

    labels = {}
    for _, r in df_fin.iterrows():
        label = " — ".join([
            str(r.get("DEMANDA","")),
            str(r.get("ARMADOR","")),
            str(r.get("TRANSPORTADORA","")),
            str(r.get("NOME","")),
            str(r.get("PLACA","")),
            str(r.get("LISTA","")),
        ])
        labels[r["_key"]] = label

    selected_keys = st.multiselect(
        "Selecione as linhas que deseja gerar (opcional)",
        options=list(labels.keys()),
        format_func=lambda k: labels[k],
        placeholder="Escolha um Carregamento",
    )
    st.session_state["selected_keys"] = selected_keys

if bt_gerar:
    df = st.session_state.df_cache
    if df is None:
        st.error("Clique em **Atualizar informações do OneDrive** primeiro.")
    else:
        df_ok = df[df["STATUS"].astype(str).str.upper().eq("FINALIZADO")].copy()

        # === aplica filtro pela seleção manual (se houver) ===
        def _make_key(row: pd.Series) -> str:
            demanda_raw = str(row.get("DEMANDA", ""))
            m = re.search(r"(\d+)", demanda_raw)
            demanda_num = m.group(1) if m else demanda_raw
            return f"{demanda_num}|{str(row.get('ARMADOR',''))}"

        df_ok["_key"] = df_ok.apply(_make_key, axis=1)
        selected = st.session_state.get("selected_keys") or []
        if selected:
            df_ok = df_ok[df_ok["_key"].isin(selected)]

        faltam = [c for c in REQUIRED_FOR_PDF if c not in df_ok.columns]
        if faltam:
            st.error("Faltam colunas obrigatórias: " + ", ".join(faltam))
        elif df_ok.empty:
            st.warning("Nenhuma linha com STATUS = FINALIZADO (ou nada selecionado).")
        else:
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for i, row in df_ok.iterrows():
                    try:
                        pdf_bytes = gerar_pdf_row(row, data_pdf)
                        demanda_raw = str_or_default(row, "DEMANDA")
                        m = re.search(r"(\d+)", demanda_raw); demanda_num = m.group(1) if m else demanda_raw
                        armador = str_or_default(row, "ARMADOR")
                        fname = f"{demanda_num}.{armador}.pdf"
                        zf.writestr(fname, pdf_bytes)
                    except Exception as e:
                        zf.writestr(f"ERRO_{i}.txt", f"Falha na linha {i}: {repr(e)}")
            zip_buf.seek(0)
            st.download_button(
                "⬇️ Baixar ZIP com PDFs",
                data=zip_buf,
                file_name=f"romaneios_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip"
            )
