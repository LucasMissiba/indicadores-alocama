from pathlib import Path
from typing import List, Optional, Tuple, Dict
import re
import unicodedata
import hashlib
import time
import base64
import os
import logging

import pandas as pd
import io
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st

# Configuração de logging simplificada para Streamlit Cloud
logging.basicConfig(level=logging.INFO)

# Configuração da página
st.set_page_config(
    page_title="Dashboard Alocama",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Configurações de segurança removidas para compatibilidade com Streamlit Cloud


# Statsmodels (opcional): previsão Holt-Winters
try:
    from statsmodels.tsa.holtwinters import ExponentialSmoothing  # type: ignore
except Exception:
    ExponentialSmoothing = None  # fallback será aplicado


APP_TITLE = "Dashboard de Contratos | Alocama"
OUTPUT_FILENAME = "resultado_itens.xlsx"
SMART_FALLBACK_CANDIDATES = [
    "item",
    "produto",
    "produtos",
    "descricao",
    "descrição",
    "descrição do produto",
    "descricao do produto",
    "produto/serviço",
    "produto/servico",
    "nome do item",
]

def validate_file_security(file_path: Path) -> bool:
    """Valida segurança do arquivo antes da leitura."""
    try:
        # Verificar se o arquivo existe
        if not file_path.exists():
            logging.warning(f"Arquivo não encontrado: {file_path}")
            return False
        
        # Verificar extensão
        allowed_extensions = ['.xlsx', '.xls']
        if file_path.suffix.lower() not in allowed_extensions:
            logging.warning(f"Extensão não permitida: {file_path.suffix}")
            return False
        
        # Verificar tamanho do arquivo (máximo 50MB)
        file_size = file_path.stat().st_size
        max_size = 50 * 1024 * 1024  # 50MB
        if file_size > max_size:
            logging.warning(f"Arquivo muito grande: {file_size} bytes")
            return False
        
        # Verificar se o arquivo não está vazio
        if file_size == 0:
            logging.warning(f"Arquivo vazio: {file_path}")
            return False
            
        return True
    except Exception as e:
        logging.error(f"Erro na validação de segurança: {e}")
        return False

def safe_read_excel(file_path: Path, **kwargs) -> Optional[pd.ExcelFile]:
    """Leitura segura de arquivos Excel com validações."""
    try:
        if not validate_file_security(file_path):
            return None
        
        # Log da operação
        logging.info(f"Lendo arquivo: {file_path.name}")
        
        # Leitura com timeout implícito
        book = pd.read_excel(file_path, **kwargs)
        return book
        
    except Exception as e:
        logging.error(f"Erro ao ler arquivo {file_path.name}: {str(e)}")
        return None

def sanitize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Sanitiza dados do DataFrame para segurança."""
    try:
        # Remover colunas com nomes suspeitos
        suspicious_patterns = ['<script', 'javascript:', 'onload=', 'onerror=']
        for col in df.columns:
            col_str = str(col).lower()
            if any(pattern in col_str for pattern in suspicious_patterns):
                logging.warning(f"Coluna suspeita removida: {col}")
                df = df.drop(columns=[col])
        
        # Limitar número de linhas (máximo 100.000)
        if len(df) > 100000:
            logging.warning(f"DataFrame truncado de {len(df)} para 100.000 linhas")
            df = df.head(100000)
        
        # Limitar número de colunas (máximo 100)
        if len(df.columns) > 100:
            logging.warning(f"DataFrame truncado de {len(df.columns)} para 100 colunas")
            df = df.iloc[:, :100]
        
        return df
    except Exception as e:
        logging.error(f"Erro na sanitização: {e}")
        return df

def render_company_selector(groups: List[str]) -> Optional[str]:
    """Painel de seleção de empresa, escalável para 200+. 
    - Até 6 empresas: botões horizontais
    - Maior que 6: selectbox com busca
    Persiste valor em session_state.
    """
    if not groups:
        return None
    # Estilo cartão
    st.markdown(
        """
        <style>
        .control-card{padding:10px 14px;border-radius:12px;background:rgba(255,255,255,0.04);
            border:1px solid rgba(255,255,255,0.08); margin:8px 0 6px 0}
        .control-title{font-weight:600;margin-bottom:6px;opacity:.9}
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.markdown("<div class='control-card'><div class='control-title'>Empresa</div>", unsafe_allow_html=True)
    selected_key = st.session_state.get("empresa_atual") or st.session_state.get("grupo_unico")
    if len(groups) <= 6:
        # Fallback para radio horizontal (compatível); persiste índice
        default = groups.index(selected_key) if selected_key in groups else 0
        choice = st.radio("Selecione o Grupo", options=groups, index=default, horizontal=True, label_visibility="collapsed", key="grupo_unico")
    else:
        default = groups.index(selected_key) if selected_key in groups else 0
        choice = st.selectbox("Selecionar empresa", options=groups, index=default, key="empresa_select")
        st.session_state["grupo_unico"] = choice
    st.session_state["empresa_atual"] = choice
    st.markdown("</div>", unsafe_allow_html=True)
    return choice

NAME_FALLBACK_CANDIDATES = [
    "paciente",
    "beneficiario",
    "beneficiário",
    "nome",
    "nome do paciente",
    "nome paciente",
    "usuario",
    "usuário",
    "assistido",
    "vida",
    "cliente",
]

def _build_hero_media_html() -> str:
    """Retorna HTML do vídeo/imagem (base64) para embutir dentro do hero.

    Procura `assets/sideview.webm`, depois `assets/sideview.mp4`, depois `assets/sideview.gif`.
    """
    try:
        assets_dir = Path.cwd() / "assets"
        for name in ["sideview.webm", "sideview.mp4", "sideview.gif"]:
            p = assets_dir / name
            if p.exists():
                suffix = p.suffix.lower()
                mime = "video/webm" if suffix == ".webm" else ("video/mp4" if suffix == ".mp4" else "image/gif")
                data_b64 = base64.b64encode(p.read_bytes()).decode("utf-8")
                src_attr = f"data:{mime};base64,{data_b64}"
                if mime.startswith("video/"):
                    return f"<div class='page-hero__bg'><video src='{src_attr}' autoplay muted loop playsinline></video></div>"
                return f"<div class='page-hero__bg'><img src='{src_attr}' alt='bg'/></div>"
    except Exception:
        pass
    return ""


# =====================
# UI helpers – KPI cards
# =====================
def render_kpi_card(title: str, value: str, subtitle: str = "", bar_pct: float = None, color: str = "#4e79a7") -> None:
    """Desenha um card de KPI com barra de progresso opcional."""
    bar_html = ""
    if bar_pct is not None:
        pct = max(0, min(100, float(bar_pct)))
        bar_html = f"""
        <div style='height:6px;background:rgba(255,255,255,.08);border-radius:6px;margin-top:6px;'>
          <div style='height:6px;width:{pct}%;background:{color};border-radius:6px'></div>
        </div>
        """
    st.markdown(
        f"""
        <div style='padding:12px 14px;border-radius:12px;background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08)'>
          <div style='font-size:12px;opacity:.82'>{title}</div>
          <div style='font-size:22px;font-weight:700;margin-top:2px'>{value}</div>
          <div style='font-size:11px;opacity:.7'>{subtitle}</div>
          {bar_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def format_currency(v: float) -> str:
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"


def _price_map_for_company(company: str) -> Dict[str, float]:
    key = (company or "").upper()
    if key == "AXX CARE":
        return {
            normalize_text_for_match("CAMA ELÉTRICA 3 MOVIMENTOS"): 10.80,
            normalize_text_for_match("CAMA MANUAL 2 MANIVELAS"): 2.83,
            normalize_text_for_match("SUPORTE DE SORO"): 0.67,
        }
    # Pode-se estender para outras empresas se necessário
    return {}


def compute_kpis_for_company(df_emp_viz: pd.DataFrame, empresa: str, last_month_order: Dict[str, int], sel_files: List[Path]) -> Dict[str, float]:
    kpis: Dict[str, float] = {}
    df_e = df_emp_viz[df_emp_viz["Empresa"].str.upper() == (empresa or "").upper()].copy()
    if df_e.empty:
        return {"qtd": 0, "itens": 0, "vidas": 0, "fat": 0.0, "mes": "-"}
    # último mês
    ultimo = df_e["Mês"].map(last_month_order).max()
    mes_label = [k for k, v in last_month_order.items() if v == ultimo]
    mes_label = mes_label[0] if mes_label else df_e["Mês"].iloc[-1]
    df_last = df_e[df_e["Mês"] == mes_label]
    kpis["qtd"] = int(df_last["Quantidade"].sum())
    kpis["itens"] = int(df_last["Item"].nunique())
    kpis["mes"] = mes_label

    # vidas ativas (únicos na coluna B) no mês
    ym_map = {"Janeiro": "2025-01", "Fevereiro": "2025-02", "Março": "2025-03", "Abril": "2025-04", "Maio": "2025-05", "Junho": "2025-06", "Julho": "2025-07", "Agosto": "2025-08"}
    alvo_ym = ym_map.get(mes_label)
    vidas_set = set()
    if alvo_ym:
        for f in sel_files:
            try:
                if primary_group_from_label(str(f)).upper() != (empresa or "").upper():
                    continue
                if year_month_from_path(f) != alvo_ym:
                    continue
                book = safe_read_excel(f, sheet_name=None)
            except Exception:
                continue
            for sh, df in (book or {}).items():
                if should_exclude_sheet(str(sh)) or not isinstance(df, pd.DataFrame) or df.empty:
                    continue
                series = None
                try:
                    if df.shape[1] >= 2:
                        series = df.iloc[:, 1]
                except Exception:
                    series = None
                if series is None:
                    name_col = select_best_name_column(df)
                    if not name_col:
                        continue
                    series = df[name_col]
                s = series.dropna().astype(str).str.strip()
                s = s[s != ""]
                vidas_set.update(s.map(normalize_text_for_match).tolist())
    kpis["vidas"] = len(vidas_set)

    # faturamento estimado (se houver mapa de preços)
    pmap = _price_map_for_company(empresa)
    if pmap:
        df_tmp = df_last.copy()
        df_tmp["key"] = df_tmp["Item"].map(normalize_text_for_match)
        df_tmp["PrecoDiaria"] = df_tmp["key"].map(pmap)
        df_tmp = df_tmp.dropna(subset=["PrecoDiaria"])  # só itens tarifados
        dias_map = {"Fevereiro": 28, "Março": 31, "Abril": 30, "Maio": 31, "Junho": 30, "Julho": 31, "Agosto": 31}
        df_tmp["Dias"] = df_tmp["Mês"].map(dias_map).fillna(30)
        kpis["fat"] = float((df_tmp["Quantidade"] * df_tmp["PrecoDiaria"] * df_tmp["Dias"]).sum())
    else:
        kpis["fat"] = 0.0
    return kpis


def hash_password(raw: str) -> str:
    try:
        return hashlib.sha256(raw.encode("utf-8")).hexdigest()
    except Exception:
        return ""


def get_auth_users() -> Dict[str, str]:
    """Retorna {usuario_lower: valor} em que valor pode ser:
    - "sha256:<hash>" (recomendado)
    - "<hash>" (compatível)
    - "plain:<senha>" (útil para testes rápidos no Cloud)
    """
    users_lower: Dict[str, str] = {}
    try:
        for k, v in (st.secrets.get("auth_users", {}) or {}).items():
            key = str(k).strip().lower()
            val = str(v)
            if val.startswith("sha256:"):
                val = val.split(":", 1)[1]
            # plain:<senha> é mantido como está
            users_lower[key] = val
    except Exception:
        users_lower = {}
    if not users_lower:
        users_lower = {"admin": hash_password("admin")}
    return users_lower


def verify_credentials(username: str, password: str) -> bool:
    users = get_auth_users()
    user_key = (username or "").strip().lower()
    pwd = (password or "")
    if not user_key or not pwd:
        return False
    stored = users.get(user_key)
    if not stored:
        return False
    if stored.startswith("plain:"):
        return password == stored.split(":", 1)[1]
    return stored == hash_password(password)


def render_splash_once() -> bool:
    """Mostra uma splash de carregamento elegante apenas na primeira visita."""
    if "splash_shown" not in st.session_state:
        st.session_state["splash_shown"] = False
    if st.session_state["splash_shown"]:
        return False
    ph = st.empty()
    ph.markdown(
        """
        <style>
        .splash-overlay{position:fixed;inset:0;display:flex;align-items:center;justify-content:center;background:#1f2a3a;z-index:999999;}
        .splash-card{padding:32px 40px;border-radius:16px;background:#1f2a3a;color:#fff;font-family:system-ui,Segoe UI,Roboto,Ubuntu,\"Helvetica Neue\",Arial}
        .brand{font-weight:700;font-size:20px;letter-spacing:.4px;margin-bottom:6px;color:#87ceeb}
        .title{font-size:28px;margin:0 0 10px 0}
        .subtitle{opacity:.9;margin-bottom:18px}
        .loader{width:96px;height:96px;border-radius:50%;margin:14px auto;border:6px solid rgba(255,255,255,.18);border-top-color:#4cc9f0;animation:spin .9s linear infinite}
        @keyframes spin{to{transform:rotate(360deg)}}
        .lgpd{margin-top:8px;font-size:12px;color:#c7d3df}
        </style>
        <div class='splash-overlay'>
          <div class='splash-card'>
            <div class='brand'>Alocama · Setor de Contratos</div>
            <div class='title'>Preparando seu painel...</div>
            <div class='subtitle'>Carregando componentes e verificando ambiente</div>
            <div class='loader'></div>
            <div class='lgpd'>Respeitamos a LGPD e tratamos dados com responsabilidade.</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    time.sleep(1.2)
    ph.empty()
    st.session_state["splash_shown"] = True
    st.rerun()
    return True


def render_login() -> bool:
    """Login desativado a pedido: sempre retorna True após a splash."""
    st.session_state["authed"] = True
    return True

def clean_item_values(series: pd.Series, selected_col_name: str, only_equipment: bool = False) -> pd.Series:
    """Normaliza e filtra valores não válidos da coluna de itens/produtos."""
    s = series.astype(str).str.strip()
    s = s[s != ""]
    invalid_names = {selected_col_name.lower(), "item", "produto", "produtos", "descrição", "descricao"}
    s = s[~s.str.lower().isin(invalid_names)]
    s = s[~s.str.lower().str.match(r"^(total|subtotal)\b")] 
    norm = s.map(normalize_text_for_match)
    bad_regex = r"(?:valor|pagina|page|quant|\bqtd\b|status|retirada|paciente|periodo|serie|unidade|unidades|\bun\b|\brs\b)"
    s = s[~norm.str.contains(bad_regex, regex=True, na=False)]
    norm = s.map(normalize_text_for_match)
    s = s[norm.str.len() >= 3]
    if only_equipment:
        s = s[norm.str.contains(r"[a-z]", regex=True, na=False)]
    return s


def normalize_text_for_match(text: str) -> str:
    """Remove acentos, baixa caixa e mantém apenas [a-z0-9 ] para facilitar o match."""
    if not isinstance(text, str):
        text = str(text)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower()
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def categorize_item_name(item_name: str) -> str:
    """Classifica o item em categorias macro (CAMA, CADEIRA DE RODAS, etc.)."""
    t = normalize_text_for_match(item_name)

    if "higien" in t or "banho" in t:
        return "CADEIRA HIGIÊNICA"
    if ("cadeira" in t and "rod" in t) or re.search(r"\brodas?\b", t):
        return "CADEIRA DE RODAS"
    if "cama" in t:
        return "CAMA"
    if "colch" in t:
        return "COLCHÃO"
    if "suporte" in t and "soro" in t:
        return "SUPORTE DE SORO"
    if "andador" in t:
        return "ANDADOR"
    if "muleta" in t:
        return "MULETA"
    if "bengala" in t:
        return "BENGALA"
    if (
        "oximetro" in t
        or "oximeter" in t
        or "cpap" in t
        or "bipap" in t
        or "ventilador" in t
        or "nebul" in t
        or "aspirador" in t
    ):
        return "RESPIRATÓRIO"

    return "OUTROS"


def attach_categories(df_result: pd.DataFrame) -> pd.DataFrame:
    """Adiciona coluna Categoria ao df de resultados por arquivo+item."""
    df = df_result.copy()
    df["Categoria"] = df["Item"].map(categorize_item_name)
    return df


def canonicalize_trio_item(name: str) -> str:
    """Unifica variações textuais dos 3 itens mais comuns para melhorar o ranking/top3.
    - CAMA MANUAL 2 MANIVELAS
    - CAMA ELÉTRICA 3 MOVIMENTOS
    - SUPORTE DE SORO
    Caso não case, retorna o nome original.
    """
    t = normalize_text_for_match(name)
    if ("cama" in t and "manual" in t and ("2" in t or "ii" in t) and "manivel" in t):
        return "CAMA MANUAL 2 MANIVELAS"
    if ("cama" in t and ("eletric" in t or "elétrica" in t or "eletrica" in t) and ("3" in t or "iii" in t) and ("mov" in t or "movimento" in t)):
        return "CAMA ELÉTRICA 3 MOVIMENTOS"
    if ("suporte" in t and "soro" in t):
        return "SUPORTE DE SORO"
    return name


def canonicalize_electric_bed_two_movements(name: str) -> str:
    """
    Unifica variações de "CAMA ELÉTRICA 2 MOVIMENTOS" (ex.: tamanhos/sufixos como 2,10Mts).
    Mantém demais itens inalterados.
    """
    t = normalize_text_for_match(name)
    if ("cama" in t and ("eletric" in t or "eletrica" in t) and ("2" in t or "ii" in t) and ("mov" in t or "movimento" in t)):
        return "CAMA ELÉTRICA 2 MOVIMENTOS"
    return name


def canonicalize_wheelchair_group(name: str) -> str:
    """
    - Se for cadeira de rodas com indicação de reclinável e tamanho 40..48 → "CADEIRA DE RODAS RECLINAVEL".
    - Se for cadeira de rodas (não reclinável) e tamanho 40..48 → "CADEIRA DE RODAS SIMPLES".
    - Caso contrário, mantém o nome original.
    """
    t = normalize_text_for_match(name)
    if not ("cadeira" in t and ("rod" in t or re.search(r"\brodas?\b", t))):
        return name
    tokens = re.findall(r"(\d{2})", t)
    has_38_48 = any(38 <= int(tok) <= 48 for tok in tokens if tok.isdigit())
    if not has_38_48:
        return name
    is_reclinavel = "reclin" in t  # cobre 'reclinavel'/'reclinável'
    return "CADEIRA DE RODAS RECLINAVEL" if is_reclinavel else "CADEIRA DE RODAS SIMPLES"


def canonicalize_wheelchair_obese_60(name: str) -> str:
    """Unifica variações de cadeiras 'obeso' 60 e correlatas.
    - 'CADEIRA DE RODAS OBESO SIMPLES ALUM 60 - 140KG' → 'CADEIRA DE RODAS OBESO SIMPLES 60'
    - 'CADEIRA DE RODAS OBESO ORTOBRAS ALUM 60 - 200KG' → idem
    - 'CADEIRA DE RODAS 65' → idem (solicitado)
    """
    t = normalize_text_for_match(name)
    if not ("cadeira" in t and ("rod" in t or re.search(r"\brodas?\b", t))):
        return name
    # Mapa por condições
    tokens = re.findall(r"(\d{2})", t)
    has_60 = any(tok == "60" for tok in tokens)
    has_65 = any(tok == "65" for tok in tokens)
    if ("obes" in t and has_60) or ("ortobras" in t and has_60) or has_65:
        return "CADEIRA DE RODAS OBESO SIMPLES 60"
    return name


def canonicalize_wheelchair_50(name: str) -> str:
    """Normaliza cadeiras de rodas 50/50,5 → 'CADEIRA DE RODAS 50'."""
    t = normalize_text_for_match(name)
    if not ("cadeira" in t and ("rod" in t or re.search(r"\brodas?\b", t))):
        return name
    # detecta 50 ou 50,5 / 50.5
    m = re.search(r"\b50([\.,]5)?\b", t)
    if m:
        return "CADEIRA DE RODAS 50"
    return name


def canonicalize_walker(name: str) -> str:
    """Agrupa quaisquer variações de andadores como 'ANDADOR'."""
    t = normalize_text_for_match(name)
    if "andador" in t:
        return "ANDADOR"
    return name


def canonicalize_bed_alt_trem(name: str) -> str:
    """Unifica variações de CAMA ELÉTRICA ALT. TREM (com dimensões/typos)."""
    t = normalize_text_for_match(name)
    if "cama" in t and "alt" in t and "trem" in t:
        return "CAMA ELÉTRICA ALT. TREM"
    return name

def infer_group_for_label(label: str, candidates: List[str]) -> str:
    """Inferência robusta do grupo a partir do caminho (aceita \\ ou / e variações)."""
    parts = re.split(r"[\\/]+", str(label))
    parts_norm = [normalize_text_for_match(p) for p in parts]
    norm_candidates = {normalize_text_for_match(c): c for c in candidates}

    if "grupo solar" in parts_norm:
        idx = parts_norm.index("grupo solar")
        if idx + 1 < len(parts_norm):
            nxt_norm = parts_norm[idx + 1]
            if nxt_norm in norm_candidates:
                return norm_candidates[nxt_norm]
            return parts[idx + 1]

    for p_norm, p in zip(parts_norm, parts):
        if p_norm in norm_candidates:
            return norm_candidates[p_norm]

    if any("hospital" in p for p in parts_norm):
        return "HOSPITALAR"
    if any("dommus" in p or "domus" in p for p in parts_norm):
        return "DOMMUS"
    if any("solar" in p for p in parts_norm):
        return "SOLAR"

    return parts[0] if parts else ""


def primary_group_from_label(label: str) -> str:
    """Inferência robusta da empresa a partir do caminho relativo.

    Prioriza a detecção por substring ('dommus'/'domus', 'hospital', 'solar', 'pronep').
    Caso não bata, usa o primeiro segmento do caminho.
    """
    s_norm = normalize_text_for_match(str(label))
    if "dommus" in s_norm or "domus" in s_norm:
        return "DOMMUS"
    if "hospital" in s_norm:
        return "HOSPITALAR"
    if "solar" in s_norm:
        return "SOLAR"
    if "pronep" in s_norm:
        return "PRONEP"
    parts = re.split(r"[\\/]+", str(label).strip())
    return (parts[0].upper() if parts and parts[0] else "").upper()


def month_from_path(path: Path) -> Optional[str]:
    """Retorna '1'..'8' se o caminho contiver:
    - pastas 1..8
    - padrão 2025-01..2025-08
    - nome do mês em português (ex.: 'janeiro', 'fevereiro', 'marco', 'março', ...)
    """
    parts = re.split(r"[\\/]+", str(path))
    text_norm = normalize_text_for_match(str(path))
    month_words = {
        "janeiro": "1",
        "fevereiro": "2",
        "marco": "3",
        "marco": "3",
        "marco": "3",
        "março": "3",
        "abril": "4",
        "maio": "5",
        "junho": "6",
        "julho": "7",
        "agosto": "8",
    }
    for p in parts:
        p_norm = p.strip()
        if re.fullmatch(r"0?[12345678]", p_norm):
            return p_norm.lstrip("0")
        m = re.fullmatch(r"\d{4}-(0[12345678])", p_norm)
        if m:
            return m.group(1).lstrip("0")
    for word, num in month_words.items():
        if word in text_norm:
            return num
    return None


def year_month_from_path(path: Path) -> Optional[str]:
    """Retorna 'YYYY-MM' se algum segmento do caminho estiver neste formato.
    Caso não exista, tenta inferir pelo nome do mês em PT-BR (assumindo ano 2025)."""
    parts = re.split(r"[\\/]+", str(path))
    for p in parts:
        p_norm = p.strip()
        m = re.fullmatch(r"(20\d{2})-(0[1-9]|1[0-2])", p_norm)
        if m:
            return m.group(0)
    # Fallback por nome do mês
    text_norm = normalize_text_for_match(str(path))
    word_to_mm = {
        "fevereiro": "02",
        "marco": "03",
        "março": "03",
        "abril": "04",
        "maio": "05",
        "junho": "06",
        "julho": "07",
        "agosto": "08",
    }
    for word, mm in word_to_mm.items():
        if word in text_norm:
            return f"2025-{mm}"
    return None


def render_top3_pies(df_by_file: pd.DataFrame, group_names: Optional[List[str]] = None) -> None:
    """Renderiza gráficos de pizza (Top 3 itens) para cada grupo informado."""
    if df_by_file.empty:
        return
    df = df_by_file.copy()
    if "Grupo" not in df.columns:
        df["Grupo"] = df["Arquivo"].apply(lambda s: infer_group_for_label(str(s), group_names))
    if not group_names:
        group_names = sorted([g for g in df["Grupo"].unique().tolist() if str(g).strip() != ""])
    df_with_cat = attach_categories(df)

    st.subheader("Top 3 itens por grupo")
    cols = st.columns(min(3, len(group_names)) or 1)
    col_idx = 0
    for group in (group_names or []):
        df_g_raw = df[df["Grupo"] == group]
        df_g = df_with_cat[df_with_cat["Grupo"] == group]
        df_g = df_g[df_g["Categoria"] != "OUTROS"]
        if df_g.empty:
            df_g = df_g_raw
        if df_g.empty:
            continue
        top3 = (
            df_g.groupby("Item", as_index=False, observed=True)["Quantidade"].sum().sort_values("Quantidade", ascending=False).head(3)
        )
        if top3.empty:
            continue
        fig = px.pie(top3, names="Item", values="Quantidade", title=f"Top 3 - {group}", hole=0.3)
        cols[col_idx % len(cols)].plotly_chart(fig, use_container_width=True)
        col_idx += 1


def should_exclude_sheet(sheet_name: str) -> bool:
    """Determina se uma aba deve ser ignorada (ex.: resumos/gráficos/totais)."""
    s = normalize_text_for_match(sheet_name)
    patterns = [
        "resumo",
        "totais",
        "total",
        "grafico",
        "graficos",
        "chart",
        "pivot",
        "tabela dinamica",
        # Mantemos somente folhas de resumo/visual/dicionário. Permitimos 'base', 'cadastro', 'validacao'.
        "mapeamento",
        "mapa",
        "dicion",
    ]
    return any(p in s for p in patterns)


def is_excel_file(path: Path) -> bool:
    """Retorna True se for um arquivo Excel válido (xlsx/xlsm, case-insensitive), exclui temporários e saída."""
    if not path.is_file():
        return False
    name_lower = path.name.lower()
    if not (name_lower.endswith(".xlsx") or name_lower.endswith(".xlsm")):
        return False
    if name_lower.startswith("~$"):
        return False
    if path.name == OUTPUT_FILENAME:
        return False
    return True


def list_excel_files(directory: Path, recursive: bool = False) -> List[Path]:
    """Lista todos os .xlsx na pasta (e opcionalmente subpastas), exceto temporários e o arquivo de saída."""
    files: List[Path] = []
    if recursive:
        for entry in directory.rglob("*"):
            if is_excel_file(entry):
                files.append(entry)
    else:
        for entry in directory.iterdir():
            if is_excel_file(entry):
                files.append(entry)
    seen = set()
    deduped: List[Path] = []
    for p in files:
        key = str(p.resolve()).lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(p)
    return sorted(deduped)


def compute_file_hash(path: Path, chunk_size: int = 1024 * 1024) -> Optional[str]:
    """Calcula hash SHA1 do arquivo (para deduplicação por conteúdo)."""
    try:
        h = hashlib.sha1()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(chunk_size), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return None


def deduplicate_files_by_content(files: List[Path]) -> Tuple[List[Path], Dict[str, List[str]]]:
    """Remove arquivos duplicados por conteúdo, mantendo o primeiro encontrado.

    Retorna (lista_sem_duplicatas, mapa_hash->lista_de_caminhos_descartados)
    """
    kept: List[Path] = []
    seen_hash_to_path: Dict[str, Path] = {}
    duplicates: Dict[str, List[str]] = {}
    for p in sorted(files):
        file_hash = compute_file_hash(p)
        if not file_hash:
            kept.append(p)
            continue
        if file_hash in seen_hash_to_path:
            duplicates.setdefault(file_hash, []).append(str(p))
            continue
        seen_hash_to_path[file_hash] = p
        kept.append(p)
    return kept, duplicates


def normalize_column_name(name: str) -> str:
    return str(name).strip().lower()


def find_matching_column(columns: List[str], target: str) -> Optional[str]:
    """Retorna o nome exato da coluna que corresponde ao alvo (case-insensitive)."""
    target_norm = normalize_column_name(target)
    for col in columns:
        if normalize_column_name(col) == target_norm:
            return col
    return None


def excel_letter_to_index(selector: str) -> Optional[int]:
    """Converte letras de coluna do Excel (ex.: 'E', 'AA') para índice 0-based."""
    if not selector:
        return None
    s = selector.strip().upper()
    if not re.fullmatch(r"[A-Z]+", s):
        return None
    index = 0
    for ch in s:
        index = index * 26 + (ord(ch) - ord('A') + 1)
    return index - 1


def resolve_column_selector(columns: List[str], selector: str) -> Optional[str]:
    """Resolve o seletor de coluna que pode ser nome, letra (A..Z) ou número (1..N)."""
    if not selector:
        return None
    s = str(selector).strip()
    if s.isdigit():
        pos = int(s) - 1
        if 0 <= pos < len(columns):
            return columns[pos]
    pos_from_letter = excel_letter_to_index(s)
    if pos_from_letter is not None and 0 <= pos_from_letter < len(columns):
        return columns[pos_from_letter]
    return find_matching_column(columns, s)


def select_best_column(df: pd.DataFrame, selector: Optional[str], use_smart: bool) -> Tuple[Optional[str], str, int]:
    """Seleciona a melhor coluna considerando:
    1) seletor explícito (nome/letra/índice)
    2) nomes candidatos comuns (SMART_FALLBACK_CANDIDATES)
    3) melhor coluna textual com mais valores não vazios

    Retorna (coluna, metodo, num_valores_nao_vazios)
    metodo em {"manual", "smart_name", "smart_fallback", "none"}
    """
    columns = list(map(str, df.columns))

    def non_empty_count(col_name: str) -> int:
        series = df[col_name].dropna()
        if series.empty:
            return 0
        series = series.astype(str).str.strip()
        return int((series != "").sum())

    if selector:
        manual_col = resolve_column_selector(columns, selector)
        if manual_col is not None:
            cnt = non_empty_count(manual_col)
            if cnt > 0:
                return manual_col, "manual", cnt

    if not use_smart:
        return (manual_col if selector else None), "none", 0

    for cand in SMART_FALLBACK_CANDIDATES:
        match = find_matching_column(columns, cand)
        if match is not None:
            cnt = non_empty_count(match)
            if cnt > 0:
                return match, "smart_name", cnt

    best_col: Optional[str] = None
    best_cnt = 0
    for col in columns:
        try:
            cnt = non_empty_count(col)
        except Exception:
            continue
        if cnt > best_cnt:
            best_col = col
            best_cnt = cnt
    if best_col and best_cnt > 0:
        return best_col, "smart_fallback", best_cnt

    return (manual_col if selector else None), "none", 0


def select_best_name_column(df: pd.DataFrame) -> Optional[str]:
    """Seleciona a melhor coluna de 'nome do paciente/vida'.

    Estratégia em camadas:
    1) Tenta casar pelos nomes candidatos (case-insensitive, com/sem acento)
    2) Se não encontrar, procura colunas cujo nome contenha tokens típicos de nomes
       ("nome", "pacient", "benefici", "usuario", "assistid", "cliente", "vida")
       e escolhe a que possuir mais valores válidos
    3) Como último recurso, escolhe a coluna mais "parecida com nomes":
       maior contagem de valores textuais com letras e pelo menos um espaço
    """
    columns = list(map(str, df.columns))

    def count_valid_names(series: pd.Series) -> int:
        if series is None or series.empty:
            return 0
        s = series.dropna().astype(str).str.strip()
        s = s[s != ""]
        if s.empty:
            return 0
        norm = s.map(normalize_text_for_match)
        looks_like_name = norm.str.contains(r"[a-z]", regex=True, na=False)
        has_space = norm.str.contains(r"\s", regex=True, na=False)
        candidates = norm[looks_like_name & has_space & (norm.str.len() >= 5)]
        return int(candidates.nunique())

    best_col: Optional[str] = None
    best_score = -1
    for cand in NAME_FALLBACK_CANDIDATES:
        col = find_matching_column(columns, cand)
        if col is None:
            continue
        score = count_valid_names(df[col])
        if score > best_score:
            best_score = score
            best_col = col
    if best_col is not None and best_score > 0:
        return best_col

    token_candidates = [
        "nome", "pacient", "benefici", "usuario", "assistid", "cliente", "vida"
    ]
    for col in columns:
        col_norm = normalize_text_for_match(col)
        if any(tok in col_norm for tok in token_candidates):
            score = count_valid_names(df[col])
            if score > best_score:
                best_score = score
                best_col = col
    if best_col is not None and best_score > 0:
        return best_col

    for col in columns:
        try:
            score = count_valid_names(df[col])
        except Exception:
            continue
        if score > best_score:
            best_score = score
            best_col = col
    return best_col


def discover_columns(files: List[Path], max_files: int = 20) -> List[str]:
    """Descobre o conjunto de colunas (união) olhando o cabeçalho das primeiras planilhas."""
    discovered = set()
    for file in files[:max_files]:
        try:
            df_head = safe_read_excel(file, nrows=0)
            for c in df_head.columns:
                discovered.add(str(c))
        except Exception:
            continue
    cols_sorted = sorted(discovered, key=lambda x: normalize_column_name(x))
    if any(normalize_column_name(c) == "item" for c in cols_sorted):
        cols_sorted = [next(c for c in cols_sorted if normalize_column_name(c) == "item")] + [
            c for c in cols_sorted if normalize_column_name(c) != "item"
        ]
    return cols_sorted


def count_items_in_files(
    files: List[Path],
    target_column: str,
    base_dir: Path,
    use_smart: bool = True,
    only_equipment: bool = False,
) -> Tuple[pd.DataFrame, List[str], List[str], Dict[str, List[Tuple[str, str, str, int]]]]:
    """
    Conta os valores da coluna indicada em todas as abas de todos os arquivos.
    Retorno:
    - DataFrame colunas [Arquivo, Item, Quantidade]
    - Lista de arquivos ignorados (sem a coluna)
    - Lista de arquivos com erro de leitura
    """
    rows = []
    ignored_missing_col: List[str] = []
    read_errors: List[str] = []
    column_debug: Dict[str, List[Tuple[str, str, str, int]]] = {}

    for file in files:
        try:
            rel = file.relative_to(base_dir)
            file_label = str(rel.with_suffix(""))
        except ValueError:
            file_label = file.stem
        try:
            book = safe_read_excel(file, sheet_name=None)
        except Exception:
            read_errors.append(file.name)
            continue

        found_in_this_file = False
        per_sheet_info: List[Tuple[str, str, str, int]] = []
        for sheet_name, df in (book or {}).items():
            if should_exclude_sheet(str(sheet_name)):
                continue
            if not isinstance(df, pd.DataFrame) or df.empty:
                continue

            selected_col, method, cnt = select_best_column(df, target_column, use_smart)
            if selected_col is None:
                continue

            found_in_this_file = True
            per_sheet_info.append((str(sheet_name), str(selected_col), method, cnt))

            series = df[selected_col].dropna()
            if series.empty:
                continue

            series = clean_item_values(series, selected_col, only_equipment=only_equipment)
            counts = series.value_counts()
            for item_value, qty in counts.items():
                rows.append({
                    "Arquivo": file_label,
                    "Item": item_value,
                    "Quantidade": int(qty),
                })

        if not found_in_this_file:
            ignored_missing_col.append(file.name)
        else:
            column_debug[file_label] = per_sheet_info

    if not rows:
        return pd.DataFrame(columns=["Arquivo", "Item", "Quantidade"]), ignored_missing_col, read_errors, column_debug

    df_result = pd.DataFrame(rows)
    df_result = (
        df_result.groupby(["Arquivo", "Item"], as_index=False, observed=True)["Quantidade"].sum()
        .sort_values(["Arquivo", "Quantidade"], ascending=[True, False])
        .reset_index(drop=True)
    )
    return df_result, ignored_missing_col, read_errors, column_debug


def discover_unique_items(files: List[Path], target_column: str, use_smart: bool = True, only_equipment: bool = False) -> List[str]:
    """Descobre a lista única de valores da coluna alvo ao longo de todos os arquivos/abas."""
    unique_values = set()
    for file in files:
        try:
            book = safe_read_excel(file, sheet_name=None)
        except Exception:
            continue
        for sheet_name, df in (book or {}).items():
            if should_exclude_sheet(str(sheet_name)):
                continue
            if not isinstance(df, pd.DataFrame) or df.empty:
                continue
            selected_col, _, _ = select_best_column(df, target_column, use_smart)
            if selected_col is None:
                continue
            series = df[selected_col].dropna()
            if series.empty:
                continue
            series = clean_item_values(series, selected_col, only_equipment=only_equipment)
            for v in series.unique():
                if v != "":
                    unique_values.add(v)
    return sorted(unique_values)


def show_plot(fig, **kwargs):
    """Exibe o gráfico com configurações padrão.
    Compat: converte use_container_width -> width ('stretch' ou 'content').
    """
    if "use_container_width" in kwargs:
        use = kwargs.pop("use_container_width")
        kwargs["width"] = "stretch" if use else "content"
    # Transição leve para gráficos re-renderizados
    try:
        fig.update_layout(transition_duration=250)
    except Exception:
        pass
    # Tema dark uniforme para todos os gráficos
    try:
        fig.update_layout(
            paper_bgcolor="#000000",
            plot_bgcolor="#000000",
            font=dict(color="#e5e7eb"),
            xaxis=dict(showgrid=True, gridcolor="#111111", zerolinecolor="#111111", linecolor="#222222", tickfont=dict(color="#cbd5e1")),
            yaxis=dict(showgrid=True, gridcolor="#111111", zerolinecolor="#111111", linecolor="#222222", tickfont=dict(color="#cbd5e1")),
            legend=dict(bgcolor="rgba(0,0,0,0.6)", font=dict(color="#e5e7eb")),
        )
    except Exception:
        pass
    st.plotly_chart(fig, **kwargs)


def save_to_excel(df_by_file: pd.DataFrame, df_totals: pd.DataFrame, path: Path, group_name: Optional[str] = None) -> None:
    ordered = df_by_file[["Item", "Quantidade", "Arquivo"]].copy()
    totals = df_totals[["Item", "Quantidade"]].copy()
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        ordered.to_excel(writer, index=False, sheet_name="resultado")
        totals.to_excel(writer, index=False, sheet_name="totais_por_item")
        if group_name:
            totals_group = totals.copy()
            totals_group.insert(0, "Grupo", group_name)
            totals_group.to_excel(writer, index=False, sheet_name="resultado_consolidado")
        else:
            totals.to_excel(writer, index=False, sheet_name="resultado_consolidado")


def render_bar_chart(df: pd.DataFrame, item_order: List[str]) -> None:
    fig = px.bar(
        df,
        x="Item",
        y="Quantidade",
        color="Arquivo",
        barmode="group",
        title="Contagem de Itens por Planilha/Mês",
        hover_data={"Quantidade": ":,"},
        category_orders={"Item": item_order},
    )
    fig.update_layout(
        xaxis_title="Item",
        yaxis_title="Quantidade",
        xaxis=dict(categoryorder="array", categoryarray=item_order),
        margin=dict(l=20, r=20, t=60, b=20),
    )
    fig.update_xaxes(tickangle=-45)
    show_plot(fig, use_container_width=True)


def render_bar_chart_consolidated(df_totals: pd.DataFrame, item_order: List[str]) -> None:
    fig = px.bar(
        df_totals,
        x="Item",
        y="Quantidade",
        title="Contagem de Itens por Planilha/Mês (Consolidado)",
        hover_data={"Quantidade": ":,"},
        category_orders={"Item": item_order},
    )
    fig.update_layout(
        xaxis_title="Item",
        yaxis_title="Quantidade",
        xaxis=dict(categoryorder="array", categoryarray=item_order),
        margin=dict(l=20, r=20, t=60, b=20),
    )
    fig.update_xaxes(tickangle=-45)
    show_plot(fig, use_container_width=True)


def main() -> None:
    
    # Tela de carregamento moderna
    if "loading_complete" not in st.session_state:
        st.session_state["loading_complete"] = False
    
    if not st.session_state["loading_complete"]:
        # CSS para tela de carregamento
        st.markdown("""
        <style>
            .stApp {
                background: #000000 !important;
            }
            .main .block-container {
                padding: 0 !important;
                max-width: 100% !important;
            }
            body {
                overflow: hidden !important;
            }
            .loading-screen {
                position: fixed;
                top: 0;
                left: 0;
                width: 100vw;
                height: 100vh;
                background: #000000;
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                z-index: 9999;
                color: white;
                font-family: 'Segoe UI', 'Roboto', sans-serif;
            }
            .title {
                font-size: 4.5rem;
                font-weight: 700;
                background: linear-gradient(45deg, #2563eb, #3b82f6, #60a5fa);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                background-clip: text;
                text-shadow: 0 0 30px rgba(37, 99, 235, 0.5);
                margin-bottom: 1rem;
                letter-spacing: 4px;
                animation: glow 2s ease-in-out infinite alternate;
            }
            .subtitle {
                font-size: 1.3rem;
                color: #e5e7eb;
                margin-bottom: 3rem;
                font-weight: 300;
                letter-spacing: 1px;
                opacity: 0.9;
            }
            .progress-container {
                width: 400px;
                height: 6px;
                background: rgba(26, 26, 26, 0.8);
                border-radius: 3px;
                overflow: hidden;
                margin-bottom: 1.5rem;
                box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.3);
            }
            .progress-bar {
                height: 100%;
                background: linear-gradient(90deg, #2563eb 0%, #3b82f6 50%, #60a5fa 100%);
                border-radius: 3px;
                box-shadow: 0 0 15px rgba(37, 99, 235, 0.6);
                transition: width 0.3s ease;
            }
            .percentage {
                font-size: 1.8rem;
                font-weight: 600;
                color: #2563eb;
                text-shadow: 0 0 10px rgba(37, 99, 235, 0.8);
                margin-bottom: 1rem;
                letter-spacing: 2px;
            }
            .status {
                font-size: 1rem;
                color: #e5e7eb;
                opacity: 0.8;
                font-weight: 300;
            }
            @keyframes glow {
                0% { text-shadow: 0 0 30px rgba(37, 99, 235, 0.5); }
                100% { text-shadow: 0 0 40px rgba(37, 99, 235, 0.8), 0 0 60px rgba(59, 130, 246, 0.4); }
            }
        </style>
        """, unsafe_allow_html=True)
        
        # Container para a tela de carregamento
        loading_container = st.empty()
        
        # Simular carregamento
        for i in range(101):
            if i < 20:
                status = "Inicializando sistema..."
            elif i < 40:
                status = "Carregando dados..."
            elif i < 60:
                status = "Processando informações..."
            elif i < 80:
                status = "Preparando visualizações..."
            else:
                status = "Finalizando carregamento..."
            
            with loading_container.container():
                st.markdown(f"""
                <div class="loading-screen">
                    <div class="title">ALOCAMA</div>
                    <div class="subtitle">Sistema de Contratos</div>
                    <div class="progress-container">
                        <div class="progress-bar" style="width: {i}%;"></div>
                    </div>
                    <div class="percentage">{i}%</div>
                    <div class="status">{status}</div>
                </div>
                """, unsafe_allow_html=True)
            
            time.sleep(0.03)
        
        # Fade out final
        with loading_container.container():
            st.markdown("""
            <div class="loading-screen" style="opacity: 1; transition: opacity 2s ease-out;">
                <div class="title">ALOCAMA</div>
                <div class="subtitle">Sistema de Contratos</div>
                <div class="progress-container">
                    <div class="progress-bar" style="width: 100%;"></div>
                </div>
                <div class="percentage">100%</div>
                <div class="status">Carregamento concluído!</div>
            </div>
            """, unsafe_allow_html=True)
        
        time.sleep(1.5)
        st.session_state["loading_complete"] = True
        st.rerun()
        return
    
    # Dashboard principal
    
    # Se solicitado, rolar automaticamente para uma âncora específica após o rerun
    if st.session_state.get("__scroll_hash"):
        target = st.session_state.get("__scroll_hash")
        st.markdown(f"""
                <script>
                    setTimeout(function(){{
                      try {{
                        const root = window.parent || window;
                        if (root && root.location) {{
                          if (root.location.hash !== '#{target}') {{ root.location.hash = '#{target}'; }}
                        }}
                        const doc = (window.parent && window.parent.document) ? window.parent.document : document;
                        const el = doc.getElementById('{target}');
                        if (el) {{ el.scrollIntoView({{behavior: 'smooth', block: 'start'}}); }}
                      }} catch(e) {{}}
                    }}, 60);
                </script>
            """, unsafe_allow_html=True)
        st.session_state["__scroll_hash"] = None
    
    if not render_login():
        return
    
    # Título do dashboard
    _, col_title, _ = st.columns([0.06, 0.88, 0.06])
    with col_title:
        st.markdown(
        f"""
        <div class='page-hero'>
          <h1 class='page-hero__title'>Dashboard de Contratos <span class='sep'>|</span> <span class='brand'>Alocama</span></h1>
          <div class='page-hero__subtitle'>Indicadores do Setor</div>
          <div class='page-hero__bar'></div>
          {_build_hero_media_html()}
        </div>
        """,
        unsafe_allow_html=True,
    )
    
    # Fade removido - dashboard carrega diretamente
    # fundo será embutido no bloco acima
    # CSS mínimo: ajustar padding superior para não cortar o título e exibir menu nativo
    st.markdown(
        """
        <style>
        a[aria-label^='Anchor link']{display:none!important}
        .block-container{padding-top:2.25rem!important}
        :root{--accent:#2563eb; --bg:#000000; --bg-soft:#000000; --text:#e5e7eb}
        .stApp, [data-testid='stAppViewContainer'], .block-container, .stMarkdown, .stSelectbox, .stRadio, .stButton, .stExpander {background-color:var(--bg)!important}
        [data-testid='stSidebar'], [data-testid='stHeader']{background-color:var(--bg)!important}
        div[data-baseweb='select']>div, .st-bf, .st-al{background-color:#0a0a0a!important;border-color:#1a1a1a!important}
        .stButton>button{background:#111!important;border-color:#1a1a1a!important}
        /* Pulse para botão primário */
        div[data-testid="stButton"] > button[data-testid="baseButton-primary"],
        div[data-testid="stButton"] > button[data-testid="baseButton-primary"]:hover,
        div[data-testid="stButton"] > button[data-testid="baseButton-primary"]:focus,
        .stButton > button[data-testid="baseButton-primary"],
        .stButton > button[data-testid="baseButton-primary"]:hover,
        .stButton > button[data-testid="baseButton-primary"]:focus {
            animation: pulse 2s infinite !important;
            box-shadow: 0 0 0 0 rgba(37, 99, 235, 0.7) !important;
        }
        @keyframes pulse {
            0% {
                transform: scale(1);
                box-shadow: 0 0 0 0 rgba(37, 99, 235, 0.7);
            }
            70% {
                transform: scale(1.05);
                box-shadow: 0 0 0 10px rgba(37, 99, 235, 0);
            }
            100% {
                transform: scale(1);
                box-shadow: 0 0 0 0 rgba(37, 99, 235, 0);
            }
        }
        .page-hero{display:flex!important;flex-direction:column;align-items:center;margin:18px 0 10px 0;text-align:center;visibility:visible!important;opacity:1!important}
        .page-hero__title{margin:0;font-weight:800;font-size:34px;line-height:1.15;letter-spacing:.2px;color:#e5e7eb!important;display:block!important;visibility:visible!important;opacity:1!important}
        .page-hero__title .brand{color:var(--accent)}
        .page-hero__title .sep{color:var(--accent);opacity:.9;margin:0 .25rem}
        .page-hero__subtitle{margin-top:6px;font-size:16px;opacity:.9}
        .page-hero__bar{margin-top:10px;width:clamp(140px,22vw,360px);height:3px;border-radius:999px;background:linear-gradient(90deg,var(--accent),#7c3aed)}
        .page-hero{position:relative;overflow:hidden;border-radius:14px;background:#000}
        .page-hero__bg{position:absolute;inset:0;z-index:-1}
        .page-hero__bg video,.page-hero__bg img{width:100%;height:100%;object-fit:cover;filter:brightness(.22) saturate(1.05)}
        </style>
        """,
        unsafe_allow_html=True,
    )

    # Removido input de pasta base/checkbox: usa diretório atual e subpastas por padrão
    base_dir = Path.cwd()
    recursive = True


    def discover_groups(dir_base: Path) -> List[str]:
        grupos: List[str] = []
        for entry in dir_base.iterdir():
            if not entry.is_dir():
                continue
            has_xlsx = False
            try:
                for p in entry.rglob("*"):
                    if is_excel_file(p):
                        has_xlsx = True
                        break
            except Exception:
                has_xlsx = False
            if has_xlsx:
                grupos.append(entry.name)
        return sorted(grupos)

    grupos_disponiveis = discover_groups(base_dir)
    if grupos_disponiveis:
        grupo_escolhido = render_company_selector(grupos_disponiveis)
        grupos_selecionados = [grupo_escolhido] if grupo_escolhido else []
    else:
        grupos_selecionados = []

    if grupos_selecionados:
        excel_files: List[Path] = []
        per_group_counts: Dict[str, int] = {}
        for g in grupos_selecionados:
            group_files = list_excel_files(base_dir / g, recursive=True)
            excel_files.extend(group_files)
            per_group_counts[g] = len(group_files)
    else:
        excel_files = list_excel_files(base_dir, recursive=recursive)

    if not excel_files:
        st.warning("Nenhum arquivo .xlsx encontrado na pasta informada. Adicione arquivos e recarregue a página.")
        return

    with st.expander("Arquivos detectados", expanded=False):
        try:
            st.write([str(f.relative_to(base_dir)) for f in excel_files])
        except Exception:
            st.write([f.name for f in excel_files])
        if grupos_selecionados:
            st.write({k: per_group_counts.get(k, 0) for k in grupos_selecionados})

    discovered_cols = discover_columns(excel_files)
    produto_col = next((c for c in discovered_cols if normalize_column_name(c) == "produto"), None)
    default_col = (
        produto_col
        if produto_col is not None
        else ("Item" if any(normalize_column_name(c) == "item" for c in discovered_cols) else (discovered_cols[0] if discovered_cols else "Item"))
    )

    st.subheader("Análise Rápida")
    
    # Definir sel_files fora do bloco if run para uso em outras partes da função
    sel_files = [f for f in excel_files if month_from_path(f) in {"1","2","3","4", "5", "6", "7", "8"}]
    
    run = st.button("Executar Análise", type="primary", key="executar_analise")
    
    # CSS com efeito de borda azul pulsante APENAS no hover
    st.markdown("""
    <style>
    /* Botões normais - sem efeito */
    button, 
    .stButton button,
    div[data-testid="stButton"] button,
    button[data-testid="baseButton-primary"] {
        border: 2px solid transparent !important;
        border-radius: 6px !important;
        outline: none !important;
        transition: all 0.3s ease !important;
    }
    
    /* Efeito APENAS no hover - pulse + borda azul */
    button:hover,
    .stButton button:hover,
    div[data-testid="stButton"] button:hover,
    button[data-testid="baseButton-primary"]:hover {
        animation: pulse 1.5s infinite !important;
        border: 3px solid #2563eb !important;
        border-radius: 8px !important;
        box-shadow: 0 0 0 0 rgba(37, 99, 235, 0.7) !important;
    }
    
    @keyframes pulse {
        0% {
            transform: scale(1);
            box-shadow: 0 0 0 0 rgba(37, 99, 235, 0.7);
            border-color: #2563eb !important;
        }
        50% {
            transform: scale(1.05);
            box-shadow: 0 0 0 6px rgba(37, 99, 235, 0.3);
            border-color: #60a5fa !important;
        }
        100% {
            transform: scale(1);
            box-shadow: 0 0 0 0 rgba(37, 99, 235, 0);
            border-color: #2563eb !important;
        }
    }
    </style>
    """, unsafe_allow_html=True)

    if run:
        with st.spinner("Processando (coluna E) e contando itens por pasta 1/2/3/4/5/6/7/8..."):
            df_result, ignored_files, error_files, column_debug = count_items_in_files(
                sel_files, "E", base_dir, use_smart=True, only_equipment=True
            )

        if df_result.empty:
            st.error("Não foi possível encontrar a coluna informada em nenhum arquivo.")
            if ignored_files:
                with st.expander("Arquivos sem a coluna informada", expanded=False):
                    st.write(ignored_files)
            if error_files:
                with st.expander("Arquivos com erro de leitura", expanded=False):
                    st.write(error_files)
            return

        def extract_pasta(label: str) -> str:
            parts = re.split(r"[\\/]+", label)
            for p in parts:
                m1 = re.fullmatch(r"(\d{4})-(0?[1-9]|1[0-2])", p)
                if m1:
                    return m1.group(2).lstrip("0")
                m2 = re.search(r"(0?[1-9]|1[0-2])", p)
                if m2:
                    return m2.group(1).lstrip("0")
                # Nomes de mês PT-BR
                pn = normalize_text_for_match(p)
                word_to_num = {
                    "janeiro": "1",
                    "fevereiro": "2",
                    "marco": "3",
                    "março": "3",
                    "abril": "4",
                    "maio": "5",
                    "junho": "6",
                    "julho": "7",
                    "agosto": "8",
                }
                if pn in word_to_num:
                    return word_to_num[pn]
            return "?"

        df_result["Pasta"] = df_result["Arquivo"].apply(extract_pasta)
        # Mantemos somente Agosto/2025 ou os 3 meses anteriores para comparação (Maio–Agosto 2025)
        def _is_2025_mm(label: str, months: set) -> bool:
            parts = re.split(r"[\\/]+", label)
            for p in parts:
                if re.fullmatch(r"2025-(0?[1-8])", p.strip()):
                    m = re.fullmatch(r"2025-(0?[1-8])", p.strip()).group(1).lstrip("0")
                    return m in months
            return False
        months_keep = {"1","2","3","4","5","6","7","8"}
        df_result = df_result[df_result["Arquivo"].apply(lambda s: _is_2025_mm(s, months_keep))]
        df_result = df_result[~(
            df_result["Item"].astype(str).str.strip().str.upper() == "DOMMUS"
        ) | (df_result["Quantidade"] > 0)].reset_index(drop=True)

        # Removida a correção de duplicidade (pedido para tirar a divisão)

        # Unificar variações específicas
        df_result["Item"] = df_result["Item"].map(canonicalize_electric_bed_two_movements)
        df_result["Item"] = df_result["Item"].map(canonicalize_wheelchair_group)
        df_result["Item"] = df_result["Item"].map(canonicalize_wheelchair_obese_60)
        df_result["Item"] = df_result["Item"].map(canonicalize_wheelchair_50)
        df_result["Item"] = df_result["Item"].map(canonicalize_walker)
        df_result["Item"] = df_result["Item"].map(canonicalize_bed_alt_trem)

        df_totais = (
            df_result.groupby("Item", as_index=False, observed=True)["Quantidade"].sum().sort_values("Quantidade", ascending=False)
        )
        item_order = df_totais["Item"].tolist()
        df_result["Item"] = pd.Categorical(df_result["Item"], categories=item_order, ordered=True)
        df_result_sorted = df_result.sort_values(["Item", "Pasta", "Quantidade"], ascending=[True, True, False])

        output_path = base_dir / OUTPUT_FILENAME
        try:
            save_df = df_result_sorted[["Item", "Quantidade", "Pasta"]]

            df_emp = df_result_sorted.copy()
            df_emp["Empresa"] = df_emp["Arquivo"].apply(primary_group_from_label).str.upper()
            df_emp["Empresa"] = df_emp["Empresa"].replace({
                "GRUPO SOLAR": "SOLAR",
            })

            months_numeric = pd.to_numeric(df_emp["Pasta"], errors="coerce")
            last_month_num = int(months_numeric.max()) if not months_numeric.dropna().empty else None
            last_month_str = str(last_month_num) if last_month_num is not None else None
            df_emp_last = df_emp[df_emp["Pasta"] == last_month_str] if last_month_str else df_emp.copy()
            df_consolidado_geral = (
                df_emp_last.groupby("Item", as_index=False, observed=True)["Quantidade"].sum().sort_values("Quantidade", ascending=False)
            )

            df_emp["ItemCanon"] = df_emp["Item"].map(canonicalize_trio_item)

            df_mes_empresa = (
                df_emp.groupby(["Empresa", "ItemCanon", "Pasta"], as_index=False, observed=True)["Quantidade"].sum()
            )
            df_peak = (
                df_mes_empresa.sort_values(["Empresa", "ItemCanon", "Quantidade"], ascending=[True, True, False])
                .drop_duplicates(["Empresa", "ItemCanon"], keep="first")
            )
            df_peak["Posição"] = (
                df_peak.groupby("Empresa")["Quantidade"].rank(ascending=False, method="first").astype(int)
            )
            df_top3_empresa = (
                df_peak[df_peak["Posição"] <= 3]
                .sort_values(["Empresa", "Posição"], ascending=[True, True])
            )

            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                save_df.to_excel(writer, index=False, sheet_name="resultado")
                df_consolidado_geral.to_excel(writer, index=False, sheet_name="consolidado_geral")
                df_top3_empresa.to_excel(writer, index=False, sheet_name="top3_por_empresa")
        except Exception as e:
            st.error(f"Falha ao salvar o arquivo de saída: {e}")
            return

        st.success("✅ Extração e contagem concluídas! Resultado salvo em resultado_itens.xlsx")

        st.markdown('<h3 style="margin:0 0 8px 0;">Dashboard</h3>', unsafe_allow_html=True)
        # Painel de KPIs e mosaico inicial removidos conforme solicitação
        st.info("Clique em 'Executar Análise' para gerar relatórios e gráficos detalhados.")

        # ====== Layout compacto em mosaico (três colunas) ======
        if False and empresa_atual_hdr:
            df_e_all = df_emp_viz_hdr[df_emp_viz_hdr["Empresa"] == empresa_atual_hdr].copy()
            meses_ordem = {"Janeiro":1, "Fevereiro":2, "Março":3, "Abril":4, "Maio":5, "Junho":6, "Julho":7, "Agosto":8}
            ultimo_idx = df_e_all["Mês"].map(meses_ordem).max()
            mes_ultimo = [k for k,v in meses_ordem.items() if v == ultimo_idx]
            mes_ultimo = mes_ultimo[0] if mes_ultimo else df_e_all["Mês"].iloc[-1]
            df_e_last = df_e_all[df_e_all["Mês"] == mes_ultimo].copy()

            # Linha 1 de gráficos
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
            g1, g2, g3 = st.columns([1,1,1])
            with g1:
                top_last = (
                    df_e_last.groupby("Item", as_index=False, observed=True)["Quantidade"].sum()
                    .sort_values("Quantidade", ascending=False).head(6)
                )
                fig_t1 = px.bar(top_last, x="Item", y="Quantidade", title=f"Top Itens – {mes_ultimo}")
                fig_t1.update_layout(width=400, height=220, margin=dict(l=10, r=10, t=38, b=60))
                fig_t1.update_xaxes(tickangle=-45)
                show_plot(fig_t1, use_container_width=True)
            with g2:
                df_sum_last = (
                    df_e_last.groupby("Item", as_index=False, observed=True)["Quantidade"].sum()
                    .sort_values("Quantidade", ascending=False)
                )
                top3 = df_sum_last.head(3)
                outros = df_sum_last["Quantidade"].sum() - top3["Quantidade"].sum()
                pie_df = top3.rename(columns={"Item": "Item", "Quantidade": "Quantidade"})[["Item","Quantidade"]].copy()
                if outros > 0:
                    pie_df = pd.concat([pie_df, pd.DataFrame([{ "Item": "Outros", "Quantidade": outros }])])
                fig_p = px.pie(pie_df, names="Item", values="Quantidade", hole=0.55, title=f"Participação Top 3 – {mes_ultimo}")
                fig_p.update_layout(width=400, height=220, margin=dict(l=10, r=10, t=38, b=10))
                show_plot(fig_p, use_container_width=True)
            with g3:
                df_cat = df_e_last.copy()
                df_cat["Categoria"] = df_cat["Item"].map(categorize_item_name)
                df_cat_sum = (
                    df_cat.groupby("Categoria", as_index=False, observed=True)["Quantidade"].sum().sort_values("Quantidade", ascending=False)
                )
                fig_cat = px.bar(df_cat_sum, x="Categoria", y="Quantidade", title=f"Resumo por categoria – {mes_ultimo}")
                fig_cat.update_layout(width=400, height=220, margin=dict(l=10, r=10, t=38, b=60))
                fig_cat.update_xaxes(tickangle=-30)
                show_plot(fig_cat, use_container_width=True)

            # Linha 2 de gráficos
            st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
            g4, g5, g6 = st.columns([1,1,1])
            with g4:
                camas_principais = ["CAMA ELÉTRICA 3 MOVIMENTOS", "CAMA MANUAL 2 MANIVELAS"]
                df_camas = (
                    df_e_all[df_e_all["Item"].isin(camas_principais)]
                    .groupby(["Mês","Item"], as_index=False)["Quantidade"].sum()
                )
                # preencher meses ausentes
                grid = pd.MultiIndex.from_product([list(meses_ordem.keys()), camas_principais], names=["Mês","Item"]).to_frame(index=False)
                df_camas = grid.merge(df_camas, on=["Mês","Item"], how="left").fillna({"Quantidade":0})
                fig_line = px.line(df_camas, x="Mês", y="Quantidade", color="Item", markers=True, title="Camas por mês")
                fig_line.update_layout(width=400, height=220, margin=dict(l=10, r=10, t=38, b=10))
                show_plot(fig_line, use_container_width=True)
            with g5:
                # Faturamento estimado por mês
                pmap = _price_map_for_company(empresa_atual_hdr)
                if pmap:
                    tmp = df_e_all.copy()
                    tmp["key"] = tmp["Item"].map(normalize_text_for_match)
                    tmp["PrecoDiaria"] = tmp["key"].map(pmap)
                    tmp = tmp.dropna(subset=["PrecoDiaria"])
                    tmp["Dias"] = tmp["Mês"].map({"Fevereiro":28, "Março":31, "Abril":30, "Maio":31, "Junho":30, "Julho":31, "Agosto":31}).fillna(30)
                    tmp["Faturamento"] = tmp["Quantidade"] * tmp["PrecoDiaria"] * tmp["Dias"]
                    df_rev_mes = tmp.groupby("Mês", as_index=False, observed=True)["Faturamento"].sum()
                    df_rev_mes["Mês"] = pd.Categorical(df_rev_mes["Mês"], categories=list(meses_ordem.keys()), ordered=True)
                    df_rev_mes = df_rev_mes.sort_values("Mês")
                    fig_fat = px.bar(df_rev_mes, x="Mês", y="Faturamento", title="Faturamento estimado por mês")
                    fig_fat.update_yaxes(tickprefix="R$ ", tickformat=",.2f")
                    fig_fat.update_layout(width=400, height=220, margin=dict(l=10, r=10, t=38, b=10))
                    show_plot(fig_fat, use_container_width=True)
                else:
                    st.info("Sem mapa de preços para estimar faturamento deste grupo.")
            with g6:
                # ARPU (estimado) = Faturamento estimado / Vidas únicas
                try:
                    month_labels = list(meses_ordem.keys())
                    month_map_ym = {"Janeiro":"2025-01", "Fevereiro":"2025-02", "Março":"2025-03","Abril":"2025-04","Maio":"2025-05","Junho":"2025-06","Julho":"2025-07","Agosto":"2025-08"}
                    sets = {m:set() for m in month_labels}
                    for f in sel_files:
                        try:
                            if primary_group_from_label(str(f)).upper() != empresa_atual_hdr:
                                continue
                            ym = year_month_from_path(f)
                            if ym not in month_map_ym.values():
                                continue
                            mes_lab = [k for k,v in month_map_ym.items() if v == ym][0]
                            # Ler com cabeçalho para permitir identificação da coluna de nomes
                            book = safe_read_excel(f, sheet_name=None)
                        except Exception:
                            continue
                        for sh, df in (book or {}).items():
                            if should_exclude_sheet(str(sh)) or not isinstance(df, pd.DataFrame) or df.empty:
                                continue
                            # Escolher melhor coluna de nomes de pacientes/vidas
                            series = None
                            name_col = None
                            try:
                                name_col = select_best_name_column(df)
                            except Exception:
                                name_col = None
                            if name_col:
                                try:
                                    series = df[name_col]
                                except Exception:
                                    series = None
                            # Fallback: coluna B apenas se os valores se parecerem com nomes
                            if series is None:
                                cand = None
                                try:
                                    if df.shape[1] >= 2:
                                        cand = df.iloc[:, 1]
                                except Exception:
                                    cand = None
                                if cand is not None:
                                    scheck = cand.dropna().astype(str).str.strip()
                                    scheck = scheck[scheck != ""]
                                    if not scheck.empty:
                                        norm = scheck.map(normalize_text_for_match)
                                        looks_like = norm.str.contains(r"[a-z]", regex=True, na=False) & norm.str.contains(r"\\s", regex=True, na=False) & (norm.str.len() >= 5)
                                        if int(looks_like.sum()) >= 5:
                                            series = cand
                            if series is None:
                                continue
                            s = series.dropna().astype(str).str.strip()
                            s = s[s != ""]
                            sets[mes_lab].update(s.map(normalize_text_for_match).tolist())
                    vidas_list = [len(sets[m]) for m in month_labels]
                    # Garante mínimo de 1 para evitar divisão por zero e ARPU zerado por gráficos vazios
                    vidas_list = [v if v > 0 else None for v in vidas_list]
                    vidas_df = pd.DataFrame({"Mês": month_labels, "Vidas": vidas_list})
                    # Faturamento geral (manual) por mês para ARPU
                    total_rev_map = {"Janeiro": 98579.58, "Fevereiro": 87831.11, "Março": 96184.47, "Abril": 92286.01, "Maio": 87803.67, "Junho": 77499.87, "Julho": 81856.05, "Agosto": 82609.95}
                    rev_df = pd.DataFrame({"Mês": month_labels, "Faturamento": [total_rev_map.get(m, 0.0) for m in month_labels]})
                    arpu_df = rev_df.merge(vidas_df, on="Mês", how="left")
                    arpu_df["ARPU"] = arpu_df.apply(lambda r: (r["Faturamento"] / r["Vidas"]) if pd.notna(r["Vidas"]) and r["Vidas"]>0 else None, axis=1)
                    fig_arpu = px.bar(arpu_df, x="Mês", y="ARPU", title="ARPU (Faturamento geral / Vidas)")
                    fig_arpu.update_yaxes(tickprefix="R$ ", tickformat=",.2f")
                    fig_arpu.update_layout(width=400, height=220, margin=dict(l=10, r=10, t=38, b=10))
                    show_plot(fig_arpu, use_container_width=True)
                except Exception:
                    st.info("ARPU não pôde ser calculado.")
        month_map = {"1":"Janeiro","2":"Fevereiro","3":"Março","4":"Abril","5":"Maio", "6": "Junho", "7": "Julho", "8": "Agosto"}
        df_viz = df_result_sorted.copy()
        df_viz["Mês"] = df_viz["Pasta"].map(month_map).fillna(df_viz["Pasta"])
        month_order = [month_map[m] for m in ["1","2","3","4","5", "6", "7", "8"]]

        # 1) Gráfico – Top 10 por Item (comparativo por mês Maio→Agosto), somando COMPL ao item base
        top10_items_orig = df_totais.head(10)["Item"].tolist()
        df_viz_top = df_viz[df_viz["Item"].isin(top10_items_orig)].copy()
        def _strip_compl_prefix(text: str) -> str:
            t = str(text)
            return re.sub(r"^\s*\(?\s*compl\.?\s*\)?\s*", "", t, flags=re.IGNORECASE)
        df_viz_top = df_viz_top.assign(ItemAgrupado=df_viz_top["Item"].apply(_strip_compl_prefix))
        df_viz_top = (
            df_viz_top.groupby(["ItemAgrupado", "Mês"], as_index=False, observed=True)["Quantidade"].sum()
            .rename(columns={"ItemAgrupado":"Item"})
        )
        top10_after_agg = (
            df_viz_top.groupby("Item", as_index=False, observed=True)["Quantidade"].sum()
            .sort_values("Quantidade", ascending=False)
            .head(10)["Item"].tolist()
        )
        # Garante presença de todos os meses (Março..Agosto) para cada um dos top10 itens
        if top10_after_agg:
            df_all_pairs = pd.MultiIndex.from_product([top10_after_agg, month_order], names=["Item","Mês"]).to_frame(index=False)
            df_viz_top = (
                df_all_pairs.merge(df_viz_top, on=["Item","Mês"], how="left")
                .fillna({"Quantidade": 0})
            )
        st.markdown('<div class="fade-in-on-scroll" style="margin-top:0;">', unsafe_allow_html=True)
        fig = px.bar(
            df_viz_top[df_viz_top["Item"].isin(top10_after_agg)],
            x="Item",
            y="Quantidade",
            color="Mês",
            barmode="group",
            category_orders={"Mês": month_order, "Item": top10_after_agg},
            title="Top 10 - Comparação de Itens (Janeiro/Fevereiro/Março/Abril/Maio/Junho/Julho/Agosto)",
            hover_data={"Mês": True, "Quantidade": ":,", "Item": True},
        )
        fig.update_traces(
            marker_line_color="#FFFFFF",
            marker_line_width=0.5,
            hovertemplate="Item: %{x}<br>Mês: %{customdata[0]}<br>Qtd: %{y:,}<extra></extra>",
        )
        fig.update_layout(
            xaxis_title="Itens (Janeiro / Fevereiro / Março / Abril / Maio / Junho / Julho / Agosto)",
            yaxis_title="Quantidade",
            showlegend=False,
            width=1200,  # Aumentar largura do gráfico
            height=500,  # Altura padrão
            margin=dict(l=10, r=10, t=60, b=120),
            font=dict(size=12),
        )
        fig.update_xaxes(tickangle=-60)
        show_plot(fig, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)


        month_map_hdr = {"3":"Março","4": "Abril","5": "Maio", "6": "Junho", "7": "Julho", "8": "Agosto"}
        last_month_hdr = df_emp_last["Pasta"].iloc[0] if not df_emp_last.empty else None
        last_month_label = month_map_hdr.get(str(last_month_hdr), str(last_month_hdr) if last_month_hdr else "-")
        with st.expander(f"Consolidado geral do último mês ({last_month_label})", expanded=False):
            def _strip_compl_prefix(text: str) -> str:
                t = str(text)
                return re.sub(r"^\s*\(?\s*compl\.?\s*\)?\s*", "", t, flags=re.IGNORECASE)

            # Recalcula consolidado limpando prefixos (COMPL.) e semelhantes
            tmp = df_emp_last.assign(ItemLimpo=df_emp_last["Item"].apply(_strip_compl_prefix))
            tmp["ItemCanon2"] = tmp["ItemLimpo"].map(canonicalize_electric_bed_two_movements)
            df_consol_limpo = (
                tmp.groupby("ItemCanon2", as_index=False)["Quantidade"].sum()
                .sort_values("Quantidade", ascending=False)
                .rename(columns={"ItemCanon2": "Item"})
            )
            df_consol_limpo.index = range(1, len(df_consol_limpo) + 1)
            df_consol_limpo.index.name = "Posição"
            st.dataframe(df_consol_limpo, use_container_width=True)
        df_peak_item = (
            df_result.sort_values(["Item", "Quantidade"], ascending=[True, False])
            .drop_duplicates(["Item"], keep="first")
            .assign(Mês=lambda d: d["Pasta"].map({"1":"Janeiro","2":"Fevereiro","3":"Março","4":"Abril","5":"Maio","6":"Junho","7":"Julho","8":"Agosto"}))
        )[["Item", "Quantidade", "Mês"]]
        # Corrige 'None' exibindo '-' para itens cujo pico não tenha mês detectado
        df_peak_item["Mês"] = df_peak_item["Mês"].fillna("-")
        with st.expander("Mês de pico por item (informativo)", expanded=False):
            df_peak_item_display = df_peak_item.copy()
            df_peak_item_display.index = range(1, len(df_peak_item_display) + 1)
            df_peak_item_display.index.name = "Posição"
            st.dataframe(df_peak_item_display, use_container_width=True)

        # Detalhamento por categoria (com limpeza de prefixo (COMPL.))
        def _strip_compl_prefix(text: str) -> str:
            t = str(text)
            return re.sub(r"^\s*\(?\s*compl\.?\s*\)?\s*", "", t, flags=re.IGNORECASE)

        # Mostrar detalhamento de categorias considerando apenas o último mês disponível
        if 'last_month_str' in locals() and last_month_str:
            df_detail = df_result_sorted[df_result_sorted["Pasta"] == last_month_str].copy()
        else:
            df_detail = df_result_sorted.copy()
        df_detail["ItemLimpo"] = df_detail["Item"].apply(_strip_compl_prefix)
        df_detail["Categoria"] = df_detail["ItemLimpo"].map(categorize_item_name)
        categorias_alvo = ["CAMA", "CADEIRA HIGIÊNICA", "CADEIRA DE RODAS"]
        for cat in categorias_alvo:
            with st.expander(cat, expanded=False):
                sub = df_detail[df_detail["Categoria"] == cat]
                if sub.empty:
                    st.info(f"Sem itens em {cat}")
                else:
                    # Para CADEIRA HIGIÊNICA, agrupar por tamanho (número) quando existir
                    if cat == "CADEIRA HIGIÊNICA":
                        sub = sub.copy()
                        # Extrai primeiro número (tamanho) quando existir no texto
                        sub["Tamanho"] = sub["ItemLimpo"].str.extract(r"(\d+)")
                        # Normalização para testes de palavras-chave (sem acentos)
                        sub["_norm"] = sub["ItemLimpo"].apply(normalize_text_for_match)
                        def _group_row(r):
                            size = r["Tamanho"]
                            norm = r["_norm"] or ""
                            has_estofada = "estofad" in norm
                            has_dobravel = "dobrav" in norm or "dobravel" in norm
                            # Regra adicional: "mod antigo" também é estofada
                            if "mod antigo" in norm and pd.isna(size):
                                size = "44"
                            # Nova regra: tamanho 40 deve somar em 44
                            if pd.notna(size) and str(size) == "40":
                                size = "44"
                            if pd.notna(size) and has_estofada:
                                return f"CADEIRA HIGIÊNICA ESTOFADA {size}"
                            if has_dobravel:
                                # Regra solicitada: somar DOBRÁVEL dentro da 44
                                return "CADEIRA HIGIÊNICA 44"
                            if pd.notna(size):
                                return f"CADEIRA HIGIÊNICA {size}"
                            return r["ItemLimpo"]
                        sub["Grupo"] = sub.apply(_group_row, axis=1)
                        tabela = (
                            sub.groupby("Grupo", as_index=False)["Quantidade"].sum()
                            .sort_values("Quantidade", ascending=False)
                        )
                        tabela = tabela.rename(columns={"Grupo": "Item"})
                    else:
                        # Para CAMA: consolidar em "CAMA 2 MANIVELAS" e "CAMA 3 MANIVELAS" quando aplicável
                        if cat == "CAMA":
                            sub = sub.copy()
                            sub["_norm"] = sub["ItemLimpo"].apply(normalize_text_for_match)
                            def _map_cama(lbl, nrm):
                                if ("2" in nrm and "manivel" in nrm) or ("duas" in nrm and "manivel" in nrm):
                                    return "CAMA 2 MANIVELAS"
                                if ("3" in nrm and "manivel" in nrm) or ("tres" in nrm and "manivel" in nrm):
                                    return "CAMA 3 MANIVELAS"
                                return lbl
                            sub["Grupo"] = [ _map_cama(lbl, nrm) for lbl, nrm in zip(sub["ItemLimpo"], sub["_norm"]) ]
                            tabela = (
                                sub.groupby("Grupo", as_index=False)["Quantidade"].sum()
                                .sort_values("Quantidade", ascending=False)
                            )
                            tabela = tabela.rename(columns={"Grupo": "Item"})
                        # Para CADEIRA DE RODAS: consolidar por tamanho quando houver número
                        elif cat == "CADEIRA DE RODAS":
                            sub = sub.copy()
                            # Extrai tamanho evitando capturar códigos como "C1"; pega números com 2+ dígitos (ex.: 40, 44, 46, 48, 50, 40,5)
                            # Usa o número do final do texto como tamanho (ex.: 40, 44, 46, 48, 50, 40,5)
                            sub["Tamanho"] = sub["ItemLimpo"].str.extract(r"(\d{2,}(?:[\.,]\d{1,2})?)\s*$")
                            def _map_rodas(lbl, size):
                                return f"CADEIRA DE RODAS {size}" if pd.notna(size) else lbl
                            sub["Grupo"] = [ _map_rodas(lbl, size) for lbl, size in zip(sub["ItemLimpo"], sub["Tamanho"]) ]
                            tabela = (
                                sub.groupby("Grupo", as_index=False)["Quantidade"].sum()
                                .sort_values("Quantidade", ascending=False)
                            )
                            tabela = tabela.rename(columns={"Grupo": "Item"})
                        else:
                            tabela = (
                                sub.groupby("ItemLimpo", as_index=False)["Quantidade"].sum()
                                .sort_values("Quantidade", ascending=False)
                            )
                            tabela = tabela.rename(columns={"ItemLimpo": "Item"})

                    tabela.index = range(1, len(tabela) + 1)
                    tabela.index.name = "Posição"
                    st.dataframe(tabela, use_container_width=True)

        st.subheader("Top 3 itens por empresa (Janeiro/Fevereiro/Março/Abril/Maio/Junho/Julho/Agosto)")
        empresas_presentes = sorted(df_top3_empresa["Empresa"].unique().tolist())
        if not empresas_presentes:
            st.info("Sem dados para os grupos selecionados")
        else:
            prefer_order = ["AXX CARE", "HOSPITALAR", "SOLAR", "DOMMUS"]
            empresas_to_show = [e for e in prefer_order if e in empresas_presentes]
            for e in empresas_presentes:
                if e not in empresas_to_show:
                    empresas_to_show.append(e)
            cols = st.columns(len(empresas_to_show))
            for empresa, col in zip(empresas_to_show, cols):
                subset = df_top3_empresa[df_top3_empresa["Empresa"].str.upper() == empresa]
                if subset.empty:
                    continue
                col.markdown(f"**{empresa}**")
                show = subset.assign(
                    Mês=subset["Pasta"].map({"1":"Janeiro","2":"Fevereiro","3":"Março","4":"Abril","5":"Maio","6":"Junho","7":"Julho","8":"Agosto"}),
                    Posição=(subset.groupby("Empresa")["Quantidade"].rank(ascending=False, method="first").astype(int))
                )[["Posição","ItemCanon","Quantidade","Mês"]]
                show = show.rename(columns={"ItemCanon": "Item"})
                # Corrige 'None' exibindo '-' para itens cujo mês não tenha sido detectado
                show["Mês"] = show["Mês"].fillna("-")
                col.dataframe(show.set_index("Posição"), use_container_width=True)


        df_emp_viz = df_result_sorted.copy()
        df_emp_viz["Empresa"] = df_emp_viz["Arquivo"].apply(primary_group_from_label).str.upper()
        df_emp_viz["Empresa"] = df_emp_viz["Empresa"].replace({"GRUPO SOLAR": "SOLAR"})
        df_emp_viz["Mês"] = df_emp_viz["Pasta"].map({"1":"Janeiro","2":"Fevereiro","3":"Março","4":"Abril","5":"Maio","6": "Junho", "7": "Julho", "8": "Agosto"})
        month_order = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho", "Julho", "Agosto"]

        empresas_presentes_viz = sorted(df_emp_viz["Empresa"].unique().tolist())
        if empresas_presentes_viz == ["AXX CARE"]:
            st.subheader("Resumo por categorias – AXX CARE")
            df_e = df_emp_viz[df_emp_viz["Empresa"] == "AXX CARE"]
            if df_e.empty:
                st.info("Sem dados para AXX CARE nos meses 6/7/8")
            else:
                # Categorias desejadas
                df_e_cat = df_e.copy()
                df_e_cat["Categoria"] = df_e_cat["Item"].map(categorize_item_name)
                alvo = ["CAMA", "CADEIRA HIGIÊNICA", "CADEIRA DE RODAS", "SUPORTE DE SORO"]
                # Considerar SOMENTE o último mês disponível
                meses_ordem_full = {"Janeiro":1, "Fevereiro":2, "Março":3, "Abril":4, "Maio":5, "Junho": 6, "Julho": 7, "Agosto": 8}
                ultimo_mes = df_e["Mês"].map(meses_ordem_full).max()
                mes_ult_label = [k for k,v in meses_ordem_full.items() if v == ultimo_mes]
                mes_ult_label = mes_ult_label[0] if mes_ult_label else "Agosto"
                df_e_last = df_e[df_e["Mês"] == mes_ult_label].copy()
                df_e_last["Categoria"] = df_e_last["Item"].map(categorize_item_name)
                df_cat_sum = (
                    df_e_last[df_e_last["Categoria"].isin(alvo)]
                    .groupby(["Categoria"], as_index=False)["Quantidade"].sum()
                    .sort_values("Quantidade", ascending=False)
                )
                fig_cat = px.bar(
                    df_cat_sum,
                    x="Categoria",
                    y="Quantidade",
                    title=f"Quantidade por categoria (Camas, Cadeira Higiene, Cadeira de Rodas, Suporte de Soro) – {mes_ult_label}",
                )
                fig_cat.update_layout(width=800, height=400)
                fig_cat.update_layout(margin=dict(l=20, r=20, t=60, b=80))
                show_plot(fig_cat, use_container_width=True)

                # Determina último mês disponível para os gráficos subsequentes
                meses_ordem = {"Janeiro":1, "Fevereiro":2, "Março":3, "Abril":4, "Maio":5, "Junho": 6, "Julho": 7, "Agosto": 8}
                ultimo_mes = df_e["Mês"].map(meses_ordem).max()
                mes_label = [k for k, v in meses_ordem.items() if v == ultimo_mes]
                mes_label = mes_label[0] if mes_label else "Junho"

                st.subheader(f"Participação dos Top 3 (quantidade) – AXX CARE ({mes_label})")
                df_p_ult = df_e[df_e["Mês"] == mes_label].copy()
                df_p_ult["ItemCanon"] = df_p_ult["Item"].map(canonicalize_trio_item)
                df_sum = df_p_ult.groupby("ItemCanon", as_index=False)["Quantidade"].sum().sort_values("Quantidade", ascending=False)
                top3 = df_sum.head(3)
                outros = df_sum["Quantidade"].sum() - top3["Quantidade"].sum()
                df_pie = top3.rename(columns={"ItemCanon": "Item"})[["Item", "Quantidade"]].copy()
                if outros > 0:
                    df_pie = pd.concat([df_pie, pd.DataFrame([{ "Item": "Outros", "Quantidade": outros }])], ignore_index=True)
                pie_order = df_pie["Item"].tolist()
                color_map = {
                    "CAMA MANUAL 2 MANIVELAS": "#4e79a7",
                    "SUPORTE DE SORO": "#f28e2c",
                    "CAMA ELÉTRICA 3 MOVIMENTOS": "#e15759",
                    "Outros": "#9aa3a8",
                }
                fig_pie = px.pie(
                    df_pie,
                    names="Item",
                    values="Quantidade",
                    hole=0.5,
                    color="Item",
                    color_discrete_map=color_map,
                    category_orders={"Item": pie_order},
                    title=f"Top 3 itens (quantidade) + Outros – AXX CARE ({mes_label})",
                )
                fig_pie.update_layout(width=600, height=500)
                fig_pie.update_traces(
                    sort=False,
                    textposition="inside",
                    texttemplate="%{label}<br>%{percent}",
                    hovertemplate="Item: %{label}<br>Qtd: %{value:,} (%{percent})<extra></extra>",
                )
                fig_pie.update_layout(
                    legend_title_text="Itens",
                    legend_orientation="v",
                    legend_y=0.5,
                    legend_x=1.02,
                    margin=dict(l=20, r=60, t=60, b=20),
                )
                st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                show_plot(fig_pie, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

                # Novo: Impacto dos Top 3 no orçamento (faturamento estimado) – pizza
                pmap_axx = {
                    normalize_text_for_match("CAMA ELÉTRICA 3 MOVIMENTOS"): 10.80,
                    normalize_text_for_match("CAMA MANUAL 2 MANIVELAS"): 2.83,
                    normalize_text_for_match("SUPORTE DE SORO"): 0.67,
                }
                df_fat_last = df_p_ult.copy()
                df_fat_last["key"] = df_fat_last["Item"].apply(normalize_text_for_match)
                df_fat_last["PrecoDiaria"] = df_fat_last["key"].map(pmap_axx)
                df_fat_last = df_fat_last.dropna(subset=["PrecoDiaria"])
                if not df_fat_last.empty:
                    dias_map = {"Fevereiro":28, "Março":31, "Abril":30, "Maio":31, "Junho":30, "Julho":31, "Agosto":31}
                    df_fat_last["Dias"] = df_fat_last["Mês"].map(dias_map).fillna(30)
                    df_fat_last["Faturamento"] = df_fat_last["Quantidade"] * df_fat_last["PrecoDiaria"] * df_fat_last["Dias"]
                    df_fat_top = (
                        df_fat_last.groupby("ItemCanon", as_index=False)["Faturamento"].sum()
                        .sort_values("Faturamento", ascending=False)
                        .head(3)
                    )
                    # Total de faturamento geral informado (para proporção correta)
                    total_rev_map = {
                        "Março": 96184.47,
                        "Abril": 92286.01,
                        "Maio": 87803.67,
                        "Junho": 77499.87,
                        "Julho": 81856.05,
                        "Agosto": 82609.95,
                    }
                    total_mes = float(total_rev_map.get(mes_label, df_fat_last["Faturamento"].sum()))
                    soma_top3 = float(df_fat_top["Faturamento"].sum())
                    outros_val = max(total_mes - soma_top3, 0.0)
                    df_pie_fat = df_fat_top.rename(columns={"ItemCanon":"Item"})[["Item","Faturamento"]]
                    if outros_val < 0:
                        outros_val = 0.0
                    # Sempre incluir "Outros" como 4ª variável (mesmo que 0)
                    df_pie_fat = pd.concat([
                        df_pie_fat,
                        pd.DataFrame([{ "Item": "Outros", "Faturamento": outros_val }])
                    ], ignore_index=True)
                    # ÚNICO gráfico: Top 3 + Outros
                    fig_pie_fat = px.pie(
                        df_pie_fat,
                        names="Item",
                        values="Faturamento",
                        hole=0.5,
                        color="Item",
                        color_discrete_map=color_map,
                        title=f"Impacto no orçamento – Top 3 + Outros (R$) – {mes_label}",
                    )
                    fig_pie_fat.update_layout(width=600, height=500)
                    fig_pie_fat.update_traces(
                        sort=False,
                        textposition="inside",
                        texttemplate="%{label}<br>R$ %{value:,.2f}",
                        hovertemplate="Item: %{label}<br>R$ %{value:,.2f} (%{percent})<extra></extra>",
                    )
                    fig_pie_fat.update_layout(margin=dict(l=20, r=60, t=60, b=20))
                    show_plot(fig_pie_fat, use_container_width=True)
        elif empresas_presentes_viz == ["PRONEP"]:
            st.subheader("Evolução PRONEP (Junho/Julho/Agosto)")
            df_pn = df_emp_viz[df_emp_viz["Empresa"] == "PRONEP"]
            if df_pn.empty:
                st.info("Sem dados para PRONEP nos meses 6/7/8")
            else:
                # Dataset consolidado apenas para auditoria (sem gráfico adicional)
                df_pn_aux = df_pn.copy()
                df_pn_aux["ItemCanon2"] = df_pn_aux["Item"].map(canonicalize_electric_bed_two_movements)

                # Gráfico Top 10 PRONEP removido a pedido

                # Auditoria removida a pedido

                df_pn_tot = (
                    df_pn.groupby("Mês", as_index=False)["Quantidade"].sum()
                    .set_index("Mês").reindex(month_order).reset_index().fillna(0)
                )
                fig_pn_line = px.line(
                    df_pn_tot,
                    x="Mês",
                    y="Quantidade",
                    markers=True,
                    title="Evolução mensal – PRONEP",
                )
                fig_pn_line.update_layout(width=800, height=400)
                fig_pn_line.update_layout(yaxis_title="Quantidade", xaxis_title="Mês")
                show_plot(fig_pn_line, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

                # Resumo por categorias – PRONEP
                df_pn_cat = df_pn.copy()
                if not df_pn_cat.empty:
                    df_pn_cat["Categoria"] = df_pn_cat["Item"].map(categorize_item_name)
                    alvo_cat_pn = ["CAMA", "CADEIRA HIGIÊNICA", "CADEIRA DE RODAS", "SUPORTE DE SORO"]
                    df_cat_pn = (
                        df_pn_cat[df_pn_cat["Categoria"].isin(alvo_cat_pn)]
                        .groupby(["Categoria"], as_index=False)["Quantidade"].sum()
                        .sort_values("Quantidade", ascending=False)
                    )
                    st.subheader("Resumo por categorias – PRONEP")
                    fig_cat_pn = px.bar(
                        df_cat_pn,
                        x="Categoria",
                        y="Quantidade",
                        title="Quantidade por categoria (Camas, Cadeira Higiene, Cadeira de Rodas, Suporte de Soro)",
                    )
                    fig_cat_pn.update_layout(width=800, height=400)
                    fig_cat_pn.update_layout(margin=dict(l=20, r=20, t=60, b=80))
                    show_plot(fig_cat_pn, use_container_width=True)

                meses_ordem = {"Junho": 6, "Julho": 7, "Agosto": 8}
                ultimo_mes = df_pn["Mês"].map(meses_ordem).max()
                mes_label = [k for k, v in meses_ordem.items() if v == ultimo_mes]
                mes_label = mes_label[0] if mes_label else "Junho"

                st.subheader(f"Participação dos Top 3 (quantidade) – PRONEP ({mes_label})")
                df_pn_last = df_pn[df_pn["Mês"] == mes_label].copy()
                canon_map_pn = {
                    normalize_text_for_match("CAMA ELÉTRICA 3 MOVIMENTOS"): "CAMA ELÉTRICA 3 MOVIMENTOS",
                    normalize_text_for_match("ARMÁRIO DE FÓRMICA"): "ARMÁRIO DE FÓRMICA",
                    normalize_text_for_match("COLCHÃO PNEUMÁTICO"): "COLCHÃO PNEUMÁTICO",
                }
                df_pn_last["key"] = df_pn_last["Item"].apply(normalize_text_for_match)
                df_pn_last["ItemCanon"] = df_pn_last["key"].map(canon_map_pn).fillna(df_pn_last["Item"]) 
                df_sum_pn = df_pn_last.groupby("ItemCanon", as_index=False)["Quantidade"].sum().sort_values("Quantidade", ascending=False)
                top3_pn = df_sum_pn.head(3)
                outros_pn = df_sum_pn["Quantidade"].sum() - top3_pn["Quantidade"].sum()
                df_pie_pn = top3_pn.rename(columns={"ItemCanon": "Item"})[["Item", "Quantidade"]].copy()
                if outros_pn > 0:
                    df_pie_pn = pd.concat([df_pie_pn, pd.DataFrame([{ "Item": "Outros", "Quantidade": outros_pn }])], ignore_index=True)
                pie_order_pn = df_pie_pn["Item"].tolist()
                color_map_pn = {
                    "CAMA ELÉTRICA 3 MOVIMENTOS": "#4e79a7",
                    "ARMÁRIO DE FÓRMICA": "#f28e2c",
                    "COLCHÃO PNEUMÁTICO": "#e15759",
                    "Outros": "#9aa3a8",
                }
                fig_pie_pn = px.pie(
                    df_pie_pn,
                    names="Item",
                    values="Quantidade",
                    hole=0.5,
                    color="Item",
                    color_discrete_map=color_map_pn,
                    category_orders={"Item": pie_order_pn},
                    title=f"Top 3 itens (quantidade) + Outros – PRONEP ({mes_label})",
                )
                fig_pie_pn.update_layout(width=600, height=500)
                fig_pie_pn.update_traces(
                    sort=False,
                    textposition="inside",
                    texttemplate="%{label}<br>%{percent}",
                    hovertemplate="Item: %{label}<br>Qtd: %{value:,} (%{percent})<extra></extra>",
                )
                fig_pie_pn.update_layout(
                    legend_title_text="Itens",
                    legend_orientation="v",
                    legend_y=0.5,
                    legend_x=1.02,
                    margin=dict(l=20, r=60, t=60, b=20),
                )
                st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                show_plot(fig_pie_pn, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

                
                # Rodapé
                st.markdown("---")
                st.markdown(
                    """
                    <div style='text-align: center; padding: 20px; color: #666; font-size: 14px;'>
                        <p><strong>Dashboard desenvolvido por Lucas Missiba</strong></p>
                        <p>Alocama · Setor de Contratos</p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )
        else:
            if empresas_presentes_viz and all(e in {"HOSPITALAR", "SOLAR", "DOMMUS"} for e in empresas_presentes_viz):
                meses_ordem = {"Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8}
                ultimo_mes = df_emp_viz["Mês"].map(meses_ordem).max()
                mes_label = [k for k, v in meses_ordem.items() if v == ultimo_mes]
                mes_label = mes_label[0] if mes_label else "Janeiro"

                st.subheader(f"Participação dos Top 3 (quantidade) – Grupo Solar ({mes_label})")
                df_gs_last = df_emp_viz[df_emp_viz["Mês"] == mes_label].copy()
                df_gs_last["ItemCanon"] = df_gs_last["Item"].map(canonicalize_trio_item)
                df_sum_gs = (
                    df_gs_last.groupby("ItemCanon", as_index=False)["Quantidade"].sum().sort_values("Quantidade", ascending=False)
                )
                top3_gs = df_sum_gs.head(3)
                outros_gs = df_sum_gs["Quantidade"].sum() - top3_gs["Quantidade"].sum()
                df_pie_gs = top3_gs.rename(columns={"ItemCanon": "Item"})[["Item", "Quantidade"]].copy()
                if outros_gs > 0:
                    df_pie_gs = pd.concat([df_pie_gs, pd.DataFrame([{ "Item": "Outros", "Quantidade": outros_gs }])], ignore_index=True)
                pie_order_gs = df_pie_gs["Item"].tolist()
                color_map = {
                    "CAMA MANUAL 2 MANIVELAS": "#4e79a7",
                    "SUPORTE DE SORO": "#f28e2c",
                    "CAMA ELÉTRICA 3 MOVIMENTOS": "#e15759",
                    "Outros": "#9aa3a8",
                }
                fig_pie_gs = px.pie(
                    df_pie_gs,
                    names="Item",
                    values="Quantidade",
                    hole=0.5,
                    color="Item",
                    color_discrete_map=color_map,
                    category_orders={"Item": pie_order_gs},
                    title=f"Top 3 itens (quantidade) + Outros – Grupo Solar ({mes_label})",
                )
                fig_pie_gs.update_layout(width=600, height=500)
                fig_pie_gs.update_traces(
                    sort=False,
                    textposition="inside",
                    texttemplate="%{label}<br>%{percent}",
                    hovertemplate="Item: %{label}<br>Qtd: %{value:,} (%{percent})<extra></extra>",
                )
                fig_pie_gs.update_layout(
                    legend_title_text="Itens",
                    legend_orientation="v",
                    legend_y=0.5,
                    legend_x=1.02,
                    margin=dict(l=20, r=60, t=60, b=20),
                )
                st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                show_plot(fig_pie_gs, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
                # Resumo por categorias – Grupo Solar (apenas último mês)
                meses_ordem = {"Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8}
                ultimo_mes = df_emp_viz["Mês"].map(meses_ordem).max()
                mes_label = [k for k, v in meses_ordem.items() if v == ultimo_mes]
                mes_label = mes_label[0] if mes_label else "Janeiro"
                
                df_gs_cat = df_emp_viz[df_emp_viz["Empresa"].isin(["HOSPITALAR","SOLAR","DOMMUS"])].copy()
                df_gs_cat = df_gs_cat[df_gs_cat["Mês"] == mes_label].copy()  # Filtrar apenas último mês
                if not df_gs_cat.empty:
                    df_gs_cat["Categoria"] = df_gs_cat["Item"].map(categorize_item_name)
                    alvo_cat = ["CAMA", "CADEIRA HIGIÊNICA", "CADEIRA DE RODAS", "SUPORTE DE SORO"]
                    df_cat_tot = (
                        df_gs_cat[df_gs_cat["Categoria"].isin(alvo_cat)]
                        .groupby(["Categoria"], as_index=False)["Quantidade"].sum()
                        .sort_values("Quantidade", ascending=False)
                    )
                    st.subheader(f"Resumo por categorias – Grupo Solar ({mes_label})")
                    fig_cat_gs = px.bar(
                        df_cat_tot,
                        x="Categoria",
                        y="Quantidade",
                        title=f"Quantidade por categoria (Camas, Cadeira Higiene, Cadeira de Rodas, Suporte de Soro) - {mes_label}",
                    )
                    fig_cat_gs.update_layout(width=800, height=400)
                    fig_cat_gs.update_layout(margin=dict(l=20, r=20, t=60, b=80))
                    show_plot(fig_cat_gs, use_container_width=True)
                    
                    
                # Gráfico de Pizza - Top 3 Itens de Agosto (Faturamento Real dos 3 Home Cares)
                st.subheader(f"Top 3 Itens por Faturamento – Agosto 2025 (SOLAR + DOMMUS + HOSPITALAR)")
                
                # Recriar df_gs_cat_detalhado para o gráfico de Top 3
                df_gs_cat_detalhado = df_gs_cat.copy()
                
                # Aplicar preços corretos
                precos_por_item = {
                    "CAMA DE 2 MOVIMENTOS ELÉTRICA": 334.80,  # R$ 10,80/dia × 31 dias = R$ 334,80/mês
                    "CAMA ELÉTRICA 2 MOVIMENTOS": 334.80,  # R$ 10,80/dia × 31 dias = R$ 334,80/mês
                    "CAMA ELÉTRICA 3 MOVIMENTOS": 334.80,  # R$ 10,80/dia × 31 dias = R$ 334,80/mês
                    "CAMA ELÉTRICA": 334.80,  # R$ 10,80/dia × 31 dias = R$ 334,80/mês
                    "COLCHÃO PNEUMÁTICO": 155.00,  # R$ 5,00/dia × 31 dias = R$ 155,00/mês
                    "CAMA": 150.00,  # Preço médio estimado para outras camas
                    "CADEIRA HIGIÊNICA": 80.00,  # Preço médio estimado por cadeira higiênica
                    "CADEIRA DE RODAS": 120.00,  # Preço médio estimado por cadeira de rodas
                    "CADEIRA DE RODAS SIMPLES": 120.00,  # Preço médio estimado por cadeira de rodas simples
                    "SUPORTE DE SORO": 60.00,  # Preço médio estimado por suporte de soro
                }
                
                df_gs_cat_detalhado["PrecoItem"] = df_gs_cat_detalhado["Item"].map(precos_por_item)
                df_gs_cat_detalhado["PrecoCategoria"] = df_gs_cat_detalhado["Categoria"].map({
                    "CAMA": 150.00,
                    "CADEIRA HIGIÊNICA": 80.00,
                    "CADEIRA DE RODAS": 120.00,
                    "SUPORTE DE SORO": 60.00,
                })
                df_gs_cat_detalhado["PrecoFinal"] = df_gs_cat_detalhado["PrecoItem"].fillna(df_gs_cat_detalhado["PrecoCategoria"])
                df_gs_cat_detalhado["FaturamentoItem"] = df_gs_cat_detalhado["Quantidade"] * df_gs_cat_detalhado["PrecoFinal"]
                
                # Filtrar apenas dados de agosto dos 3 home cares
                df_agosto = df_gs_cat_detalhado[df_gs_cat_detalhado["Mês"] == "Agosto"].copy()
                
                # Faturamento já foi calculado acima
                
                # Calcular faturamento total real de agosto
                faturamento_total_agosto = 312029.51  # Valor correto informado pelo usuário
                
                
                if not df_agosto.empty:
                    # Agrupar por item para obter faturamento total por item em agosto
                    df_itens_agosto = df_agosto.groupby("Item", as_index=False).agg({
                        "Quantidade": "sum",
                        "FaturamentoItem": "sum"
                    }).sort_values("FaturamentoItem", ascending=False)
                    
                    # Separar top 3 e calcular "Outros"
                    top_3_agosto = df_itens_agosto.head(3)
                    outros_fat_agosto = df_itens_agosto.iloc[3:]["FaturamentoItem"].sum()
                    outros_qtd_agosto = df_itens_agosto.iloc[3:]["Quantidade"].sum()
                    
                    # Criar DataFrame para o gráfico de pizza
                    df_pizza_agosto = top_3_agosto[["Item", "FaturamentoItem", "Quantidade"]].copy()
                    if outros_fat_agosto > 0:
                        df_pizza_agosto = pd.concat([
                            df_pizza_agosto,
                            pd.DataFrame({
                                "Item": ["Outros"],
                                "FaturamentoItem": [outros_fat_agosto],
                                "Quantidade": [outros_qtd_agosto]
                            })
                        ], ignore_index=True)
                    
                    # Calcular percentuais baseado no faturamento total de agosto
                    df_pizza_agosto["Percentual"] = (df_pizza_agosto["FaturamentoItem"] / faturamento_total_agosto * 100).round(1)
                    
                    # Cores mais bonitas e específicas para cada item
                    cores_personalizadas = {
                        "CAMA DE 2 MOVIMENTOS ELÉTRICA": "#FF6B6B",  # Vermelho coral
                        "CAMA ELÉTRICA 2 MOVIMENTOS": "#FF6B6B",  # Mesmo vermelho
                        "COLCHÃO PNEUMÁTICO": "#4ECDC4",  # Turquesa
                        "CADEIRA HIGIÊNICA": "#45B7D1",  # Azul claro
                        "CADEIRA DE RODAS": "#96CEB4",  # Verde menta
                        "SUPORTE DE SORO": "#FFEAA7",  # Amarelo claro
                        "Outros": "#DDA0DD"  # Ameixa
                    }
                    
                    # Mapear cores para os itens (com fallback para cores padrão)
                    df_pizza_agosto["Cor"] = df_pizza_agosto["Item"].map(cores_personalizadas)
                    
                    # Preencher cores faltantes com cores padrão
                    cores_padrao = ["#FF6B6B", "#4ECDC4", "#45B7D1", "#96CEB4", "#FFEAA7", "#DDA0DD"]
                    cores_finais = []
                    for i, cor in enumerate(df_pizza_agosto["Cor"]):
                        if pd.isna(cor):
                            cores_finais.append(cores_padrao[i % len(cores_padrao)])
                        else:
                            cores_finais.append(cor)
                    
                    # Criar gráfico de pizza mais bonito
                    fig_pizza_agosto = px.pie(
                        df_pizza_agosto,
                        values="FaturamentoItem",
                        names="Item",
                        title=f"Top 3 Itens por Faturamento – Agosto 2025<br><sub>Faturamento Total Real: R$ {faturamento_total_agosto:,.2f}</sub>",
                        hole=0.4,  # Buraco maior para visual mais moderno
                        color_discrete_sequence=cores_finais,
                    )
                    fig_pizza_agosto.update_layout(width=700, height=600)
                    
                    # Configurações mais bonitas
                    # Preparar customdata corretamente (após cálculo do percentual)
                    customdata_list = []
                    for _, row in df_pizza_agosto.iterrows():
                        customdata_list.append([row["Percentual"], row["Quantidade"]])
                    
                    fig_pizza_agosto.update_traces(
                        textinfo="percent+label",
                        textfont_size=12,
                        textfont_color="white",
                        textposition="outside",
                        hovertemplate="<b>%{label}</b><br>" +
                        "Faturamento: R$ %{value:,.2f}<br>" +
                        "Percentual do Total: %{customdata[0]:.1f}%<br>" +
                        "Quantidade: %{customdata[1]:,d} unidades<br>" +
                        "Faturamento Total Agosto: R$ " + f"{faturamento_total_agosto:,.2f}" + "<br>" +
                        "<extra></extra>",
                        customdata=customdata_list,
                        pull=[0.1 if "CAMA" in item or "COLCHÃO" in item else 0 for item in df_pizza_agosto["Item"]]
                    )
                    
                    fig_pizza_agosto.update_layout(
                        margin=dict(l=40, r=40, t=80, b=40),
                        font=dict(size=14, color="white"),
                        paper_bgcolor="rgba(0,0,0,0)",
                        plot_bgcolor="rgba(0,0,0,0)",
                        title_font_size=18,
                        title_font_color="white"
                    )
                    
                    show_plot(fig_pizza_agosto, use_container_width=True)
                    
                else:
                    st.warning(f"⚠️ Nenhum dado encontrado para {mes_label}")
                

            tabs = st.tabs(empresas_presentes_viz)
            for tab, empresa in zip(tabs, empresas_presentes_viz):
                with tab:
                    df_e = df_emp_viz[df_emp_viz["Empresa"] == empresa]
                    if df_e.empty:
                        st.info(f"Sem dados para {empresa} nos meses 6/7/8")
                    else:
                        # Dataset consolidado apenas para auditoria (sem gráfico adicional)
                        df_e_aux = df_e.copy()
                        df_e_aux["ItemCanon2"] = df_e_aux["Item"].map(canonicalize_electric_bed_two_movements)

                        top10_e = (
                            df_e.groupby("Item", as_index=False)["Quantidade"].sum()
                            .sort_values("Quantidade", ascending=False)["Item"].head(10).tolist()
                        )
                        df_e_top = df_e[df_e["Item"].isin(top10_e)]
                        fig_e_bar = px.bar(
                            df_e_top,
                            x="Item",
                            y="Quantidade",
                            color="Mês",
                            category_orders={"Mês": month_order, "Item": top10_e},
                            barmode="group",
                            title=f"Top 10 itens – {empresa}",
                            hover_data={"Mês": True, "Quantidade": ":,", "Item": True},
                        )
                        fig_e_bar.update_layout(
                            xaxis_title="Itens (Junho / Julho / Agosto)",
                            yaxis_title="Quantidade",
                            width=1200,  # Aumentar largura horizontal
                            height=400,  # Altura padrão
                            margin=dict(l=20, r=20, t=60, b=150),
                            showlegend=False,
                        )
                        fig_e_bar.update_xaxes(tickangle=-60)
                        st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                        show_plot(fig_e_bar, use_container_width=True)
                        st.markdown('</div>', unsafe_allow_html=True)

                        # Auditoria removida a pedido

                        df_e_tot = (
                            df_e.groupby("Mês", as_index=False)["Quantidade"].sum()
                            .set_index("Mês").reindex(month_order).reset_index().fillna(0)
                        )
                        fig_e_line = px.line(
                            df_e_tot,
                            x="Mês",
                            y="Quantidade",
                            markers=True,
                            title=f"Evolução mensal – {empresa}",
                        )
                        fig_e_line.update_layout(width=1200, height=400)
                        fig_e_line.update_layout(
                            yaxis_title="Quantidade", 
                            xaxis_title="Mês",
                            width=1200,  # Aumentar largura horizontal
                            height=350,  # Altura padrão
                            margin=dict(l=20, r=20, t=60, b=60)
                        )
                        st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                        show_plot(fig_e_line, use_container_width=True)
                        st.markdown('</div>', unsafe_allow_html=True)

        empresas_presentes_fat = sorted(df_emp_viz["Empresa"].unique().tolist())
        # Título central entre os gráficos de participação (pizza) e os gráficos de faturamento
        if empresas_presentes_fat:
            st.markdown("<h3 style='margin:12px 0 6px 0; text-align:center;'>FATURAMENTO</h3>", unsafe_allow_html=True)
        if empresas_presentes_fat == ["AXX CARE"]:
            st.subheader("Faturamento AXX CARE – Top 3 Itens (Fevereiro/Março/Abril/Maio/Junho/Julho/Agosto)")
            price_map = {
                normalize_text_for_match("CAMA ELÉTRICA 3 MOVIMENTOS"): 10.80,
                normalize_text_for_match("CAMA MANUAL 2 MANIVELAS"): 2.83,
                normalize_text_for_match("SUPORTE DE SORO"): 0.67,
            }
            canonical_map = {
                normalize_text_for_match("CAMA ELÉTRICA 3 MOVIMENTOS"): "CAMA ELÉTRICA 3 MOVIMENTOS",
                normalize_text_for_match("CAMA MANUAL 2 MANIVELAS"): "CAMA MANUAL 2 MANIVELAS",
                normalize_text_for_match("SUPORTE DE SORO"): "SUPORTE DE SORO",
            }
            df_rev = df_emp_viz[df_emp_viz["Empresa"] == "AXX CARE"].copy()
            def _strip_compl_prefix(text: str) -> str:
                t = str(text)
                return re.sub(r"^\s*\(?\s*compl\.?\s*\)?\s*", "", t, flags=re.IGNORECASE)
            df_rev = df_rev.assign(Item=df_rev["Item"].apply(_strip_compl_prefix))
            df_rev["key"] = df_rev["Item"].apply(normalize_text_for_match)
            df_rev["PrecoDiaria"] = df_rev["key"].map(price_map)
            df_rev = df_rev.dropna(subset=["PrecoDiaria"])  # mantém apenas os 3 itens
            if df_rev.empty:
                st.info("Sem ocorrências dos itens tarifados para AXX CARE nos meses 6/7/8.")
            else:
                df_rev["ItemCanonical"] = df_rev["key"].map(canonical_map)
                df_rev_sum = (
                    df_rev.groupby(["Empresa", "Mês", "ItemCanonical"], as_index=False)
                    .agg(Quantidade=("Quantidade", "sum"), PrecoDiaria=("PrecoDiaria", "first"))
                )
                # Completar meses ausentes com zero (Fevereiro→Agosto)
                months_for_fat = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]
                if not df_rev_sum.empty:
                    idx = pd.MultiIndex.from_product([["AXX CARE"], months_for_fat, df_rev_sum["ItemCanonical"].unique()], names=["Empresa","Mês","ItemCanonical"]).to_frame(index=False)
                    df_rev_sum = idx.merge(df_rev_sum, on=["Empresa","Mês","ItemCanonical"], how="left").fillna({"Quantidade":0})
                dias_map = {"Janeiro": 31, "Fevereiro": 28, "Março": 31, "Abril": 30, "Maio": 31, "Junho": 30, "Julho": 31, "Agosto": 31}
                df_rev_sum["Dias"] = df_rev_sum["Mês"].map(dias_map).fillna(30)
                df_rev_sum["Faturamento"] = df_rev_sum["Quantidade"] * df_rev_sum["PrecoDiaria"] * df_rev_sum["Dias"]

                item_order = [
                    canonical_map[normalize_text_for_match("CAMA ELÉTRICA 3 MOVIMENTOS")],
                    canonical_map[normalize_text_for_match("CAMA MANUAL 2 MANIVELAS")],
                    canonical_map[normalize_text_for_match("SUPORTE DE SORO")],
                ]
                fig_rev = px.bar(
                    df_rev_sum,
                    x="Mês",
                    y="Faturamento",
                    color="ItemCanonical",
                    facet_col="ItemCanonical",
                    category_orders={"Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho", "Julho", "Agosto"], "ItemCanonical": item_order},
                    title="Faturamento AXX CARE por Mês (diária x ocorrências)",
                    hover_data={"Faturamento": ":.2f", "Quantidade": True, "Dias": True},
                    labels={"ItemCanonical": "Item"},
                )
                fig_rev.update_layout(width=1200, height=600)
                fig_rev.update_layout(
                    yaxis_title="Faturamento (R$)", legend_title_text="Item",
                    legend_orientation="h", legend_y=-0.2, separators=".,",
                    margin=dict(l=20, r=20, t=60, b=80),
                )
                fig_rev.update_yaxes(tickprefix="R$ ", tickformat=".2f")
                st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                show_plot(fig_rev, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

                # Faturamento geral (AXX CARE) – valores informados
                st.subheader("Faturamento geral (AXX CARE) – Janeiro/Fevereiro/Março/Abril/Maio/Junho/Julho/Agosto")
                df_total_axx = pd.DataFrame({
                    "Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho", "Julho", "Agosto"],
                    "Faturamento": [98579.58, 87831.11, 96184.47, 92286.01, 87803.67, 77499.87, 81856.05, 82609.95],
                })
                fig_total_axx = px.bar(
                    df_total_axx,
                    x="Mês",
                    y="Faturamento",
                    text="Faturamento",
                    title="Faturamento geral por mês (valores fornecidos)",
                    category_orders={"Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]},
                )
                fig_total_axx.update_layout(width=1000, height=500)
                fig_total_axx.update_traces(texttemplate="R$ %{y:,.2f}", textposition="outside")
                ymax_axx = float(df_total_axx["Faturamento"].max())
                fig_total_axx.update_yaxes(tickprefix="R$ ", tickformat=",.2f", range=[0, ymax_axx * 1.15])
                fig_total_axx.update_layout(yaxis_title="Faturamento (R$)", margin=dict(l=20, r=20, t=80, b=60))
                show_plot(fig_total_axx, use_container_width=True)
        elif all(e in {"HOSPITALAR", "SOLAR", "DOMMUS"} for e in empresas_presentes_fat):
            st.subheader("Faturamento Grupo Solar – Top Itens (Janeiro→Agosto)")
            price_map_solar = {
                normalize_text_for_match("CAMA ELÉTRICA 3 MOVIMENTOS"): 10.80,
                normalize_text_for_match("SUPORTE DE SORO"): 0.67,
                normalize_text_for_match("COLCHÃO PNEUMÁTICO"): 5.00,
            }
            canonical_map_solar = {
                normalize_text_for_match("CAMA ELÉTRICA 3 MOVIMENTOS"): "CAMA ELÉTRICA 3 MOVIMENTOS",
                normalize_text_for_match("SUPORTE DE SORO"): "SUPORTE DE SORO",
                normalize_text_for_match("COLCHÃO PNEUMÁTICO"): "COLCHÃO PNEUMÁTICO",
            }
            df_gs = df_emp_viz[df_emp_viz["Empresa"] != "AXX CARE"].copy()
            df_gs["key"] = df_gs["Item"].apply(normalize_text_for_match)
            df_gs["PrecoDiaria"] = df_gs["key"].map(price_map_solar)
            df_gs = df_gs.dropna(subset=["PrecoDiaria"])  # mantém apenas itens tarifados
            if df_gs.empty:
                st.info("Sem dados tarifados para Grupo Solar nos meses 6/7/8.")
            else:
                df_gs["ItemCanonical"] = df_gs["key"].map(canonical_map_solar)
                df_gs_sum = (
                    df_gs.groupby(["Empresa", "Mês", "ItemCanonical"], as_index=False)
                    .agg(Quantidade=("Quantidade", "sum"), PrecoDiaria=("PrecoDiaria", "first"))
                )
                # Completar meses ausentes com zero (Janeiro→Agosto)
                months_for_fat_gs = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]
                empresas_gs = ["SOLAR", "HOSPITALAR", "DOMMUS"]
                if not df_gs_sum.empty:
                    idx = pd.MultiIndex.from_product([empresas_gs, months_for_fat_gs, df_gs_sum["ItemCanonical"].unique()], names=["Empresa","Mês","ItemCanonical"]).to_frame(index=False)
                    df_gs_sum = idx.merge(df_gs_sum, on=["Empresa","Mês","ItemCanonical"], how="left").fillna({"Quantidade":0})
                dias_map = {"Janeiro": 31, "Fevereiro": 28, "Março": 31, "Abril": 30, "Maio": 31, "Junho": 30, "Julho": 31, "Agosto": 31}
                df_gs_sum["Dias"] = df_gs_sum["Mês"].map(dias_map).fillna(30)
                df_gs_sum["Faturamento"] = df_gs_sum["Quantidade"] * df_gs_sum["PrecoDiaria"] * df_gs_sum["Dias"]
                
                item_order_gs = [
                    canonical_map_solar[normalize_text_for_match("CAMA ELÉTRICA 3 MOVIMENTOS")],
                    canonical_map_solar[normalize_text_for_match("SUPORTE DE SORO")],
                    canonical_map_solar[normalize_text_for_match("COLCHÃO PNEUMÁTICO")],
                ]
                df_gs_total = (
                    df_gs_sum.groupby(["Mês", "ItemCanonical"], as_index=False)["Faturamento"].sum()
                )
                fig_gs = px.bar(
                    df_gs_total,
                    x="Mês", y="Faturamento", color="ItemCanonical",
                    facet_col="ItemCanonical",
                    category_orders={"Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho", "Julho", "Agosto"], "ItemCanonical": item_order_gs},
                    title="Faturamento por Mês – Grupo Solar (diária x ocorrências)",
                    hover_data={"Faturamento": ":.2f"},
                    labels={"ItemCanonical": "Item"},
                )
                fig_gs.update_layout(yaxis_title="Faturamento (R$)", legend_title_text="Item",
                                     legend_orientation="h", legend_y=-0.2, separators=".,",
                                     margin=dict(l=20, r=20, t=60, b=80))
                fig_gs.update_yaxes(tickprefix="R$ ", tickformat=".2f")
                st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                show_plot(fig_gs, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)
                # Faturamento geral (Grupo Solar) – valores informados
                st.subheader("Faturamento geral (Grupo Solar) – Janeiro→Agosto")
                df_total_gs = pd.DataFrame({
                    "Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho", "Julho", "Agosto"],
                    "Faturamento": [290194.24, 267454.14, 298768.28, 286571.08, 294592.68, 287981.61, 309546.25, 312029.51],
                })
                fig_total_gs = px.bar(
                    df_total_gs,
                    x="Mês",
                    y="Faturamento",
                    text="Faturamento",
                    title="Faturamento geral por mês (valores fornecidos)",
                    category_orders={"Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho", "Julho", "Agosto"]},
                )
                fig_total_gs.update_traces(texttemplate="R$ %{y:,.2f}", textposition="outside")
                ymax_gs = float(df_total_gs["Faturamento"].max())
                fig_total_gs.update_yaxes(tickprefix="R$ ", tickformat=",.2f", range=[0, ymax_gs * 1.15])
                fig_total_gs.update_layout(yaxis_title="Faturamento (R$)", margin=dict(l=20, r=20, t=80, b=60))
                show_plot(fig_total_gs, use_container_width=True)
                
                df_rev_sum = df_gs_sum

                # ================================
                # ARPU - Faturamento por vida (Grupo Solar)
                # ================================
                try:
                    st.subheader("ARPU - Faturamento por vida (Grupo Solar)")
                    month_labels_gs = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]
                    month_sets_arpu_gs = {m: set() for m in month_labels_gs}
                    mes_label_gs = {"2025-01":"Janeiro","2025-02":"Fevereiro","2025-03":"Março","2025-04":"Abril","2025-05":"Maio","2025-06": "Junho", "2025-07": "Julho", "2025-08": "Agosto"}
                    
                    for file in sel_files:
                        ym = year_month_from_path(file)
                        if ym not in {"2025-01","2025-02","2025-03","2025-04", "2025-05", "2025-06", "2025-07", "2025-08"}:
                            continue
                        mes_label = mes_label_gs.get(ym, None)
                        if mes_label is None:
                            continue
                        group = primary_group_from_label(str(file))
                        if group not in ["HOSPITALAR", "SOLAR", "DOMMUS"]:
                            continue
                        try:
                            book = safe_read_excel(file, sheet_name=None)
                        except Exception:
                            continue
                        for sheet_name, df_sheet in book.items():
                            if should_exclude_sheet(sheet_name):
                                continue
                            series = select_best_name_column(df_sheet)
                            if series is None:
                                continue
                            series = series.dropna().astype(str).str.strip()
                            series = series[series != ""]
                            if series.empty:
                                continue
                            nomes_norm = series.apply(normalize_text_for_match)
                            month_sets_arpu_gs[mes_label].update(nomes_norm.tolist())
                    
                    df_vidas_arpu_gs = pd.DataFrame({
                        "Mês": month_labels_gs,
                        "Vidas": [len(month_sets_arpu_gs[m]) for m in month_labels_gs],
                    })
                    # Faturamento informado manualmente por mês (Grupo Solar)
                    total_rev_map_gs = {
                        "Janeiro": 290194.24,
                        "Fevereiro": 267454.14,
                        "Março": 298768.28,
                        "Abril": 286571.08,
                        "Maio": 294592.68,
                        "Junho": 287981.61,
                        "Julho": 309546.25,
                        "Agosto": 312029.51,
                    }
                    rev_df_gs = pd.DataFrame({"Mês": month_labels_gs, "Faturamento": [total_rev_map_gs.get(m, 0.0) for m in month_labels_gs]})
                    df_arpu_gs = rev_df_gs.merge(df_vidas_arpu_gs, on="Mês", how="left").fillna(0)
                    df_arpu_gs["ARPU"] = df_arpu_gs.apply(lambda r: (float(r["Faturamento"]) / r["Vidas"]) if r["Vidas"] > 0 else 0.0, axis=1)
                    fig_arpu_gs = px.bar(
                        df_arpu_gs, x="Mês", y="ARPU",
                        title="ARPU por mês (Geral – Grupo Solar)",
                        category_orders={"Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]},
                    )
                    fig_arpu_gs.update_yaxes(tickprefix="R$ ", tickformat=",.2f")
                    show_plot(fig_arpu_gs, use_container_width=True)
                except Exception:
                    pass
                
                # ================================
                # Ticket Médio por Cliente (Grupo Solar)
                # ================================
                try:
                    st.subheader("Ticket Médio por Cliente – Grupo Solar")
                    
                    # Calcular ticket médio mensal
                    month_labels_ticket = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]
                    ticket_medio_data = []
                    
                    # Faturamento consolidado por mês
                    faturamento_mensal = {
                        "Janeiro": 290194.24,
                        "Fevereiro": 267454.14,
                        "Março": 298768.28,
                        "Abril": 286571.08,
                        "Maio": 294592.68,
                        "Junho": 287981.61,
                        "Julho": 309546.25,
                        "Agosto": 312029.51,
                    }
                    
                    # Coletar vidas ativas por mês
                    month_sets_ticket = {m: set() for m in month_labels_ticket}
                    mes_label_ticket = {"2025-01":"Janeiro","2025-02":"Fevereiro","2025-03":"Março","2025-04":"Abril","2025-05":"Maio","2025-06": "Junho", "2025-07": "Julho", "2025-08": "Agosto"}
                    
                    for file in sel_files:
                        ym = year_month_from_path(file)
                        if ym not in {"2025-01","2025-02","2025-03","2025-04", "2025-05", "2025-06", "2025-07", "2025-08"}:
                            continue
                        mes_label = mes_label_ticket.get(ym, None)
                        if mes_label is None:
                            continue
                        group = primary_group_from_label(str(file))
                        if group not in ["HOSPITALAR", "SOLAR", "DOMMUS"]:
                            continue
                        try:
                            book = safe_read_excel(file, sheet_name=None)
                        except Exception:
                            continue
                        for sheet_name, df_sheet in book.items():
                            if should_exclude_sheet(sheet_name):
                                continue
                            series = select_best_name_column(df_sheet)
                            if series is None:
                                continue
                            series = series.dropna().astype(str).str.strip()
                            series = series[series != ""]
                            if series.empty:
                                continue
                            nomes_norm = series.apply(normalize_text_for_match)
                            month_sets_ticket[mes_label].update(nomes_norm.tolist())
                    
                    # Calcular ticket médio
                    for mes in month_labels_ticket:
                        vidas = len(month_sets_ticket[mes])
                        faturamento = faturamento_mensal.get(mes, 0)
                        ticket_medio = (faturamento / vidas) if vidas > 0 else 0
                        ticket_medio_data.append({
                            "Mês": mes,
                            "Ticket Médio": ticket_medio,
                            "Vidas": vidas,
                            "Faturamento": faturamento
                        })
                    
                    df_ticket_medio = pd.DataFrame(ticket_medio_data)
                    
                    fig_ticket = px.bar(
                        df_ticket_medio, x="Mês", y="Ticket Médio",
                        title="Ticket Médio por Cliente – Grupo Solar",
                        category_orders={"Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]},
                        color="Ticket Médio",
                        color_continuous_scale="Viridis"
                    )
                    fig_ticket.update_traces(
                        hovertemplate="<b>%{x}</b><br>" +
                        "Ticket Médio: R$ %{y:,.2f}<br>" +
                        "Vidas Ativas: %{customdata[0]:,}<br>" +
                        "Faturamento: R$ %{customdata[1]:,.2f}<br>" +
                        "<extra></extra>",
                        customdata=df_ticket_medio[["Vidas", "Faturamento"]]
                    )
                    fig_ticket.update_yaxes(tickprefix="R$ ", tickformat=",.2f")
                    show_plot(fig_ticket, use_container_width=True)
                except Exception:
                    pass

                # ================================
                # Heatmap Item × Mês (Grupo Solar)
                # ================================
                try:
                    st.subheader("Heatmap Item × Mês – Grupo Solar")
                    df_heat_gs = (
                        df_emp_viz[df_emp_viz["Empresa"].isin(["HOSPITALAR", "SOLAR", "DOMMUS"])][["Mês","Item","Quantidade"]]
                        .groupby(["Item","Mês"], as_index=False)["Quantidade"].sum()
                    )
                    # Mantém somente top 20 itens por soma para foco visual
                    tops_gs = (
                        df_heat_gs.groupby("Item")["Quantidade"].sum()
                        .sort_values(ascending=False).head(20).index.tolist()
                    )
                    df_heat_gs = df_heat_gs[df_heat_gs["Item"].isin(tops_gs)]
                    df_pvt_gs = df_heat_gs.pivot(index="Item", columns="Mês", values="Quantidade").fillna(0)
                    df_pvt_gs = df_pvt_gs[[m for m in ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"] if m in df_pvt_gs.columns]]
                    fig_heat_gs = px.imshow(
                        df_pvt_gs,
                        color_continuous_scale="Blues",
                        aspect="auto",
                        labels=dict(color="Quantidade"),
                        title="Intensidade de ocorrências por item e mês (Grupo Solar)",
                    )
                    show_plot(fig_heat_gs, use_container_width=True)
                except Exception:
                    pass

                # ================================
                # Gráfico de Pizza - Faturamento por Empresa (Grupo Solar)
                # ================================
                try:
                    st.subheader("Distribuição de Faturamento por Empresa (Grupo Solar)")
                    # Calcular faturamento total por empresa
                    faturamento_por_empresa = df_rev_sum.groupby("Empresa")["Faturamento"].sum().reset_index()
                    faturamento_por_empresa = faturamento_por_empresa.sort_values("Faturamento", ascending=False)
                    
                    # Ajustar para o valor total real (R$ 2.347.137,79)
                    total_real = 2347137.79
                    total_calculado = faturamento_por_empresa["Faturamento"].sum()
                    fator_ajuste = total_real / total_calculado
                    
                    # Aplicar o fator de ajuste mantendo as proporções
                    faturamento_por_empresa["Faturamento_Ajustado"] = faturamento_por_empresa["Faturamento"] * fator_ajuste
                    
                    
                    fig_pizza_gs = px.pie(
                        faturamento_por_empresa,
                        values="Faturamento_Ajustado",
                        names="Empresa",
                        title="Participação no Faturamento - Grupo Solar (Últimos 8 Meses)",
                        hole=0.3,
                        color_discrete_sequence=px.colors.qualitative.Set3
                    )
                    # Preparar customdata para o gráfico do Grupo Solar
                    customdata_gs = []
                    for _, row in faturamento_por_empresa.iterrows():
                        customdata_gs.append([row["Faturamento_Ajustado"]])
                    
                    fig_pizza_gs.update_traces(
                        textposition='inside',
                        textinfo='percent+label',
                        hovertemplate='<b>%{label}</b><br>' +
                        'Faturamento: R$ %{value:,.2f}<br>' +
                        'Participação: %{percent}<br>' +
                        'Valor Absoluto: R$ %{customdata[0]:,.2f}<br>' +
                        '<extra></extra>',
                        customdata=customdata_gs
                    )
                    fig_pizza_gs.update_layout(
                        margin=dict(l=20, r=20, t=60, b=20),
                        showlegend=True,
                        legend=dict(
                            orientation="v",
                            yanchor="middle",
                            y=0.5,
                            xanchor="left",
                            x=1.01
                        )
                    )
                    st.plotly_chart(fig_pizza_gs, use_container_width=True)
                    
                except Exception as e:
                    st.error(f"❌ **Erro na seção do Grupo Solar**: {str(e)}")
                    import traceback
                    st.write(f"**Traceback completo**: {traceback.format_exc()}")

                # ================================
                # ARPU Consolidado - Faturamento por vida (3 Home Cares)
                # ================================
                # ARPU Consolidado - Faturamento por vida (3 Home Cares)
                # ================================
                try:
                    st.subheader("ARPU Consolidado - Faturamento por vida (3 Home Cares)")
                    
                    month_labels_consolidado = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]
                    month_sets_arpu_consolidado = {m: set() for m in month_labels_consolidado}
                    mes_label_consolidado = {"2025-01":"Janeiro","2025-02":"Fevereiro","2025-03":"Março","2025-04":"Abril","2025-05":"Maio","2025-06": "Junho", "2025-07": "Julho", "2025-08": "Agosto"}
                    
                    # Debug: contar arquivos processados
                    arquivos_processados = 0
                    total_arquivos = len(sel_files)
                    
                    for file in sel_files:
                        ym = year_month_from_path(file)
                        if ym not in {"2025-01","2025-02","2025-03","2025-04", "2025-05", "2025-06", "2025-07", "2025-08"}:
                            continue
                        mes_label = mes_label_consolidado.get(ym, None)
                        if mes_label is None:
                            continue
                        group = primary_group_from_label(str(file))
                        if group not in ["HOSPITALAR", "SOLAR", "DOMMUS", "AXX CARE"]:
                            continue
                        
                        try:
                            book = safe_read_excel(file, sheet_name=None)
                            arquivos_processados += 1
                        except Exception as e:
                            st.warning(f"⚠️ Erro ao ler arquivo {file.name}: {str(e)}")
                            continue
                            
                        for sheet_name, df_sheet in book.items():
                            if should_exclude_sheet(sheet_name):
                                continue
                            name_col = select_best_name_column(df_sheet)
                            if name_col is None:
                                continue
                            try:
                                series = df_sheet[name_col].dropna().astype(str).str.strip()
                                series = series[series != ""]
                                if series.empty:
                                    continue
                                nomes_norm = series.apply(normalize_text_for_match)
                                month_sets_arpu_consolidado[mes_label].update(nomes_norm.tolist())
                            except Exception as e:
                                st.warning(f"⚠️ Erro ao processar coluna {name_col} em {file.name}: {str(e)}")
                                continue
                    
                    # Debug: mostrar estatísticas
                    st.info(f"📊 **Arquivos processados**: {arquivos_processados} de {total_arquivos}")
                    
                    df_vidas_arpu_consolidado = pd.DataFrame({
                        "Mês": month_labels_consolidado,
                        "Vidas": [len(month_sets_arpu_consolidado[m]) for m in month_labels_consolidado],
                    })
                    
                    # Faturamento consolidado dos 3 home cares
                    total_rev_map_consolidado = {
                        "Janeiro": 290194.24,
                        "Fevereiro": 267454.14,
                        "Março": 298768.28,
                        "Abril": 286571.08,
                        "Maio": 294592.68,
                        "Junho": 287981.61,
                        "Julho": 309546.25,
                        "Agosto": 312029.51,
                    }
                    
                    rev_df_consolidado = pd.DataFrame({
                        "Mês": month_labels_consolidado, 
                        "Faturamento": [total_rev_map_consolidado.get(m, 0.0) for m in month_labels_consolidado]
                    })
                    
                    df_arpu_consolidado = rev_df_consolidado.merge(df_vidas_arpu_consolidado, on="Mês", how="left").fillna(0)
                    df_arpu_consolidado["ARPU"] = df_arpu_consolidado.apply(
                        lambda r: (float(r["Faturamento"]) / r["Vidas"]) if r["Vidas"] > 0 else 0.0, axis=1
                    )
                    
                    
                    # Verificar se há dados válidos
                    if df_arpu_consolidado["Vidas"].sum() == 0:
                        st.error("❌ **Problema**: Nenhuma vida ativa encontrada nos arquivos!")
                        st.write("**Possíveis causas:**")
                        st.write("- Arquivos não estão sendo lidos corretamente")
                        st.write("- Filtros estão excluindo todos os dados")
                        st.write("- Estrutura dos arquivos Excel mudou")
                    else:
                        fig_arpu_consolidado = px.bar(
                            df_arpu_consolidado, x="Mês", y="ARPU",
                            title="ARPU Consolidado por mês (3 Home Cares)",
                            category_orders={"Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]},
                            color="ARPU",
                            color_continuous_scale="Viridis"
                        )
                        fig_arpu_consolidado.update_yaxes(tickprefix="R$ ", tickformat=",.2f")
                        fig_arpu_consolidado.update_traces(
                            hovertemplate="<b>%{x}</b><br>" +
                            "ARPU: R$ %{y:,.2f}<br>" +
                            "Vidas: %{customdata[0]:,}<br>" +
                            "Faturamento: R$ %{customdata[1]:,.2f}<br>" +
                            "<extra></extra>",
                            customdata=df_arpu_consolidado[["Vidas", "Faturamento"]]
                        )
                        show_plot(fig_arpu_consolidado, use_container_width=True)
                        
                        # Estatísticas resumidas
                        arpu_medio = df_arpu_consolidado["ARPU"].mean()
                        arpu_max = df_arpu_consolidado["ARPU"].max()
                        arpu_min = df_arpu_consolidado["ARPU"].min()
                        
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("ARPU Médio", f"R$ {arpu_medio:,.2f}")
                        with col2:
                            st.metric("ARPU Máximo", f"R$ {arpu_max:,.2f}")
                        with col3:
                            st.metric("ARPU Mínimo", f"R$ {arpu_min:,.2f}")
                            
                except Exception as e:
                    st.error(f"❌ **Erro no ARPU Consolidado**: {str(e)}")
                    st.write("**Detalhes do erro:**")
                    st.code(str(e))

            st.subheader("Faturamento por empresa (Janeiro→Agosto)")
            col_h, col_s, col_d = st.columns(3)
            for empresa, col in [("HOSPITALAR", col_h), ("SOLAR", col_s), ("DOMMUS", col_d)]:
                with col:
                    sub = df_rev_sum[df_rev_sum["Empresa"] == empresa]
                    if sub.empty:
                        st.info(f"Sem dados para {empresa}")
                    else:
                        fig_e = px.bar(
                            sub,
                            x="Mês",
                            y="Faturamento",
                            color="ItemCanonical",
                            barmode="group",
                            category_orders={"Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho", "Julho", "Agosto"], "ItemCanonical": item_order},
                            title=empresa,
                            hover_data={"Faturamento": ":.2f", "Quantidade": True},
                        )
                        fig_e.update_layout(
                            yaxis_title="Faturamento (R$)",
                            showlegend=False,
                        )
                        fig_e.update_yaxes(tickprefix="R$ ", tickformat=".2f")
                        st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                        show_plot(fig_e, use_container_width=True)
                        st.markdown('</div>', unsafe_allow_html=True)

        if empresas_presentes_fat == ["PRONEP"]:
            st.subheader("Faturamento PRONEP – Top Itens (Janeiro→Agosto)")
            price_map_pronep = {
                normalize_text_for_match("CAMA ELÉTRICA 3 MOVIMENTOS"): 10.30,
                normalize_text_for_match("ARMÁRIO DE FÓRMICA"): 2.80,
                normalize_text_for_match("COLCHÃO PNEUMÁTICO"): 5.00,
            }
            canonical_map_pronep = {
                normalize_text_for_match("CAMA ELÉTRICA 3 MOVIMENTOS"): "CAMA ELÉTRICA 3 MOVIMENTOS",
                normalize_text_for_match("ARMÁRIO DE FÓRMICA"): "ARMÁRIO DE FÓRMICA",
                normalize_text_for_match("COLCHÃO PNEUMÁTICO"): "COLCHÃO PNEUMÁTICO",
            }
            df_pn = df_emp_viz[df_emp_viz["Empresa"] == "PRONEP"].copy()
            df_pn["key"] = df_pn["Item"].apply(normalize_text_for_match)
            df_pn["PrecoDiaria"] = df_pn["key"].map(price_map_pronep)
            df_pn = df_pn.dropna(subset=["PrecoDiaria"])  # mantém apenas os 3 itens
            if df_pn.empty:
                st.info("Sem ocorrências dos itens tarifados para PRONEP nos meses 1/2/3/4/5/6/7/8.")
                df_pn_sum = pd.DataFrame()  # DataFrame vazio para evitar erro
            else:
                df_pn["ItemCanonical"] = df_pn["key"].map(canonical_map_pronep)
                df_pn_sum = (
                    df_pn.groupby(["Empresa", "Mês", "ItemCanonical"], as_index=False)
                    .agg(Quantidade=("Quantidade", "sum"), PrecoDiaria=("PrecoDiaria", "first"))
                )
                dias_map = {"Janeiro": 31, "Fevereiro": 28, "Março": 31, "Abril": 30, "Maio": 31, "Junho": 30, "Julho": 31, "Agosto": 31}
                df_pn_sum["Dias"] = df_pn_sum["Mês"].map(dias_map).fillna(30)
                df_pn_sum["Faturamento"] = df_pn_sum["Quantidade"] * df_pn_sum["PrecoDiaria"] * df_pn_sum["Dias"]

            # ================================
            # GRÁFICO DE PIZZA - TOP 3 FATURAMENTO PRONEP (AGOSTO)
            # ================================
            if not df_pn_sum.empty and "Agosto" in df_pn_sum["Mês"].values:
                st.subheader("Participação dos Top 3 (faturamento) – PRONEP (Agosto)")
                
                # Filtrar dados de agosto da PRONEP
                df_agosto_pronep = df_pn_sum[df_pn_sum["Mês"] == "Agosto"].copy()
                
                if not df_agosto_pronep.empty:
                    # Calcular faturamento total por item em agosto
                    df_fat_agosto = df_agosto_pronep.groupby("ItemCanonical", as_index=False)["Faturamento"].sum()
                    df_fat_agosto = df_fat_agosto.sort_values("Faturamento", ascending=False)
                    
                    # Pegar top 3
                    top3_agosto = df_fat_agosto.head(3)
                    
                    # Valor total correto informado pelo usuário
                    total_fat_agosto = 54204.32
                    
                    # Calcular "Outros" como diferença para completar o total
                    fat_top3 = top3_agosto["Faturamento"].sum()
                    outros_fat = total_fat_agosto - fat_top3
                    
                    # Criar DataFrame para o gráfico de pizza
                    df_pizza_agosto = top3_agosto.copy()
                    if outros_fat > 0:
                        df_pizza_agosto = pd.concat([
                            df_pizza_agosto,
                            pd.DataFrame({"ItemCanonical": ["Outros"], "Faturamento": [outros_fat]})
                        ], ignore_index=True)
                    
                    # Calcular percentuais baseado no total correto
                    df_pizza_agosto["Percentual"] = (df_pizza_agosto["Faturamento"] / total_fat_agosto * 100).round(1)
                    
                    # Cores personalizadas
                    cores_pizza = {
                        "CAMA ELÉTRICA 3 MOVIMENTOS": "#1f77b4",
                        "ARMÁRIO DE FÓRMICA": "#ff7f0e", 
                        "COLCHÃO PNEUMÁTICO": "#2ca02c",
                        "Outros": "#d62728"
                    }
                    
                    df_pizza_agosto["Cor"] = df_pizza_agosto["ItemCanonical"].map(cores_pizza)
                    
                    # Criar gráfico de pizza
                    fig_pizza_agosto = px.pie(
                        df_pizza_agosto,
                        values="Faturamento",
                        names="ItemCanonical",
                        title="Top 3 itens (faturamento) + Outros - PRONEP (Agosto)",
                        hole=0.4,
                        color_discrete_sequence=df_pizza_agosto["Cor"].tolist()
                    )
                    
                    # Preparar customdata para o gráfico PRONEP
                    customdata_pronep = []
                    for _, row in df_pizza_agosto.iterrows():
                        customdata_pronep.append([row["Percentual"]])
                    
                    # Atualizar layout
                    fig_pizza_agosto.update_traces(
                        textposition='inside',
                        textinfo='percent+label',
                        hovertemplate="<b>%{label}</b><br>" +
                                     "Faturamento: R$ %{value:,.2f}<br>" +
                                     "Percentual do Total: %{customdata[0]:.1f}%<br>" +
                                     "Faturamento Total Agosto: R$ " + f"{total_fat_agosto:,.2f}" + "<br>" +
                                     "<extra></extra>",
                        customdata=customdata_pronep
                    )
                    
                    fig_pizza_agosto.update_layout(
                        margin=dict(l=20, r=20, t=60, b=20),
                        showlegend=True,
                        legend=dict(
                            orientation="v",
                            yanchor="middle",
                            y=0.5,
                            xanchor="left",
                            x=1.01
                        )
                    )
                    
                    show_plot(fig_pizza_agosto, use_container_width=True)
                    
                    # Mostrar resumo
                    st.write("**Resumo do Faturamento - Agosto 2025:**")
                    for _, row in df_pizza_agosto.iterrows():
                        st.write(f"• **{row['ItemCanonical']}**: R$ {row['Faturamento']:,.2f} ({row['Percentual']}%)")
                    
                    st.success(f"💰 **Total Faturamento Agosto**: R$ {total_fat_agosto:,.2f}")
                    
                else:
                    st.info("ℹ️ Sem dados de faturamento para agosto da PRONEP")

                item_order_pn = [
                    canonical_map_pronep[normalize_text_for_match("CAMA ELÉTRICA 3 MOVIMENTOS")],
                    canonical_map_pronep[normalize_text_for_match("ARMÁRIO DE FÓRMICA")],
                    canonical_map_pronep[normalize_text_for_match("COLCHÃO PNEUMÁTICO")],
                ]
                fig_pn_rev = px.bar(
                    df_pn_sum,
                    x="Mês",
                    y="Faturamento",
                    color="ItemCanonical",
                    facet_col="ItemCanonical",
                    category_orders={"Mês": ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto"], "ItemCanonical": item_order_pn},
                    title="Faturamento PRONEP por Mês (diária x ocorrências)",
                    hover_data={"Faturamento": ":.2f", "Quantidade": True, "Dias": True},
                    labels={"ItemCanonical": "Item"},
                )
                fig_pn_rev.update_layout(
                    yaxis_title="Faturamento (R$)", legend_title_text="Item",
                    legend_orientation="h", legend_y=-0.2, separators=".,",
                    margin=dict(l=20, r=20, t=60, b=80),
                )
                fig_pn_rev.update_yaxes(tickprefix="R$ ", tickformat=".2f")
                st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                show_plot(fig_pn_rev, use_container_width=True)
                st.markdown('</div>', unsafe_allow_html=True)

                # Faturamento geral (valores informados)
                st.subheader("Faturamento geral (PRONEP) – Janeiro→Agosto")
                df_total_pronep = pd.DataFrame({
                    "Mês": ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto"],
                    "Faturamento": [64280.11, 55913.26, 61094.47, 57437.30, 59120.20, 55664.48, 56251.60, 54204.32],
                })
                fig_total_pronep = px.bar(
                    df_total_pronep,
                    x="Mês",
                    y="Faturamento",
                    text="Faturamento",
                    title="Faturamento geral por mês (valores fornecidos)",
                )
                fig_total_pronep.update_traces(texttemplate="R$ %{y:,.2f}", textposition="outside")
                ymax_pronep = float(df_total_pronep["Faturamento"].max())
                fig_total_pronep.update_yaxes(tickprefix="R$ ", tickformat=",.2f", range=[0, ymax_pronep * 1.15])
                fig_total_pronep.update_layout(yaxis_title="Faturamento (R$)", margin=dict(l=20, r=20, t=80, b=60))
                show_plot(fig_total_pronep, use_container_width=True)
                df_rev_sum_pronep = df_pn_sum

                # ================================
                # VIDAS ATIVAS PRONEP
                # ================================
                st.subheader("Vidas ativas no Home Care – PRONEP (Janeiro→Agosto)")
                month_sets = {"Janeiro": set(), "Fevereiro": set(), "Março": set(), "Abril": set(), "Maio": set(), "Junho": set(), "Julho": set(), "Agosto": set()}
                mes_label_map = {"2025-01": "Janeiro", "2025-02": "Fevereiro", "2025-03": "Março", "2025-04": "Abril", "2025-05": "Maio", "2025-06": "Junho", "2025-07": "Julho", "2025-08": "Agosto"}
                
                # Debug: mostrar arquivos processados
                arquivos_pronep = 0
                total_arquivos = len(sel_files)
                
                for file in sel_files:
                    try:
                        book = safe_read_excel(file, sheet_name=None)
                    except Exception:
                        continue
                    ym = year_month_from_path(file)
                    if ym not in {"2025-01", "2025-02", "2025-03", "2025-04", "2025-05", "2025-06", "2025-07", "2025-08"}:
                        continue
                    mes_label = mes_label_map.get(ym, None)
                    if mes_label not in month_sets:
                        continue
                        
                    # Verificar se é arquivo da PRONEP
                    group = primary_group_from_label(str(file))
                    if group != "PRONEP":
                        continue
                        
                    arquivos_pronep += 1
                        
                    for sheet_name, df_sheet in (book or {}).items():
                        if should_exclude_sheet(str(sheet_name)):
                            continue
                        if not isinstance(df_sheet, pd.DataFrame) or df_sheet.empty:
                            continue
                        series = None
                        try:
                            if df_sheet.shape[1] >= 2:
                                series = df_sheet.iloc[:, 1]
                        except Exception:
                            series = None
                        if series is None:
                            name_col = select_best_name_column(df_sheet)
                            if not name_col:
                                continue
                            series = df_sheet[name_col]
                        series = series.dropna().astype(str).str.strip()
                        series = series[series != ""]
                        if series.empty:
                            continue
                        nomes_norm = series.apply(normalize_text_for_match)
                        month_sets[mes_label].update(nomes_norm.tolist())
                        
                
                df_vidas_mes = pd.DataFrame({
                    "Mês": ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto"],
                    "VidasUnicas": [len(month_sets[m]) for m in ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto"]],
                })
                
                media_vidas = (
                    df_vidas_mes["VidasUnicas"][df_vidas_mes["VidasUnicas"] > 0].mean() if (df_vidas_mes["VidasUnicas"] > 0).any() else 0
                )
                fig_vidas = px.bar(
                    df_vidas_mes,
                    x="Mês",
                    y="VidasUnicas",
                    title="Vidas ativas únicas por mês (PRONEP)",
                    text="VidasUnicas",
                )
                fig_vidas.update_traces(textposition="outside")
                fig_vidas.update_layout(yaxis_title="Vidas únicas", xaxis_title="Mês", margin=dict(l=20, r=20, t=60, b=40))
                show_plot(fig_vidas, use_container_width=True)
                target = int(round(media_vidas))
                placeholder = st.empty()
                for val in range(0, target + 1, max(1, target // 30)):
                    placeholder.metric("Média de vidas ativas (8 meses)", f"{val}")
                    time.sleep(0.02)
                if target % max(1, target // 30) != 0:
                    placeholder.metric("Média de vidas ativas (8 meses)", f"{target}")


                # ================================
                # ARPU - Faturamento por vida (PRONEP)
                # ================================
                try:
                    st.subheader("ARPU - Faturamento por vida (PRONEP)")
                    month_labels_pronep = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto"]
                    month_sets_arpu_pronep = {m: set() for m in month_labels_pronep}
                    mes_label_pronep = {"2025-01": "Janeiro", "2025-02": "Fevereiro", "2025-03": "Março", "2025-04": "Abril", "2025-05": "Maio", "2025-06": "Junho", "2025-07": "Julho", "2025-08": "Agosto"}
                    
                    # Debug: mostrar arquivos processados
                    arquivos_processados = 0
                    total_arquivos = len(sel_files)
                    
                    for file in sel_files:
                        ym = year_month_from_path(file)
                        if ym not in {"2025-01", "2025-02", "2025-03", "2025-04", "2025-05", "2025-06", "2025-07", "2025-08"}:
                            continue
                        mes_label = mes_label_pronep.get(ym, None)
                        if mes_label is None:
                            continue
                        group = primary_group_from_label(str(file))
                        if group != "PRONEP":
                            continue
                        
                        arquivos_processados += 1
                        
                        try:
                            book = safe_read_excel(file, sheet_name=None)
                        except Exception as e:
                            st.warning(f"⚠️ Erro ao ler arquivo {file.name}: {str(e)}")
                            continue
                        
                        for sheet_name, df_sheet in (book or {}).items():
                            if should_exclude_sheet(str(sheet_name)):
                                continue
                            if not isinstance(df_sheet, pd.DataFrame) or df_sheet.empty:
                                continue
                            
                            # Tentar diferentes métodos para encontrar a coluna de nomes
                            series = None
                            try:
                                # Método 1: Segunda coluna
                                if df_sheet.shape[1] >= 2:
                                    series = df_sheet.iloc[:, 1]
                                    if series.dropna().empty:
                                        series = None
                            except Exception:
                                series = None
                            
                            if series is None:
                                # Método 2: Usar função select_best_name_column
                                try:
                                    name_col = select_best_name_column(df_sheet)
                                    if name_col:
                                        series = df_sheet[name_col]
                                except Exception as e:
                                    st.warning(f"⚠️ Erro ao selecionar coluna em {sheet_name}: {str(e)}")
                                    continue
                            
                            if series is None or series.empty:
                                continue
                                
                            # Limpar e processar dados
                            series = series.dropna().astype(str).str.strip()
                            series = series[series != ""]
                            if series.empty:
                                continue
                            
                            # Normalizar nomes
                            try:
                                nomes_norm = series.apply(normalize_text_for_match)
                                month_sets_arpu_pronep[mes_label].update(nomes_norm.tolist())
                            except Exception as e:
                                st.warning(f"⚠️ Erro ao normalizar nomes em {sheet_name}: {str(e)}")
                                continue
                    
                    
                    df_vidas_arpu_pronep = pd.DataFrame({
                        "Mês": month_labels_pronep,
                        "Vidas": [len(month_sets_arpu_pronep[m]) for m in month_labels_pronep]
                    })
                    
                    # Mostrar dados de vidas para debug
                    
                    # Faturamento por mês (valores fornecidos)
                    rev_df_pronep = pd.DataFrame({
                        "Mês": month_labels_pronep,
                        "Faturamento": [64280.11, 55913.26, 61094.47, 57437.30, 59120.20, 55664.48, 56251.60, 54204.32]
                    })
                    
                    df_arpu_pronep = rev_df_pronep.merge(df_vidas_arpu_pronep, on="Mês", how="left").fillna(0)
                    df_arpu_pronep["ARPU"] = df_arpu_pronep.apply(
                        lambda r: (float(r["Faturamento"]) / r["Vidas"]) if r["Vidas"] > 0 else 0.0, axis=1
                    )
                    
                    if df_arpu_pronep["Vidas"].sum() == 0:
                        st.error("❌ **Problema**: Nenhuma vida ativa encontrada nos arquivos da PRONEP!")
                        st.info("💡 **Sugestão**: Verifique se os arquivos da PRONEP estão sendo carregados corretamente e se contêm dados de pacientes.")
                    else:
                        fig_arpu_pronep = px.bar(
                            df_arpu_pronep, x="Mês", y="ARPU",
                            title="ARPU (Average Revenue Per User) - PRONEP",
                            text="ARPU",
                            color="ARPU",
                            color_continuous_scale="Viridis"
                        )
                        fig_arpu_pronep.update_traces(
                            texttemplate="R$ %{text:,.2f}",
                            textposition="outside",
                            hovertemplate="<b>%{x}</b><br>ARPU: R$ %{y:,.2f}<br>Vidas: %{customdata[0]}<br>Faturamento: R$ %{customdata[1]:,.2f}<extra></extra>",
                            customdata=df_arpu_pronep[["Vidas", "Faturamento"]].values
                        )
                        fig_arpu_pronep.update_layout(
                            yaxis_title="ARPU (R$)",
                            xaxis_title="Mês",
                            margin=dict(l=20, r=20, t=60, b=40)
                        )
                        fig_arpu_pronep.update_yaxes(tickprefix="R$ ", tickformat=".2f")
                        show_plot(fig_arpu_pronep, use_container_width=True)
                        
                        # Métricas do ARPU
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("ARPU Médio", f"R$ {df_arpu_pronep['ARPU'].mean():,.2f}")
                        with col2:
                            st.metric("ARPU Máximo", f"R$ {df_arpu_pronep['ARPU'].max():,.2f}")
                        with col3:
                            st.metric("ARPU Mínimo", f"R$ {df_arpu_pronep['ARPU'].min():,.2f}")
                            
                except Exception as e:
                    st.error(f"❌ **Erro no ARPU PRONEP**: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())

                # ================================
                # PROJEÇÕES - PRONEP
                # ================================
                st.markdown("<h3 style='text-align:center; margin-top:32px;'>PROJEÇÕES - PRONEP</h3>", unsafe_allow_html=True)
                
                # Projeção de Vidas Atendidas
                st.subheader("Projeção de Vidas Atendidas")
                
                # Calcular tendência baseada nos últimos 3 meses
                ultimos_3_meses_vidas = df_vidas_arpu_pronep.tail(3)["Vidas"].tolist()
                tendencia_vidas = (ultimos_3_meses_vidas[-1] - ultimos_3_meses_vidas[0]) / 2
                
                # Cenários para vidas
                vidas_otimista_pronep = []
                vidas_realista_pronep = []
                vidas_pessimista_pronep = []
                
                ultima_vida_pronep = df_vidas_arpu_pronep["Vidas"].iloc[-1]
                meses_projecao_pronep = ["Setembro", "Outubro", "Novembro"]
                
                for i in range(3):
                    # Otimista: crescimento de 3% ao mês
                    vidas_otimista_pronep.append(int(ultima_vida_pronep * (1.03 ** (i+1))))
                    # Realista: manutenção ou pequena queda (era pessimista)
                    vidas_realista_pronep.append(max(int(ultima_vida_pronep - 2 * (i+1)), int(ultima_vida_pronep * 0.95)))
                    # Pessimista: tendência atual + pequeno crescimento (era realista)
                    vidas_pessimista_pronep.append(int(ultima_vida_pronep + tendencia_vidas * (i+1) + 5))
                
                # Criar DataFrame para projeção de vidas
                df_proj_vidas_pronep = pd.DataFrame({
                    "Mês": meses_projecao_pronep,
                    "Otimista": vidas_otimista_pronep,
                    "Realista": vidas_realista_pronep,
                    "Pessimista": vidas_pessimista_pronep
                })
                
                # Gráfico de projeção de vidas - Design moderno
                fig_proj_vidas_pronep = px.line(
                    df_proj_vidas_pronep,
                    x="Mês",
                    y=["Otimista", "Realista", "Pessimista"],
                    title="<b>Projeção de Vidas Atendidas</b><br><sub>Próximos 3 Meses</sub>",
                    labels={"value": "Vidas Atendidas", "variable": "Cenário"},
                    color_discrete_sequence=["#00D4AA", "#FF6B6B", "#4ECDC4"]
                )
                
                # Atualizar layout com design moderno
                fig_proj_vidas_pronep.update_layout(
                    yaxis_title="<b>Vidas Atendidas</b>",
                    xaxis_title="<b>Mês</b>",
                    margin=dict(l=40, r=40, t=80, b=60),
                    legend_title="<b>Cenários</b>",
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font=dict(family="Arial, sans-serif", size=12),
                    title_font_size=18,
                    title_x=0.5,
                    hovermode='x unified',
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    )
                )
                
                # Formatação dos eixos
                fig_proj_vidas_pronep.update_xaxes(
                    showgrid=True,
                    gridwidth=1,
                    gridcolor='rgba(128,128,128,0.2)',
                    showline=True,
                    linewidth=2,
                    linecolor='rgba(128,128,128,0.3)'
                )
                fig_proj_vidas_pronep.update_yaxes(
                    showgrid=True,
                    gridwidth=1,
                    gridcolor='rgba(128,128,128,0.2)',
                    showline=True,
                    linewidth=2,
                    linecolor='rgba(128,128,128,0.3)',
                    tickformat=".0f"
                )
                
                # Adicionar marcadores nos pontos
                fig_proj_vidas_pronep.update_traces(
                    mode='lines+markers',
                    marker=dict(size=8, line=dict(width=2, color='white')),
                    line=dict(width=3)
                )
                
                show_plot(fig_proj_vidas_pronep, use_container_width=True)
                
                # Projeção de Faturamento
                st.subheader("Projeção de Faturamento")
                
                # Cenários para faturamento
                fat_otimista_pronep = []
                fat_realista_pronep = []
                fat_pessimista_pronep = []
                
                ultimo_fat_pronep = df_total_pronep["Faturamento"].iloc[-1]
                
                for i in range(3):
                    # Otimista: crescimento de 5% ao mês
                    fat_otimista_pronep.append(round(ultimo_fat_pronep * (1.05 ** (i+1)), 2))
                    # Realista: pequena queda de 1% ao mês (era pessimista)
                    fat_realista_pronep.append(round(ultimo_fat_pronep * (0.99 ** (i+1)), 2))
                    # Pessimista: manutenção com pequeno crescimento de 2% ao mês (era realista)
                    fat_pessimista_pronep.append(round(ultimo_fat_pronep * (1.02 ** (i+1)), 2))
                
                # Criar DataFrame para projeção de faturamento
                df_proj_fat_pronep = pd.DataFrame({
                    "Mês": meses_projecao_pronep,
                    "Otimista": fat_otimista_pronep,
                    "Realista": fat_realista_pronep,
                    "Pessimista": fat_pessimista_pronep
                })
                
                # Gráfico de projeção de faturamento - Design moderno
                fig_proj_fat_pronep = px.line(
                    df_proj_fat_pronep,
                    x="Mês",
                    y=["Otimista", "Realista", "Pessimista"],
                    title="<b>Projeção de Faturamento</b><br><sub>Próximos 3 Meses</sub>",
                    labels={"value": "Faturamento (R$)", "variable": "Cenário"},
                    color_discrete_sequence=["#00D4AA", "#FF6B6B", "#4ECDC4"]
                )
                
                # Atualizar layout com design moderno
                fig_proj_fat_pronep.update_layout(
                    yaxis_title="<b>Faturamento (R$)</b>",
                    xaxis_title="<b>Mês</b>",
                    margin=dict(l=40, r=40, t=80, b=60),
                    legend_title="<b>Cenários</b>",
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    font=dict(family="Arial, sans-serif", size=12),
                    title_font_size=18,
                    title_x=0.5,
                    hovermode='x unified',
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    )
                )
                
                # Formatação dos eixos
                fig_proj_fat_pronep.update_xaxes(
                    showgrid=True,
                    gridwidth=1,
                    gridcolor='rgba(128,128,128,0.2)',
                    showline=True,
                    linewidth=2,
                    linecolor='rgba(128,128,128,0.3)'
                )
                fig_proj_fat_pronep.update_yaxes(
                    showgrid=True,
                    gridwidth=1,
                    gridcolor='rgba(128,128,128,0.2)',
                    showline=True,
                    linewidth=2,
                    linecolor='rgba(128,128,128,0.3)',
                    tickformat="R$ ,.0f"
                )
                
                # Adicionar marcadores nos pontos
                fig_proj_fat_pronep.update_traces(
                    mode='lines+markers',
                    marker=dict(size=8, line=dict(width=2, color='white')),
                    line=dict(width=3)
                )
                
                # Formatação do hover para valores monetários
                fig_proj_fat_pronep.update_traces(
                    hovertemplate="<b>%{fullData.name}</b><br>" +
                                 "Mês: %{x}<br>" +
                                 "Faturamento: R$ %{y:,.2f}<br>" +
                                 "<extra></extra>"
                )
                
                show_plot(fig_proj_fat_pronep, use_container_width=True)

        if empresas_presentes_viz and all(e in {"HOSPITALAR", "SOLAR", "DOMMUS"} for e in empresas_presentes_viz):
            st.subheader("Vidas ativas no Home Care – Grupo Solar (Janeiro→Agosto)")
            month_sets = {"Janeiro": set(), "Fevereiro": set(), "Março": set(), "Abril": set(), "Maio": set(), "Junho": set(), "Julho": set(), "Agosto": set()}
            for file in sel_files:
                try:
                    book = safe_read_excel(file, sheet_name=None)
                except Exception:
                    continue
                ym = year_month_from_path(file)
                if ym not in {"2025-01","2025-02","2025-03","2025-04", "2025-05", "2025-06", "2025-07", "2025-08"}:
                    continue
                mes_label = {"2025-01":"Janeiro","2025-02":"Fevereiro","2025-03":"Março","2025-04":"Abril","2025-05":"Maio","2025-06": "Junho", "2025-07": "Julho", "2025-08": "Agosto"}.get(ym, None)
                if mes_label not in month_sets:
                    continue
                for sheet_name, df_sheet in (book or {}).items():
                    if should_exclude_sheet(str(sheet_name)):
                        continue
                    if not isinstance(df_sheet, pd.DataFrame) or df_sheet.empty:
                        continue
                    series = None
                    try:
                        if df_sheet.shape[1] >= 2:
                            # Para arquivos SOLAR com colunas "Unnamed", pular primeira linha (cabeçalho)
                            if any("Unnamed" in str(col) for col in df_sheet.columns):
                                series = df_sheet.iloc[1:, 1]  # Pular linha 0, usar coluna B
                            else:
                                series = df_sheet.iloc[:, 1]
                    except Exception:
                        series = None
                    if series is None:
                        name_col = select_best_name_column(df_sheet)
                        if not name_col:
                            continue
                        series = df_sheet[name_col]
                    series = series.dropna().astype(str).str.strip()
                    series = series[series != ""]
                    if series.empty:
                        continue
                    nomes_norm = series.apply(normalize_text_for_match)
                    month_sets[mes_label].update(nomes_norm.tolist())
            df_vidas_mes = pd.DataFrame({
                "Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho", "Julho", "Agosto"],
                "VidasUnicas": [
                    len(month_sets.get("Janeiro", set())),
                    len(month_sets.get("Fevereiro", set())),
                    len(month_sets.get("Março", set())),
                    len(month_sets.get("Abril", set())),
                    len(month_sets.get("Maio", set())),
                    len(month_sets.get("Junho", set())),
                    len(month_sets.get("Julho", set())),
                    len(month_sets.get("Agosto", set())),
                ],
            })
            media_vidas = (
                df_vidas_mes["VidasUnicas"][df_vidas_mes["VidasUnicas"] > 0].mean() if (df_vidas_mes["VidasUnicas"] > 0).any() else 0
            )
            fig_vidas = px.bar(
                df_vidas_mes,
                x="Mês",
                y="VidasUnicas",
                title="Vidas ativas únicas por mês (Grupo Solar)",
                text="VidasUnicas",
            )
            fig_vidas.update_traces(textposition="outside")
            fig_vidas.update_layout(yaxis_title="Vidas únicas", xaxis_title="Mês", margin=dict(l=20, r=20, t=60, b=40))
            show_plot(fig_vidas, use_container_width=True)
            target = int(round(media_vidas))
            placeholder = st.empty()
            for val in range(0, target + 1, max(1, target // 30)):
                placeholder.metric("Média de vidas ativas (8 meses)", f"{val}")
                time.sleep(0.02)
            if target % max(1, target // 30) != 0:
                placeholder.metric("Média de vidas ativas (8 meses)", f"{target}")
            
            # ===============================================================
            # PROJEÇÕES - Grupo Solar
            # ===============================================================
            st.markdown("<h3 style='text-align:center; margin-top:32px;'>PROJEÇÕES</h3>", unsafe_allow_html=True)
            st.markdown("<p style='text-align:center; margin-top:-12px;'>Projeções para os próximos 3 meses – Grupo Solar</p>", unsafe_allow_html=True)
            
            # Dados históricos para projeções
            df_historico_vidas = df_vidas_mes.copy()
            df_historico_faturamento = pd.DataFrame({
                "Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"],
                "Faturamento": [312029.51, 312029.51, 312029.51, 312029.51, 312029.51, 312029.51, 312029.51, 312029.51],  # Usando valor real de agosto
            })
            
            # Projeções para os próximos 3 meses
            meses_projecao = ["Setembro", "Outubro", "Novembro"]
            
            # Projeção de Vidas Atendidas
            st.subheader("Projeção de Vidas Atendidas")
            
            # Calcular tendência baseada nos últimos 3 meses
            ultimos_3_meses = df_historico_vidas.tail(3)["VidasUnicas"].tolist()
            tendencia = (ultimos_3_meses[-1] - ultimos_3_meses[0]) / 2  # Tendência linear
            
            # Cenários para vidas
            vidas_otimista = []
            vidas_realista = []
            vidas_pessimista = []
            
            ultima_vida = df_historico_vidas["VidasUnicas"].iloc[-1]
            
            for i in range(3):
                # Otimista: crescimento de 3% ao mês
                vidas_otimista.append(int(ultima_vida * (1.03 ** (i+1))))
                # Realista: tendência atual + pequeno crescimento
                vidas_realista.append(int(ultima_vida + tendencia * (i+1) + 5))
                # Pessimista: manutenção ou pequena queda
                vidas_pessimista.append(max(int(ultima_vida - 2 * (i+1)), int(ultima_vida * 0.95)))
            
            # Criar DataFrame para projeção de vidas
            df_proj_vidas = pd.DataFrame({
                "Mês": meses_projecao,
                "Otimista": vidas_otimista,
                "Realista": vidas_realista,
                "Pessimista": vidas_pessimista
            })
            
            # Gráfico de projeção de vidas - Design moderno
            fig_proj_vidas = px.line(
                df_proj_vidas,
                x="Mês",
                y=["Otimista", "Realista", "Pessimista"],
                title="<b>Projeção de Vidas Atendidas</b><br><sub>Próximos 3 Meses</sub>",
                labels={"value": "Vidas Atendidas", "variable": "Cenário"},
                color_discrete_sequence=["#00D4AA", "#FF6B6B", "#4ECDC4"]
            )
            
            # Atualizar layout com design moderno
            fig_proj_vidas.update_layout(
                yaxis_title="<b>Vidas Atendidas</b>",
                xaxis_title="<b>Mês</b>",
                margin=dict(l=40, r=40, t=80, b=60),
                legend_title="<b>Cenários</b>",
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(family="Arial, sans-serif", size=12),
                title_font_size=18,
                title_x=0.5,
                hovermode='x unified',
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1
                )
            )
            
            # Formatação dos eixos
            fig_proj_vidas.update_xaxes(
                showgrid=True,
                gridwidth=1,
                gridcolor='rgba(128,128,128,0.2)',
                showline=True,
                linewidth=2,
                linecolor='rgba(128,128,128,0.3)'
            )
            fig_proj_vidas.update_yaxes(
                showgrid=True,
                gridwidth=1,
                gridcolor='rgba(128,128,128,0.2)',
                showline=True,
                linewidth=2,
                linecolor='rgba(128,128,128,0.3)',
                tickformat=".0f"
            )
            
            # Adicionar marcadores nos pontos
            fig_proj_vidas.update_traces(
                mode='lines+markers',
                marker=dict(size=8, line=dict(width=2, color='white')),
                line=dict(width=3)
            )
            
            show_plot(fig_proj_vidas, use_container_width=True)
            
            # Projeção de Faturamento
            st.subheader("Projeção de Faturamento")
            
            # Cenários para faturamento
            fat_otimista = []
            fat_realista = []
            fat_pessimista = []
            
            ultimo_fat = df_historico_faturamento["Faturamento"].iloc[-1]
            
            for i in range(3):
                # Otimista: crescimento de 5% ao mês
                fat_otimista.append(round(ultimo_fat * (1.05 ** (i+1)), 2))
                # Realista: manutenção com pequeno crescimento de 2% ao mês
                fat_realista.append(round(ultimo_fat * (1.02 ** (i+1)), 2))
                # Pessimista: pequena queda de 1% ao mês
                fat_pessimista.append(round(ultimo_fat * (0.99 ** (i+1)), 2))
            
            # Criar DataFrame para projeção de faturamento
            df_proj_fat = pd.DataFrame({
                "Mês": meses_projecao,
                "Otimista": fat_otimista,
                "Realista": fat_realista,
                "Pessimista": fat_pessimista
            })
            
            # Gráfico de projeção de faturamento - Design moderno
            fig_proj_fat = px.line(
                df_proj_fat,
                x="Mês",
                y=["Otimista", "Realista", "Pessimista"],
                title="<b>Projeção de Faturamento</b><br><sub>Próximos 3 Meses</sub>",
                labels={"value": "Faturamento (R$)", "variable": "Cenário"},
                color_discrete_sequence=["#00D4AA", "#FF6B6B", "#4ECDC4"]
            )
            
            # Atualizar layout com design moderno
            fig_proj_fat.update_layout(
                yaxis_title="<b>Faturamento (R$)</b>",
                xaxis_title="<b>Mês</b>",
                margin=dict(l=40, r=40, t=80, b=60),
                legend_title="<b>Cenários</b>",
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(family="Arial, sans-serif", size=12),
                title_font_size=18,
                title_x=0.5,
                hovermode='x unified',
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1
                )
            )
            
            # Formatação dos eixos
            fig_proj_fat.update_xaxes(
                showgrid=True,
                gridwidth=1,
                gridcolor='rgba(128,128,128,0.2)',
                showline=True,
                linewidth=2,
                linecolor='rgba(128,128,128,0.3)'
            )
            fig_proj_fat.update_yaxes(
                showgrid=True,
                gridwidth=1,
                gridcolor='rgba(128,128,128,0.2)',
                showline=True,
                linewidth=2,
                linecolor='rgba(128,128,128,0.3)',
                tickformat="R$ ,.0f"
            )
            
            # Adicionar marcadores nos pontos
            fig_proj_fat.update_traces(
                mode='lines+markers',
                marker=dict(size=8, line=dict(width=2, color='white')),
                line=dict(width=3)
            )
            
            # Formatação do hover para valores monetários
            fig_proj_fat.update_traces(
                hovertemplate="<b>%{fullData.name}</b><br>" +
                             "Mês: %{x}<br>" +
                             "Faturamento: R$ %{y:,.2f}<br>" +
                             "<extra></extra>"
            )
            
            show_plot(fig_proj_fat, use_container_width=True)
        elif empresas_presentes_viz == ["AXX CARE"]:
            st.subheader("Vidas ativas no Home Care – AXX CARE (Janeiro→Agosto)")
            month_sets = {"Janeiro": set(), "Fevereiro": set(), "Março": set(), "Abril": set(), "Maio": set(), "Junho": set(), "Julho": set(), "Agosto": set()}
            for file in sel_files:
                try:
                    book = safe_read_excel(file, sheet_name=None)
                except Exception:
                    continue
                ym = year_month_from_path(file)
                if ym not in {"2025-01","2025-02","2025-03","2025-04", "2025-05", "2025-06", "2025-07", "2025-08"}:
                    continue
                mes_label = {"2025-01":"Janeiro","2025-02":"Fevereiro","2025-03":"Março","2025-04":"Abril","2025-05":"Maio","2025-06": "Junho", "2025-07": "Julho", "2025-08": "Agosto"}.get(ym, None)
                if mes_label not in month_sets:
                    continue
                for sheet_name, df_sheet in (book or {}).items():
                    if should_exclude_sheet(str(sheet_name)):
                        continue
                    if not isinstance(df_sheet, pd.DataFrame) or df_sheet.empty:
                        continue
                    series = None
                    try:
                        if df_sheet.shape[1] >= 2:
                            series = df_sheet.iloc[:, 1]
                    except Exception:
                        series = None
                    if series is None:
                        name_col = select_best_name_column(df_sheet)
                        if not name_col:
                            continue
                        series = df_sheet[name_col]
                    series = series.dropna().astype(str).str.strip()
                    series = series[series != ""]
                    if series.empty:
                        continue
                    nomes_norm = series.apply(normalize_text_for_match)
                    month_sets[mes_label].update(nomes_norm.tolist())
            df_vidas_mes = pd.DataFrame({
                "Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho", "Julho", "Agosto"],
                "VidasUnicas": [
                    len(month_sets.get("Janeiro", set())),
                    len(month_sets.get("Fevereiro", set())),
                    len(month_sets.get("Março", set())),
                    len(month_sets.get("Abril", set())),
                    len(month_sets.get("Maio", set())),
                    len(month_sets.get("Junho", set())),
                    len(month_sets.get("Julho", set())),
                    len(month_sets.get("Agosto", set())),
                ],
            })
            # Ordena eixo X começando em Janeiro → Agosto
            _order_fev_to_aug = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]
            try:
                df_vidas_mes["Mês"] = pd.Categorical(df_vidas_mes["Mês"], categories=_order_fev_to_aug, ordered=True)
                df_vidas_mes = df_vidas_mes.sort_values("Mês")
            except Exception:
                pass
            media_vidas = (
                df_vidas_mes["VidasUnicas"][df_vidas_mes["VidasUnicas"] > 0].mean() if (df_vidas_mes["VidasUnicas"] > 0).any() else 0
            )
            fig_vidas = px.bar(
                df_vidas_mes,
                x="Mês",
                y="VidasUnicas",
                title="Vidas ativas únicas por mês (AXX CARE)",
                text="VidasUnicas",
            )
            fig_vidas.update_traces(textposition="outside")
            fig_vidas.update_layout(yaxis_title="Vidas únicas", xaxis_title="Mês", margin=dict(l=20, r=20, t=60, b=40))
            show_plot(fig_vidas, use_container_width=True)
            target = int(round(media_vidas))
            placeholder = st.empty()
            for val in range(0, target + 1, max(1, target // 30)):
                placeholder.metric("Média de vidas ativas (3 meses)", f"{val}")
                time.sleep(0.02)
            if target % max(1, target // 30) != 0:
                placeholder.metric("Média de vidas ativas (3 meses)", f"{target}")


        # ===============================================================
        # PREVISÃO – AXX CARE baseada no Faturamento Geral informado
        # ===============================================================
        if "AXX CARE" in df_emp_viz["Empresa"].unique().tolist():
            st.markdown("<h3 style='text-align:center; margin-top:32px;'>PREVISÃO</h3>", unsafe_allow_html=True)
            st.markdown("<p style='text-align:center; margin-top:-12px;'>Receita projetada para os próximos 3 meses – AXX CARE (base: Faturamento Geral)</p>", unsafe_allow_html=True)

            # Série histórica do Faturamento geral (valores fornecidos)
            df_total_axx_hist = pd.DataFrame({
                "Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"],
                "Faturamento": [98579.58, 87831.11, 96184.47, 92286.01, 87803.67, 77499.87, 81856.05, 82609.95],
            })
            ordem = {m:i for i,m in enumerate(["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]) }
            df_total_axx_hist["ord"] = df_total_axx_hist["Mês"].map(ordem)
            df_total_axx_hist = df_total_axx_hist.sort_values("ord").drop(columns=["ord"]).reset_index(drop=True)

            # Projeções com base nessa série
            vals = df_total_axx_hist["Faturamento"].tolist()
            def mov_avg_next(v, k=3):
                buf = v[-k:] if len(v) >= k else v
                return (sum(buf)/len(buf)) if buf else 0
            proj_realista = []
            cur = vals[:]
            for _ in range(3):
                nxt = mov_avg_next(cur, 3)
                proj_realista.append(nxt)
                cur.append(nxt)
            proj_otimista = [v * 1.07 for v in proj_realista]

            # Próximos 3 meses com nomes reais
            meses_pt = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]
            last_label = df_total_axx_hist["Mês"].iloc[-1] if not df_total_axx_hist.empty else "Agosto"
            try:
                idx = meses_pt.index(last_label)
            except ValueError:
                idx = 7  # Agosto
            proximos_meses = [meses_pt[(idx + i) % 12] for i in range(1, 4)]

            # Linhas – histórico vs projeções
            fig_l = go.Figure()
            fig_l.add_trace(go.Scatter(x=df_total_axx_hist["Mês"], y=df_total_axx_hist["Faturamento"], name="Real", mode="lines+markers", line=dict(color="#3b82f6")))
            fig_l.add_trace(go.Scatter(x=proximos_meses, y=proj_realista, name="Realista", mode="lines+markers", line=dict(color="#f59e0b", dash="dash")))
            fig_l.add_trace(go.Scatter(x=proximos_meses, y=proj_otimista, name="Otimista", mode="lines+markers", line=dict(color="#10b981", dash="dot")))
            fig_l.update_layout(title="Tendência histórica + Projeção (AXX CARE – Geral)", yaxis_title="Faturamento (R$)")
            fig_l.update_yaxes(tickprefix="R$ ", tickformat=",.2f")
            show_plot(fig_l, use_container_width=True)

            # Barras – últimos 3 meses vs próximos 3 (realista)
            df_bar_hist = df_total_axx_hist.tail(3).copy()
            df_bar_hist["Tipo"] = "Histórico"
            df_bar_proj = pd.DataFrame({"Mês": proximos_meses, "Faturamento": proj_realista, "Tipo": "Projetado (Realista)"})
            df_bar = pd.concat([df_bar_hist[["Mês","Faturamento","Tipo"]], df_bar_proj], ignore_index=True)
            fig_b = px.bar(df_bar, x="Mês", y="Faturamento", color="Tipo", barmode="group", title="Últimos 3 meses vs próximos 3 (Realista – Geral)")
            fig_b.update_yaxes(tickprefix="R$ ", tickformat=",.2f")
            show_plot(fig_b, use_container_width=True)

            # KPIs de projeção
            k1, k2, k3 = st.columns(3)
            k1.metric(f"{proximos_meses[0]} (Realista)", f"R$ {proj_realista[0]:,.2f}")
            k2.metric(f"{proximos_meses[1]} (Realista)", f"R$ {proj_realista[1]:,.2f}")
            k3.metric(f"{proximos_meses[2]} (Realista)", f"R$ {proj_realista[2]:,.2f}")

            # Seção Holt‑Winters removida a pedido (estava duplicada)

            # Área de análise: por que caiu/subiu e o que fazer
            try:
                meses_ordem_full = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]
                ultimo = df_total_axx_hist["Mês"].iloc[-1]
                idx_u = meses_ordem_full.index(ultimo) if ultimo in meses_ordem_full else len(meses_ordem_full) - 1
                anterior = meses_ordem_full[idx_u - 1] if idx_u > 0 else ultimo
                fat_u = float(df_total_axx_hist.loc[df_total_axx_hist["Mês"] == ultimo, "Faturamento"].fillna(0).iloc[0])
                fat_a = float(df_total_axx_hist.loc[df_total_axx_hist["Mês"] == anterior, "Faturamento"].fillna(0).iloc[0])
                delta_abs = fat_u - fat_a
                delta_pct = (delta_abs / fat_a * 100.0) if fat_a > 0 else None
                direcao = "subiu" if delta_abs >= 0 else "caiu"

                # Drivers por item (quando disponível)
                top_pos, top_neg = [], []
                if 'df_rev_sum' in locals() and isinstance(df_rev_sum, pd.DataFrame) and not df_rev_sum.empty:
                    last_item = df_rev_sum[df_rev_sum["Mês"] == ultimo].groupby("ItemCanonical")["Faturamento"].sum()
                    prev_item = df_rev_sum[df_rev_sum["Mês"] == anterior].groupby("ItemCanonical")["Faturamento"].sum()
                    items = sorted(set(last_item.index).union(set(prev_item.index)))
                    deltas = []
                    for it in items:
                        deltas.append((it, float(last_item.get(it, 0.0) - prev_item.get(it, 0.0))))
                    # Maiores altas e quedas
                    top_pos = [f"{it} (+R$ {v:,.2f})" for it, v in sorted([d for d in deltas if d[1] > 0], key=lambda x: x[1], reverse=True)[:2]]
                    top_neg = [f"{it} (-R$ {abs(v):,.2f})" for it, v in sorted([d for d in deltas if d[1] < 0], key=lambda x: x[1])[:2]]

                sugestoes = []
                if delta_abs < 0:
                    if top_neg:
                        sugestoes.append(f"Recuperar volumes em {', '.join(top_neg)}")
                    if top_pos:
                        sugestoes.append(f"Acelerar itens em alta: {', '.join(top_pos)}")
                    sugestoes.extend([
                        "Reativar pacientes inativos e evitar churn",
                        "Revisar disponibilidade logística/estoque para evitar rupturas",
                    ])
                else:
                    if top_pos:
                        sugestoes.append(f"Manter foco no item: {', '.join(top_pos)}")
                    sugestoes.extend([
                        "Prospectar novos contratos semelhantes aos de melhor desempenho",
                        "Aprimorar conversão de solicitações e tempo de instalação",
                    ])

                st.markdown("<hr>", unsafe_allow_html=True)
                st.markdown("<h4 style='margin:6px 0 6px 0;'>Análise de variação do faturamento</h4>", unsafe_allow_html=True)
                resumo = f"Entre {anterior} e {ultimo}, o faturamento {direcao} em R$ {abs(delta_abs):,.2f}"
                if delta_pct is not None:
                    resumo += f" (\u2248 {delta_pct:.1f}%)."
                st.markdown(resumo)
                if top_pos or top_neg:
                    if top_pos:
                        st.markdown(f"- Itens que mais puxaram para cima: {', '.join(top_pos)}")
                    if top_neg:
                        st.markdown(f"- Itens que mais puxaram para baixo: {', '.join(top_neg)}")
                if sugestoes:
                    st.markdown("**O que pode ser feito para subir mais:**")
                    for s in sugestoes:
                        st.markdown(f"- {s}")

                # Visualização em gráficos
                # Linha: duas principais camas que mais impactam no faturamento (série mensal)
                if 'df_rev_sum' in locals() and isinstance(df_rev_sum, pd.DataFrame) and not df_rev_sum.empty:
                    camas_principais = [
                        "CAMA ELÉTRICA 3 MOVIMENTOS",
                        "CAMA MANUAL 2 MANIVELAS",
                    ]
                    meses_ordem_full = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]
                    df_camas = (
                        df_rev_sum[df_rev_sum["ItemCanonical"].isin(camas_principais)]
                        .groupby(["Mês","ItemCanonical"], as_index=False)["Faturamento"].sum()
                    )
                    # Garante série completa e ordenada (evita "voltar" no fim)
                    grid = pd.MultiIndex.from_product([meses_ordem_full, camas_principais], names=["Mês","ItemCanonical"]).to_frame(index=False)
                    df_camas = grid.merge(df_camas, on=["Mês","ItemCanonical"], how="left").fillna({"Faturamento": 0})
                    df_camas["Mês"] = pd.Categorical(df_camas["Mês"], categories=meses_ordem_full, ordered=True)
                    df_camas = df_camas.sort_values(["ItemCanonical", "Mês"])  
                    fig_camas = px.line(
                        df_camas,
                        x="Mês",
                        y="Faturamento",
                        color="ItemCanonical",
                        markers=True,
                        category_orders={"Mês": meses_ordem_full},
                        title="Evolução mensal – Camas que mais impactam o faturamento",
                    )
                    fig_camas.update_traces(mode="lines+markers", line_shape="linear")
                    fig_camas.update_yaxes(tickprefix="R$ ", tickformat=",.2f")
                    show_plot(fig_camas, use_container_width=True)
            except Exception:
                pass

            # Rodapé
            st.markdown("---")
            st.markdown(
                """
                <div style='text-align: center; padding: 20px; color: #666; font-size: 14px;'>
                    <p><strong>Dashboard desenvolvido por Lucas Missiba</strong></p>
                    <p>Alocama · Setor de Contratos</p>
                </div>
                """,
                unsafe_allow_html=True
            )

            # ================================
            # Substituição: Waterfall – variação mensal do faturamento
            # ================================
            st.subheader("Variação mensal do faturamento – AXX CARE (Waterfall)")
            try:
                df_total_axx_hist = pd.DataFrame({
                    "Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"],
                    "Faturamento": [98579.58, 87831.11, 96184.47, 92286.01, 87803.67, 77499.87, 81856.05, 82609.95],
                })
                vals = df_total_axx_hist["Faturamento"].tolist()
                meses = df_total_axx_hist["Mês"].tolist()
                measures = ["absolute"] + ["relative"] * (len(vals) - 2) + ["total"]
                y = [vals[0]] + [vals[i] - vals[i-1] for i in range(1, len(vals)-1)] + [vals[-1]]
                labels = [f"{meses[0]} (base)"] + [f"Δ {m}" for m in meses[1:-1]] + [f"{meses[-1]} (final)"]
                fig_w = go.Figure(go.Waterfall(
                    x=labels,
                    measure=measures,
                    y=y,
                    connector=dict(line=dict(color="#374151")),
                    increasing=dict(marker=dict(color="#10b981")),
                    decreasing=dict(marker=dict(color="#ef4444")),
                    totals=dict(marker=dict(color="#3b82f6")),
                    textposition="outside",
                ))
                fig_w.update_layout(title="Como o faturamento evoluiu mês a mês")
                show_plot(fig_w, use_container_width=True)
            except Exception:
                pass

            # Downloads removidos a pedido

            # ================================
            # ARPU (Ticket médio) = Faturamento / Vidas
            # ================================
            try:
                st.subheader("ARPU – Faturamento por vida (AXX CARE)")
                # Reconta vidas por mês (B) para 2025-02..08
                month_labels = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]
                month_map_ym = {"2025-01":"Janeiro","2025-02":"Fevereiro","2025-03":"Março","2025-04":"Abril","2025-05":"Maio","2025-06":"Junho","2025-07":"Julho","2025-08":"Agosto"}
                month_sets_arpu = {m:set() for m in month_labels}
                for file in sel_files:
                    try:
                        book = safe_read_excel(file, sheet_name=None)
                    except Exception:
                        continue
                    ym = year_month_from_path(file)
                    if ym not in month_map_ym:
                        continue
                    mes_label = month_map_ym[ym]
                    for sheet_name, df_sheet in (book or {}).items():
                        if should_exclude_sheet(str(sheet_name)):
                            continue
                        if not isinstance(df_sheet, pd.DataFrame) or df_sheet.empty:
                            continue
                        # Tenta identificar a melhor coluna de nomes
                        series = None
                        name_col = None
                        try:
                            name_col = select_best_name_column(df_sheet)
                        except Exception:
                            name_col = None
                        if name_col:
                            try:
                                series = df_sheet[name_col]
                            except Exception:
                                series = None
                        if series is None and df_sheet.shape[1] >= 2:
                            cand = df_sheet.iloc[:, 1]
                            scheck = cand.dropna().astype(str).str.strip()
                            scheck = scheck[scheck != ""]
                            if not scheck.empty:
                                norm = scheck.map(normalize_text_for_match)
                                looks_like = norm.str.contains(r"[a-z]", regex=True, na=False) & norm.str.contains(r"\\s", regex=True, na=False) & (norm.str.len() >= 5)
                                if int(looks_like.sum()) >= 5:
                                    series = cand
                        if series is None:
                            continue
                        series = series.dropna().astype(str).str.strip()
                        series = series[series != ""]
                        if series.empty:
                            continue
                        nomes_norm = series.apply(normalize_text_for_match)
                        month_sets_arpu[mes_label].update(nomes_norm.tolist())
                df_vidas_arpu = pd.DataFrame({
                    "Mês": month_labels,
                    "Vidas": [len(month_sets_arpu[m]) for m in month_labels],
                })
                # Faturamento informado manualmente por mês (inclui Janeiro e Fevereiro)
                total_rev_map = {
                    "Janeiro": 98579.58,
                    "Fevereiro": 87831.11,
                    "Março": 96184.47,
                    "Abril": 92286.01,
                    "Maio": 87803.67,
                    "Junho": 77499.87,
                    "Julho": 81856.05,
                    "Agosto": 82609.95,
                }
                rev_df = pd.DataFrame({"Mês": month_labels, "Faturamento": [total_rev_map.get(m, 0.0) for m in month_labels]})
                df_arpu = rev_df.merge(df_vidas_arpu, on="Mês", how="left").fillna(0)
                df_arpu["ARPU"] = df_arpu.apply(lambda r: (float(r["Faturamento"]) / r["Vidas"]) if r["Vidas"] > 0 else 0.0, axis=1)
                fig_arpu = px.bar(
                    df_arpu, x="Mês", y="ARPU",
                    title="ARPU (Average Revenue Per User) - AXX CARE",
                    text="ARPU",
                    color="ARPU",
                    color_continuous_scale="Viridis",
                    category_orders={"Mês": ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"]},
                )
                fig_arpu.update_traces(
                    texttemplate="R$ %{text:,.2f}",
                    textposition="outside",
                    hovertemplate="<b>%{x}</b><br>ARPU: R$ %{y:,.2f}<br>Vidas: %{customdata[0]}<br>Faturamento: R$ %{customdata[1]:,.2f}<extra></extra>",
                    customdata=df_arpu[["Vidas", "Faturamento"]].values
                )
                fig_arpu.update_layout(
                    yaxis_title="ARPU (R$)",
                    xaxis_title="Mês",
                    margin=dict(l=20, r=20, t=60, b=40)
                )
                fig_arpu.update_yaxes(tickprefix="R$ ", tickformat=",.2f")
                show_plot(fig_arpu, use_container_width=True)
            except Exception:
                pass

            # Retenção removida a pedido

            # ================================
            # Heatmap Item × Mês (AXX CARE)
            # ================================
            try:
                st.subheader("Heatmap Item × Mês – AXX CARE")
                df_heat = (
                    df_rev[df_rev["Empresa"] == "AXX CARE"][["Mês","Item","Quantidade"]]
                    .groupby(["Item","Mês"], as_index=False)["Quantidade"].sum()
                )
                # Mantém somente top 20 itens por soma para foco visual
                tops = (
                    df_heat.groupby("Item")["Quantidade"].sum()
                    .sort_values(ascending=False).head(20).index.tolist()
                )
                df_heat = df_heat[df_heat["Item"].isin(tops)]
                df_pvt = df_heat.pivot(index="Item", columns="Mês", values="Quantidade").fillna(0)
                df_pvt = df_pvt[[m for m in ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto"] if m in df_pvt.columns]]
                fig_heat = px.imshow(
                    df_pvt,
                    color_continuous_scale="Blues",
                    aspect="auto",
                    labels=dict(color="Quantidade"),
                    title="Intensidade de ocorrências por item e mês",
                )
                show_plot(fig_heat, use_container_width=True)
            except Exception:
                pass

            # Pareto 80/20 removido a pedido

    # Rodapé
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; padding: 20px; color: #666; font-size: 14px;'>
            <p><strong>Dashboard desenvolvido por Lucas Missiba</strong></p>
            <p>Alocama · Setor de Contratos</p>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    # Fade removido - não precisa fechar div


if __name__ == "__main__":
    main()


