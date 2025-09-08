from pathlib import Path
from typing import List, Optional, Tuple, Dict
import re
import unicodedata
import hashlib
import time

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components


APP_TITLE = "Painel Gerencial do Setor de Contratos | Alô Cama"
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

# Candidatos para detectar coluna de pacientes/vidas
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


# ==============================
# Autenticação simples (sem lib)
# ==============================
def hash_password(raw: str) -> str:
    try:
        return hashlib.sha256(raw.encode("utf-8")).hexdigest()
    except Exception:
        return ""


def get_auth_users() -> Dict[str, str]:
    """Retorna {usuario: senha_sha256}. Configure em st.secrets['auth_users'].

    Exemplo do secrets.toml (Streamlit Cloud):
    [auth_users]
    admin = "sha256:..."
    lucas = "sha256:..."

    Use: python -c "import hashlib;print('sha256:'+hashlib.sha256('SUA_SENHA'.encode()).hexdigest())"
    """
    users: Dict[str, str] = {}
    try:
        # Pode ser dict normal, ou dict com prefixo 'sha256:'
        for k, v in (st.secrets.get("auth_users", {}) or {}).items():
            if isinstance(v, str) and v.startswith("sha256:"):
                users[k] = v.split(":", 1)[1]
            elif isinstance(v, str):
                users[k] = v  # assume já ser sha256
    except Exception:
        users = {}
    # Fallback seguro para desenvolvimento local
    if not users:
        users = {"admin": hash_password("admin")}
    return users


def verify_credentials(username: str, password: str) -> bool:
    users = get_auth_users()
    if not username or not password:
        return False
    stored = users.get(username)
    return stored == hash_password(password)


def render_splash_once() -> bool:
    """Mostra uma splash de carregamento elegante apenas na primeira visita."""
    if "splash_shown" not in st.session_state:
        st.session_state["splash_shown"] = False
    if st.session_state["splash_shown"]:
        return False
    # Overlay em toda a viewport (sem iframe) para garantir cor única
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
            <div class='brand'>Alô Cama · Setor de Contratos</div>
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
    """Página de login moderna/animada. Retorna True se logado."""
    if st.session_state.get("authed"):
        return True

    # Card de login único centralizado
    col_center = st.columns([1, 2, 1])[1]
    with col_center:
        # Removido pulse/animacão
        with st.container(border=True):
            # Logo: procura em assets/logo.(png|jpg|jpeg|svg) ou usa URL em secrets
            logo_path: Optional[str] = None
            try:
                for ext in ("png", "jpg", "jpeg", "svg"):
                    p = Path("assets") / f"logo.{ext}"
                    if p.exists():
                        logo_path = str(p)
                        break
                if not logo_path:
                    logo_path = st.secrets.get("logo_url", None) if hasattr(st, "secrets") else None
            except Exception:
                logo_path = None
            if logo_path:
                c1, c2, c3 = st.columns([1, 2, 1])
                with c2:
                    try:
                        st.image(logo_path, width=180)
                    except Exception:
                        pass
            st.markdown(
                """
                <div style="display:flex;flex-direction:column;gap:4px">
                  <div style="font-weight:700;font-size:20px;color:#87ceeb">Acesso ao Painel</div>
                  <div style="color:#b7c7d9;font-size:14px;margin-bottom:6px">Insira suas credenciais para continuar</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            user = st.text_input("Usuário", key="auth_user", placeholder="Seu usuário")
            pwd = st.text_input("Senha", type="password", key="auth_pwd", placeholder="Sua senha")
            submit = st.button("Entrar", type="primary", use_container_width=True)

            st.caption("Respeitamos a LGPD e tratamos dados pessoais com segurança e finalidade específica.")
        # fim do card

    if submit:
        if verify_credentials(user, pwd):
            st.session_state["authed"] = True
            st.session_state["auth_user_name"] = user
            st.success("Autenticação realizada. Redirecionando...")
            time.sleep(0.6)
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos")

    return False

def clean_item_values(series: pd.Series, selected_col_name: str, only_equipment: bool = False) -> pd.Series:
    """Normaliza e filtra valores não válidos da coluna de itens/produtos."""
    s = series.astype(str).str.strip()
    # Remove vazios
    s = s[s != ""]
    # Remove valores iguais ao cabeçalho/nome da coluna (e sinônimos comuns)
    invalid_names = {selected_col_name.lower(), "item", "produto", "produtos", "descrição", "descricao"}
    s = s[~s.str.lower().isin(invalid_names)]
    # Remove linhas de totais/subtotais
    s = s[~s.str.lower().str.match(r"^(total|subtotal)\b")] 
    # Normalização para filtros adicionais
    norm = s.map(normalize_text_for_match)
    # Remove indicadores que não representam equipamentos
    bad_regex = r"(?:valor|pagina|page|quant|\bqtd\b|status|retirada|paciente|periodo|serie|unidade|unidades|\bun\b|\brs\b)"
    s = s[~norm.str.contains(bad_regex, regex=True, na=False)]
    # Remove tokens muito curtos após normalização
    norm = s.map(normalize_text_for_match)
    s = s[norm.str.len() >= 3]
    if only_equipment:
        # Mantém apenas valores que contenham letras (desconsidera números aleatórios ou somente símbolos)
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

    # Ordem importa: categorias mais específicas primeiro
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


def infer_group_for_label(label: str, candidates: List[str]) -> str:
    """Inferência robusta do grupo a partir do caminho (aceita \ ou / e variações)."""
    parts = re.split(r"[\\/]+", str(label))
    parts_norm = [normalize_text_for_match(p) for p in parts]
    norm_candidates = {normalize_text_for_match(c): c for c in candidates}

    # Preferência: após 'GRUPO SOLAR'
    if "grupo solar" in parts_norm:
        idx = parts_norm.index("grupo solar")
        if idx + 1 < len(parts_norm):
            nxt_norm = parts_norm[idx + 1]
            if nxt_norm in norm_candidates:
                return norm_candidates[nxt_norm]
            return parts[idx + 1]

    # Checa qualquer parte que bata exatamente com candidatos
    for p_norm, p in zip(parts_norm, parts):
        if p_norm in norm_candidates:
            return norm_candidates[p_norm]

    # Heurísticas de substring
    if any("hospital" in p for p in parts_norm):
        return "HOSPITALAR"
    if any("dommus" in p or "domus" in p for p in parts_norm):
        return "DOMMUS"
    if any("solar" in p for p in parts_norm):
        return "SOLAR"

    return parts[0] if parts else ""


def primary_group_from_label(label: str) -> str:
    """Inferência robusta da empresa a partir do caminho relativo.

    Prioriza a detecção por substring ('dommus'/'domus', 'hospital', 'solar').
    Caso não bata, usa o primeiro segmento do caminho.
    """
    s_norm = normalize_text_for_match(str(label))
    if "dommus" in s_norm or "domus" in s_norm:
        return "DOMMUS"
    if "hospital" in s_norm:
        return "HOSPITALAR"
    if "solar" in s_norm:
        return "SOLAR"
    parts = re.split(r"[\\/]+", str(label).strip())
    return (parts[0].upper() if parts and parts[0] else "").upper()


def month_from_path(path: Path) -> Optional[str]:
    """Retorna '6', '7' ou '8' se o caminho contiver pastas 6/7/8 ou 2025-06/07/08."""
    parts = re.split(r"[\\/]+", str(path))
    for p in parts:
        p_norm = p.strip()
        if re.fullmatch(r"0?[678]", p_norm):
            return p_norm.lstrip("0")
        m = re.fullmatch(r"\d{4}-(0[678])", p_norm)
        if m:
            return m.group(1).lstrip("0")
    return None


def render_top3_pies(df_by_file: pd.DataFrame, group_names: Optional[List[str]] = None) -> None:
    """Renderiza gráficos de pizza (Top 3 itens) para cada grupo informado."""
    if df_by_file.empty:
        return
    df = df_by_file.copy()
    if "Grupo" not in df.columns:
        df["Grupo"] = df["Arquivo"].apply(lambda s: infer_group_for_label(str(s), group_names))
    # Se group_names não for informado, usa os grupos detectados nos dados
    if not group_names:
        group_names = sorted([g for g in df["Grupo"].unique().tolist() if str(g).strip() != ""])
    # Anexa categorias para possível filtro; se um grupo ficar vazio, cai no fallback sem filtro
    df_with_cat = attach_categories(df)

    st.subheader("Top 3 itens por grupo")
    cols = st.columns(min(3, len(group_names)) or 1)
    col_idx = 0
    for group in (group_names or []):
        df_g_raw = df[df["Grupo"] == group]
        df_g = df_with_cat[df_with_cat["Grupo"] == group]
        df_g = df_g[df_g["Categoria"] != "OUTROS"]
        if df_g.empty:
            # Fallback: usa dados crus (sem filtrar categoria)
            df_g = df_g_raw
        if df_g.empty:
            continue
        top3 = (
            df_g.groupby("Item", as_index=False)["Quantidade"].sum().sort_values("Quantidade", ascending=False).head(3)
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
        "base",
        "validacao",
        "valida",
        "mapeamento",
        "mapa",
        "cadastro",
        "dicion",
    ]
    return any(p in s for p in patterns)


def is_excel_file(path: Path) -> bool:
    """Retorna True se for um arquivo Excel válido (xlsx/xlsm, case-insensitive), exclui temporários e saída."""
    if not path.is_file():
        return False
    name_lower = path.name.lower()
    # Aceita .xlsx e .xlsm
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
    # Deduplica por caminho absoluto case-insensitive
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
            # Arquivo duplicado por conteúdo – ignora
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
    # Número 1-based
    if s.isdigit():
        pos = int(s) - 1
        if 0 <= pos < len(columns):
            return columns[pos]
    # Letra(s) do Excel
    pos_from_letter = excel_letter_to_index(s)
    if pos_from_letter is not None and 0 <= pos_from_letter < len(columns):
        return columns[pos_from_letter]
    # Nome da coluna (case-insensitive)
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

    # 1) Manual
    if selector:
        manual_col = resolve_column_selector(columns, selector)
        if manual_col is not None:
            cnt = non_empty_count(manual_col)
            if cnt > 0:
                return manual_col, "manual", cnt
            # se não houver dados, ainda podemos tentar smart

    if not use_smart:
        return (manual_col if selector else None), "none", 0

    # 2) Por nomes comuns
    for cand in SMART_FALLBACK_CANDIDATES:
        match = find_matching_column(columns, cand)
        if match is not None:
            cnt = non_empty_count(match)
            if cnt > 0:
                return match, "smart_name", cnt

    # 3) Melhor coluna textual
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
        # Normaliza e filtra valores com letras e preferência por nomes compostos
        norm = s.map(normalize_text_for_match)
        looks_like_name = norm.str.contains(r"[a-z]", regex=True, na=False)
        # favorece strings com espaço (nome e sobrenome)
        has_space = norm.str.contains(r"\s", regex=True, na=False)
        candidates = norm[looks_like_name & has_space & (norm.str.len() >= 5)]
        return int(candidates.nunique())

    # 1) Exatos a partir dos candidatos
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

    # 2) Por tokens no nome da coluna
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

    # 3) Heurística geral
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
            df_head = pd.read_excel(file, nrows=0)
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
        # Rótulo do arquivo como caminho relativo (sem extensão)
        try:
            rel = file.relative_to(base_dir)
            file_label = str(rel.with_suffix(""))
        except ValueError:
            file_label = file.stem
        try:
            book = pd.read_excel(file, sheet_name=None)
        except Exception:
            read_errors.append(file.name)
            continue

        found_in_this_file = False
        per_sheet_info: List[Tuple[str, str, str, int]] = []
        for sheet_name, df in (book or {}).items():
            # Ignora abas de resumo/gráficos/totais/etc.
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

            # Limpeza de valores para evitar contar cabeçalhos/total/vazios
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
        df_result.groupby(["Arquivo", "Item"], as_index=False)["Quantidade"].sum()
        .sort_values(["Arquivo", "Quantidade"], ascending=[True, False])
        .reset_index(drop=True)
    )
    # Reordena colunas para [Item, Quantidade, Arquivo] ao salvar, mas mantém aqui para gráfico
    return df_result, ignored_missing_col, read_errors, column_debug


def discover_unique_items(files: List[Path], target_column: str, use_smart: bool = True, only_equipment: bool = False) -> List[str]:
    """Descobre a lista única de valores da coluna alvo ao longo de todos os arquivos/abas."""
    unique_values = set()
    for file in files:
        try:
            book = pd.read_excel(file, sheet_name=None)
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


def save_to_excel(df_by_file: pd.DataFrame, df_totals: pd.DataFrame, path: Path, group_name: Optional[str] = None) -> None:
    # Salva em ordem solicitada: Item, Quantidade, Arquivo (aba principal)
    ordered = df_by_file[["Item", "Quantidade", "Arquivo"]].copy()
    totals = df_totals[["Item", "Quantidade"]].copy()
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        ordered.to_excel(writer, index=False, sheet_name="resultado")
        totals.to_excel(writer, index=False, sheet_name="totais_por_item")
        # Aba adicional consolidada (igual aos totais, opcionalmente com nome do grupo)
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
    st.plotly_chart(fig, use_container_width=True)


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
    st.plotly_chart(fig, use_container_width=True)


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    # Splash (primeiro acesso)
    if render_splash_once():
        return
    # Login obrigatório
    if not render_login():
        return
    # Cabeçalho compacto sem barra acima
    st.markdown(
        f"<div style='margin:0 0 8px 0; font-size:26px; font-weight:700;'>{APP_TITLE}</div>",
        unsafe_allow_html=True,
    )
    # Injeta CSS/JS global (remove âncoras e define animação de entrada)
    components.html(
        """
        <style>
        a[aria-label^="Anchor link"]{display:none!important}
        /* Oculta cabeçalho/menus nativos do Streamlit em todas as telas */
        #MainMenu, footer{visibility:hidden}
        [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"], .stDeployButton{display:none!important}
        header{visibility:hidden;height:0!important}
        .block-container{padding-top:0.25rem!important}
        .fade-in-on-scroll{opacity:0; transform: translateY(16px); transition: opacity .6s ease, transform .6s ease}
        .fade-in-on-scroll.is-visible{opacity:1; transform: translateY(0)}
        </style>
        <script>
        (function(){
          const init=()=>{
            const els=document.querySelectorAll('.fade-in-on-scroll');
            const io=new IntersectionObserver((entries)=>{
              entries.forEach(e=>{ if(e.isIntersecting){ e.target.classList.add('is-visible'); io.unobserve(e.target); } });
            },{threshold:0.15});
            els.forEach(el=>io.observe(el));
          };
          if(document.readyState==='complete' || document.readyState==='interactive') init();
          else window.addEventListener('DOMContentLoaded', init);
        })();
        </script>
        """,
        height=0,
    )

    # Seção de origem dos arquivos (sem cabeçalho para manter painel mais limpo)
    colp1, colp2 = st.columns([3, 1])
    with colp1:
        base_dir_str = st.text_input("Pasta base (caminho completo)", value=str(Path.cwd()))
    with colp2:
        recursive = st.checkbox("Incluir subpastas", value=True, help="Se marcado, lê também os .xlsx de subpastas.")

    base_dir = Path(base_dir_str).expanduser()
    if not base_dir.exists() or not base_dir.is_dir():
        st.error("Pasta base inválida. Verifique o caminho informado.")
        return

    # Dica removida para simplificar o painel

    def discover_groups(dir_base: Path) -> List[str]:
        grupos: List[str] = []
        for entry in dir_base.iterdir():
            if not entry.is_dir():
                continue
            # Verifica se há pelo menos um .xlsx (case-insensitive) em alguma subpasta
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
    # Filtro de grupo: escolha única (não permite múltiplos)
    if grupos_disponiveis:
        grupo_escolhido = st.radio(
            "Grupos (pastas principais)", options=grupos_disponiveis, horizontal=True, key="grupo_unico"
        )
        grupos_selecionados = [grupo_escolhido] if grupo_escolhido else []
    else:
        grupos_selecionados = []

    if grupos_selecionados:
        excel_files: List[Path] = []
        per_group_counts: Dict[str, int] = {}
        for g in grupos_selecionados:
            group_files = list_excel_files(base_dir / g, recursive=True)
            # Mantemos todos os arquivos do grupo/mesmo conteúdo — não deduplicar por conteúdo
            excel_files.extend(group_files)
            per_group_counts[g] = len(group_files)
    else:
        # Mantemos todos — não deduplicar por conteúdo
        excel_files = list_excel_files(base_dir, recursive=recursive)

    if not excel_files:
        st.warning("Nenhum arquivo .xlsx encontrado na pasta informada. Adicione arquivos e recarregue a página.")
        return

    with st.expander("Arquivos detectados", expanded=False):
        try:
            st.write([str(f.relative_to(base_dir)) for f in excel_files])
        except Exception:
            st.write([f.name for f in excel_files])
        # Diagnóstico: contagem por grupo
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
    run = st.button("Executar Análise", type="primary")

    if run:
        with st.spinner("Processando (coluna E) e contando itens por pasta 6/7/8..."):
            # Mantém somente arquivos de 6/7/8, aceitando 2025-06/07/08
            sel_files = [f for f in excel_files if month_from_path(f) in {"6", "7", "8"}]
            # Força uso da coluna E e considera somente equipamentos
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

        # Deriva a pasta (mês) a partir do caminho relativo (primeiro segmento após grupo)
        def extract_pasta(label: str) -> str:
            parts = re.split(r"[\\/]+", label)
            # Ex.: DOMMUS/2025-06/arquivo -> pasta = 6
            for p in parts:
                m = re.search(r"(0?[678])", p)
                if m:
                    return m.group(1).lstrip("0")
            return "?"

        df_result["Pasta"] = df_result["Arquivo"].apply(extract_pasta)
        # Remove itens artificiais 'DOMMUS' com quantidade 0
        df_result = df_result[~(
            df_result["Item"].astype(str).str.strip().str.upper() == "DOMMUS"
        ) | (df_result["Quantidade"] > 0)].reset_index(drop=True)

        # Totais por item (global) e ordenação
        df_totais = (
            df_result.groupby("Item", as_index=False)["Quantidade"].sum().sort_values("Quantidade", ascending=False)
        )
        item_order = df_totais["Item"].tolist()
        df_result["Item"] = pd.Categorical(df_result["Item"], categories=item_order, ordered=True)
        df_result_sorted = df_result.sort_values(["Item", "Pasta", "Quantidade"], ascending=[True, True, False])

        output_path = base_dir / OUTPUT_FILENAME
        try:
            # Salva como requisitado: Item, Quantidade, Pasta
            save_df = df_result_sorted[["Item", "Quantidade", "Pasta"]]

            # Preparação de consolidados por empresa
            df_emp = df_result_sorted.copy()
            df_emp["Empresa"] = df_emp["Arquivo"].apply(primary_group_from_label).str.upper()
            df_emp["Empresa"] = df_emp["Empresa"].replace({
                "GRUPO SOLAR": "SOLAR",
            })

            # Consolidação geral apenas do último mês disponível (6/7/8)
            months_numeric = pd.to_numeric(df_emp["Pasta"], errors="coerce")
            last_month_num = int(months_numeric.max()) if not months_numeric.dropna().empty else None
            last_month_str = str(last_month_num) if last_month_num is not None else None
            df_emp_last = df_emp[df_emp["Pasta"] == last_month_str] if last_month_str else df_emp.copy()
            df_consolidado_geral = (
                df_emp_last.groupby("Item", as_index=False)["Quantidade"].sum().sort_values("Quantidade", ascending=False)
            )

            # Normaliza nomes para ranking mais consistente (trio principal)
            df_emp["ItemCanon"] = df_emp["Item"].map(canonicalize_trio_item)

            # Top 3 por empresa considerando o mês de pico (não soma os meses)
            df_mes_empresa = (
                df_emp.groupby(["Empresa", "ItemCanon", "Pasta"], as_index=False)["Quantidade"].sum()
            )
            # Seleciona, para cada (Empresa, Item), o registro com maior Quantidade
            df_peak = (
                df_mes_empresa.sort_values(["Empresa", "ItemCanon", "Quantidade"], ascending=[True, True, False])
                .drop_duplicates(["Empresa", "ItemCanon"], keep="first")
            )
            # Ranking por empresa usando a quantidade de pico (posição inicia em 1)
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

        # Gráfico de barras comparativo (Top 10 itens entre Junho/Julho/Agosto)
        st.markdown('<h3 style="margin:0 0 8px 0;">Dashboard</h3>', unsafe_allow_html=True)
        month_map = {"6": "Junho", "7": "Julho", "8": "Agosto"}
        df_viz = df_result_sorted.copy()
        df_viz["Mês"] = df_viz["Pasta"].map(month_map).fillna(df_viz["Pasta"])
        month_order = [month_map[m] for m in ["6", "7", "8"]]

        top10_items = df_totais.head(10)["Item"].tolist()
        df_viz_top = df_viz[df_viz["Item"].isin(top10_items)]
        # Consolida somatório mês a mês SOMANDO todos os arquivos/CNPJs do grupo selecionado
        df_viz_top = (
            df_viz_top.groupby(["Item", "Mês"], as_index=False)["Quantidade"].sum()
        )
        # Garante exatamente Top 10 após a consolidação
        top10_after_agg = (
            df_viz_top.groupby("Item", as_index=False)["Quantidade"].sum()
            .sort_values("Quantidade", ascending=False)
            .head(10)["Item"].tolist()
        )
        df_viz_top = df_viz_top[df_viz_top["Item"].isin(top10_after_agg)]
        # Inserimos um wrapper com animação de fade-in ao entrar na viewport
        components.html(
            """
            <style>
            .fade-in-on-scroll{opacity:0; transform: translateY(16px); transition: opacity .6s ease, transform .6s ease}
            .fade-in-on-scroll.is-visible{opacity:1; transform: translateY(0)}
            </style>
            <script>
            (function(){
              const init=()=>{
                const els=document.querySelectorAll('.fade-in-on-scroll');
                const io=new IntersectionObserver((entries)=>{
                  entries.forEach(e=>{ if(e.isIntersecting){ e.target.classList.add('is-visible'); io.unobserve(e.target); } });
                },{threshold:0.15});
                els.forEach(el=>io.observe(el));
              };
              if(document.readyState==='complete' || document.readyState==='interactive') init();
              else window.addEventListener('DOMContentLoaded', init);
            })();
            </script>
            """,
            height=0,
        )

        st.markdown('<div class="fade-in-on-scroll" style="margin-top:0;">', unsafe_allow_html=True)
        fig = px.bar(
            df_viz_top,
            x="Item",
            y="Quantidade",
            color="Mês",
            barmode="group",
            category_orders={"Mês": month_order, "Item": top10_after_agg},
            title="Top 10 - Comparação de Itens entre Junho/Julho/Agosto",
            hover_data={"Mês": True, "Quantidade": ":,", "Item": True},
        )
        fig.update_traces(
            marker_line_color="#FFFFFF",
            marker_line_width=0.5,
            hovertemplate="Item: %{x}<br>Mês: %{customdata[0]}<br>Qtd: %{y:,}<extra></extra>",
        )
        fig.update_layout(
            xaxis_title="Itens (Junho / Julho / Agosto)",
            yaxis_title="Quantidade",
            legend_title_text="Meses",
            legend_orientation="h",
            legend_y=-0.2,
            showlegend=False,
            margin=dict(l=10, r=10, t=48, b=110),
        )
        fig.update_xaxes(tickangle=-60)
        st.plotly_chart(fig, width="stretch")
        st.markdown('</div>', unsafe_allow_html=True)

        # Consolidação geral somando as três empresas
        # Mostra consolidado apenas do último mês
        month_map_hdr = {"6": "Junho", "7": "Julho", "8": "Agosto"}
        last_month_hdr = df_emp_last["Pasta"].iloc[0] if not df_emp_last.empty else None
        last_month_label = month_map_hdr.get(str(last_month_hdr), str(last_month_hdr) if last_month_hdr else "-")
        st.subheader(f"Consolidado geral do último mês ({last_month_label})")
        df_consolidado_display = df_consolidado_geral.copy()
        df_consolidado_display.index = range(1, len(df_consolidado_display) + 1)
        df_consolidado_display.index.name = "Posição"
        st.dataframe(df_consolidado_display, use_container_width=True)
        # Mostra também o mês de pico por item (informação pedida)
        df_peak_item = (
            df_result.sort_values(["Item", "Quantidade"], ascending=[True, False])
            .drop_duplicates(["Item"], keep="first")
            .assign(Mês=lambda d: d["Pasta"].map({"6":"Junho","7":"Julho","8":"Agosto"}))
        )[["Item", "Quantidade", "Mês"]]
        with st.expander("Mês de pico por item (informativo)", expanded=False):
            df_peak_item_display = df_peak_item.copy()
            df_peak_item_display.index = range(1, len(df_peak_item_display) + 1)
            df_peak_item_display.index.name = "Posição"
            st.dataframe(df_peak_item_display, use_container_width=True)

        # Top 3 por empresa (mês de pico por item) – dinâmico conforme grupos filtrados
        st.subheader("Top 3 itens por empresa (Junho/Julho/Agosto)")
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
                    Mês=subset["Pasta"].map({"6":"Junho","7":"Julho","8":"Agosto"}),
                    Posição=(subset.groupby("Empresa")["Quantidade"].rank(ascending=False, method="first").astype(int))
                )[["Posição","ItemCanon","Quantidade","Mês"]]
                show.rename(columns={"ItemCanon": "Item"}, inplace=True)
                col.dataframe(show.set_index("Posição"), use_container_width=True)

        # Gráfico comparativo removido conforme solicitado

        # Evolução por empresa: mostra apenas AXX CARE quando for o único grupo; senão, abas dinâmicas
        df_emp_viz = df_result_sorted.copy()
        df_emp_viz["Empresa"] = df_emp_viz["Arquivo"].apply(primary_group_from_label).str.upper()
        df_emp_viz["Empresa"].replace({"GRUPO SOLAR": "SOLAR"}, inplace=True)
        df_emp_viz["Mês"] = df_emp_viz["Pasta"].map({"6": "Junho", "7": "Julho", "8": "Agosto"})
        month_order = ["Junho", "Julho", "Agosto"]

        empresas_presentes_viz = sorted(df_emp_viz["Empresa"].unique().tolist())
        if empresas_presentes_viz == ["AXX CARE"]:
            st.subheader("Evolução AXX CARE (Junho/Julho/Agosto)")
            df_e = df_emp_viz[df_emp_viz["Empresa"] == "AXX CARE"]
            if df_e.empty:
                st.info("Sem dados para AXX CARE nos meses 6/7/8")
            else:
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
                    title="Top 10 itens – AXX CARE",
                    hover_data={"Mês": True, "Quantidade": ":,", "Item": True},
                )
                fig_e_bar.update_layout(
                    xaxis_title="Itens (Junho / Julho / Agosto)",
                    yaxis_title="Quantidade",
                    margin=dict(l=20, r=20, t=60, b=150),
                    showlegend=False,
                )
                fig_e_bar.update_xaxes(tickangle=-60)
                st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                st.plotly_chart(fig_e_bar, width="stretch")
                st.markdown('</div>', unsafe_allow_html=True)

                df_e_tot = (
                    df_e.groupby("Mês", as_index=False)["Quantidade"].sum()
                    .set_index("Mês").reindex(month_order).reset_index().fillna(0)
                )
                fig_e_line = px.line(
                    df_e_tot,
                    x="Mês",
                    y="Quantidade",
                    markers=True,
                    title="Evolução mensal – AXX CARE",
                )
                fig_e_line.update_layout(yaxis_title="Quantidade", xaxis_title="Mês")
                st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                st.plotly_chart(fig_e_line, width="stretch")
                st.markdown('</div>', unsafe_allow_html=True)

                # Top 3 consolidado do último mês disponível
                # Determina último mês presente (6/7/8)
                meses_ordem = {"Junho": 6, "Julho": 7, "Agosto": 8}
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
                st.plotly_chart(fig_pie, width="stretch")
                st.markdown('</div>', unsafe_allow_html=True)
        elif empresas_presentes_viz == ["PRONEP"]:
            st.subheader("Evolução PRONEP (Junho/Julho/Agosto)")
            df_pn = df_emp_viz[df_emp_viz["Empresa"] == "PRONEP"]
            if df_pn.empty:
                st.info("Sem dados para PRONEP nos meses 6/7/8")
            else:
                top10_pn = (
                    df_pn.groupby("Item", as_index=False)["Quantidade"].sum()
                    .sort_values("Quantidade", ascending=False)["Item"].head(10).tolist()
                )
                df_pn_top = df_pn[df_pn["Item"].isin(top10_pn)]
                fig_pn_bar = px.bar(
                    df_pn_top,
                    x="Item",
                    y="Quantidade",
                    color="Mês",
                    category_orders={"Mês": month_order, "Item": top10_pn},
                    barmode="group",
                    title="Top 10 itens – PRONEP",
                    hover_data={"Mês": True, "Quantidade": ":,", "Item": True},
                )
                fig_pn_bar.update_layout(
                    xaxis_title="Itens (Junho / Julho / Agosto)",
                    yaxis_title="Quantidade",
                    margin=dict(l=20, r=20, t=60, b=150),
                    showlegend=False,
                )
                fig_pn_bar.update_xaxes(tickangle=-60)
                st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                st.plotly_chart(fig_pn_bar, width="stretch")
                st.markdown('</div>', unsafe_allow_html=True)

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
                fig_pn_line.update_layout(yaxis_title="Quantidade", xaxis_title="Mês")
                st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                st.plotly_chart(fig_pn_line, width="stretch")
                st.markdown('</div>', unsafe_allow_html=True)

                # Top 3 consolidado do último mês disponível – PRONEP
                meses_ordem = {"Junho": 6, "Julho": 7, "Agosto": 8}
                ultimo_mes = df_pn["Mês"].map(meses_ordem).max()
                mes_label = [k for k, v in meses_ordem.items() if v == ultimo_mes]
                mes_label = mes_label[0] if mes_label else "Junho"

                st.subheader(f"Participação dos Top 3 (quantidade) – PRONEP ({mes_label})")
                df_pn_last = df_pn[df_pn["Mês"] == mes_label].copy()
                # Canonicalização dos 3 itens principais da PRONEP
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
                st.plotly_chart(fig_pie_pn, width="stretch")
                st.markdown('</div>', unsafe_allow_html=True)
        else:
            st.subheader("Evolução por empresa (Junho/Julho/Agosto)")
            # Donut consolidado do Grupo Solar (aparece quando o filtro não inclui AXX CARE)
            if empresas_presentes_viz and all(e in {"HOSPITALAR", "SOLAR", "DOMMUS"} for e in empresas_presentes_viz):
                meses_ordem = {"Junho": 6, "Julho": 7, "Agosto": 8}
                ultimo_mes = df_emp_viz["Mês"].map(meses_ordem).max()
                mes_label = [k for k, v in meses_ordem.items() if v == ultimo_mes]
                mes_label = mes_label[0] if mes_label else "Junho"

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
                st.plotly_chart(fig_pie_gs, width="stretch")
                st.markdown('</div>', unsafe_allow_html=True)

            tabs = st.tabs(empresas_presentes_viz)
            for tab, empresa in zip(tabs, empresas_presentes_viz):
                with tab:
                    df_e = df_emp_viz[df_emp_viz["Empresa"] == empresa]
                    if df_e.empty:
                        st.info(f"Sem dados para {empresa} nos meses 6/7/8")
                    else:
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
                            margin=dict(l=20, r=20, t=60, b=150),
                            showlegend=False,
                        )
                        fig_e_bar.update_xaxes(tickangle=-60)
                        st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                        st.plotly_chart(fig_e_bar, width="stretch")
                        st.markdown('</div>', unsafe_allow_html=True)

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
                        fig_e_line.update_layout(yaxis_title="Quantidade", xaxis_title="Mês")
                        st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                        st.plotly_chart(fig_e_line, width="stretch")
                        st.markdown('</div>', unsafe_allow_html=True)

        # Faturamento: AXX CARE somente OU Grupo Solar somente
        empresas_presentes_fat = sorted(df_emp_viz["Empresa"].unique().tolist())
        if empresas_presentes_fat == ["AXX CARE"]:
            st.subheader("Faturamento AXX CARE – Top 3 Itens (Junho/Julho/Agosto)")
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
                dias_map = {"Junho": 30, "Julho": 31, "Agosto": 31}
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
                    category_orders={"Mês": ["Junho", "Julho", "Agosto"], "ItemCanonical": item_order},
                    title="Faturamento AXX CARE por Mês (diária x ocorrências)",
                    hover_data={"Faturamento": ":.2f", "Quantidade": True, "Dias": True},
                    labels={"ItemCanonical": "Item"},
                )
                fig_rev.update_layout(
                    yaxis_title="Faturamento (R$)", legend_title_text="Item",
                    legend_orientation="h", legend_y=-0.2, separators=".,",
                    margin=dict(l=20, r=20, t=60, b=80),
                )
                fig_rev.update_yaxes(tickprefix="R$ ", tickformat=".2f")
                st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                st.plotly_chart(fig_rev, width="stretch")
                st.markdown('</div>', unsafe_allow_html=True)
        elif all(e in {"HOSPITALAR", "SOLAR", "DOMMUS"} for e in empresas_presentes_fat):
            # Grupo Solar (quando filtro só tem empresas do grupo)
            st.subheader("Faturamento Grupo Solar – Top Itens (Junho/Julho/Agosto)")
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
                dias_map = {"Junho": 30, "Julho": 31, "Agosto": 31}
                df_gs_sum["Dias"] = df_gs_sum["Mês"].map(dias_map).fillna(30)
                df_gs_sum["Faturamento"] = df_gs_sum["Quantidade"] * df_gs_sum["PrecoDiaria"] * df_gs_sum["Dias"]
                item_order_gs = [
                    canonical_map_solar[normalize_text_for_match("CAMA ELÉTRICA 3 MOVIMENTOS")],
                    canonical_map_solar[normalize_text_for_match("SUPORTE DE SORO")],
                    canonical_map_solar[normalize_text_for_match("COLCHÃO PNEUMÁTICO")],
                ]
                # Agrega o faturamento do Grupo Solar (soma entre empresas) por mês e item
                df_gs_total = (
                    df_gs_sum.groupby(["Mês", "ItemCanonical"], as_index=False)["Faturamento"].sum()
                )
                fig_gs = px.bar(
                    df_gs_total,
                    x="Mês", y="Faturamento", color="ItemCanonical",
                    facet_col="ItemCanonical",
                    category_orders={"Mês": ["Junho", "Julho", "Agosto"], "ItemCanonical": item_order_gs},
                    title="Faturamento por Mês – Grupo Solar (diária x ocorrências)",
                    hover_data={"Faturamento": ":.2f"},
                    labels={"ItemCanonical": "Item"},
                )
                fig_gs.update_layout(yaxis_title="Faturamento (R$)", legend_title_text="Item",
                                     legend_orientation="h", legend_y=-0.2, separators=".,",
                                     margin=dict(l=20, r=20, t=60, b=80))
                fig_gs.update_yaxes(tickprefix="R$ ", tickformat=".2f")
                st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                st.plotly_chart(fig_gs, width="stretch")
                st.markdown('</div>', unsafe_allow_html=True)
                # Disponibiliza df_gs_sum como df_rev_sum para seções abaixo que usam esse nome
                df_rev_sum = df_gs_sum

            # Faturamento separado por empresa (apenas quando for Grupo Solar)
            st.subheader("Faturamento por empresa (Junho/Julho/Agosto)")
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
                            category_orders={"Mês": ["Junho", "Julho", "Agosto"], "ItemCanonical": item_order},
                            title=empresa,
                            hover_data={"Faturamento": ":.2f", "Quantidade": True},
                        )
                        fig_e.update_layout(
                            yaxis_title="Faturamento (R$)",
                            legend_title_text="Item",
                            legend_orientation="h",
                            legend_y=-0.2,
                        )
                        fig_e.update_yaxes(tickprefix="R$ ", tickformat=".2f")
                        st.markdown('<div class="fade-in-on-scroll">', unsafe_allow_html=True)
                        st.plotly_chart(fig_e, width="stretch")
                        st.markdown('</div>', unsafe_allow_html=True)

        # Faturamento PRONEP – quando apenas PRONEP está filtrada
        if empresas_presentes_fat == ["PRONEP"]:
            st.subheader("Faturamento PRONEP – Top Itens (Junho/Julho/Agosto)")
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
                st.info("Sem ocorrências dos itens tarifados para PRONEP nos meses 6/7/8.")
            else:
                df_pn["ItemCanonical"] = df_pn["key"].map(canonical_map_pronep)
                df_pn_sum = (
                    df_pn.groupby(["Empresa", "Mês", "ItemCanonical"], as_index=False)
                    .agg(Quantidade=("Quantidade", "sum"), PrecoDiaria=("PrecoDiaria", "first"))
                )
                dias_map = {"Junho": 30, "Julho": 31, "Agosto": 31}
                df_pn_sum["Dias"] = df_pn_sum["Mês"].map(dias_map).fillna(30)
                df_pn_sum["Faturamento"] = df_pn_sum["Quantidade"] * df_pn_sum["PrecoDiaria"] * df_pn_sum["Dias"]

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
                    category_orders={"Mês": ["Junho", "Julho", "Agosto"], "ItemCanonical": item_order_pn},
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
                st.plotly_chart(fig_pn_rev, width="stretch")
                st.markdown('</div>', unsafe_allow_html=True)

        # Vidas ativas no Home Care (Grupo Solar) – últimos 3 meses (Junho/Julho/Agosto)
        if empresas_presentes_viz and all(e in {"HOSPITALAR", "SOLAR", "DOMMUS"} for e in empresas_presentes_viz):
            st.subheader("Vidas ativas no Home Care – Grupo Solar (últimos 3 meses)")
            # Conjunto de nomes únicos por mês (Junho/Julho/Agosto)
            month_sets = {"Junho": set(), "Julho": set(), "Agosto": set()}
            for file in sel_files:
                try:
                    book = pd.read_excel(file, sheet_name=None)
                except Exception:
                    continue
                pasta_mes = month_from_path(file)
                mes_label = {"6": "Junho", "7": "Julho", "8": "Agosto"}.get(pasta_mes or "", None)
                if mes_label not in month_sets:
                    continue
                for sheet_name, df_sheet in (book or {}).items():
                    # Ignora abas de resumo/totais/gráficos
                    if should_exclude_sheet(str(sheet_name)):
                        continue
                    if not isinstance(df_sheet, pd.DataFrame) or df_sheet.empty:
                        continue
                    # 1) Força coluna B (segunda coluna). 2) Se não existir, fallback inteligente
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
                    # Normaliza levemente para evitar duplicatas por variação de acento/caixa
                    nomes_norm = series.apply(normalize_text_for_match)
                    month_sets[mes_label].update(nomes_norm.tolist())
            # Monta DataFrame com contagem por mês e média
            df_vidas_mes = pd.DataFrame({
                "Mês": ["Junho", "Julho", "Agosto"],
                "VidasUnicas": [len(month_sets["Junho"]), len(month_sets["Julho"]), len(month_sets["Agosto"])],
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
            st.plotly_chart(fig_vidas, width="stretch")
            # Animação do contador até a média
            target = int(round(media_vidas))
            placeholder = st.empty()
            for val in range(0, target + 1, max(1, target // 30)):
                placeholder.metric("Média de vidas ativas (3 meses)", f"{val}")
                time.sleep(0.02)
            if target % max(1, target // 30) != 0:
                placeholder.metric("Média de vidas ativas (3 meses)", f"{target}")
        elif empresas_presentes_viz == ["AXX CARE"]:
            st.subheader("Vidas ativas no Home Care – AXX CARE (últimos 3 meses)")
            # Conjunto de nomes únicos por mês (Junho/Julho/Agosto)
            month_sets = {"Junho": set(), "Julho": set(), "Agosto": set()}
            for file in sel_files:
                try:
                    book = pd.read_excel(file, sheet_name=None)
                except Exception:
                    continue
                pasta_mes = month_from_path(file)
                mes_label = {"6": "Junho", "7": "Julho", "8": "Agosto"}.get(pasta_mes or "", None)
                if mes_label not in month_sets:
                    continue
                for sheet_name, df_sheet in (book or {}).items():
                    if should_exclude_sheet(str(sheet_name)):
                        continue
                    if not isinstance(df_sheet, pd.DataFrame) or df_sheet.empty:
                        continue
                    # Coluna B como padrão; fallback inteligente
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
                "Mês": ["Junho", "Julho", "Agosto"],
                "VidasUnicas": [len(month_sets["Junho"]), len(month_sets["Julho"]), len(month_sets["Agosto"])],
            })
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
            st.plotly_chart(fig_vidas, width="stretch")
            target = int(round(media_vidas))
            placeholder = st.empty()
            for val in range(0, target + 1, max(1, target // 30)):
                placeholder.metric("Média de vidas ativas (3 meses)", f"{val}")
                time.sleep(0.02)
            if target % max(1, target // 30) != 0:
                placeholder.metric("Média de vidas ativas (3 meses)", f"{target}")
        elif empresas_presentes_viz == ["PRONEP"]:
            st.subheader("Vidas ativas no Home Care – PRONEP (últimos 3 meses)")
            month_sets = {"Junho": set(), "Julho": set(), "Agosto": set()}
            for file in sel_files:
                try:
                    book = pd.read_excel(file, sheet_name=None)
                except Exception:
                    continue
                pasta_mes = month_from_path(file)
                mes_label = {"6": "Junho", "7": "Julho", "8": "Agosto"}.get(pasta_mes or "", None)
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
                "Mês": ["Junho", "Julho", "Agosto"],
                "VidasUnicas": [len(month_sets["Junho"]), len(month_sets["Julho"]), len(month_sets["Agosto"])],
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
            st.plotly_chart(fig_vidas, width="stretch")
            target = int(round(media_vidas))
            placeholder = st.empty()
            for val in range(0, target + 1, max(1, target // 30)):
                placeholder.metric("Média de vidas ativas (3 meses)", f"{val}")
                time.sleep(0.02)
            if target % max(1, target // 30) != 0:
                placeholder.metric("Média de vidas ativas (3 meses)", f"{target}")

        # Botão de download removido conforme solicitado


if __name__ == "__main__":
    main()


