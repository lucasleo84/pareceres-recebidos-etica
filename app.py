import os
from datetime import datetime
from pathlib import Path
import pandas as pd
import streamlit as st

# ================== CONFIG ==================
st.set_page_config(page_title="Consulta de Pareceres Recebidos", layout="centered")

ARQ_XLSX = "pareceres_recebidos.xlsx"  # planilha de mapeamento
PASTA_PARECERES = "Pareceres"          # onde estão os arquivos dos pareceres

Path(PASTA_PARECERES).mkdir(exist_ok=True)

# ================== HELPERS ==================
def _normalize_str(s):
    return str(s).strip()

@st.cache_data(show_spinner=False)
def carregar_mapa(_sig: float) -> pd.DataFrame:
    """
    Lê o XLSX e retorna um DataFrame com colunas padronizadas:
    - aluno
    - arquivo  (pode ser nome de arquivo local ou URL)
    Regras:
      * Se houver cabeçalhos, tenta usar 'Aluno' e 'Arquivo' (ou similares).
      * Senão, assume: Coluna A = aluno, Coluna B = arquivo.
      * Permite múltiplos arquivos separados por vírgula/pipe/; em uma mesma célula.
    """
    df_raw = pd.read_excel(ARQ_XLSX, header=0)

    # tenta detectar cabeçalho; se estiver sem nomes significativos, recarrega sem header
    lower_cols = [str(c).strip().lower() for c in df_raw.columns]
    has_names = any(c in ("aluno", "alunos", "discente", "arquivo", "parecer", "arquivo do parecer") for c in lower_cols)

    if not has_names:
        # sem cabeçalho confiável -> usa posições (A=aluno, B=arquivo)
        df_raw = pd.read_excel(ARQ_XLSX, header=None)
        df_raw = df_raw.rename(columns={0: "aluno", 1: "arquivo"})
    else:
        # tenta mapear nomes comuns
        col_aluno = None
        for cand in ["aluno", "alunos", "discente", "nome", "nome do aluno"]:
            if cand in lower_cols:
                col_aluno = df_raw.columns[lower_cols.index(cand)]
                break
        col_arquivo = None
        for cand in ["arquivo", "parecer", "arquivo do parecer", "pdf", "coluna b"]:
            if cand in lower_cols:
                col_arquivo = df_raw.columns[lower_cols.index(cand)]
                break
        # fallback para posições (A/B) caso não ache
        if col_aluno is None or col_arquivo is None:
            df_raw = pd.read_excel(ARQ_XLSX, header=None)
            df_raw = df_raw.rename(columns={0: "aluno", 1: "arquivo"})
        else:
            df_raw = df_raw.rename(columns={col_aluno: "aluno", col_arquivo: "arquivo"})

    # padroniza tipos/strings
    for c in ["aluno", "arquivo"]:
        if c not in df_raw.columns:
            df_raw[c] = ""
        df_raw[c] = df_raw[c].astype(str).map(_normalize_str)

    # expande células com múltiplos arquivos (se houver)
    rows = []
    for _, r in df_raw.iterrows():
        aluno = r.get("aluno", "").strip()
        arquivo = r.get("arquivo", "").strip()
        if not aluno or not arquivo or arquivo.lower() == "nan":
            continue
        # separadores comuns
        partes = [p.strip() for p in pd.Series(str(arquivo)).str.split(r"[|,;]").iloc[0] if p and p.strip()]
        if not partes:
            partes = [arquivo]
        for p in partes:
            rows.append({"aluno": aluno, "arquivo": p})

    df = pd.DataFrame(rows) if rows else df_raw[["aluno", "arquivo"]].copy()
    # remove duplicatas triviais
    df = df.drop_duplicates().reset_index(drop=True)
    return df

def listar_arquivos_do_aluno(df: pd.DataFrame, aluno: str):
    """Retorna lista de dicts {tipo, label, path/url} para os arquivos do aluno."""
    sel = df[df["aluno"].str.casefold() == aluno.casefold()]
    resultados = []
    for _, r in sel.iterrows():
        item = r["arquivo"]
        # URL?
        if str(item).lower().startswith(("http://", "https://")):
            resultados.append({"tipo": "url", "label": item, "alvo": item})
        else:
            # caminho local: se vier só o nome, procurar na pasta PARECERES
            caminho = item
            if not os.path.isabs(caminho):
                caminho = os.path.join(PASTA_PARECERES, item)
            resultados.append({"tipo": "file", "label": os.path.basename(caminho), "alvo": caminho})
    return resultados

# ================== GUARDAS ==================
if not os.path.exists(ARQ_XLSX):
    st.error(f"Arquivo não encontrado: **{ARQ_XLSX}**. Coloque o Excel na raiz do app.")
    st.stop()

mtime = os.path.getmtime(ARQ_XLSX)
df_map = carregar_mapa(mtime)

# ================== UI ==================
st.title("📄 Pareceres Recebidos")
st.caption("Selecione seu nome para baixar os pareceres que chegaram para você.")

# lista de alunos
alunos = sorted(df_map["aluno"].dropna().unique(), key=lambda s: s.casefold())
aluno_sel = st.selectbox("Seu nome", alunos)

colA, colB = st.columns([1, 1])
with colA:
    if st.button("🔄 Recarregar planilha"):
        st.cache_data.clear()
        st.rerun()
with colB:
    st.write("")  # espaçador

if aluno_sel:
    arquivos = listar_arquivos_do_aluno(df_map, aluno_sel)
    if not arquivos:
        st.warning("Nenhum parecer encontrado para este nome. Verifique se o Excel foi atualizado.")
    else:
        st.subheader("Seus pareceres")
        for i, item in enumerate(arquivos, start=1):
            if item["tipo"] == "url":
                st.link_button(f"🔗 Abrir parecer #{i}", item["alvo"])
            else:
                caminho = item["alvo"]
                if os.path.exists(caminho):
                    with open(caminho, "rb") as f:
                        st.download_button(
                            label=f"⬇️ Baixar parecer #{i} — {item['label']}",
                            data=f.read(),
                            file_name=item["label"],
                            mime="application/octet-stream",
                            key=f"dl_{aluno_sel}_{i}"
                        )
                else:
                    st.error(f"Arquivo não encontrado: `{item['label']}`. Verifique a pasta **{PASTA_PARECERES}/** e a coluna B do Excel.")

st.divider()
with st.expander("Como preparar a planilha (rápido)"):
    st.markdown(
        """
        - O app aceita planilha **com** ou **sem** cabeçalho.
        - Sem cabeçalho: **Coluna A = Aluno**, **Coluna B = Arquivo/Link**.
        - Com cabeçalho: use nomes como **Aluno** e **Arquivo** (ou **Parecer**).
        - Na coluna B você pode colocar:
            - **nome do arquivo** que está dentro da pasta `Pareceres/` (ex.: `Joao_Silva.pdf`);
            - **URL** (Google Drive, OneDrive etc).
        - Se houver **vários pareceres** para o mesmo aluno, você pode:
            - repetir linhas para o mesmo aluno **ou**
            - separar na coluna B por **vírgula**, **ponto e vírgula** ou **pipe** (`|`).
        """
    )
