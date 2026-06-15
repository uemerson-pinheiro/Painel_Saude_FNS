"""
Painel APS – Financiamento da Saúde
Município: São Gabriel do Oeste / MS (IBGE 500769)
Fontes: e-Gestor APS + FNS Fundo a Fundo
"""
import base64
import json
import time
from datetime import datetime
from pathlib import Path
from urllib.request import Request, urlopen

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

try:
    import xlrd
except ImportError:
    xlrd = None

# ──────────────────────────────────────────────────────────────
# CONFIGURAÇÃO DA PÁGINA
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Painel APS – São Gabriel do Oeste",
    page_icon="⚕️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────────────────────
# PALETA – Manual de Identidade Visual Meu SUS Digital
# ──────────────────────────────────────────────────────────────
AZUL          = "#183eff"
AZUL_ESCURO   = "#0a2177"
LARANJA       = "#ec641c"
LARANJA_CLARO = "#ff7a00"
VERMELHO      = "#e52722"
AMARELO       = "#ffcf00"
VERDE         = "#00cf00"
VERDE_ESCURO  = "#00a000"
CINZA_BG      = "#f2f5fb"

PALETA = [AZUL, LARANJA, VERDE_ESCURO, VERMELHO, "#ffb300", LARANJA_CLARO, "#7b96ff", "#8b5cf6"]

CORES_QUALIDADE = {
    "ÓTIMO":       VERDE_ESCURO,
    "BOM":         AZUL,
    "SUFICIENTE":  "#ffb300",
    "REGULAR":     LARANJA,
    "RUIM":        VERMELHO,
}

# ──────────────────────────────────────────────────────────────
# CSS
# ──────────────────────────────────────────────────────────────
st.markdown(f"""
<style>
  .stApp {{ background-color:{CINZA_BG}; }}

  /* Sidebar */
  [data-testid="stSidebar"]>div:first-child {{
    background:linear-gradient(175deg,{AZUL_ESCURO} 0%,{AZUL} 100%);
    padding-top:.5rem;
  }}
  [data-testid="stSidebar"] label,
  [data-testid="stSidebar"] p,
  [data-testid="stSidebar"] span {{color:rgba(255,255,255,.85)!important; font-size:.84rem!important;}}
  [data-testid="stSidebar"] h3 {{color:white!important;}}
  [data-testid="stSidebar"] .stSelectbox>div>div {{background:rgba(255,255,255,.12)!important; border-color:rgba(255,255,255,.25)!important;}}

  /* Banner topo */
  .header-banner {{
    background:linear-gradient(120deg,{AZUL_ESCURO} 0%,{AZUL} 65%,{LARANJA} 100%);
    color:white; padding:1.3rem 2rem; border-radius:14px;
    margin-bottom:1.4rem; display:flex; align-items:center; gap:1.2rem;
  }}
  .hb-icon {{font-size:2.8rem; line-height:1;}}
  .hb-titulo {{font-size:1.55rem; font-weight:800; margin:0; line-height:1.15;}}
  .hb-sub {{font-size:.88rem; opacity:.82; margin:.2rem 0 0 0;}}

  /* Card KPI */
  .kpi-box {{
    background:white; border-radius:12px; padding:1rem 1.2rem;
    border-top:4px solid {AZUL};
    box-shadow:0 2px 10px rgba(24,62,255,.09);
  }}
  .kpi-box.laranja  {{border-top-color:{LARANJA};}}
  .kpi-box.verde    {{border-top-color:{VERDE_ESCURO};}}
  .kpi-box.vermelho {{border-top-color:{VERMELHO};}}
  .kpi-box.amarelo  {{border-top-color:{AMARELO};}}
  .kpi-lbl {{font-size:.68rem;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.6px;margin-bottom:.2rem;}}
  .kpi-val {{font-size:1.6rem;font-weight:800;color:{AZUL_ESCURO};line-height:1.1;}}
  .kpi-delta {{font-size:.76rem;color:#999;margin-top:.2rem;}}

  /* Seção */
  .sec-title {{
    font-size:.98rem; font-weight:700; color:{AZUL_ESCURO};
    padding-bottom:.35rem; border-bottom:2.5px solid {AZUL};
    margin:.8rem 0 .8rem 0;
  }}

  /* Tabs */
  .stTabs [data-baseweb="tab-list"] {{
    gap:3px; background:white; border-radius:10px;
    padding:4px; box-shadow:0 1px 5px rgba(0,0,0,.07);
  }}
  .stTabs [data-baseweb="tab"] {{
    border-radius:7px; padding:.38rem 1rem;
    font-size:.86rem; font-weight:600; color:#555;
  }}
  .stTabs [aria-selected="true"] {{background:{AZUL}!important; color:white!important;}}

  #MainMenu {{visibility:hidden;}} footer {{visibility:hidden;}}
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────
# UTILITÁRIOS
# ──────────────────────────────────────────────────────────────
MESES = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]

def fmt_brl(v):
    try:
        v = float(v)
        s = f"{abs(v):,.2f}".replace(",","X").replace(".",",").replace("X",".")
        return f"R$ {s}" if v >= 0 else f"−R$ {s}"
    except Exception:
        return "—"

def parse_brl(s):
    try:
        if isinstance(s, (int, float)):
            return float(s)
        return float(str(s).replace(".","").replace(",","."))
    except Exception:
        return 0.0

def mes_label(parcela):
    p = str(int(parcela))
    if len(p) == 6:
        return f"{MESES[int(p[4:6])-1]}/{p[:4]}"
    return p

def delta_html(novo, antigo):
    if antigo and antigo != 0:
        d = (novo - antigo) / abs(antigo) * 100
        sinal, cor = ("▲", "#27ae60") if d >= 0 else ("▼", "#e74c3c")
        return f'<span style="color:{cor};font-weight:600">{sinal} {abs(d):.1f}%</span> vs mês ant.'
    return ""

def kpi_card(label, valor, delta="", cor=""):
    cls = f"kpi-box {cor}" if cor else "kpi-box"
    d = f'<div class="kpi-delta">{delta}</div>' if delta else ""
    return f"""<div class="{cls}">
  <div class="kpi-lbl">{label}</div>
  <div class="kpi-val">{valor}</div>
  {d}
</div>"""

def estilo(fig, titulo=None, h=360):
    fig.update_layout(
        height=h, plot_bgcolor="white", paper_bgcolor="white",
        font=dict(family="Arial, sans-serif", size=11.5, color="#444"),
        title=dict(text=f"<b>{titulo}</b>" if titulo else "", font=dict(size=13.5, color=AZUL_ESCURO), x=0),
        margin=dict(l=10, r=10, t=44 if titulo else 10, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.01, xanchor="right", x=1, font=dict(size=9.5)),
        xaxis=dict(showgrid=False, showline=True, linecolor="#ddd", tickfont=dict(size=10)),
        yaxis=dict(showgrid=True, gridcolor="#f0f0f0", showline=False, tickfont=dict(size=10)),
        colorway=PALETA,
    )
    return fig

# ──────────────────────────────────────────────────────────────
# DADOS COM CACHE
# ──────────────────────────────────────────────────────────────
_H = {"User-Agent": "Mozilla/5.0 (compatible; PainelAPS/1.0)", "Accept": "application/json"}

def _get(url, n=3):
    for i in range(n):
        try:
            with urlopen(Request(url, headers=_H), timeout=30) as r:
                return json.loads(r.read())
        except Exception as e:
            if i == n - 1:
                raise RuntimeError(f"Falha: {url} → {e}")
            time.sleep(2)

def _post_bytes(url, payload):
    body = json.dumps(payload).encode()
    hdrs = {**_H, "Content-Type": "application/json", "Accept": "*/*"}
    with urlopen(Request(url, data=body, headers=hdrs, method="POST"), timeout=60) as r:
        return r.read()

@st.cache_data(ttl=3600, show_spinner=False)
def load_egestor(co_uf: str, co_mun: str, p_ini: str, p_fim: str):
    url = (
        "https://relatorioaps-prd.saude.gov.br/financiamento/pagamento"
        f"?unidadeGeografica=MUNICIPIO&coUf={co_uf}&coMunicipio={co_mun}"
        f"&nuParcelaInicio={p_ini}&nuParcelaFim={p_fim}&tipoRelatorio=COMPLETO"
    )
    d = _get(url)
    return pd.DataFrame(d.get("resumosPlanosOrcamentarios", [])), pd.DataFrame(d.get("pagamentos", []))

@st.cache_data(ttl=3600, show_spinner=False)
def load_entidade(ano: int, uf: str, co_mun: str):
    url = (
        "https://consultafns.saude.gov.br/recursos/consulta-detalhada/entidades"
        f"?ano={ano}&count=10&estado={uf}&mes=1&municipio={co_mun}&page=1&tipoConsulta=2"
    )
    d = _get(url)
    lst = d.get("resultado", {}).get("dados", [])
    return lst[0] if lst else {}

@st.cache_data(ttl=3600, show_spinner=False)
def load_fns_acoes(ano: int, uf: str, co_mun: str, cnpj: str):
    rows, p = [], 1
    while True:
        url = (
            "https://consultafns.saude.gov.br/recursos/consulta-detalhada/detalhe-acao"
            f"?ano={ano}&count=50&cpfCnpjUg={cnpj}&estado={uf}&municipio={co_mun}&page={p}&tipoConsulta=2"
        )
        d = _get(url)
        res = d.get("resultado", {})
        for r in res.get("dados", []):
            rows.append({
                "Ação":          r.get("descricao", ""),
                "Forma Repasse": r.get("formaRepasse", ""),
                "Componente":    (r.get("componenteBloco") or {}).get("nome", ""),
                "Bloco":         (r.get("blocoPacto") or {}).get("nome", ""),
                "Grupo":         (r.get("grupoAcao") or {}).get("nome", ""),
                "Valor Total":   float(r.get("valorTotal") or 0),
                "Desconto":      float(r.get("valorDescontoTotal") or 0),
                "Valor Líquido": float(r.get("valorLiquido") or 0),
            })
        if p >= res.get("totalPaginas", 1):
            break
        p += 1
        time.sleep(0.3)
    return pd.DataFrame(rows)

@st.cache_data(ttl=3600, show_spinner=False)
def load_obs(meses: tuple, co_mun_ibge: str, nome_mun: str,
             razao_social: str, cnpj_fmt: str, cpf_cnpj: str, uf: str):
    if xlrd is None:
        return pd.DataFrame()
    url = "https://consultafns.saude.gov.br/recursos/consulta-detalhada/planilha-detalhada/"
    frames = []
    for ano, mes in meses:
        try:
            payload = {
                "ano": str(ano), "coAcao": "", "coBloco": "", "coComponente": "",
                "coGrupoAcao": "", "coMesAno": mes, "coMunicipioIbge": co_mun_ibge,
                "coPlanoOrcamentario": "", "dtFinalOb": "", "dtInicioOb": "",
                "formaRepasse": "", "noMunicipio": nome_mun, "noRazaoSocial": razao_social,
                "nuCnpj": cnpj_fmt, "nuCpfCnpjUg": cpf_cnpj,
                "sgUf": uf, "tipoConsulta": 2,
            }
            wb = xlrd.open_workbook(file_contents=_post_bytes(url, payload))
            sh = wb.sheet_by_index(0)
            hr = next((r for r in range(sh.nrows)
                       if "Bloco" in sh.row_values(r) and "Ação Detalhada" in sh.row_values(r)), None)
            if hr is None:
                continue
            cols = {i: str(c).strip() for i, c in enumerate(sh.row_values(hr)) if str(c).strip()}
            # índice da coluna "Nº OB" para validar linhas reais de lançamento
            ob_idx = next((i for i, n in cols.items() if "ob" in n.lower() and "nº" in n.lower()), None)
            for r in range(hr + 1, sh.nrows):
                row = sh.row_values(r)
                if not any(str(v).strip() for v in row):
                    continue
                # descarta qualquer linha onde ALGUMA célula contenha "total" (subtotais/totais gerais)
                if any(str(v).strip().lower().startswith("total") for v in row):
                    continue
                # descarta linhas sem Nº OB (são agregações, não lançamentos individuais)
                if ob_idx is not None and not str(row[ob_idx]).strip():
                    continue
                ln = {"Ano": ano, "Mês": mes}
                for idx, nome in cols.items():
                    ln[nome] = row[idx]
                frames.append(ln)
            time.sleep(0.5)
        except Exception:
            pass
    return pd.DataFrame(frames) if frames else pd.DataFrame()

# ──────────────────────────────────────────────────────────────
# SIDEBAR – FILTROS
# ──────────────────────────────────────────────────────────────
_brasao_path = Path(__file__).parent / "brasao_sgo.jpeg"
_brasao_b64 = base64.b64encode(_brasao_path.read_bytes()).decode() if _brasao_path.exists() else ""

with st.sidebar:
    brasao_html = (
        f'<img src="data:image/jpeg;base64,{_brasao_b64}" '
        f'style="width:110px;border-radius:8px;margin-bottom:.6rem;">'
        if _brasao_b64 else ""
    )
    st.markdown(f"""
    <div style="text-align:center;padding:.4rem 0 1.4rem 0;">
      {brasao_html}
      <div style="color:rgba(255,255,255,.7);font-size:.72rem;margin-top:.4rem;
                  letter-spacing:.6px;font-weight:600;text-transform:uppercase;">
        Ministério da Saúde
      </div>
    </div>
    <hr style="border-color:rgba(255,255,255,.2);margin:0 0 .8rem 0;">
    """, unsafe_allow_html=True)

    st.markdown("### ⚙️ Filtros")
    ano      = st.selectbox("Ano", [2026, 2025, 2024], index=0)
    mes_ini  = st.selectbox("Mês inicial", range(1, 13), format_func=lambda m: MESES[m-1], index=0)
    mes_fim  = st.selectbox("Mês final",   range(1, 13), format_func=lambda m: MESES[m-1],
                            index=min(datetime.now().month, 12) - 1 if ano == datetime.now().year else 11)

    if mes_fim < mes_ini:
        st.warning("Mês final deve ser ≥ mês inicial.")
        st.stop()

    st.markdown("---")
    st.markdown(f"""
    <div style="color:rgba(255,255,255,.65);font-size:.72rem;line-height:1.8;">
      📍 <b style="color:white;">São Gabriel do Oeste</b><br>
      MS &nbsp;|&nbsp; IBGE 500769<br>
      CNPJ 13.659.627/0001-09<br>
      Fundo Municipal de Saúde
    </div>
    <div style="color:rgba(255,255,255,.4);font-size:.65rem;margin-top:.8rem;">
      Fonte: e-Gestor APS + FNS<br>
      Cache renovado a cada 1h
    </div>
    """, unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────
# PARÂMETROS FIXOS
# ──────────────────────────────────────────────────────────────
UF           = "MS"
CO_UF        = "50"
CO_MUNICIPIO = "500769"
P_INI        = f"{ano}{mes_ini:02d}"
P_FIM        = f"{ano}{mes_fim:02d}"
MESES_SEL    = tuple((ano, m) for m in range(mes_ini, mes_fim + 1))

# ──────────────────────────────────────────────────────────────
# CABEÇALHO
# ──────────────────────────────────────────────────────────────
periodo_str = (f"{MESES[mes_ini-1]}/{ano}" if mes_ini == mes_fim
               else f"{MESES[mes_ini-1]} → {MESES[mes_fim-1]}/{ano}")
st.markdown(f"""
<div class="header-banner">
  <div class="hb-icon">⚕️</div>
  <div>
    <p class="hb-titulo">Painel APS – Financiamento da Saúde</p>
    <p class="hb-sub">São Gabriel do Oeste · MS · IBGE 500769 &nbsp;|&nbsp; {periodo_str}</p>
  </div>
</div>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────────────────────
# CARREGAMENTO
# ──────────────────────────────────────────────────────────────
with st.spinner("🔄 Buscando dados nas APIs do Ministério da Saúde..."):
    try:
        df_resumo, df_pag = load_egestor(CO_UF, CO_MUNICIPIO, P_INI, P_FIM)
        entidade = load_entidade(ano, UF, CO_MUNICIPIO)
        cnpj = entidade.get("cpfCnpj", "13659627000109")
        df_acoes = load_fns_acoes(ano, UF, CO_MUNICIPIO, cnpj)
        df_obs   = load_obs(
            MESES_SEL,
            entidade.get("codigoMunicipioIBGE", CO_MUNICIPIO),
            entidade.get("municipio", "SAO GABRIEL DO OESTE"),
            entidade.get("razaoSocial", ""),
            entidade.get("cpfCnpjFormatado", ""),
            cnpj, UF,
        )
        dados_ok = True
    except Exception as exc:
        st.error(f"❌ Erro ao carregar dados: {exc}")
        dados_ok = False

if not dados_ok:
    st.stop()

# ──────────────────────────────────────────────────────────────
# ABAS
# ──────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Visão Geral",
    "🏥 Detalhamento APS",
    "💰 Recursos FNS",
    "📋 Repasses / OBs",
])

# ══════════════════════════════════════════════
# ABA 1 – VISÃO GERAL
# ══════════════════════════════════════════════
with tab1:
    if df_resumo.empty:
        st.info("Sem dados de resumo para o período.")
        st.stop()

    df_r = df_resumo.copy()
    for c in ["vlEfetivoRepasse","vlDesconto","vlIntegral","vlAjuste"]:
        if c in df_r.columns:
            df_r[c] = pd.to_numeric(df_r[c], errors="coerce").fillna(0)

    df_mensal = (df_r.groupby("nuParcela")
                 .agg(repasse=("vlEfetivoRepasse","sum"), desconto=("vlDesconto","sum"))
                 .reset_index()
                 .sort_values("nuParcela"))
    df_mensal["label"] = df_mensal["nuParcela"].apply(mes_label)

    total_ano  = df_mensal["repasse"].sum()
    total_desc = df_mensal["desconto"].sum()
    ult        = df_mensal["repasse"].iloc[-1]
    pen        = df_mensal["repasse"].iloc[-2] if len(df_mensal) > 1 else None

    # KPIs
    st.markdown('<div class="sec-title">Indicadores do Período</div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(kpi_card("Repasse Último Mês", fmt_brl(ult), delta_html(ult, pen)), unsafe_allow_html=True)
    with c2:
        st.markdown(kpi_card("Acumulado no Período", fmt_brl(total_ano), f"{len(df_mensal)} mês(es)", "laranja"), unsafe_allow_html=True)
    with c3:
        pct_desc = f"{total_desc/total_ano*100:.1f}% do bruto" if total_ano else ""
        st.markdown(kpi_card("Total de Descontos", fmt_brl(total_desc), pct_desc, "vermelho"), unsafe_allow_html=True)
    with c4:
        n_prog = df_r["dsPlanoOrcamentario"].nunique() if "dsPlanoOrcamentario" in df_r.columns else "—"
        st.markdown(kpi_card("Programas com Repasse", str(n_prog), "planos orçamentários", "verde"), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Gráfico de barras + linha de desconto
    ca, cb = st.columns([3, 2])
    with ca:
        fig = go.Figure()
        fig.add_trace(go.Bar(x=df_mensal["label"], y=df_mensal["repasse"],
                             name="Repasse Efetivo", marker_color=AZUL, opacity=.88))
        if total_desc > 0:
            fig.add_trace(go.Scatter(x=df_mensal["label"], y=df_mensal["desconto"],
                                     name="Desconto", mode="lines+markers",
                                     line=dict(color=VERMELHO, width=2.5), marker=dict(size=7)))
        estilo(fig, "Evolução Mensal dos Repasses (R$)", h=360)
        fig.update_yaxes(tickprefix="R$ ", tickformat=",.0f")
        st.plotly_chart(fig, use_container_width=True)

    with cb:
        if "dsPlanoOrcamentario" in df_r.columns:
            df_pl = (df_r.groupby("dsPlanoOrcamentario")["vlEfetivoRepasse"]
                     .sum().reset_index()
                     .sort_values("vlEfetivoRepasse", ascending=False))
            fig2 = px.pie(df_pl, names="dsPlanoOrcamentario", values="vlEfetivoRepasse",
                          color_discrete_sequence=PALETA, hole=.42)
            fig2.update_traces(textposition="outside", textinfo="percent+label", textfont_size=8.5)
            estilo(fig2, "Composição por Programa", h=360)
            st.plotly_chart(fig2, use_container_width=True)

    # Tabela resumo mensal
    st.markdown('<div class="sec-title">Resumo por Mês</div>', unsafe_allow_html=True)
    df_tab = df_mensal[["label","repasse","desconto"]].copy()
    df_tab.columns = ["Mês/Ano","Repasse Efetivo","Desconto"]
    df_tab["% Desconto"] = (df_tab["Desconto"] / df_tab["Repasse Efetivo"] * 100).map("{:.2f}%".format)
    df_tab["Repasse Efetivo"] = df_tab["Repasse Efetivo"].apply(fmt_brl)
    df_tab["Desconto"]        = df_tab["Desconto"].apply(fmt_brl)
    st.dataframe(df_tab, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# ABA 2 – DETALHAMENTO APS
# ══════════════════════════════════════════════
with tab2:
    if df_pag.empty:
        st.info("Sem dados detalhados de pagamentos APS para o período.")
    else:
        df_p = df_pag.copy()

        # Converter numéricos
        cols_texto = {"nuParcela","sgUf","noMunicipio","dsEsferaAdministrativa",
                      "dsClassificacaoQualidadeEsfEap","dsClassificacaoVinculoEsfEap",
                      "dsClassificacaoQualidadeEmulti"}
        for c in df_p.columns:
            if c not in cols_texto:
                df_p[c] = pd.to_numeric(df_p[c], errors="coerce").fillna(0)

        if "nuParcela" in df_p.columns:
            df_p["label"] = df_p["nuParcela"].apply(mes_label)

        # ── Gráfico: vlEfetivoRepasse por Programa/Mês (fonte: resumosPlanosOrcamentarios)
        st.markdown('<div class="sec-title">Repasse Efetivo por Programa (por Mês)</div>', unsafe_allow_html=True)
        df_r2 = df_resumo.copy()
        df_r2["vlEfetivoRepasse"] = pd.to_numeric(df_r2["vlEfetivoRepasse"], errors="coerce").fillna(0)
        df_r2["label"] = df_r2["nuParcela"].apply(mes_label)

        if "dsPlanoOrcamentario" in df_r2.columns:
            df_prog_mes = (df_r2.groupby(["label", "dsPlanoOrcamentario"])["vlEfetivoRepasse"]
                           .sum().reset_index()
                           .sort_values("label"))
            fig3 = px.bar(
                df_prog_mes, x="label", y="vlEfetivoRepasse",
                color="dsPlanoOrcamentario",
                color_discrete_sequence=PALETA,
                barmode="stack",
                labels={"vlEfetivoRepasse": "Repasse Efetivo (R$)",
                        "label": "Mês/Ano",
                        "dsPlanoOrcamentario": "Programa"},
            )
            estilo(fig3, "Repasse Efetivo por Programa (por Mês)", h=400)
            fig3.update_yaxes(tickprefix="R$ ", tickformat=",.0f")
            st.plotly_chart(fig3, use_container_width=True)

        # ── Classificações de qualidade
        qa = "dsClassificacaoQualidadeEsfEap"
        va = "dsClassificacaoVinculoEsfEap"
        qa_ok = qa in df_p.columns and df_p[qa].notna().any()
        va_ok = va in df_p.columns and df_p[va].notna().any()

        if (qa_ok or va_ok) and "label" in df_p.columns:
            st.markdown('<div class="sec-title">Classificação de Qualidade e Vínculo – ESF/EAP</div>', unsafe_allow_html=True)
            cc1, cc2 = st.columns(2)

            def _qual_chart(df_col, titulo):
                fig = go.Figure()
                for _, row in df_p[["label", df_col]].dropna().iterrows():
                    q = str(row[df_col])
                    fig.add_trace(go.Bar(
                        x=[row["label"]], y=[1],
                        marker_color=CORES_QUALIDADE.get(q, "#aaa"),
                        text=q, textposition="inside",
                        textfont=dict(color="white", size=10, family="Arial"),
                        showlegend=False,
                        hovertemplate=f"<b>{row['label']}</b><br>{titulo}: {q}<extra></extra>",
                    ))
                fig.update_yaxes(visible=False, range=[0, 1.6])
                estilo(fig, titulo, h=200)
                return fig

            with cc1:
                if qa_ok:
                    st.plotly_chart(_qual_chart(qa, "Qualidade ESF/EAP"), use_container_width=True)
            with cc2:
                if va_ok:
                    st.plotly_chart(_qual_chart(va, "Vínculo ESF/EAP"), use_container_width=True)

        # ── Cobertura ESF
        cob_map = {
            "qtEsfCredenciado": "Credenciadas",
            "qtEsfHomologado":  "Homologadas",
            "qtEsfTotalPgto":   "Com Pagamento",
        }
        cob_ok = {k: v for k, v in cob_map.items() if k in df_p.columns}
        if cob_ok and "label" in df_p.columns:
            st.markdown('<div class="sec-title">Cobertura – Equipes de Saúde da Família</div>', unsafe_allow_html=True)
            df_cob = df_p[["label"] + list(cob_ok.keys())].copy()
            df_cob = df_cob.melt(id_vars="label", var_name="Tipo", value_name="Qtd")
            df_cob["Tipo"] = df_cob["Tipo"].map(cob_ok)
            fig6 = px.line(df_cob, x="label", y="Qtd", color="Tipo",
                           color_discrete_sequence=[AZUL, LARANJA, VERDE_ESCURO], markers=True)
            estilo(fig6, "Equipes ESF – Credenciadas vs Homologadas vs Com Pagamento", h=300)
            st.plotly_chart(fig6, use_container_width=True)

        # ── Tabelas separadas por programa
        with st.expander("📄 Ver tabelas de pagamentos por programa"):

            def _colunas_grupo(palavras, excluir=None):
                ex = excluir or []
                return [c for c in df_p.columns
                        if any(p in c.lower() for p in palavras)
                        and not any(e in c.lower() for e in ex)
                        and c not in ("nuParcela", "label")]

            GRUPOS_APS = {
                "eSF":           _colunas_grupo(["esf"],      excluir=["esfr"]),
                "eAP":           _colunas_grupo(["eap"]),
                "eMulti":        _colunas_grupo(["emulti"]),
                "eSB":           _colunas_grupo(["esb"]),
                "CEO":           _colunas_grupo(["ceo"]),
                "LRPD":          _colunas_grupo(["lrpd"]),
                "ACS":           _colunas_grupo(["acs"]),
                "UOM":           _colunas_grupo(["uom"]),
                "SESB":          _colunas_grupo(["sesb"]),
                "eCR":           _colunas_grupo(["ecr"]),
                "eSFR":          _colunas_grupo(["esfr"]),
                "UBSF":          _colunas_grupo(["ubsf"]),
                "Microscopista": _colunas_grupo(["microsc"]),
                "Residência":    _colunas_grupo(["resid"]),
            }
            # Mantém apenas grupos que têm colunas com dados
            GRUPOS_APS = {k: v for k, v in GRUPOS_APS.items() if v}

            if not GRUPOS_APS:
                st.dataframe(df_pag, use_container_width=True, hide_index=True)
            else:
                tabs_prog = st.tabs(list(GRUPOS_APS.keys()))
                for tab_g, (nome_g, cols_g) in zip(tabs_prog, GRUPOS_APS.items()):
                    with tab_g:
                        cols_g_ok = [c for c in cols_g if c in df_p.columns]
                        if not cols_g_ok:
                            st.info(f"Sem colunas para {nome_g}")
                            continue
                        df_show = df_p[["label"] + cols_g_ok].copy()
                        df_show = df_show.rename(columns={"label": "Mês/Ano"})
                        # Formatar colunas monetárias (vl*) como R$
                        for c in cols_g_ok:
                            if c.startswith("vl"):
                                df_show[c] = df_show[c].apply(fmt_brl)
                        st.dataframe(df_show, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# ABA 3 – RECURSOS FNS
# ══════════════════════════════════════════════
with tab3:
    if df_acoes.empty:
        st.info("Sem dados FNS para o período.")
    else:
        total_fns   = df_acoes["Valor Líquido"].sum()
        total_d_fns = df_acoes["Desconto"].sum()
        n_acoes     = len(df_acoes)

        # KPIs
        st.markdown('<div class="sec-title">Indicadores FNS</div>', unsafe_allow_html=True)
        k1, k2, k3 = st.columns(3)
        with k1:
            st.markdown(kpi_card("Total Líquido FNS", fmt_brl(total_fns), f"{ano}"), unsafe_allow_html=True)
        with k2:
            st.markdown(kpi_card("Total Descontos", fmt_brl(total_d_fns), "", "vermelho"), unsafe_allow_html=True)
        with k3:
            st.markdown(kpi_card("Ações com Repasse", str(n_acoes), "ações distintas", "verde"), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Barras horizontais por bloco
        cl, cr = st.columns([3, 2])
        with cl:
            df_bloco = (df_acoes.groupby("Bloco")["Valor Líquido"].sum()
                        .reset_index().sort_values("Valor Líquido"))
            df_bloco = df_bloco[df_bloco["Bloco"].str.strip() != ""]
            fig7 = go.Figure(go.Bar(
                x=df_bloco["Valor Líquido"], y=df_bloco["Bloco"],
                orientation="h",
                marker=dict(color=PALETA[:len(df_bloco)]),
                text=df_bloco["Valor Líquido"].apply(lambda v: fmt_brl(v)),
                textposition="outside", textfont=dict(size=9),
            ))
            estilo(fig7, "Valor Líquido por Bloco de Financiamento", h=420)
            fig7.update_xaxes(tickprefix="R$ ", tickformat=",.0f", showgrid=True, gridcolor="#f0f0f0")
            fig7.update_yaxes(showgrid=False, tickfont=dict(size=9.5))
            st.plotly_chart(fig7, use_container_width=True)

        with cr:
            df_comp = (df_acoes[df_acoes["Componente"].str.strip() != ""]
                       .groupby("Componente")["Valor Líquido"].sum()
                       .reset_index())
            if not df_comp.empty:
                fig8 = px.pie(df_comp, names="Componente", values="Valor Líquido",
                              color_discrete_sequence=PALETA, hole=.4)
                fig8.update_traces(textinfo="percent", textfont_size=9.5)
                estilo(fig8, "Por Componente", h=420)
                st.plotly_chart(fig8, use_container_width=True)

        # Tabela detalhada
        st.markdown('<div class="sec-title">Detalhamento das Ações</div>', unsafe_allow_html=True)
        df_show = df_acoes.sort_values("Valor Líquido", ascending=False).copy()
        for c in ["Valor Total","Desconto","Valor Líquido"]:
            df_show[c] = df_show[c].apply(fmt_brl)
        st.dataframe(df_show, use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════
# ABA 4 – REPASSES / OBs
# ══════════════════════════════════════════════
with tab4:
    if xlrd is None:
        st.warning("⚠️ Biblioteca `xlrd` não instalada. Execute: `pip install xlrd`")
    elif df_obs.empty:
        st.info("Sem dados de Ordens Bancárias para o período selecionado.")
    else:
        # Converter valores
        for c in ["Valor Total","Desconto","Valor Líquido"]:
            if c in df_obs.columns:
                df_obs[c] = df_obs[c].apply(parse_brl)

        # Filtros
        st.markdown('<div class="sec-title">Filtros</div>', unsafe_allow_html=True)
        fc1, fc2 = st.columns(2)
        with fc1:
            meses_disp = sorted(df_obs["Mês"].unique().tolist()) if "Mês" in df_obs.columns else []
            meses_sel  = st.multiselect("Mês", meses_disp, default=meses_disp,
                                        format_func=lambda m: MESES[int(m)-1])
        with fc2:
            blocos_disp = sorted(df_obs["Bloco"].dropna().unique().tolist()) if "Bloco" in df_obs.columns else []
            blocos_sel  = st.multiselect("Bloco", blocos_disp, default=blocos_disp)

        df_f = df_obs.copy()
        if meses_sel and "Mês" in df_f.columns:
            df_f = df_f[df_f["Mês"].isin(meses_sel)]
        if blocos_sel and "Bloco" in df_f.columns:
            df_f = df_f[df_f["Bloco"].isin(blocos_sel)]

        # KPIs
        tot_liq  = df_f["Valor Líquido"].sum() if "Valor Líquido" in df_f.columns else 0
        tot_desc = df_f["Desconto"].sum()       if "Desconto"      in df_f.columns else 0
        n_obs    = len(df_f)

        k1, k2, k3 = st.columns(3)
        with k1:
            st.markdown(kpi_card("Total Líquido OBs", fmt_brl(tot_liq)), unsafe_allow_html=True)
        with k2:
            st.markdown(kpi_card("Total Descontos", fmt_brl(tot_desc), "", "vermelho"), unsafe_allow_html=True)
        with k3:
            st.markdown(kpi_card("Nº de OBs", str(n_obs), "lançamentos", "laranja"), unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        # Gráfico por mês
        if "Mês" in df_f.columns and "Valor Líquido" in df_f.columns:
            df_m = (df_f.groupby("Mês")["Valor Líquido"].sum()
                    .reset_index().sort_values("Mês"))
            df_m["label"] = df_m["Mês"].apply(lambda m: MESES[int(m)-1])

            fig9 = go.Figure(go.Bar(
                x=df_m["label"], y=df_m["Valor Líquido"],
                marker_color=LARANJA,
                text=df_m["Valor Líquido"].apply(fmt_brl),
                textposition="outside", textfont=dict(size=9),
            ))
            estilo(fig9, "Valor Líquido por Mês de Emissão da OB", h=300)
            fig9.update_yaxes(tickprefix="R$ ", tickformat=",.0f")
            st.plotly_chart(fig9, use_container_width=True)

        # Tabela
        st.markdown('<div class="sec-title">Ordens Bancárias</div>', unsafe_allow_html=True)
        cols_show = [c for c in [
            "Ano","Mês","Bloco","Grupo","Ação Detalhada","Competência/Parcela",
            "Nº OB","Data OB","Banco OB","Agência OB","Conta OB",
            "Valor Total","Desconto","Valor Líquido","Tipo Repasse",
        ] if c in df_f.columns]
        df_show_obs = df_f[cols_show].copy()
        for c in ["Valor Total","Desconto","Valor Líquido"]:
            if c in df_show_obs.columns:
                df_show_obs[c] = df_show_obs[c].apply(fmt_brl)
        st.dataframe(df_show_obs, use_container_width=True, hide_index=True)
