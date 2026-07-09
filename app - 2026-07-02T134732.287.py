"""
Controle de Fatura — app.py
Streamlit + Google Sheets (mesmo padrão arquitetural do SmartWallet)

CONFIGURAÇÃO NECESSÁRIA ANTES DO DEPLOY:
1. Criar uma planilha Google em branco, copiar o ID da URL e colar em SPREADSHEET_ID abaixo.
2. Compartilhar essa planilha (permissão de Editor) com a service account:
   carteira-python@ccmensal.iam.gserviceaccount.com
3. Nos Secrets do Streamlit Cloud, adicionar o bloco [gcp_service_account] com o JSON da
   mesma service account já usada no SmartWallet.
4. Rodar o app uma vez — ele cria sozinho as abas 'lancamentos', 'categorias' e 'config'
   com os cabeçalhos e categorias padrão, se ainda não existirem.
"""

import streamlit as st
import pandas as pd
import gspread
from gspread.utils import ValueRenderOption
from google.oauth2.service_account import Credentials
from datetime import datetime, date, timedelta
import plotly.graph_objects as go

# ───────────────────────── CONFIG ─────────────────────────

SPREADSHEET_ID = "1JvAXJm6ThknEv3j8xd8Q1xAhLGnrGLt3MRGgsXBIsow"

ABA_LANCAMENTOS = "lancamentos"
ABA_CATEGORIAS = "categorias"
ABA_CONFIG = "config"

FIXED_ID = "parcel"

CATEGORIAS_PADRAO = [
    {"id": "alim", "nome": "alimentação", "cor": "#D85A30", "valor_alvo": 1700, "fixo": False},
    {"id": "transp", "nome": "transporte", "cor": "#BA7517", "valor_alvo": 850, "fixo": False},
    {"id": "lazer", "nome": "lazer", "cor": "#D4537E", "valor_alvo": 425, "fixo": False},
    {"id": "cuidado", "nome": "cuidado pessoal", "cor": "#7F77DD", "valor_alvo": 340, "fixo": False},
    {"id": "super", "nome": "supermercado", "cor": "#639922", "valor_alvo": 680, "fixo": False},
    {"id": "assina", "nome": "assinaturas", "cor": "#888780", "valor_alvo": 170, "fixo": False},
    {"id": "saude", "nome": "saúde", "cor": "#1D9E75", "valor_alvo": 340, "fixo": False},
    {"id": "outros", "nome": "outros", "cor": "#666666", "valor_alvo": 340, "fixo": False},
    {"id": FIXED_ID, "nome": "parcelamentos", "cor": "#E5B800", "valor_alvo": 0, "fixo": True},
]

LIMITE_PADRAO = 8500

st.set_page_config(page_title="controle de fatura", layout="centered")

# Impede que st.columns empilhe verticalmente em telas estreitas (celular).
# Sem isso, qualquer layout lado a lado (setas de navegação, valor + botão de
# excluir, etc.) quebra em blocos empilhados no mobile.
st.markdown("""
<style>
@media (max-width: 640px) {
    div[data-testid="stHorizontalBlock"] {
        flex-wrap: nowrap !important;
    }
    div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        min-width: 0 !important;
    }
}
div[data-testid="stButton"] button {
    padding: 0.1rem 0.5rem !important;
    min-height: 1.8rem !important;
    height: 1.8rem !important;
    font-size: 0.8rem !important;
    line-height: 1 !important;
}
div[data-testid="stTextInput"] input,
div[data-testid="stDateInput"] input,
div[data-testid="stNumberInput"] input,
div[data-baseweb="select"] > div {
    min-height: 1.8rem !important;
    height: 1.8rem !important;
    padding: 0.2rem 0.5rem !important;
    font-size: 0.8rem !important;
}
div[data-baseweb="input"],
div[data-baseweb="base-input"] {
    min-height: 1.8rem !important;
}
div[data-testid="stTextInput"],
div[data-testid="stDateInput"],
div[data-testid="stNumberInput"] {
    margin-bottom: 0 !important;
}
div[data-testid="stWidgetLabel"] p {
    font-size: 0.78rem !important;
}
div[data-testid="stWidgetLabel"] {
    margin-bottom: 0.1rem !important;
}
</style>
""", unsafe_allow_html=True)

# ───────────────────────── CLIENTE GOOGLE SHEETS ─────────────────────────

@st.cache_resource
def get_client():
    secret = st.secrets["gcp_service_account"]
    if isinstance(secret, str):
        import json
        info = json.loads(secret)
    else:
        info = dict(secret)
    creds = Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"],
    )
    return gspread.authorize(creds)


def get_spreadsheet():
    return get_client().open_by_key(SPREADSHEET_ID)


def garantir_abas(sh):
    """Cria as abas necessárias com cabeçalho/dados padrão se ainda não existirem."""
    titulos_existentes = [ws.title for ws in sh.worksheets()]

    if ABA_LANCAMENTOS not in titulos_existentes:
        ws = sh.add_worksheet(title=ABA_LANCAMENTOS, rows=1000, cols=4)
        ws.append_row(["data", "categoria", "descricao", "valor"])

    if ABA_CATEGORIAS not in titulos_existentes:
        ws = sh.add_worksheet(title=ABA_CATEGORIAS, rows=50, cols=5)
        ws.append_row(["id", "nome", "cor", "valor_alvo", "fixo"])
        for c in CATEGORIAS_PADRAO:
            ws.append_row([c["id"], c["nome"], c["cor"], c["valor_alvo"], str(c["fixo"])])

    if ABA_CONFIG not in titulos_existentes:
        ws = sh.add_worksheet(title=ABA_CONFIG, rows=10, cols=2)
        ws.append_row(["chave", "valor"])
        ws.append_row(["limite_mensal", LIMITE_PADRAO])


# ───────────────────────── LEITURA / ESCRITA ─────────────────────────

@st.cache_data(ttl=3600)
def carregar_categorias():
    sh = get_spreadsheet()
    garantir_abas(sh)
    ws = sh.worksheet(ABA_CATEGORIAS)
    df = pd.DataFrame(ws.get_all_records(value_render_option=ValueRenderOption.unformatted))
    if df.empty or "valor_alvo" not in df.columns:
        # Planilha ainda no esquema antigo (percentual) ou vazia — usa os padrões limpos.
        # Abra a aba Orçamento e clique em "Salvar orçamento" uma vez para gravar
        # esse esquema novo na planilha.
        return pd.DataFrame(CATEGORIAS_PADRAO)
    df["valor_alvo"] = pd.to_numeric(df["valor_alvo"], errors="coerce").fillna(0).astype(float)
    df["fixo"] = df["fixo"].astype(str).str.lower().eq("true")
    return df



@st.cache_data(ttl=3600)
def carregar_config():
    sh = get_spreadsheet()
    garantir_abas(sh)
    ws = sh.worksheet(ABA_CONFIG)
    df = pd.DataFrame(ws.get_all_records(value_render_option=ValueRenderOption.unformatted))
    limite = LIMITE_PADRAO
    if not df.empty:
        linha = df[df["chave"] == "limite_mensal"]
        if not linha.empty:
            limite = float(linha.iloc[0]["valor"])
    return {"limite_mensal": limite}


def carregar_lancamentos():
    """Sem cache de longo prazo — recarrega a cada save/delete via session_state."""
    if "df_lancamentos" not in st.session_state:
        sh = get_spreadsheet()
        garantir_abas(sh)
        ws = sh.worksheet(ABA_LANCAMENTOS)
        df = pd.DataFrame(ws.get_all_records(value_render_option=ValueRenderOption.unformatted))
        if not df.empty:
            df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0)
            df["data"] = pd.to_datetime(df["data"]).dt.date
        st.session_state["df_lancamentos"] = df
    return st.session_state["df_lancamentos"]


def recarregar_lancamentos():
    st.session_state.pop("df_lancamentos", None)


def salvar_lancamento(data_str, categoria, descricao, valor):
    ws = get_spreadsheet().worksheet(ABA_LANCAMENTOS)
    ws.append_row([data_str, categoria, descricao, valor])
    recarregar_lancamentos()


def excluir_lancamento(data_str, categoria, descricao, valor):
    """Exclusão por conteúdo, não por índice — evita quebra de posição em uso concorrente."""
    ws = get_spreadsheet().worksheet(ABA_LANCAMENTOS)
    registros = ws.get_all_records(value_render_option=ValueRenderOption.unformatted)
    for i, r in enumerate(registros, start=2):  # linha 1 = cabeçalho
        if (str(r["data"]) == str(data_str) and r["categoria"] == categoria
                and str(r["descricao"]) == str(descricao) and float(r["valor"]) == float(valor)):
            ws.delete_rows(i)
            break
    recarregar_lancamentos()


def salvar_configuracoes(limite, df_categorias):
    sh = get_spreadsheet()

    ws_cfg = sh.worksheet(ABA_CONFIG)
    ws_cfg.clear()
    ws_cfg.update([["chave", "valor"], ["limite_mensal", limite]])

    ws_cat = sh.worksheet(ABA_CATEGORIAS)
    ws_cat.clear()
    linhas = [["id", "nome", "cor", "valor_alvo", "fixo"]]
    for _, r in df_categorias.iterrows():
        linhas.append([r["id"], r["nome"], r["cor"], r["valor_alvo"], str(r["fixo"])])
    ws_cat.update(linhas)

    st.cache_data.clear()


# ───────────────────────── FORMATAÇÃO ─────────────────────────

def formatar_brl(valor):
    return "R$ " + f"{valor:.0f}"


def fmt_pct(p):
    if p == 0:
        return "0%"
    return f"{p:.1f}".replace(".", ",") + "%"


def parse_valor(txt):
    """Converte texto pt-BR ('1.234,56' ou '1234,56' ou '1234.56') para float."""
    if not txt:
        return 0.0
    txt = str(txt).strip().replace(".", "").replace(",", ".") if "," in str(txt) else str(txt).strip()
    try:
        return float(txt)
    except ValueError:
        return 0.0


# ───────────────────────── LÓGICA DE NEGÓCIO ─────────────────────────

def mes_label(d):
    meses = ["jan", "fev", "mar", "abr", "mai", "jun",
             "jul", "ago", "set", "out", "nov", "dez"]
    return f"{meses[d.month - 1]}/{d.year % 100:02d}"


def filtrar_mes(df, ano, mes):
    if df.empty:
        return df
    return df[(pd.to_datetime(df["data"]).dt.year == ano) & (pd.to_datetime(df["data"]).dt.month == mes)]


def mes_anterior(ref):
    primeiro_dia_mes_atual = ref.replace(day=1)
    return (primeiro_dia_mes_atual - timedelta(days=1)).replace(day=1)


def dias_ate_proximo_mes(hoje):
    proximo_mes = hoje.month + 1
    proximo_ano = hoje.year
    if proximo_mes == 13:
        proximo_mes = 1
        proximo_ano += 1
    primeiro_dia_proximo = date(proximo_ano, proximo_mes, 1)
    return (primeiro_dia_proximo - hoje).days


def orcamento_categoria(row, limite=None):
    """O valor alvo já é salvo em R$ direto — não há mais round-trip via percentual."""
    return row["valor_alvo"]


# ───────────────────────── ESTADO DE NAVEGAÇÃO DE MÊS ─────────────────────────

if "mes_ref" not in st.session_state:
    hoje = date.today()
    st.session_state["mes_ref"] = date(hoje.year, hoje.month, 1)


def mudar_mes(delta):
    ref = st.session_state["mes_ref"]
    novo_mes = ref.month + delta
    novo_ano = ref.year
    if novo_mes == 0:
        novo_mes = 12
        novo_ano -= 1
    elif novo_mes == 13:
        novo_mes = 1
        novo_ano += 1
    candidato = date(novo_ano, novo_mes, 1)
    hoje = date.today().replace(day=1)
    if candidato <= hoje:
        st.session_state["mes_ref"] = candidato


def render_nav_mes(prefixo):
    """Navegador de mês com setas nas extremidades (mesmo espírito dos cards)."""
    colA, colB, colC = st.columns([1, 10, 1])
    with colA:
        st.button("‹", on_click=mudar_mes, args=(-1,), key=f"{prefixo}_prev",
                   use_container_width=True, type="tertiary")
    with colB:
        st.markdown(
            f"<div style='text-align:center;font-weight:500;padding-top:0.4rem;'>{mes_label(st.session_state['mes_ref'])}</div>",
            unsafe_allow_html=True,
        )
    with colC:
        st.button("›", on_click=mudar_mes, args=(1,), key=f"{prefixo}_next",
                   use_container_width=True, type="tertiary")


# ───────────────────────── APP ─────────────────────────

df_cat = carregar_categorias()
config = carregar_config()
limite_mensal = config["limite_mensal"]

def render_cards_limite(df_mes, limite_mensal):
    """Cards de total da fatura / limite disponível + texto central de dias + barra de progresso."""
    total_gasto = df_mes["valor"].sum() if not df_mes.empty else 0
    disponivel = limite_mensal - total_gasto

    cap_esq = fmt_pct(total_gasto / limite_mensal * 100) + " do limite" if limite_mensal > 0 else ""
    if disponivel >= 0 and limite_mensal > 0:
        cap_dir = fmt_pct(disponivel / limite_mensal * 100) + " restante"
    elif disponivel < 0:
        cap_dir = "excedido"
    else:
        cap_dir = ""

    hoje = date.today()
    prox_mes_num = hoje.month + 1 if hoje.month < 12 else 1
    prox_ano = hoje.year if hoje.month < 12 else hoje.year + 1
    meses_extenso = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
                      "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
    nome_prox_mes = meses_extenso[prox_mes_num - 1]
    texto_dias = f"{dias_ate_proximo_mes(hoje)} dias até {nome_prox_mes}"

    st.markdown(f"""
    <div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:0.5rem;'>
        <div style='text-align:left;'>
            <div style='font-size:14px;color:rgba(250,250,250,0.6);'>total da fatura (R$)</div>
            <div style='font-size:2.25rem;font-weight:600;line-height:1.2;'>{total_gasto:.0f}</div>
            <div style='font-size:13px;color:#888;margin-top:2px;'>{cap_esq}</div>
        </div>
        <div style='text-align:center;color:#888;font-size:12px;flex:1;'>{texto_dias}</div>
        <div style='text-align:right;'>
            <div style='font-size:14px;color:rgba(250,250,250,0.6);'>limite disponível (R$)</div>
            <div style='font-size:2.25rem;font-weight:600;line-height:1.2;'>{abs(disponivel):.0f}</div>
            <div style='font-size:13px;color:#888;margin-top:2px;'>{cap_dir}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.progress(min(total_gasto / limite_mensal, 1.0) if limite_mensal > 0 else 0)


def render_grafico_categorias(df_mes, limite_mensal, df_cat, chave):
    """Gráfico de rosca + barras de progresso por categoria."""
    total_gasto = df_mes["valor"].sum() if not df_mes.empty else 0
    gasto_parcel = df_mes[df_mes["categoria"] == FIXED_ID]["valor"].sum() if not df_mes.empty else 0
    gasto_outros = max(0, total_gasto - gasto_parcel)
    disponivel = limite_mensal - total_gasto

    # Ordem FIXA (não reordena conforme os valores mudam). O Plotly sempre desenha
    # a 1ª fatia da lista começando às 12h e indo para a direita, não importa o
    # "direction" — confirmado empiricamente. parcelamentos precisa ser a ÚLTIMA
    # fatia (termina às 12h, ocupando o 2º quadrante); disponível vem logo antes
    # dela (fica adjacente, no 3º quadrante); despesas fica com o restante.
    fatia_disp = max(0, disponivel)
    labels, valores, cores = [], [], []
    if fatia_disp > 0:
        labels.append("disponível"); valores.append(fatia_disp); cores.append("#2a2a2a")
    if gasto_outros > 0:
        labels.append("despesas"); valores.append(gasto_outros); cores.append("#D85A30")
    if gasto_parcel > 0:
        labels.append("parcelamentos"); valores.append(gasto_parcel); cores.append("#E5B800")

    if valores:
        fig = go.Figure(data=[go.Pie(
            labels=labels, values=valores, hole=0.65, marker=dict(colors=cores),
            sort=False, direction="clockwise", rotation=0,
        )])
        fig.update_layout(
            showlegend=True,
            legend=dict(orientation="h", yanchor="top", y=-0.05, xanchor="center", x=0.5),
            margin=dict(t=10, b=10, l=10, r=10), height=320,
        )
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False}, key=f"pie_{chave}")
    else:
        st.caption("sem lançamentos neste mês.")

    gasto_por_cat = df_mes.groupby("categoria")["valor"].sum().to_dict() if not df_mes.empty else {}
    cats_ordenadas = df_cat[df_cat["id"] != FIXED_ID].copy()
    cats_ordenadas["gasto"] = cats_ordenadas["id"].map(gasto_por_cat).fillna(0)
    cats_ordenadas = cats_ordenadas.sort_values("gasto", ascending=False)

    if cats_ordenadas.empty:
        st.caption("nenhuma categoria configurada.")
    else:
        for _, c in cats_ordenadas.iterrows():
            orcamento = orcamento_categoria(c, limite_mensal)
            gasto = c["gasto"]
            pct = min(gasto / orcamento, 1.0) if orcamento > 0 else 0
            estourou = gasto > orcamento and orcamento > 0
            st.markdown(
                f"<div style='display:flex;justify-content:space-between;font-size:13px;margin-bottom:2px;'>"
                f"<span>{c['nome'].lower()}</span>"
                f"<span style='color:{'#e05252' if estourou else '#888'};'>"
                f"{formatar_brl(gasto)} / {formatar_brl(orcamento)}</span></div>",
                unsafe_allow_html=True,
            )
            st.progress(pct)

        gasto_parcel_mes = gasto_por_cat.get(FIXED_ID, 0)
        if gasto_parcel_mes > 0:
            st.markdown(
                f"<div style='display:flex;justify-content:space-between;font-size:13px;margin-top:6px;'>"
                f"<span>parcelamentos</span><span style='color:#888;'>{formatar_brl(gasto_parcel_mes)}</span></div>",
                unsafe_allow_html=True,
            )


aba_lancar, aba_dash, aba_hist, aba_orc = st.tabs(
    ["lançar", "dashboard", "histórico", "orçamento"]
)

# ───────────────────────── ABA: LANÇAR ─────────────────────────

with aba_lancar:
    hoje = date.today()
    df_mes_atual_lancar = filtrar_mes(carregar_lancamentos(), hoje.year, hoje.month)
    render_cards_limite(df_mes_atual_lancar, limite_mensal)
    st.markdown("---")

    @st.fragment
    def form_lancamento():
        mes_ant = mes_anterior(hoje)
        df_mes_ant = filtrar_mes(carregar_lancamentos(), mes_ant.year, mes_ant.month)
        gasto_mes_ant = df_mes_ant.groupby("categoria")["valor"].sum().to_dict() if not df_mes_ant.empty else {}

        todas_cats = df_cat.copy()
        todas_cats["nome_lower"] = todas_cats["nome"].str.lower()
        todas_cats = todas_cats.sort_values("nome_lower")

        opcoes_ids = todas_cats["id"].tolist()
        opcoes_labels = dict(zip(todas_cats["id"], todas_cats["nome_lower"]))

        if gasto_mes_ant:
            categoria_default = max(gasto_mes_ant, key=gasto_mes_ant.get)
            index_default = opcoes_ids.index(categoria_default) if categoria_default in opcoes_ids else 0
        else:
            index_default = 0

        with st.form("form_lancar", clear_on_submit=True):
            col_cat, col_desc = st.columns([1, 1.4])
            with col_cat:
                categoria_id = st.selectbox(
                    "categoria", options=opcoes_ids, index=index_default,
                    format_func=lambda x: opcoes_labels[x]
                )
            with col_desc:
                descricao = st.text_input("descrição (opcional)", placeholder="ex: almoço no tabuã, uber p/ valinhos...")

            col_val, col_data = st.columns([1, 1])
            with col_val:
                valor_txt = st.text_input("valor (R$)", placeholder="0,00")
            with col_data:
                data_lanc = st.date_input("data", value=date.today(), format="DD/MM/YYYY")

            enviado = st.form_submit_button("registrar gasto", use_container_width=True)

            if enviado:
                valor = parse_valor(valor_txt)
                if valor <= 0:
                    st.error("informe um valor válido.")
                else:
                    salvar_lancamento(data_lanc.isoformat(), categoria_id, descricao, valor)
                    st.success("registrado!")
                    st.rerun(scope="app")

    form_lancamento()

with aba_dash:
    render_nav_mes("dash")

    ref = st.session_state["mes_ref"]
    df_lanc = carregar_lancamentos()
    df_mes = filtrar_mes(df_lanc, ref.year, ref.month)

    render_grafico_categorias(df_mes, limite_mensal, df_cat, chave="dash")

with aba_hist:
    render_nav_mes("hist")

    ref = st.session_state["mes_ref"]
    df_lanc = carregar_lancamentos()
    df_mes = filtrar_mes(df_lanc, ref.year, ref.month)

    if df_mes.empty:
        st.caption("nenhum lançamento neste mês.")
    else:
        df_mes_ordenado = df_mes.sort_values("data", ascending=False)
        mapa_nomes = dict(zip(df_cat["id"], df_cat["nome"]))
        for idx_row, (_, r) in enumerate(df_mes_ordenado.iterrows()):
            nome_cat = mapa_nomes.get(r["categoria"], "sem categoria").lower()
            item_id = f"{idx_row}_{r['data']}_{r['categoria']}_{r['valor']}"
            pendente_key = f"pendente_del_{item_id}"

            col_acao, col_desc, col_val = st.columns([0.4, 4, 2])
            with col_acao:
                if st.session_state.get(pendente_key):
                    subok, subcancel = st.columns(2)
                    with subok:
                        if st.button("✓", key=f"ok_{item_id}", use_container_width=True, type="primary"):
                            excluir_lancamento(r["data"].isoformat(), r["categoria"], r["descricao"], r["valor"])
                            st.session_state.pop(pendente_key, None)
                            st.rerun(scope="app")
                    with subcancel:
                        if st.button("✕", key=f"cancel_{item_id}", use_container_width=True):
                            st.session_state.pop(pendente_key, None)
                            st.rerun()
                else:
                    if st.button("✕", key=f"del_{item_id}", use_container_width=True):
                        st.session_state[pendente_key] = True
                        st.rerun()
            with col_desc:
                st.markdown(f"**{r['descricao'] or nome_cat}**  \n<span style='color:#888;font-size:12px;'>{nome_cat} · {r['data'].strftime('%d/%m/%y')}</span>", unsafe_allow_html=True)
            with col_val:
                st.markdown(f"<div style='text-align:right;padding-top:0.5rem;'>− {formatar_brl(r['valor'])}</div>", unsafe_allow_html=True)

        csv = df_mes_ordenado.rename(columns={"data": "data", "categoria": "categoria", "descricao": "descrição", "valor": "valor (R$)"})
        st.download_button(
            "↓ exportar histórico (csv)",
            csv.to_csv(index=False, sep=";").encode("utf-8-sig"),
            file_name=f"fatura_{ref.year}_{ref.month:02d}.csv",
            use_container_width=True,
        )

# ───────────────────────── ABA: ORÇAMENTO ─────────────────────────

with aba_orc:
    if "orc_versao" not in st.session_state:
        st.session_state["orc_versao"] = 0
    versao = st.session_state["orc_versao"]

    _, col_limite = st.columns([3, 1])
    with col_limite:
        limite_txt = st.text_input("limite mensal do cartão (R$)", value=f"{limite_mensal:.0f}",
                                    key=f"limite_{versao}")
    limite_novo = parse_valor(limite_txt)

    st.write("")

    df_cat_edit = df_cat.copy()
    valores_editados = {}
    for _, c in df_cat_edit[df_cat_edit["fixo"] == False].iterrows():
        orcamento_atual = orcamento_categoria(c, limite_mensal)
        col1, col2 = st.columns([5, 1.1])
        with col1:
            st.markdown(f"<div style='padding-top:0.5rem;'>{c['nome'].lower()}</div>", unsafe_allow_html=True)
        with col2:
            v_txt = st.text_input("valor", value=f"{orcamento_atual:.0f}",
                                   key=f"orc_{c['id']}_{versao}", label_visibility="collapsed")
        valores_editados[c["id"]] = parse_valor(v_txt)
        st.markdown("<hr style='margin:2px 0;border-color:#242424;'>", unsafe_allow_html=True)

    soma = sum(valores_editados.values())
    st.write("")
    if limite_novo > 0 and soma > limite_novo:
        st.markdown(f"<span style='color:#e05252;'>excede em {formatar_brl(soma - limite_novo)}</span>", unsafe_allow_html=True)
    else:
        st.caption(f"disponível: {formatar_brl(max(0, limite_novo - soma))}")

    with st.expander("+ nova categoria"):
        col_nome, col_val = st.columns([3, 1])
        with col_nome:
            novo_nome = st.text_input("nome da categoria", key="novo_nome")
        with col_val:
            novo_valor_txt = st.text_input("limite em R$", key="novo_valor")
        if st.button("adicionar categoria"):
            if not novo_nome.strip():
                st.error("digite o nome da categoria.")
            else:
                novo_valor = parse_valor(novo_valor_txt)
                if novo_valor > 1_000_000:
                    st.error("valor muito alto — confira se não digitou zeros a mais.")
                elif limite_novo > 0 and (soma + novo_valor) > limite_novo:
                    st.error("soma ultrapassaria o limite.")
                else:
                    novo_id = "cat_" + str(int(datetime.now().timestamp()))
                    nova_linha = pd.DataFrame([{
                        "id": novo_id, "nome": novo_nome.strip().lower(), "cor": "#888780",
                        "valor_alvo": novo_valor,
                        "fixo": False,
                    }])
                    df_final = pd.concat([df_cat_edit, nova_linha], ignore_index=True)
                    salvar_configuracoes(limite_novo, df_final)
                    st.session_state["orc_versao"] += 1
                    st.success(f"'{novo_nome.strip().lower()}' adicionada.")
                    st.rerun(scope="app")

    st.write("")
    if st.button("salvar orçamento", use_container_width=True):
        if limite_novo <= 0:
            st.error("informe um limite mensal válido antes de salvar.")
        elif any(v > 1_000_000 for v in valores_editados.values()):
            st.error("algum valor de categoria está muito alto — confira se não digitou zeros a mais.")
        elif soma > limite_novo:
            st.error("soma ultrapassa o limite.")
        else:
            df_cat_edit.loc[df_cat_edit["fixo"] == False, "valor_alvo"] = df_cat_edit[df_cat_edit["fixo"] == False]["id"].map(
                lambda i: valores_editados[i]
            )
            salvar_configuracoes(limite_novo, df_cat_edit)
            st.session_state["orc_versao"] += 1
            st.success("orçamento salvo!")
            st.rerun(scope="app")

