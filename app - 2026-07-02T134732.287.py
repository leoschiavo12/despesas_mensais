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
from google.oauth2.service_account import Credentials
from datetime import datetime, date
import plotly.graph_objects as go

# ───────────────────────── CONFIG ─────────────────────────

SPREADSHEET_ID = "1JvAXJm6ThknEv3j8xd8Q1xAhLGnrGLt3MRGgsXBIsow"

ABA_LANCAMENTOS = "lancamentos"
ABA_CATEGORIAS = "categorias"
ABA_CONFIG = "config"

FIXED_ID = "parcel"

CATEGORIAS_PADRAO = [
    {"id": "alim", "nome": "Alimentação", "cor": "#D85A30", "pct_alvo": 20, "fixo": False},
    {"id": "transp", "nome": "Transporte", "cor": "#BA7517", "pct_alvo": 10, "fixo": False},
    {"id": "lazer", "nome": "Lazer", "cor": "#D4537E", "pct_alvo": 5, "fixo": False},
    {"id": "cuidado", "nome": "Cuidado pessoal", "cor": "#7F77DD", "pct_alvo": 4, "fixo": False},
    {"id": "super", "nome": "Supermercado", "cor": "#639922", "pct_alvo": 8, "fixo": False},
    {"id": "assina", "nome": "Assinaturas", "cor": "#888780", "pct_alvo": 2, "fixo": False},
    {"id": "saude", "nome": "Saúde", "cor": "#1D9E75", "pct_alvo": 4, "fixo": False},
    {"id": "outros", "nome": "Outros", "cor": "#666666", "pct_alvo": 4, "fixo": False},
    {"id": FIXED_ID, "nome": "Parcelamentos", "cor": "#E5B800", "pct_alvo": 0, "fixo": True},
]

LIMITE_PADRAO = 8500

st.set_page_config(page_title="Controle de Fatura", page_icon="💳", layout="centered")

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
        ws.append_row(["id", "nome", "cor", "pct_alvo", "fixo"])
        for c in CATEGORIAS_PADRAO:
            ws.append_row([c["id"], c["nome"], c["cor"], c["pct_alvo"], str(c["fixo"])])

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
    df = pd.DataFrame(ws.get_all_records())
    if df.empty:
        return pd.DataFrame(CATEGORIAS_PADRAO).rename(columns={"nome": "nome", "cor": "cor"})
    df["pct_alvo"] = pd.to_numeric(df["pct_alvo"], errors="coerce").fillna(0)
    df["fixo"] = df["fixo"].astype(str).str.lower().eq("true")
    return df


@st.cache_data(ttl=3600)
def carregar_config():
    sh = get_spreadsheet()
    garantir_abas(sh)
    ws = sh.worksheet(ABA_CONFIG)
    df = pd.DataFrame(ws.get_all_records())
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
        df = pd.DataFrame(ws.get_all_records())
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
    registros = ws.get_all_records()
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
    linhas = [["id", "nome", "cor", "pct_alvo", "fixo"]]
    for _, r in df_categorias.iterrows():
        linhas.append([r["id"], r["nome"], r["cor"], r["pct_alvo"], str(r["fixo"])])
    ws_cat.update(linhas)

    st.cache_data.clear()


# ───────────────────────── FORMATAÇÃO ─────────────────────────

def formatar_brl(valor):
    return "R$ " + f"{valor:,.0f}".replace(",", ".")


def fmt_pct(p):
    if p == 0:
        return "0%"
    texto = f"{p:.2f}".rstrip("0").rstrip(".")
    return f"{texto}%"


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
    meses = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
             "agosto", "setembro", "outubro", "novembro", "dezembro"]
    return f"{meses[d.month - 1]} de {d.year}"


def filtrar_mes(df, ano, mes):
    if df.empty:
        return df
    return df[(pd.to_datetime(df["data"]).dt.year == ano) & (pd.to_datetime(df["data"]).dt.month == mes)]


def orcamento_categoria(row, limite):
    return limite * (row["pct_alvo"] / 100)


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


# ───────────────────────── APP ─────────────────────────

st.title("💳 Controle de Fatura")

df_cat = carregar_categorias()
config = carregar_config()
limite_mensal = config["limite_mensal"]

aba_lancar, aba_dash, aba_hist, aba_orc = st.tabs(
    ["➕ Lançar", "📊 Dashboard", "📜 Histórico", "⚙️ Orçamento"]
)

# ───────────────────────── ABA: LANÇAR ─────────────────────────

with aba_lancar:

    @st.fragment
    def form_lancamento():
        df_mes_atual = filtrar_mes(carregar_lancamentos(), date.today().year, date.today().month)
        gasto_por_cat = df_mes_atual.groupby("categoria")["valor"].sum().to_dict() if not df_mes_atual.empty else {}

        cats_ordenadas = df_cat[df_cat["id"] != FIXED_ID].copy()
        cats_ordenadas["gasto"] = cats_ordenadas["id"].map(gasto_por_cat).fillna(0)
        cats_ordenadas = cats_ordenadas.sort_values("gasto", ascending=False)

        opcoes_ids = [FIXED_ID] + cats_ordenadas["id"].tolist()
        opcoes_labels = {FIXED_ID: "Parcelamentos"}
        opcoes_labels.update(dict(zip(cats_ordenadas["id"], cats_ordenadas["nome"])))

        with st.form("form_lancar", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                categoria_id = st.selectbox(
                    "Categoria", options=opcoes_ids, format_func=lambda x: opcoes_labels[x]
                )
            with col2:
                valor_txt = st.text_input("Valor (R$)", placeholder="0,00")
            descricao = st.text_input("Descrição (opcional)", placeholder="ex: almoço no Tabuã, Uber p/ Valinhos...")
            data_lanc = st.date_input("Data", value=date.today(), format="DD/MM/YYYY")
            enviado = st.form_submit_button("Registrar gasto", use_container_width=True)

            if enviado:
                valor = parse_valor(valor_txt)
                if valor <= 0:
                    st.error("Informe um valor válido.")
                else:
                    salvar_lancamento(data_lanc.isoformat(), categoria_id, descricao, valor)
                    st.success("Registrado!")
                    st.rerun(scope="app")

        st.markdown("##### Resumo do mês")
        if cats_ordenadas.empty:
            st.caption("Nenhuma categoria configurada.")
        else:
            for _, c in cats_ordenadas.iterrows():
                orcamento = orcamento_categoria(c, limite_mensal)
                gasto = c["gasto"]
                pct = min(gasto / orcamento, 1.0) if orcamento > 0 else 0
                estourou = gasto > orcamento and orcamento > 0
                st.markdown(
                    f"<div style='display:flex;justify-content:space-between;font-size:13px;margin-bottom:2px;'>"
                    f"<span>{c['nome']}</span>"
                    f"<span style='color:{'#e05252' if estourou else '#888'};'>"
                    f"{formatar_brl(gasto)} / {formatar_brl(orcamento)}</span></div>",
                    unsafe_allow_html=True,
                )
                st.progress(pct)

    form_lancamento()

# ───────────────────────── ABA: DASHBOARD ─────────────────────────

with aba_dash:
    colA, colB, colC = st.columns([1, 3, 1])
    with colA:
        st.button("‹", on_click=mudar_mes, args=(-1,), key="dash_prev")
    with colB:
        st.markdown(f"<div style='text-align:center;font-weight:500;'>{mes_label(st.session_state['mes_ref'])}</div>", unsafe_allow_html=True)
    with colC:
        st.button("›", on_click=mudar_mes, args=(1,), key="dash_next")

    ref = st.session_state["mes_ref"]
    df_lanc = carregar_lancamentos()
    df_mes = filtrar_mes(df_lanc, ref.year, ref.month)

    total_gasto = df_mes["valor"].sum() if not df_mes.empty else 0
    gasto_parcel = df_mes[df_mes["categoria"] == FIXED_ID]["valor"].sum() if not df_mes.empty else 0
    gasto_outros = max(0, total_gasto - gasto_parcel)
    limite_efetivo = limite_mensal + gasto_parcel
    disponivel = limite_efetivo - total_gasto

    c1, c2, c3 = st.columns(3)
    c1.metric("Limite", formatar_brl(limite_efetivo), "base + parcel." if gasto_parcel > 0 else "100%")
    c2.metric("Despesas do mês", formatar_brl(gasto_outros),
              fmt_pct(gasto_outros / limite_efetivo * 100) + " do limite" if limite_efetivo > 0 else None)
    c3.metric("Disponível", formatar_brl(abs(disponivel)),
              (fmt_pct(max(0, disponivel) / limite_efetivo * 100) + " restante") if disponivel >= 0 else "excedido",
              delta_color="normal" if disponivel >= 0 else "inverse")

    st.progress(min(total_gasto / limite_efetivo, 1.0) if limite_efetivo > 0 else 0)

    st.markdown("##### Composição da fatura")
    fatia_disp = max(0, disponivel)
    labels, valores, cores = [], [], []
    if gasto_outros > 0:
        labels.append("Despesas"); valores.append(gasto_outros); cores.append("#D85A30")
    if gasto_parcel > 0:
        labels.append("Parcelamentos"); valores.append(gasto_parcel); cores.append("#E5B800")
    if fatia_disp > 0:
        labels.append("Disponível"); valores.append(fatia_disp); cores.append("#2a2a2a")

    if valores:
        fig = go.Figure(data=[go.Pie(labels=labels, values=valores, hole=0.65, marker=dict(colors=cores))])
        fig.update_layout(showlegend=True, margin=dict(t=10, b=10, l=10, r=10), height=280)
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.caption("Sem lançamentos neste mês.")

    st.markdown("##### Por categoria")
    if not df_mes.empty:
        gasto_por_cat = df_mes.groupby("categoria")["valor"].sum().to_dict()
    else:
        gasto_por_cat = {}
    cats_view = df_cat[df_cat["id"] != FIXED_ID].copy()
    cats_view["gasto"] = cats_view["id"].map(gasto_por_cat).fillna(0)
    cats_view = cats_view.sort_values("gasto", ascending=False)
    for _, c in cats_view.iterrows():
        orcamento = orcamento_categoria(c, limite_mensal)
        if orcamento == 0 and c["gasto"] == 0:
            continue
        gasto = c["gasto"]
        pct = min(gasto / orcamento, 1.0) if orcamento > 0 else 0
        estourou = gasto > orcamento and orcamento > 0
        st.markdown(
            f"<div style='display:flex;justify-content:space-between;font-size:13px;margin-bottom:2px;'>"
            f"<span>{c['nome']}</span>"
            f"<span style='color:{'#e05252' if estourou else '#888'};'>"
            f"{formatar_brl(gasto)} / {formatar_brl(orcamento)}</span></div>",
            unsafe_allow_html=True,
        )
        st.progress(pct)

# ───────────────────────── ABA: HISTÓRICO ─────────────────────────

with aba_hist:
    colA, colB, colC = st.columns([1, 3, 1])
    with colA:
        st.button("‹", on_click=mudar_mes, args=(-1,), key="hist_prev")
    with colB:
        st.markdown(f"<div style='text-align:center;font-weight:500;'>{mes_label(st.session_state['mes_ref'])}</div>", unsafe_allow_html=True)
    with colC:
        st.button("›", on_click=mudar_mes, args=(1,), key="hist_next")

    ref = st.session_state["mes_ref"]
    df_lanc = carregar_lancamentos()
    df_mes = filtrar_mes(df_lanc, ref.year, ref.month)

    if df_mes.empty:
        st.caption("Nenhum lançamento neste mês.")
    else:
        df_mes_ordenado = df_mes.sort_values("data", ascending=False)
        mapa_nomes = dict(zip(df_cat["id"], df_cat["nome"]))
        for _, r in df_mes_ordenado.iterrows():
            nome_cat = mapa_nomes.get(r["categoria"], "Sem categoria")
            col1, col2, col3 = st.columns([5, 2, 1])
            with col1:
                st.markdown(f"**{r['descricao'] or nome_cat}**  \n<span style='color:#888;font-size:12px;'>{nome_cat} · {r['data'].strftime('%d/%m/%y')}</span>", unsafe_allow_html=True)
            with col2:
                st.markdown(f"<div style='text-align:right;'>− {formatar_brl(r['valor'])}</div>", unsafe_allow_html=True)
            with col3:
                if st.button("✕", key=f"del_{r['data']}_{r['categoria']}_{r['descricao']}_{r['valor']}"):
                    excluir_lancamento(r["data"].isoformat(), r["categoria"], r["descricao"], r["valor"])
                    st.rerun(scope="app")

        csv = df_mes_ordenado.rename(columns={"data": "Data", "categoria": "Categoria", "descricao": "Descrição", "valor": "Valor (R$)"})
        st.download_button(
            "↓ Exportar histórico (CSV)",
            csv.to_csv(index=False, sep=";").encode("utf-8-sig"),
            file_name=f"fatura_{ref.year}_{ref.month:02d}.csv",
            use_container_width=True,
        )

# ───────────────────────── ABA: ORÇAMENTO ─────────────────────────

with aba_orc:
    limite_txt = st.text_input("Limite mensal do cartão (R$)", value=f"{limite_mensal:.2f}".replace(".", ","))
    limite_novo = parse_valor(limite_txt)

    st.markdown("##### Categorias")
    st.caption("Digite o valor em R$ — a % é calculada automaticamente")

    df_cat_edit = df_cat.copy()
    valores_editados = {}
    for _, c in df_cat_edit[df_cat_edit["fixo"] == False].iterrows():
        orcamento_atual = orcamento_categoria(c, limite_mensal)
        col1, col2, col3 = st.columns([3, 2, 1])
        with col1:
            st.markdown(f"🔵 {c['nome']}")
        with col2:
            v_txt = st.text_input("valor", value=f"{orcamento_atual:.2f}".replace(".", ","),
                                   key=f"orc_{c['id']}", label_visibility="collapsed")
        with col3:
            v = parse_valor(v_txt)
            pct_calc = (v / limite_novo * 100) if limite_novo > 0 else 0
            st.markdown(f"<div style='text-align:right;color:#888;'>{fmt_pct(pct_calc)}</div>", unsafe_allow_html=True)
        valores_editados[c["id"]] = v

    soma = sum(valores_editados.values())
    pct_usado = (soma / limite_novo * 100) if limite_novo > 0 else 0
    st.progress(min(pct_usado / 100, 1.0))
    if pct_usado > 100:
        st.markdown(f"<span style='color:#e05252;'>Excede em {formatar_brl(soma - limite_novo)}</span>", unsafe_allow_html=True)
    else:
        st.caption(f"Alocado: {fmt_pct(pct_usado)} · Disponível: {formatar_brl(max(0, limite_novo - soma))}")

    with st.expander("+ Nova categoria"):
        novo_nome = st.text_input("Nome da categoria", key="novo_nome")
        novo_valor_txt = st.text_input("Limite em R$", key="novo_valor")
        if st.button("Adicionar categoria"):
            if not novo_nome.strip():
                st.error("Digite o nome da categoria.")
            else:
                novo_valor = parse_valor(novo_valor_txt)
                if limite_novo > 0 and (soma + novo_valor) > limite_novo:
                    st.error("Soma ultrapassaria o limite.")
                else:
                    novo_id = "cat_" + str(int(datetime.now().timestamp()))
                    nova_linha = pd.DataFrame([{
                        "id": novo_id, "nome": novo_nome.strip(), "cor": "#888780",
                        "pct_alvo": (novo_valor / limite_novo * 100) if limite_novo > 0 else 0,
                        "fixo": False,
                    }])
                    df_final = pd.concat([df_cat_edit, nova_linha], ignore_index=True)
                    salvar_configuracoes(limite_novo, df_final)
                    st.success(f"'{novo_nome}' adicionada.")
                    st.rerun(scope="app")

    st.markdown("---")
    col_save, col_del = st.columns(2)
    with col_save:
        if st.button("💾 Salvar orçamento", use_container_width=True):
            if limite_novo > 0 and soma > limite_novo:
                st.error("Soma ultrapassa o limite.")
            else:
                df_cat_edit.loc[df_cat_edit["fixo"] == False, "pct_alvo"] = df_cat_edit[df_cat_edit["fixo"] == False]["id"].map(
                    lambda i: (valores_editados[i] / limite_novo * 100) if limite_novo > 0 else 0
                )
                salvar_configuracoes(limite_novo, df_cat_edit)
                st.success("Orçamento salvo!")
                st.rerun(scope="app")
    with col_del:
        if st.button("🗑️ Apagar todos os lançamentos", use_container_width=True):
            st.session_state["confirmar_delete"] = True
    if st.session_state.get("confirmar_delete"):
        st.warning("Tem certeza? Essa ação apaga TODOS os lançamentos e não pode ser desfeita.")
        c1, c2 = st.columns(2)
        if c1.button("Sim, apagar tudo"):
            ws = get_spreadsheet().worksheet(ABA_LANCAMENTOS)
            ws.clear()
            ws.append_row(["data", "categoria", "descricao", "valor"])
            recarregar_lancamentos()
            st.session_state["confirmar_delete"] = False
            st.rerun(scope="app")
        if c2.button("Cancelar"):
            st.session_state["confirmar_delete"] = False
            st.rerun()
