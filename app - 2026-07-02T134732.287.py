"""
Controle de Fatura — app.py
Streamlit + Google Sheets (mesmo padrão arquitetural do SmartWallet)

CONFIGURAÇÃO (Secrets do Streamlit Cloud) — nenhum ID ou dado pessoal fica no código:
  [gcp_service_account]   → JSON da service account (a mesma do SmartWallet)
  [google_sheets]
  spreadsheet_id          = "ID da planilha deste app"
  carteira_spreadsheet_id = "ID da planilha da carteira (SmartWallet)"
As duas planilhas precisam estar compartilhadas (Editor) com a service account.

Na primeira execução o app cria sozinho as abas 'lancamentos', 'categorias', 'config' e 'pix'.
Parcelas usam a 5ª coluna 'parcelamento' de 'lancamentos'; receitas da PJ e investimentos
gravam na planilha da carteira, que o SmartWallet só lê.
"""

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import gspread
from gspread.utils import ValueRenderOption, DateTimeOption
import calendar
import re
from google.oauth2.service_account import Credentials
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo

# o Streamlit Cloud roda em UTC: toda data/hora do app usa o horário de Brasília
FUSO = ZoneInfo("America/Sao_Paulo")


def agora_br():
    return datetime.now(FUSO)


def hoje_br():
    return agora_br().date()

# ───────────────────────── CONFIG ─────────────────────────

_GS = st.secrets.get("google_sheets", {})
SPREADSHEET_ID = _GS.get("spreadsheet_id")

ABA_LANCAMENTOS = "lancamentos"
ABA_CATEGORIAS = "categorias"
ABA_CONFIG = "config"
# notas fiscais e reembolsos da PJ ficam na planilha da carteira (SmartWallet), aba pj_notas
CARTEIRA_SPREADSHEET_ID = _GS.get("carteira_spreadsheet_id")
ABA_NOTAS = "pj_notas"
RECEITA_TIPOS = ["nota fiscal", "reembolso", "pix"]
MESES3 = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
ABA_PIX = "pix"   # pagamentos e recebimentos via PIX — fora da fatura (não entra em lancamentos)

# investimentos: aba lancamentos da planilha da carteira (o SmartWallet só lê)
ABA_INVEST = "lancamentos"
ABA_ATIVOS = "ativos_info"
INVEST_HEADERS = ["data", "tipo", "ativo", "classe", "quantidade", "preco_unitario", "total"]
ATIVOS_FIXOS = [("BTC", "Cripto"), ("Renda+ 2050", "Tesouro Direto")]
PAISES_ETF = ["Brasil", "EUA", "China", "Europa", "Emergentes", "Global"]
COL_PARCELAMENTO = "parcelamento"   # 5ª coluna de lancamentos: id que liga as parcelas de uma compra
MAX_PARCELAS = 24

FIXED_ID = "parcel"

CATEGORIAS_PADRAO = [
    {"id": "alim", "nome": "alimentação", "cor": "#D85A30", "valor_alvo": 0, "fixo": False},
    {"id": "transp", "nome": "transporte", "cor": "#BA7517", "valor_alvo": 0, "fixo": False},
    {"id": "lazer", "nome": "lazer", "cor": "#D4537E", "valor_alvo": 0, "fixo": False},
    {"id": "cuidado", "nome": "cuidado pessoal", "cor": "#7F77DD", "valor_alvo": 0, "fixo": False},
    {"id": "super", "nome": "supermercado", "cor": "#639922", "valor_alvo": 0, "fixo": False},
    {"id": "assina", "nome": "assinaturas", "cor": "#888780", "valor_alvo": 0, "fixo": False},
    {"id": "saude", "nome": "saúde", "cor": "#1D9E75", "valor_alvo": 0, "fixo": False},
    {"id": "outros", "nome": "outros", "cor": "#666666", "valor_alvo": 0, "fixo": False},
    {"id": FIXED_ID, "nome": "parcelamentos", "cor": "#E5B800", "valor_alvo": 0, "fixo": True},
]

LIMITE_PADRAO = 5000   # só para planilha nova; o valor real fica na aba config

st.set_page_config(page_title="controle de fatura", layout="centered")

if not SPREADSHEET_ID or not CARTEIRA_SPREADSHEET_ID:
    st.error("faltam os IDs das planilhas nos Secrets: bloco [google_sheets] com "
             "spreadsheet_id e carteira_spreadsheet_id.")
    st.stop()

# Impede que st.columns empilhe verticalmente em telas estreitas (celular).
# Sem isso, qualquer layout lado a lado (setas de navegação, valor + botão de
# excluir, etc.) quebra em blocos empilhados no mobile.
st.markdown("""
<style>
/* Estreita o app inteiro (mesmo no desktop) — dá cara de app, não de planilha larga */
div[data-testid="stAppViewContainer"] .block-container,
section.main .block-container {
    max-width: 480px !important;
    margin-left: auto !important;
    margin-right: auto !important;
    padding-top: 2rem !important;
}

@media (max-width: 640px) {
    div[data-testid="stHorizontalBlock"] {
        flex-wrap: nowrap !important;
    }
    div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"] {
        min-width: 0 !important;
    }
}
/* altura única para campos e botão do formulário (texto alinhado entre eles) */
:root { --altura-campo: 2.5rem; }

/* botões pequenos do app (✕, setas, exportar) */
div[data-testid="stButton"] button,
div[data-testid="stDownloadButton"] button {
    padding: 0.1rem 0.5rem !important;
    min-height: 1.8rem !important;
    height: 1.8rem !important;
    font-size: 0.8rem !important;
    line-height: 1 !important;
}
/* botão de enviar do formulário: mesma altura dos campos */
div[data-testid="stFormSubmitButton"] button {
    min-height: var(--altura-campo) !important;
    height: var(--altura-campo) !important;
    font-size: 0.9rem !important;
    line-height: 1 !important;
    border-radius: 8px !important;
}

/* campos: a borda fica só no contêiner externo; o <input> interno é transparente
   e preenche o contêiner — evita a "borda dupla" que desalinhava o texto */
div[data-baseweb="input"] {
    min-height: var(--altura-campo) !important;
    height: var(--altura-campo) !important;
    border: 1px solid #2c2c33 !important;
    border-radius: 8px !important;
    box-sizing: border-box !important;
}
div[data-baseweb="base-input"] {
    min-height: 0 !important;
    height: 100% !important;
    border: none !important;
    background: transparent !important;
}
div[data-baseweb="input"] input {
    height: 100% !important;
    min-height: 0 !important;
    border: none !important;
    background: transparent !important;
    padding: 0 0.7rem !important;
    font-size: 0.9rem !important;
}
div[data-testid="stTextInput"],
div[data-testid="stDateInput"],
div[data-testid="stNumberInput"],
div[data-testid="stSelectbox"],
div[data-testid="stElementContainer"] {
    margin-bottom: 0 !important;
}
div[data-testid="stWidgetLabel"] p,
div[data-testid="stWidgetLabel"] label {
    font-size: 0.68rem !important;
    text-transform: uppercase !important;
    letter-spacing: 0.04em !important;
    color: rgba(230,230,230,0.55) !important;
    font-weight: 500 !important;
}
div[data-testid="stWidgetLabel"] {
    margin-bottom: 0.2rem !important;
}

/* seletor de categoria: mesma altura e borda única dos outros campos.
   Cobre o Streamlit novo (react-aria: role="group") e o antigo (baseweb). */
div[data-testid="stSelectbox"] div[role="group"],
div[data-baseweb="select"] > div {
    min-height: var(--altura-campo) !important;
    height: var(--altura-campo) !important;
    border: 1px solid #2c2c33 !important;
    border-radius: 8px !important;
    box-sizing: border-box !important;
    cursor: pointer !important;
}
div[data-testid="stSelectbox"] input {
    height: 100% !important;
    min-height: 0 !important;
    border: none !important;
    background: transparent !important;
    padding: 0 0.7rem !important;
    font-size: 0.9rem !important;
    caret-color: transparent !important;
    cursor: pointer !important;
}

div[data-testid="stForm"] {
    padding: 1rem 1.1rem !important;
    border: 1px solid #202024 !important;
    border-radius: 12px !important;
}
div[data-testid="stForm"] div[data-testid="stVerticalBlock"] {
    gap: 0.9rem !important;
}
/* formulário dentro da janela (st.dialog): sem moldura própria */
div[role="dialog"] div[data-testid="stForm"],
div[data-testid="stDialog"] div[data-testid="stForm"],
div[data-testid="stDialog"] [data-testid="stForm"] {
    border: none !important;
    padding: 0 !important;
}
/* botões que abrem os formulários: mesma altura dos campos */
div[class*="st-key-novo_"] button {
    min-height: var(--altura-campo) !important;
    height: var(--altura-campo) !important;
    font-size: 0.9rem !important;
    border-radius: 8px !important;
}
/* lista do resumo da aba receitas: sem grade, só uma linha fina entre os itens (como no SmartWallet) */
table.pf-tab { width: 100% !important; border-collapse: collapse !important; border: none !important;
    margin: 0 !important; font-size: 0.95rem; }
table.pf-tab tr, table.pf-tab tr:nth-child(even) { border: none !important; background: transparent !important; }
table.pf-tab td { padding: 0.55rem 0.3rem !important; border: none !important;
    border-bottom: 1px solid rgba(250,250,250,0.08) !important; background: transparent !important; }
table.pf-tab td:last-child { text-align: right; white-space: nowrap; font-variant-numeric: tabular-nums; }
table.pf-tab td.pos { color: #22c55e; }
table.pf-tab td.neg { color: #ef4444; }
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

    if ABA_PIX not in titulos_existentes:
        ws = sh.add_worksheet(title=ABA_PIX, rows=1000, cols=5)
        ws.append_row(["data", "tipo", "categoria", "descricao", "valor"])

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
            if COL_PARCELAMENTO not in df.columns:
                df[COL_PARCELAMENTO] = ""
            df[COL_PARCELAMENTO] = df[COL_PARCELAMENTO].fillna("").astype(str)
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


# ── parcelamentos ──
# cada parcela é um lançamento normal, datado no mês em que cai na fatura; a coluna
# "parcelamento" guarda o id da compra, e a descrição termina em " (i/n)".

RE_PARCELA = re.compile(r"^(.*) \((\d+)/(\d+)\)$")


def _somar_meses(d, n):
    m = d.month - 1 + n
    ano, mes = d.year + m // 12, m % 12 + 1
    return date(ano, mes, min(d.day, calendar.monthrange(ano, mes)[1]))


def _garantir_coluna_parcelamento(ws):
    if ws.col_count < 5:
        ws.add_cols(5 - ws.col_count)
    cab = ws.row_values(1)
    if len(cab) < 5 or cab[4] != COL_PARCELAMENTO:
        ws.update_cell(1, 5, COL_PARCELAMENTO)


def salvar_parcelado(data_compra, categoria, descricao, total, n):
    """divide o total em n parcelas mensais (a última leva os centavos da divisão)"""
    ws = get_spreadsheet().worksheet(ABA_LANCAMENTOS)
    _garantir_coluna_parcelamento(ws)
    gid = agora_br().strftime("%y%m%d%H%M%S")
    parcela = int(round(total * 100)) // n / 100
    linhas = []
    for i in range(n):
        valor = parcela if i < n - 1 else round(total - parcela * (n - 1), 2)
        linhas.append([_somar_meses(data_compra, i).isoformat(), categoria,
                       f"{descricao} ({i + 1}/{n})", valor, gid])
    ws.append_rows(linhas)
    recarregar_lancamentos()


def parcelamentos_abertos(df):
    """compras com parcelas em meses futuros: [{gid, nome, categoria, n, futuras: [(data, descricao, valor)]}]"""
    if df.empty or COL_PARCELAMENTO not in df.columns:
        return []
    hoje = hoje_br()
    fut = df[(df[COL_PARCELAMENTO] != "") &
             df["data"].apply(lambda d: (d.year, d.month) > (hoje.year, hoje.month))]
    grupos = []
    for gid, g in fut.groupby(COL_PARCELAMENTO):
        g = g.sort_values("data")
        m = RE_PARCELA.match(str(g.iloc[0]["descricao"]))
        grupos.append({
            "gid": gid, "nome": m.group(1) if m else str(g.iloc[0]["descricao"]),
            "n": int(m.group(3)) if m else len(g), "categoria": g.iloc[0]["categoria"],
            "futuras": [(r["data"], str(r["descricao"]), float(r["valor"])) for _, r in g.iterrows()],
        })
    return sorted(grupos, key=lambda x: x["futuras"][0][0])


def adiantar_parcelas(grupo, k, desconto):
    """troca as k últimas parcelas futuras por um lançamento hoje, com o desconto abatido"""
    alvo = grupo["futuras"][-k:]
    soma = round(sum(v for _, _, v in alvo), 2)
    nums = [RE_PARCELA.match(d) for _, d, _ in alvo]
    faixa = (f"{nums[0].group(2)}–{nums[-1].group(2)}/{grupo['n']}" if all(nums) and k > 1
             else (f"{nums[0].group(2)}/{grupo['n']}" if nums[0] else f"{k}x"))
    desc = f"{grupo['nome']} (adiant. {faixa}" + (f", desc. {fmt_brl2(desconto)})" if desconto > 0 else ")")
    ws = get_spreadsheet().worksheet(ABA_LANCAMENTOS)
    # primeiro grava o adiantamento; depois apaga as parcelas, de baixo para cima
    ws.append_row([hoje_br().isoformat(), grupo["categoria"], desc, round(soma - desconto, 2), grupo["gid"]])
    registros = ws.get_all_records(value_render_option=ValueRenderOption.unformatted)
    chaves = {(d.isoformat(), ds) for d, ds, _ in alvo}
    linhas = [i for i, r in enumerate(registros, start=2)
              if str(r.get(COL_PARCELAMENTO, "")) == str(grupo["gid"])
              and (str(r["data"]), str(r["descricao"])) in chaves]
    for i in sorted(linhas, reverse=True):
        ws.delete_rows(i)
    recarregar_lancamentos()


# ── investimentos (planilha da carteira) ──

def _ws_carteira(nome, cabecalho=None):
    sh = get_client().open_by_key(CARTEIRA_SPREADSHEET_ID)
    if nome not in [w.title for w in sh.worksheets()]:
        if cabecalho is None:
            return None
        ws = sh.add_worksheet(title=nome, rows=200, cols=len(cabecalho))
        ws.append_row(cabecalho)
        return ws
    return sh.worksheet(nome)


def _num_br(x):
    """número vindo da planilha: float, '1.234,56', '0,0001' ou '12.5' → float (inválido → None)"""
    if isinstance(x, (int, float)):
        return float(x)
    t = str(x).strip().replace("R$", "").replace(" ", "")
    if not t:
        return None
    if "," in t:
        t = t.replace(".", "").replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def _data_carteira(x):
    """data da planilha da carteira: 'dd/mm/aaaa' (ou número de série) → date"""
    if isinstance(x, (int, float)):
        return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(x))).date()
    d = pd.to_datetime(str(x).strip(), format="%d/%m/%Y", errors="coerce")
    if pd.isna(d):
        d = pd.to_datetime(str(x).strip(), dayfirst=True, errors="coerce")
    return None if pd.isna(d) else d.date()


def _ler_invest_bruto(ws):
    """linhas da aba lancamentos da carteira: [(nº da linha, data, tipo, ativo, classe, qtd, preço, total)]"""
    vals = ws.get("A:G", value_render_option=ValueRenderOption.unformatted,
                  date_time_render_option=DateTimeOption.formatted_string)
    out = []
    for i, r in enumerate(vals[1:], start=2):
        r = (list(r) + [""] * 7)[:7]
        d, q, p, t = _data_carteira(r[0]), _num_br(r[4]), _num_br(r[5]), _num_br(r[6])
        if d is None or q is None:
            continue
        out.append((i, d, str(r[1]).strip().lower(), str(r[2]).strip(), str(r[3]).strip(),
                    q, p or 0.0, t if t is not None else q * (p or 0.0)))
    return out


def carregar_investimentos():
    if "df_invest" not in st.session_state:
        ws = _ws_carteira(ABA_INVEST, INVEST_HEADERS)
        df = pd.DataFrame(_ler_invest_bruto(ws),
                          columns=["linha", "data", "tipo", "ativo", "classe", "quantidade", "preco", "total"])
        st.session_state["df_invest"] = df
    return st.session_state["df_invest"]


def carregar_ativos():
    """[(ativo, classe)] para o formulário: só o que está em carteira (posição > 0), mais BTC,
    Renda+ 2050 e os ativos cadastrados que ainda não têm lançamento. quem foi vendido por
    inteiro sai da lista (volta por "+ novo ativo")"""
    if "ativos_invest" not in st.session_state:
        info = {}
        ws = _ws_carteira(ABA_ATIVOS)
        if ws is not None:
            for r in ws.get_all_values()[1:]:
                # linhas com ticker_atual são nomes antigos (de-para) e não entram na lista
                if r and r[0].strip() and len(r) > 1 and r[1].strip() and not (len(r) > 5 and r[5].strip()):
                    info.setdefault(r[0].strip().upper(), r[1].strip())
        df = carregar_investimentos()
        classes, posicao = {}, {}
        for _, r in df.iterrows():
            if not r["ativo"]:
                continue
            if r["classe"]:
                classes[r["ativo"]] = r["classe"]
            sinal = 1 if r["tipo"] == "compra" else (-1 if r["tipo"] == "venda" else 0)
            posicao[r["ativo"]] = posicao.get(r["ativo"], 0.0) + sinal * r["quantidade"]
        extras = st.session_state.get("ativos_extra", set())
        ativos = {a: classes.get(a) or info.get(a) for a, q in posicao.items() if q > 1e-9}
        ativos.update({a: c for a, c in info.items() if a not in posicao or a in extras})
        for a, c in ATIVOS_FIXOS:
            ativos.setdefault(a, c)
        ativos = {a: c for a, c in ativos.items() if c}
        ordem = {"ETF": 0, "FII": 1, "Cripto": 2, "Tesouro Direto": 3}
        st.session_state["ativos_invest"] = sorted(ativos.items(), key=lambda x: (ordem.get(x[1], 9), x[0]))
        st.session_state["ativos_todos"] = {**info, **classes}
    return st.session_state["ativos_invest"]


def _limpar_invest():
    st.session_state.pop("df_invest", None)
    st.session_state.pop("ativos_invest", None)


def _txt_num(v, casas):
    """float → texto pt-BR para a planilha (mesmo formato que o SmartWallet gravava)"""
    return f"{v:.{casas}f}".rstrip("0").rstrip(".").replace(".", ",")


def salvar_investimento(data_lanc, tipo, ativo, classe, qtd, preco):
    ws = _ws_carteira(ABA_INVEST, INVEST_HEADERS)
    ws.append_row([data_lanc.strftime("%d/%m/%Y"), tipo, ativo, classe,
                   _txt_num(qtd, 8), _txt_num(preco, 8), _txt_num(round(qtd * preco, 2), 2)],
                  value_input_option="USER_ENTERED")
    _limpar_invest()


def excluir_investimento(data_lanc, tipo, ativo, qtd, total):
    """exclusão por conteúdo: relê a aba e apaga a primeira linha igual"""
    ws = _ws_carteira(ABA_INVEST, INVEST_HEADERS)
    for linha, d, t, a, _, q, _, tot in _ler_invest_bruto(ws):
        if (d == data_lanc and t == tipo and a == ativo
                and abs(q - qtd) < 1e-9 and abs(tot - total) < 0.005):
            ws.delete_rows(linha)
            break
    _limpar_invest()


def salvar_ativo_novo(ativo, classe, tipo, indexador, pais):
    """cadastra (ou atualiza) o ativo na aba ativos_info da carteira — classificação usada pelo SmartWallet"""
    ws = _ws_carteira(ABA_ATIVOS, ["ativo", "classe", "tipo", "indexador", "pais", "ticker_atual"])
    nova = [ativo, classe, tipo or "", indexador or "", pais or ""]
    for i, r in enumerate(ws.get_all_values()[1:], start=2):
        if r and r[0].strip().upper() == ativo:
            ws.update(range_name=f"A{i}:E{i}", values=[nova])
            break
    else:
        ws.append_row(nova)
    # ativo que já teve lançamentos (vendido por inteiro) volta para a lista nesta sessão
    st.session_state["ativos_extra"] = st.session_state.get("ativos_extra", set()) | {ativo}
    _limpar_invest()


def atualizar_lancamento(orig, novo):
    """orig/novo: (data aaaa-mm-dd, categoria, descrição, valor) — a coluna parcelamento fica como está"""
    ws = get_spreadsheet().worksheet(ABA_LANCAMENTOS)
    registros = ws.get_all_records(value_render_option=ValueRenderOption.unformatted)
    for i, r in enumerate(registros, start=2):
        if (str(r["data"]) == str(orig[0]) and r["categoria"] == orig[1]
                and str(r["descricao"]) == str(orig[2]) and float(r["valor"]) == float(orig[3])):
            ws.update(range_name=f"A{i}:D{i}", values=[list(novo)])
            break
    recarregar_lancamentos()


def atualizar_pix(orig, novo):
    """orig/novo: (data aaaa-mm-dd, tipo, categoria, descrição, valor)"""
    ws = get_spreadsheet().worksheet(ABA_PIX)
    registros = ws.get_all_records(value_render_option=ValueRenderOption.unformatted)
    for i, r in enumerate(registros, start=2):
        if (str(r["data"]) == str(orig[0]) and str(r["tipo"]).strip().lower() == orig[1]
                and str(r["categoria"]) == str(orig[2]) and str(r["descricao"]) == str(orig[3])
                and float(r["valor"]) == float(orig[4])):
            ws.update(range_name=f"A{i}:E{i}", values=[list(novo)])
            break
    st.session_state.pop("df_pix", None)


def atualizar_nota(tipo, data_txt, valor, descricao, nova_data, novo_valor, nova_descricao):
    ws = get_client().open_by_key(CARTEIRA_SPREADSHEET_ID).worksheet(ABA_NOTAS)
    regs = ws.get_all_records(value_render_option=ValueRenderOption.unformatted)
    for i, r in enumerate(regs, start=2):
        if (str(r["tipo"]).strip().lower() == tipo and str(r["data"]) == str(data_txt)
                and abs(float(r["valor"] or 0) - float(valor)) < 0.005
                and str(r.get("descricao", "")) == str(descricao)):
            ws.update(range_name=f"A{i}:E{i}", values=[[
                tipo, nova_data.strftime("%d/%m/%Y"), float(novo_valor),
                "mês seguinte" if tipo == "nota fiscal" else "", nova_descricao]])
            break
    st.session_state.pop("df_notas", None)


def atualizar_investimento(orig, data_lanc, tipo, ativo, classe, qtd, preco, total):
    """orig: (data, tipo, ativo, quantidade, total) da linha a trocar"""
    ws = _ws_carteira(ABA_INVEST, INVEST_HEADERS)
    for linha, d, t, a, _, q, _, tot in _ler_invest_bruto(ws):
        if (d == orig[0] and t == orig[1] and a == orig[2]
                and abs(q - orig[3]) < 1e-9 and abs(tot - orig[4]) < 0.005):
            ws.update(range_name=f"A{linha}:G{linha}", values=[[
                data_lanc.strftime("%d/%m/%Y"), tipo, ativo, classe,
                _txt_num(qtd, 8), _txt_num(preco, 8), _txt_num(round(total, 2), 2)]],
                value_input_option="USER_ENTERED")
            break
    _limpar_invest()


def carregar_pix():
    """lançamentos de PIX (aba pix) — totalmente separados da fatura do cartão"""
    if "df_pix" not in st.session_state:
        sh = get_spreadsheet()
        garantir_abas(sh)
        ws = sh.worksheet(ABA_PIX)
        df = pd.DataFrame(ws.get_all_records(value_render_option=ValueRenderOption.unformatted))
        if df.empty:
            df = pd.DataFrame(columns=["data", "tipo", "categoria", "descricao", "valor"])
        else:
            df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0)
            df["data"] = pd.to_datetime(df["data"]).dt.date
            df["tipo"] = df["tipo"].astype(str).str.strip().str.lower()
        st.session_state["df_pix"] = df
    return st.session_state["df_pix"]


def salvar_pix(data_str, tipo, categoria, descricao, valor):
    ws = get_spreadsheet().worksheet(ABA_PIX)
    ws.append_row([data_str, tipo, categoria, descricao, valor])
    st.session_state.pop("df_pix", None)


def excluir_pix(data_str, tipo, categoria, descricao, valor):
    """exclusão por conteúdo, como nos lançamentos do cartão"""
    ws = get_spreadsheet().worksheet(ABA_PIX)
    registros = ws.get_all_records(value_render_option=ValueRenderOption.unformatted)
    for i, r in enumerate(registros, start=2):
        if (str(r["data"]) == str(data_str) and str(r["tipo"]).strip().lower() == tipo
                and str(r["categoria"]) == str(categoria) and str(r["descricao"]) == str(descricao)
                and float(r["valor"]) == float(valor)):
            ws.delete_rows(i)
            break
    st.session_state.pop("df_pix", None)


def carregar_notas():
    """notas fiscais e reembolsos (aba pj_notas da carteira) · None se a aba não existe"""
    if "df_notas" not in st.session_state:
        sh = get_client().open_by_key(CARTEIRA_SPREADSHEET_ID)
        if ABA_NOTAS not in [ws.title for ws in sh.worksheets()]:
            return None
        regs = sh.worksheet(ABA_NOTAS).get_all_records(value_render_option=ValueRenderOption.unformatted)
        df = pd.DataFrame(regs)
        if df.empty:
            df = pd.DataFrame(columns=["tipo", "data", "valor", "recebimento", "descricao"])
        else:
            df["data_txt"] = df["data"].astype(str)
            df["data"] = pd.to_datetime(df["data"].astype(str), dayfirst=True, errors="coerce").dt.date
            df["valor"] = pd.to_numeric(df["valor"], errors="coerce").fillna(0)
            df["tipo"] = df["tipo"].astype(str).str.strip().str.lower()
            df = df.dropna(subset=["data"])
        st.session_state["df_notas"] = df
    return st.session_state["df_notas"]


def salvar_nota(tipo, data_lanc, valor, descricao):
    """mesmo formato do SmartWallet: tipo, data dd/mm/aaaa, valor, recebimento, descrição"""
    ws = get_client().open_by_key(CARTEIRA_SPREADSHEET_ID).worksheet(ABA_NOTAS)
    ws.append_row([tipo, data_lanc.strftime("%d/%m/%Y"), float(valor),
                   "mês seguinte" if tipo == "nota fiscal" else "", descricao])
    st.session_state.pop("df_notas", None)


def excluir_nota(tipo, data_txt, valor, descricao):
    ws = get_client().open_by_key(CARTEIRA_SPREADSHEET_ID).worksheet(ABA_NOTAS)
    regs = ws.get_all_records(value_render_option=ValueRenderOption.unformatted)
    for i, r in enumerate(regs, start=2):
        if (str(r["tipo"]).strip().lower() == tipo and str(r["data"]) == str(data_txt)
                and abs(float(r["valor"] or 0) - float(valor)) < 0.005
                and str(r.get("descricao", "")) == str(descricao)):
            ws.delete_rows(i)
            break
    st.session_state.pop("df_notas", None)


# ── fluxo da PF (resumos das abas receitas e despesas) ──
# mesma conta do fluxo do SmartWallet: no mês k entram na PF o pró-labore líquido e os lucros
# das notas emitidas em k−1, os reembolsos, as entradas PIX e os dividendos; saem a fatura do
# cartão de k−1, o FIES e as saídas PIX. parâmetros lidos das abas pj_parametros /
# pj_parametros_hist / pf_cartao / dividendos_mensais da planilha da carteira.

PJ_PAR_PADRAO = {"faturamento_minimo": 0.0, "prev_pct": 0.0, "pro_labore": 0.0, "inss_pct": 11.0,
                 "das_pct": 6.0, "contador": 0.0, "fies": 0.0, "reembolso_base": 0.0,
                 "taxa_prefeitura": 0.0, "cartao_estimado": 0.0,
                 "fies_fim": "", "cartao_real_desde": ""}
PJ_PAR_TEXTO = {"fies_fim", "cartao_real_desde"}


def _ano_mes(s):
    """'12/2032' · '01/12/2032' · '2032-12-01' · '2032-12' → '2032-12' (ou None)"""
    s = str(s or "").strip()
    for fmt in ("%m/%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y-%m"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m")
        except ValueError:
            pass
    return None


def carregar_base_pf():
    """parâmetros PJ (atuais + histórico), faturas informadas no SmartWallet e dividendos por mês"""
    if "base_pf" not in st.session_state:
        sh = get_client().open_by_key(CARTEIRA_SPREADSHEET_ID)
        abas = {w.title: w for w in sh.worksheets()}

        def linhas(nome, rng):
            if nome not in abas:
                return []
            return abas[nome].get(rng, value_render_option=ValueRenderOption.unformatted)[1:]

        par = dict(PJ_PAR_PADRAO)
        for r in linhas("pj_parametros", "A:B"):
            if len(r) >= 2 and str(r[0]).strip() in par:
                k = str(r[0]).strip()
                v = str(r[1]).strip() if k in PJ_PAR_TEXTO else _num_br(r[1])
                if v is not None:
                    par[k] = v
        hist = []
        for r in linhas("pj_parametros_hist", "A:C"):
            if len(r) >= 3 and str(r[0]).strip():
                d = str(r[1]).strip()
                d = "0000-00" if d in ("", "0000-00", "início", "inicio") else _ano_mes(d)
                v = _num_br(r[2])
                if d and v is not None:
                    hist.append({"chave": str(r[0]).strip(), "desde": d, "valor": v})
        hist.sort(key=lambda e: (e["chave"], e["desde"]))
        mensal = lambda nome: {str(r[0]).strip(): _num_br(r[1]) for r in linhas(nome, "A:B")
                               if len(r) >= 2 and str(r[0]).strip() and _num_br(r[1]) is not None}
        st.session_state["base_pf"] = {"par": par, "hist": hist, "cartao": mensal("pf_cartao"),
                                       "divs": mensal("dividendos_mensais")}
    return st.session_state["base_pf"]


def par_vigente(par, hist, ano_mes):
    """parâmetros valendo no mês: último valor do histórico com desde <= mês"""
    out = dict(par)
    for k in {e["chave"] for e in hist}:
        regs = [e for e in hist if e["chave"] == k]
        validos = [e for e in regs if e["desde"] <= ano_mes]
        out[k] = (validos[-1] if validos else regs[0])["valor"]
    return out


def cartao_do_mes(ref):
    """gastos no cartão de um mês (lancamentos + abas arquivo_AAAA deste app)"""
    if "cartao_arquivo" not in st.session_state:
        tot = {}
        for ws in get_spreadsheet().worksheets():
            if ws.title.startswith("arquivo_"):
                for r in ws.get_all_records(value_render_option=ValueRenderOption.unformatted):
                    d, v = pd.to_datetime(str(r.get("data", "")), errors="coerce"), _num_br(r.get("valor", ""))
                    if not pd.isna(d) and v is not None:
                        tot[d.strftime("%Y-%m")] = tot.get(d.strftime("%Y-%m"), 0.0) + v
        st.session_state["cartao_arquivo"] = tot
    df = filtrar_mes(carregar_lancamentos(), ref.year, ref.month)
    return (float(df["valor"].sum()) if not df.empty else 0.0) \
        + st.session_state["cartao_arquivo"].get(ref.strftime("%Y-%m"), 0.0)


def fluxo_pf(ref, df_notas):
    """entradas e saídas da conta PF no mês ref: {'entradas': [(rótulo, valor)], 'saidas': [...],
    'tem_div': dividendos do mês já gravados}"""
    base = carregar_base_pf()
    k, ant = ref.strftime("%Y-%m"), mes_anterior(ref)
    a = ant.strftime("%Y-%m")
    par = par_vigente(base["par"], base["hist"], k)

    def _n(m, t):
        if df_notas is None or df_notas.empty:
            return 0.0
        d = filtrar_mes(df_notas, m.year, m.month)
        return float(d[d["tipo"] == t]["valor"].sum()) if not d.empty else 0.0

    fat = _n(ant, "nota fiscal")                       # notas de k−1, recebidas (e com DAS pago) em k
    prev = par["faturamento_minimo"] * par["prev_pct"] / 100
    das = fat * par["das_pct"] / 100
    taxa = par["taxa_prefeitura"] if ref.month == 2 else 0.0
    lucros = fat - prev - das - par["contador"] - taxa - par["pro_labore"]
    pl_liq = par["pro_labore"] * (1 - par["inss_pct"] / 100)
    fim = _ano_mes(par.get("fies_fim"))
    fies = min(par["fies"], max(lucros, 0.0)) if (fim is None or k <= fim) else 0.0
    # fatura paga em k = gastos de k−1 (valor informado no SmartWallet > média estimada > real)
    desde = _ano_mes(par.get("cartao_real_desde")) or "0000-00"
    if a in base["cartao"]:
        cartao = base["cartao"][a]
    elif a < desde:
        cartao = par["cartao_estimado"]
    else:
        cartao = cartao_do_mes(ant)
    pix = filtrar_mes(carregar_pix(), ref.year, ref.month)
    _px = lambda t: float(pix[pix["tipo"] == t]["valor"].sum()) if not pix.empty else 0.0
    return {
        "entradas": [("lucros distribuídos", lucros), ("pró-labore líquido", pl_liq),
                     ("reembolsos", par["reembolso_base"] + _n(ref, "reembolso")),
                     ("outras receitas PIX", _px("entrada")), ("dividendos", base["divs"].get(k, 0.0))],
        "saidas": [(f"fatura de {mes_label(ant)}", cartao), ("FIES", fies),
                   ("outras despesas PIX", _px("saida"))],
        "tem_div": k in base["divs"],
    }


def render_resumo(rotulo, total, itens, cor_total=None):
    """painel no estilo do resumo da carteira (SmartWallet › mês › resumo): total em cima, lista embaixo.
    itens: [(rótulo, valor, classe)] · classe 'pos' (verde), 'neg' (vermelho) ou '' (cor padrão)"""
    brl = lambda v: ("R$ " + f"{abs(v):.0f}").replace("$", "&#36;")
    linhas = "".join(f"<tr><td>{r}</td><td class='{c}'>{'−' if v < 0 else ''}{brl(v)}</td></tr>"
                     for r, v, c in itens) \
        or "<tr><td style='color:#888;'>nada neste mês</td></tr>"
    st.markdown(
        "<div style='border:1px solid #202024;border-radius:12px;padding:1rem 1.1rem 0.6rem;'>"
        f"<p style='font-size:0.85rem;color:rgba(250,250,250,0.55);margin:0 0 2px 0;'>{rotulo}</p>"
        f"<p style='font-size:1.6rem;font-weight:600;line-height:1.2;margin:0 0 0.9rem 0;"
        f"color:{cor_total or 'inherit'};'>{'−' if total < 0 else ''}{brl(total)}</p>"
        f"<table class='pf-tab'>{linhas}</table></div>", unsafe_allow_html=True)


def render_lista_pf(rotulo, lista, cor):
    """lista da PF do maior para o menor, sem os zerados; valor negativo fica com a cor oposta"""
    oposta = {"pos": "neg", "neg": "pos"}[cor]
    itens = sorted([(r, v, cor if v >= 0 else oposta) for r, v in lista if abs(v) >= 0.005], key=lambda x: -x[1])
    render_resumo(rotulo, sum(v for _, v, _ in itens), itens)


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


def fmt_brl2(valor):
    """com centavos, pt-BR: R$ 1.234,56"""
    return "R$ " + f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_qtd(q):
    return f"{q:.8f}".rstrip("0").rstrip(".").replace(".", ",")


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
    hoje = hoje_br()
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
    hoje = hoje_br().replace(day=1)
    if candidato <= hoje:
        st.session_state["mes_ref"] = candidato


def ativar_campos_formulario(aria_label_valor):
    """Ajustes de entrada via JS (o Streamlit não expõe isso nativamente):
    - seletor de categoria: não abre o teclado do celular (inputmode none);
    - campo de valor: teclado numérico e máscara de centavos (digita só números,
      a vírgula fica fixa com 2 casas: 1 -> 0,01 | 1250 -> 12,50)."""
    components.html(f"""
    <script>
    const doc = window.parent.document;
    function formatarCentavos(txt) {{
        const dig = txt.replace(/\\D/g, '');
        if (!dig) return '';
        const c = parseInt(dig, 10);
        return Math.floor(c / 100) + ',' + String(c % 100).padStart(2, '0');
    }}
    function ajustar() {{
        doc.querySelectorAll('div[data-testid="stSelectbox"] input').forEach(el => {{
            el.setAttribute('inputmode', 'none');
        }});
        doc.querySelectorAll('input[aria-label="{aria_label_valor}"]').forEach(el => {{
            el.setAttribute('inputmode', 'numeric');
            if (el.dataset.mascara) return;
            el.dataset.mascara = '1';
            el.addEventListener('input', () => {{
                const novo = formatarCentavos(el.value);
                if (novo === el.value) return;
                const setter = Object.getOwnPropertyDescriptor(
                    window.parent.HTMLInputElement.prototype, 'value').set;
                setter.call(el, novo);
                el.dispatchEvent(new Event('input', {{ bubbles: true }}));
            }});
        }});
    }}
    ajustar();
    new MutationObserver(ajustar).observe(doc.body, {{ childList: true, subtree: true }});
    </script>
    """, height=0)


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

def _html_brl(v):
    """R$ formatado com o cifrão como entidade HTML (evita o Markdown ler dois '$' como LaTeX)."""
    return formatar_brl(v).replace("$", "&#36;")


ICONE_CARTAO = (
    "<svg width='18' height='18' viewBox='0 0 24 24' fill='none' stroke='#9a9a9a' stroke-width='1.8' "
    "stroke-linecap='round' stroke-linejoin='round'><rect x='2' y='5' width='20' height='14' rx='2.5'/>"
    "<line x1='2' y1='10' x2='22' y2='10'/><line x1='6' y1='15' x2='10' y2='15'/></svg>"
)
ICONE_CALENDARIO = (
    "<svg width='18' height='18' viewBox='0 0 24 24' fill='none' stroke='#9a9a9a' stroke-width='1.8' "
    "stroke-linecap='round' stroke-linejoin='round'><rect x='3' y='4.5' width='18' height='16.5' rx='2.5'/>"
    "<line x1='3' y1='10' x2='21' y2='10'/><line x1='8' y1='2.5' x2='8' y2='6.5'/>"
    "<line x1='16' y1='2.5' x2='16' y2='6.5'/></svg>"
)


def _barra_html(titulo, direita, pct, cor, marcador=None):
    """Barra de progresso própria: rótulo à esquerda, números à direita, marcador opcional."""
    largura = min(max(pct, 0), 1) * 100
    marca = ""
    return (
        "<div style='margin-top:16px;'>"
        "<div style='display:flex;justify-content:space-between;align-items:center;"
        "font-size:12px;color:#9a9a9a;margin-bottom:6px;'>"
        f"<span style='display:flex;align-items:center;'>{titulo}</span>"
        f"<span>{direita}</span></div>"
        "<div style='position:relative;height:8px;background:#202024;border-radius:6px;'>"
        f"<div style='width:{largura:.1f}%;height:100%;background:{cor};border-radius:6px;'></div>"
        f"{marca}</div></div>"
    )


def render_cards_limite(df_mes, limite_mensal, mes_ref=None):
    """Painel da aba lançar: total da fatura / limite disponível + barras de fatura e de mês."""
    import calendar
    total_gasto = df_mes["valor"].sum() if not df_mes.empty else 0
    disponivel = limite_mensal - total_gasto

    hoje = hoje_br()
    dias_no_mes = calendar.monthrange(hoje.year, hoje.month)[1]
    pct_mes = hoje.day / dias_no_mes
    if mes_ref is not None and (mes_ref.year, mes_ref.month) < (hoje.year, hoje.month):
        pct_mes = 1.0
    pct_fatura = total_gasto / limite_mensal if limite_mensal > 0 else 0

    # cor/estado da fatura comparando com o ritmo do mês
    if pct_fatura > 1:
        cor_fatura, status = "#e05252", "limite excedido"
    elif pct_fatura > pct_mes:
        cor_fatura, status = "#E5B800", "acima do ritmo do mês"
    else:
        cor_fatura, status = "#6b7280", "dentro do ritmo do mês"

    label_disp = "limite disponível (R$)" if disponivel >= 0 else "limite excedido (R$)"
    cor_disp = "inherit" if disponivel >= 0 else "#e05252"

    html = (
        "<div style='border:1px solid #202024;border-radius:12px;padding:1rem 1.1rem 1.1rem;'>"
        "<div style='display:grid;grid-template-columns:1fr 1fr;align-items:end;column-gap:12px;'>"
        "<div style='justify-self:start;text-align:center;'>"
        "<div style='font-size:12px;color:rgba(250,250,250,0.55);margin-bottom:2px;'>total da fatura (R&#36;)</div>"
        f"<div style='font-size:1.6rem;font-weight:600;line-height:1.15;'>{total_gasto:.0f}</div>"
        "</div>"
        "<div style='justify-self:end;text-align:center;'>"
        f"<div style='font-size:12px;color:rgba(250,250,250,0.55);margin-bottom:2px;'>{label_disp.replace('$', '&#36;')}</div>"
        f"<div style='font-size:1.6rem;font-weight:600;line-height:1.15;color:{cor_disp};'>{abs(disponivel):.0f}</div>"
        "</div></div>"
        + _barra_html(
            ICONE_CARTAO,
            f"<span style='color:{cor_fatura};'>{fmt_pct(pct_fatura * 100)} consumido</span>",
            pct_fatura, cor_fatura,
        )
        + _barra_html(
            ICONE_CALENDARIO,
            f"{fmt_pct(pct_mes * 100)} decorrido",
            pct_mes, "#6b7280",
        )
        + "</div>"
    )
    st.markdown(html, unsafe_allow_html=True)


def render_barras_categorias(df_mes, limite_mensal, df_cat):
    """barras de progresso por categoria: gasto / orçamento"""
    gasto_por_cat = df_mes.groupby("categoria")["valor"].sum().to_dict() if not df_mes.empty else {}
    cats_ordenadas = df_cat[df_cat["id"] != FIXED_ID].copy()
    cats_ordenadas["gasto"] = cats_ordenadas["id"].map(gasto_por_cat).fillna(0)
    cats_ordenadas = cats_ordenadas.sort_values("gasto", ascending=False)

    if cats_ordenadas.empty:
        st.caption("nenhuma categoria configurada.")
        return
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


def linha_item(item_id, titulo, sub, valor_html, ao_editar, ao_excluir, pergunta):
    """linha do histórico: título + detalhe · ✎ editar · ✕ excluir (com confirmação) · valor"""
    pendente_key = f"pendente_del_{item_id}"
    col_desc, col_ed, col_del, col_val = st.columns([3.0, 0.8, 0.8, 1.7])
    with col_desc:
        st.markdown(f"**{titulo}**  \n<span style='color:#888;font-size:12px;'>{sub}</span>".replace("$", "&#36;"),
                    unsafe_allow_html=True)
    with col_ed:
        if st.button("✎", key=f"ed_{item_id}", use_container_width=True):
            ao_editar()
    with col_del:
        if not st.session_state.get(pendente_key):
            if st.button("✕", key=f"del_{item_id}", use_container_width=True):
                st.session_state[pendente_key] = True
                st.rerun()
    with col_val:
        st.markdown(f"<div style='text-align:right;padding-top:0.5rem;'>{valor_html}</div>", unsafe_allow_html=True)
    if st.session_state.get(pendente_key):
        st.markdown(f"<div style='font-size:12px;color:#888;margin-top:-4px;margin-bottom:4px;'>{pergunta}</div>",
                    unsafe_allow_html=True)
        col_confirma, col_cancela = st.columns(2)
        with col_confirma:
            if st.button("confirmar exclusão", key=f"ok_{item_id}", use_container_width=True, type="primary"):
                ao_excluir()
                st.session_state.pop(pendente_key, None)
                st.rerun()
        with col_cancela:
            if st.button("cancelar", key=f"cancel_{item_id}", use_container_width=True):
                st.session_state.pop(pendente_key, None)
                st.rerun()


def _txt_valor(v):
    """valor para pré-preencher o campo com máscara de centavos"""
    return f"{v:.2f}".replace(".", ",")


def _fechar(msg):
    st.session_state["_msg"] = msg
    st.rerun()


# ───────────────────────── FORMULÁRIOS (janelas) ─────────────────────────
# cada formulário abre numa janela por cima da página; com item=... vem preenchido para edição.

@st.dialog("despesa")
def dialogo_despesa(item=None):
    todas_cats = df_cat.copy()
    todas_cats["nome_lower"] = todas_cats["nome"].str.lower()
    todas_cats = todas_cats.sort_values("nome_lower")
    opcoes_ids = todas_cats["id"].tolist()
    opcoes_labels = dict(zip(todas_cats["id"], todas_cats["nome_lower"]))
    if item:
        index_default = opcoes_ids.index(item["categoria"]) if item["categoria"] in opcoes_ids else 0
    else:
        # padrão: a categoria com mais gasto no mês anterior
        mes_ant = mes_anterior(hoje_br())
        df_mes_ant = filtrar_mes(carregar_lancamentos(), mes_ant.year, mes_ant.month)
        gasto_mes_ant = df_mes_ant.groupby("categoria")["valor"].sum().to_dict() if not df_mes_ant.empty else {}
        categoria_default = max(gasto_mes_ant, key=gasto_mes_ant.get) if gasto_mes_ant else None
        index_default = opcoes_ids.index(categoria_default) if categoria_default in opcoes_ids else 0

    with st.form("form_despesa_edit" if item else "form_despesa", border=False):
        if item:
            st.caption(f"pagamento: {item['pagamento']}")
            categoria_id = st.selectbox("categoria", options=opcoes_ids, index=index_default,
                                        format_func=lambda x: opcoes_labels[x])
            pagamento, n_parcelas = item["pagamento"], 1
        else:
            col_pag, col_cat = st.columns([1, 1.6])
            with col_pag:
                pagamento = st.selectbox("pagamento", options=["cartão", "pix"], index=0)
            with col_cat:
                categoria_id = st.selectbox("categoria", options=opcoes_ids, index=index_default,
                                            format_func=lambda x: opcoes_labels[x])
        col_data, col_val = st.columns([1, 1])
        with col_data:
            data_lanc = st.date_input("data", value=item["data"] if item else hoje_br(), format="DD/MM/YYYY")
        with col_val:
            valor_txt = st.text_input("valor (R$)", value=_txt_valor(item["valor"]) if item else "",
                                      placeholder="0,00")
        if item:
            descricao = st.text_input("descrição", value=item["descricao"])
        else:
            col_desc, col_parc = st.columns([1.6, 1])
            with col_desc:
                descricao = st.text_input("descrição")
            with col_parc:
                n_parcelas = st.selectbox("parcelas", options=list(range(1, MAX_PARCELAS + 1)), index=0,
                                          format_func=lambda n: "à vista" if n == 1 else f"{n}x")
        ativar_campos_formulario("valor (R$)")
        enviado = st.form_submit_button("salvar" if item else "registrar despesa", use_container_width=True)

        if enviado:
            valor = parse_valor(valor_txt)
            if valor <= 0:
                st.error("informe um valor válido.")
            elif item and pagamento == "pix":
                atualizar_pix((item["data"].isoformat(), "saida", item["categoria"], item["descricao"], item["valor"]),
                              (data_lanc.isoformat(), "saida", categoria_id, descricao, valor))
                _fechar("alterado!")
            elif item:
                atualizar_lancamento((item["data"].isoformat(), item["categoria"], item["descricao"], item["valor"]),
                                     (data_lanc.isoformat(), categoria_id, descricao, valor))
                _fechar("alterado!")
            elif n_parcelas > 1 and pagamento == "pix":
                st.error("parcelamento só no cartão.")
            elif n_parcelas > 1:
                # valor = total da compra; cada parcela cai na fatura do seu mês
                salvar_parcelado(data_lanc, categoria_id, descricao.strip() or opcoes_labels[categoria_id],
                                 valor, n_parcelas)
                _fechar("registrado!")
            elif pagamento == "pix":
                # pix: aba pix (fora da fatura do cartão)
                salvar_pix(data_lanc.isoformat(), "saida", categoria_id, descricao, valor)
                _fechar("registrado!")
            else:
                salvar_lancamento(data_lanc.isoformat(), categoria_id, descricao, valor)
                _fechar("registrado!")


@st.dialog("adiantar parcelas")
def dialogo_adiantar(grupo):
    fut = grupo["futuras"]
    st.markdown(f"**{grupo['nome']}** · restam {len(fut)} de {grupo['n']}")
    k = st.selectbox("parcelas a adiantar", options=list(range(len(fut), 0, -1)), index=0,
                     format_func=lambda n: (f"todas ({n})" if n == len(fut) else
                                            ("a última" if n == 1 else f"as {n} últimas")))
    desconto_txt = st.text_input("desconto (R$)", placeholder="0,00")
    ativar_campos_formulario("desconto (R$)")
    soma = round(sum(v for _, _, v in fut[-k:]), 2)
    desconto = parse_valor(desconto_txt)
    st.markdown(
        f"<div style='font-size:13px;color:#888;'>parcelas: {fmt_brl2(soma)} · desconto: {fmt_brl2(desconto)}</div>"
        f"<div style='font-size:1.3rem;font-weight:600;margin:2px 0 10px;'>a pagar hoje: "
        f"{fmt_brl2(max(soma - desconto, 0))}</div>".replace("$", "&#36;"), unsafe_allow_html=True)
    st.caption("as parcelas adiantadas saem dos meses futuros e o valor com desconto entra na fatura deste mês.")
    if st.button("adiantar", type="primary", use_container_width=True):
        if desconto < 0 or desconto >= soma:
            st.error("o desconto precisa ser menor que o valor das parcelas.")
        else:
            adiantar_parcelas(grupo, k, desconto)
            _fechar("parcelas adiantadas!")


@st.dialog("receita")
def dialogo_receita(item=None, erro_notas=None):
    """nota fiscal e reembolso → aba pj_notas da carteira; pix → aba pix (entrada), fora da fatura"""
    with st.form("form_receita_edit" if item else "form_receita", border=False):
        if item:
            st.caption(f"tipo: {item['tipo']}")
            tipo = item["tipo"]
        else:
            tipo = st.selectbox("tipo", options=RECEITA_TIPOS, index=0)
        col_data, col_val = st.columns([1, 1])
        with col_data:
            data_lanc = st.date_input("data", value=item["data"] if item else hoje_br(), format="DD/MM/YYYY")
        with col_val:
            valor_txt = st.text_input("valor (R$)", value=_txt_valor(item["valor"]) if item else "",
                                      placeholder="0,00")
        descricao = st.text_input("descrição", value=item["descricao"] if item else "")
        ativar_campos_formulario("valor (R$)")
        enviado = st.form_submit_button("salvar" if item else "registrar receita", use_container_width=True)
        if enviado:
            valor = parse_valor(valor_txt)
            if valor <= 0:
                st.error("informe um valor válido.")
            elif tipo == "pix" and item:
                atualizar_pix((item["data"].isoformat(), "entrada", item["categoria"], item["descricao"], item["valor"]),
                              (data_lanc.isoformat(), "entrada", item["categoria"], descricao, valor))
                _fechar("alterado!")
            elif tipo == "pix":
                salvar_pix(data_lanc.isoformat(), "entrada", "receita", descricao, valor)
                _fechar("registrado!")
            elif erro_notas:
                st.error(erro_notas)
            elif item:
                atualizar_nota(tipo, item["data_txt"], item["valor"], item["descricao"], data_lanc, valor, descricao)
                _fechar("alterado!")
            else:
                salvar_nota(tipo, data_lanc, valor, descricao)
                if tipo == "nota fiscal":       # nota é recebida (e paga DAS) no mês seguinte
                    cx = date(data_lanc.year + (data_lanc.month == 12), data_lanc.month % 12 + 1, 1)
                    _fechar(f"registrada! entra no caixa de {MESES3[cx.month - 1]}/{str(cx.year)[2:]}")
                _fechar("registrado!")


@st.dialog("aporte")
def dialogo_investimento(df_inv, ativos_inv, item=None):
    ativos_inv = list(ativos_inv)
    if item and item["ativo"] not in dict(ativos_inv):
        ativos_inv.append((item["ativo"], item.get("classe") or
                           st.session_state.get("ativos_todos", {}).get(item["ativo"], "")))
    nomes = [a for a, _ in ativos_inv]
    classes = dict(ativos_inv)
    preco_ini = _txt_valor(item["preco"]) if item else ""
    with st.form("form_invest_edit" if item else "form_invest", border=False):
        col_tipo, col_ativo = st.columns([1, 1.6])
        with col_tipo:
            # lançamentos de outro tipo já existentes na planilha (ex.: dividendo manual) mantêm o tipo
            _tipos = ["compra", "venda"] + ([item["tipo"]] if item and item["tipo"] not in ("compra", "venda") else [])
            tipo = st.selectbox("tipo", options=_tipos, index=_tipos.index(item["tipo"]) if item else 0)
        with col_ativo:
            ativo = st.selectbox("ativo", options=nomes,
                                 index=(nomes.index(item["ativo"]) if item and item["ativo"] in nomes
                                        else (0 if nomes else None)))
        col_data, col_qtd = st.columns([1, 1])
        with col_data:
            data_lanc = st.date_input("data", value=item["data"] if item else hoje_br(),
                                      max_value=hoje_br(), format="DD/MM/YYYY")
        with col_qtd:
            qtd_txt = st.text_input("quantidade", value=fmt_qtd(item["quantidade"]) if item else "",
                                    placeholder="ex.: 10")
        preco_txt = st.text_input("preço unitário (R$)", value=preco_ini, placeholder="0,00")
        ativar_campos_formulario("preço unitário (R$)")
        enviado = st.form_submit_button("salvar" if item else "registrar aporte", use_container_width=True)
        if enviado:
            qtd = _num_br(qtd_txt) or 0.0
            # preço não mexido na edição: mantém o original (pode ter mais de 2 casas)
            preco = item["preco"] if item and preco_txt == preco_ini else parse_valor(preco_txt)
            posicao = 0.0
            if ativo and not df_inv.empty:
                _a = df_inv[df_inv["ativo"] == ativo]
                posicao = float(_a[_a["tipo"] == "compra"]["quantidade"].sum()
                                - _a[_a["tipo"] == "venda"]["quantidade"].sum())
                if item and item["ativo"] == ativo:      # desconta o próprio lançamento em edição
                    posicao -= item["quantidade"] * (1 if item["tipo"] == "compra" else -1)
            if not ativo:
                st.error("escolha um ativo.")
            elif qtd <= 0 or preco <= 0:
                st.error("informe quantidade e preço.")
            elif tipo == "venda" and qtd > posicao + 1e-9:
                st.error(f"venda maior que a posição ({fmt_qtd(max(posicao, 0))} de {ativo}).")
            elif item:
                atualizar_investimento((item["data"], item["tipo"], item["ativo"], item["quantidade"], item["total"]),
                                       data_lanc, tipo, ativo, classes[ativo], qtd, preco, qtd * preco)
                _fechar("alterado!")
            else:
                salvar_investimento(data_lanc, tipo, ativo, classes[ativo], qtd, preco)
                _fechar("registrado!")


# classificação do ativo novo num campo só: (rótulo, classe, tipo, indexador, país)
TIPOS_ATIVO = ([("FII de tijolo", "FII", "tijolo", None, None),
                ("FII de papel · CDI", "FII", "papel", "CDI", None),
                ("FII de papel · IPCA", "FII", "papel", "IPCA", None)]
               + [(f"ETF · {p}", "ETF", None, None, p) for p in PAISES_ETF]
               + [("ETF de renda fixa · CDI", "ETF", "renda fixa", "CDI", "Brasil"),
                  ("ETF de renda fixa · IPCA", "ETF", "renda fixa", "IPCA", "Brasil")])


@st.dialog("novo ativo")
def dialogo_novo_ativo(ativos_inv):
    with st.form("form_ativo", border=False):
        col_tk, col_tp = st.columns([1, 1.6])
        with col_tk:
            ticker = st.text_input("ticker", placeholder="ex.: ABCD11")
        with col_tp:
            i_tipo = st.selectbox("tipo", options=list(range(len(TIPOS_ATIVO))), index=0,
                                  format_func=lambda i: TIPOS_ATIVO[i][0])
        enviado = st.form_submit_button("cadastrar ativo", use_container_width=True)
        if enviado:
            ticker = (ticker or "").strip().upper()
            _, classe, tipo, indexador, pais = TIPOS_ATIVO[i_tipo]
            if not re.fullmatch(r"[A-Z]{4}\d{1,2}", ticker):
                st.error("ticker inválido — use o formato da B3 (ex.: ABCD11).")
            elif ticker in dict(ativos_inv):
                st.error(f"{ticker} já está na lista.")
            else:
                salvar_ativo_novo(ticker, classe, tipo, indexador, pais)
                _fechar(f"{ticker} cadastrado!")


if st.session_state.get("_msg"):
    st.toast(st.session_state.pop("_msg"))

aba_lancar, aba_receitas, aba_invest, aba_orc = st.tabs(
    ["despesas", "receitas", "investimentos", "orçamento"]
)

# ───────────────────────── ABA: DESPESAS ─────────────────────────
# painel da fatura · botão do formulário · barras por categoria · parcelamentos · histórico

with aba_lancar:
    render_nav_mes("desp")
    ref = st.session_state["mes_ref"]
    df_mes = filtrar_mes(carregar_lancamentos(), ref.year, ref.month)
    render_cards_limite(df_mes, limite_mensal, ref)
    st.write("")
    # resumo: o que saiu da conta PF no mês (fatura do mês anterior, FIES e outras despesas PIX)
    try:
        render_lista_pf("pago pela PF", fluxo_pf(ref, carregar_notas())["saidas"], "neg")
    except Exception as e:
        st.warning(f"não foi possível montar o resumo da PF ({str(e)[:60]})")
    st.write("")
    with st.container(key="novo_despesa"):
        if st.button("+ nova despesa", use_container_width=True):
            dialogo_despesa()
    st.write("")

    render_barras_categorias(df_mes, limite_mensal, df_cat)

    abertos = parcelamentos_abertos(carregar_lancamentos())
    if abertos:
        st.write("")
        st.markdown("<div style='font-size:12px;color:#888;margin-bottom:4px;'>parcelamentos em aberto</div>",
                    unsafe_allow_html=True)
        for g in abertos:
            fut = g["futuras"]
            col_desc, col_acao, col_val = st.columns([3.2, 1.7, 1.6])
            with col_desc:
                st.markdown(
                    f"**{g['nome']}**  \n<span style='color:#888;font-size:12px;'>restam {len(fut)} de {g['n']} · "
                    f"{fmt_brl2(fut[0][2])}/mês · até {mes_label(fut[-1][0])}</span>".replace("$", "&#36;"),
                    unsafe_allow_html=True)
            with col_acao:
                if st.button("adiantar", key=f"adi_{g['gid']}", use_container_width=True):
                    dialogo_adiantar(g)
            with col_val:
                st.markdown(f"<div style='text-align:right;padding-top:0.5rem;'>"
                            f"{formatar_brl(sum(v for _, _, v in fut))}</div>", unsafe_allow_html=True)

    # histórico do mês: cartão + pix na mesma lista (o pix fica marcado e fora dos totais da fatura)
    df_pix_h = filtrar_mes(carregar_pix(), ref.year, ref.month)
    df_pix_h = df_pix_h[df_pix_h["tipo"] == "saida"] if not df_pix_h.empty else df_pix_h
    st.write("")
    if df_mes.empty and df_pix_h.empty:
        st.caption("nenhuma despesa neste mês.")
    else:
        # inverte a ordem original (último inserido primeiro) e ordena por data de forma estável
        _cart_h = df_mes.iloc[::-1].assign(pagamento="cartão")
        _pix_h = (df_pix_h.iloc[::-1][["data", "categoria", "descricao", "valor"]].assign(pagamento="pix")
                  if not df_pix_h.empty else df_pix_h)
        df_mes_ordenado = pd.concat([x for x in [_cart_h, _pix_h] if not x.empty], ignore_index=True) \
            .sort_values("data", ascending=False, kind="mergesort")
        mapa_nomes = dict(zip(df_cat["id"], df_cat["nome"]))
        for idx_row, (_, r) in enumerate(df_mes_ordenado.iterrows()):
            nome_cat = mapa_nomes.get(r["categoria"], "sem categoria").lower()
            it = {"pagamento": r["pagamento"], "data": r["data"], "categoria": r["categoria"],
                  "descricao": str(r["descricao"] or ""), "valor": float(r["valor"])}

            def _excluir(it=it):
                if it["pagamento"] == "pix":
                    excluir_pix(it["data"].isoformat(), "saida", it["categoria"], it["descricao"], it["valor"])
                else:
                    excluir_lancamento(it["data"].isoformat(), it["categoria"], it["descricao"], it["valor"])

            linha_item(
                f"d{idx_row}_{it['data']}_{it['categoria']}_{it['valor']}_{it['pagamento']}",
                it["descricao"] or nome_cat,
                f"{nome_cat}{' · pix' if it['pagamento'] == 'pix' else ''} · {it['data'].strftime('%d/%m/%y')}",
                f"− {formatar_brl(it['valor'])}",
                lambda it=it: dialogo_despesa(it), _excluir, "excluir este lançamento?")

        st.write("")
        csv = df_mes_ordenado[["data", "categoria", "descricao", "valor", "pagamento"]].rename(
            columns={"descricao": "descrição", "valor": "valor (R$)"}).to_csv(index=False).encode("utf-8")
        st.download_button("exportar csv", csv, f"fatura_{mes_label(ref).replace('/', '_')}.csv", "text/csv")

# ───────────────────────── ABA: RECEITAS ─────────────────────────

with aba_receitas:
    render_nav_mes("rec")
    ref_rec = st.session_state["mes_ref"]
    try:
        df_notas = carregar_notas()
        erro_notas = None if df_notas is not None else "abra o SmartWallet uma vez para criar a aba pj_notas."
    except Exception as e:
        df_notas, erro_notas = None, f"planilha da carteira indisponível ({str(e)[:60]})"
    df_notas_mes = (filtrar_mes(df_notas, ref_rec.year, ref_rec.month)
                    if df_notas is not None and not df_notas.empty else pd.DataFrame())
    df_pix_rec = filtrar_mes(carregar_pix(), ref_rec.year, ref_rec.month)
    df_pix_rec = df_pix_rec[df_pix_rec["tipo"] == "entrada"] if not df_pix_rec.empty else df_pix_rec

    # resumo: só o que entrou na conta PF no mês (o histórico abaixo mostra as notas brutas)
    if erro_notas:
        st.warning(erro_notas)
    else:
        try:
            _fx = fluxo_pf(ref_rec, df_notas)
            render_lista_pf("recebido na PF", _fx["entradas"], "pos")
            if not _fx["tem_div"] and ref_rec == hoje_br().replace(day=1):
                st.caption("dividendos do mês entram no resumo quando o mês fecha.")
        except Exception as e:
            st.warning(f"não foi possível montar o resumo da PF ({str(e)[:60]})")
    st.write("")
    with st.container(key="novo_receita"):
        if st.button("+ nova receita", use_container_width=True):
            dialogo_receita(erro_notas=erro_notas)
    st.write("")

    itens = []
    if not df_notas_mes.empty:
        for _, r in df_notas_mes.iterrows():
            itens.append({"origem": "nota", "tipo": r["tipo"], "data": r["data"], "valor": float(r["valor"]),
                          "descricao": str(r.get("descricao", "") or ""), "data_txt": r["data_txt"]})
    if not df_pix_rec.empty:
        for _, r in df_pix_rec.iloc[::-1].iterrows():
            itens.append({"origem": "pix", "tipo": "pix", "data": r["data"], "valor": float(r["valor"]),
                          "descricao": str(r.get("descricao", "") or ""), "categoria": r["categoria"]})
    if not itens:
        st.caption("nenhuma receita neste mês.")
    for idx_row, it in enumerate(sorted(itens, key=lambda x: x["data"], reverse=True)):
        def _excluir_rec(it=it):
            if it["origem"] == "pix":
                excluir_pix(it["data"].isoformat(), "entrada", it["categoria"], it["descricao"], it["valor"])
            else:
                excluir_nota(it["tipo"], it["data_txt"], it["valor"], it["descricao"])

        linha_item(
            f"r{idx_row}_{it['data']}_{it['tipo']}_{it['valor']}",
            it["descricao"] or it["tipo"], f"{it['tipo']} · {it['data'].strftime('%d/%m/%y')}",
            f"<span style='color:#1D9E75;'>+ {formatar_brl(it['valor'])}</span>",
            lambda it=it: dialogo_receita(it, erro_notas), _excluir_rec, "excluir esta receita?")

# ───────────────────────── ABA: INVESTIMENTOS ─────────────────────────
# compras e vendas da carteira: gravam na aba lancamentos da planilha da carteira,
# no mesmo formato de antes — o SmartWallet só lê.

with aba_invest:
    render_nav_mes("inv")
    ref_inv = st.session_state["mes_ref"]
    try:
        df_inv = carregar_investimentos()
        ativos_inv = carregar_ativos()
        erro_inv = None
    except Exception as e:
        df_inv = pd.DataFrame(columns=["linha", "data", "tipo", "ativo", "classe", "quantidade", "preco", "total"])
        ativos_inv, erro_inv = [], f"planilha da carteira indisponível ({str(e)[:60]})"

    df_inv_mes = filtrar_mes(df_inv, ref_inv.year, ref_inv.month)
    _tot = lambda d, t: float(d[d["tipo"] == t]["total"].sum()) if not d.empty else 0.0
    # média do líquido (compras − vendas) dos 6 meses anteriores ao mês exibido
    _liq, _m = [], ref_inv
    for _ in range(6):
        _m = mes_anterior(_m)
        _d = filtrar_mes(df_inv, _m.year, _m.month)
        _liq.append(_tot(_d, "compra") - _tot(_d, "venda"))
    _compras, _vendas = _tot(df_inv_mes, "compra"), _tot(df_inv_mes, "venda")
    render_resumo("saldo do mês", _compras - _vendas,
                  [("compras", _compras, "pos" if _compras else ""), ("vendas", -_vendas, "neg" if _vendas else ""),
                   ("média dos últimos 6 meses", sum(_liq) / 6, "")])
    if erro_inv:
        st.warning(erro_inv)
    st.write("")
    with st.container(key="novo_invest"):
        col_inv, col_atv = st.columns([1.3, 1])
        with col_inv:
            if st.button("+ novo aporte", use_container_width=True, disabled=bool(erro_inv)):
                dialogo_investimento(df_inv, ativos_inv)
        with col_atv:
            if st.button("+ novo ativo", use_container_width=True, disabled=bool(erro_inv)):
                dialogo_novo_ativo(ativos_inv)
    st.write("")

    if df_inv_mes.empty:
        st.caption("nenhum investimento neste mês.")
    else:
        df_inv_ord = df_inv_mes.iloc[::-1].sort_values("data", ascending=False, kind="mergesort")
        for idx_row, (_, r) in enumerate(df_inv_ord.iterrows()):
            it = {"data": r["data"], "tipo": r["tipo"], "ativo": r["ativo"], "classe": r["classe"],
                  "quantidade": float(r["quantidade"]), "preco": float(r["preco"]), "total": float(r["total"])}
            compra = it["tipo"] == "compra"
            linha_item(
                f"i{idx_row}_{it['data']}_{it['ativo']}_{it['total']}",
                it["ativo"],
                f"{it['tipo']} · {fmt_qtd(it['quantidade'])} × {fmt_brl2(it['preco'])} · {it['data'].strftime('%d/%m/%y')}",
                (f"<span style='color:#1D9E75;'>{formatar_brl(it['total'])}</span>" if compra
                 else (f"− {formatar_brl(it['total'])}" if it["tipo"] == "venda" else formatar_brl(it["total"]))),
                lambda it=it: dialogo_investimento(df_inv, ativos_inv, it),
                lambda it=it: excluir_investimento(it["data"], it["tipo"], it["ativo"], it["quantidade"], it["total"]),
                "excluir este investimento?")

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
                    novo_id = "cat_" + str(int(agora_br().timestamp()))
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

