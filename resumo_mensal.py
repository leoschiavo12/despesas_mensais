"""
resumo_mensal.py
Gera o resumo mensal (PDF + resumo no corpo do e-mail) do Controle de Fatura e envia por e-mail.
Rodado automaticamente via GitHub Actions (ver .github/workflows/resumo-mensal.yml),
mas também pode ser rodado manualmente:  python resumo_mensal.py

Variáveis de ambiente (Secrets no GitHub):
  GCP_SERVICE_ACCOUNT_JSON  -> conteúdo JSON completo da service account
  SPREADSHEET_ID            -> ID da planilha do Controle de Fatura
  CARTEIRA_SPREADSHEET_ID   -> ID da planilha da carteira (SmartWallet) — notas, parâmetros PJ,
                               investimentos e dividendos. sem ele, o resumo sai só com a fatura e o PIX
  GMAIL_USER                -> e-mail Gmail que envia
  GMAIL_APP_PASSWORD        -> senha de app de 16 caracteres gerada no Gmail
  EMAIL_DESTINO             -> e-mail que recebe o resumo

Envio: no 5º dia útil do mês seguinte (seg–sex, sem feriados nacionais), para dar tempo de
ajustar lançamentos do mês fechado. O workflow roda todo dia útil do 1º ao 10º; o script só
envia no 5º dia útil. FORCAR_ENVIO=1 (disparo manual) envia na hora.
REF_MES=aaaa-mm gera o resumo de um mês específico (útil para testar).
"""

import os
import re
import json
import smtplib
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import holidays
import gspread
from gspread.utils import ValueRenderOption, DateTimeOption
import pandas as pd
from google.oauth2.service_account import Credentials
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.barcharts import VerticalBarChart

FIXED_ID = "parcel"
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]
MESES3 = [m[:3] for m in MESES]
FLUXO_INICIO = "2026-12"      # primeiro mês com histórico do fluxo da PF (como no SmartWallet)

VERDE, VERMELHO, CINZA, LINHA = "#16a34a", "#dc2626", "#6b7280", "#e5e7eb"

PJ_PAR_PADRAO = {"faturamento_minimo": 0.0, "prev_pct": 0.0, "pro_labore": 0.0, "inss_pct": 11.0,
                 "das_pct": 6.0, "contador": 0.0, "fies": 0.0, "reembolso_base": 0.0,
                 "taxa_prefeitura": 0.0, "cartao_estimado": 0.0,
                 "fies_fim": "", "cartao_real_desde": ""}
PJ_PAR_TEXTO = {"fies_fim", "cartao_real_desde"}


# ───────────────────────── UTILITÁRIOS ─────────────────────────

def brl(v, sinal=False):
    """R$ 1.234,56 · com sinal=True, negativo vira '- R$ ...' (Helvetica não tem o '−')"""
    s = "R$ " + f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("- " + s) if (sinal and v < -0.004) else s


def pct(p):
    return f"{p:.1f}".replace(".", ",") + "%"


def num_br(x):
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


def para_data(x):
    """iso (aaaa-mm-dd), dd/mm/aaaa ou número de série do Sheets → date (ou None)"""
    if isinstance(x, (int, float)):
        return (pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(x))).date()
    s = str(x).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            pass
    return None


def ano_mes(s):
    s = str(s or "").strip()
    for fmt in ("%m/%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y-%m"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m")
        except ValueError:
            pass
    return None


def somar_meses(d, n):
    m = d.month - 1 + n
    return date(d.year + m // 12, m % 12 + 1, 1)


def chave(d):
    return d.strftime("%Y-%m")


def rot(d):
    return f"{MESES3[d.month - 1]}/{d.year % 100:02d}"


# ───────────────────────── LEITURA DAS PLANILHAS ─────────────────────────

def get_client():
    info = json.loads(os.environ["GCP_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/spreadsheets",
                      "https://www.googleapis.com/auth/drive"])
    return gspread.authorize(creds)


def _registros(ws):
    return ws.get_all_records(value_render_option=ValueRenderOption.unformatted)


def _df_datas(regs, cols):
    df = pd.DataFrame(regs)
    if df.empty:
        return pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df["data"] = df["data"].apply(para_data)
    df["valor"] = df["valor"].apply(num_br)
    return df.dropna(subset=["data", "valor"])[cols]


def carregar_fatura(gc):
    """planilha do Controle de Fatura: lançamentos (com arquivo_AAAA), categorias, limite e PIX"""
    sh = gc.open_by_key(os.environ["SPREADSHEET_ID"])
    abas = {w.title: w for w in sh.worksheets()}
    cols = ["data", "categoria", "descricao", "valor", "parcelamento"]
    regs = []
    for t, ws in abas.items():
        if t == "lancamentos" or t.startswith("arquivo_"):
            regs += _registros(ws)
    lanc = _df_datas(regs, cols)
    lanc["descricao"] = lanc["descricao"].fillna("").astype(str)
    lanc["parcelamento"] = lanc["parcelamento"].fillna("").astype(str)

    cat = pd.DataFrame(_registros(abas["categorias"]))
    cat["valor_alvo"] = pd.to_numeric(cat["valor_alvo"], errors="coerce").fillna(0)
    cat["fixo"] = cat["fixo"].astype(str).str.lower().eq("true")

    limite = 0.0
    for r in _registros(abas["config"]):
        if str(r.get("chave")) == "limite_mensal":
            limite = float(num_br(r.get("valor")) or 0)

    pix = _df_datas(_registros(abas["pix"]) if "pix" in abas else [],
                    ["data", "tipo", "categoria", "descricao", "valor"])
    pix["tipo"] = pix["tipo"].astype(str).str.strip().str.lower()
    pix["descricao"] = pix["descricao"].fillna("").astype(str)
    return lanc, cat, limite, pix


def carregar_carteira(gc):
    """planilha da carteira: notas/reembolsos, parâmetros PJ, faturas informadas, dividendos e
    investimentos. None se o secret CARTEIRA_SPREADSHEET_ID não estiver configurado."""
    sid = os.environ.get("CARTEIRA_SPREADSHEET_ID")
    if not sid:
        return None
    sh = gc.open_by_key(sid)
    abas = {w.title: w for w in sh.worksheets()}

    def linhas(nome, rng):
        if nome not in abas:
            return []
        return abas[nome].get(rng, value_render_option=ValueRenderOption.unformatted,
                              date_time_render_option=DateTimeOption.formatted_string)[1:]

    notas = []
    for r in linhas("pj_notas", "A:E"):
        r = (list(r) + [""] * 5)[:5]
        d, v = para_data(r[1]), num_br(r[2])
        if d and v is not None:
            tipo = "reembolso" if str(r[0]).strip().lower().startswith("reemb") else "nota fiscal"
            notas.append({"tipo": tipo, "data": d, "valor": v, "descricao": str(r[4]).strip()})
    notas = pd.DataFrame(notas, columns=["tipo", "data", "valor", "descricao"])

    par = dict(PJ_PAR_PADRAO)
    for r in linhas("pj_parametros", "A:B"):
        if len(r) >= 2 and str(r[0]).strip() in par:
            k = str(r[0]).strip()
            v = str(r[1]).strip() if k in PJ_PAR_TEXTO else num_br(r[1])
            if v is not None:
                par[k] = v
    hist = []
    for r in linhas("pj_parametros_hist", "A:C"):
        if len(r) >= 3 and str(r[0]).strip():
            d = str(r[1]).strip()
            d = "0000-00" if d in ("", "0000-00", "início", "inicio") else ano_mes(d)
            v = num_br(r[2])
            if d and v is not None:
                hist.append({"chave": str(r[0]).strip(), "desde": d, "valor": v})
    hist.sort(key=lambda e: (e["chave"], e["desde"]))
    mensal = lambda nome: {str(r[0]).strip(): num_br(r[1]) for r in linhas(nome, "A:B")
                           if len(r) >= 2 and str(r[0]).strip() and num_br(r[1]) is not None}

    inv = []
    for r in linhas("lancamentos", "A:G"):
        r = (list(r) + [""] * 7)[:7]
        d, q, p, t = para_data(r[0]), num_br(r[4]), num_br(r[5]), num_br(r[6])
        if d and q is not None:
            inv.append({"data": d, "tipo": str(r[1]).strip().lower(), "ativo": str(r[2]).strip(),
                        "classe": str(r[3]).strip(), "quantidade": q, "preco": p or 0.0,
                        "total": t if t is not None else q * (p or 0.0)})
    inv = pd.DataFrame(inv, columns=["data", "tipo", "ativo", "classe", "quantidade", "preco", "total"])
    return {"notas": notas, "par": par, "hist": hist, "cartao": mensal("pf_cartao"),
            "divs": mensal("dividendos_mensais"), "inv": inv}


# ───────────────────────── CÁLCULOS ─────────────────────────

def do_mes(df, ref):
    if df.empty:
        return df
    return df[df["data"].apply(lambda d: d.year == ref.year and d.month == ref.month)]


def soma(df, col="valor"):
    return float(df[col].sum()) if not df.empty else 0.0


def par_vigente(par, hist, k):
    out = dict(par)
    for c in {e["chave"] for e in hist}:
        regs = [e for e in hist if e["chave"] == c]
        validos = [e for e in regs if e["desde"] <= k]
        out[c] = (validos[-1] if validos else regs[0])["valor"]
    return out


def fluxo_pf(ref, cart, lanc, pix):
    """entradas e saídas da conta PF no mês ref — mesma conta do fluxo do SmartWallet e da aba
    resumo do app: pró-labore e lucros vêm das notas emitidas no mês anterior; a fatura paga
    no mês é a do mês anterior"""
    k, ant = chave(ref), somar_meses(ref, -1)
    a = chave(ant)
    par = par_vigente(cart["par"], cart["hist"], k)
    notas = cart["notas"]
    _n = lambda m, t: soma(do_mes(notas[notas["tipo"] == t], m)) if not notas.empty else 0.0
    fat = _n(ant, "nota fiscal")
    prev = par["faturamento_minimo"] * par["prev_pct"] / 100
    das = fat * par["das_pct"] / 100
    taxa = par["taxa_prefeitura"] if ref.month == 2 else 0.0
    lucros = fat - prev - das - par["contador"] - taxa - par["pro_labore"]
    pl_liq = par["pro_labore"] * (1 - par["inss_pct"] / 100)
    fim = ano_mes(par.get("fies_fim"))
    fies = min(par["fies"], max(lucros, 0.0)) if (fim is None or k <= fim) else 0.0
    desde = ano_mes(par.get("cartao_real_desde")) or "0000-00"
    if a in cart["cartao"]:
        cartao = cart["cartao"][a]
    elif a < desde:
        cartao = par["cartao_estimado"]
    else:
        cartao = soma(do_mes(lanc, ant))
    px = do_mes(pix, ref)
    inv = do_mes(cart["inv"], ref)
    investido = soma(inv[inv["tipo"] == "compra"], "total") - soma(inv[inv["tipo"] == "venda"], "total") \
        if not inv.empty else 0.0
    entradas = [("lucros distribuídos", lucros), ("pró-labore líquido", pl_liq),
                ("reembolsos", par["reembolso_base"] + _n(ref, "reembolso")),
                ("outras receitas PIX", soma(px[px["tipo"] == "entrada"]) if not px.empty else 0.0),
                ("dividendos", cart["divs"].get(k, 0.0))]
    saidas = [(f"fatura de {rot(ant)}", cartao), ("FIES", fies),
              ("outras despesas PIX", soma(px[px["tipo"] == "saida"]) if not px.empty else 0.0),
              ("investido", investido)]
    lim = lambda l: sorted([x for x in l if abs(x[1]) >= 0.005], key=lambda x: -x[1])
    entradas, saidas = lim(entradas), lim(saidas)
    return {"entradas": entradas, "saidas": saidas, "tem_div": k in cart["divs"],
            "entrou": sum(v for _, v in entradas), "saiu": sum(v for _, v in saidas),
            "saldo": sum(v for _, v in entradas) - sum(v for _, v in saidas)}


def montar_dados(ref):
    gc = get_client()
    lanc, cat, limite, pix = carregar_fatura(gc)
    cart = carregar_carteira(gc)
    nomes = dict(zip(cat["id"], cat["nome"]))
    orc = dict(zip(cat["id"], cat["valor_alvo"]))

    # ── fatura do mês ──
    mes = do_mes(lanc, ref)
    ant = do_mes(lanc, somar_meses(ref, -1))
    total, total_ant = soma(mes), soma(ant)
    parcel = soma(mes[mes["categoria"] == FIXED_ID]) if not mes.empty else 0.0
    por_cat = mes.groupby("categoria")["valor"].sum().to_dict() if not mes.empty else {}
    por_cat_ant = ant.groupby("categoria")["valor"].sum().to_dict() if not ant.empty else {}
    cats = []
    for cid in set(por_cat) | {c for c in nomes if c != FIXED_ID}:
        g, o = por_cat.get(cid, 0.0), orc.get(cid, 0.0)
        if cid == FIXED_ID or (g == 0 and o == 0):
            continue
        cats.append({"nome": nomes.get(cid, cid), "gasto": g, "orc": o,
                     "uso": g / o * 100 if o > 0 else None, "delta": g - por_cat_ant.get(cid, 0.0)})
    cats.sort(key=lambda c: -c["gasto"])

    # ── destaques ──
    top = mes.sort_values("valor", ascending=False).head(5) if not mes.empty else mes
    top = [(r["descricao"] or nomes.get(r["categoria"], r["categoria"]), r["data"], r["valor"])
           for _, r in top.iterrows()]
    estouros = [c for c in cats if c["orc"] > 0 and c["gasto"] > c["orc"]]

    # ── parcelamentos futuros: o que já está comprometido nas faturas depois do mês do resumo ──
    fut = lanc[(lanc["parcelamento"] != "") & lanc["data"].apply(lambda d: chave(d) > chave(ref))] \
        if not lanc.empty else lanc
    parc_mes = fut.groupby(fut["data"].apply(lambda d: date(d.year, d.month, 1)))["valor"].sum() \
        .sort_index().to_dict() if not fut.empty else {}
    parc_compras = []
    for gid, g in (fut.groupby("parcelamento") if not fut.empty else []):
        g = g.sort_values("data")
        m = re.match(r"^(.*) \((\d+)/(\d+)\)$", str(g.iloc[0]["descricao"]))
        parc_compras.append({"nome": m.group(1) if m else str(g.iloc[0]["descricao"]),
                             "restam": len(g), "n": int(m.group(3)) if m else len(g),
                             "parcela": float(g.iloc[0]["valor"]), "ate": g.iloc[-1]["data"],
                             "total": soma(g)})
    parc_compras.sort(key=lambda p: -p["total"])

    # ── lançamentos do mês (cartão + PIX de saída) ──
    lst = [(r["data"], nomes.get(r["categoria"], "parcelamentos" if r["categoria"] == FIXED_ID
                                 else r["categoria"]), r["descricao"], "cartão", r["valor"])
           for _, r in mes.iterrows()]
    px = do_mes(pix, ref)
    for _, r in (px[px["tipo"] == "saida"].iterrows() if not px.empty else []):
        lst.append((r["data"], nomes.get(r["categoria"], r["categoria"]), r["descricao"], "pix", r["valor"]))
    lst.sort(key=lambda x: x[0])

    # ── receitas do mês ──
    receitas = []
    if cart is not None:
        for _, r in do_mes(cart["notas"], ref).iterrows():
            receitas.append((r["data"], r["tipo"], r["descricao"], r["valor"]))
    for _, r in (px[px["tipo"] == "entrada"].iterrows() if not px.empty else []):
        receitas.append((r["data"], "pix", r["descricao"], r["valor"]))
    receitas.sort(key=lambda x: x[0])

    # ── investimentos do mês ──
    inv = None
    if cart is not None:
        im = do_mes(cart["inv"], ref)
        im = im[im["tipo"].isin(["compra", "venda"])] if not im.empty else im
        tot = lambda d, t: soma(d[d["tipo"] == t], "total") if not d.empty else 0.0
        liq6 = [tot(do_mes(cart["inv"], somar_meses(ref, -i)), "compra")
                - tot(do_mes(cart["inv"], somar_meses(ref, -i)), "venda") for i in range(1, 7)]
        inv = {"lanc": im.sort_values("data") if not im.empty else im,
               "compras": tot(im, "compra"), "vendas": tot(im, "venda"), "media6": sum(liq6) / 6}

    # ── fluxo da PF e evolução de 12 meses ──
    meses12 = [somar_meses(ref, -i) for i in range(11, -1, -1)]
    evol_fatura = [(m, soma(do_mes(lanc, m))) for m in meses12]
    fluxo, evol_saldo = None, []
    if cart is not None:
        fluxo = fluxo_pf(ref, cart, lanc, pix)
        evol_saldo = [(m, fluxo_pf(m, cart, lanc, pix)["saldo"]) for m in meses12 if chave(m) >= FLUXO_INICIO]

    return {"ref": ref, "limite": limite, "total": total, "total_ant": total_ant, "parcel": parcel,
            "cats": cats, "top": top, "estouros": estouros, "parc_mes": parc_mes,
            "parc_compras": parc_compras, "lancamentos": lst, "receitas": receitas, "inv": inv,
            "fluxo": fluxo, "evol_fatura": evol_fatura, "evol_saldo": evol_saldo,
            "sem_carteira": cart is None}


# ───────────────────────── PDF ─────────────────────────
# mesmo espírito dos resumos do app: listas limpas (só uma linha fina entre os itens),
# rótulos em minúsculas, verde para entradas e vermelho para saídas

ST = getSampleStyleSheet()
S_TIT = ParagraphStyle("tit", parent=ST["Normal"], fontName="Helvetica-Bold", fontSize=17, leading=21)
S_SUB = ParagraphStyle("sub", parent=ST["Normal"], fontSize=9, textColor=colors.HexColor(CINZA), spaceAfter=10)
S_SEC = ParagraphStyle("sec", parent=ST["Normal"], fontName="Helvetica-Bold", fontSize=11.5,
                       spaceBefore=16, spaceAfter=5)
S_TXT = ParagraphStyle("txt", parent=ST["Normal"], fontSize=8.5, textColor=colors.HexColor(CINZA))
S_CEL = ParagraphStyle("cel", parent=ST["Normal"], fontSize=9, leading=11)


def lista(linhas, larguras, cores=None, total=False, cab=None, alinhar_dir=None, fonte=9):
    """tabela no visual do app: sem grade, linha fina entre itens, última linha em negrito se total.
    cores: {(col, lin): '#hex'} · alinhar_dir: colunas à direita (padrão: a última)"""
    dados = ([cab] if cab else []) + linhas
    t = Table(dados, colWidths=larguras, hAlign="LEFT")
    dir_cols = alinhar_dir if alinhar_dir is not None else [len(larguras) - 1]
    est = [("FONTSIZE", (0, 0), (-1, -1), fonte), ("TOPPADDING", (0, 0), (-1, -1), 4),
           ("BOTTOMPADDING", (0, 0), (-1, -1), 4), ("LEFTPADDING", (0, 0), (-1, -1), 2),
           ("RIGHTPADDING", (0, 0), (-1, -1), 2), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
           ("LINEBELOW", (0, 0), (-1, -2 if total else -1), 0.5, colors.HexColor(LINHA))]
    for c in dir_cols:
        est.append(("ALIGN", (c, 0), (c, -1), "RIGHT"))
    if cab:
        est += [("FONTSIZE", (0, 0), (-1, 0), 7.5), ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor(CINZA))]
    if total:
        est += [("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("LINEABOVE", (0, -1), (-1, -1), 0.8, colors.HexColor("#9ca3af"))]
    off = 1 if cab else 0
    for (c, l), cor in (cores or {}).items():
        est.append(("TEXTCOLOR", (c, l + off), (c, l + off), colors.HexColor(cor)))
    t.setStyle(TableStyle(est))
    return t


def grafico(titulo, pontos, cor):
    """barras simples de 12 meses (negativos em vermelho)"""
    d = Drawing(17 * cm, 4.6 * cm)
    d.add(String(0, 4.25 * cm, titulo, fontName="Helvetica", fontSize=8.5, fillColor=colors.HexColor(CINZA)))
    ch = VerticalBarChart()
    ch.x, ch.y, ch.width, ch.height = 1.4 * cm, 0.6 * cm, 15.4 * cm, 3.3 * cm
    vals = [v for _, v in pontos]
    ch.data = [vals]
    ch.categoryAxis.categoryNames = [rot(m) for m, _ in pontos]
    ch.categoryAxis.labels.fontSize = 6.5
    ch.categoryAxis.labels.fontName = "Helvetica"
    ch.valueAxis.labels.fontName = "Helvetica"
    ch.categoryAxis.visibleTicks = False
    ch.categoryAxis.strokeColor = colors.HexColor(LINHA)
    ch.valueAxis.labels.fontSize = 6.5
    ch.valueAxis.strokeColor = colors.white
    ch.valueAxis.gridStrokeColor = colors.HexColor("#f1f1f1")
    ch.valueAxis.visibleGrid = True
    ch.valueAxis.valueMin = min(0, min(vals) * 1.15)
    ch.valueAxis.valueMax = max(1, max(vals) * 1.15)
    ch.valueAxis.labelTextFormat = lambda v: f"{v / 1000:.0f}k" if abs(v) >= 1000 else f"{v:.0f}"
    ch.bars.strokeColor = None
    ch.barWidth = 6
    for i, v in enumerate(vals):
        ch.bars[(0, i)].fillColor = colors.HexColor(VERMELHO if v < 0 else cor)
    d.add(ch)
    return d


def gerar_pdf(caminho, D):
    ref = D["ref"]
    doc = SimpleDocTemplate(caminho, pagesize=A4, topMargin=1.6 * cm, bottomMargin=1.6 * cm,
                            leftMargin=2 * cm, rightMargin=2 * cm)
    W = 17 * cm
    el = [Paragraph(f"resumo de {MESES[ref.month - 1]} de {ref.year}", S_TIT),
          Paragraph("controle de fatura · pessoa física", S_SUB)]
    F = D["fluxo"]

    # ── visão geral ──
    disp = D["limite"] - D["total"]
    lin = [["fatura do cartão", brl(D["total"])],
           ["limite disponível" if disp >= 0 else "limite excedido", brl(abs(disp))]]
    cores = {(1, 1): VERDE if disp >= 0 else VERMELHO}
    if F:
        lin += [["entrou na conta", brl(F["entrou"])], ["saiu da conta", brl(F["saiu"])],
                ["saldo do mês", brl(F["saldo"], True)]]
        cores.update({(1, 2): VERDE, (1, 3): VERMELHO, (1, 4): VERDE if F["saldo"] >= 0 else VERMELHO})
    el += [Paragraph("visão geral", S_SEC), lista(lin, [11 * cm, 6 * cm], cores, total=bool(F))]

    # ── fluxo da PF ──
    if F:
        lin, cores = [], {}
        lin.append([Paragraph("<font size=7 color='#6b7280'>ENTRADAS</font>", S_CEL), ""])
        for r, v in F["entradas"]:
            cores[(1, len(lin))] = VERDE if v >= 0 else VERMELHO
            lin.append([r, brl(v, True)])
        lin.append([Paragraph("<font size=7 color='#6b7280'>SAÍDAS</font>", S_CEL), ""])
        for r, v in F["saidas"]:
            cores[(1, len(lin))] = VERMELHO if v >= 0 else VERDE
            lin.append([r, brl(-v, True)])
        lin.append(["saldo", brl(F["saldo"], True)])
        cores[(1, len(lin) - 1)] = VERDE if F["saldo"] >= 0 else VERMELHO
        el += [Paragraph(f"movimentação da conta em {MESES[ref.month - 1]}", S_SEC),
               lista(lin, [11 * cm, 6 * cm], cores, total=True)]
        if not F["tem_div"]:
            el.append(Paragraph("dividendos do mês ainda não gravados na planilha da carteira "
                                "(abra o SmartWallet para atualizar).", S_TXT))

    # ── fatura do cartão ──
    delta = D["total"] - D["total_ant"]
    lin = [["despesas (sem parcelamentos)", brl(D["total"] - D["parcel"])],
           ["parcelamentos", brl(D["parcel"])],
           ["total", brl(D["total"])],
           ["mês anterior", brl(D["total_ant"])],
           ["variação", f"{'+' if delta >= 0 else '-'} {brl(abs(delta))}"
                        + (f" ({pct(delta / D['total_ant'] * 100)})" if D["total_ant"] > 0 else "")]]
    el += [Paragraph(f"fatura de {MESES[ref.month - 1]}", S_SEC),
           lista(lin, [11 * cm, 6 * cm], {(1, 4): VERMELHO if delta > 0 else VERDE})]
    if D["cats"]:
        lin, cores = [], {}
        for c in D["cats"]:
            if c["uso"] is not None and c["uso"] > 100:
                cores[(3, len(lin))] = VERMELHO
            lin.append([c["nome"], brl(c["gasto"]), brl(c["orc"]),
                        pct(c["uso"]) if c["uso"] is not None else "—",
                        f"{'+' if c['delta'] >= 0 else '-'} {brl(abs(c['delta']))}"])
        el += [Spacer(1, 8), lista(lin, [5 * cm, 3 * cm, 3 * cm, 2.2 * cm, 3.8 * cm], cores,
                                   cab=["categoria", "gasto", "orçamento", "% usado", "vs. mês ant."],
                                   alinhar_dir=[1, 2, 3, 4])]

    # ── destaques ──
    if D["top"] or D["estouros"]:
        bloco = [Paragraph("destaques do mês", S_SEC)]
        if D["top"]:
            bloco.append(lista([[str(n)[:50], d.strftime("%d/%m"), brl(v)] for n, d, v in D["top"]],
                            [11 * cm, 2 * cm, 4 * cm], cab=["maiores gastos", "data", "valor"],
                            alinhar_dir=[2]))
        if D["estouros"]:
            bloco += [Spacer(1, 6), lista(
                [[c["nome"], brl(c["gasto"] - c["orc"]), pct(c["uso"])] for c in D["estouros"]],
                [11 * cm, 3.5 * cm, 2.5 * cm], {(1, i): VERMELHO for i in range(len(D["estouros"]))},
                cab=["acima do orçamento", "excesso", "% usado"], alinhar_dir=[1, 2])]
        el.append(KeepTogether(bloco))

    # ── receitas ──
    if D["receitas"]:
        lin = [[d.strftime("%d/%m"), t, Paragraph(str(ds)[:60], S_CEL), brl(v)] for d, t, ds, v in D["receitas"]]
        el += [Paragraph(f"receitas lançadas em {MESES[ref.month - 1]}", S_SEC),
               lista(lin, [1.6 * cm, 2.8 * cm, 9 * cm, 3.6 * cm],
                     {(3, i): VERDE for i in range(len(lin))}, cab=["data", "tipo", "descrição", "valor"]),
               Paragraph("notas fiscais pelo valor bruto, na data de emissão (entram na conta no mês seguinte).",
                         S_TXT)]

    # ── investimentos ──
    I = D["inv"]
    if I is not None:
        lin = [["compras", brl(I["compras"])], ["vendas", brl(I["vendas"])],
               ["média dos últimos 6 meses", brl(I["media6"], True)],
               ["saldo", brl(I["compras"] - I["vendas"], True)]]
        el += [Paragraph(f"investimentos em {MESES[ref.month - 1]}", S_SEC),
               lista(lin, [11 * cm, 6 * cm], {(1, 0): VERDE if I["compras"] else CINZA,
                                             (1, 1): VERMELHO if I["vendas"] else CINZA}, total=True)]
        if not I["lanc"].empty:
            lin = [[r["data"].strftime("%d/%m"), r["ativo"], r["tipo"],
                    f"{r['quantidade']:.8f}".rstrip("0").rstrip(".").replace(".", ","),
                    brl(r["preco"]), brl(r["total"])] for _, r in I["lanc"].iterrows()]
            el += [Spacer(1, 6), lista(lin, [1.6 * cm, 3.4 * cm, 2 * cm, 2.6 * cm, 3.4 * cm, 4 * cm],
                                       cab=["data", "ativo", "tipo", "quantidade", "preço", "total"],
                                       alinhar_dir=[3, 4, 5])]

    # ── parcelamentos futuros ──
    if D["parc_mes"]:
        meses = list(D["parc_mes"].items())
        lin = [[rot(m), brl(v)] for m, v in meses] + [["total comprometido", brl(sum(v for _, v in meses))]]
        bloco = [Paragraph("parcelamentos nas próximas faturas", S_SEC),
                 lista(lin, [11 * cm, 6 * cm], total=True)]
        if D["parc_compras"]:
            bloco += [Spacer(1, 6), lista(
                [[p["nome"][:45], f"{p['restam']} de {p['n']}", brl(p["parcela"]), rot(p["ate"]), brl(p["total"])]
                 for p in D["parc_compras"]],
                [6 * cm, 2.2 * cm, 3.2 * cm, 2.2 * cm, 3.4 * cm],
                cab=["compra", "restam", "parcela", "até", "a pagar"], alinhar_dir=[2, 3, 4])]
        el.append(KeepTogether(bloco))

    # ── evolução ──
    graf = [Paragraph("evolução dos últimos 12 meses", S_SEC),
            grafico("fatura do cartão por mês (R$)", D["evol_fatura"], "#6b7280")]
    if len(D["evol_saldo"]) >= 2:
        graf.append(grafico(f"saldo da conta PF por mês (R$) · desde {rot(D['evol_saldo'][0][0])}",
                            D["evol_saldo"], VERDE))
    el.append(KeepTogether(graf))

    # ── todos os lançamentos ──
    el.append(Paragraph(f"todos os lançamentos de {MESES[ref.month - 1]}", S_SEC))
    if not D["lancamentos"]:
        el.append(Paragraph("nenhum lançamento neste mês.", S_TXT))
    else:
        lin = [[d.strftime("%d/%m"), c, Paragraph(str(ds)[:60], S_CEL), p, brl(v)]
               for d, c, ds, p, v in D["lancamentos"]]
        el.append(lista(lin, [1.6 * cm, 3.2 * cm, 7.4 * cm, 1.6 * cm, 3.2 * cm],
                        cab=["data", "categoria", "descrição", "pagto.", "valor"], fonte=8.5))

    el += [Spacer(1, 16), Paragraph(
        f"gerado automaticamente pelo controle de fatura em "
        f"{datetime.now(ZoneInfo('America/Sao_Paulo')):%d/%m/%Y %H:%M}.", S_TXT)]
    if D["sem_carteira"]:
        el.append(Paragraph("configure o secret CARTEIRA_SPREADSHEET_ID no GitHub para incluir o fluxo da PF, "
                            "as receitas da PJ e os investimentos.", S_TXT))
    doc.build(el)


# ───────────────────────── E-MAIL ─────────────────────────

def corpo_html(D):
    """resumo no corpo do e-mail: os números principais, no visual do app"""
    ref, F = D["ref"], D["fluxo"]

    def bloco(titulo, linhas, total=None):
        tr = "".join(
            f"<tr><td style='padding:6px 0;border-bottom:1px solid #eee;'>{r}</td>"
            f"<td style='padding:6px 0;border-bottom:1px solid #eee;text-align:right;color:{c};"
            f"white-space:nowrap;'>{v}</td></tr>" for r, v, c in linhas)
        if total:
            tr += (f"<tr><td style='padding:6px 0;font-weight:600;'>{total[0]}</td>"
                   f"<td style='padding:6px 0;text-align:right;font-weight:600;color:{total[2]};'>{total[1]}</td></tr>")
        return (f"<p style='margin:18px 0 4px;font-size:12px;color:#6b7280;'>{titulo}</p>"
                f"<table style='width:100%;border-collapse:collapse;font-size:14px;'>{tr}</table>")

    disp = D["limite"] - D["total"]
    h = [f"<div style='font-family:-apple-system,Segoe UI,Helvetica,Arial,sans-serif;max-width:480px;color:#111;'>"
         f"<p style='font-size:18px;font-weight:600;margin:0;'>resumo de {MESES[ref.month - 1]} de {ref.year}</p>"]
    h.append(bloco(f"fatura de {MESES[ref.month - 1]}", [
        ("total da fatura", brl(D["total"]), "#111"),
        ("limite disponível" if disp >= 0 else "limite excedido", brl(abs(disp)), VERDE if disp >= 0 else VERMELHO),
        ("vs. mês anterior", f"{'+' if D['total'] >= D['total_ant'] else '-'} {brl(abs(D['total'] - D['total_ant']))}",
         VERMELHO if D["total"] > D["total_ant"] else VERDE)]))
    if F:
        h.append(bloco(f"movimentação da conta em {MESES[ref.month - 1]}",
                       [(r, brl(v, True), VERDE if v >= 0 else VERMELHO) for r, v in F["entradas"]]
                       + [(r, brl(-v, True), VERMELHO if v >= 0 else VERDE) for r, v in F["saidas"]],
                       ("saldo", brl(F["saldo"], True), VERDE if F["saldo"] >= 0 else VERMELHO)))
    if D["inv"] is not None:
        I = D["inv"]
        h.append(bloco(f"investimentos em {MESES[ref.month - 1]}", [
            ("compras", brl(I["compras"]), VERDE if I["compras"] else CINZA),
            ("vendas", brl(I["vendas"]), VERMELHO if I["vendas"] else CINZA),
            ("média dos últimos 6 meses", brl(I["media6"], True), "#111")],
            ("saldo", brl(I["compras"] - I["vendas"], True), "#111")))
    dest = [(f"{str(n)[:40]} · {d:%d/%m}", brl(v), "#111") for n, d, v in D["top"][:3]]
    dest += [(f"{c['nome']} acima do orçamento", f"+ {brl(c['gasto'] - c['orc'])}", VERMELHO) for c in D["estouros"]]
    if dest:
        h.append(bloco("destaques", dest))
    if D["parc_mes"]:
        prox = list(D["parc_mes"].items())[0]
        h.append(bloco("parcelamentos", [
            (f"já na fatura de {rot(prox[0])}", brl(prox[1]), "#111"),
            ("total comprometido", brl(sum(D["parc_mes"].values())), "#111")]))
    h.append("<p style='margin:20px 0 0;font-size:12px;color:#6b7280;'>o detalhamento completo está no PDF anexo."
             "</p></div>")
    return "".join(h)


def enviar_email(caminho_pdf, D):
    ref = D["ref"]
    msg = MIMEMultipart("mixed")
    msg["From"] = os.environ["GMAIL_USER"]
    msg["To"] = os.environ["EMAIL_DESTINO"]
    msg["Subject"] = f"resumo de {MESES[ref.month - 1]} de {ref.year}"
    alt = MIMEMultipart("alternative")
    F = D["fluxo"]
    texto = (f"fatura de {MESES[ref.month - 1]}: {brl(D['total'])}"
             + (f"\nentrou na conta: {brl(F['entrou'])}\nsaiu da conta: {brl(F['saiu'])}"
                f"\nsaldo: {brl(F['saldo'], True)}" if F else "")
             + "\n\no detalhamento completo está no PDF anexo.")
    alt.attach(MIMEText(texto, "plain", "utf-8"))
    alt.attach(MIMEText(corpo_html(D), "html", "utf-8"))
    msg.attach(alt)
    with open(caminho_pdf, "rb") as f:
        anexo = MIMEApplication(f.read(), _subtype="pdf")
        anexo.add_header("Content-Disposition", "attachment", filename=os.path.basename(caminho_pdf))
        msg.attach(anexo)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(os.environ["GMAIL_USER"], os.environ["GMAIL_APP_PASSWORD"])
        server.send_message(msg)


# ───────────────────────── AGENDA ─────────────────────────

def dia_util(d):
    return d.weekday() < 5 and d not in holidays.Brazil(years=d.year)


def quinto_dia_util(ano, mes):
    d, n = date(ano, mes, 1), 0
    while True:
        if dia_util(d):
            n += 1
            if n == 5:
                return d
        d += timedelta(days=1)


def marcar_saida(enviado):
    """avisa o workflow se o resumo foi enviado (o keepalive só roda nesse caso)"""
    if os.environ.get("GITHUB_OUTPUT"):
        with open(os.environ["GITHUB_OUTPUT"], "a") as f:
            f.write(f"enviado={'true' if enviado else 'false'}\n")


# ───────────────────────── MAIN ─────────────────────────

def main():
    hoje = datetime.now(ZoneInfo("America/Sao_Paulo")).date()   # horário de Brasília
    if os.environ.get("REF_MES"):
        ref = datetime.strptime(os.environ["REF_MES"], "%Y-%m").date()
    else:
        alvo = quinto_dia_util(hoje.year, hoje.month)
        if hoje != alvo and os.environ.get("FORCAR_ENVIO") != "1":
            print(f"hoje ({hoje:%d/%m}) não é o 5º dia útil ({alvo:%d/%m}) — nada a enviar.")
            marcar_saida(False)
            return
        ref = somar_meses(hoje.replace(day=1), -1)   # o mês que acabou de fechar

    D = montar_dados(ref)
    nome_arquivo = f"resumo_{ref.year}_{ref.month:02d}.pdf"
    gerar_pdf(nome_arquivo, D)
    if os.environ.get("SO_PDF") == "1":
        print(f"PDF gerado: {nome_arquivo}")
        return
    enviar_email(nome_arquivo, D)
    print(f"Resumo de {ref.month:02d}/{ref.year} enviado com sucesso.")
    marcar_saida(True)


if __name__ == "__main__":
    main()
