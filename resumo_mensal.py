"""
resumo_mensal.py
Gera o PDF de resumo mensal de despesas (Controle de Fatura) e envia por e-mail.
Rodado automaticamente via GitHub Actions (ver .github/workflows/resumo-mensal.yml),
mas também pode ser rodado manualmente:  python resumo_mensal.py

Variáveis de ambiente necessárias (configuradas como Secrets no GitHub):
  GCP_SERVICE_ACCOUNT_JSON  -> conteúdo JSON completo da service account
  SPREADSHEET_ID            -> ID da planilha ControleFatura
  GMAIL_USER                -> e-mail Gmail que envia (ex: seuemail@gmail.com)
  GMAIL_APP_PASSWORD        -> senha de app de 16 caracteres gerada no Gmail
  EMAIL_DESTINO             -> e-mail que vai receber o resumo (pode ser o mesmo do GMAIL_USER)
"""

import os
import json
import smtplib
import calendar
from datetime import date, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

import gspread
from gspread.utils import ValueRenderOption
import pandas as pd
from google.oauth2.service_account import Credentials
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
)

FIXED_ID = "parcel"

# ───────────────────────── CONEXÃO GOOGLE SHEETS ─────────────────────────

def get_client():
    info = json.loads(os.environ["GCP_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"],
    )
    return gspread.authorize(creds)


def carregar_dados():
    sh = get_client().open_by_key(os.environ["SPREADSHEET_ID"])

    df_lanc = pd.DataFrame(sh.worksheet("lancamentos").get_all_records(value_render_option=ValueRenderOption.unformatted))
    if not df_lanc.empty:
        df_lanc["valor"] = pd.to_numeric(df_lanc["valor"], errors="coerce").fillna(0)
        df_lanc["data"] = pd.to_datetime(df_lanc["data"]).dt.date

    df_cat = pd.DataFrame(sh.worksheet("categorias").get_all_records(value_render_option=ValueRenderOption.unformatted))
    df_cat["valor_alvo"] = pd.to_numeric(df_cat["valor_alvo"], errors="coerce").fillna(0)
    df_cat["fixo"] = df_cat["fixo"].astype(str).str.lower().eq("true")

    df_cfg = pd.DataFrame(sh.worksheet("config").get_all_records(value_render_option=ValueRenderOption.unformatted))
    limite = 0.0
    linha = df_cfg[df_cfg["chave"] == "limite_mensal"]
    if not linha.empty:
        limite = float(linha.iloc[0]["valor"])

    return df_lanc, df_cat, limite


# ───────────────────────── LÓGICA DO RESUMO ─────────────────────────

def formatar_brl(valor):
    return "R$ " + f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def mes_anterior(ref):
    primeiro_dia_mes_atual = ref.replace(day=1)
    return (primeiro_dia_mes_atual - timedelta(days=1)).replace(day=1)


def filtrar_mes(df, ano, mes):
    if df.empty:
        return df
    return df[(pd.to_datetime(df["data"]).dt.year == ano) & (pd.to_datetime(df["data"]).dt.month == mes)]


def resumo_do_mes(df_lanc, ano, mes):
    df_mes = filtrar_mes(df_lanc, ano, mes)
    total = df_mes["valor"].sum() if not df_mes.empty else 0.0
    por_categoria = df_mes.groupby("categoria")["valor"].sum().to_dict() if not df_mes.empty else {}
    return df_mes, total, por_categoria


# ───────────────────────── GERAÇÃO DO PDF ─────────────────────────

def gerar_pdf(caminho, df_cat, limite, ref, df_mes_atual, total_atual, cat_atual,
              total_anterior, cat_anterior):

    styles = getSampleStyleSheet()
    titulo_style = ParagraphStyle("titulo", parent=styles["Heading1"], fontSize=18, spaceAfter=4)
    sub_style = ParagraphStyle("sub", parent=styles["Normal"], fontSize=10, textColor=colors.grey, spaceAfter=16)
    secao_style = ParagraphStyle("secao", parent=styles["Heading2"], fontSize=13, spaceBefore=18, spaceAfter=8)

    meses_pt = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
                "agosto", "setembro", "outubro", "novembro", "dezembro"]

    doc = SimpleDocTemplate(caminho, pagesize=A4,
                             topMargin=2*cm, bottomMargin=2*cm, leftMargin=2*cm, rightMargin=2*cm)
    elementos = []

    elementos.append(Paragraph("Resumo Mensal de Despesas", titulo_style))
    elementos.append(Paragraph(f"{meses_pt[ref.month - 1].capitalize()} de {ref.year}", sub_style))

    # ── Total vs limite ──
    gasto_parcel = df_mes_atual[df_mes_atual["categoria"] == FIXED_ID]["valor"].sum() if not df_mes_atual.empty else 0
    gasto_outros = max(0, total_atual - gasto_parcel)
    limite_efetivo = limite + gasto_parcel
    disponivel = limite_efetivo - total_atual

    elementos.append(Paragraph("Total gasto vs. limite", secao_style))
    dados_totais = [
        ["Limite do mês", formatar_brl(limite_efetivo)],
        ["Despesas (sem parcelamentos)", formatar_brl(gasto_outros)],
        ["Parcelamentos", formatar_brl(gasto_parcel)],
        ["Total gasto", formatar_brl(total_atual)],
        ["Disponível" if disponivel >= 0 else "Excedeu em", formatar_brl(abs(disponivel))],
    ]
    t = Table(dados_totais, colWidths=[9*cm, 6*cm])
    t.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("LINEBELOW", (0, 0), (-1, -2), 0.5, colors.lightgrey),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("TEXTCOLOR", (1, -1), (1, -1), colors.red if disponivel < 0 else colors.green),
    ]))
    elementos.append(t)

    # ── Comparação com mês anterior ──
    elementos.append(Paragraph("Comparação com o mês anterior", secao_style))
    delta_total = total_atual - total_anterior
    pct_delta = (delta_total / total_anterior * 100) if total_anterior > 0 else 0
    sinal = "+" if delta_total >= 0 else ""
    dados_comp = [
        ["Mês anterior", formatar_brl(total_anterior)],
        ["Mês atual", formatar_brl(total_atual)],
        ["Variação", f"{sinal}{formatar_brl(delta_total)} ({sinal}{pct_delta:.1f}%)"],
    ]
    t2 = Table(dados_comp, colWidths=[9*cm, 6*cm])
    t2.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("LINEBELOW", (0, 0), (-1, -2), 0.5, colors.lightgrey),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("TEXTCOLOR", (1, -1), (1, -1), colors.red if delta_total > 0 else colors.green),
    ]))
    elementos.append(t2)

    # ── Detalhamento por categoria ──
    elementos.append(Paragraph("Detalhamento por categoria", secao_style))
    linhas_cat = [["Categoria", "Gasto", "Orçamento", "% usado", "vs. mês ant."]]
    mapa_nomes = dict(zip(df_cat["id"], df_cat["nome"]))
    mapa_orcamento = dict(zip(df_cat["id"], df_cat["valor_alvo"]))
    todos_ids = sorted(set(list(cat_atual.keys()) + list(mapa_nomes.keys())),
                        key=lambda i: cat_atual.get(i, 0), reverse=True)
    for cid in todos_ids:
        if cid == FIXED_ID:
            continue
        nome = mapa_nomes.get(cid, cid)
        gasto = cat_atual.get(cid, 0)
        orc = mapa_orcamento.get(cid, 0)
        pct_uso = (gasto / orc * 100) if orc > 0 else 0
        gasto_ant = cat_anterior.get(cid, 0)
        delta_cat = gasto - gasto_ant
        sinal_cat = "+" if delta_cat >= 0 else ""
        if gasto == 0 and orc == 0:
            continue
        linhas_cat.append([
            nome, formatar_brl(gasto), formatar_brl(orc), f"{pct_uso:.0f}%",
            f"{sinal_cat}{formatar_brl(delta_cat)}"
        ])
    t3 = Table(linhas_cat, colWidths=[4.5*cm, 3*cm, 3*cm, 2*cm, 3*cm])
    t3.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a1a1a")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.lightgrey),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
    ]))
    elementos.append(t3)

    # ── Lista de todos os lançamentos ──
    elementos.append(Paragraph("Todos os lançamentos do mês", secao_style))
    if df_mes_atual.empty:
        elementos.append(Paragraph("Nenhum lançamento neste mês.", styles["Normal"]))
    else:
        linhas_lanc = [["Data", "Categoria", "Descrição", "Valor"]]
        df_ordenado = df_mes_atual.sort_values("data")
        for _, r in df_ordenado.iterrows():
            nome_cat = mapa_nomes.get(r["categoria"], "Parcelamentos" if r["categoria"] == FIXED_ID else r["categoria"])
            linhas_lanc.append([
                r["data"].strftime("%d/%m"), nome_cat, (r["descricao"] or "")[:40], formatar_brl(r["valor"])
            ])
        t4 = Table(linhas_lanc, colWidths=[2*cm, 3.5*cm, 6.5*cm, 3.5*cm], repeatRows=1)
        t4.setStyle(TableStyle([
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a1a1a")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.lightgrey),
            ("ALIGN", (3, 0), (3, -1), "RIGHT"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f7f7f7")]),
        ]))
        elementos.append(t4)

    elementos.append(Spacer(1, 20))
    elementos.append(Paragraph("Gerado automaticamente pelo Controle de Fatura.", sub_style))

    doc.build(elementos)


# ───────────────────────── ENVIO DE E-MAIL ─────────────────────────

def enviar_email(caminho_pdf, ref):
    meses_pt = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
                "agosto", "setembro", "outubro", "novembro", "dezembro"]
    assunto = f"Resumo de despesas — {meses_pt[ref.month - 1]} de {ref.year}"

    msg = MIMEMultipart()
    msg["From"] = os.environ["GMAIL_USER"]
    msg["To"] = os.environ["EMAIL_DESTINO"]
    msg["Subject"] = assunto
    msg.attach(MIMEText(
        f"Segue em anexo o resumo de despesas de {meses_pt[ref.month - 1]} de {ref.year}.\n\n"
        f"Gerado automaticamente pelo Controle de Fatura.", "plain"
    ))

    with open(caminho_pdf, "rb") as f:
        anexo = MIMEApplication(f.read(), _subtype="pdf")
        anexo.add_header("Content-Disposition", "attachment", filename=os.path.basename(caminho_pdf))
        msg.attach(anexo)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(os.environ["GMAIL_USER"], os.environ["GMAIL_APP_PASSWORD"])
        server.send_message(msg)


# ───────────────────────── MAIN ─────────────────────────

def main():
    hoje = date.today()
    ref = mes_anterior(hoje)  # resumo é sempre do mês que ACABOU de fechar

    df_lanc, df_cat, limite = carregar_dados()

    df_mes_atual, total_atual, cat_atual = resumo_do_mes(df_lanc, ref.year, ref.month)
    ref_anterior = mes_anterior(ref)
    _, total_anterior, cat_anterior = resumo_do_mes(df_lanc, ref_anterior.year, ref_anterior.month)

    nome_arquivo = f"resumo_{ref.year}_{ref.month:02d}.pdf"
    gerar_pdf(nome_arquivo, df_cat, limite, ref, df_mes_atual, total_atual, cat_atual,
              total_anterior, cat_anterior)

    enviar_email(nome_arquivo, ref)
    print(f"Resumo de {ref.month:02d}/{ref.year} enviado com sucesso.")


if __name__ == "__main__":
    main()
