import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import urllib
import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime
from pathlib import Path

# =====================================================
# CONFIGURAÇÕES
# =====================================================

PASTA_SAIDA = r"C:\Users\srvautbicar\Desktop\PROCESSOS\import para SQL\Base PDD\Relatorios"

# =====================================================
# PARÂMETRO: MÊS ATUAL
# =====================================================

print("=" * 55)
print("  RELATÓRIO PDD - COMPARATIVO MENSAL")
print("=" * 55)
print()
print("Informe o mês que deseja analisar.")
print("Formato esperado: AAAA-MM-DD  (ex: 2026-04-30)")
print()

def pedir_data(label):
    while True:
        entrada = input(f"📅 {label}: ").strip()
        try:
            return datetime.strptime(entrada, "%Y-%m-%d").date()
        except ValueError:
            print("❌ Formato inválido. Use AAAA-MM-DD\n")

dt_atual = pedir_data("Mês de análise (ex: 2026-04-30)")

# =====================================================
# CONEXÃO SQL SERVER
# =====================================================

params = urllib.parse.quote_plus(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=SRJDLK01\\SSPDLKAZ01;"
    "DATABASE=P_OURO_CAR_DB;"
    "UID=powerbi;"
    "PWD=Ah79Wlk999;"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}", fast_executemany=True)

print("\n🔌 Conectando ao banco e buscando dados...")

# =====================================================
# BUSCAR MÊS ANTERIOR NO BANCO AUTOMATICAMENTE
# =====================================================

with engine.connect() as conn:
    dt_anterior = conn.execute(
        text("SELECT MAX(dt_fechamento) FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR WHERE dt_fechamento < :dt"),
        {"dt": str(dt_atual)}
    ).scalar()

if dt_anterior is None:
    print(f"\n❌ Não existe mês anterior a {dt_atual} na tabela.")
    print("   Verifique se os dados foram importados corretamente.")
    exit()

MESES_PT = {
    "JANUARY": "JANEIRO", "FEBRUARY": "FEVEREIRO", "MARCH": "MARÇO",
    "APRIL": "ABRIL", "MAY": "MAIO", "JUNE": "JUNHO",
    "JULY": "JULHO", "AUGUST": "AGOSTO", "SEPTEMBER": "SETEMBRO",
    "OCTOBER": "OUTUBRO", "NOVEMBER": "NOVEMBRO", "DECEMBER": "DEZEMBRO"
}

def nome_mes(dt):
    nome = dt.strftime("%B/%Y").upper()
    for en, pt in MESES_PT.items():
        nome = nome.replace(en, pt)
    return nome

nome_atual    = nome_mes(dt_atual)
nome_anterior = nome_mes(dt_anterior)

print(f"✅ Comparativo identificado: {nome_anterior}  →  {nome_atual}")

# =====================================================
# LOGO (arquivo local)
# =====================================================

import base64
from pathlib import Path

PASTA_SAIDA = r"C:\Users\srvautbicar\Desktop\PROCESSOS\import para SQL\Base PDD\Relatorios"
LOGO_PATH   = r"C:\Users\srvautbicar\Desktop\PROCESSOS\import para SQL\Base PDD\logo_cury.png"

try:
    with open(LOGO_PATH, "rb") as f:
        logo_src = "data:image/png;base64," + base64.b64encode(f.read()).decode()
except:
    logo_src = ""

# =====================================================
# QUERIES — agrupadas direto no SQL
# =====================================================

with engine.connect() as conn:

    # 1. PDD por empreendimento
    r = conn.execute(text("""
        SELECT empreendimento, regional,
               SUM(valor_pdd) AS vat, SUM(valor_pdd_anterior) AS vant
        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR WHERE dt_fechamento = :dt
        GROUP BY empreendimento, regional
    """), {"dt": str(dt_atual)})
    df_emp = pd.DataFrame(r.fetchall(), columns=r.keys())

    # 2. PDD por aging
    r = conn.execute(text("""
        SELECT aging_da_unidade,
               SUM(valor_pdd) AS vat, SUM(valor_pdd_anterior) AS vant
        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR WHERE dt_fechamento = :dt
        GROUP BY aging_da_unidade
    """), {"dt": str(dt_atual)})
    df_aging = pd.DataFrame(r.fetchall(), columns=r.keys())

    # 3. PDD por status de financiamento
    r = conn.execute(text("""
        SELECT status_de_financiamento,
               SUM(valor_pdd) AS vat, SUM(valor_pdd_anterior) AS vant
        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR WHERE dt_fechamento = :dt
          AND status_de_financiamento IS NOT NULL
        GROUP BY status_de_financiamento
    """), {"dt": str(dt_atual)})
    df_fin = pd.DataFrame(r.fetchall(), columns=r.keys())

    # 4. PDD por status de obra
    r = conn.execute(text("""
        SELECT status_de_obra,
               SUM(valor_pdd) AS vat, SUM(valor_pdd_anterior) AS vant
        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR WHERE dt_fechamento = :dt
        GROUP BY status_de_obra
    """), {"dt": str(dt_atual)})
    df_obra = pd.DataFrame(r.fetchall(), columns=r.keys())

    # 5. Movimentação de aging — qtd por status atual vs anterior (agrupado por status)
    r = conn.execute(text("""
        SELECT aging_da_unidade                   AS status_atual,
               COUNT(*)                           AS qtd_atual,
               SUM(CASE WHEN aging_da_unidade_anterior = aging_da_unidade THEN 1 ELSE 0 END) AS qtd_igual,
               SUM(CASE WHEN aging_da_unidade_anterior != aging_da_unidade OR aging_da_unidade_anterior IS NULL THEN 1 ELSE 0 END) AS qtd_mudou
        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR WHERE dt_fechamento = :dt
        GROUP BY aging_da_unidade
    """), {"dt": str(dt_atual)})
    df_aging_at = pd.DataFrame(r.fetchall(), columns=r.keys())

    r = conn.execute(text("""
        SELECT aging_da_unidade_anterior          AS status_ant,
               COUNT(*)                           AS qtd_ant
        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR WHERE dt_fechamento = :dt
          AND aging_da_unidade_anterior IS NOT NULL
        GROUP BY aging_da_unidade_anterior
    """), {"dt": str(dt_atual)})
    df_aging_ant = pd.DataFrame(r.fetchall(), columns=r.keys())

    # 6. Movimentação de financiamento — qtd por status atual vs anterior
    r = conn.execute(text("""
        SELECT status_de_financiamento            AS status_atual,
               COUNT(*)                           AS qtd_atual,
               SUM(CASE WHEN status_de_financiamento_ant = status_de_financiamento THEN 1 ELSE 0 END) AS qtd_igual,
               SUM(CASE WHEN status_de_financiamento_ant != status_de_financiamento OR status_de_financiamento_ant IS NULL THEN 1 ELSE 0 END) AS qtd_mudou
        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR WHERE dt_fechamento = :dt
          AND status_de_financiamento IS NOT NULL
        GROUP BY status_de_financiamento
    """), {"dt": str(dt_atual)})
    df_fin_at = pd.DataFrame(r.fetchall(), columns=r.keys())

    r = conn.execute(text("""
        SELECT status_de_financiamento_ant        AS status_ant,
               COUNT(*)                           AS qtd_ant
        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR WHERE dt_fechamento = :dt
          AND status_de_financiamento_ant IS NOT NULL
          AND status_de_financiamento     IS NOT NULL
        GROUP BY status_de_financiamento_ant
    """), {"dt": str(dt_atual)})
    df_fin_ant = pd.DataFrame(r.fetchall(), columns=r.keys())

    # 7. Empreendimentos que saíram de Em construção para Pronto
    r = conn.execute(text("""
        SELECT DISTINCT empreendimento, regional
        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR
        WHERE dt_fechamento = :dt
          AND status_de_obra = 'Pronto'
          AND status_de_obra_anterior = 'Em construção'
    """), {"dt": str(dt_atual)})
    df_prontos = pd.DataFrame(r.fetchall(), columns=r.keys())

    # 8. Dataset AGRUPADO para filtros JS (muito menor que o completo)
    r = conn.execute(text("""
        SELECT empreendimento, aging_da_unidade, aging_da_unidade_anterior,
               status_de_financiamento, status_de_financiamento_ant,
               status_de_obra, regional,
               COUNT(*)                             AS qtd,
               CAST(SUM(valor_pdd)          AS FLOAT) AS valor_pdd_atual,
               CAST(SUM(valor_pdd_anterior) AS FLOAT) AS valor_pdd_anterior
        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR WHERE dt_fechamento = :dt
        GROUP BY empreendimento, aging_da_unidade, aging_da_unidade_anterior,
                 status_de_financiamento, status_de_financiamento_ant,
                 status_de_obra, regional
    """), {"dt": str(dt_atual)})
    df_full = pd.DataFrame(r.fetchall(), columns=r.keys())

print(f"📊 Dados carregados | {len(df_prontos)} empreend. concluídos")

# =====================================================
# PREPARAR DATAFRAMES
# =====================================================

import json

def prep(df, col):
    df = df.rename(columns={"vat": "valor_pdd_atual", "vant": "valor_pdd_anterior"})
    for c in ["valor_pdd_atual", "valor_pdd_anterior"]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    df["eh_novo"]   = (df["valor_pdd_anterior"] == 0) & (df["valor_pdd_atual"] > 0)
    df["diferenca"] = df["valor_pdd_atual"] - df["valor_pdd_anterior"]
    df["pct"]       = df.apply(
        lambda r: 100.0 if r["eh_novo"]
        else ((r["diferenca"] / r["valor_pdd_anterior"] * 100) if r["valor_pdd_anterior"] != 0 else 0),
        axis=1
    )
    return df.sort_values("diferenca", ascending=False)

df_emp   = prep(df_emp,   "empreendimento")
df_aging = prep(df_aging, "aging_da_unidade")
df_fin   = prep(df_fin,   "status_de_financiamento")
df_obra  = prep(df_obra,  "status_de_obra")

total_ant = df_emp["valor_pdd_anterior"].sum()
total_at  = df_emp["valor_pdd_atual"].sum()
total_dif = total_at - total_ant
total_pct = (total_dif / total_ant * 100) if total_ant != 0 else 0

# Movimentação — merge atual vs anterior em uma linha por status
def prep_mov(df_at, df_ant, col_status):
    df = pd.merge(
        df_at.rename(columns={"status_atual": col_status}),
        df_ant.rename(columns={"status_ant": col_status, "qtd_ant": "qtd_anterior"}),
        on=col_status, how="outer"
    ).fillna(0)
    df["qtd_atual"]    = pd.to_numeric(df["qtd_atual"],    errors="coerce").fillna(0).astype(int)
    df["qtd_anterior"] = pd.to_numeric(df["qtd_anterior"], errors="coerce").fillna(0).astype(int)
    df["qtd_mudou"]    = pd.to_numeric(df.get("qtd_mudou",   0), errors="coerce").fillna(0).astype(int)
    df["diferenca"]    = df["qtd_atual"] - df["qtd_anterior"]
    df["pct"]          = df.apply(
        lambda r: 100.0 if r["qtd_anterior"] == 0 and r["qtd_atual"] > 0
        else ((r["diferenca"] / r["qtd_anterior"] * 100) if r["qtd_anterior"] != 0 else 0),
        axis=1
    )
    df["eh_novo"] = (df["qtd_anterior"] == 0) & (df["qtd_atual"] > 0)
    return df.sort_values("diferenca", ascending=False)

df_aging_mov = prep_mov(df_aging_at, df_aging_ant, "aging_da_unidade")
df_fin_mov   = prep_mov(df_fin_at,   df_fin_ant,   "status_de_financiamento")

# Filtros — apenas Regional, Empreendimento e Status de Obra
emps  = sorted(df_emp["empreendimento"].dropna().unique().tolist())
obras = sorted(df_obra["status_de_obra"].dropna().unique().tolist())
regs  = sorted(df_full["regional"].dropna().unique().tolist())

print("✅ Dados processados!")

# =====================================================
# HELPERS HTML
# =====================================================

def fmt_brl(v):
    neg = v < 0
    p = f"{abs(v):.2f}".split(".")
    intp, out = p[0], ""
    for k, ch in enumerate(reversed(intp)):
        if k > 0 and k % 3 == 0: out = "." + out
        out = ch + out
    s = f"R$ {out},{p[1]}"
    return f"-{s}" if neg else s

def badge_pct(pct, novo=False):
    if novo:
        return '<span class="badge badge-new">&#9733; Novo</span>'
    if pct == 0:
        return '<span class="badge neutral">0,0%</span>'
    cls = "badge-up" if pct > 0 else "badge-down"
    arr = "&#9650;" if pct > 0 else "&#9660;"
    return f'<span class="badge {cls}">{arr} {abs(pct):.1f}%</span>'

def row_class(dif, novo=False):
    if novo:   return "row-new"
    if dif > 0: return "row-up"
    if dif < 0: return "row-down"
    return ""

def sel_opts(lst):
    return '<option value="">Todos</option>' + "".join([f'<option value="{x}">{x}</option>' for x in lst])

# =====================================================
# TABELA PDD
# =====================================================

def tabela_pdd(df_t, col, titulo, cid):
    rows = ""
    for _, r in df_t.iterrows():
        novo = bool(r["eh_novo"])
        cls  = row_class(r["diferenca"], novo)
        cd   = "#c0392b" if r["diferenca"] > 0 else ("#1a7a4a" if r["diferenca"] < 0 else "#666")
        if novo: cd = "#1B3A8C"
        nota = '<span style="font-size:10px;color:#1B3A8C;margin-left:6px">&#9432; sem saldo anterior</span>' if novo else ""
        rows += (f'<tr class="{cls}" data-val="{r[col]}"'
                 f' data-ant="{r["valor_pdd_anterior"]}" data-at="{r["valor_pdd_atual"]}">'
                 f'<td>{r[col]}{nota}</td>'
                 f'<td class="num">{fmt_brl(r["valor_pdd_anterior"])}</td>'
                 f'<td class="num">{fmt_brl(r["valor_pdd_atual"])}</td>'
                 f'<td class="num fw" style="color:{cd}">{fmt_brl(r["diferenca"])}</td>'
                 f'<td class="num">{badge_pct(r["pct"], novo)}</td></tr>')
    # Linha de total
    t_ant = df_t["valor_pdd_anterior"].sum()
    t_at  = df_t["valor_pdd_atual"].sum()
    t_dif = t_at - t_ant
    t_pct = (t_dif / t_ant * 100) if t_ant != 0 else 0
    t_cd  = "#c0392b" if t_dif > 0 else ("#1a7a4a" if t_dif < 0 else "#666")
    total_row = (f'<tr class="total-row" id="{cid}-total">'
                 f'<td><strong>Total Geral</strong></td>'
                 f'<td class="num"><strong>{fmt_brl(t_ant)}</strong></td>'
                 f'<td class="num"><strong>{fmt_brl(t_at)}</strong></td>'
                 f'<td class="num fw" style="color:{t_cd}"><strong>{fmt_brl(t_dif)}</strong></td>'
                 f'<td class="num">{badge_pct(t_pct)}</td></tr>')
    return (f'<div class="card" id="{cid}">'
            f'<div class="card-hdr"><span class="card-ico">&#9670;</span>{titulo}</div>'
            f'<div class="tw"><table>'
            f'<thead><tr><th>{titulo}</th>'
            f'<th class="num">{nome_anterior}</th>'
            f'<th class="num">{nome_atual}</th>'
            f'<th class="num">Diferença</th>'
            f'<th class="num">Variação</th>'
            f'</tr></thead><tbody>{rows}{total_row}</tbody></table></div></div>')

# =====================================================
# TABELA MOVIMENTAÇÃO — mesmo padrão do detalhamento
# Status | Qtd Anterior | Qtd Atual | Diferença | %
# =====================================================

def tabela_mov(df_m, col, titulo, cid):
    rows = ""
    for _, r in df_m.iterrows():
        novo = bool(r["eh_novo"])
        dif  = int(r["diferenca"])
        pct  = r["pct"]
        if novo:      cls = "row-new"
        elif dif > 0: cls = "row-up"
        elif dif < 0: cls = "row-down"
        else:         cls = ""
        if novo:      cd = "#1B3A8C"
        elif dif > 0: cd = "#c0392b"
        elif dif < 0: cd = "#1a7a4a"
        else:         cd = "#666"
        bdg     = badge_pct(pct, novo)
        nota    = '<span style="font-size:10px;color:#1B3A8C;margin-left:6px">&#9432; sem saldo anterior</span>' if novo else ""
        ant_fmt = f'{int(r["qtd_anterior"]):,}' if not novo else "—"
        dif_fmt = (f'+{dif:,}' if dif > 0 else f'{dif:,}')
        rows += (f'<tr class="{cls}" data-val="{r[col]}"'
                 f' data-ant="{int(r["qtd_anterior"])}" data-at="{int(r["qtd_atual"])}">'
                 f'<td>{r[col]}{nota}</td>'
                 f'<td class="num">{ant_fmt}</td>'
                 f'<td class="num fw">{int(r["qtd_atual"]):,}</td>'
                 f'<td class="num fw" style="color:{cd}">{dif_fmt}</td>'
                 f'<td class="num">{bdg}</td></tr>')
    # Total
    t_ant = int(df_m["qtd_anterior"].sum())
    t_at  = int(df_m["qtd_atual"].sum())
    t_dif = t_at - t_ant
    t_pct = (t_dif / t_ant * 100) if t_ant != 0 else 0
    t_cd  = "#c0392b" if t_dif > 0 else ("#1a7a4a" if t_dif < 0 else "#666")
    t_dif_fmt = f'+{t_dif:,}' if t_dif > 0 else f'{t_dif:,}'
    total_row = (f'<tr class="total-row" id="{cid}-total">'
                 f'<td><strong>Total Geral</strong></td>'
                 f'<td class="num"><strong>{t_ant:,}</strong></td>'
                 f'<td class="num fw"><strong>{t_at:,}</strong></td>'
                 f'<td class="num fw" style="color:{t_cd}"><strong>{t_dif_fmt}</strong></td>'
                 f'<td class="num">{badge_pct(t_pct)}</td></tr>')
    return (f'<div class="card" id="{cid}">'
            f'<div class="card-hdr"><span class="card-ico">&#9670;</span>{titulo}</div>'
            f'<div class="tw"><table>'
            f'<thead><tr><th>{titulo}</th>'
            f'<th class="num">{nome_anterior}</th>'
            f'<th class="num">{nome_atual}</th>'
            f'<th class="num">Diferença</th>'
            f'<th class="num">Variação</th>'
            f'</tr></thead><tbody>{rows}{total_row}</tbody></table></div></div>')

# =====================================================
# MONTAR HTML
# =====================================================

tab_emp   = tabela_pdd(df_emp,   "empreendimento",         "Por Empreendimento",        "card-emp")
tab_aging = tabela_pdd(df_aging, "aging_da_unidade",       "Por Aging da Unidade",      "card-aging")
tab_fin   = tabela_pdd(df_fin,   "status_de_financiamento","Por Status de Financiamento","card-fin")
tab_obra  = tabela_pdd(df_obra,  "status_de_obra",         "Por Status de Obra",        "card-obra")

tab_aging_mov = tabela_mov(df_aging_mov, "aging_da_unidade",
    f"Movimentação de Aging — Unidades por Status", "card-aging-mov")
tab_fin_mov   = tabela_mov(df_fin_mov, "status_de_financiamento",
    f"Movimentação de Financiamento — Unidades por Status", "card-fin-mov")

# KPI cards
def kpi_cards(df_k, tipo):
    cor = "#c0392b" if tipo == "up" else "#1a7a4a"
    ico = "&#9650;" if tipo == "up" else "&#9660;"
    html = ""
    for _, r in df_k.iterrows():
        novo = bool(r["eh_novo"])
        cor_k = "#1B3A8C" if novo else cor
        html += (f'<div class="kpi" style="border-top:3px solid {cor_k}">'
                 f'<div class="kpi-ico" style="color:{cor_k}">{ico}</div>'
                 f'<div class="kpi-nm">{r["empreendimento"]}</div>'
                 f'<div class="kpi-vl" style="color:{cor_k}">{fmt_brl(r["diferenca"])}</div>'
                 f'<div>{badge_pct(r["pct"], novo)}</div></div>')
    return html

cards_up   = kpi_cards(df_emp.nlargest(3,  "diferenca"), "up")
cards_down = kpi_cards(df_emp.nsmallest(3, "diferenca"), "down")
cor_total  = "#c0392b" if total_dif > 0 else "#1a7a4a"

# Novos prontos
np_html = ""
if len(df_prontos) > 0:
    items = "".join(
        f'<div class="np-item"><span class="np-reg">{r["regional"]}</span>{r["empreendimento"]}</div>'
        for _, r in df_prontos.iterrows()
    )
    np_html = (f'<div class="card np-card">'
               f'<div class="card-hdr"><span class="card-ico">&#127959;</span>'
               f'Empreendimentos Concluídos &mdash; Em construção &rarr; Pronto ({len(df_prontos)})</div>'
               f'<div class="np-grid">{items}</div></div>')

# JSON para filtros — garantir float para somas no JS
df_full["valor_pdd_atual"]    = pd.to_numeric(df_full["valor_pdd_atual"],    errors="coerce").fillna(0.0)
df_full["valor_pdd_anterior"] = pd.to_numeric(df_full["valor_pdd_anterior"], errors="coerce").fillna(0.0)
data_json = df_full[["empreendimento","aging_da_unidade","aging_da_unidade_anterior",
                      "status_de_financiamento","status_de_financiamento_ant",
                      "status_de_obra","regional","qtd",
                      "valor_pdd_atual","valor_pdd_anterior"]
                    ].to_json(orient="records", force_ascii=False, double_precision=4)

# =====================================================
# HTML
# =====================================================

html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Relatório PDD &mdash; {nome_anterior} &times; {nome_atual}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Sora:wght@700;800&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
:root{{--cury:#1B3A8C;--cury-d:#0f2460;--cury-l:#e8edf8;--up:#c0392b;--up-l:#fdecea;--dn:#1a7a4a;--dn-l:#e6f7ee;--mv:#5c35c9;--mv-l:#ede7ff;--new:#1B3A8C;--new-l:#dde8ff;--bg:#f0f2f7;--sf:#fff;--bd:#dde2ef;--tx:#111827;--mt:#6b7280}}
body{{font-family:'Inter',sans-serif;background:var(--bg);color:var(--tx);font-size:14px}}
.topbar{{background:var(--cury-d);padding:0 40px;display:flex;align-items:center;justify-content:space-between;height:56px}}
.topbar img{{height:30px}}
.topbar-r{{font-size:11px;color:rgba(255,255,255,.4);letter-spacing:1px;text-transform:uppercase}}
.hdr{{background:var(--cury);color:#fff;padding:36px 40px 28px;position:relative;overflow:hidden}}
.hdr::after{{content:'';position:absolute;right:-60px;top:-60px;width:280px;height:280px;border-radius:50%;background:rgba(255,255,255,.04);pointer-events:none}}
.hdr-lbl{{font-size:11px;font-weight:600;letter-spacing:3px;text-transform:uppercase;color:rgba(255,255,255,.5);margin-bottom:10px}}
.hdr-ttl{{font-family:'Sora',sans-serif;font-size:30px;font-weight:800;line-height:1.15;margin-bottom:6px}}
.hdr-sub{{font-size:13px;color:rgba(255,255,255,.55);margin-bottom:28px}}
.periodo{{display:flex;align-items:center;gap:12px}}
.pb{{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.2);border-radius:6px;padding:7px 16px;font-size:13px;font-weight:600}}
.pa{{color:rgba(255,255,255,.4);font-size:16px}}
.pn{{font-size:11px;color:rgba(255,255,255,.3)}}
.filters{{background:var(--sf);border-bottom:2px solid var(--cury-l);padding:12px 40px;display:flex;gap:16px;flex-wrap:wrap;align-items:center}}
.filters label{{font-size:10px;font-weight:700;color:var(--mt);text-transform:uppercase;letter-spacing:1px;white-space:nowrap}}
.filters select{{border:1px solid var(--bd);border-radius:6px;padding:5px 10px;font-size:12px;font-family:'Inter',sans-serif;color:var(--tx);background:#fafbff;cursor:pointer;outline:none;min-width:130px}}
.filters select:focus{{border-color:var(--cury)}}
.btn-r{{background:var(--cury);color:#fff;border:none;border-radius:6px;padding:7px 16px;font-size:12px;font-weight:600;cursor:pointer;margin-left:auto;white-space:nowrap}}
.btn-r:hover{{background:var(--cury-d)}}
.fbadge{{background:var(--cury-l);color:var(--cury);border-radius:12px;padding:2px 8px;font-size:11px;font-weight:700;display:none;white-space:nowrap}}
.sumario{{display:grid;grid-template-columns:repeat(3,1fr);gap:1px;background:var(--bd);border-bottom:1px solid var(--bd)}}
.sum{{background:var(--sf);padding:22px 32px}}
.sum-lbl{{font-size:10px;font-weight:600;letter-spacing:2px;text-transform:uppercase;color:var(--mt);margin-bottom:8px}}
.sum-val{{font-family:'Sora',sans-serif;font-size:22px;font-weight:700}}
.sum-sub{{font-size:12px;color:var(--mt);margin-top:4px}}
.sec{{padding:28px 40px 0}}
.sec-ttl{{font-size:10px;font-weight:700;letter-spacing:3px;text-transform:uppercase;color:var(--mt);margin-bottom:16px;display:flex;align-items:center;gap:10px}}
.sec-ttl::after{{content:'';flex:1;height:1px;background:var(--bd)}}
.kgrid{{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:28px}}
.kpi{{background:var(--sf);border:1px solid var(--bd);border-radius:10px;padding:16px 18px}}
.kpi-ico{{font-size:15px;margin-bottom:6px}}
.kpi-nm{{font-size:12px;font-weight:500;line-height:1.3;margin-bottom:5px}}
.kpi-vl{{font-family:'Sora',sans-serif;font-size:16px;font-weight:700;margin-bottom:5px}}
.ca{{padding:0 40px 40px}}
.card{{background:var(--sf);border:1px solid var(--bd);border-radius:10px;margin-bottom:18px;overflow:hidden}}
.card-hdr{{display:flex;align-items:center;gap:8px;padding:13px 18px;border-bottom:1px solid var(--bd);font-family:'Sora',sans-serif;font-size:14px;font-weight:700}}
.card-ico{{color:var(--mt);font-size:13px}}
.tw{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
thead tr{{background:#f5f7ff;border-bottom:2px solid var(--bd)}}
th{{padding:9px 14px;font-size:10px;font-weight:700;letter-spacing:1.5px;text-transform:uppercase;color:var(--mt);white-space:nowrap}}
td{{padding:9px 14px;border-bottom:1px solid #f0f2f7}}
tr:last-child td{{border-bottom:none}}
tr.row-up{{background:var(--up-l)}}
tr.row-down{{background:var(--dn-l)}}
tr.row-move{{background:var(--mv-l)}}
tr.row-new{{background:var(--new-l)}}
tr:hover td{{filter:brightness(.97)}}
.num{{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}}
.fw{{font-weight:700}}
.badge{{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:700}}
.badge-up{{background:var(--up-l);color:var(--up)}}
.badge-down{{background:var(--dn-l);color:var(--dn)}}
.badge-new{{background:var(--new-l);color:var(--new)}}
.badge.neutral{{background:#f0f0f0;color:#888}}
.np-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:10px;padding:14px 18px}}
.np-item{{background:var(--cury-l);border:1px solid #c5d0e8;border-radius:8px;padding:9px 13px;font-size:13px;font-weight:500;display:flex;align-items:center;gap:10px}}
.np-reg{{background:var(--cury);color:#fff;border-radius:4px;padding:2px 7px;font-size:10px;font-weight:700;white-space:nowrap}}
.footer{{background:var(--cury-d);color:rgba(255,255,255,.4);padding:18px 40px;font-size:11px;letter-spacing:1px;display:flex;justify-content:space-between}}
.hidden{{display:none!important}}
tr.total-row{{background:#f5f7ff;border-top:2px solid var(--bd)}}
tr.total-row td{{font-weight:700;padding:10px 14px}}
@media print{{.filters{{display:none}}.hdr::after{{display:none}}}}
</style>
</head>
<body>
<div class="topbar">
  <img src="{logo_src}" alt="Cury">
  <div class="topbar-r">Relatório Interno &middot; Provisão para Devedores Duvidosos</div>
</div>
<div class="hdr">
  <div class="hdr-lbl">Provisão para Devedores Duvidosos</div>
  <div class="hdr-ttl">Análise Comparativa de PDD</div>
  <div class="hdr-sub">Variação do valor provisionado por empreendimento, aging, financiamento e status de obra</div>
  <div class="periodo">
    <div class="pb">{nome_anterior}</div><div class="pa">&rarr;</div><div class="pb">{nome_atual}</div>
    <div class="pn">&middot; Mês anterior identificado automaticamente</div>
  </div>
</div>
<div class="filters" id="filters">
  <label>Regional</label><select id="f-reg" onchange="aF()">{sel_opts(regs)}</select>
  <label>Empreendimento</label><select id="f-emp" onchange="aF()">{sel_opts(emps)}</select>
  <label>Status de Obra</label><select id="f-obra" onchange="aF()">{sel_opts(obras)}</select>
  <span class="fbadge" id="fbadge"></span>
  <button class="btn-r" onclick="rF()">&#8635; Limpar</button>
</div>
<div class="sumario">
  <div class="sum"><div class="sum-lbl">PDD {nome_anterior}</div><div class="sum-val" id="s-ant">{fmt_brl(total_ant)}</div><div class="sum-sub">Total provisionado anterior</div></div>
  <div class="sum"><div class="sum-lbl">PDD {nome_atual}</div><div class="sum-val" id="s-at">{fmt_brl(total_at)}</div><div class="sum-sub">Total provisionado atual</div></div>
  <div class="sum"><div class="sum-lbl">Variação Total</div><div class="sum-val" id="s-dif" style="color:{cor_total}">{fmt_brl(total_dif)}</div><div class="sum-sub" id="s-pct">{badge_pct(total_pct)} em relação ao mês anterior</div></div>
</div>
<div class="sec">
  <div class="sec-ttl">Maiores Aumentos de PDD &mdash; Empreendimentos</div>
  <div class="kgrid">{cards_up}</div>
  <div class="sec-ttl">Maiores Reduções de PDD &mdash; Empreendimentos</div>
  <div class="kgrid">{cards_down}</div>
</div>
<div class="ca" style="padding-top:28px">
  <div class="sec-ttl" style="margin-bottom:18px">Detalhamento por Visão <span id="finfo" style="font-weight:400;color:#aaa;font-size:11px;letter-spacing:0;text-transform:none;margin-left:6px"></span></div>
  {tab_emp}{tab_aging}{tab_fin}{tab_obra}
  <div class="sec-ttl" style="margin:28px 0 18px">Movimentação no Período</div>
  {tab_aging_mov}{tab_fin_mov}{np_html}
</div>
<div class="footer">
  <span>CURY CONSTRUTORA &middot; RELATÓRIO PDD &mdash; CONFIDENCIAL</span>
  <span>Gerado em {datetime.now().strftime("%d/%m/%Y às %H:%M")} &middot; Competência {nome_atual}</span>
</div>
<script>
const D={data_json};

// ── Formatadores ──
function fB(v){{
  const n=v<0,a=Math.abs(v),p=a.toFixed(2).split(".");
  const i=p[0].split("").reverse().map((c,k)=>k>0&&k%3==0?c+".":c).reverse().join("");
  return(n?"-":"")+"R$ "+i+","+p[1];
}}
function fN(v){{
  const n=v<0,a=Math.abs(Math.round(v));
  const s=a.toString().split("").reverse().map((c,k)=>k>0&&k%3==0?c+".":c).reverse().join("");
  return(n?"-":"")+s;
}}
function bJ(p){{
  if(p===0)return'<span class="badge neutral">0,0%</span>';
  const c=p>0?"badge-up":"badge-down",a=p>0?"&#9650;":"&#9660;";
  return`<span class="badge ${{c}}">${{a}} ${{Math.abs(p).toFixed(1)}}%</span>`;
}}
function rowCls(dif){{
  if(dif>0)return"row-up";if(dif<0)return"row-down";return"";
}}
function corDif(dif){{
  return dif>0?"#c0392b":dif<0?"#1a7a4a":"#666";
}}

// ── Agrupamento do dataset filtrado ──
function groupBy(data, key){{
  const map={{}};
  data.forEach(r=>{{
    const k=r[key]||"(vazio)";
    if(!map[k])map[k]={{at:0,ant:0}};
    map[k].at +=(r.valor_pdd_atual   ||0);
    map[k].ant+=(r.valor_pdd_anterior||0);
  }});
  return Object.entries(map)
    .map(([k,v])=>{{const d=v.at-v.ant;const p=v.ant!==0?d/v.ant*100:(v.at>0?100:0);return{{k,at:v.at,ant:v.ant,dif:d,pct:p}};}})
    .sort((a,b)=>b.dif-a.dif);
}}

// ── Reescreve tbody de uma tabela PDD ──
function rewritePDD(cid, rows){{
  const tbody=document.querySelector("#"+cid+" tbody");
  if(!tbody)return;
  let html="";
  let tAt=0,tAnt=0;
  rows.forEach(r=>{{
    tAt+=r.at;tAnt+=r.ant;
    const cls=rowCls(r.dif),cd=corDif(r.dif);
    const dFmt=fB(r.dif);
    html+=`<tr class="${{cls}}" data-val="${{r.k}}" data-ant="${{r.ant}}" data-at="${{r.at}}">
      <td>${{r.k}}</td>
      <td class="num">${{fB(r.ant)}}</td>
      <td class="num">${{fB(r.at)}}</td>
      <td class="num fw" style="color:${{cd}}">${{dFmt}}</td>
      <td class="num">${{bJ(r.pct)}}</td></tr>`;
  }});
  const tDif=tAt-tAnt,tP=tAnt!==0?tDif/tAnt*100:0,co=corDif(tDif);
  html+=`<tr class="total-row" id="${{cid}}-total">
    <td><strong>Total Geral</strong></td>
    <td class="num"><strong>${{fB(tAnt)}}</strong></td>
    <td class="num"><strong>${{fB(tAt)}}</strong></td>
    <td class="num fw" style="color:${{co}}"><strong>${{fB(tDif)}}</strong></td>
    <td class="num">${{bJ(tP)}}</td></tr>`;
  tbody.innerHTML=html;
}}

// ── Reescreve tbody de uma tabela Movimentação ──
// Soma qtd (contagem de unidades) por status atual e anterior
function rewriteMov(cid, data, colAt, colAnt){{
  const tbody=document.querySelector("#"+cid+" tbody");
  if(!tbody)return;
  // soma qtd por status ATUAL
  const mapAt={{}};
  data.forEach(r=>{{
    const k=r[colAt]||null;
    if(k===null)return;
    mapAt[k]=(mapAt[k]||0)+(r.qtd||1);
  }});
  // soma qtd por status ANTERIOR
  const mapAnt={{}};
  data.forEach(r=>{{
    const k=r[colAnt]||null;
    if(k===null)return;
    mapAnt[k]=(mapAnt[k]||0)+(r.qtd||1);
  }});
  // une os dois
  const keys=new Set([...Object.keys(mapAt),...Object.keys(mapAnt)]);
  let rows=[...keys].map(k=>{{
    const at=mapAt[k]||0,ant=mapAnt[k]||0,dif=at-ant;
    const pct=ant!==0?dif/ant*100:(at>0?100:0);
    return{{k,at,ant,dif,pct}};
  }}).sort((a,b)=>b.dif-a.dif);

  let html="",tAt=0,tAnt=0;
  rows.forEach(r=>{{
    tAt+=r.at;tAnt+=r.ant;
    const cls=rowCls(r.dif),cd=corDif(r.dif);
    const antFmt=r.ant>0?fN(r.ant):"—";
    const dFmt=(r.dif>=0?"+":"")+fN(r.dif);
    html+=`<tr class="${{cls}}" data-val="${{r.k}}" data-ant="${{r.ant}}" data-at="${{r.at}}">
      <td>${{r.k}}</td>
      <td class="num">${{antFmt}}</td>
      <td class="num fw">${{fN(r.at)}}</td>
      <td class="num fw" style="color:${{cd}}">${{dFmt}}</td>
      <td class="num">${{bJ(r.pct)}}</td></tr>`;
  }});
  const tDif=tAt-tAnt,tP=tAnt!==0?tDif/tAnt*100:0,co=corDif(tDif);
  const tDifFmt=(tDif>=0?"+":"")+fN(tDif);
  html+=`<tr class="total-row" id="${{cid}}-total">
    <td><strong>Total Geral</strong></td>
    <td class="num"><strong>${{fN(tAnt)}}</strong></td>
    <td class="num fw"><strong>${{fN(tAt)}}</strong></td>
    <td class="num fw" style="color:${{co}}"><strong>${{tDifFmt}}</strong></td>
    <td class="num">${{bJ(tP)}}</td></tr>`;
  tbody.innerHTML=html;
}}

// ── Filtro principal ──
function aF(){{
  const re=document.getElementById("f-reg").value,
        em=document.getElementById("f-emp").value,
        ob=document.getElementById("f-obra").value;
  const ac=[re,em,ob].filter(Boolean).length;
  const fb=document.getElementById("fbadge");
  fb.style.display=ac>0?"inline":"none";
  if(ac>0)fb.textContent=ac+" filtro"+(ac>1?"s":"")+" ativo"+(ac>1?"s":"");
  document.getElementById("finfo").textContent=ac>0?"("+ac+" filtro"+(ac>1?"s":"")+" ativo"+(ac>1?"s":"")+")" :"";

  // Dataset filtrado
  const f=D.filter(r=>
    (!re||r.regional===re)&&
    (!em||r.empreendimento===em)&&
    (!ob||r.status_de_obra===ob)
  );

  // Sumário
  const tA=f.reduce((s,r)=>s+(r.valor_pdd_atual||0),0);
  const tAnt=f.reduce((s,r)=>s+(r.valor_pdd_anterior||0),0);
  const tD=tA-tAnt,tP=tAnt!==0?tD/tAnt*100:0,co=corDif(tD);
  document.getElementById("s-ant").textContent=fB(tAnt);
  document.getElementById("s-at").textContent=fB(tA);
  document.getElementById("s-dif").textContent=fB(tD);
  document.getElementById("s-dif").style.color=co;
  document.getElementById("s-pct").innerHTML=bJ(tP)+" em relação ao mês anterior";

  // Reescreve todas as tabelas com dados filtrados
  rewritePDD("card-emp",   groupBy(f,"empreendimento"));
  rewritePDD("card-aging", groupBy(f,"aging_da_unidade"));
  rewritePDD("card-fin",   groupBy(f.filter(r=>r.status_de_financiamento),"status_de_financiamento"));
  rewritePDD("card-obra",  groupBy(f,"status_de_obra"));

  // Reescreve movimentação
  rewriteMov("card-aging-mov", f, "aging_da_unidade", "aging_da_unidade_anterior");
  rewriteMov("card-fin-mov",   f.filter(r=>r.status_de_financiamento), "status_de_financiamento", "status_de_financiamento_ant");
}}

function rF(){{
  document.querySelectorAll(".filters select").forEach(s=>s.value="");
  aF();
}}
</script>
</body></html>"""

Path(PASTA_SAIDA).mkdir(parents=True, exist_ok=True)
nome_arquivo  = f"relatorio_pdd_{dt_atual.strftime('%Y_%m')}.html"
caminho_saida = Path(PASTA_SAIDA) / nome_arquivo
with open(caminho_saida, "w", encoding="utf-8") as f:
    f.write(html)

print(f"\n{'='*55}")
print(f"  RELATORIO GERADO COM SUCESSO!")
print(f"{'='*55}")
print(f"  Comparativo : {nome_anterior} x {nome_atual}")
print(f"  Arquivo     : {nome_arquivo}")
print(f"  Local       : {PASTA_SAIDA}")
print(f"{'='*55}")