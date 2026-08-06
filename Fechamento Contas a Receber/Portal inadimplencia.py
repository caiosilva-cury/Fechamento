import warnings
warnings.filterwarnings("ignore")

import urllib
import json
import base64
import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime
from pathlib import Path

# =====================================================
# CONFIGURAÇÕES
# =====================================================

PASTA_SAIDA = r"C:\Users\srvautbicar\Desktop\PROCESSOS\import para SQL\Base PDD\Relatorios"
LOGO_PATH   = r"C:\Users\srvautbicar\Desktop\PROCESSOS\import para SQL\Base PDD\logo_cury.png"

print("=" * 60)
print("  PORTAL DE INADIMPLÊNCIA — GERAÇÃO DO RELATÓRIO HTML")
print("=" * 60)
print()

def pedir_data(label):
    while True:
        entrada = input(f"📅 {label}: ").strip()
        try:
            return datetime.strptime(entrada, "%Y-%m-%d").date()
        except ValueError:
            print("❌ Formato inválido. Use AAAA-MM-DD\n")

dt_atual = pedir_data("Mês de fechamento (ex: 2026-04-30)")

params = urllib.parse.quote_plus(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=SRJDLK01\\SSPDLKAZ01;"
    "DATABASE=P_OURO_CAR_DB;"
    "UID=powerbi;"
    "PWD=Ah79Wlk999;"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}", fast_executemany=True)
print("\n🔌 Conectando ao banco...")

try:
    with open(LOGO_PATH, "rb") as f:
        logo_src = "data:image/png;base64," + base64.b64encode(f.read()).decode()
except Exception:
    logo_src = ""

# ── últimos 3 meses ──
with engine.connect() as conn:
    r = conn.execute(text("""
        SELECT DISTINCT TOP 3 dt_fechamento FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR
        WHERE dt_fechamento <= :dt ORDER BY dt_fechamento DESC
    """), {"dt": str(dt_atual)})
    datas = [row[0] for row in r.fetchall()]

if len(datas) < 2:
    print(f"\n❌ Necessário pelo menos 2 meses. Encontrado: {len(datas)}")
    exit()

dt_at  = datas[0]
dt_ant = datas[1]
dt_ret = datas[2] if len(datas) >= 3 else None

MESES_PT = {1:"Janeiro",2:"Fevereiro",3:"Março",4:"Abril",5:"Maio",6:"Junho",
            7:"Julho",8:"Agosto",9:"Setembro",10:"Outubro",11:"Novembro",12:"Dezembro"}

def nome_mes(dt): return f"{MESES_PT[dt.month]}/{dt.year}"
def mes_label(dt): return f"{MESES_PT[dt.month][:3]}/{str(dt.year)[2:]}"

mes_at  = nome_mes(dt_at)
mes_ant = nome_mes(dt_ant)
mes_ret = nome_mes(dt_ret) if dt_ret else "—"
print(f"✅ Período identificado: {mes_ret} → {mes_ant} → {mes_at}")

print("📊 Carregando dados...")

with engine.connect() as conn:

    # ── A. DADOS GRANULARES POR EMPREENDIMENTO (3 meses) ──
    # Tudo que o JS precisa para recalcular qualquer painel ao filtrar

    r = conn.execute(text("""
        SELECT
            a.dt_fechamento,
            a.empreendimento,
            a.regional,
            a.status_de_obra,
            a.total_unid,
            a.unid_inad,
            a.saldo_dev,
            a.inadimplencia,
            a.rec_carteira,
            a.pdd,
            a.pdd_poc,
            a.pdd_contab,
            a.pdd_poc_contab,
            a.poc_medio,
            a.aging_dom,
            a.is_lancamento
        FROM (
            SELECT
                dt_fechamento,
                empreendimento,
                regional,
                MAX(status_de_obra)                                            AS status_de_obra,
                COUNT(*)                                                        AS total_unid,
                COUNT(CASE WHEN status_da_unidade='Inadimplente' THEN 1 END)   AS unid_inad,
                SUM(CAST(saldo_dev_carteira        AS FLOAT))                   AS saldo_dev,
                SUM(CAST(inadimplencia_carteira    AS FLOAT))                   AS inadimplencia,
                SUM(ISNULL(CAST(valor_recebido_carteira AS FLOAT), 0) + ISNULL(CAST(valor_recebido_ato AS FLOAT), 0))           AS rec_carteira,
                SUM(ISNULL(CAST(valor_pdd AS FLOAT), 0))                         AS pdd,
                SUM(ISNULL(CAST(valor_pdd_poc AS FLOAT), 0))                    AS pdd_poc,
                SUM(CAST(valor_pdd_contabilidade   AS FLOAT))                   AS pdd_contab,
                SUM(ISNULL(CAST(valor_pdd_poc_contabilidade AS FLOAT), 0))      AS pdd_poc_contab,
                AVG(CAST(poc_obras                 AS FLOAT))                   AS poc_medio,
                (
                    SELECT TOP 1 aging_da_unidade
                    FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR sub
                    WHERE sub.dt_fechamento   = FAT.TB_TRN_CAR_CARTEIRAORI_OUR.dt_fechamento
                      AND sub.empreendimento  = FAT.TB_TRN_CAR_CARTEIRAORI_OUR.empreendimento
                      AND sub.status_da_unidade = 'Inadimplente'
                      AND sub.aging_da_unidade IS NOT NULL
                      AND sub.aging_da_unidade NOT IN ('Quitado','Estoque')
                    GROUP BY sub.aging_da_unidade ORDER BY COUNT(*) DESC
                )                                                               AS aging_dom,
                CASE
                    WHEN YEAR(MIN(data_lancamento))  = YEAR(dt_fechamento)
                     AND MONTH(MIN(data_lancamento)) = MONTH(dt_fechamento)
                    THEN 1 ELSE 0
                END                                                             AS is_lancamento
            FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR
            WHERE dt_fechamento IN (:dt_at, :dt_ant, :dt_ret)
            GROUP BY dt_fechamento, empreendimento, regional
        ) a
        ORDER BY a.dt_fechamento, a.inadimplencia DESC
    """), {"dt_at": str(dt_at), "dt_ant": str(dt_ant),
           "dt_ret": str(dt_ret) if dt_ret else "1900-01-01"})
    df_emp = pd.DataFrame(r.fetchall(), columns=r.keys())

    # ── B. AGING POR EMPREENDIMENTO (3 meses) ──
    r = conn.execute(text("""
        SELECT
            dt_fechamento,
            empreendimento,
            regional,
            aging_da_unidade,
            COUNT(*)                                                       AS qtd_unid,
            SUM(CAST(inadimplencia_carteira AS FLOAT))                     AS saldo_inad,
            SUM(CAST(valor_pdd              AS FLOAT))                     AS pdd_gerado
        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR
        WHERE dt_fechamento IN (:dt_at, :dt_ant, :dt_ret)
          AND aging_da_unidade IS NOT NULL
          AND aging_da_unidade NOT IN ('Quitado','Estoque','Adimplente')
          AND (status_da_unidade='Inadimplente' OR CAST(inadimplencia_carteira AS FLOAT)>0)
        GROUP BY dt_fechamento, empreendimento, regional, aging_da_unidade
    """), {"dt_at": str(dt_at), "dt_ant": str(dt_ant),
           "dt_ret": str(dt_ret) if dt_ret else "1900-01-01"})
    df_aging_emp = pd.DataFrame(r.fetchall(), columns=r.keys())

    # ── C. SÉRIE HISTÓRICA POR EMPREENDIMENTO (12 meses) ──
    r = conn.execute(text("""
        SELECT
            dt_fechamento,
            empreendimento,
            regional,
            SUM(CAST(inadimplencia_carteira   AS FLOAT)) AS inad,
            SUM(CAST(saldo_dev_carteira       AS FLOAT)) AS saldo,
            SUM(CAST(valor_pdd                AS FLOAT)) AS pdd,
            SUM(ISNULL(CAST(valor_pdd_poc AS FLOAT), 0)) AS pdd_poc,
            COUNT(CASE WHEN status_da_unidade='Inadimplente' THEN 1 END) AS unid_inad
        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR
        WHERE dt_fechamento <= CAST(:dt AS DATE)
          AND dt_fechamento >= DATEADD(MONTH,-11,CAST(:dt AS DATE))
        GROUP BY dt_fechamento, empreendimento, regional
        ORDER BY dt_fechamento ASC
    """), {"dt": str(dt_at)})
    df_hist_emp = pd.DataFrame(r.fetchall(), columns=r.keys())

    # ── D. LANÇAMENTOS (dados extras p/ card) ──
    r = conn.execute(text("""
        SELECT
            empreendimento, regional,
            MIN(data_lancamento) AS data_lancamento,
            MAX(previsao_de_entrega) AS previsao_entrega
        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR
        WHERE dt_fechamento = :dt_at
        GROUP BY empreendimento, regional
    """), {"dt_at": str(dt_at)})
    df_lanc_info = pd.DataFrame(r.fetchall(), columns=r.keys())
    lanc_info = {row["empreendimento"]: {
        "dl": str(row["data_lancamento"])[:7] if row["data_lancamento"] else "—",
        "pe": str(row["previsao_entrega"] or "—")
    } for _, row in df_lanc_info.iterrows()}

print("✅ Dados carregados!")

# ── DIAGNÓSTICO: valor bruto do poc no banco ──
poc_sample = df_emp[df_emp["dt_fechamento"] == dt_at][["empreendimento","pdd_poc","pdd_poc_contab"]].head(5)
print("\n🔍 DIAGNÓSTICO — primeiros valores de pdd_poc no banco (mês atual):")
print(poc_sample.to_string(index=False))
total_poc = df_emp[df_emp["dt_fechamento"] == dt_at]["pdd_poc"].sum()
print(f"   Total pdd_poc mês atual: {total_poc}")
nan_count = df_emp[df_emp["dt_fechamento"] == dt_at]["pdd_poc"].isna().sum()
print(f"   Registros nulos/NaN: {nan_count}")
print()

# =====================================================
# MONTAR JSON MESTRE — tudo que o JS precisa
# =====================================================

def _aging_sort_key(s):
    try: return int(''.join(filter(str.isdigit, s.split()[0])))
    except: return 9999

# faixas de aging ordenadas
all_aging_faixas = sorted(
    {f for f in df_aging_emp["aging_da_unidade"].dropna().unique()
     if f not in ('Quitado','Estoque','Adimplente')},
    key=_aging_sort_key
)
_aging_cores = ["#EF9F27","#E24B4A","#C0392B","#A32D2D","#791F1F",
                "#601515","#4A0E0E","#330808","#1F0303","#0D0101"]
aging_cores_map = {b: _aging_cores[min(i,len(_aging_cores)-1)] for i,b in enumerate(all_aging_faixas)}

# datas únicas da histórica (labels)
hist_dts = sorted(df_hist_emp["dt_fechamento"].unique())
hist_labels = [mes_label(d) for d in hist_dts]

# converter datas para string para serialização
def dt_str(d): return str(d)[:10]

# dados de empreendimento por mês
emp_rows = []
for _, row in df_emp.iterrows():
    emp_rows.append({
        "dt":    dt_str(row["dt_fechamento"]),
        "emp":   str(row["empreendimento"]),
        "reg":   str(row["regional"]),
        "obra":  str(row["status_de_obra"] or "—"),
        "tu":    int(row["total_unid"] or 0),
        "ui":    int(row["unid_inad"] or 0),
        "sd":    round(float(row["saldo_dev"] or 0), 2),
        "inad":  round(float(row["inadimplencia"] or 0), 2),
        "rec":   round(float(row["rec_carteira"] or 0), 2),
        "pdd":   round(float(row["pdd"] or 0), 2),
        "poc":   (lambda v: 0.0 if v is None or (isinstance(v, float) and v != v) else round(float(v), 2))(row["pdd_poc"]),
        "pc":    round(float(row["pdd_contab"] or 0), 2),
        "pcc":   (lambda v: 0.0 if v is None or (isinstance(v, float) and v != v) else round(float(v), 2))(row["pdd_poc_contab"]),
        "pocm":  round(float(row["poc_medio"] or 0), 4),
        "aging": str(row["aging_dom"] or "—"),
        "lanc":  int(row["is_lancamento"] or 0),
    })

# aging por empreendimento/mês
aging_rows = []
for _, row in df_aging_emp.iterrows():
    aging_rows.append({
        "dt":    dt_str(row["dt_fechamento"]),
        "emp":   str(row["empreendimento"]),
        "reg":   str(row["regional"]),
        "faixa": str(row["aging_da_unidade"]),
        "qtd":   int(row["qtd_unid"] or 0),
        "si":    round(float(row["saldo_inad"] or 0), 2),
        "pdd":   round(float(row["pdd_gerado"] or 0), 2),
    })

# histórica por empreendimento
hist_rows = []
for _, row in df_hist_emp.iterrows():
    hist_rows.append({
        "dt":   dt_str(row["dt_fechamento"]),
        "emp":  str(row["empreendimento"]),
        "reg":  str(row["regional"]),
        "inad": round(float(row["inad"] or 0), 2),
        "sd":   round(float(row["saldo"] or 0), 2),
        "pdd":  round(float(row["pdd"] or 0), 2),
        "poc":  (lambda v: 0.0 if v is None or (isinstance(v, float) and v != v) else round(float(v), 2))(row["pdd_poc"]),
        "ui":   int(row["unid_inad"] or 0),
    })

# info de lançamentos (datas para card)
lanc_info_list = [{"emp": k, **v} for k, v in lanc_info.items()]

# regionais e empreendimentos para autocomplete
regionais_list = sorted(df_emp[df_emp["dt_fechamento"]==dt_at]["regional"].dropna().unique().tolist())
emps_por_regional = {}
for reg in regionais_list:
    emps = sorted(df_emp[(df_emp["dt_fechamento"]==dt_at) & (df_emp["regional"]==reg)]["empreendimento"].dropna().unique().tolist())
    emps_por_regional[reg] = emps

master_json = json.dumps({
    "dt_at":  dt_str(dt_at),
    "dt_ant": dt_str(dt_ant),
    "dt_ret": dt_str(dt_ret) if dt_ret else None,
    "mes_at":  mes_at,
    "mes_ant": mes_ant,
    "mes_ret": mes_ret,
    "hist_labels": hist_labels,
    "hist_dts":    [dt_str(d) for d in hist_dts],
    "aging_faixas": all_aging_faixas,
    "aging_cores":  [aging_cores_map[b] for b in all_aging_faixas],
    "regionais": regionais_list,
    "emps_por_regional": emps_por_regional,
    "emp_rows":   emp_rows,
    "aging_rows": aging_rows,
    "hist_rows":  hist_rows,
    "lanc_info":  lanc_info_list,
}, ensure_ascii=False)

print(f"📦 JSON master: {len(master_json)//1024} KB")

# =====================================================
# HTML
# =====================================================

html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Portal de Inadimplência — {mes_at}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Sora:wght@700;800&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
:root{{
  --cury:#1B3A8C;--cury-d:#0f2460;--cury-l:#e8edf8;
  --up:#c0392b;--up-l:#fdecea;
  --dn:#1a7a4a;--dn-l:#e6f7ee;
  --amb:#BA7517;--amb-l:#FAEEDA;
  --info:#185FA5;--info-l:#E6F1FB;
  --bg:#f0f2f7;--sf:#fff;--bd:#dde2ef;
  --tx:#111827;--mt:#6b7280
}}
body{{font-family:'Inter',sans-serif;background:var(--bg);color:var(--tx);font-size:14px}}
.topbar{{background:var(--cury-d);padding:0 36px;display:flex;align-items:center;justify-content:space-between;height:52px}}
.topbar img{{height:28px}}
.topbar-r{{font-size:11px;color:rgba(255,255,255,.35);letter-spacing:1px;text-transform:uppercase}}
.hdr{{background:var(--cury);color:#fff;padding:32px 36px 28px}}
.hdr-lbl{{font-size:11px;font-weight:600;letter-spacing:3px;text-transform:uppercase;color:rgba(255,255,255,.5);margin-bottom:8px}}
.hdr-ttl{{font-family:'Sora',sans-serif;font-size:26px;font-weight:800;margin-bottom:4px}}
.hdr-sub{{font-size:13px;color:rgba(255,255,255,.5)}}
/* ── FILTRO BAR ── */
.filter-bar{{background:var(--sf);border-bottom:2px solid var(--cury-l);padding:10px 36px;display:flex;align-items:center;gap:12px;flex-wrap:wrap}}
.filter-bar label{{font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--mt)}}
.fbtn-clear{{padding:5px 12px;border:1px solid var(--bd);border-radius:6px;font-size:12px;cursor:pointer;background:var(--sf);color:var(--mt);font-family:'Inter',sans-serif}}
.fbtn-clear:hover{{background:var(--bg)}}
.filter-tag{{background:var(--cury-l);color:var(--cury);padding:3px 10px;border-radius:20px;font-size:12px;font-weight:600;display:none}}
.filter-tag.on{{display:inline-block}}
.fi{{position:relative;display:inline-flex;align-items:center}}
.fi input{{padding:5px 28px 5px 10px;border:1px solid var(--bd);border-radius:6px;font-size:13px;font-family:'Inter',sans-serif;color:var(--tx);background:var(--sf);min-width:210px;outline:none}}
.fi input:focus{{border-color:var(--cury);box-shadow:0 0 0 2px rgba(27,58,140,.15)}}
.fi input::placeholder{{color:#aab}}
.fi .fi-clr{{position:absolute;right:7px;background:none;border:none;cursor:pointer;font-size:14px;color:var(--mt);padding:0;line-height:1;display:none}}
.fi .fi-clr.on{{display:block}}
.fi-drop{{position:absolute;top:calc(100% + 4px);left:0;background:#fff;border:1px solid var(--bd);border-radius:8px;box-shadow:0 8px 24px rgba(0,0,0,.12);z-index:500;min-width:250px;max-height:260px;overflow-y:auto;display:none}}
.fi-drop.on{{display:block}}
.fi-drop li{{list-style:none;padding:7px 12px;font-size:13px;cursor:pointer;white-space:nowrap}}
.fi-drop li:hover,.fi-drop li.hi{{background:var(--cury-l);color:var(--cury)}}
.fi-drop li.no-res{{color:var(--mt);font-style:italic;cursor:default}}
.fi-drop li.no-res:hover{{background:none;color:var(--mt)}}
/* ── TABS ── */
.tabs{{background:var(--sf);border-bottom:1px solid var(--bd);padding:0 36px;display:flex;gap:0;overflow-x:auto;align-items:center}}
.tab{{padding:13px 15px;font-size:13px;color:var(--mt);cursor:pointer;border:none;border-bottom:2px solid transparent;background:none;white-space:nowrap;font-family:'Inter',sans-serif}}
.tab.on{{color:var(--cury);border-bottom-color:var(--cury);font-weight:600}}
.tab-meto{{margin-left:auto;padding:7px 14px;font-size:12px;font-weight:600;color:var(--cury);border:1.5px solid var(--cury);border-radius:20px;cursor:pointer;background:var(--cury-l);white-space:nowrap;font-family:'Inter',sans-serif;border-bottom:1.5px solid var(--cury)!important}}
.tab-meto:hover{{background:var(--cury);color:#fff}}
.pnl{{display:none;padding:28px 36px 40px}}
.pnl.on{{display:block}}
/* ── KPIs ── */
.kg{{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:10px;margin-bottom:20px}}
.kc{{background:var(--sf);border:1px solid var(--bd);border-radius:10px;padding:14px 16px}}
.kc-lbl{{font-size:10px;font-weight:600;letter-spacing:1.5px;text-transform:uppercase;color:var(--mt);margin-bottom:6px}}
.kc-val{{font-family:'Sora',sans-serif;font-size:19px;font-weight:700;line-height:1.2}}
.kc-d{{font-size:12px;margin-top:3px}}
.ub{{color:var(--up)}} .dg{{color:var(--dn)}} .nm{{color:var(--mt)}}
/* ── MISC ── */
.sec{{font-size:10px;font-weight:700;letter-spacing:2.5px;text-transform:uppercase;color:var(--mt);margin:24px 0 14px;display:flex;align-items:center;gap:10px}}
.sec::after{{content:'';flex:1;height:1px;background:var(--bd)}}
.ins{{display:flex;gap:10px;padding:10px 14px;border-left:3px solid;margin-bottom:8px;font-size:13px;line-height:1.6;border-radius:0 6px 6px 0}}
.ins.g{{border-color:var(--dn);background:var(--dn-l)}}
.ins.r{{border-color:var(--up);background:var(--up-l)}}
.ins.b{{border-color:var(--info);background:var(--info-l)}}
.ins.a{{border-color:var(--amb);background:var(--amb-l)}}
.row2{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px}}
.card{{background:var(--sf);border:1px solid var(--bd);border-radius:10px;overflow:hidden}}
.card-hdr{{display:flex;align-items:center;gap:8px;padding:12px 16px;border-bottom:1px solid var(--bd);font-family:'Sora',sans-serif;font-size:13px;font-weight:700}}
.card-hdr .ico{{color:var(--mt);font-size:13px}}
.tw{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
thead tr{{background:#f5f7ff;border-bottom:2px solid var(--bd)}}
th{{padding:8px 10px;font-size:10px;font-weight:700;letter-spacing:1px;text-transform:uppercase;color:var(--mt);white-space:nowrap;text-align:left}}
td{{padding:8px 10px;border-bottom:1px solid #f0f2f7;vertical-align:middle}}
tr:last-child td{{border-bottom:none}}
.bb{{background:var(--bd);border-radius:3px;height:5px;margin-top:4px;min-width:60px}}
.bf{{height:5px;border-radius:3px}}
.emp-nm{{font-weight:500;color:var(--tx);font-size:13px}}
.emp-sub{{font-size:11px;color:var(--mt)}}
.bdg{{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600}}
.bdg-r{{background:var(--up-l);color:var(--up)}}
.bdg-g{{background:var(--dn-l);color:var(--dn)}}
.bdg-a{{background:var(--amb-l);color:var(--amb)}}
.bdg-b{{background:var(--info-l);color:var(--info)}}
.lc{{background:var(--sf);border:1px solid var(--bd);border-radius:10px;padding:14px 16px;margin-bottom:10px}}
.lc-head{{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:10px}}
.lc-nm{{font-size:14px;font-weight:600;font-family:'Sora',sans-serif}}
.lc-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:10px}}
.lc-item .lbl{{font-size:11px;color:var(--mt);margin-bottom:2px}}
.lc-item .vl{{font-weight:600;font-size:13px}}
.leg{{display:flex;flex-wrap:wrap;gap:12px;margin-bottom:10px;font-size:12px;color:var(--mt)}}
.leg span{{display:flex;align-items:center;gap:5px}}
.dot{{width:10px;height:10px;border-radius:2px;flex-shrink:0}}
.ch{{position:relative;width:100%}}
.frow{{display:flex;gap:6px;margin-bottom:14px;flex-wrap:wrap}}
.fb{{padding:5px 12px;border:1px solid var(--bd);border-radius:20px;font-size:12px;cursor:pointer;background:var(--sf);color:var(--mt);font-family:'Inter',sans-serif}}
.fb.on{{background:var(--cury);color:#fff;border-color:var(--cury)}}
.no-data{{color:var(--mt);font-style:italic;padding:16px 0;font-size:13px}}
/* ── MODAL ── */
.modal-overlay{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:1000;align-items:center;justify-content:center}}
.modal-overlay.on{{display:flex}}
.modal{{background:#fff;border-radius:14px;width:min(720px,94vw);max-height:88vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.25)}}
.modal-hdr{{padding:20px 24px 16px;border-bottom:1px solid var(--bd);display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;background:#fff;z-index:1}}
.modal-hdr h2{{font-family:'Sora',sans-serif;font-size:18px;font-weight:800;color:var(--cury)}}
.modal-close{{border:none;background:none;font-size:22px;cursor:pointer;color:var(--mt);line-height:1;padding:2px 6px;border-radius:4px}}
.modal-close:hover{{background:var(--bg)}}
.modal-body{{padding:20px 24px 28px}}
.meto-sec{{margin-bottom:22px}}
.meto-sec h3{{font-size:13px;font-weight:700;color:var(--cury);text-transform:uppercase;letter-spacing:1px;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid var(--cury-l)}}
.meto-row{{display:flex;gap:10px;margin-bottom:8px;font-size:13px;line-height:1.6}}
.meto-row .meto-lbl{{font-weight:600;min-width:180px;color:var(--tx);flex-shrink:0}}
.formula{{background:var(--cury-l);border-left:3px solid var(--cury);padding:6px 12px;border-radius:0 6px 6px 0;font-size:12px;font-family:monospace;margin:6px 0 10px}}
.footer{{background:var(--cury-d);color:rgba(255,255,255,.35);padding:16px 36px;font-size:11px;letter-spacing:1px;display:flex;justify-content:space-between;flex-wrap:wrap;gap:6px}}
@media print{{.tabs{{display:none}}.filter-bar{{display:none}}.pnl{{display:block!important;page-break-before:always}}}}
</style>
</head>
<body>

<!-- MODAL METODOLOGIA -->
<div class="modal-overlay" id="modal-meto" onclick="if(event.target===this)closeMeto()">
  <div class="modal">
    <div class="modal-hdr">
      <h2>&#9783; Metodologia dos Indicadores</h2>
      <button class="modal-close" onclick="closeMeto()">&#10005;</button>
    </div>
    <div class="modal-body">
      <div class="meto-sec">
        <h3>&#9632; Inadimplência &amp; Carteira</h3>
        <div class="meto-row"><span class="meto-lbl">Inadimplência carteira</span><span>Soma das parcelas vencidas não pagas da carteira própria (<code>inadimplencia_carteira</code>). Exclui encargos e financiamento de instituição.</span></div>
        <div class="formula">Inadimplência = SUM(inadimplencia_carteira)</div>
        <div class="meto-row"><span class="meto-lbl">% Inad. / Carteira</span><span>Percentual da inadimplência sobre o saldo devedor total no mês de fechamento.</span></div>
        <div class="formula">% Inad. / carteira = inadimplencia_carteira ÷ saldo_dev_carteira × 100</div>
        <div class="meto-row"><span class="meto-lbl">% Inad. / cobrança enviada</span><span>Percentual da inadimplência sobre o total cobrado no período: soma do que foi recebido mais o que ficou em aberto.</span></div>
        <div class="formula">% Inad. / cobrança = inadimplencia_carteira ÷ (valor_recebido_carteira + inadimplencia_carteira) × 100</div>
        <div class="meto-row"><span class="meto-lbl">Meta de referência</span><span>8,5% — limite interno. Acima de 12% = crítico; 8–12% = atenção; abaixo de 8% = saudável.</span></div>
      </div>
      <div class="meto-sec">
        <h3>&#9632; Aging da Carteira</h3>
        <div class="meto-row"><span class="meto-lbl">Faixas de aging</span><span>Agrupamento pelo número de dias em atraso (<code>aging_da_unidade</code>). Aging dominante = faixa com mais unidades inadimplentes no empreendimento.</span></div>
        <div class="meto-row"><span class="meto-lbl">Excluídos</span><span>Unidades com status Quitado, Estoque e Adimplente são excluídas.</span></div>
      </div>
      <div class="meto-sec">
        <h3>&#9632; PDD — Provisão para Devedores Duvidosos</h3>
        <div class="meto-row"><span class="meto-lbl">PDD contábil</span><span>Provisão pelo percentual de risco (<code>percentual_pdd_contabilidade</code>) sobre o saldo devedor total, independente do avanço físico.</span></div>
        <div class="formula">PDD contábil = saldo_dev_carteira × percentual_pdd_contabilidade</div>
        <div class="meto-row"><span class="meto-lbl">PDD sobre POC</span><span>Provisão proporcional ao avanço físico da obra (POC). Mais conservadora nos estágios iniciais.</span></div>
        <div class="formula">PDD POC = PDD contábil × poc_obras</div>
      </div>
      <div class="meto-sec">
        <h3>&#9632; Movimentação de PDD</h3>
        <div class="meto-row"><span class="meto-lbl">Constituição</span><span>Aumento do PDD — envelhecimento do atraso ou novas unidades inadimplentes.</span></div>
        <div class="meto-row"><span class="meto-lbl">PDD Recuperado</span><span>Redução do PDD — regularização (pagamento/acordo), distrato ou migração para faixa de menor risco.</span></div>
        <div class="formula">Variação PDD = valor_pdd (mês atual) − valor_pdd (mês anterior)</div>
      </div>
      <div class="meto-sec">
        <h3>&#9632; Periodicidade &amp; Fonte</h3>
        <div class="meto-row"><span class="meto-lbl">Fechamento</span><span>Dados consolidados mensalmente em <code>pdd_base</code> (banco <code>GestaoFinanceira</code>), snapshot por <code>dt_fechamento</code>.</span></div>
        <div class="meto-row"><span class="meto-lbl">Comparativos</span><span>Variações sempre em relação ao mês imediatamente anterior disponível. Série histórica cobre até 12 meses.</span></div>
      </div>
    </div>
  </div>
</div>

<div class="topbar">
  <img src="{logo_src}" alt="Cury">
  <span class="topbar-r">Portal de Inadimplência &middot; Uso interno</span>
</div>
<div class="hdr">
  <div class="hdr-lbl">Gestão Financeira &middot; PDD &middot; Carteira Residencial</div>
  <div class="hdr-ttl">Portal de Inadimplência — {mes_at}</div>
  <div class="hdr-sub">Empreendimentos · Lançamentos · PDD contábil vs POC · Tendência trimestral</div>
</div>

<!-- FILTROS -->
<div class="filter-bar">
  <label>&#9660; Filtrar por</label>
  <div class="fi" id="fi-reg">
    <input id="inp-regional" type="text" placeholder="&#128269; Regional..." autocomplete="off"
           oninput="fiInput('reg')" onfocus="fiOpen('reg')" onkeydown="fiKey(event,'reg')">
    <button class="fi-clr" id="clr-reg" onclick="fiClear('reg')" title="Limpar">&#10005;</button>
    <ul class="fi-drop" id="drop-reg"></ul>
  </div>
  <div class="fi" id="fi-emp">
    <input id="inp-emp" type="text" placeholder="&#128269; Empreendimento..." autocomplete="off"
           oninput="fiInput('emp')" onfocus="fiOpen('emp')" onkeydown="fiKey(event,'emp')">
    <button class="fi-clr" id="clr-emp" onclick="fiClear('emp')" title="Limpar">&#10005;</button>
    <ul class="fi-drop" id="drop-emp"></ul>
  </div>
  <button class="fbtn-clear" onclick="clearFilters()">&#10005; Limpar tudo</button>
  <span class="filter-tag" id="filter-tag">Filtro ativo</span>
</div>

<!-- TABS -->
<div class="tabs">
  <button class="tab on"  onclick="sw('resumo',this)">&#9632; Resumo executivo</button>
  <button class="tab"     onclick="sw('lanc',this)">&#128640; Lançamentos</button>
  <button class="tab"     onclick="sw('rank',this)">&#127942; Ranking inadimplência</button>
  <button class="tab"     onclick="sw('movpdd',this)">&#8645; Movimentação PDD</button>
  <button class="tab"     onclick="sw('pddpoc',this)">&#9783; PDD contábil vs POC</button>
  <button class="tab"     onclick="sw('tend',this)">&#128200; Tendência &amp; trimestre</button>
  <button class="tab-meto" onclick="openMeto()">&#9432; Metodologia</button>
</div>

<!-- ════════════ RESUMO ════════════ -->
<div id="pnl-resumo" class="pnl on">
  <div class="kg" id="kpi-grid"><!-- JS --></div>
  <div id="alertas-resumo"></div>
  <div class="sec" id="top5-sec">Top 5 — maior inadimplência</div>
  <div class="tw"><table>
    <thead><tr>
      <th>#</th><th>Empreendimento</th><th style="text-align:right">Unid. inad.</th>
      <th>Inadimplência</th><th style="text-align:right">% inad./saldo</th>
      <th style="text-align:right">Var. mês</th><th style="text-align:right">PDD</th>
      <th style="text-align:right">Var. PDD</th><th>Status</th>
    </tr></thead>
    <tbody id="resumo-top5"></tbody>
  </table></div>
  <div class="sec">Maiores movimentações de PDD</div>
  <div class="row2">
    <div class="card">
      <div class="card-hdr"><span class="ico">&#9660;</span> PDD Recuperado</div>
      <div class="tw"><table>
        <thead><tr><th>Empreendimento</th><th style="text-align:right">Variação</th><th style="text-align:right">Unidades</th></tr></thead>
        <tbody id="resumo-rec"></tbody>
      </table></div>
    </div>
    <div class="card">
      <div class="card-hdr"><span class="ico">&#9650;</span> Maiores constituições</div>
      <div class="tw"><table>
        <thead><tr><th>Empreendimento</th><th style="text-align:right">Variação</th><th style="text-align:right">Unidades</th></tr></thead>
        <tbody id="resumo-const"></tbody>
      </table></div>
    </div>
  </div>
</div>

<!-- ════════════ LANÇAMENTOS ════════════ -->
<div id="pnl-lanc" class="pnl">
  <div id="lanc-info-bar" class="ins b"></div>
  <div id="lanc-container"></div>
</div>

<!-- ════════════ RANKING ════════════ -->
<div id="pnl-rank" class="pnl">
  <div class="frow">
    <button class="fb on" onclick="fRank(this,'todos')">Todos</button>
    <button class="fb"    onclick="fRank(this,'critico')">Crítico (&gt;12%)</button>
    <button class="fb"    onclick="fRank(this,'atencao')">Atenção (8–12%)</button>
    <button class="fb"    onclick="fRank(this,'ok')">Saudável (&lt;8%)</button>
  </div>
  <div class="tw"><table>
    <thead><tr>
      <th>#</th><th>Empreendimento</th><th style="text-align:right">Unid. inad.</th>
      <th>Inadimplência</th><th style="text-align:right">% inad./saldo</th>
      <th style="text-align:right">Var. mês</th><th style="text-align:right">PDD</th>
      <th style="text-align:right">Var. PDD</th><th>Status</th>
    </tr></thead>
    <tbody id="rank-body"></tbody>
  </table></div>
</div>

<!-- ════════════ MOVIMENTAÇÃO PDD ════════════ -->
<div id="pnl-movpdd" class="pnl">
  <div class="kg" style="grid-template-columns:repeat(auto-fit,minmax(160px,1fr))" id="mov-kpis"></div>
  <div class="sec">Detalhamento por empreendimento</div>
  <div class="tw"><table>
    <thead><tr>
      <th>Empreendimento</th><th>Tipo</th><th style="text-align:right">Unidades</th>
      <th style="text-align:right">Variação PDD</th><th style="text-align:right">Var. unid.</th>
      <th>Aging dominante</th>
    </tr></thead>
    <tbody id="mov-body"></tbody>
  </table></div>
  <div class="sec">Evolução do aging — carteira inadimplente</div>
  <div class="leg" id="aging-legend"></div>
  <div class="ch" style="height:240px"><canvas id="chAging"></canvas></div>
</div>

<!-- ════════════ PDD CONTÁBIL vs POC ════════════ -->
<div id="pnl-pddpoc" class="pnl">
  <div class="kg" style="grid-template-columns:repeat(auto-fit,minmax(160px,1fr))" id="poc-kpis"></div>
  <div class="ins b">&#9432; <b>PDD contábil</b> aplica o % de provisão sobre todo o saldo devedor.
    <b>PDD sobre POC</b> considera apenas a parcela reconhecida pelo avanço físico da obra.</div>
  <div class="sec">Evolução comparada — série histórica</div>
  <div class="leg">
    <span><span class="dot" style="background:#3B6D11"></span>PDD contábil</span>
    <span><span class="dot" style="background:#97C459"></span>PDD sobre POC</span>
    <span><span class="dot" style="background:#E24B4A;width:10px;height:2px;border-radius:0;margin-top:4px"></span>% PDD/saldo (eixo dir.)</span>
  </div>
  <div class="ch" style="height:240px;margin-bottom:24px"><canvas id="chPddPoc"></canvas></div>
  <div class="sec">PDD por faixa de aging</div>
  <div class="tw"><table>
    <thead><tr>
      <th>Aging</th><th style="text-align:right">Unidades</th>
      <th style="text-align:right">Saldo inad.</th><th style="text-align:right">% provisão</th>
      <th style="text-align:right">PDD gerado</th><th style="text-align:right">Var. unid.</th>
    </tr></thead>
    <tbody id="aging-pdd-body"></tbody>
  </table></div>
  <div class="sec">Top 5 — maior PDD constituído</div>
  <div class="tw"><table>
    <thead><tr>
      <th>Empreendimento</th><th style="text-align:right">PDD contábil</th>
      <th style="text-align:right">PDD POC</th><th style="text-align:right">% PDD/saldo</th>
      <th style="text-align:right">POC acumulado</th><th style="text-align:right">Var. PDD</th>
    </tr></thead>
    <tbody id="top-pdd-body"></tbody>
  </table></div>
</div>

<!-- ════════════ TENDÊNCIA ════════════ -->
<div id="pnl-tend" class="pnl">
  <div class="sec">Taxa de inadimplência — série histórica</div>
  <div class="leg">
    <span><span class="dot" style="background:#E24B4A"></span>% inadimplência / carteira</span>
    <span><span class="dot" style="background:#185FA5;width:10px;height:2px;border-radius:0;margin-top:4px"></span>Meta (8,5%)</span>
  </div>
  <div class="ch" style="height:220px;margin-bottom:24px"><canvas id="chTend"></canvas></div>
  <div class="sec">Inadimplência &amp; PDD — série histórica</div>
  <div class="leg">
    <span><span class="dot" style="background:#E24B4A"></span>Inadimplência</span>
    <span><span class="dot" style="background:#3B6D11"></span>PDD contábil</span>
  </div>
  <div class="ch" style="height:220px;margin-bottom:24px"><canvas id="chInadPdd"></canvas></div>
  <div class="sec">Comparativo trimestral</div>
  <div class="tw"><table>
    <thead><tr id="tri-head"></tr></thead>
    <tbody id="tri-body"></tbody>
  </table></div>
</div>

<div class="footer">
  <span>CURY CONSTRUTORA &middot; PORTAL DE INADIMPLÊNCIA &mdash; CONFIDENCIAL</span>
  <span>Gerado em {datetime.now().strftime("%d/%m/%Y às %H:%M")} &middot; Competência {mes_at}</span>
</div>

<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<script>
// ══════════════════════════════════════════════
// DADOS MASTER (gerados pelo Python)
// ══════════════════════════════════════════════
const M = {master_json};

// ══════════════════════════════════════════════
// ESTADO
// ══════════════════════════════════════════════
let selReg = '', selEmp = '';
let rankStatusFilter = 'todos';
let charts = {{}};   // instâncias Chart.js reutilizáveis

// ══════════════════════════════════════════════
// HELPERS DE FORMATO
// ══════════════════════════════════════════════
function fB(v){{
  if(v==null||isNaN(v)||v===undefined)return'R$ 0,00';
  const n=v<0,a=Math.abs(v);
  if(a>=1e9)return(n?'-':'')+'R$ '+(a/1e9).toFixed(2).replace('.',',')+' bi';
  if(a>=1e6)return(n?'-':'')+'R$ '+(a/1e6).toFixed(2).replace('.',',')+' M';
  const p=a.toFixed(2).split('.');
  return(n?'-':'')+'R$ '+p[0].replace(/\B(?=(\d{{3}})+(?!\d))/g,'.')+','+p[1];
}}
function fP(v,sinal=true){{
  const s=Math.abs(v).toFixed(2).replace('.',',')+"%";
  return sinal?(v>=0?'+':'-')+s:s;
}}
function fN(v){{return Math.round(v).toLocaleString('pt-BR');}}
function dc(v,inv=false){{return v>0?(inv?'ub':'dg'):v<0?(inv?'dg':'ub'):'nm';}}
function seta(v){{return v>0?'▲':v<0?'▼':'—';}}
function niv(pi){{return pi>=12?'critico':pi>=8?'atencao':'ok';}}
function bdgNiv(nivel){{
  const m={{critico:['bdg-r','crítico'],atencao:['bdg-a','atenção'],ok:['bdg-g','saudável']}};
  const[c,l]=m[nivel]||['bdg-g',nivel];
  return`<span class="bdg ${{c}}">${{l}}</span>`;
}}

// ══════════════════════════════════════════════
// FILTRO DE DADOS
// ══════════════════════════════════════════════
function matchFilter(reg, emp){{
  return (!selReg || reg===selReg) && (!selEmp || emp===selEmp);
}}
function empRows(dt){{
  return M.emp_rows.filter(r=>r.dt===dt && matchFilter(r.reg,r.emp));
}}
function empRowsAnt(dt, dt2){{
  // retorna map emp->row para dt2
  const m={{}};
  M.emp_rows.filter(r=>r.dt===dt2&&matchFilter(r.reg,r.emp)).forEach(r=>m[r.emp]=r);
  return m;
}}
function agingRows(dt){{
  return M.aging_rows.filter(r=>r.dt===dt && matchFilter(r.reg,r.emp));
}}
function histRows(){{
  return M.hist_rows.filter(r=>matchFilter(r.reg,r.emp));
}}

// agrega array de emp_rows em totais
function agg(rows){{
  return rows.reduce((a,r)=>{{
    a.tu  +=(+r.tu  ||0); a.ui  +=(+r.ui  ||0);
    a.sd  +=(+r.sd  ||0); a.inad+=(+r.inad||0);
    a.rec +=(+r.rec ||0);
    a.pdd +=(+r.pdd ||0); a.poc +=(+r.poc ||0);
    a.pc  +=(+r.pc  ||0); a.pcc +=(+r.pcc ||0);
    return a;
  }},{{tu:0,ui:0,sd:0,inad:0,rec:0,pdd:0,poc:0,pc:0,pcc:0}});
}}

// agrega histórica por dt
function aggHist(){{
  const map={{}};
  histRows().forEach(r=>{{
    if(!map[r.dt])map[r.dt]={{inad:0,sd:0,pdd:0,poc:0,ui:0}};
    map[r.dt].inad+=(+r.inad||0); map[r.dt].sd+=(+r.sd||0);
    map[r.dt].pdd +=(+r.pdd ||0); map[r.dt].poc+=(+r.poc||0);
    map[r.dt].ui  +=(+r.ui  ||0);
  }});
  return M.hist_dts.map(dt=>map[dt]||{{inad:0,sd:0,pdd:0,poc:0,ui:0}});
}}

// ══════════════════════════════════════════════
// RENDER PRINCIPAL — chamado sempre que filtro muda
// ══════════════════════════════════════════════
function render(){{
  const at  = M.dt_at;
  const ant = M.dt_ant;
  const ret = M.dt_ret;

  const rows_at  = empRows(at);
  const rows_ant = empRows(ant);
  const rows_ret = ret ? empRows(ret) : [];

  const m_at  = agg(rows_at);
  const m_ant = agg(rows_ant);
  const m_ret = ret ? agg(rows_ret) : m_ant;

  const pct_inad_at  = m_at.sd  ? m_at.inad /m_at.sd *100 : 0;
  const pct_inad_ant = m_ant.sd ? m_ant.inad/m_ant.sd*100 : 0;
  const pct_pdd_sd   = m_at.sd  ? m_at.pdd  /m_at.sd *100 : 0;

  // ── mov data (filtrado) ──
  const map_ant = empRowsAnt(at, ant);
  const mov_data = rows_at.map(r=>{{
    const prev = map_ant[r.emp]||{{pdd:0,ui:0}};
    return {{...r, pdd_ant:prev.pdd, ui_ant:prev.ui,
             variacao:r.pdd-prev.pdd, var_ui:r.ui-prev.ui}};
  }}).filter(r=>Math.abs(r.variacao)>0.01);

  const pdd_rec   = mov_data.filter(r=>r.variacao<0).reduce((s,r)=>s+r.variacao,0);
  const pdd_const = mov_data.filter(r=>r.variacao>0).reduce((s,r)=>s+r.variacao,0);
  const pdd_liq   = m_at.pdd - m_ant.pdd;

  // ── lançamentos (filtrado) ──
  const lanc_rows = rows_at.filter(r=>r.lanc===1);

  renderKPIs(m_at, m_ant, pct_inad_at, pct_inad_ant, pct_pdd_sd, pdd_rec, lanc_rows.length);
  renderAlertas(m_at, m_ant, pct_inad_at, lanc_rows.length);
  renderTop5(rows_at, rows_ant);
  renderMovCards(mov_data);
  renderLanc(lanc_rows);
  renderRank(rows_at, rows_ant);
  renderMovKpis(pdd_rec, pdd_const, pdd_liq, M.mes_at);
  renderMovTable(mov_data);
  renderAgingChart(at, ant, ret);
  renderPocKpis(m_at, m_ant, pct_pdd_sd);
  renderPocChart();
  renderAgingPddTable(at, ant);
  renderTopPdd(rows_at, rows_ant);
  renderTend();
  renderTri(m_at, m_ant, m_ret, pct_inad_at, pct_inad_ant, pct_pdd_sd, lanc_rows.length, rows_ant, rows_ret);
}}

// ══════════════════════════════════════════════
// KPIs DO RESUMO
// ══════════════════════════════════════════════
function kpiCard(lbl, val, delta, dcls, suffix='', valCls=''){{
  return `<div class="kc">
    <div class="kc-lbl">${{lbl}}</div>
    <div class="kc-val ${{valCls}}">${{val}}</div>
    <div class="kc-d ${{dcls}}">${{delta}}${{suffix}}</div>
  </div>`;
}}
function renderKPIs(at, ant, pi_at, pi_ant, ppdd, pdd_rec, n_lanc){{
  const pi_cob_at  = (at.rec +at.inad) ? at.inad /(at.rec +at.inad)*100 : 0;
  const pi_cob_ant = (ant.rec+ant.inad) ? ant.inad/(ant.rec+ant.inad)*100 : 0;
  const inad_pct = ant.inad ? (at.inad-ant.inad)/ant.inad*100 : 0;
  const ui_pct   = ant.ui   ? (at.ui  -ant.ui  )/ant.ui  *100 : 0;
  const sd_pct   = ant.sd   ? (at.sd  -ant.sd  )/ant.sd  *100 : 0;
  const pdd_pct  = ant.pdd  ? (at.pdd -ant.pdd )/ant.pdd *100 : 0;
  const poc_pct  = ant.poc  ? (at.poc -ant.poc )/ant.poc *100 : 0;
  document.getElementById('kpi-grid').innerHTML =
    kpiCard(`Inadimplência (${{M.mes_at}})`, fB(at.inad),
      `${{seta(inad_pct)}} ${{fP(inad_pct)}} vs ${{M.mes_ant}}`, dc(inad_pct,true)) +
    kpiCard('Unidades inad.', fN(at.ui),
      `${{seta(ui_pct)}} ${{fP(ui_pct)}} vs ${{M.mes_ant}}`, dc(ui_pct,true)) +
    kpiCard('% inad. / carteira', pi_at.toFixed(2)+'%',
      `${{fP(pi_at-pi_ant)}} p.p. vs ${{M.mes_ant}}`, dc(pi_at-pi_ant,true)) +
    kpiCard('% inad. / cobrança', pi_cob_at.toFixed(2)+'%',
      `${{fP(pi_cob_at-pi_cob_ant)}} p.p. vs ${{M.mes_ant}}`, dc(pi_cob_at-pi_cob_ant,true)) +
    kpiCard('Saldo devedor', fB(at.sd),
      `${{seta(sd_pct)}} ${{fP(sd_pct)}} vs ${{M.mes_ant}}`, dc(sd_pct)) +
    kpiCard('PDD contábil', fB(at.pdd),
      `${{seta(pdd_pct)}} ${{fP(pdd_pct)}} vs ${{M.mes_ant}}`, dc(pdd_pct,true)) +
    kpiCard('PDD sobre POC', fB(at.poc),
      `${{seta(poc_pct)}} ${{fP(poc_pct)}} vs ${{M.mes_ant}}`, dc(poc_pct,true)) +
    kpiCard('Empreend. novos', n_lanc, `novos em ${{M.mes_at}}`, 'nm') +
    kpiCard('PDD Recuperado', fB(pdd_rec), 'regularizações no mês', 'dg', '', 'dg');
}}

// ══════════════════════════════════════════════
// ALERTAS
// ══════════════════════════════════════════════
function renderAlertas(at, ant, pi_at, n_lanc){{
  let h = '';
  if(at.inad < ant.inad)
    h+=`<div class="ins g">&#10003; <b>Inadimplência em queda:</b> ${{fB(Math.abs(at.inad-ant.inad))}} a menos que ${{M.mes_ant}} — redução de ${{Math.abs(at.ui-ant.ui)}} unidades.</div>`;
  else
    h+=`<div class="ins r">&#9650; <b>Inadimplência em alta:</b> ${{fB(Math.abs(at.inad-ant.inad))}} a mais que ${{M.mes_ant}} — crescimento de ${{Math.abs(at.ui-ant.ui)}} unidades.</div>`;
  if(at.pdd > ant.pdd && at.inad <= ant.inad)
    h+=`<div class="ins r">&#9650; <b>PDD sobe mesmo com inad. em queda:</b> envelhecimento da carteira eleva o percentual de provisionamento.</div>`;
  if(n_lanc>0)
    h+=`<div class="ins b">&#9432; <b>${{n_lanc}} empreendimento(s) novo(s) em ${{M.mes_at}}</b> — ainda sem impacto relevante na inadimplência.</div>`;
  document.getElementById('alertas-resumo').innerHTML = h;
}}

// ══════════════════════════════════════════════
// TOP 5 RESUMO
// ══════════════════════════════════════════════
function rankRow(r, ant_map, idx){{
  const prev = ant_map[r.emp]||{{inad:0,pdd:0}};
  const pi   = r.sd ? r.inad/r.sd*100 : 0;
  const var_i= prev.inad ? (r.inad-prev.inad)/prev.inad*100 : 0;
  const var_p= prev.pdd  ? (r.pdd -prev.pdd )/prev.pdd *100 : 0;
  const bw   = Math.min(Math.round(pi*4),100);
  const nivel= niv(pi);
  return `<tr>
    <td style="color:var(--mt);font-size:12px">${{idx}}</td>
    <td><div class="emp-nm">${{r.emp}}</div><div class="emp-sub">${{r.reg}} · ${{r.obra}}</div></td>
    <td style="text-align:right">${{fN(r.ui)}}</td>
    <td><div style="font-weight:500">${{fB(r.inad)}}</div>
        <div class="bb"><div class="bf" style="width:${{bw}}%;background:#E24B4A"></div></div></td>
    <td style="text-align:right">${{pi.toFixed(1)}}%</td>
    <td class="${{dc(var_i,true)}}" style="text-align:right">${{fP(var_i)}}</td>
    <td style="text-align:right">${{fB(r.pdd)}}</td>
    <td class="${{dc(var_p,true)}}" style="text-align:right">${{fP(var_p)}}</td>
    <td>${{bdgNiv(nivel)}}</td>
  </tr>`;
}}
function renderTop5(rows_at, rows_ant){{
  const ant_map={{}};
  rows_ant.forEach(r=>ant_map[r.emp]=r);
  const sorted = [...rows_at].sort((a,b)=>b.inad-a.inad).slice(0,5);
  document.getElementById('resumo-top5').innerHTML =
    sorted.length ? sorted.map((r,i)=>rankRow(r,ant_map,i+1)).join('') :
    '<tr><td colspan="9" class="no-data">Sem dados para o filtro selecionado</td></tr>';
}}

// ══════════════════════════════════════════════
// CARDS MOVIMENTAÇÃO (RESUMO)
// ══════════════════════════════════════════════
function renderMovCards(mov){{
  const rec   = [...mov].filter(r=>r.variacao<0).sort((a,b)=>a.variacao-b.variacao).slice(0,5);
  const const_ = [...mov].filter(r=>r.variacao>0).sort((a,b)=>b.variacao-a.variacao).slice(0,5);
  document.getElementById('resumo-rec').innerHTML = rec.length ?
    rec.map(r=>`<tr>
      <td><div class="emp-nm">${{r.emp}}</div><div class="emp-sub">${{r.reg}}</div></td>
      <td class="dg" style="text-align:right;font-weight:600">${{fB(r.variacao)}}</td>
      <td class="dg" style="text-align:right">${{fN(Math.abs(r.var_ui))}}</td></tr>`).join('') :
    '<tr><td colspan="3" class="no-data">Sem recuperações no período</td></tr>';
  document.getElementById('resumo-const').innerHTML = const_.length ?
    const_.map(r=>`<tr>
      <td><div class="emp-nm">${{r.emp}}</div><div class="emp-sub">${{r.reg}}</div></td>
      <td class="ub" style="text-align:right;font-weight:600">+${{fB(r.variacao)}}</td>
      <td class="ub" style="text-align:right">+${{fN(r.var_ui)}}</td></tr>`).join('') :
    '<tr><td colspan="3" class="no-data">Sem constituições no período</td></tr>';
}}

// ══════════════════════════════════════════════
// LANÇAMENTOS
// ══════════════════════════════════════════════
function renderLanc(lanc_rows){{
  const infoMap = {{}};
  M.lanc_info.forEach(x=>infoMap[x.emp]=x);
  document.getElementById('lanc-info-bar').innerHTML =
    `&#128640; <b>${{lanc_rows.length}} empreendimento(s) novo(s) em ${{M.mes_at}}</b> — identificados no fechamento atual.`;
  document.getElementById('lanc-container').innerHTML = lanc_rows.length ?
    lanc_rows.map(r=>{{
      const info = infoMap[r.emp]||{{}};
      const dl = (info.dl||'').substring(0,7).replace('-','/') || '—';
      const pe = info.pe||'—';
      return `<div class="lc">
        <div class="lc-head">
          <div><div class="lc-nm">${{r.emp}}</div>
          <div class="emp-sub">${{r.reg}} · Lançado ${{dl}}</div></div>
          <span class="bdg bdg-b">Novo</span>
        </div>
        <div class="lc-grid">
          <div class="lc-item"><div class="lbl">Unidades</div><div class="vl">${{fN(r.tu)}}</div></div>
          <div class="lc-item"><div class="lbl">Saldo devedor</div><div class="vl">${{fB(r.sd)}}</div></div>
          <div class="lc-item"><div class="lbl">Inadimplência</div><div class="vl dg">${{fB(r.inad)}}</div></div>
          <div class="lc-item"><div class="lbl">PDD</div><div class="vl dg">${{fB(r.pdd)}}</div></div>
          <div class="lc-item"><div class="lbl">Status obra</div><div class="vl">${{r.obra}}</div></div>
          <div class="lc-item"><div class="lbl">Entrega prevista</div><div class="vl">${{pe}}</div></div>
          <div class="lc-item"><div class="lbl">POC médio</div><div class="vl">${{(r.pocm*100).toFixed(1)}}%</div></div>
        </div>
      </div>`;
    }}).join('') :
    '<p class="no-data">Nenhum empreendimento lançado neste período para o filtro selecionado.</p>';
}}

// ══════════════════════════════════════════════
// RANKING COMPLETO
// ══════════════════════════════════════════════
function renderRank(rows_at, rows_ant){{
  const ant_map={{}};
  rows_ant.forEach(r=>ant_map[r.emp]=r);
  let sorted = [...rows_at].sort((a,b)=>b.inad-a.inad);
  if(rankStatusFilter!=='todos'){{
    sorted = sorted.filter(r=>{{
      const pi = r.sd?r.inad/r.sd*100:0;
      return niv(pi)===rankStatusFilter;
    }});
  }}
  document.getElementById('rank-body').innerHTML = sorted.length ?
    sorted.map((r,i)=>rankRow(r,ant_map,i+1)).join('') :
    '<tr><td colspan="9" class="no-data">Sem dados</td></tr>';
}}
function fRank(el, f){{
  document.querySelectorAll('.fb').forEach(b=>b.classList.remove('on'));
  el.classList.add('on');
  rankStatusFilter = f;
  renderRank(empRows(M.dt_at), empRows(M.dt_ant));
}}

// ══════════════════════════════════════════════
// MOVIMENTAÇÃO PDD
// ══════════════════════════════════════════════
function renderMovKpis(rec, const_, liq, mes){{
  document.getElementById('mov-kpis').innerHTML =
    kpiCard(`PDD Recuperado (${{mes}})`, fB(rec), 'via regularizações / distratos', 'nm', '', 'dg') +
    kpiCard(`PDD constituído (${{mes}})`, fB(const_), 'envelhecimento da carteira', 'nm', '', 'ub') +
    kpiCard('PDD líquido', fB(liq), 'variação total do período', dc(liq,true), '', dc(liq,true));
}}
function renderMovTable(mov){{
  const const_ = [...mov].filter(r=>r.variacao>0).sort((a,b)=>b.variacao-a.variacao);
  const rec    = [...mov].filter(r=>r.variacao<0).sort((a,b)=>a.variacao-b.variacao);
  const total  = mov.reduce((s,r)=>s+r.variacao,0);
  let h = '';
  [...const_.map(r=>([r,'Constituição','bdg-r'])), ...rec.map(r=>([r,'Recuperação','bdg-g']))].forEach(([r,lbl,cls])=>{{
    const su = r.var_ui>=0?'+':'';
    h+=`<tr>
      <td><div class="emp-nm">${{r.emp}}</div><div class="emp-sub">${{r.reg}}</div></td>
      <td><span class="bdg ${{cls}}">${{lbl}}</span></td>
      <td style="text-align:right">${{fN(Math.abs(r.var_ui))}}</td>
      <td class="${{dc(r.variacao,true)}}" style="text-align:right;font-weight:500">${{fB(r.variacao)}}</td>
      <td class="${{dc(r.var_ui,true)}}" style="text-align:right">${{su}}${{r.var_ui}}</td>
      <td style="text-align:right">${{r.aging}}</td>
    </tr>`;
  }});
  h+=`<tr style="background:var(--bg)"><td colspan="2"><strong>Total líquido</strong></td><td></td>
    <td class="${{dc(total,true)}}" style="text-align:right;font-weight:500">${{fB(total)}}</td>
    <td></td><td></td></tr>`;
  document.getElementById('mov-body').innerHTML = h ||
    '<tr><td colspan="6" class="no-data">Sem movimentações no período para este filtro</td></tr>';
}}

// ══════════════════════════════════════════════
// AGING CHART
// ══════════════════════════════════════════════
function renderAgingChart(at, ant, ret){{
  const meses = ret ? [M.mes_ret, M.mes_ant, M.mes_at] : [M.mes_ant, M.mes_at];
  const dts   = ret ? [ret, ant, at] : [ant, at];
  // agrupa faixas por mês
  const byDt = {{}};
  dts.forEach(dt=>{{
    byDt[dt] = {{}};
    agingRows(dt).forEach(r=>{{
      byDt[dt][r.faixa] = (byDt[dt][r.faixa]||0) + r.qtd;
    }});
  }});
  const datasets = M.aging_faixas.map((f,i)=>{{
    return {{
      label: f,
      data: dts.map(dt=>byDt[dt][f]||0),
      backgroundColor: M.aging_cores[i],
      stack: 's', borderRadius: 2,
    }};
  }});
  // legenda
  document.getElementById('aging-legend').innerHTML = M.aging_faixas.map((f,i)=>
    `<span style="display:flex;align-items:center;gap:5px">
      <span style="width:10px;height:10px;border-radius:2px;background:${{M.aging_cores[i]}};flex-shrink:0"></span>${{f}}
    </span>`).join('');
  rebuildChart('chAging',{{
    type:'bar',
    data:{{labels:meses,datasets}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{
        x:{{grid:{{color:'rgba(0,0,0,.05)'}},ticks:{{color:'#888'}}}},
        y:{{grid:{{color:'rgba(0,0,0,.05)'}},ticks:{{color:'#888'}},stacked:true,beginAtZero:true}}
      }}
    }}
  }});
}}

// ══════════════════════════════════════════════
// PDD POC KPIs + CHART
// ══════════════════════════════════════════════
function renderPocKpis(at, ant, ppdd){{
  const pdd_pct = ant.pdd ? (at.pdd-ant.pdd)/ant.pdd*100 : 0;
  const poc_pct = ant.poc ? (at.poc-ant.poc)/ant.poc*100 : 0;
  document.getElementById('poc-kpis').innerHTML =
    kpiCard(`PDD contábil (${{M.mes_at}})`, fB(at.pdd), fP(pdd_pct)+' vs '+M.mes_ant, dc(pdd_pct,true)) +
    kpiCard(`PDD sobre POC (${{M.mes_at}})`, fB(at.poc), fP(poc_pct)+' vs '+M.mes_ant, dc(poc_pct,true)) +
    kpiCard('% PDD / saldo dev.', ppdd.toFixed(2)+'%', 'base de cálculo contábil', 'nm') +
    kpiCard('Variação PDD (abs.)', fB(at.pdd-ant.pdd), 'vs '+M.mes_ant, dc(at.pdd-ant.pdd,true));
}}
function renderPocChart(){{
  const hist = aggHist();
  const pdds  = M.hist_dts.map((_,i)=>+(hist[i].pdd||0).toFixed(2));
  const pocs  = M.hist_dts.map((_,i)=>+(hist[i].poc||0).toFixed(2));
  const sdArr = M.hist_dts.map((_,i)=>+(hist[i].sd||0));
  const pctArr= M.hist_dts.map((_,i)=>sdArr[i]?+(pdds[i]/sdArr[i]*100).toFixed(2):null);
  const minY  = Math.max(0, Math.min(...pdds.filter(v=>v>0))*0.8);
  rebuildChart('chPddPoc',{{
    type:'bar',
    data:{{labels:M.hist_labels,datasets:[
      {{label:'PDD contábil',data:pdds,backgroundColor:'#3B6D11',borderRadius:4,yAxisID:'y',order:2}},
      {{label:'PDD sobre POC',data:pocs,backgroundColor:'#97C459',borderRadius:4,yAxisID:'y',order:2}},
      {{label:'% PDD/saldo',data:pctArr,type:'line',borderColor:'#E24B4A',backgroundColor:'transparent',
        pointRadius:4,borderWidth:2,yAxisID:'y2',order:1}},
    ]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{
        x:{{grid:{{color:'rgba(0,0,0,.05)'}},ticks:{{color:'#888',autoSkip:false,maxRotation:30}}}},
        y:{{grid:{{color:'rgba(0,0,0,.05)'}},ticks:{{color:'#888',callback:v=>fB(v)}},min:minY,beginAtZero:false}},
        y2:{{position:'right',grid:{{display:false}},ticks:{{color:'#888',callback:v=>v.toFixed(2)+'%'}}}}
      }}
    }}
  }});
}}

// ══════════════════════════════════════════════
// AGING PDD TABLE
// ══════════════════════════════════════════════
function renderAgingPddTable(at, ant){{
  const at_rows  = agingRows(at);
  const ant_rows = agingRows(ant);
  // agrega por faixa
  const byFaixaAt={{}}, byFaixaAnt={{}};
  at_rows.forEach(r=>{{
    if(!byFaixaAt[r.faixa])byFaixaAt[r.faixa]={{qtd:0,si:0,pdd:0}};
    byFaixaAt[r.faixa].qtd+=r.qtd; byFaixaAt[r.faixa].si+=r.si; byFaixaAt[r.faixa].pdd+=r.pdd;
  }});
  ant_rows.forEach(r=>{{
    if(!byFaixaAnt[r.faixa])byFaixaAnt[r.faixa]={{qtd:0}};
    byFaixaAnt[r.faixa].qtd+=r.qtd;
  }});
  let h = M.aging_faixas.map(f=>{{
    const a = byFaixaAt[f]; if(!a||!a.qtd) return '';
    const b = byFaixaAnt[f]||{{qtd:0}};
    const pct_prov = a.si ? a.pdd/a.si*100 : 0;
    const var_u = a.qtd - b.qtd;
    return `<tr>
      <td>${{f}}</td>
      <td style="text-align:right">${{fN(a.qtd)}}</td>
      <td style="text-align:right">${{fB(a.si)}}</td>
      <td style="text-align:right">${{pct_prov.toFixed(1)}}%</td>
      <td style="text-align:right;font-weight:500">${{fB(a.pdd)}}</td>
      <td class="${{dc(var_u,true)}}" style="text-align:right">${{var_u>=0?'+':''}}${{var_u}}</td>
    </tr>`;
  }}).join('');
  document.getElementById('aging-pdd-body').innerHTML = h ||
    '<tr><td colspan="6" class="no-data">Sem dados</td></tr>';
}}

// ══════════════════════════════════════════════
// TOP 5 PDD
// ══════════════════════════════════════════════
function renderTopPdd(rows_at, rows_ant){{
  const ant_map={{}};
  rows_ant.forEach(r=>ant_map[r.emp]=r);
  const sorted = [...rows_at].sort((a,b)=>b.pc-a.pc).slice(0,5);
  document.getElementById('top-pdd-body').innerHTML = sorted.length ?
    sorted.map(r=>{{
      const prev = ant_map[r.emp]||{{pc:0}};
      const pct_s = r.sd ? r.pc/r.sd*100 : 0;
      const var_p = r.pc - prev.pc;
      return `<tr>
        <td><div class="emp-nm">${{r.emp}}</div></td>
        <td style="text-align:right">${{fB(r.pc)}}</td>
        <td style="text-align:right">${{fB(r.pcc)}}</td>
        <td style="text-align:right">${{pct_s.toFixed(2)}}%</td>
        <td style="text-align:right">${{(r.pocm*100).toFixed(1)}}%</td>
        <td class="${{dc(var_p,true)}}" style="text-align:right">${{fB(var_p)}}</td>
      </tr>`;
    }}).join('') :
    '<tr><td colspan="6" class="no-data">Sem dados</td></tr>';
}}

// ══════════════════════════════════════════════
// TENDÊNCIA
// ══════════════════════════════════════════════
function renderTend(){{
  const hist = aggHist();
  const inad  = M.hist_dts.map((_,i)=>+(hist[i].inad||0).toFixed(2));
  const pdds  = M.hist_dts.map((_,i)=>+(hist[i].pdd||0).toFixed(2));
  const sdArr = M.hist_dts.map((_,i)=>+(hist[i].sd||0));
  const pctI  = M.hist_dts.map((_,i)=>sdArr[i]?+(inad[i]/sdArr[i]*100).toFixed(2):0);
  rebuildChart('chTend',{{
    type:'line',
    data:{{labels:M.hist_labels,datasets:[
      {{label:'% inad.',data:pctI,borderColor:'#E24B4A',backgroundColor:'rgba(226,75,74,.08)',
        tension:.4,pointRadius:4,fill:true,borderWidth:2}},
      {{label:'Meta 8,5%',data:M.hist_labels.map(()=>8.5),borderColor:'#185FA5',
        borderDash:[6,4],pointRadius:0,borderWidth:1.5,fill:false}},
    ]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{
        x:{{grid:{{color:'rgba(0,0,0,.05)'}},ticks:{{color:'#888',autoSkip:false,maxRotation:30}}}},
        y:{{grid:{{color:'rgba(0,0,0,.05)'}},ticks:{{color:'#888',callback:v=>v.toFixed(1)+'%'}}}}
      }}
    }}
  }});
  rebuildChart('chInadPdd',{{
    type:'line',
    data:{{labels:M.hist_labels,datasets:[
      {{label:'Inadimplência',data:inad,borderColor:'#E24B4A',backgroundColor:'rgba(226,75,74,.08)',
        tension:.4,pointRadius:4,fill:true,borderWidth:2}},
      {{label:'PDD contábil',data:pdds,borderColor:'#3B6D11',backgroundColor:'rgba(59,109,17,.07)',
        tension:.4,pointRadius:4,fill:true,borderWidth:2}},
    ]}},
    options:{{responsive:true,maintainAspectRatio:false,
      plugins:{{legend:{{display:false}}}},
      scales:{{
        x:{{grid:{{color:'rgba(0,0,0,.05)'}},ticks:{{color:'#888',autoSkip:false,maxRotation:30}}}},
        y:{{grid:{{color:'rgba(0,0,0,.05)'}},ticks:{{color:'#888',callback:v=>fB(v)}}}}
      }}
    }}
  }});
}}

// ══════════════════════════════════════════════
// COMPARATIVO TRIMESTRAL
// ══════════════════════════════════════════════
function renderTri(at, ant, ret, pi_at, pi_ant, ppdd, n_lanc, rows_ant, rows_ret){{
  const pi_ret = ret&&ret.sd ? ret.inad/ret.sd*100 : 0;
  const ppdd_ant = ant.sd ? ant.pdd/ant.sd*100 : 0;
  const hasRet = !!M.dt_ret;
  // cabeçalho
  document.getElementById('tri-head').innerHTML =
    '<th>Indicador</th>' +
    (hasRet?`<th>${{M.mes_ret}}</th>`:'') +
    `<th>${{M.mes_ant}}</th><th>${{M.mes_at}}</th><th>Var. período</th>`;
  const base_inad = hasRet ? ret.inad : ant.inad;
  const base_sd   = hasRet ? ret.sd   : ant.sd;
  const base_pdd  = hasRet ? ret.pdd  : ant.pdd;
  const base_poc  = hasRet ? ret.poc  : ant.poc;
  const base_ui   = hasRet ? ret.ui   : ant.ui;
  const n_ant = rows_ant.filter(r=>r.lanc===1).length;
  const n_ret = rows_ret.filter(r=>r.lanc===1).length;
  function pctVar(a,b){{return b?(a-b)/b*100:0;}}
  function row(lbl, vr, va, vat, fmt, inv=false, suffix=''){{
    const vp = pctVar(vat,hasRet?vr:va);
    return `<tr>
      <td>${{lbl}}</td>
      ${{hasRet?`<td>${{fmt(vr)}}${{suffix}}</td>`:''}}
      <td>${{fmt(va)}}${{suffix}}</td>
      <td>${{fmt(vat)}}${{suffix}}</td>
      <td class="${{dc(vp,inv)}}">${{fP(vp)}}</td>
    </tr>`;
  }}
  document.getElementById('tri-body').innerHTML =
    row('<b>Saldo devedor total</b>', base_sd, ant.sd, at.sd, fB) +
    row('Inadimplência carteira', base_inad, ant.inad, at.inad, fB, true) +
    row('Unidades inadimplentes', base_ui, ant.ui, at.ui, fN, true) +
    `<tr><td>% inad. / carteira</td>
      ${{hasRet?`<td>${{pi_ret.toFixed(2)}}%</td>`:''}}
      <td>${{pi_ant.toFixed(2)}}%</td><td>${{pi_at.toFixed(2)}}%</td>
      <td class="${{dc(pi_at-pi_ant,true)}}">${{fP(pi_at-pi_ant)}} p.p.</td></tr>` +
    row('PDD contábil', base_pdd, ant.pdd, at.pdd, fB, true) +
    row('PDD sobre POC', base_poc, ant.poc, at.poc, fB, true) +
    `<tr><td>% PDD / saldo dev.</td>
      ${{hasRet?`<td>${{pi_ret.toFixed(2)}}%</td>`:''}}
      <td>${{ppdd_ant.toFixed(2)}}%</td><td>${{ppdd.toFixed(2)}}%</td>
      <td class="nm">—</td></tr>` +
    `<tr><td>Empreend. novos (mês)</td>
      ${{hasRet?`<td>${{n_ret}}</td>`:''}}
      <td>${{n_ant}}</td><td>${{n_lanc}}</td>
      <td class="nm">lançamentos por mês</td></tr>`;
}}

// ══════════════════════════════════════════════
// CHART HELPER — reconstrói sem duplicar
// ══════════════════════════════════════════════
function rebuildChart(id, cfg){{
  if(charts[id]){{ charts[id].destroy(); delete charts[id]; }}
  const canvas = document.getElementById(id);
  if(!canvas) return;
  charts[id] = new Chart(canvas, cfg);
}}

// ══════════════════════════════════════════════
// TABS
// ══════════════════════════════════════════════
function sw(id, el){{
  document.querySelectorAll('.pnl').forEach(p=>p.classList.remove('on'));
  document.querySelectorAll('.tab').forEach(t=>t.classList.remove('on'));
  document.getElementById('pnl-'+id).classList.add('on');
  el.classList.add('on');
}}

// ══════════════════════════════════════════════
// MODAL
// ══════════════════════════════════════════════
function openMeto(){{document.getElementById('modal-meto').classList.add('on');}}
function closeMeto(){{document.getElementById('modal-meto').classList.remove('on');}}
document.addEventListener('keydown',e=>{{if(e.key==='Escape')closeMeto();}});

// ══════════════════════════════════════════════
// AUTOCOMPLETE DROPDOWNS
// ══════════════════════════════════════════════
const _allRegs = M.regionais;
const _allEmps = [...new Set(M.emp_rows.filter(r=>r.dt===M.dt_at).map(r=>r.emp))].sort();

function fiRender(type, query){{
  const drop = document.getElementById('drop-'+type);
  const pool = type==='reg' ? _allRegs
             : (selReg ? (M.emps_por_regional[selReg]||_allEmps) : _allEmps);
  const q = query.toLowerCase().trim();
  const matches = q ? pool.filter(x=>x.toLowerCase().includes(q)) : pool;
  drop.innerHTML = '';
  if(!matches.length){{
    drop.innerHTML='<li class="no-res">Nenhum resultado</li>';
  }} else {{
    matches.slice(0,80).forEach(x=>{{
      const li = document.createElement('li');
      if(q){{
        const idx = x.toLowerCase().indexOf(q);
        li.innerHTML = x.slice(0,idx)+'<b>'+x.slice(idx,idx+q.length)+'</b>'+x.slice(idx+q.length);
      }} else {{ li.textContent=x; }}
      li.dataset.val = x;
      li.addEventListener('mousedown',e=>{{e.preventDefault();fiSelect(type,x);}});
      drop.appendChild(li);
    }});
  }}
}}
function fiOpen(type){{
  fiRender(type, document.getElementById('inp-'+(type==='reg'?'regional':'emp')).value);
  document.getElementById('drop-'+type).classList.add('on');
}}
function fiInput(type){{
  const inp = document.getElementById('inp-'+(type==='reg'?'regional':'emp'));
  const q = inp.value;
  document.getElementById('clr-'+type).classList.toggle('on',q.length>0);
  fiRender(type,q);
  document.getElementById('drop-'+type).classList.add('on');
  if(!q){{ if(type==='reg'){{selReg='';}}else{{selEmp='';}} render(); }}
}}
function fiSelect(type, val){{
  if(type==='reg'){{
    selReg=val;
    document.getElementById('inp-regional').value=val;
    document.getElementById('clr-reg').classList.add('on');
    if(selEmp && !(M.emps_por_regional[val]||[]).includes(selEmp)){{
      selEmp='';
      document.getElementById('inp-emp').value='';
      document.getElementById('clr-emp').classList.remove('on');
    }}
  }} else {{
    selEmp=val;
    document.getElementById('inp-emp').value=val;
    document.getElementById('clr-emp').classList.add('on');
  }}
  document.getElementById('drop-'+type).classList.remove('on');
  updateFilterTag();
  render();
}}
function fiClear(type){{
  if(type==='reg'){{selReg='';document.getElementById('inp-regional').value='';document.getElementById('clr-reg').classList.remove('on');}}
  else{{selEmp='';document.getElementById('inp-emp').value='';document.getElementById('clr-emp').classList.remove('on');}}
  document.getElementById('drop-'+type).classList.remove('on');
  updateFilterTag();
  render();
}}
function fiKey(e,type){{
  const drop=document.getElementById('drop-'+type);
  const items=[...drop.querySelectorAll('li:not(.no-res)')];
  let hi=items.findIndex(l=>l.classList.contains('hi'));
  if(e.key==='ArrowDown'){{e.preventDefault();if(hi<items.length-1){{if(hi>=0)items[hi].classList.remove('hi');items[hi+1].classList.add('hi');items[hi+1].scrollIntoView({{block:'nearest'}});}}}}
  else if(e.key==='ArrowUp'){{e.preventDefault();if(hi>0){{items[hi].classList.remove('hi');items[hi-1].classList.add('hi');items[hi-1].scrollIntoView({{block:'nearest'}});}}}}
  else if(e.key==='Enter'){{if(hi>=0)fiSelect(type,items[hi].dataset.val);}}
  else if(e.key==='Escape'){{drop.classList.remove('on');}}
}}
document.addEventListener('click',e=>{{
  if(!e.target.closest('#fi-reg'))document.getElementById('drop-reg').classList.remove('on');
  if(!e.target.closest('#fi-emp'))document.getElementById('drop-emp').classList.remove('on');
}});
function updateFilterTag(){{
  const tag = document.getElementById('filter-tag');
  const label = [selReg,selEmp].filter(Boolean).join(' › ');
  if(label){{tag.classList.add('on');tag.textContent=label;}}
  else{{tag.classList.remove('on');}}
}}
function clearFilters(){{fiClear('reg');fiClear('emp');}}

// ══════════════════════════════════════════════
// BOOT
// ══════════════════════════════════════════════
render();
</script>
</body>
</html>"""

# =====================================================
# SALVAR
# =====================================================
Path(PASTA_SAIDA).mkdir(parents=True, exist_ok=True)
arquivo = Path(PASTA_SAIDA) / f"portal_inadimplencia_{dt_at.strftime('%Y_%m')}.html"
with open(arquivo, "w", encoding="utf-8") as f:
    f.write(html)

print()
print("=" * 60)
print("  PORTAL GERADO COM SUCESSO!")
print(f"  Arquivo : portal_inadimplencia_{dt_at.strftime('%Y_%m')}.html")
print(f"  Local   : {PASTA_SAIDA}")
print(f"  Período : {mes_ret} → {mes_ant} → {mes_at}")
print("=" * 60)