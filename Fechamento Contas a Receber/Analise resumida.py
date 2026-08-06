import warnings
warnings.filterwarnings("ignore")

import urllib
import pandas as pd
from sqlalchemy import create_engine, text
from datetime import datetime
from pathlib import Path
import base64

# =====================================================
# CONFIGURAÇÕES
# =====================================================

PASTA_SAIDA = r"C:\Users\srvautbicar\Desktop\PROCESSOS\import para SQL\Base PDD\Emails"

LOGO_PATH = r"C:\Users\srvautbicar\Desktop\PROCESSOS\import para SQL\Base PDD\logo_cury.png"

# =====================================================
# PARÂMETRO
# =====================================================

print("=" * 55)
print("  EMAIL FECHAMENTO MENSAL - PDD")
print("=" * 55)
print()

def pedir_data(label):

    while True:

        entrada = input(f"📅 {label}: ").strip()

        try:
            return datetime.strptime(
                entrada,
                "%Y-%m-%d"
            ).date()

        except ValueError:
            print("❌ Formato inválido. Use AAAA-MM-DD\n")

dt_atual = pedir_data(
    "Mês de fechamento (ex: 2026-04-30)"
)

# =====================================================
# CONEXÃO
# =====================================================

params = urllib.parse.quote_plus(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=SRJDLK01\\SSPDLKAZ01;"
    "DATABASE=P_OURO_CAR_DB;"
    "UID=powerbi;"
    "PWD=Ah79Wlk999;"
)

engine = create_engine(
    f"mssql+pyodbc:///?odbc_connect={params}",
    fast_executemany=True
)

print("\n🔌 Conectando ao banco...")

# =====================================================
# LOGO
# =====================================================

try:

    with open(LOGO_PATH, "rb") as f:

        logo_src = (
            "data:image/png;base64,"
            + base64.b64encode(f.read()).decode()
        )

except:
    logo_src = ""

# =====================================================
# BUSCA ÚLTIMOS 3 MESES
# =====================================================

with engine.connect() as conn:

    r = conn.execute(text("""

        SELECT DISTINCT TOP 3 dt_fechamento

        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR

        WHERE dt_fechamento <= :dt

        ORDER BY dt_fechamento DESC

    """), {"dt": str(dt_atual)})

    datas = [row[0] for row in r.fetchall()]

if len(datas) < 3:

    print(
        f"\n❌ Necessário pelo menos "
        f"3 meses no banco. "
        f"Encontrado: {len(datas)}"
    )

    exit()

dt_at  = datas[0]
dt_ant = datas[1]
dt_ret = datas[2]

MESES_PT = {
    1:"janeiro",
    2:"fevereiro",
    3:"março",
    4:"abril",
    5:"maio",
    6:"junho",
    7:"julho",
    8:"agosto",
    9:"setembro",
    10:"outubro",
    11:"novembro",
    12:"dezembro"
}

mes_at  = MESES_PT[dt_at.month]
mes_ant = MESES_PT[dt_ant.month]
mes_ret = MESES_PT[dt_ret.month]

print(
    f"✅ Meses identificados: "
    f"{mes_ret} → {mes_ant} → {mes_at}"
)

# =====================================================
# QUERY PRINCIPAL
# =====================================================

def metricas(conn, dt):

    r = conn.execute(text("""

        SELECT

            COUNT(*) AS total_unidades,

            COUNT(
                CASE
                    WHEN status_da_unidade = 'Inadimplente'
                    THEN 1
                END
            ) AS unid_inadimplentes,

            SUM(
                CAST(saldo_dev_carteira AS FLOAT)
            ) AS saldo_devedor,

            SUM(
                CAST(inadimplencia_carteira AS FLOAT)
            ) AS inadimplencia,

            SUM(
                CAST(valor_pdd AS FLOAT)
            ) AS pdd,

            SUM(
                CAST(valor_pdd_poc AS FLOAT)
            ) AS pdd_poc

        FROM FAT.TB_TRN_CAR_CARTEIRAORI_OUR

        WHERE dt_fechamento = :dt

    """), {"dt": str(dt)})

    row = r.fetchone()

    return {

        "total_unidades":
            row[0],

        "unid_inadimplentes":
            row[1],

        "saldo_devedor":
            row[2] or 0,

        "inadimplencia":
            row[3] or 0,

        "pdd":
            row[4] or 0,

        "pdd_poc":
            row[5] or 0,
    }

with engine.connect() as conn:

    m_at  = metricas(conn, dt_at)
    m_ant = metricas(conn, dt_ant)
    m_ret = metricas(conn, dt_ret)

print("✅ Dados carregados!")

# =====================================================
# FUNÇÕES
# =====================================================

def pct(atual, anterior):

    if anterior == 0:
        return 0

    return (
        (atual - anterior)
        / anterior
        * 100
    )

def fmt_brl(v, decimais=2):

    neg = v < 0

    if abs(v) >= 1_000_000_000:

        s = (
            f"R$ {abs(v)/1_000_000_000:.2f} bilhões"
            .replace(".", ",")
        )

    elif abs(v) >= 1_000_000:

        s = (
            f"R$ {abs(v)/1_000_000:.2f} milhões"
            .replace(".", ",")
        )

    else:

        p = (
            f"{abs(v):,.{decimais}f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

        s = f"R$ {p}"

    return f"-{s}" if neg else s

def fmt_n(v):

    return (
        f"{int(v):,}"
        .replace(",", ".")
    )

def fmt_pct(v):

    s = (
        f"{abs(v):.2f}%"
        .replace(".", ",")
    )

    return f"+{s}" if v >= 0 else f"-{s}"

# =====================================================
# CÁLCULOS
# =====================================================

# UNIDADES

unid_at  = m_at["total_unidades"]
unid_ant = m_ant["total_unidades"]

unid_dif = unid_at - unid_ant

unid_pct = pct(
    unid_at,
    unid_ant
)

# SALDO DEVEDOR

sd_at  = m_at["saldo_devedor"]
sd_ant = m_ant["saldo_devedor"]
sd_ret = m_ret["saldo_devedor"]

sd_pct_at = pct(
    sd_at,
    sd_ant
)

sd_pct_ant = pct(
    sd_ant,
    sd_ret
)

# INADIMPLÊNCIA

inad_at  = m_at["inadimplencia"]
inad_ant = m_ant["inadimplencia"]
inad_ret = m_ret["inadimplencia"]

inad_pct_at = pct(
    inad_at,
    inad_ant
)

inad_pct_ant = pct(
    inad_ant,
    inad_ret
)

# PDD

pdd_at  = m_at["pdd"]
pdd_ant = m_ant["pdd"]
pdd_ret = m_ret["pdd"]

pdd_pct_at = pct(
    pdd_at,
    pdd_ant
)

pdd_pct_ant = pct(
    pdd_ant,
    pdd_ret
)

# PDD POC

poc_at  = m_at["pdd_poc"]
poc_ant = m_ant["pdd_poc"]

poc_pct = pct(
    poc_at,
    poc_ant
)

# INADIMPLENTES

unid_inad_at  = m_at["unid_inadimplentes"]
unid_inad_ant = m_ant["unid_inadimplentes"]

pct_inad_at = (
    unid_inad_at
    / unid_at
    * 100
)

pct_inad_ant = (
    unid_inad_ant
    / unid_ant
    * 100
)

# =====================================================
# TEXTOS EXPLICATIVOS
# =====================================================

# UNIDADES

if unid_dif > 0:

    txt_unidades = f"""
    No período, houve variação na quantidade de unidades,
    passando de {fmt_n(unid_ant)} para {fmt_n(unid_at)},
    em função da realização de novos lançamentos.

    Dessa forma, registramos uma evolução de
    {fmt_pct(unid_pct)} no comparativo mensal.
    """

else:

    txt_unidades = f"""
    No período, a quantidade de unidades permaneceu estável,
    totalizando {fmt_n(unid_at)} unidades no fechamento atual.
    """

# SALDO DEVEDOR

sd_tendencia = (
    "positiva"
    if sd_pct_at > sd_pct_ant
    else "negativa"
)

txt_saldo = f"""
Em relação ao saldo devedor, registramos uma variação de
{fmt_pct(sd_pct_at)} em {mes_at},
na comparação com {mes_ant}.

No período anterior, a variação registrada havia sido de
{fmt_pct(sd_pct_ant)},
representando uma diferença de
{abs(sd_pct_at - sd_pct_ant):.2f}
pontos percentuais entre os períodos,
indicando tendência {sd_tendencia}.
"""

# INADIMPLÊNCIA

resultado_inad = (
    "positivo"
    if inad_pct_at < inad_pct_ant
    else "negativo"
)

txt_inad = f"""
A inadimplência apresentou uma variação de
{fmt_pct(inad_pct_at)} em {mes_at},
resultado {resultado_inad}
quando comparado ao período anterior,
que havia registrado {fmt_pct(inad_pct_ant)}.

Em termos absolutos,
o montante inadimplente
{("recuou" if inad_at < inad_ant else "avançou")}
de {fmt_brl(inad_ant)}
para {fmt_brl(inad_at)}.

O percentual de unidades inadimplentes
sobre o total
{("reduziu" if pct_inad_at < pct_inad_ant else "aumentou")}
de {fmt_pct(pct_inad_ant)}
para {fmt_pct(pct_inad_at)},
representando uma variação de
{abs(pct_inad_at - pct_inad_ant):.2f}
pontos percentuais.
"""

# PDD

relacao_pdd = (
    "inversamente relacionado"
    if (
        (pdd_pct_at > 0 and inad_pct_at < 0)
        or
        (pdd_pct_at < 0 and inad_pct_at > 0)
    )
    else "alinhado"
)

txt_pdd = f"""
Em relação ao PDD,
registramos uma variação de
{fmt_pct(pdd_pct_at)} no período,
equivalente a
{fmt_brl(abs(pdd_at - pdd_ant))}
em comparação ao fechamento anterior.

O resultado apresentado está
{relacao_pdd}
ao comportamento da inadimplência
no período.

No fechamento anterior,
a variação registrada havia sido de
{fmt_pct(pdd_pct_ant)}.

Sobre o PDD vinculado ao POC,
foi registrada uma variação de
{fmt_pct(poc_pct)},
equivalente a
{fmt_brl(abs(poc_at - poc_ant))}
frente ao mês anterior.
"""

# COMPARATIVO 3 MESES

txt_3m = f"""
O comparativo trimestral demonstra
a evolução consolidada
dos principais indicadores da carteira
nos últimos três fechamentos.

No período analisado,
o saldo devedor apresentou movimentação
de {fmt_brl(sd_ret)} em {mes_ret},
para {fmt_brl(sd_at)} em {mes_at}.

A inadimplência apresentou comportamento
{("decrescente" if inad_at < inad_ret else "crescente")}
ao longo do trimestre,
encerrando o período em
{fmt_brl(inad_at)}.

Já o PDD acompanhou
a movimentação da carteira,
passando de
{fmt_brl(pdd_ret)}
para
{fmt_brl(pdd_at)}
no fechamento atual.
"""

# =====================================================
# HTML
# =====================================================

html = f"""
<!DOCTYPE html>

<html lang="pt-BR">

<head>

<meta charset="UTF-8">

<style>

body{{
    font-family:Calibri,Arial,sans-serif;
    font-size:14px;
    color:#222;
    background:#f4f4f4;
    margin:0;
    padding:20px
}}

.container{{
    max-width:900px;
    margin:0 auto;
    background:#fff;
    border-radius:8px;
    overflow:hidden;
    box-shadow:0 2px 8px rgba(0,0,0,0.1)
}}

.header{{
    background:#1B3A8C;
    color:#fff;
    padding:28px 32px
}}

.header h1{{
    font-size:24px;
    margin:0 0 4px
}}

.header p{{
    margin:0;
    font-size:13px;
    opacity:.7
}}

.logo{{
    height:32px;
    margin-bottom:14px
}}

.body{{
    padding:28px 32px
}}

.bloco{{
    margin-bottom:28px;
    padding-bottom:28px;
    border-bottom:1px solid #eee
}}

.bloco:last-child{{
    border-bottom:none
}}

.bloco-titulo{{
    font-size:13px;
    font-weight:700;
    color:#1B3A8C;
    text-transform:uppercase;
    letter-spacing:1px;
    margin-bottom:12px
}}

.bloco p{{
    margin:0 0 8px;
    line-height:1.7;
    color:#333
}}

.kpis{{
    display:grid;
    grid-template-columns:repeat(3,1fr);
    gap:12px;
    margin:16px 0
}}

.kpi{{
    background:#f5f7ff;
    border:1px solid #dde2ef;
    border-radius:8px;
    padding:12px 14px
}}

.kpi-label{{
    font-size:10px;
    font-weight:700;
    color:#6b7280;
    text-transform:uppercase;
    letter-spacing:1px;
    margin-bottom:4px
}}

.kpi-val{{
    font-size:18px;
    font-weight:700;
    color:#1B3A8C
}}

.kpi-sub{{
    font-size:11px;
    color:#888;
    margin-top:2px
}}

.up{{
    color:#c0392b;
    font-weight:700
}}

.dn{{
    color:#1a7a4a;
    font-weight:700
}}

table{{
    width:100%;
    border-collapse:collapse;
    margin-top:12px
}}

thead tr{{
    background:#f5f7ff
}}

th{{
    padding:10px;
    border:1px solid #dde2ef;
    text-align:center;
    font-size:12px
}}

td{{
    padding:10px;
    border:1px solid #eee;
    text-align:center
}}

.footer{{
    background:#0f2460;
    color:rgba(255,255,255,.5);
    padding:16px 32px;
    font-size:11px;
    text-align:center
}}

</style>

</head>

<body>

<div class="container">

    <div class="header">

        <img src="{logo_src}" class="logo">

        <h1>
            Fechamento Mensal — {mes_at.capitalize()}/{dt_at.year}
        </h1>

        <p>
            Provisão para Devedores Duvidosos ·
            Gerado em
            {datetime.now().strftime("%d/%m/%Y às %H:%M")}
        </p>

    </div>

    <div class="body">

        <p style="
            margin:0 0 24px;
            color:#555;
            line-height:1.7
        ">

            Relatório gerencial referente
            ao fechamento mensal da carteira,
            contendo os principais indicadores
            relacionados ao saldo devedor,
            inadimplência e provisão
            para devedores duvidosos (PDD),
            com comparativo entre os períodos analisados
            e acompanhamento da evolução
            dos indicadores da carteira.

        </p>

        <!-- UNIDADES -->

        <div class="bloco">

            <div class="bloco-titulo">
                📦 Unidades
            </div>

            <div class="kpis">

                <div class="kpi">

                    <div class="kpi-label">
                        {mes_ant.capitalize()}
                    </div>

                    <div class="kpi-val">
                        {fmt_n(unid_ant)}
                    </div>

                    <div class="kpi-sub">
                        unidades
                    </div>

                </div>

                <div class="kpi">

                    <div class="kpi-label">
                        {mes_at.capitalize()}
                    </div>

                    <div class="kpi-val">
                        {fmt_n(unid_at)}
                    </div>

                    <div class="kpi-sub">
                        unidades
                    </div>

                </div>

                <div class="kpi">

                    <div class="kpi-label">
                        Variação
                    </div>

                    <div class="kpi-val {'up' if unid_pct >= 0 else 'dn'}">
                        {fmt_pct(unid_pct)}
                    </div>

                    <div class="kpi-sub">
                        {fmt_n(abs(unid_dif))}
                        unidades
                    </div>

                </div>

            </div>

            <p>{txt_unidades}</p>

        </div>

        <!-- SALDO -->

        <div class="bloco">

            <div class="bloco-titulo">
                💰 Saldo Devedor
            </div>

            <div class="kpis">

                <div class="kpi">

                    <div class="kpi-label">
                        {mes_ant.capitalize()}
                    </div>

                    <div class="kpi-val">
                        {fmt_brl(sd_ant)}
                    </div>

                </div>

                <div class="kpi">

                    <div class="kpi-label">
                        {mes_at.capitalize()}
                    </div>

                    <div class="kpi-val">
                        {fmt_brl(sd_at)}
                    </div>

                </div>

                <div class="kpi">

                    <div class="kpi-label">
                        Variação
                    </div>

                    <div class="kpi-val {'up' if sd_pct_at >= 0 else 'dn'}">
                        {fmt_pct(sd_pct_at)}
                    </div>

                    <div class="kpi-sub">
                        anterior:
                        {fmt_pct(sd_pct_ant)}
                    </div>

                </div>

            </div>

            <p>{txt_saldo}</p>

        </div>

        <!-- INADIMPLÊNCIA -->

        <div class="bloco">

            <div class="bloco-titulo">
                ⚠️ Inadimplência
            </div>

            <div class="kpis">

                <div class="kpi">

                    <div class="kpi-label">
                        {mes_ant.capitalize()}
                    </div>

                    <div class="kpi-val">
                        {fmt_brl(inad_ant)}
                    </div>

                    <div class="kpi-sub">
                        {fmt_n(unid_inad_ant)}
                        unidades ·
                        {fmt_pct(pct_inad_ant)}
                    </div>

                </div>

                <div class="kpi">

                    <div class="kpi-label">
                        {mes_at.capitalize()}
                    </div>

                    <div class="kpi-val">
                        {fmt_brl(inad_at)}
                    </div>

                    <div class="kpi-sub">
                        {fmt_n(unid_inad_at)}
                        unidades ·
                        {fmt_pct(pct_inad_at)}
                    </div>

                </div>

                <div class="kpi">

                    <div class="kpi-label">
                        Variação
                    </div>

                    <div class="kpi-val {'dn' if inad_pct_at < 0 else 'up'}">
                        {fmt_pct(inad_pct_at)}
                    </div>

                    <div class="kpi-sub">
                        anterior:
                        {fmt_pct(inad_pct_ant)}
                    </div>

                </div>

            </div>

            <p>{txt_inad}</p>

        </div>

        <!-- PDD -->

        <div class="bloco">

            <div class="bloco-titulo">
                📊 PDD — Provisão para Devedores Duvidosos
            </div>

            <div class="kpis">

                <div class="kpi">

                    <div class="kpi-label">
                        {mes_ant.capitalize()}
                    </div>

                    <div class="kpi-val">
                        {fmt_brl(pdd_ant)}
                    </div>

                </div>

                <div class="kpi">

                    <div class="kpi-label">
                        {mes_at.capitalize()}
                    </div>

                    <div class="kpi-val">
                        {fmt_brl(pdd_at)}
                    </div>

                </div>

                <div class="kpi">

                    <div class="kpi-label">
                        Variação
                    </div>

                    <div class="kpi-val {'up' if pdd_pct_at >= 0 else 'dn'}">
                        {fmt_pct(pdd_pct_at)}
                    </div>

                    <div class="kpi-sub">
                        anterior:
                        {fmt_pct(pdd_pct_ant)}
                    </div>

                </div>

            </div>

<div class="kpis">

    <div
        class="kpi"
        style="grid-column:span 2"
    >

        <div class="kpi-label">
            PDD sobre POC
        </div>

        <div class="kpi-val">
            {fmt_brl(poc_at)}
        </div>

        <div class="kpi-sub">
            {fmt_pct(poc_pct)}
            vs {mes_ant} ·
            variação de
            {fmt_brl(abs(poc_at - poc_ant))}
        </div>

    </div>

    <div class="kpi">

        <div class="kpi-label">
            Variação PDD POC
        </div>

        <div class="kpi-val {'up' if poc_pct >= 0 else 'dn'}">
            {fmt_pct(poc_pct)}
        </div>

        <div class="kpi-sub">
            {fmt_brl(abs(poc_at - poc_ant))}
        </div>

    </div>

</div>

            <p>{txt_pdd}</p>

        </div>

        <!-- COMPARATIVO 3 MESES -->

        <div class="bloco">

            <div class="bloco-titulo">
                📈 Comparativo Últimos 3 Meses
            </div>

            <table>

                <thead>

                    <tr>

                        <th>Métrica</th>

                        <th>
                            {mes_ret.capitalize()}
                        </th>

                        <th>
                            {mes_ant.capitalize()}
                        </th>

                        <th>
                            {mes_at.capitalize()}
                        </th>

                    </tr>

                </thead>

                <tbody>

                    <tr>

                        <td>
                            Saldo Devedor
                        </td>

                        <td>
                            {fmt_brl(sd_ret)}
                        </td>

                        <td>
                            {fmt_brl(sd_ant)}
                        </td>

                        <td>
                            {fmt_brl(sd_at)}
                        </td>

                    </tr>

                    <tr>

                        <td>
                            Inadimplência
                        </td>

                        <td>
                            {fmt_brl(inad_ret)}
                        </td>

                        <td>
                            {fmt_brl(inad_ant)}
                        </td>

                        <td>
                            {fmt_brl(inad_at)}
                        </td>

                    </tr>

                    <tr>

                        <td>
                            PDD
                        </td>

                        <td>
                            {fmt_brl(pdd_ret)}
                        </td>

                        <td>
                            {fmt_brl(pdd_ant)}
                        </td>

                        <td>
                            {fmt_brl(pdd_at)}
                        </td>

                    </tr>

                </tbody>

            </table>

            <p>{txt_3m}</p>

        </div>

    </div>

    <div class="footer">

        Relatório automático ·
        Gestão Financeira ·
        PDD

    </div>

</div>

</body>
</html>
"""

# =====================================================
# SALVAR
# =====================================================

Path(PASTA_SAIDA).mkdir(
    parents=True,
    exist_ok=True
)

arquivo_saida = (
    Path(PASTA_SAIDA)
    / f"email_fechamento_{dt_at.strftime('%Y_%m')}.html"
)

with open(
    arquivo_saida,
    "w",
    encoding="utf-8"
) as f:

    f.write(html)

print()
print("=" * 55)
print("✅ EMAIL GERADO COM SUCESSO")
print(f"📁 {arquivo_saida}")
print("=" * 55)