import pandas as pd
import os
import urllib
from sqlalchemy import create_engine, text
from datetime import datetime

# =====================================================
# CONFIGURAÇÕES
# =====================================================

PASTA_EXCEL = r"C:\Users\caio.silva\Desktop\CAIO\IMPORTACAO\Excel"
ABA_EXCEL = "Relatório"
TABELA_SQL = "Import_py"
CHUNKSIZE = 50000

# =====================================================
# MAPA DE COLUNAS (EXCEL -> SQL)
# =====================================================

MAPA_COLUNAS = {
    "Código da empresa": "Codigo_Empresa",
    "Empresa": "Empresa",
    "Código do empreendimento": "Codigo_Empreendimento",
    "Empreendimento": "Empreendimento",
    "Unidade Principal": "Unidade_Principal",
    "Número do contrato": "Numero_Contrato",
    "Nome do Cliente": "Nome_Cliente",
    "CPF/CNPJ cliente": "CPF_CNPJ_Cliente",
    "Nome do Cliente Secundário": "Nome_Cliente_Secundario",
    "CPF/CNPJ cliente secundário": "CPF_CNPJ_Cliente_Secundario",
    "Nome do Avalista": "Nome_Avalista",
    "CPF/CNPJ do Avalista": "CPF_CNPJ_Avalista",
    "Telefone do Avalista": "Telefone_Avalista",
    "Título": "Titulo",
    "Tipo de parcela": "Tipo_Parcela",
    "Número da parcela": "Numero_Parcela",
    "Data de vencimento": "Data_Vencimento",
    "Valor original": "Valor_Original",
    "Valor corrigido": "Valor_Corrigido",
    "Pro rata": "Pro_rata",
    "Juros": "Juros",
    "Multa": "Multa",
    "Valor total": "Valor_Total",
    "Valor corrigido Congelado": "Valor_Corrigido_Congelado",
    "Pro rata Congelado": "Pro_rata_Congelado",
    "Juros Congelado": "Juros_Congelado",
    "Multa Congelado": "Multa_Congelado",
    "Valor total Congelado": "Valor_Total_Congelado",
    "Data de pagamento": "Data_Pagamento",
    "Valor de pagamento": "Valor_Pagamento",
    "Dias de Desconto": "Dias_Desconto",
    "Desconto VP": "Desconto_VP",
    "Valor Presente": "Valor_Presente",
    "Categoria de Venda": "Categoria_Venda",
    "Status de Financiamento": "Status_Financiamento",
    "Data da Entrega das Chaves ao Cliente": "Data_Entrega_Chaves",
    "Cliente com Programa de Fidelidade": "Programa_Fidelidade",
    "Data de Adesão Programa de Fidelidade": "Data_Adesao_Fidelidade",
    "Contrato de financiamento": "Contrato_Financiamento",
    "Data de repasse": "Data_Repasse",
    "Data de registro imóvel": "Data_Registro_Imovel"
}

COLUNAS_SQL = list(MAPA_COLUNAS.values())

# =====================================================
# CONEXÃO SQL SERVER
# =====================================================

params = urllib.parse.quote_plus(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=SRJDLK01\\SSPDLKAZ01;"
    "DATABASE=P_PRATA_CAR_DB;"
    "Trusted_Connection=yes;"
)

engine = create_engine(
    f"mssql+pyodbc:///?odbc_connect={params}",
    fast_executemany=True
)

# =====================================================
# INÍCIO
# =====================================================

inicio_execucao = datetime.now()
print(f"🚀 Início da execução: {inicio_execucao}")

# =====================================================
# TRUNCATE DA TABELA
# =====================================================

print("🧹 Limpando tabela Import_py (TRUNCATE)...")
with engine.begin() as conn:
    conn.execute(text(f"TRUNCATE TABLE {TABELA_SQL}"))

# =====================================================
# LISTAR ARQUIVOS
# =====================================================

arquivos = [f for f in os.listdir(PASTA_EXCEL) if f.endswith(".xlsx")]
print(f"📂 Total de arquivos encontrados: {len(arquivos)}")

# =====================================================
# PROCESSAMENTO
# =====================================================

for idx, arquivo in enumerate(arquivos, start=1):
    print(f"\n📄 ({idx}/{len(arquivos)}) Processando {arquivo}")
    caminho = os.path.join(PASTA_EXCEL, arquivo)

    # -------- LEITURA --------
    df = pd.read_excel(
        caminho,
        sheet_name=ABA_EXCEL,
        dtype=str
    )

    print(f"📊 Linhas lidas: {len(df)}")

    # -------- LIMPEZA --------
    df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

    # -------- RENOMEAR COLUNAS --------
    df.rename(columns=MAPA_COLUNAS, inplace=True)

    # -------- MANTER APENAS COLUNAS DO SQL --------
    df = df[[c for c in df.columns if c in COLUNAS_SQL]]

    # -------- DATAS --------
    colunas_data = [
        "Data_Vencimento",
        "Data_Pagamento",
        "Data_Entrega_Chaves",
        "Data_Adesao_Fidelidade",
        "Data_Repasse",
        "Data_Registro_Imovel"
    ]

    for col in colunas_data:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # -------- VALORES NUMÉRICOS --------
    colunas_valor = [
        "Valor_Original",
        "Valor_Corrigido",
        "Juros",
        "Multa",
        "Valor_Total",
        "Valor_Corrigido_Congelado",
        "Valor_Total_Congelado",
        "Valor_Pagamento",
        "Valor_Presente"
    ]

    for col in colunas_valor:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False)
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # -------- INSERT EM LOTE --------
    df.to_sql(
        TABELA_SQL,
        engine,
        if_exists="append",
        index=False,
        chunksize=CHUNKSIZE
    )

    print(f"✅ {arquivo} importado com sucesso")

# =====================================================
# FINALIZAÇÃO
# =====================================================

fim_execucao = datetime.now()
print("\n🏁 Importação concluída com sucesso!")
print(f"⏱ Início : {inicio_execucao}")
print(f"⏱ Fim    : {fim_execucao}")
print(f"⏳ Duração: {fim_execucao - inicio_execucao}")
