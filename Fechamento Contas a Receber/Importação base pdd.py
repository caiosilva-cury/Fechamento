import pandas as pd
import os
import urllib
from sqlalchemy import create_engine, text
from datetime import datetime
import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# =====================================================
# CONFIGURAÇÕES
# =====================================================

ARQUIVO_EXCEL = r"C:\Users\srvautbicar\Desktop\PROCESSOS\import para SQL\Base PDD\Arquivos\base_pdd.xlsx"
ABA_EXCEL     = "Export"
TABELA_SQL    = "TB_TRN_CAR_CARTEIRAORI_OUR"
SCHEMA_SQL    = "FAT"
CHUNKSIZE     = 50000

# =====================================================
# PARÂMETRO: DATA DE FECHAMENTO
# =====================================================

print("=" * 55)
print("  IMPORTAÇÃO PDD - BASE MENSAL")
print("=" * 55)
print()
print("Informe a data de fechamento do mês que está importando.")
print("Formato esperado: AAAA-MM-DD  (ex: 2026-03-31)")
print()

while True:
    entrada = input("📅 Data de fechamento: ").strip()
    try:
        data_fechamento = datetime.strptime(entrada, "%Y-%m-%d").date()
        print(f"\n✅ Data de fechamento definida: {data_fechamento}")
        break
    except ValueError:
        print("❌ Formato inválido. Use AAAA-MM-DD (ex: 2026-03-31)\n")

# =====================================================
# MAPA DE COLUNAS (EXCEL -> SQL)
# =====================================================

MAPA_COLUNAS = {
    "Chave ID Emp Unid"                                         : "chave_id_emp_unid",
    "ID Empresa"                                                : "id_empresa",
    "Empresa"                                                   : "empresa",
    "CNPJ Empresa"                                              : "cnpj_empresa",
    "ID Empreendimento"                                         : "id_empreendimento",
    "Empreendimento"                                            : "empreendimento",
    "Unidade"                                                   : "unidade",
    "Centro de Custo"                                           : "centro_de_custo",
    "Regional"                                                  : "regional",
    "Status de Obra"                                            : "status_de_obra",
    # status_de_obra_anterior       -> não vem no arquivo fonte (ficará NULL)
    "Previsão de Entrega"                                       : "previsao_de_entrega",
    "ID Cliente"                                                : "id_cliente",
    "Nome do Cliente"                                           : "nome_do_cliente",
    "CPF/CNPJ Cliente"                                          : "cpf_cnpj_cliente",
    "Data de Nascimento"                                        : "data_de_nascimento",
    "E-mail"                                                    : "email",
    "Estado Civil"                                              : "estado_civil",
    "Profissão"                                                 : "profissao",
    "Telefones"                                                 : "telefones",
    "Estoque Comercial"                                         : "estoque_comercial",
    "Categoria da Venda"                                        : "categoria_da_venda",
    "Status de Financiamento"                                   : "status_de_financiamento",
    # status_de_financiamento_ant   -> não vem no arquivo fonte (ficará NULL)
    "Unidade com Saldo Devedor?"                                : "unidade_com_saldo_devedor",
    "Status da Unidade"                                         : "status_da_unidade",
    "Aging da Unidade"                                          : "aging_da_unidade",
    # aging_da_unidade_anterior     -> não vem no arquivo fonte (ficará NULL)
    "Saldo Devedor Carteira"                                    : "saldo_dev_carteira",
    "Saldo Devedor Carteira (Com Multa e Juros)"                : "saldo_dev_carteira_multa_juros",
    "Saldo Devedor FI (Principal)"                              : "saldo_dev_fi_principal",
    "Saldo Devedor FI (Complemento)"                            : "saldo_dev_fi_complemento",
    "Saldo Devedor FI (P+C)"                                    : "saldo_dev_fi_pc",
    "Saldo Devedor FI (Com Multa e Juros)"                      : "saldo_dev_fi_multa_juros",
    "Saldo Devedor Total"                                       : "saldo_dev_total",
    "Saldo Devedor Total (Com Multa e Juros)"                   : "saldo_dev_total_multa_juros",
    "Percentual de Pró Soluto"                                  : "percentual_pro_soluto",
    "Inadimplência Carteira"                                    : "inadimplencia_carteira",
    "Inadimplência Carteira (Multa)"                            : "inadimplencia_carteira_multa",
    "Inadimplência Carteira (Juros)"                            : "inadimplencia_carteira_juros",
    "Inadimplência Carteira Total"                              : "inadimplencia_carteira_total",
    "Inadimplência Carteira (Juros, congelamento até 360 dias)" : "inadimplencia_carteira_juros_cong",
    "Carteira a Vencer"                                         : "carteira_a_vencer",
    "Inadimplência FI"                                          : "inadimplencia_fi",
    "Inadimplência FI (Multa)"                                  : "inadimplencia_fi_multa",
    "Inadimplência FI (Juros)"                                  : "inadimplencia_fi_juros",
    "Inadimplência FI Total"                                    : "inadimplencia_fi_total",
    "Inadimplência FI (Juros, congelamento até 360 dias)"       : "inadimplencia_fi_juros_cong",
    "FI a Vencer"                                               : "fi_a_vencer",
    "Valor Recebido Ato"                                        : "valor_recebido_ato",
    "Valor Recebido Carteira"                                   : "valor_recebido_carteira",
    "Valor Recebido FI"                                         : "valor_recebido_fi",
    "Valor Recebido Total"                                      : "valor_recebido_total",
    "Parcela de DI"                                             : "parcela_di",
    "Parcela de DI (Multa)"                                     : "parcela_di_multa",
    "Parcela de DI (Juros)"                                     : "parcela_di_juros",
    "Parcela de DI Total"                                       : "parcela_di_total",
    "Parcela de DI (Juros, congelamento até 360 dias)"          : "parcela_di_juros_cong",
    "Valor Total da Carteira"                                   : "valor_total_carteira",
    "Valor Total da Carteira (Com Multa e Juros)"               : "valor_total_carteira_multa_juros",
    "Valor Total da Carteira (Congelamento)"                    : "valor_total_carteira_congelamento",
    "Saldo Devedor para PDD"                                    : "saldo_dev_pdd",
    "Saldo Devedor para PDD (Com Multa e Juros)"                : "saldo_dev_pdd_multa_juros",
    "Saldo Devedor para PDD (Congelamento)"                     : "saldo_dev_pdd_congelamento",
    "Percentual de PDD"                                         : "percentual_pdd",
    # percentual_pdd_anterior       -> não vem no arquivo fonte (ficará NULL)
    "Valor de PDD"                                              : "valor_pdd",
    # valor_pdd_anterior            -> não vem no arquivo fonte (ficará NULL)
    "Percentual de PDD - Contabilidade"                         : "percentual_pdd_contabilidade",
    "Valor de PDD - Contabilidade"                              : "valor_pdd_contabilidade",
    "Valor de PDD (Com Multa e Juros)"                          : "valor_pdd_multa_juros",
    "Valor de PDD_Com Multa e Juros - Contabilidade"            : "valor_pdd_multa_juros_contabilidade",
    "Valor de PDD (Congelamento)"                               : "valor_pdd_congelamento",
    "Valor de PDD_Congelamento - Contabilidade"                 : "valor_pdd_congelamento_contabilidade",
    "POC Obras"                                                 : "poc_obras",
    # poc_obras_anterior            -> não vem no arquivo fonte (ficará NULL)
    "Valor de PDD sobre POC"                                    : "valor_pdd_poc",
    "Valor de PDD sobre POC - Contabilidade"                    : "valor_pdd_poc_contabilidade",
    "Valor de PDD sobre POC (Com Multa e Juros)"                : "valor_pdd_poc_multa_juros",
    "Valor de PDD sobre POC_Com Multa e Juros - Contabilidade"  : "valor_pdd_poc_multa_juros_contab",
    "Valor de PDD sobre POC (Congelamento)"                     : "valor_pdd_poc_congelamento",
    "Valor de PDD sobre POC_Congelamento - Contabilidade"       : "valor_pdd_poc_congelamento_contab",
    "Data de Lançamento"                                        : "data_lancamento",
    "Dias de Lançamento da Obra"                                : "dias_lancamento_obra",
    "Data Habite-se Prevista"                                   : "data_habitese_prevista",
    "Data Habite-se Realizada"                                  : "data_habitese_realizada",
    "Data AGI Prevista"                                         : "data_agi_prevista",
    "Data AGI Realizada"                                        : "data_agi_realizada",
    "Data de início das obras"                                  : "data_inicio_obras",
    "Data de término das obras"                                 : "data_termino_obras",
    "Físico Acumulado"                                          : "fisico_acumulado",
    "Data do Contrato Cury"                                     : "data_contrato_cury",
    "Ano da Venda"                                              : "ano_da_venda",
    "Valor do Contrato Atualizado"                              : "valor_contrato_atualizado",
    "Número do Contrato Cury"                                   : "numero_contrato_cury",
    "Status Chaves"                                             : "status_chaves",
    "Data da Entrega de Chaves"                                 : "data_entrega_chaves",
}

COLUNAS_DATA = [
    "data_lancamento",
    "data_habitese_prevista",
    "data_habitese_realizada",
    "data_agi_prevista",
    "data_agi_realizada",
    "data_inicio_obras",
    "data_termino_obras",
    "data_contrato_cury",
    "data_entrega_chaves",
]

COLUNAS_NUMERICAS = [
    "saldo_dev_carteira", "saldo_dev_carteira_multa_juros",
    "saldo_dev_fi_principal", "saldo_dev_fi_complemento",
    "saldo_dev_fi_pc", "saldo_dev_fi_multa_juros",
    "saldo_dev_total", "saldo_dev_total_multa_juros",
    "percentual_pro_soluto",
    "inadimplencia_carteira", "inadimplencia_carteira_multa",
    "inadimplencia_carteira_juros", "inadimplencia_carteira_total",
    "inadimplencia_carteira_juros_cong", "carteira_a_vencer",
    "inadimplencia_fi", "inadimplencia_fi_multa",
    "inadimplencia_fi_juros", "inadimplencia_fi_total",
    "inadimplencia_fi_juros_cong", "fi_a_vencer",
    "valor_recebido_ato", "valor_recebido_carteira",
    "valor_recebido_fi", "valor_recebido_total",
    "parcela_di", "parcela_di_multa", "parcela_di_juros",
    "parcela_di_total", "parcela_di_juros_cong",
    "valor_total_carteira", "valor_total_carteira_multa_juros",
    "valor_total_carteira_congelamento",
    "saldo_dev_pdd", "saldo_dev_pdd_multa_juros", "saldo_dev_pdd_congelamento",
    "percentual_pdd",
    "valor_pdd",
    "percentual_pdd_contabilidade", "valor_pdd_contabilidade",
    "valor_pdd_multa_juros", "valor_pdd_multa_juros_contabilidade",
    "valor_pdd_congelamento", "valor_pdd_congelamento_contabilidade",
    "poc_obras",
    "valor_pdd_poc", "valor_pdd_poc_contabilidade",
    "valor_pdd_poc_multa_juros", "valor_pdd_poc_multa_juros_contab",
    "valor_pdd_poc_congelamento", "valor_pdd_poc_congelamento_contab",
    "fisico_acumulado", "valor_contrato_atualizado",
]

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

engine = create_engine(
    f"mssql+pyodbc:///?odbc_connect={params}",
    fast_executemany=True
)

# =====================================================
# INÍCIO
# =====================================================

inicio_execucao = datetime.now()
print(f"\n🚀 Início da execução: {inicio_execucao}")

# =====================================================
# VERIFICAR SE JÁ EXISTE DADOS PARA ESSA COMPETÊNCIA
# =====================================================

with engine.begin() as conn:
    resultado = conn.execute(
        text(f"SELECT COUNT(*) FROM {SCHEMA_SQL}.{TABELA_SQL} WHERE data_fechamento = :dt"),
        {"dt": data_fechamento}
    ).scalar()

if resultado > 0:
    print(f"\n⚠️  Atenção: já existem {resultado:,} registros para {data_fechamento} na tabela.")
    print("   Se continuar, os dados existentes serão DELETADOS e substituídos.")
    confirmacao = input("\n   Deseja continuar? (S/N): ").strip().upper()
    if confirmacao != "S":
        print("\n❌ Importação cancelada pelo usuário.")
        exit()
    print(f"\n🧹 Deletando registros existentes de {data_fechamento}...")
    with engine.begin() as conn:
        conn.execute(
            text(f"DELETE FROM {SCHEMA_SQL}.{TABELA_SQL} WHERE data_fechamento = :dt"),
            {"dt": data_fechamento}
        )
    print("✅ Registros anteriores removidos.")

# =====================================================
# LEITURA DO EXCEL
# =====================================================

print(f"\n📂 Lendo arquivo: {ARQUIVO_EXCEL}")
print(f"📋 Aba: {ABA_EXCEL}")

df = pd.read_excel(
    ARQUIVO_EXCEL,
    sheet_name=ABA_EXCEL,
    dtype=str
)

# Remove coluna sem nome (Unnamed)
df = df.loc[:, ~df.columns.str.startswith("Unnamed")]

print(f"📊 Linhas lidas: {len(df):,} | Colunas: {len(df.columns)}")

# =====================================================
# LIMPEZA
# =====================================================

df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

# =====================================================
# REMOVER RODAPÉ (Total, filtros aplicados, linhas vazias)
# A coluna "Chave ID Emp Unid" sempre tem valor nas linhas de dados
# e fica vazia ou com texto livre nas linhas de rodapé
# =====================================================

linhas_antes = len(df)
df = df[df["Chave ID Emp Unid"].notna()]                          # remove nulos
df = df[df["Chave ID Emp Unid"].str.strip() != ""]                # remove vazios
df = df[~df["Chave ID Emp Unid"].str.lower().str.startswith("total")]     # remove linha Total
df = df[~df["Chave ID Emp Unid"].str.lower().str.startswith("filtros")]   # remove rodapé de filtros
linhas_removidas = linhas_antes - len(df)
if linhas_removidas > 0:
    print(f"🗑️  Linhas de rodapé removidas: {linhas_removidas} (Total, filtros, vazias)")

# =====================================================
# RENOMEAR COLUNAS
# =====================================================

df.rename(columns=MAPA_COLUNAS, inplace=True)

# Manter apenas colunas mapeadas
colunas_validas = [c for c in df.columns if c in MAPA_COLUNAS.values()]
df = df[colunas_validas]

# =====================================================
# DATAS
# =====================================================

for col in COLUNAS_DATA:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors="coerce")

# =====================================================
# NUMÉRICOS
# =====================================================

for col in COLUNAS_NUMERICAS:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

# =====================================================
# INTEIROS
# =====================================================

for col in ["id_empresa", "id_empreendimento"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")

for col in ["id_cliente", "ano_da_venda"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")

# =====================================================
# INSERIR COLUNA data_fechamento NA PRIMEIRA POSIÇÃO
# =====================================================

df.insert(0, "data_fechamento", data_fechamento)

# =====================================================
# INSERT EM LOTE
# =====================================================

print(f"\n⬆️  Iniciando insert na tabela [{TABELA_SQL}]...")
print(f"   Registros a inserir: {len(df):,}")
print(f"   Chunk size         : {CHUNKSIZE:,}")

df.to_sql(
    TABELA_SQL,
    engine,
    schema="FAT",
    if_exists="append",
    index=False,
    chunksize=CHUNKSIZE
)

# =====================================================
# FINALIZAÇÃO
# =====================================================

fim_execucao = datetime.now()
duracao = fim_execucao - inicio_execucao

print(f"\n{'=' * 55}")
print(f"  ✅ IMPORTAÇÃO CONCLUÍDA COM SUCESSO!")
print(f"{'=' * 55}")
print(f"  Competência : {data_fechamento}")
print(f"  Registros   : {len(df):,}")
print(f"  Início      : {inicio_execucao.strftime('%H:%M:%S')}")
print(f"  Fim         : {fim_execucao.strftime('%H:%M:%S')}")
print(f"  Duração     : {duracao}")
print(f"{'=' * 55}")
