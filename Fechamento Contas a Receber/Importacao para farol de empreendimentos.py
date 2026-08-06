import pandas as pd
import urllib
from sqlalchemy import create_engine, text
from datetime import datetime
import numpy as np

# =====================================================
# CONFIGURAÇÕES
# =====================================================

ARQUIVO_EXCEL = r"C:\Users\srvautbicar\Desktop\PROCESSOS\import para SQL\Saldo devedor\Arquivos\Farol de empreendimento.xlsx"
ABA_EXCEL = "Farol de empreendimentos"
SCHEMA_SQL = "DIM"
TABELA_SQL = "TB_COP_CAR_FAROLEMPREENDIMENTOS_PRA"
CHUNKSIZE = 5000

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
# MAPEAMENTO Excel → SQL
# =====================================================

RENAME_MAP = {
    "Código Sienge":                            "codigo_sienge",
    "SPE: Nome da conta":                       "spe_nome_conta",
    "CNPJ":                                     "cnpj",
    "Nome de Empreendimento":                   "nome_empreendimento",
    "Endereço":                                 "endereco",
    "Cidade":                                   "cidade",
    "Filial":                                   "filial",
    "Total Unidades":                           "total_unidades",
    "Data de lançamento":                       "data_lancamento",
    "GO Responsável":                           "go_responsavel",
    "Opção de planta":                          "opcao_planta",
    "Status macro":                             "status_macro",
    "Data início das obras":                    "data_inicio_obras",
    "Data término de obras (Interna)":          "data_termino_obras_interna",
    "Data término de obras":                    "data_termino_obras",
    "% físico acumulado":                       "percentual_fisico_acumulado",
    "Mobilização de canteiro":                  "mobilizacao_canteiro",
    "Alvenaria":                                "alvenaria",
    "Acabamento Externo":                       "acabamento_externo",
    "Acabamento Interno":                       "acabamento_interno",
    "Estrutura":                                "estrutura",
    "Fundação":                                 "fundacao",
    "Instalações Elétricas":                    "instalacoes_eletricas",
    "Instalações Hidráulicas":                  "instalacoes_hidraulicas",
    "Pintura":                                  "pintura",
    "Última atualização do  andamento de obra": "ultima_atualizacao_andamento_obra",
    "Data da última atualização mídia":         "data_ultima_atualizacao_midia",
    "Última atualização - Vídeo Drone":         "ultima_atualizacao_video_drone",
    "Assembleia Patrimônio Realizada":          "assembleia_patrimonio_realizada",
    "Data envio ata - jurídico":                "data_envio_ata_juridico",
    "Data de registro da ata":                  "data_registro_ata",
    "Data do próximo relatório":                "data_proximo_relatorio",
    "Data do último relatório":                 "data_ultimo_relatorio",
    "E-mail comissão de representantes":        "email_comissao_representantes",
    "Última atualização - Tour":                "ultima_atualizacao_tour",
    "1ª Visita Realizada":                      "primeira_visita_realizada",
    "2ª Visita Realizada":                      "segunda_visita_realizada",
    "3ª Visita Realizada":                      "terceira_visita_realizada",
    "Cobrança avulsa de TC até:":               "cobranca_avulsa_tc_ate",
    "Agência da conta de financiamento":        "agencia_conta_financiamento",
    "Contratação PJ":                           "contratacao_pj",
    "Modalidade Financiamento PJ":              "modalidade_financiamento_pj",
    "Tipo de financiamento":                    "tipo_financiamento",
    "Data início Financiamento Imobiliário":    "data_inicio_financiamento_imobiliario",
    "Data de entrega contratual CEF":           "data_entrega_contratual_cef",
    "Data de entrega contr. c/ carência CEF":   "data_entrega_contratual_carencia_cef",
    "Data de entrega contratual Cury":          "data_entrega_contratual_cury",
    "Data de entrega contr. c/ carência Cury":  "data_entrega_contratual_carencia_cury",
    "Cartório de registro do Imóvel":           "cartorio_registro_imovel",
    "Escrevente":                               "escrevente",
    "Telefone cartório":                        "telefone_cartorio",
    "Cartório de Notas":                        "cartorio_notas",
    "Data habite-se prevista":                  "data_habite_se_prevista",
    "Data realizada habite-se":                 "data_habite_se_realizada",
    "Data AGI Prevista":                        "data_agi_prevista",
    "Data AGI Realizada":                       "data_agi_realizada",
    "Data Vistoria Área Comum Previsto":        "data_vistoria_area_comum_prevista",
    "Data Vistoria Área Comum Realizada":       "data_vistoria_area_comum_realizada",
    "Data Liberação Chaves Previsto":           "data_liberacao_chaves_prevista",
    "Data Liberação Chaves Realizada":          "data_liberacao_chaves_realizada",
    "Data Workshop de gestão cond. Realizada":  "data_workshop_gestao_cond_realizada",
    "Data realizada CND ISS":                   "data_realizada_cnd_iss",
    "Data prevista matrícula individualizada":  "data_prevista_matricula_individualizada",
    "Data realizada matrícula individualizada": "data_realizada_matricula_individualizada",
    "Data realizada averbação do habite-se":    "data_realizada_averbacao_habite_se",
    "Multa":                                    "multa",
    "Quantidade meses multa":                   "quantidade_meses_multa",
    "Taxa de concessionária":                   "taxa_concessionaria",
    "IPTU individualizado":                     "iptu_individualizado",
    "Administradora do condomínio: Nome da conta": "administradora_condominio_nome_conta",
    "Telefone":                                 "telefone",
    "Nível de criticidade":                     "nivel_criticidade",
    "Relacionamento clientes/Contas a receber": "relacionamento_clientes_contas_receber",
    "Assistência Técnica":                      "assistencia_tecnica",
    "Jurídico":                                 "juridico",
    "Comercial/Repasse":                        "comercial_repasse",
    "Obras":                                    "obras",
    "Financeiro":                               "financeiro",
    "Crédito imobiliário":                      "credito_imobiliario",
    "Estado":                                   "estado",
    "Total de unidades - Pemutante":            "total_unidades_permutante",
    "VGV MÉDIO":                                "vgv_medio",
    "Data visita obra":                         "data_visita_obra",
    "Opção de planta encerrada em:":            "opcao_planta_encerrada_em",
}

COLUNAS_FLOAT = [
    "codigo_sienge", "total_unidades", "percentual_fisico_acumulado",
    "mobilizacao_canteiro", "alvenaria", "acabamento_externo",
    "acabamento_interno", "estrutura", "fundacao", "instalacoes_eletricas",
    "instalacoes_hidraulicas", "pintura", "quantidade_meses_multa",
    "total_unidades_permutante",
]

COLUNAS_MONEY = ["vgv_medio"]

COLUNAS_DATA = [
    "data_lancamento", "data_inicio_obras", "data_termino_obras_interna",
    "data_termino_obras", "ultima_atualizacao_andamento_obra",
    "data_ultima_atualizacao_midia", "ultima_atualizacao_video_drone",
    "data_envio_ata_juridico", "data_registro_ata", "data_proximo_relatorio",
    "data_ultimo_relatorio", "ultima_atualizacao_tour",
    "primeira_visita_realizada", "segunda_visita_realizada",
    "terceira_visita_realizada", "cobranca_avulsa_tc_ate",
    "data_inicio_financiamento_imobiliario", "data_entrega_contratual_cef",
    "data_entrega_contratual_carencia_cef", "data_entrega_contratual_cury",
    "data_entrega_contratual_carencia_cury", "data_habite_se_prevista",
    "data_habite_se_realizada", "data_agi_prevista", "data_agi_realizada",
    "data_vistoria_area_comum_prevista", "data_vistoria_area_comum_realizada",
    "data_liberacao_chaves_prevista", "data_liberacao_chaves_realizada",
    "data_workshop_gestao_cond_realizada", "data_realizada_cnd_iss",
    "data_prevista_matricula_individualizada",
    "data_realizada_matricula_individualizada",
    "data_realizada_averbacao_habite_se", "data_visita_obra",
    "opcao_planta_encerrada_em", "assembleia_patrimonio_realizada",
]

# =====================================================
# INÍCIO
# =====================================================

inicio_execucao = datetime.now()
print(f"🚀 Início: {inicio_execucao}")

# =====================================================
# LEITURA EXCEL
# =====================================================

df = pd.read_excel(ARQUIVO_EXCEL, sheet_name=ABA_EXCEL, engine="openpyxl")
print(f"📊 Linhas lidas: {len(df)}")

# =====================================================
# LIMPEZA E RENOMEAÇÃO
# =====================================================

df.columns = df.columns.str.strip()
df = df.rename(columns=RENAME_MAP)
df = df[[col for col in df.columns if col in RENAME_MAP.values()]]

cols_str = df.select_dtypes(include=["object", "string"]).columns
df[cols_str] = df[cols_str].apply(lambda col: col.str.strip() if col.dtype == "object" else col)

# =====================================================
# CONVERSÃO NUMÉRICA
# =====================================================

def limpa_numero(serie):
    return (
        serie.astype(str)
        .str.replace(r"[^\d,.-]", "", regex=True)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

for col in COLUNAS_FLOAT + COLUNAS_MONEY:
    if col in df.columns:
        if not pd.api.types.is_numeric_dtype(df[col]):
            df[col] = limpa_numero(df[col])
        df[col] = pd.to_numeric(df[col], errors="coerce")

# =====================================================
# CONVERSÃO DE DATAS
# =====================================================

for col in COLUNAS_DATA:
    if col in df.columns:
        df[col] = pd.to_datetime(
            df[col],
            dayfirst=True,
            errors="coerce"
        ).dt.date

# =====================================================
# NULLS
# =====================================================

df = df.replace({np.nan: None})

# =====================================================
# DEBUG
# =====================================================

print(df.head())
print(df.dtypes)

# Se não veio nenhuma linha do Excel, não faz sentido truncar a tabela
# e ficar sem dado nenhum — aborta antes do TRUNCATE.
if df.empty:
    print("⚠️ Nenhuma linha lida do Excel — abortando sem truncar a tabela.")
    raise SystemExit(1)

# =====================================================
# TRUNCATE + INSERT
# (tabela representa o estado atual dos empreendimentos,
# não um histórico — cada execução substitui o conteúdo anterior)
# =====================================================

with engine.begin() as conn:
    conn.execute(text(f"TRUNCATE TABLE [{SCHEMA_SQL}].[{TABELA_SQL}]"))

    df.to_sql(
        TABELA_SQL,
        conn,
        schema=SCHEMA_SQL,
        if_exists="append",
        index=False,
        chunksize=CHUNKSIZE,
    )

# =====================================================
# FINAL
# =====================================================

fim_execucao = datetime.now()
print("\n✅ Importação concluída (truncate + insert)")
print(f"⏱ Tempo total: {fim_execucao - inicio_execucao}")