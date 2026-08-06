# Pipeline Cobrança/PDD — Saldo Devedor → SQL → Relatórios

Fluxo mensal de 4 etapas: extrai o relatório de saldo devedor do Sienge, importa o que for necessário para o SQL Server, carrega a base de PDD (vinda do Power BI) e por fim gera os relatórios finais em HTML.

## Visão geral do fluxo

```
ETAPA 1 — EXTRAÇÃO
  1a. Exportação do drive atualizado do Chrome.py   (só se o chromedriver estiver desatualizado)
  1b. Extração Saldo devedor.py                      → login no Sienge, exporta 1 xlsx por empresa

  (apoio/organização dos arquivos baixados, conforme necessidade)
  1c. Renomear arquivos - saldo devedor.py
  1d. Analise de arquivos.py
  1e. Mover arquivos sem informações.py

ETAPA 2 — IMPORTAÇÃO PARA O SQL (o que for necessário)
  2a. Importação Excel para SQL - Saldo devedor.py   → tabela Import_py       (P_PRATA_CAR_DB)
  2b. Importacao_para_saldo_devedor.py                → tabela STG.TB_TRN_CAR_IMPORTSALDODEVEDOR_PYTHON_ING (P_PRATA_CAR_DB)
  2c. Importacao para farol de empreendimentos.py     → tabela DIM.TB_COP_CAR_FAROLEMPREENDIMENTOS_PRA (P_OURO_CAR_DB)

ETAPA 3 — BASE PDD (Power BI → SQL)
  3.  Importação_base_pdd.py                          → tabela pdd_base (GestaoFinanceira)

ETAPA 4 — GERAÇÃO DOS RELATÓRIOS
  4a. Portal_inadimplencia.py   → portal_inadimplencia_AAAA_MM.html
  4b. Portal_pdd.py             → relatorio_pdd_AAAA_MM.html
  4c. Analise_resumida.py       → email_fechamento_AAAA_MM.html
```

---

## Etapa 1 — Extração do Saldo Devedor

### 1a. `Exportação do drive atualizado do Chrome.py`
Atualiza o `chromedriver` via `webdriver_manager` e imprime o caminho instalado. Só precisa rodar quando o Chrome for atualizado e o driver antigo parar de funcionar na Etapa 1b (a inicialização do navegador está comentada — hoje o script só baixa o driver, não navega em nada).

### 1b. `Extração Saldo devedor.py`
Automação Selenium que faz login no Sienge (SSO Microsoft) e exporta um `.xlsx` por empresa (IDs 1 a 342, com uma lista de exclusões) para a pasta de Downloads, usando os filtros de juros/multa/congelamento do mês anterior.

⚠️ **E-mail e senha estão em texto puro no código** (`EMAIL`, `SENHA`). Vale mover para variável de ambiente antes de compartilhar/versionar este script.

### 1c–1e. Suporte (organização dos arquivos baixados)
- **`Renomear arquivos - saldo devedor.py`** — renomeia `relatorio (ID).xlsx` → `SLDDEV_JUL26_SPE__ID.xlsx`. O mês/ano está fixo no código (`JUL26`), precisa ajustar manualmente a cada rodada.
- **`Analise de arquivos.py`** — auditoria rápida: tamanho, se está vazio, primeira célula de cada arquivo (conferência antes de importar).
- **`Mover arquivos sem informações.py`** — isola em outra pasta os arquivos que vieram só com cabeçalho (sem linhas de dado).

---

## Etapa 2 — Importação para o SQL

Você tem **dois scripts para saldo devedor** que apontam para o mesmo banco (`P_PRATA_CAR_DB`) mas para tabelas diferentes — importe apenas o que for necessário para o seu fluxo atual:

### 2a. `Importação Excel para SQL - Saldo devedor.py`
- Conexão via `Trusted_Connection` (autenticação Windows).
- `TRUNCATE` + carga em `Import_py`.
- Lê pasta `C:\...\CAIO\IMPORTACAO\Excel`, aba `Relatório`, `dtype=str`.

### 2b. `Importacao_para_saldo_devedor.py`
- Conexão via usuário/senha (`UID=powerbi`).
- Sem `TRUNCATE` — insere direto (`append`) em `STG.TB_TRN_CAR_IMPORTSALDODEVEDOR_PYTHON_ING`.
- Lê pasta `...\Saldo devedor\Arquivos`, converte tipos numéricos (incluindo `Pro_rata`) e inteiros (`Codigo_Empresa`, `Codigo_Empreendimento` etc.) — versão mais completa/robusta que a 2a.

⚠️ Como não há `TRUNCATE` nem filtro por competência aqui, rodar esse script duas vezes com os mesmos arquivos duplica as linhas na STG.

### 2c. `Importacao para farol de empreendimentos.py`
- Carrega o Farol de Empreendimentos em `DIM.TB_COP_CAR_FAROLEMPREENDIMENTOS_PRA` (`P_OURO_CAR_DB`).
- Já tem `TRUNCATE` + `INSERT` dentro da mesma transação, e aborta antes do truncate se o Excel vier vazio — tabela representa o estado atual, não histórico.

---

## Etapa 3 — Base PDD (do Power BI)

### `Importação_base_pdd.py`
Carga mensal da base de PDD (Provisão para Devedores Duvidosos) em `pdd_base` (`GestaoFinanceira`):
- Pede a **data de fechamento** (`AAAA-MM-DD`) interativamente.
- Se já existir dado para essa competência, pede confirmação antes de deletar e recarregar — não afeta meses anteriores.
- Remove linhas de rodapé (Total, filtros, vazias) usando `Chave ID Emp Unid` como referência.
- Mapeia ~80 colunas (saldo devedor, inadimplência, PDD, POC, datas de obra) e insere `dt_fechamento` como primeira coluna.

⚠️ Mesma observação: senha do SQL (`UID=powerbi`) em texto puro.

---

## Etapa 4 — Geração dos Relatórios

Todos os três leem de `FAT.TB_TRN_CAR_CARTEIRAORI_OUR` (`P_OURO_CAR_DB`) e pedem a data de fechamento interativamente — **rodar só depois que a Etapa 3 já tiver os dados do mês carregados**, senão o script aborta por falta de meses suficientes no banco.

### 4a. `Portal_inadimplencia.py`
Gera um dashboard HTML interativo e standalone (Chart.js embutido, dados dos últimos 3 e 12 meses embarcados no próprio arquivo) com:
- KPIs de saldo devedor, inadimplência, PDD e PDD sobre POC, comparando os últimos 3 meses.
- Aging por empreendimento, filtros por regional/empreendimento com autocomplete.
- Precisa de pelo menos 2 meses de dados no banco.
- Salva em `portal_inadimplencia_AAAA_MM.html`.

### 4b. `Portal_pdd.py`
Relatório HTML comparativo (mês atual x mês anterior, encontrado automaticamente no banco) com cards por empreendimento, aging, status de financiamento e status de obra, com filtros e recálculo em JS.
Salva em `relatorio_pdd_AAAA_MM.html`.

### 4c. `Analise_resumida.py`
E-mail de fechamento mensal em HTML (com logo embutido em base64) — resume unidades, saldo devedor, inadimplência e PDD com texto explicativo gerado automaticamente a partir das variações percentuais. Precisa de pelo menos 3 meses de dados no banco.
Salva em `email_fechamento_AAAA_MM.html`.

---

## Pontos de atenção gerais

| Item | Onde | Sugestão |
|---|---|---|
| Credenciais em texto puro (Sienge e SQL) | `Extração Saldo devedor.py`, `Importação_base_pdd.py`, `Importacao_para_saldo_devedor.py`, `Importacao para farol...py`, `Portal_inadimplencia.py`, `Portal_pdd.py`, `Analise_resumida.py` | Mover para variável de ambiente / cofre de segredos antes de versionar |
| Dois scripts fazendo a mesma importação de saldo devedor para tabelas diferentes | `Importação Excel para SQL - Saldo devedor.py` (2a) vs `Importacao_para_saldo_devedor.py` (2b) | Confirmar qual é a versão vigente e aposentar a outra, para não gerar confusão sobre qual rodar |
| Carga sem controle de competência (`append` puro, sem truncate/delete por mês) | `Importacao_para_saldo_devedor.py` (STG) | Adicionar coluna de competência + lógica de substituição, como já existe em `Importação_base_pdd.py` |
| Mês/ano fixo no nome do arquivo | `Renomear arquivos - saldo devedor.py` (`JUL26`) | Parametrizar com `datetime.now()` |
| `df.applymap` (deprecado no pandas ≥ 2.1) | `Importação Excel para SQL - Saldo devedor.py` | Trocar por `df.map(...)` (já corrigido nos scripts mais novos) |
| Relatórios (Etapa 4) dependem de a Etapa 3 já ter rodado para o mês | `Portal_inadimplencia.py`, `Portal_pdd.py`, `Analise_resumida.py` | Rodar sempre na ordem 1 → 2 → 3 → 4 num mesmo fechamento mensal |

---

*Gerado a partir da leitura dos scripts fornecidos, organizados na ordem de execução: extração → importação SQL → base PDD → relatórios.*
