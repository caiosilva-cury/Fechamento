# Pipeline de Importação — Saldo Devedor / PDD / Farol de Empreendimentos

Conjunto de scripts Python que automatizam a extração de relatórios do Sienge, o tratamento dos arquivos Excel gerados e a carga em SQL Server (bases `cardbbi` e `GestaoFinanceira`).

## Visão geral do fluxo

```
1. Exportação do drive atualizado do Chrome.py   → garante chromedriver atualizado
2. Extração Saldo devedor.py                      → login no Sienge + exporta 1 xlsx por empresa
3. Renomear arquivos - saldo devedor.py            → renomeia "relatorio (ID).xlsx" → padrão SLDDEV
4. Analise de arquivos.py                          → auditoria: tamanho, linhas, 1ª célula
5. Mover arquivos sem informações.py               → isola arquivos vazios (só cabeçalho)
6. Importação Excel para SQL - Saldo devedor.py    → TRUNCATE + carga em massa (Import_py)
7. Importação_base_pdd.py                          → carga mensal de PDD (com dedupe por competência)
8. Importacao para farol de empreendimentos.py     → carga do Farol de Empreendimentos
```

Os passos 1–5 processam a pasta local:
`C:\Users\caio.silva\Desktop\CAIO\IMPORTACAO\Excel`

---

## 1. `Exportação do drive atualizado do Chrome.py`
**Função:** baixa/atualiza o `chromedriver` via `webdriver_manager` e imprime o caminho instalado.
**Status:** driver de fato (linhas de inicialização do Chrome) está comentado — hoje só serve para atualizar o driver, não navega em nada.
**Dependências:** `selenium`, `webdriver_manager`.

## 2. `Extração Saldo devedor.py`
**Função:** automação Selenium que:
- Faz login no Sienge (via Microsoft SSO) com e-mail/senha fixos no código.
- Configura filtros do relatório (juros, congelamento de acréscimos, datas — usa automaticamente o mês anterior ao atual).
- Percorre as empresas de ID 1 a 342 (pulando uma lista de IDs excluídos) e exporta um `.xlsx` por empresa para a pasta de Downloads.
- Tem fallback: se o download não for detectado em 60s, copia um arquivo modelo (`relatorio.xlsx`) renomeado como placeholder.

⚠️ **Atenção de segurança:** e-mail e senha estão em texto puro no código (`EMAIL`, `SENHA`). Recomendo mover para variável de ambiente ou um cofre de segredos (ex: `keyring`, Azure Key Vault, `.env` fora do controle de versão) antes de versionar ou compartilhar este script.

**Dependências:** `selenium`, chromedriver compatível com o Chrome instalado.

## 3. `Renomear arquivos - saldo devedor.py`
**Função:** varre a pasta `Excel`, identifica arquivos no padrão `relatorio (ID).xlsx` e renomeia para `SLDDEV_JUL26_SPE__ID.xlsx`.
**Ponto de atenção:** o mês/ano (`JUL26`) está fixo no código — precisa ser atualizado manualmente a cada rodada mensal, ou parametrizado (ex: usando `datetime.now()`).

## 4. `Analise de arquivos.py`
**Função:** auditoria rápida de todos os `.xlsx` da pasta — reporta tamanho em KB, se está vazio, quantas linhas de dados tem e o valor da primeira célula (útil para conferência visual antes da carga).

## 5. `Mover arquivos sem informações.py`
**Função:** lê cada `.xlsx`/`.xls` da pasta de origem; se não houver nenhuma linha de dado (só cabeçalho), move o arquivo para uma subpasta `Arquivos sem informação`, mantendo a pasta principal limpa para a importação.

## 6. `Importação Excel para SQL - Saldo devedor.py`
**Função:**
- Conecta no SQL Server (`cardbbi`, autenticação Windows/Trusted Connection).
- Faz `TRUNCATE` na tabela `Import_py` (⚠️ apaga tudo antes de recarregar — não é incremental).
- Lê todos os `.xlsx` da pasta, aba `Relatório`.
- Renomeia colunas conforme mapa Excel→SQL, converte datas e valores numéricos (tratando formato BR: `.` milhar / `,` decimal).
- Insere em lote (`chunksize=50000`) por arquivo.

## 7. `Importação_base_pdd.py`
**Função:** carga mensal da base de PDD (Provisão para Devedores Duvidosos) na tabela `pdd_base` (`GestaoFinanceira`).
- Pede interativamente a **data de fechamento** (`AAAA-MM-DD`) no console.
- Verifica se já existem registros para essa competência; se sim, pede confirmação antes de deletar e recarregar (evita duplicidade sem apagar meses anteriores).
- Remove linhas de rodapé (Total, filtros aplicados, vazias) usando a coluna-chave `Chave ID Emp Unid`.
- Mapeia ~80 colunas (saldo devedor, inadimplência, PDD, POC, datas de obra etc.), converte tipos e insere `dt_fechamento` como primeira coluna.

⚠️ Mesma observação de segurança: senha do SQL (`UID=powerbi`) está em texto puro no código.

## 8. `Importacao para farol de empreendimentos.py`
**Função:** carga da planilha "Farol de empreendimentos" (status de obra, entrega, jurídico, financiamento etc.) na tabela `FAROL_DE_EMPREENDIMENTOS` (`GestaoFinanceira`).
- Mapeia ~90 colunas, converte percentuais/valores numéricos e ~35 colunas de data.
- Usa `if_exists="append"` — também não é incremental por competência (diferente do script de PDD), então repetir a execução duplica linhas.

---

## Pontos de atenção gerais (recomendações)

| Item | Onde | Sugestão |
|---|---|---|
| Credenciais em texto puro | `Extração Saldo devedor.py`, `Importação_base_pdd.py`, `Importacao para farol de empreendimentos.py`, `Importação Excel para SQL...py` | Mover para variáveis de ambiente / cofre de segredos |
| Mês/ano fixo no nome do arquivo | `Renomear arquivos...py` (`JUL26`) | Parametrizar com `datetime.now()` |
| Carga não incremental (`append` sem controle de competência) | `Importacao para farol de empreendimentos.py`, `Importação Excel para SQL...py` | Adicionar coluna de data de carga + lógica de `DELETE`/`TRUNCATE` por período, como já existe no script de PDD |
| `df.applymap` (deprecado no pandas ≥ 2.1) | `Importação Excel para SQL...py` | Trocar por `df.map(...)` (já usado corretamente no script de PDD) |
| Driver do Chrome baixado mas não usado | `Exportação do drive atualizado do Chrome.py` | Ou remover comentários e integrar de fato, ou simplificar o script para só o que ele faz hoje |
| Loop de 342 empresas sem paralelismo | `Extração Saldo devedor.py` | Processo é sequencial e pode levar horas; considerar rodar por lote/paralelo se o portal permitir |

---

*Gerado automaticamente a partir da leitura dos 8 scripts fornecidos.*
