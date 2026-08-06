import os
import pandas as pd

# caminho da pasta
pasta = r"C:\Users\caio.silva\Desktop\CAIO\IMPORTACAO\Excel"

# percorre os arquivos da pasta
for arquivo in os.listdir(pasta):
    if arquivo.lower().endswith(".xlsx"):
        caminho_arquivo = os.path.join(pasta, arquivo)

        try:
            # tamanho do arquivo em KB
            tamanho_kb = os.path.getsize(caminho_arquivo) / 1024

            if tamanho_kb == 0:
                print(f"❌ {arquivo} | Arquivo vazio (0 KB)")
                continue

            # lê o excel (considerando primeira linha como cabeçalho)
            df = pd.read_excel(caminho_arquivo, header=0)

            if df.shape[0] < 1:
                print(f"⚠️ {arquivo} | Sem dados após o cabeçalho | {tamanho_kb:.2f} KB")
                continue

            # primeira célula da segunda linha
            primeira_celula = df.iloc[0, 0]

            print(
                f"✅ {arquivo} | "
                f"Tamanho: {tamanho_kb:.2f} KB | "
                f"Primeira célula da 2ª linha: {primeira_celula}"
            )

        except Exception as e:
            print(f"❌ {arquivo} | Erro ao ler arquivo: {e}")
