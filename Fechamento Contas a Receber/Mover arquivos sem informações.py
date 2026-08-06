import os
import pandas as pd
import shutil

# Caminho das pastas
caminho_origem = r"C:\Users\caio.silva\Desktop\CAIO\IMPORTACAO\Excel"
caminho_destino = r"C:\Users\caio.silva\Desktop\CAIO\IMPORTACAO\Arquivos sem informação"

# Cria a pasta de destino se não existir
os.makedirs(caminho_destino, exist_ok=True)

# Loop pelos arquivos da pasta de origem
for arquivo in os.listdir(caminho_origem):
    if arquivo.endswith(".xlsx") or arquivo.endswith(".xls"):
        caminho_arquivo = os.path.join(caminho_origem, arquivo)
        
        try:
            # Lê o arquivo
            df = pd.read_excel(caminho_arquivo)
            
            # Verifica se só tem cabeçalho (nenhum dado além da linha 1)
            if df.shape[0] == 0:
                destino = os.path.join(caminho_destino, arquivo)
                shutil.move(caminho_arquivo, destino)
                print(f"📂 Movido (sem dados): {arquivo}")
            else:
                print(f"✅ Mantido (com dados): {arquivo}")
        
        except Exception as e:
            print(f"⚠️ Erro ao processar {arquivo}: {e}")
