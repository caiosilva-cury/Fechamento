import os

# Caminho da pasta
caminho = r"C:\Users\caio.silva\Desktop\CAIO\IMPORTACAO\Excel"

# Percorrer arquivos da pasta
for arquivo in os.listdir(caminho):
    
    # Procurar arquivos no padrão "relatorio (x)."
    if arquivo.startswith("relatorio (") and arquivo.endswith(").xlsx"):
        
        # Extrair o número entre parênteses
        numero = arquivo.split("(")[1].split(")")[0]

        # Novo nome
        novo_nome = f"SLDDEV_JUL26_SPE__{numero}.xlsx"  

        # Renomear
        caminho_antigo = os.path.join(caminho, arquivo)
        caminho_novo = os.path.join(caminho, novo_nome)
        os.rename(caminho_antigo, caminho_novo)

        print(f"Renomeado: {arquivo} -> {novo_nome}")
