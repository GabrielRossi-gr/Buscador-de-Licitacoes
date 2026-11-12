
import pandas as pd
from pathlib import Path

# ==========================================
# CONFIGURAÇÕES
# ==========================================

# Pasta onde estão os arquivos
pasta_entrada = Path(__file__).resolve().parent

# Nome do arquivo consolidado
arquivo_saida = pasta_entrada / "BOLETINS_CONSOLIDADOS.xlsx"

# Prefixo dos arquivos que serão encontrados
prefixo = "boletim_web"

# ==========================================
# LOCALIZAR ARQUIVOS
# ==========================================

arquivos = [
    arquivo for arquivo in pasta_entrada.iterdir()
    if arquivo.is_file()
    and arquivo.name.lower().startswith(prefixo.lower())
    and arquivo.suffix.lower() in [".xlsx", ".xlsm"]
    and not arquivo.name.startswith("~$")
    and arquivo != arquivo_saida
]

if not arquivos:
    print("❌ Nenhum arquivo encontrado.")
    exit()

print(f"📂 Foram encontrados {len(arquivos)} arquivos.")

# ==========================================
# UNIFICAR PLANILHAS
# ==========================================

lista_dados = []

for arquivo in arquivos:
    print(f"📖 Lendo: {arquivo.name}")

    try:
        # Lê a primeira aba de cada arquivo
        df = pd.read_excel(arquivo, engine="openpyxl")

        if not df.empty:
            lista_dados.append(df)

    except Exception as erro:
        print(f"⚠️ Erro ao ler {arquivo.name}: {erro}")

# ==========================================
# CONSOLIDAR E SALVAR
# ==========================================

if lista_dados:

    df_final = pd.concat(lista_dados, ignore_index=True)

    df_final.to_excel(
        arquivo_saida,
        index=False,
        engine="openpyxl"
    )

    print("\n✨ Consolidação concluída!")
    print(f"📊 Total de registros: {len(df_final)}")
    print(f"📁 Arquivo gerado: {arquivo_saida}")

else:
    print("⚠️ Nenhum dado encontrado para consolidar.")
