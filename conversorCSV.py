import pandas as pd
import os

def converter_excel_para_csv(arquivo_entrada, arquivo_saida):
    try:
        print(f"🚀 Iniciando conversão de: {arquivo_entrada}")
        
        # 1. Carrega a planilha
        # O pandas, por padrão, lê a primeira aba (a mais à esquerda)
        df = pd.read_excel(arquivo_entrada, engine='openpyxl')
        
        # 2. Salva como CSV
        # sep=';' define o ponto e vírgula como separador
        # index=False impede que o Pandas crie uma coluna extra de números
        # encoding='utf-8-sig' garante compatibilidade de acentos no Excel
        df.to_csv(arquivo_saida, sep=';', index=False, encoding='utf-8-sig')
        
        print(f"✅ Sucesso! Arquivo salvo em: {arquivo_saida}")
        
    except Exception as e:
        print(f"❌ Ocorreu um erro: {e}")

# ==========================================
# CONFIGURAÇÃO
# ==========================================
if __name__ == "__main__":
    caminho_entrada = 'dados.xlsm'  # Nome do seu arquivo original
    caminho_saida = 'dados_convertidos.csv'
    
    if os.path.exists(caminho_entrada):
        converter_excel_para_csv(caminho_entrada, caminho_saida)
    else:
        print("Arquivo de entrada não encontrado!")