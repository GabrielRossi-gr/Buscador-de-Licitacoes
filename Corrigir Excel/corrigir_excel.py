import zipfile
import os

def corrigir_planilha_corrompida(caminho_entrada, caminho_saida):
    """
    Remove caracteres inválidos (como \ufffe) de arquivos XML internos
    de um arquivo .xlsx corrompido, gerando um novo arquivo funcional.
    """
    if not os.path.exists(caminho_entrada):
        print(f"Erro: O arquivo '{caminho_entrada}' não foi encontrado.")
        return

    print(f"Iniciando a correção de: {caminho_entrada}")
    
    # Abre o arquivo corrompido para leitura e cria o novo arquivo corrigido
    with zipfile.ZipFile(caminho_entrada, 'r') as arquivo_zip_in:
        with zipfile.ZipFile(caminho_saida, 'w', compression=zipfile.ZIP_DEFLATED) as arquivo_zip_out:
            
            for item in arquivo_zip_in.infolist():
                # Lê os dados binários do arquivo interno atual
                dados = arquivo_zip_in.read(item.filename)
                
                # O caractere corrompido costuma ficar no 'xl/sharedStrings.xml'
                # mas inspecionamos qualquer arquivo .xml por segurança.
                if item.filename.endswith('.xml'):
                    try:
                        # Decodifica o conteúdo em UTF-8
                        conteudo_texto = dados.decode('utf-8', errors='replace')
                        
                        if '\ufffe' in conteudo_texto:
                            # Remove o caractere problemático
                            conteudo_texto = conteudo_texto.replace('\ufffe', '')
                            # Converte de volta para bytes
                            dados = conteudo_texto.encode('utf-8')
                            print(f"-> Caracteres inválidos removidos de: {item.filename}")
                    except Exception as e:
                        # Se não conseguir decodificar o XML, mantém o arquivo original intacto
                        print(f"Aviso: Não foi possível processar {item.filename}: {e}")
                
                # Escreve o arquivo limpo no novo pacote .xlsx
                arquivo_zip_out.writestr(item, dados)
                
    print(f"Sucesso! Planilha corrigida salva em: {caminho_saida}\n")

# --- EXEMPLO DE USO ---
if __name__ == "__main__":

    # Substitua pelos nomes dos seus arquivos
    arquivo_com_erro = "corrigir.xlsx"
    arquivo_recuperado = "planilha_corrigida.xlsx"
    

    corrigir_planilha_corrompida(arquivo_com_erro, arquivo_recuperado)
