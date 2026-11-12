# Motor de filtragem dinâmico.
print("✅ Motor de filtragem dinâmico")

import pandas as pd
import re
import os
from openpyxl import load_workbook, Workbook
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText

# ==========================================
# CONFIGURAÇÕES
# ==========================================
tamanho_lote = 500
arquivo_entrada = 'dados.xlsm'
arquivo_saida = 'RELATORIO.xlsx'

# ==========================================
# SUAS CLASSES DE FILTRO (ORIGINAIS INALTERADAS)
# ==========================================
class Filtro:
    def __init__(self, coluna_idx, nome):
        self.coluna_idx = coluna_idx
        self.nome = nome
    def aplicar(self, df):
        raise NotImplementedError()
    def _obter_coluna(self, df):
        return df.columns[self.coluna_idx]

class FiltroSimples(Filtro):
    def __init__(self, coluna_idx, tipo, palavras, nome):
        super().__init__(coluna_idx, nome)
        self.tipo = tipo
        self.padrao_regex = r'\b(' + '|'.join(re.escape(p) for p in palavras) + r')\b'
    def aplicar(self, df):
        if df.empty: return df
        nome_coluna = self._obter_coluna(df)
        serie_texto = df[nome_coluna].fillna("").astype(str)
        if self.tipo == 0:
            extracted = serie_texto.str.extract(self.padrao_regex, flags=re.IGNORECASE, expand=False)
            mask = extracted.notna()
            df_filtrado = df[mask].copy()
            df_filtrado['Termo_Match'] = extracted[mask].values
            return df_filtrado
        elif self.tipo == 1:
            mask = serie_texto.str.contains(self.padrao_regex, flags=re.IGNORECASE, na=False, regex=True)
            return df[~mask].copy()

class FiltroCondicional(Filtro):
    def __init__(self, coluna_idx, tipo, gatilhos, adicionais, nome):
        super().__init__(coluna_idx, nome)
        self.tipo = tipo 
        self.regex_gatilhos = '|'.join(re.escape(g) for g in gatilhos)
        self.regex_adicionais = '|'.join(re.escape(a) for a in adicionais)
    def aplicar(self, df):
        if df.empty: return df
        nome_coluna = self._obter_coluna(df)
        serie_texto = df[nome_coluna].fillna("").astype(str)
        mask_gatilho = serie_texto.str.contains(self.regex_gatilhos, flags=re.IGNORECASE, na=False, regex=True)
        mask_adicional = serie_texto.str.contains(self.regex_adicionais, flags=re.IGNORECASE, na=False, regex=True)
        if self.tipo == 2: mask_final = ~mask_gatilho | mask_adicional
        elif self.tipo == 3: mask_final = ~(mask_gatilho & mask_adicional)
        return df[mask_final].copy()

# ==========================================
# MOTOR DE FILTRAGEM (COM ACUMULADOR PARA TABELA FINAL)
# ==========================================
class MotorFiltragem:
    def __init__(self, titulo):
        self.titulo = titulo
        self.filtros = []
        self.historico_abas = []
        self.acumulador_resultados_finais = [] # Guarda o resultado da última etapa de cada lote

    def adicionar_filtro(self, filtro_obj):
        self.filtros.append(filtro_obj)

    def processar_lote(self, df_chunk):
        df_atual = df_chunk.copy()
        self.historico_abas = [("0_Base_Original", df_atual.copy())]
        for filtro in self.filtros:
            df_atual = filtro.aplicar(df_atual)
            self.historico_abas.append((filtro.nome, df_atual.copy() if df_atual is not None else pd.DataFrame()))
            if df_atual is None or df_atual.empty: break
        
        # Se restou algo após o último filtro, guarda na lista para a Tabela Final
        if df_atual is not None and not df_atual.empty:
            self.acumulador_resultados_finais.append(df_atual.copy())

    def salvar_append(self, caminho_saida, is_first_chunk):
        if not self.historico_abas: return
        
        if is_first_chunk:
            with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
                for nome, df in reversed(self.historico_abas):
                    df.to_excel(writer, sheet_name=nome[:31], index=False)
        else:
            wb = load_workbook(caminho_saida)
            for nome, df in reversed(self.historico_abas):
                if df.empty: continue
                nome_aba = nome[:31]
                if nome_aba not in wb.sheetnames: wb.create_sheet(nome_aba)
                ws = wb[nome_aba]
                for _, row in df.iterrows():
                    linha_valores = [row[c] for c in df.columns]
                    ws.append(linha_valores)
            wb.save(caminho_saida)
            wb.close()
        self.historico_abas = []

# ==========================================
# PASSO 1: INJETAR A TABELA FINAL NA 1ª POSIÇÃO USANDO A MEMÓRIA
# ==========================================
def inserir_tabela_final(caminho_arquivo, motor_instancia):
    print("📋 Inserindo a 'Tabela Final' consolidada na primeira posição...")
    
    if not motor_instancia.acumulador_resultados_finais:
        print("⚠️ Nenhum resultado final acumulado.")
        return

    # Junta todos os pedaços finais de cada lote em um único DataFrame consolidado
    df_tabela_final = pd.concat(motor_instancia.acumulador_resultados_finais, ignore_index=True)

    # Lê todas as abas que já foram salvas no arquivo Excel
    xl = pd.ExcelFile(caminho_arquivo)
    dados_outras_abas = {aba: pd.read_excel(caminho_arquivo, sheet_name=aba) for aba in xl.sheet_names}

    # Reescreve o arquivo colocando a 'Tabela Final' como a primeira aba (mais à esquerda)
    with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
        df_tabela_final.to_excel(writer, sheet_name='Tabela Final', index=False)
        
        for nome_aba, df_aba in dados_outras_abas.items():
            if nome_aba != 'Tabela Final':
                df_aba.to_excel(writer, sheet_name=nome_aba[:31], index=False)
        
    print("✨ 'Tabela Final' inserida e posicionada totalmente à esquerda com sucesso!")

# ==========================================
# PASSO 2: REMOVER DUPLICADAS DA TABELA FINAL
# ==========================================
def remover_duplicadas_tabela_final(caminho_arquivo):
    print("🧹 Removendo linhas duplicadas da 'Tabela Final'...")
    
    wb = load_workbook(caminho_arquivo)
    if 'Tabela Final' not in wb.sheetnames:
        wb.close()
        return
    
    ws = wb['Tabela Final']
    seen = set()
    rows_to_delete = []
    
    for row_idx in range(2, ws.max_row + 1):
        row_values_str = tuple(str(cell.value) if cell.value is not None else "" for cell in ws[row_idx])
        
        if row_values_str in seen:
            rows_to_delete.append(row_idx)
        else:
            seen.add(row_values_str)
            
    for row_idx in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(row_idx)
        
    wb.save(caminho_arquivo)
    wb.close()
    print("✨ Duplicadas removidas da 'Tabela Final' com sucesso!")

# ==========================================
# PASSO 3: ALGORITMO DE PINTURA NA TABELA FINAL
# ==========================================
def pintar_tabela_final(caminho_arquivo, lista_filtros):
    print("🎨 Iniciando pintura das palavras na 'Tabela Final'...")
    
    wb = load_workbook(caminho_arquivo)
    if 'Tabela Final' not in wb.sheetnames:
        wb.close()
        return

    palavras_alvo = set()
    for f in lista_filtros:
        if isinstance(f, FiltroSimples):
            palavras = f.padrao_regex.replace(r'\b(', '').replace(r')\b', '').split('|')
            for p in palavras:
                palavras_alvo.add(re.sub(r'\\(.)', r'\1', p))
        elif isinstance(f, FiltroCondicional):
            for g in f.regex_gatilhos.split('|'):
                palavras_alvo.add(re.sub(r'\\(.)', r'\1', g))
            for a in f.regex_adicionais.split('|'):
                palavras_alvo.add(re.sub(r'\\(.)', r'\1', a))

    if not palavras_alvo:
        wb.close()
        return

    padrao_global = r'\b(' + '|'.join(re.escape(p) for p in palavras_alvo if p) + r')\b'
    regex_compilada = re.compile(padrao_global, re.IGNORECASE)
    fonte_vermelha = InlineFont(color="FF0000", b=True)

    ws = wb['Tabela Final']
    
    for row in range(2, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col)
            valor_celula = cell.value
            
            if valor_celula and not isinstance(valor_celula, CellRichText):
                texto_str = str(valor_celula)
                match = regex_compilada.search(texto_str)
                
                if match:
                    inicio, fim = match.span()
                    palavra_encontrada = texto_str[inicio:fim]
                    
                    texto_rico = CellRichText()
                    if inicio > 0:
                        texto_rico.append(texto_str[:inicio])
                        
                    texto_rico.append(TextBlock(fonte_vermelha, palavra_encontrada))
                    
                    if fim < len(texto_str):
                        texto_rico.append(texto_str[fim:])
                        
                    cell.value = texto_rico
                    break 

    wb.save(caminho_arquivo)
    wb.close()
    print("✨ Pintura concluída na 'Tabela Final' com sucesso!")

# ==========================================
# EXECUÇÃO PRINCIPAL
# ==========================================
if __name__ == "__main__":
    motor = MotorFiltragem("Relatorio de Ativos")

    f1 = FiltroSimples(7, 0, ['BOOSTER','BOOSTERs','SBL','BOMBEAMENTO','BOMBA','bombas','reservatório', 'reservatorio','tanque','estação','estacao','ETE','ETA','EEE','abraçadeira','abracadeira','acoplamento',], "Filtro geral")
    f2 = FiltroCondicional(7, 3, ['bomba','BOMBEAMENTO'], ['manutenção','manutencao','calor','concreto','combustível','dosadoras','dosador','SUCÇÃO','banheiro','banheiros',], 'filtro bomba')
    f3 = FiltroCondicional(7, 3, ['estação','estacao','ETE','ETA','EEE'], ['Remoção','limpeza','manutenção','manutencao','Rádio','REFORMA','laboratoriais','laboratorial',"laboratório",'laboratorio','GÁS','AQUECEDOR'], 'filtro estação')
    f4 = FiltroCondicional(7, 3, ['tanque','reservatório','reservatorio'], ["CAIXA D'ÁGUA","barril","MEDICINAIS",'medicina','Nitrogênio','agitadores','ESPARGIDOR','fossa'], 'filtro reservatório')
    f5 = FiltroSimples(7, 1, ['EPIS','proteção','incendio','incendios','incêndio','ferramentas','salas','CARRETA','caminhão','caminhões','ares-condicionados','ar-condicionado','pipa','andaime','retífica','odontológicas','odontológia','Solar','Diesel','Farmacêutico'], "filtro de lixo")
    f6 = FiltroSimples(7, 1, ['manutenção', 'limpeza'], "Filtro manutenção")
    f7 = FiltroSimples(5, 1, ['urgente'], "Filtro urgente")

    motor.adicionar_filtro(f1)
    motor.adicionar_filtro(f2)
    motor.adicionar_filtro(f3)
    motor.adicionar_filtro(f4)
    motor.adicionar_filtro(f5)
    motor.adicionar_filtro(f6)
    motor.adicionar_filtro(f7)

    lista_filtros_objeto = [f1, f2, f3, f4, f5, f6, f7]

    if os.path.exists(arquivo_entrada):
        pulo_linhas = 0
        continuar_processando = True
        primeiro_lote = True

        colunas = pd.read_excel(arquivo_entrada, nrows=0).columns.tolist()

        while continuar_processando:
            print(f"📖 Lendo lote a partir da linha {pulo_linhas}...")
            
            df_lote = pd.read_excel(
                arquivo_entrada, 
                skiprows=pulo_linhas + 1, 
                nrows=tamanho_lote, 
                names=colunas,
                engine='openpyxl'
            )

            if df_lote.empty:
                continuar_processando = False
                break

            print(f"⚙️ Processando {len(df_lote)} linhas...")
            motor.processar_lote(df_lote)

            motor.salvar_append(arquivo_saida, primeiro_lote)
            primeiro_lote = False

            pulo_linhas += len(df_lote)

            if len(df_lote) < tamanho_lote:
                continuar_processando = False

        print(f"\n🔄 Todos os lotes processados. Iniciando sequência de pós-processamento...")
        
        inserir_tabela_final(arquivo_saida, motor)      # 1. Cria e insere a Tabela Final na 1ª posição usando a memória
        remover_duplicadas_tabela_final(arquivo_saida) # 2. Remove duplicados dela
        pintar_tabela_final(arquivo_saida, lista_filtros_objeto) # 3. Roda o algoritmo de pintura

        print(f"\n✨ Processo completo finalizado com sucesso!")
    else:
        print("❌ Arquivo de entrada não encontrado.")