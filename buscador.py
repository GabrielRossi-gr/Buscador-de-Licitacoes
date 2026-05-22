#      Motor de filtragem dinâmico.

#      Estruturas dos filtros
#      Estrutura 1 SIMPLES -> [coluna, tipo, ['palavra1','palavra2'], 'nome do filtro'],
#      Estrutura 2 COMPOSTO -> [coluna, tipo, ['gatilho1','gatilho2'], ['palavra1','palavra2'], 'nome do filtro'],


#     0 - Incluir se palavra exata existir.
#     1 - Excluir se palavra exata existir.                                     -->   [coluna, 1, ['EPIS','proteção'], "lixo"],
#     2 - Refinamento Condicional ([Gatilho] -> MANTÉM se tiver a [palavra]).
#     3 - Exclusão Condicional ([Gatilho] -> EXCLUI se tiver a [palavra]).      -->   [coluna, 3, ['bomba','BOMBEAMENTO'], ['manutenção','calor'], 'filtro bomba']



# # Configuração dos seus filtros validados
#     motor.adicionar_filtro(FiltroSimples(8, 0, ['BOOSTER','BOOSTERs','SBL','BOMBEAMENTO','BOMBA','bombas','reservatório', 'reservatorio','tanque','estação','estacao','ETE','ETA','EEE','abraçadeira','abracadeira','acoplamento', ], "Filtro geral"))
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['bomba','BOMBEAMENTO'], ['manutenção','manutencao','calor','concreto','combustível','dosadoras','dosador','SUCÇÃO','banheiro','banheiros',], 'filtro bomba'))
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['estação','estacao','ETE','ETA','EEE'], ['prfv','Remoção','limpeza','manutenção','manutencao','Rádio','REFORMA','laboratoriais','laboratorial',"laboratório",'laboratorio','GÁS','AQUECEDOR'], 'filtro estação'))
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['tanque','reservatório','reservatorio'], ["CAIXA D'ÁGUA","barril","MEDICINAIS",'medicina','Nitrogênio','agitadores','ESPARGIDOR'], 'filtro reservatório'))    
    
#     motor.adicionar_filtro(FiltroSimples(8, 1, ['EPIS','proteção','incendio','incendios','ferramentas','salas','CARRETA','caminhão','caminhões','ares-condicionados','ar-condicionado','pipa'], "filtro de lixo"))
#     motor.adicionar_filtro(FiltroSimples(8, 1, ['manutenção', 'limpeza'], "Filtro manutenção"))
#     motor.adicionar_filtro(FiltroSimples(5, 1, ['urgente'], "Filtro urgente"))

# Motor de filtragem dinâmico.
print("✅ Motor de filtragem dinâmico")



import pandas as pd
import re
import os
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# ==========================================
# CONFIGURAÇÕES
# ==========================================
tamanho_lote = 500
arquivo_entrada = 'dados.xlsm'
arquivo_saida = 'RELATORIO.xlsx'

# ==========================================
# SUAS CLASSES DE FILTRO (MANTIDAS EXATAMENTE IGUAIS)
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
# MOTOR ADAPTADO PARA GERADOR DE LINHAS
# ==========================================
class MotorFiltragem:
    def __init__(self, titulo):
        self.titulo = titulo
        self.filtros = []
        self.historico_abas = []

    def adicionar_filtro(self, filtro_obj):
        self.filtros.append(filtro_obj)

    def processar_lote(self, df_chunk):
        df_atual = df_chunk.copy()
        self.historico_abas = [("0_Base_Original", df_atual.copy())]
        for filtro in self.filtros:
            df_atual = filtro.aplicar(df_atual)
            self.historico_abas.append((filtro.nome, df_atual.copy() if df_atual is not None else pd.DataFrame()))
            if df_atual is None or df_atual.empty: break

    def salvar_append(self, caminho_saida, is_first_chunk):
        if not self.historico_abas: return
        if is_first_chunk:
            with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
                for nome, df in reversed(self.historico_abas):
                    if 'Termo_Match' in df.columns and nome == "0_Base_Original":
                         df = df.drop(columns=['Termo_Match'], errors='ignore')
                    df.to_excel(writer, sheet_name=nome[:31], index=False)
        else:
            wb = load_workbook(caminho_saida)
            for nome, df in reversed(self.historico_abas):
                if df.empty: continue
                if 'Termo_Match' in df.columns and nome == "0_Base_Original":
                     df = df.drop(columns=['Termo_Match'], errors='ignore')
                nome_aba = nome[:31]
                if nome_aba not in wb.sheetnames: wb.create_sheet(nome_aba)
                ws = wb[nome_aba]
                for r in dataframe_to_rows(df, index=False, header=False):
                    ws.append(r)
            wb.save(caminho_saida)
            wb.close()
        self.historico_abas = []

# ==========================================
# EXECUÇÃO COM LÓGICA DE ABERTURA/FECHAMENTO ESTRITA
# ==========================================
if __name__ == "__main__":
    motor = MotorFiltragem("Relatorio de Ativos")

    # (Seus filtros aqui...)
    motor.adicionar_filtro(FiltroSimples(7, 0, ['BOOSTER','BOOSTERs','SBL','BOMBEAMENTO','BOMBA','bombas','reservatório', 'reservatorio','tanque','estação','estacao','ETE','ETA','EEE','abraçadeira','abracadeira','acoplamento', ], "Filtro geral"))
    motor.adicionar_filtro(FiltroCondicional(7, 3, ['bomba','BOMBEAMENTO'], ['manutenção','manutencao','calor','concreto','combustível','dosadoras','dosador','SUCÇÃO','banheiro','banheiros',], 'filtro bomba'))
    motor.adicionar_filtro(FiltroCondicional(7, 3, ['estação','estacao','ETE','ETA','EEE'], ['prfv','Remoção','limpeza','manutenção','manutencao','Rádio','REFORMA','laboratoriais','laboratorial',"laboratório",'laboratorio','GÁS','AQUECEDOR'], 'filtro estação'))
    motor.adicionar_filtro(FiltroCondicional(7, 3, ['tanque','reservatório','reservatorio'], ["CAIXA D'ÁGUA","barril","MEDICINAIS",'medicina','Nitrogênio','agitadores','ESPARGIDOR','fossa'], 'filtro reservatório'))    
    
    motor.adicionar_filtro(FiltroSimples(7, 1, ['EPIS','proteção','incendio','incendios','ferramentas','salas','CARRETA','caminhão','caminhões','ares-condicionados','ar-condicionado','pipa'], "filtro de lixo"))
    motor.adicionar_filtro(FiltroSimples(7, 1, ['manutenção', 'limpeza'], "Filtro manutenção"))
    motor.adicionar_filtro(FiltroSimples(5, 1, ['urgente'], "Filtro urgente"))


    if os.path.exists(arquivo_entrada):
        primeiro_lote = True
        pulo_linhas = 0
        continuar_processando = True

        # Pegamos os nomes das colunas uma única vez para não perder o cabeçalho
        colunas = pd.read_excel(arquivo_entrada, nrows=0).columns.tolist()

        while continuar_processando:
            print(f"📖 Lendo lote a partir da linha {pulo_linhas}...")
            
            # 1. ABRE e COLOCA NA MEMÓRIA APENAS O LOTE
            # 2. FECHA AUTOMATICAMENTE após o término do read_excel
            df_lote = pd.read_excel(
                arquivo_entrada, 
                skiprows=pulo_linhas + 1, # +1 para pular o cabeçalho original
                nrows=tamanho_lote, 
                names=colunas,
                engine='openpyxl' # Use 'pyxlsb' se for .xlsb para ser ainda mais rápido
            )

            if df_lote.empty:
                continuar_processando = False
                break

            # 3. PROCESSA O LOTE (Arquivo de entrada já está fechado aqui)
            print(f"⚙️ Processando {len(df_lote)} linhas...")
            motor.processar_lote(df_lote)

            # 4. SALVA NO RELATÓRIO (Abre, escreve e fecha)
            motor.salvar_append(arquivo_saida, primeiro_lote)

            # 5. ATUALIZA ENDEREÇO DA PRÓXIMA LINHA E ESTADO
            pulo_linhas += len(df_lote)
            primeiro_lote = False

            # Se o lote veio menor que o tamanho solicitado, chegamos ao fim
            if len(df_lote) < tamanho_lote:
                continuar_processando = False

        print(f"\n✨ Concluído! Processamento finalizado com sucesso.")
    else:
        print("❌ Arquivo de entrada não encontrado.")

   


