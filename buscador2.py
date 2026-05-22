
# Motor de filtragem dinâmico.

#      Estruturas dos filtros
#      Estrutura 1 SIMPLES -> [coluna, tipo, ['palavra1','palavra2'], 'nome do filtro'],
#      Estrutura 2 COMPOSTO -> [coluna, tipo, ['gatilho1','gatilho2'], ['palavra1','palavra2'], 'nome do filtro'],


#     0 - Incluir se palavra exata existir.
#     1 - Excluir se palavra exata existir.                                     -->   [coluna, 1, ['EPIS','proteção'], "lixo"],
#     2 - Refinamento Condicional ([Gatilho] -> MANTÉM se tiver a [palavra]).
#     3 - Exclusão Condicional ([Gatilho] -> EXCLUI se tiver a [palavra]).      -->   [coluna, 3, ['bomba','BOMBEAMENTO'], ['manutenção','calor'], 'filtro bomba']
 


import pandas as pd
import re
import os

# ==========================================
# CLASSE BASE E FILTROS ESPECÍFICOS
# ==========================================

class Filtro:
    """Classe base para todos os filtros."""
    def __init__(self, coluna_idx, nome):
        self.coluna_idx = coluna_idx
        self.nome = nome

    def aplicar(self, dados):
        raise NotImplementedError("Cada filtro deve implementar seu próprio método aplicar.")

class FiltroSimples(Filtro):
    """Tipos 0 (Inclusão) e 1 (Exclusão) SEM o bloqueio de palavras exatas."""
    def __init__(self, coluna_idx, tipo, palavras, nome):
        super().__init__(coluna_idx, nome)
        self.tipo = tipo  # 0 ou 1
        self.palavras = palavras

    def aplicar(self, item, nome_coluna):
        valor = str(item.get(nome_coluna, "")).lower()
        match_detectado = None
        
        for termo in self.palavras:
            # REMOVIDO O \b para permitir que "BOMBA" encontre "BOMBAS" ou "BOMBEAMENTO"
            padrao = re.escape(termo.lower()) 
            if re.search(padrao, valor):
                match_detectado = termo.lower()
                break
        
        if self.tipo == 0:  # Inclusão
            if match_detectado:
                item['Termo_Match'] = match_detectado
                return item
        elif self.tipo == 1:  # Exclusão
            if not match_detectado:
                return item
        return None

class FiltroCondicional(Filtro):
    """Tipos 2 (Manutenção) e 3 (Exclusão Condicional)."""
    def __init__(self, coluna_idx, tipo, gatilhos, adicionais, nome):
        super().__init__(coluna_idx, nome)
        self.tipo = tipo  # 2 ou 3
        self.gatilhos = [g.lower() for g in gatilhos]
        self.adicionais = [a.lower() for a in adicionais]

    def aplicar(self, item, nome_coluna):
        valor = str(item.get(nome_coluna, "")).lower()
        tem_gatilho = any(g in valor for g in self.gatilhos)
        
        if tem_gatilho:
            tem_termo_ref = any(p in valor for p in self.adicionais)
            if self.tipo == 2 and tem_termo_ref: return item
            if self.tipo == 3 and not tem_termo_ref: return item
            return None
        return item  

# ==========================================
# GERENCIADOR DO MOTOR
# ==========================================

class MotorFiltragem:
    def __init__(self, titulo):
        self.titulo = titulo
        self.filtros = []
        self.historico_abas = []

    def adicionar_filtro(self, filtro_obj):
        self.filtros.append(filtro_obj)

    def processar(self, df_inicial):
        if df_inicial.empty:
            print("❌ O DataFrame de entrada está vazio. A aba do Excel pode estar em branco.")
            return

        dados_atuais = df_inicial.to_dict(orient='records')
        self.historico_abas = [("0_Base_Original", df_inicial.copy())]
        
        print(f"\n🚀 Iniciando processamento de {len(dados_atuais)} linhas...")

        for filtro in self.filtros:
            colunas_disponiveis = list(dados_atuais[0].keys())
            
            if filtro.coluna_idx >= len(colunas_disponiveis):
                print(f"⚠️ Erro Crítico no Filtro '{filtro.nome}': A coluna de índice {filtro.coluna_idx} não existe.")
                print(f"   A planilha tem as seguintes {len(colunas_disponiveis)} colunas: {colunas_disponiveis}")
                return

            nome_coluna = colunas_disponiveis[filtro.coluna_idx]
            novos_dados = []

            for item in dados_atuais:
                resultado = filtro.aplicar(item.copy(), nome_coluna)
                if resultado is not None:
                    novos_dados.append(resultado)
            
            dados_atuais = novos_dados
            qtd_restante = len(dados_atuais)
            
            if qtd_restante == 0:
                print(f"🛑 O funil PAROU no filtro '{filtro.nome}' (Analisando a coluna: '{nome_coluna}').")
                print("   Motivo: Nenhuma linha sobrou após esta filtragem.")
                break
            
            self.historico_abas.append((filtro.nome, pd.DataFrame(dados_atuais)))
            print(f"✅ Filtro '{filtro.nome}' aplicou com sucesso: {qtd_restante} itens restantes.")

        print(f"\n🏁 Fim do processamento. Linhas prontas para salvar: {len(dados_atuais)}")

    def salvar(self, caminho_saida):
        if not self.historico_abas: return
        with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
            for nome, df in reversed(self.historico_abas):
                if 'Termo_Match' in df.columns and nome == "0_Base_Original":
                     df = df.drop(columns=['Termo_Match'])
                df.to_excel(writer, sheet_name=nome[:31], index=False)
        print(f"✨ Arquivo '{caminho_saida}' gerado com sucesso!")

# ==========================================
# EXECUÇÃO E DIAGNÓSTICO
# ==========================================

if __name__ == "__main__":
    nome_arquivo = 'dados.xlsm'
    
    print("="*40)
    print("🔍 INICIANDO DIAGNÓSTICO DO SISTEMA")
    print("="*40)
    print(f"Pasta atual onde o Python está rodando: {os.getcwd()}")
    
    if os.path.exists(nome_arquivo):
        print(f"✅ SUCESSO: O arquivo '{nome_arquivo}' foi encontrado nesta pasta!")
        try:
            # Lendo a primeira aba com engine openpyxl para evitar problemas com .xlsm
            print("Lendo a planilha do Excel... (isso pode demorar alguns segundos se for muito grande)")
            df = pd.read_excel(nome_arquivo, engine='openpyxl')
            
            motor = MotorFiltragem("Relatorio de Ativos")
            
            # ADICIONANDO SEUS FILTROS (Mudei todos para a coluna 8 conforme você pediu)
            # motor.adicionar_filtro(FiltroSimples(8, 0, ['BOOSTER','SBL','BOMBEAMENTO','BOMBA','reservatório', 'reservatorio','tanque','estação','estacao','ETE','ETA','EEE','abraçadeira','abracadeira','acoplamento', ], "Filtro geral"))
            # motor.adicionar_filtro(FiltroCondicional(8, 3, ['bomba','BOMBEAMENTO'], ['manutenção','calor','concreto','combustível','dosadoras','dosador','SUCÇÃO','banheiro','banheiros',], 'filtro bomba'))
            # motor.adicionar_filtro(FiltroCondicional(8, 3, ['estação','estacao','ETE','ETA','EEE'], ['prfv','Remoção','limpeza','manutenção','Rádio','REFORMA','laboratoriais','laboratorial',"laboratório",'laboratorio','GÁS',], 'filtro estação'))
            # motor.adicionar_filtro(FiltroCondicional(8, 3, ['tanque'], ["CAIXA D'ÁGUA"], 'filtro reservatório'))    
            # motor.adicionar_filtro(FiltroSimples(8, 1, ['EPIS','proteção','incendio','incendios','ferramentas','salas','CARRETA','caminhão','caminhões','ares-condicionados','ar-condicionado','pipa'], "filtro de lixo"))
            # motor.adicionar_filtro(FiltroSimples(8, 1, ['manutenção', 'limpeza'], "Filtro manutenção"))
            
            # Aqui no seu código original estava coluna 5. Mantive 5, se for 8 mude aqui:
            # motor.adicionar_filtro(FiltroSimples(5, 1, ['urgente'], "Filtro urgente"))

                # 2. Adicionar os filtros usando a nova estrutura
            motor.adicionar_filtro(FiltroSimples(8, 0, ['BOOSTER','BOOSTERs','SBL','BOMBEAMENTO','BOMBA','bombas','reservatório', 'reservatorio','tanque','estação','estacao','ETE','ETA','EEE','abraçadeira','abracadeira','acoplamento', ], "Filtro geral"))
            motor.adicionar_filtro(FiltroCondicional(8, 3, ['bomba','BOMBEAMENTO'], ['manutenção','calor','concreto','combustível','dosadoras','dosador','SUCÇÃO','banheiro','banheiros',], 'filtro bomba'))
            motor.adicionar_filtro(FiltroCondicional(8, 3, ['estação','estacao','ETE','ETA','EEE'], ['prfv','Remoção','limpeza','manutenção','Rádio','REFORMA','laboratoriais','laboratorial',"laboratório",'laboratorio','GÁS',], 'filtro estação'))
            motor.adicionar_filtro(FiltroCondicional(8, 3, ['tanque'], ["CAIXA D'ÁGUA"], 'filtro reservatório'))    
            
            motor.adicionar_filtro(FiltroSimples(8, 1, ['EPIS','proteção','incendio','incendios','ferramentas','salas','CARRETA','caminhão','caminhões','ares-condicionados','ar-condicionado','pipa'], "filtro de lixo"))
            motor.adicionar_filtro(FiltroSimples(8, 1, ['manutenção', 'limpeza'], "Filtro manutenção"))
            motor.adicionar_filtro(FiltroSimples(5, 1, ['urgente'], "Filtro urgente"))



            motor.processar(df)
            
            if len(motor.historico_abas) > 1: # Verifica se passou do original
                motor.salvar('relatorio.xlsx')
            else:
                print("⚠️ O relatório final NÃO foi salvo porque não restaram dados após os filtros.")
                
        except Exception as e:
            print(f"❌ ERRO AO LER O ARQUIVO EXCEL: {e}")
    else:
        print(f"❌ ERRO CRÍTICO: O arquivo '{nome_arquivo}' NÃO foi encontrado!")
        print("   Verifique se o nome está correto e se ele está EXATAMENTE na pasta mostrada acima.")
















# import pandas as pd
# import re
# import os

# # ==========================================
# # CLASSE BASE E FILTROS ESPECÍFICOS
# # ==========================================

# class Filtro:
#     """Classe base para todos os filtros."""
#     def __init__(self, coluna_idx, nome):
#         self.coluna_idx = coluna_idx
#         self.nome = nome

#     def aplicar(self, dados):
#         raise NotImplementedError("Cada filtro deve implementar seu próprio método aplicar.")

# class FiltroSimples(Filtro):
#     """Tipos 0 (Inclusão) e 1 (Exclusão) usando Regex."""
#     def __init__(self, coluna_idx, tipo, palavras, nome):
#         super().__init__(coluna_idx, nome)
#         self.tipo = tipo  # 0 ou 1
#         self.palavras = palavras

#     def aplicar(self, item, nome_coluna):
#         valor = str(item.get(nome_coluna, "")).lower()
#         match_detectado = None
        
#         for termo in self.palavras:
#             padrao = r'\b' + re.escape(termo.lower()) + r'\b'
#             if re.search(padrao, valor):
#                 match_detectado = termo.lower()
#                 break
        
#         if self.tipo == 0:  # Inclusão
#             if match_detectado:
#                 item['Termo_Match'] = match_detectado
#                 return item
#         elif self.tipo == 1:  # Exclusão
#             if not match_detectado:
#                 return item
#         return None

# class FiltroCondicional(Filtro):
#     """Tipos 2 (Manutenção) e 3 (Exclusão Condicional)."""
#     def __init__(self, coluna_idx, tipo, gatilhos, adicionais, nome):
#         super().__init__(coluna_idx, nome)
#         self.tipo = tipo  # 2 ou 3
#         self.gatilhos = [g.lower() for g in gatilhos]
#         self.adicionais = [a.lower() for a in adicionais]

#     def aplicar(self, item, nome_coluna):
#         valor = str(item.get(nome_coluna, "")).lower()
#         tem_gatilho = any(g in valor for g in self.gatilhos)
        
#         if tem_gatilho:
#             tem_termo_ref = any(p in valor for p in self.adicionais)
#             if self.tipo == 2 and tem_termo_ref: return item
#             if self.tipo == 3 and not tem_termo_ref: return item
#             return None
#         return item  # Se não tem gatilho, passa direto

# # ==========================================
# # GERENCIADOR DO MOTOR
# # ==========================================

# class MotorFiltragem:
#     def __init__(self, titulo):
#         self.titulo = titulo
#         self.filtros = []
#         self.historico_abas = []

#     def adicionar_filtro(self, filtro_obj):
#         self.filtros.append(filtro_obj)
#         print(f"➕ Filtro '{filtro_obj.nome}' adicionado.")

#     def remover_filtro(self, nome_filtro):
#         self.filtros = [f for f in self.filtros if f.nome != nome_filtro]
#         print(f"🗑️ Filtro '{nome_filtro}' removido.")

#     def processar(self, df_inicial):
#         dados_atuais = df_inicial.to_dict(orient='records')
#         self.historico_abas = [("0_Base_Original", df_inicial.copy())]

#         for filtro in self.filtros:
#             nome_coluna = list(dados_atuais[0].keys())[filtro.coluna_idx]
#             novos_dados = []

#             for item in dados_atuais:
#                 resultado = filtro.aplicar(item.copy(), nome_coluna)
#                 if resultado:
#                     novos_dados.append(resultado)
            
#             dados_atuais = novos_dados
#             if not dados_atuais:
#                 print(f"⚠️ O funil parou em '{filtro.nome}': 0 itens restantes.")
#                 break
            
#             self.historico_abas.append((filtro.nome, pd.DataFrame(dados_atuais)))
#             print(f"✅ Etapa '{filtro.nome}' concluída ({len(dados_atuais)} itens).")


#     def salvar(self, caminho_saida):
#         if not self.historico_abas: return
        
#         with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
#             for nome, df in reversed(self.historico_abas):
#                 # Limpa coluna de match se não for a aba de inclusão
#                 if 'Termo_Match' in df.columns and nome == "0_Base_Original":
#                      df = df.drop(columns=['Termo_Match'])
                
#                 df.to_excel(writer, sheet_name=nome[:31], index=False)
#         print(f"\n✨ Relatório '{self.titulo}' gerado com sucesso!")

# # ==========================================
# # EXECUÇÃO 
# # ==========================================

# if __name__ == "__main__":
#     # 1. Instanciar o Motor
#     motor = MotorFiltragem("Relatorio de Ativos")

#     # 2. Adicionar os filtros usando a nova estrutura
#     motor.adicionar_filtro(FiltroSimples(8, 0, ['BOOSTER','BOOSTERs','SBL','BOMBEAMENTO','BOMBA','bombas','reservatório', 'reservatorio','tanque','estação','estacao','ETE','ETA','EEE','abraçadeira','abracadeira','acoplamento', ], "Filtro geral"))
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['bomba','BOMBEAMENTO'], ['manutenção','calor','concreto','combustível','dosadoras','dosador','SUCÇÃO','banheiro','banheiros',], 'filtro bomba'))
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['estação','estacao','ETE','ETA','EEE'], ['prfv','Remoção','limpeza','manutenção','Rádio','REFORMA','laboratoriais','laboratorial',"laboratório",'laboratorio','GÁS',], 'filtro estação'))
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['tanque'], ["CAIXA D'ÁGUA"], 'filtro reservatório'))    
    
#     motor.adicionar_filtro(FiltroSimples(8, 1, ['EPIS','proteção','incendio','incendios','ferramentas','salas','CARRETA','caminhão','caminhões','ares-condicionados','ar-condicionado','pipa'], "filtro de lixo"))
#     motor.adicionar_filtro(FiltroSimples(8, 1, ['manutenção', 'limpeza'], "Filtro manutenção"))
#     motor.adicionar_filtro(FiltroSimples(5, 1, ['urgente'], "Filtro urgente"))


#     # 3. Executar (supondo que o arquivo exista)
#     if os.path.exists('dados.xlsm'):
#         df = pd.read_excel('dados.xlsm', sheet_name=-1)
#         motor.processar(df)
#         motor.salvar('relatorio.xlsx')

