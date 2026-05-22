
# import pandas as pd
# import re
# import os
# from openpyxl import load_workbook
# from openpyxl.utils.dataframe import dataframe_to_rows

# # ==========================================
# # CONFIGURAÇÕES TÉCNICAS
# # ==========================================
# tamanho_lote = 20  # Processa 200 por vez para segurança total de RAM
# arquivo_entrada = 'dados.xlsm'
# arquivo_saida = 'RELATORIO.xlsx'

# # ==========================================
# # CLASSE BASE E FILTROS ESPECÍFICOS (SUA LÓGICA VALIDADA)
# # ==========================================

# class Filtro:
#     def __init__(self, coluna_idx, nome):
#         self.coluna_idx = coluna_idx
#         self.nome = nome

#     def aplicar(self, df):
#         raise NotImplementedError("Cada filtro deve implementar seu próprio método aplicar.")
        
#     def _obter_coluna(self, df):
#         return df.columns[self.coluna_idx]

# class FiltroSimples(Filtro):
#     def __init__(self, coluna_idx, tipo, palavras, nome):
#         super().__init__(coluna_idx, nome)
#         self.tipo = tipo
#         self.padrao_regex = r'\b(' + '|'.join(re.escape(p) for p in palavras) + r')\b'

#     def aplicar(self, df):
#         if df.empty: return df
#         nome_coluna = self._obter_coluna(df)
#         serie_texto = df[nome_coluna].fillna("").astype(str)

#         if self.tipo == 0: # Inclusão
#             extracted = serie_texto.str.extract(self.padrao_regex, flags=re.IGNORECASE, expand=False)
#             mask = extracted.notna()
#             df_filtrado = df[mask].copy()
#             df_filtrado['Termo_Match'] = extracted[mask].values
#             return df_filtrado

#         elif self.tipo == 1: # Exclusão
#             mask = serie_texto.str.contains(self.padrao_regex, flags=re.IGNORECASE, na=False, regex=True)
#             return df[~mask].copy()

# class FiltroCondicional(Filtro):
#     def __init__(self, coluna_idx, tipo, gatilhos, adicionais, nome):
#         super().__init__(coluna_idx, nome)
#         self.tipo = tipo 
#         self.regex_gatilhos = '|'.join(re.escape(g) for g in gatilhos)
#         self.regex_adicionais = '|'.join(re.escape(a) for a in adicionais)

#     def aplicar(self, df):
#         if df.empty: return df
#         nome_coluna = self._obter_coluna(df)
#         serie_texto = df[nome_coluna].fillna("").astype(str)
        
#         mask_gatilho = serie_texto.str.contains(self.regex_gatilhos, flags=re.IGNORECASE, na=False, regex=True)
#         mask_adicional = serie_texto.str.contains(self.regex_adicionais, flags=re.IGNORECASE, na=False, regex=True)
        
#         if self.tipo == 2:
#             mask_final = ~mask_gatilho | mask_adicional
#         elif self.tipo == 3:
#             mask_final = ~(mask_gatilho & mask_adicional)
            
#         return df[mask_final].copy()

# # ==========================================
# # MOTOR DE FILTRAGEM (ADAPTADO PARA PONTEIRO)
# # ==========================================

# class MotorFiltragem:
#     def __init__(self, titulo):
#         self.titulo = titulo
#         self.filtros = []
#         self.historico_abas = []

#     def adicionar_filtro(self, filtro_obj):
#         self.filtros.append(filtro_obj)

#     def processar_lote(self, df_chunk):
#         """Processa o lote e guarda o estado de cada etapa na memória temporária."""
#         df_atual = df_chunk.copy()
#         # Nota: Base_Original recebe o lote atual
#         self.historico_abas = [("0_Base_Original", df_atual.copy())]

#         for filtro in self.filtros:
#             df_atual = filtro.aplicar(df_atual)
#             self.historico_abas.append((filtro.nome, df_atual.copy() if df_atual is not None else pd.DataFrame()))
            
#             if df_atual is None or df_atual.empty:
#                 break

#     def salvar_append(self, caminho_saida, is_first_chunk):
#         """Escreve os resultados no Excel de forma incremental."""
#         if not self.historico_abas: return
        
#         if is_first_chunk:
#             # Cria o arquivo novo na primeira iteração
#             with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
#                 for nome, df in reversed(self.historico_abas):
#                     if 'Termo_Match' in df.columns and nome == "0_Base_Original":
#                          df = df.drop(columns=['Termo_Match'], errors='ignore')
#                     df.to_excel(writer, sheet_name=nome[:31], index=False)
#         else:
#             # Anexa o conteúdo nos arquivos existentes
#             wb = load_workbook(caminho_saida)
#             for nome, df in reversed(self.historico_abas):
#                 if df.empty: continue
                
#                 if 'Termo_Match' in df.columns and nome == "0_Base_Original":
#                      df = df.drop(columns=['Termo_Match'], errors='ignore')
                
#                 nome_aba = nome[:31]
#                 if nome_aba not in wb.sheetnames:
#                     wb.create_sheet(nome_aba)
                
#                 ws = wb[nome_aba]
#                 for r in dataframe_to_rows(df, index=False, header=False):
#                     ws.append(r)
#             wb.save(caminho_saida)
#             wb.close()
            
#         # Limpa o lote da memória RAM após salvar
#         self.historico_abas = []

# # ==========================================
# # EXECUÇÃO COM LÓGICA DE PONTEIRO
# # ==========================================

# if __name__ == "__main__":
#     motor = MotorFiltragem("Relatorio de Ativos")

#     # Configuração dos seus filtros validados
#     motor.adicionar_filtro(FiltroSimples(8, 0, ['BOOSTER','BOOSTERs','SBL','BOMBEAMENTO','BOMBA','bombas','reservatório', 'reservatorio','tanque','estação','estacao','ETE','ETA','EEE','abraçadeira','abracadeira','acoplamento', ], "Filtro geral"))
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['bomba','BOMBEAMENTO'], ['manutenção','manutencao','calor','concreto','combustível','dosadoras','dosador','SUCÇÃO','banheiro','banheiros',], 'filtro bomba'))
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['estação','estacao','ETE','ETA','EEE'], ['prfv','Remoção','limpeza','manutenção','manutencao','Rádio','REFORMA','laboratoriais','laboratorial',"laboratório",'laboratorio','GÁS','AQUECEDOR'], 'filtro estação'))
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['tanque','reservatório','reservatorio'], ["CAIXA D'ÁGUA","barril","MEDICINAIS",'medicina','Nitrogênio','agitadores','ESPARGIDOR'], 'filtro reservatório'))    
    
#     motor.adicionar_filtro(FiltroSimples(8, 1, ['EPIS','proteção','incendio','incendios','ferramentas','salas','CARRETA','caminhão','caminhões','ares-condicionados','ar-condicionado','pipa'], "filtro de lixo"))
#     motor.adicionar_filtro(FiltroSimples(8, 1, ['manutenção', 'limpeza'], "Filtro manutenção"))
#     motor.adicionar_filtro(FiltroSimples(5, 1, ['urgente'], "Filtro urgente"))

#     if os.path.exists(arquivo_entrada):
#         # 1. Captura as colunas antes de começar
#         df_headers = pd.read_excel(arquivo_entrada, sheet_name=-1, nrows=0)
#         colunas_originais = df_headers.columns.tolist()
        
#         ponteiro = 0
#         primeiro_lote = True

#         print(f"🔄 Iniciando processamento dinâmico em lotes de {tamanho_lote}...")

#         while True:
#             # 2. Carrega apenas o pedaço necessário (pula o que já foi processado)
#             try:
#                 df_chunk = pd.read_excel(
#                     arquivo_entrada, 
#                     sheet_name=-1, 
#                     skiprows=ponteiro + 1, 
#                     nrows=tamanho_lote,
#                     names=colunas_originais
#                 )
#             except Exception:
#                 break

#             if df_chunk.empty:
#                 break

#             print(f"⚙️ Processando ponteiro: linha {ponteiro} até {ponteiro + len(df_chunk)}...")
            
#             # 3. Processa e Salva
#             motor.processar_lote(df_chunk)
#             motor.salvar_append(arquivo_saida, primeiro_lote)
            
#             # 4. Atualiza estado
#             ponteiro += len(df_chunk)
#             primeiro_lote = False

#         print(f"\n✨ Relatório '{motor.titulo}' gerado com sucesso!")
#     else:
#         print(f"❌ Erro: Arquivo '{arquivo_entrada}' não encontrado.")























# import pandas as pd
# import re
# import os
# from openpyxl import load_workbook
# from openpyxl.utils.dataframe import dataframe_to_rows

# # TAMANHO DO LOTE AQUI
# tamanho_lote = 200 

# # ==========================================
# # CLASSE BASE E FILTROS ESPECÍFICOS (VETORIZADOS)
# # ==========================================

# class Filtro:
#     """Classe base para todos os filtros."""
#     def __init__(self, coluna_idx, nome):
#         self.coluna_idx = coluna_idx
#         self.nome = nome

#     def aplicar(self, df):
#         raise NotImplementedError("Cada filtro deve implementar seu próprio método aplicar.")
        
#     def _obter_coluna(self, df):
#         """Retorna o nome da coluna baseada no índice fornecido."""
#         return df.columns[self.coluna_idx]

# class FiltroSimples(Filtro):
#     """Tipos 0 (Inclusão) e 1 (Exclusão) usando Regex Vetorizado."""
#     def __init__(self, coluna_idx, tipo, palavras, nome):
#         super().__init__(coluna_idx, nome)
#         self.tipo = tipo
#         self.padrao_regex = r'\b(' + '|'.join(re.escape(p) for p in palavras) + r')\b'

#     def aplicar(self, df):
#         if df.empty: return df
#         nome_coluna = self._obter_coluna(df)
#         serie_texto = df[nome_coluna].fillna("").astype(str)

#         if self.tipo == 0: # Inclusão
#             extracted = serie_texto.str.extract(self.padrao_regex, flags=re.IGNORECASE, expand=False)
#             mask = extracted.notna()
#             df_filtrado = df[mask].copy()
#             df_filtrado['Termo_Match'] = extracted[mask].values
#             return df_filtrado

#         elif self.tipo == 1: # Exclusão
#             mask = serie_texto.str.contains(self.padrao_regex, flags=re.IGNORECASE, na=False, regex=True)
#             return df[~mask].copy()

# class FiltroCondicional(Filtro):
#     """Tipos 2 (Manutenção) e 3 (Exclusão Condicional) usando Máscaras Booleanas."""
#     def __init__(self, coluna_idx, tipo, gatilhos, adicionais, nome):
#         super().__init__(coluna_idx, nome)
#         self.tipo = tipo 
#         self.regex_gatilhos = '|'.join(re.escape(g) for g in gatilhos)
#         self.regex_adicionais = '|'.join(re.escape(a) for a in adicionais)

#     def aplicar(self, df):
#         if df.empty: return df
#         nome_coluna = self._obter_coluna(df)
#         serie_texto = df[nome_coluna].fillna("").astype(str)
        
#         mask_gatilho = serie_texto.str.contains(self.regex_gatilhos, flags=re.IGNORECASE, na=False, regex=True)
#         mask_adicional = serie_texto.str.contains(self.regex_adicionais, flags=re.IGNORECASE, na=False, regex=True)
        
#         if self.tipo == 2:
#             mask_final = ~mask_gatilho | mask_adicional
#         elif self.tipo == 3:
#             mask_final = ~(mask_gatilho & mask_adicional)
            
#         return df[mask_final].copy()

# # ==========================================
# # GERENCIADOR DO MOTOR POR LOTES (CHUNKS)
# # ==========================================

# class MotorFiltragem:
#     def __init__(self, titulo):
#         self.titulo = titulo
#         self.filtros = []
#         self.historico_abas = []

#     def adicionar_filtro(self, filtro_obj):
#         self.filtros.append(filtro_obj)
#         print(f"➕ Filtro '{filtro_obj.nome}' adicionado.")

#     def processar_lote(self, df_chunk):
#         """Processa apenas o lote atual de dados."""
#         df_atual = df_chunk.copy()
#         self.historico_abas = [("0_Base_Original", df_atual.copy())]

#         for filtro in self.filtros:
#             df_atual = filtro.aplicar(df_atual)
#             self.historico_abas.append((filtro.nome, df_atual.copy()))
            
#             if df_atual is None or df_atual.empty:
#                 # Se o funil zerou neste lote, não precisamos continuar os próximos filtros para este lote
#                 break

#     def salvar_append(self, caminho_saida, is_first_chunk):
#         """Salva o lote no Excel. Se for o primeiro, cria o arquivo. Senão, faz append nas abas."""
#         if not self.historico_abas: return
        
#         if is_first_chunk:
#             # Cria o arquivo do zero
#             with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
#                 for nome, df in reversed(self.historico_abas):
#                     if 'Termo_Match' in df.columns and nome == "0_Base_Original":
#                          df = df.drop(columns=['Termo_Match'], errors='ignore')
                    
#                     df.to_excel(writer, sheet_name=nome[:31], index=False)
#         else:
#             # Abre o arquivo existente e adiciona as linhas no final
#             wb = load_workbook(caminho_saida)
#             for nome, df in reversed(self.historico_abas):
#                 if 'Termo_Match' in df.columns and nome == "0_Base_Original":
#                      df = df.drop(columns=['Termo_Match'], errors='ignore')
                
#                 nome_aba = nome[:31]
#                 if nome_aba in wb.sheetnames:
#                     ws = wb[nome_aba]
#                     # Adiciona as linhas sem o cabeçalho (pois já existe)
#                     for r in dataframe_to_rows(df, index=False, header=False):
#                         ws.append(r)
#             wb.save(caminho_saida)
#             wb.close()
            
#         # Limpa o histórico deste lote para liberar a memória RAM
#         self.historico_abas = []

# # ==========================================
# # EXECUÇÃO 
# # ==========================================

# if __name__ == "__main__":
#     motor = MotorFiltragem("Relatorio de Ativos")

#     motor.adicionar_filtro(FiltroSimples(8, 0, ['BOOSTER','BOOSTERs','SBL','BOMBEAMENTO','BOMBA','bombas','reservatório', 'reservatorio','tanque','estação','estacao','ETE','ETA','EEE','abraçadeira','abracadeira','acoplamento', ], "Filtro geral"))
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['bomba','BOMBEAMENTO'], ['manutenção','manutencao','calor','concreto','combustível','dosadoras','dosador','SUCÇÃO','banheiro','banheiros',], 'filtro bomba'))
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['estação','estacao','ETE','ETA','EEE'], ['prfv','Remoção','limpeza','manutenção','manutencao','Rádio','REFORMA','laboratoriais','laboratorial',"laboratório",'laboratorio','GÁS','AQUECEDOR'], 'filtro estação'))
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['tanque','reservatório','reservatorio'], ["CAIXA D'ÁGUA","barril","MEDICINAIS",'medicina','Nitrogênio','agitadores','ESPARGIDOR'], 'filtro reservatório'))    
#     motor.adicionar_filtro(FiltroSimples(8, 1, ['EPIS','proteção','incendio','incendios','ferramentas','salas','CARRETA','caminhão','caminhões','ares-condicionados','ar-condicionado','pipa'], "filtro de lixo"))
#     motor.adicionar_filtro(FiltroSimples(8, 1, ['manutenção', 'limpeza'], "Filtro manutenção"))
#     motor.adicionar_filtro(FiltroSimples(5, 1, ['urgente'], "Filtro urgente"))

#     arquivo_entrada = 'dados.xlsm'
#     arquivo_saida = 'RELATORIO.xlsx'

#     if os.path.exists(arquivo_entrada):
#         print("\n⏳ Carregando a base de dados original na memória...")
#         df_total = pd.read_excel(arquivo_entrada, sheet_name=-1)
#         total_linhas = len(df_total)
        
        
        
#         print(f"📊 Total de {total_linhas} linhas encontradas.")
#         print(f"🔄 Iniciando processamento em blocos de {tamanho_lote} linhas...\n")

#         primeiro_lote = True
        
#         # Loop para processar e salvar por lotes
#         for inicio in range(0, total_linhas, tamanho_lote):
#             fim = inicio + tamanho_lote
#             df_chunk = df_total.iloc[inicio:fim]
            
#             print(f"⚙️ Processando linhas {inicio} a {min(fim, total_linhas)}...")
            
#             motor.processar_lote(df_chunk)
#             motor.salvar_append(arquivo_saida, primeiro_lote)
            
#             primeiro_lote = False
            
#         print(f"\n✨ Relatório '{motor.titulo}' gerado com sucesso!")
#     else:
#         print(f"❌ Arquivo '{arquivo_entrada}' não encontrado.")

























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

# # class FiltroSimples(Filtro):
# #     """Tipos 0 (Inclusão) e 1 (Exclusão) usando Regex."""
# #     def __init__(self, coluna_idx, tipo, palavras, nome):
# #         super().__init__(coluna_idx, nome)
# #         self.tipo = tipo  # 0 ou 1
# #         self.palavras = palavras

# #     def aplicar(self, item, nome_coluna):
# #         valor = str(item.get(nome_coluna, "")).lower()
# #         match_detectado = None
        
# #         for termo in self.palavras:
# #             padrao = r'\b' + re.escape(termo.lower()) + r'\b'
# #             if re.search(padrao, valor):
# #                 match_detectado = termo.lower()
# #                 break
        
# #         if self.tipo == 0:  # Inclusão
# #             if match_detectado:
# #                 item['Termo_Match'] = match_detectado
# #                 return item
# #         elif self.tipo == 1:  # Exclusão
# #             if not match_detectado:
# #                 return item
# #         return None



# class FiltroSimples(Filtro):
#     def __init__(self, coluna_idx, tipo, palavras, nome):
#         super().__init__(coluna_idx, nome)
#         self.tipo = tipo
#         self.palavras = palavras
#         # Pré-compila a regex uma única vez ao criar o filtro
#         padrao_unido = r'\b(' + '|'.join(re.escape(p.lower()) for p in palavras) + r')\b'
#         self.regex_compilada = re.compile(padrao_unido)

#     def aplicar(self, item, nome_coluna):
#         valor = str(item.get(nome_coluna, "")).lower()
#         match = self.regex_compilada.search(valor) # Muito mais rápido
        
#         if self.tipo == 0: # Inclusão
#             if match:
#                 item['Termo_Match'] = match.group(0)
#                 return item
#         elif self.tipo == 1: # Exclusão
#             if not match:
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
    
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['bomba','BOMBEAMENTO'], ['manutenção','manutencao','calor','concreto','combustível','dosadoras','dosador','SUCÇÃO','banheiro','banheiros',], 'filtro bomba'))
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['estação','estacao','ETE','ETA','EEE'], ['prfv','Remoção','limpeza','manutenção','manutencao','Rádio','REFORMA','laboratoriais','laboratorial',"laboratório",'laboratorio','GÁS','AQUECEDOR'], 'filtro estação'))
#     motor.adicionar_filtro(FiltroCondicional(8, 3, ['tanque','reservatório','reservatorio'], ["CAIXA D'ÁGUA","barril","MEDICINAIS",'medicina','Nitrogênio','agitadores','ESPARGIDOR'], 'filtro reservatório'))    
    
#     motor.adicionar_filtro(FiltroSimples(8, 1, ['EPIS','proteção','incendio','incendios','ferramentas','salas','CARRETA','caminhão','caminhões','ares-condicionados','ar-condicionado','pipa'], "filtro de lixo"))
#     motor.adicionar_filtro(FiltroSimples(8, 1, ['manutenção', 'limpeza'], "Filtro manutenção"))
#     motor.adicionar_filtro(FiltroSimples(5, 1, ['urgente'], "Filtro urgente"))


#     # 3. Executar (supondo que o arquivo exista)
#     if os.path.exists('dados.xlsm'):
#         df = pd.read_excel('dados.xlsm', sheet_name=-1)
#         motor.processar(df)
#         motor.salvar('RELATORIO.xlsx')
