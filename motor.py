import pandas as pd
import re
import os

# ==========================================
# 1. CLASSES DE FILTROS (REGRAS)
# ==========================================

class Filtro:
    """Classe base para todos os tipos de filtros."""
    def __init__(self, coluna_idx, nome):
        self.coluna_idx = coluna_idx
        self.nome = nome

    def aplicar(self, item, nome_coluna):
        """Deve ser sobrescrito pelas subclasses."""
        raise NotImplementedError("O método aplicar deve ser implementado.")

    def to_dict(self):
        """Transforma o objeto em dicionário para salvar em JSON."""
        raise NotImplementedError("O método to_dict deve ser implementado.")


class FiltroSimples(Filtro):
    """Implementa os tipos 0 (Inclusão) e 1 (Exclusão) via Regex."""
    def __init__(self, coluna_idx, tipo, palavras, nome):
        super().__init__(coluna_idx, nome)
        self.tipo = tipo  # 0: Incluir se existir, 1: Excluir se existir
        self.palavras = palavras

    def aplicar(self, item, nome_coluna):
        valor = str(item.get(nome_coluna, "")).lower()
        match_detectado = None
        
        for termo in self.palavras:
            termo_lower = termo.lower()
            # \b garante que a palavra seja exata (evita 'bomba' em 'bombardeio')
            padrao = r'\b' + re.escape(termo_lower) + r'\b'
            if re.search(padrao, valor):
                match_detectado = termo_lower
                break

        if self.tipo == 0:  # Tipo 0: INCLUIR
            if match_detectado:
                item['Termo_Match'] = match_detectado
                return item
        elif self.tipo == 1:  # Tipo 1: EXCLUIR
            if not match_detectado:
                return item
        return None

    def to_dict(self):
        return {
            "classe": "FiltroSimples",
            "coluna_idx": self.coluna_idx,
            "tipo": self.tipo,
            "palavras": self.palavras,
            "nome": self.nome
        }


class FiltroCondicional(Filtro):
    """Implementa os tipos 2 (Refinamento) e 3 (Exclusão Condicional)."""
    def __init__(self, coluna_idx, tipo, gatilhos, adicionais, nome):
        super().__init__(coluna_idx, nome)
        self.tipo = tipo  # 2: Mantém se tiver adicional, 3: Exclui se tiver adicional
        self.gatilhos = [g.lower() for g in gatilhos]
        self.adicionais = [a.lower() for a in adicionais]

    def aplicar(self, item, nome_coluna):
        valor = str(item.get(nome_coluna, "")).lower()
        tem_gatilho = any(g in valor for g in self.gatilhos)
        
        if tem_gatilho:
            tem_termo_ref = any(p in valor for p in self.adicionais)
            
            if self.tipo == 2:  # Mantém apenas se tiver o gatilho E o adicional
                return item if tem_termo_ref else None
            
            if self.tipo == 3:  # Exclui se tiver o gatilho E o adicional
                return None if tem_termo_ref else item
        
        return item  # Se não tem o gatilho, passa ileso

    def to_dict(self):
        return {
            "classe": "FiltroCondicional",
            "coluna_idx": self.coluna_idx,
            "tipo": self.tipo,
            "gatilhos": self.gatilhos,
            "adicionais": self.adicionais,
            "nome": self.nome
        }


# ==========================================
# 2. GERENCIADOR DE EXECUÇÃO
# ==========================================

class MotorFiltragem:
    """Motor responsável por processar uma lista de filtros em sequência."""
    def __init__(self, titulo_cenario):
        self.titulo_cenario = titulo_cenario
        self.filtros = []  # Lista de objetos FiltroSimples ou FiltroCondicional
        self.historico_abas = []

    def adicionar_filtro(self, filtro_obj):
        self.filtros.append(filtro_obj)

    def processar(self, df_inicial):
        """Executa a lógica de funil sequencial."""
        if df_inicial.empty:
            return None

        # Converte para lista de dicts para performance no loop
        dados_atuais = df_inicial.to_dict(orient='records')
        self.historico_abas = [("0_Base_Original", df_inicial.copy())]

        for filtro in self.filtros:
            if not dados_atuais:
                break
            
            # Descobre o nome da coluna pelo índice
            nome_coluna = list(dados_atuais[0].keys())[filtro.coluna_idx]
            novos_dados = []

            for item in dados_atuais:
                resultado = filtro.aplicar(item.copy(), nome_coluna)
                if resultado:
                    novos_dados.append(resultado)
            
            dados_atuais = novos_dados
            
            # Armazena o DataFrame resultante desta etapa
            if dados_atuais:
                self.historico_abas.append((filtro.nome, pd.DataFrame(dados_atuais)))
            else:
                break

        return dados_atuais

    def salvar_excel(self, caminho_saida):
        """Gera o Excel final com abas reversas (da última etapa para a primeira)."""
        if not self.historico_abas:
            print("❌ Nada para salvar.")
            return

        try:
            with pd.ExcelWriter(caminho_saida, engine='openpyxl') as writer:
                # Invertemos para que o resultado final seja a primeira aba do Excel
                for nome, df in reversed(self.historico_abas):
                    # Limpeza de colunas auxiliares
                    if 'Termo_Match' in df.columns and nome != self.filtros[0].nome:
                        if any(f.tipo == 0 for f in self.filtros): # se for do tipo inclusão mantém
                            pass
                        else:
                            df = df.drop(columns=['Termo_Match'])
                    
                    # Nome da aba limitado a 31 caracteres
                    nome_aba = re.sub(r'[\\/*?:\[\]]', '', nome)[:31]
                    df.to_excel(writer, sheet_name=nome_aba, index=False)
            return True
        except Exception as e:
            print(f"❌ Erro ao salvar Excel: {e}")
            return False

# ==========================================
# 3. UTILITÁRIO DE CARREGAMENTO (FÁBRICA)
# ==========================================

def criar_filtro_de_dict(d):
    """Reconstrói o objeto do filtro a partir de um dicionário (vinda do JSON)."""
    if d['classe'] == "FiltroSimples":
        return FiltroSimples(d['coluna_idx'], d['tipo'], d['palavras'], d['nome'])
    elif d['classe'] == "FiltroCondicional":
        return FiltroCondicional(d['coluna_idx'], d['tipo'], d['gatilhos'], d['adicionais'], d['nome'])
    return None