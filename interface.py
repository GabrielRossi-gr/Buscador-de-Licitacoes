import customtkinter as ctk
import json
import os
from componentes import JanelaAdicionar 

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class InterfaceConfig(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Gerenciador de Filtros Pro")
        self.geometry("600x700")
        self.arquivo_db = "configuracoes.json"
        
        # Agora lista_configs será uma lista de DICTs: 
        # [{"titulo": "Cenário 1", "filtros": [...]}, ...]
        self.lista_configs = []

        self.setup_ui()
        self.carregar_dados()

    def setup_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Cabeçalho
        self.frame_topo = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_topo.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        self.frame_topo.grid_columnconfigure(0, weight=1)

        self.label_titulo = ctk.CTkLabel(self.frame_topo, text="Meus Cenários", font=ctk.CTkFont(size=22, weight="bold"))
        self.label_titulo.grid(row=0, column=0, sticky="w")

        self.btn_abrir_popup = ctk.CTkButton(self.frame_topo, text="+ Novo Cenário", command=self.abrir_janela_add)
        self.btn_abrir_popup.grid(row=0, column=1)

        # Lista com Scroll
        self.scroll_lista = ctk.CTkScrollableFrame(self, label_text="Configurações Salvas")
        self.scroll_lista.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.scroll_lista.grid_columnconfigure(0, weight=1)

    def abrir_janela_add(self):
        JanelaAdicionar(self, self.processar_adicao, titulo_janela="Nome do Cenário")

    def processar_adicao(self, texto, janela_popup):
        texto = texto.strip()
        if texto:
            # Criamos a estrutura do objeto de configuração
            nova_config = {
                "titulo": texto,
                "filtros": [] # Aqui ficarão os filtros (Tipo 0, 1, 2, 3)
            }
            self.lista_configs.append(nova_config)
            self.salvar_dados()
            self.atualizar_interface_lista()
            janela_popup.destroy()

    def abrir_configuracao(self, config_obj):
        """Função que será chamada ao clicar em uma configuração"""
        print(f"Abrindo filtros de: {config_obj['titulo']}")
        # Aqui você abriria uma nova Janela ou mudaria o Frame para exibir os filtros
        # Exemplo: JanelaFiltrosInternos(self, config_obj)

    def carregar_dados(self):
        if os.path.exists(self.arquivo_db):
            try:
                with open(self.arquivo_db, "r", encoding="utf-8") as f:
                    self.lista_configs = json.load(f)
            except: self.lista_configs = []
        self.atualizar_interface_lista()

    def salvar_dados(self):
        with open(self.arquivo_db, "w", encoding="utf-8") as f:
            json.dump(self.lista_configs, f, indent=4)

    def excluir_item(self, config_obj):
        self.lista_configs.remove(config_obj)
        self.salvar_dados()
        self.atualizar_interface_lista()

    def atualizar_interface_lista(self):
        for widget in self.scroll_lista.winfo_children():
            widget.destroy()

        for i, config in enumerate(self.lista_configs):
            frame_item = ctk.CTkFrame(self.scroll_lista, fg_color="#2B2B2B", corner_radius=8)
            frame_item.grid(row=i, column=0, padx=10, pady=5, sticky="ew")
            frame_item.grid_columnconfigure(0, weight=1)

            # Clique no nome abre a configuração
            lbl = ctk.CTkLabel(frame_item, text=config["titulo"], font=ctk.CTkFont(size=15, weight="bold"), cursor="hand2")
            lbl.grid(row=0, column=0, padx=15, pady=10, sticky="w")
            lbl.bind("<Button-1>", lambda e, c=config: self.abrir_configuracao(c))

            # Indicador de quantos filtros existem
            qtd_filtros = len(config.get("filtros", []))
            lbl_info = ctk.CTkLabel(frame_item, text=f"{qtd_filtros} filtros", text_color="gray")
            lbl_info.grid(row=0, column=1, padx=10)

            btn_del = ctk.CTkButton(frame_item, text="Excluir", fg_color="#CC3333", width=60, height=25,
                                   command=lambda c=config: self.excluir_item(c))
            btn_del.grid(row=0, column=2, padx=10, pady=10)

if __name__ == "__main__":
    app = InterfaceConfig()
    app.mainloop()




























# import customtkinter as ctk
# import json
# import os
# from componentes import JanelaAdicionar # Importando o seu componente

# ctk.set_appearance_mode("System")
# ctk.set_default_color_theme("blue")

# class InterfaceConfig(ctk.CTk):
#     def __init__(self):
#         super().__init__()

#         self.title("Gerenciador de Configurações")
#         self.geometry("500x600")
#         self.arquivo_db = "configuracoes.json"
#         self.lista_configs = []

#         self.grid_columnconfigure(0, weight=1)
#         self.grid_rowconfigure(1, weight=1)

#         # Cabeçalho
#         self.frame_topo = ctk.CTkFrame(self, fg_color="transparent")
#         self.frame_topo.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
#         self.frame_topo.grid_columnconfigure(0, weight=1)

#         self.label_titulo = ctk.CTkLabel(self.frame_topo, text="Configurações", font=ctk.CTkFont(size=22, weight="bold"))
#         self.label_titulo.grid(row=0, column=0, sticky="w")

#         self.btn_abrir_popup = ctk.CTkButton(self.frame_topo, text="+ Nova Config", command=self.abrir_janela_add)
#         self.btn_abrir_popup.grid(row=0, column=1)

#         # Lista com Scroll
#         self.scroll_lista = ctk.CTkScrollableFrame(self, label_text="")
#         self.scroll_lista.grid(row=1, column=0, padx=20, pady=(0, 20), sticky="nsew")
#         self.scroll_lista.grid_columnconfigure(0, weight=1)

#         self.carregar_dados()

#     def abrir_janela_add(self):
#         # Instancia a classe que está no outro arquivo
#         JanelaAdicionar(self, self.processar_adicao, titulo_janela="Configurações")

#     def processar_adicao(self, texto, janela_popup):
#         texto = texto.strip()
#         if texto:
#             self.lista_configs.append(texto)
#             self.salvar_dados()
#             self.atualizar_interface_lista()
#             janela_popup.destroy()

#     def carregar_dados(self):
#         if os.path.exists(self.arquivo_db):
#             try:
#                 with open(self.arquivo_db, "r", encoding="utf-8") as f:
#                     self.lista_configs = json.load(f)
#             except: self.lista_configs = []
#         self.atualizar_interface_lista()

#     def salvar_dados(self):
#         with open(self.arquivo_db, "w", encoding="utf-8") as f:
#             json.dump(self.lista_configs, f, indent=4)

#     def excluir_item(self, item):
#         if item in self.lista_configs:
#             self.lista_configs.remove(item)
#             self.salvar_dados()
#             self.atualizar_interface_lista()

#     def atualizar_interface_lista(self):
#         for widget in self.scroll_lista.winfo_children():
#             widget.destroy()

#         for i, item in enumerate(self.lista_configs):
#             # O frame que serve de fundo para o item
#             frame_item = ctk.CTkFrame(
#                 self.scroll_lista, 
#                 fg_color="#939393",        # Define o fundo como preto
#                 corner_radius=8,             # Arredonda os cantos para um visual moderno
#                 border_width=1,              # Opcional: adiciona uma borda fina
#                 border_color="#6D6D6D"     # Cor da borda (cinza escuro)
#             )
#             frame_item.grid(row=i, column=0, padx=10, pady=2, sticky="ew")
#             frame_item.grid_columnconfigure(0, weight=1)

#             # Importante: Ajuste a cor do texto para branco ou cinza claro para dar contraste
#             lbl = ctk.CTkLabel(
#                 frame_item, 
#                 text=item, 
#                 text_color="white",      # Texto branco sobre fundo preto
#                 font=ctk.CTkFont(size=14)
#             )
#             lbl.grid(row=0, column=0, padx=15, pady=10, sticky="w")

#             btn_del = ctk.CTkButton(
#                 frame_item, 
#                 text="Excluir", 
#                 fg_color="#CC3333", 
#                 width=60,
#                 command=lambda x=item: self.excluir_item(x)
#             )
#             btn_del.grid(row=0, column=1, padx=10, pady=10)

# if __name__ == "__main__":
#     app = InterfaceConfig()
#     app.mainloop()