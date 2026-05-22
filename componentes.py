import customtkinter as ctk

class JanelaAdicionar(ctk.CTkToplevel):
    def __init__(
        self, 
        parent, 
        callback_confirmar, 
        titulo_janela="Nova Configuração", # Valor padrão
        texto_label="Digite o título:"      # Valor padrão
    ):
        super().__init__(parent)
        
        self.title(titulo_janela)
        self.geometry("350x220")
        
        # Garante que o pop-up fique à frente
        self.attributes("-topmost", True)
        self.grab_set()

        self.grid_columnconfigure((0, 1), weight=1)

        # Label dinâmica baseada no parâmetro
        self.label = ctk.CTkLabel(self, text=texto_label, font=ctk.CTkFont(size=14))
        self.label.grid(row=0, column=0, columnspan=2, padx=20, pady=20)

        self.entry_titulo = ctk.CTkEntry(self, placeholder_text="Escreva aqui...", width=250)
        self.entry_titulo.grid(row=1, column=0, columnspan=2, padx=20, pady=10)
        self.entry_titulo.focus()

        # Botões
        self.btn_cancelar = ctk.CTkButton(
            self, text="Cancelar", fg_color="gray", command=self.destroy
        )
        self.btn_cancelar.grid(row=2, column=0, padx=20, pady=20, sticky="ew")

        self.btn_confirmar = ctk.CTkButton(
            self, text="Adicionar", 
            command=lambda: callback_confirmar(self.entry_titulo.get(), self)
        )
        self.btn_confirmar.grid(row=2, column=1, padx=20, pady=20, sticky="ew")