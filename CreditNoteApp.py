import os
import datetime
import re
import tkinter as tk
from tkinter import messagebox, Toplevel, Listbox, Scrollbar
from PIL import Image, ImageTk 

# Importa todas as funções de backend e constantes
try:
    import customtkinter as ctk
    # Certifique-se de que o backend_data.py está no mesmo diretório
    from backend_data import (
        load_clientes, save_clientes, load_estado, save_estado,
        process_and_save_note, SAIDA_FOLDER
    )
except ImportError as e:
    print(f"Erro ao importar backend ou customtkinter: {e}")
    print("Verifique se backend_data.py existe e se 'customtkinter' e 'Pillow' estão instalados.")
    exit()

# --- Constantes de Arquivo ---
LOGO_FILEPATH = "Gemini_Generated_Image_tlt5qdtlt5qdtlt5.png"

# --- Cores Customizadas (Tema Verde Moderno) ---
CTK_COLOR_BACKGROUND = ("#ECF0F1", "#2C3E50")  # Cinza claro/Chumbo escuro (Fundo geral)
CTK_COLOR_PRIMARY = "#27AE60"                   # Verde Esmeralda (Foco, botões principais, títulos)
CTK_COLOR_SECONDARY = "#196F3D"                 # Verde Escuro (Hover e detalhes)
CTK_COLOR_SUCCESS = "#27AE60"                   # Mantido verde para Sucesso
CTK_COLOR_DANGER = "#E74C3C"                     # Vermelho (Excluir)
CTK_COLOR_ACCENT = "#95A5A6"                    # Cinza Chumbo (Imprimir)
CTK_COLOR_HEADER = ("#BDC3C7", "#1E2A38")       # Cinza Claro Suave/Azul Escuro para Header (Mantenho contraste)

# --- Cores para Listbox (Zebrado) ---
COLOR_LIST_EVEN = "#F8F8F8"  # Cinza muito claro para linhas pares
COLOR_LIST_ODD = "#FFFFFF"   # Branco para linhas ímpares


# --- Classe Principal da Aplicação GUI ---

class CreditNoteApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Setup da Janela Principal ---
        self.title("Sistema de Gestão de Notas de Crédito")
        self.geometry("1000x750")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue") # Theme base

        # --- Dados do Backend ---
        self.clientes = load_clientes()
        self.estado = load_estado()
        self.selected_client = None
        self.last_saved_file = None 
        
        # Variável para armazenar a referência da imagem (necessário para Tkinter)
        self.logo_image = None

        # --- Configuração de Layout (Grid) ---
        self.grid_columnconfigure((0, 1), weight=1)
        self.grid_rowconfigure(1, weight=1) # Row 1 é o conteúdo principal

        # --- Header com Logo e Título ---
        self._setup_header()

        # --- Frames de Conteúdo Principal ---
        content_frame = ctk.CTkFrame(self, fg_color="transparent")
        content_frame.grid(row=1, column=0, columnspan=2, padx=15, pady=(0, 15), sticky="nsew")
        content_frame.grid_columnconfigure((0, 1), weight=1)
        content_frame.grid_rowconfigure(0, weight=1)

        # Frame de Gerenciamento de Clientes (Esquerda)
        self.client_frame = ctk.CTkFrame(content_frame, fg_color=CTK_COLOR_BACKGROUND)
        self.client_frame.grid(row=0, column=0, padx=10, pady=0, sticky="nsew")

        # Frame de Geração de Nota (Direita)
        self.note_frame = ctk.CTkFrame(content_frame, fg_color=("white", "gray15"))
        self.note_frame.grid(row=0, column=1, padx=10, pady=0, sticky="nsew")

        # Configurações internas dos frames
        self._setup_client_management(self.client_frame)
        self._setup_note_generation(self.note_frame)

    def _setup_header(self):
        """Cria o cabeçalho superior para logo e título principal, centralizando o título."""
        header_frame = ctk.CTkFrame(self, fg_color=CTK_COLOR_HEADER, height=80, corner_radius=0)
        header_frame.grid(row=0, column=0, columnspan=2, sticky="new")
        header_frame.grid_columnconfigure(0, weight=1) # Coluna para logo e título

        # Frame interno para centralizar logo e título
        center_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        center_frame.grid(row=0, column=0, sticky="nsew", padx=20)
        center_frame.grid_columnconfigure((0, 2), weight=1) # Espaços vazios para centralizar
        center_frame.grid_columnconfigure(1, weight=0) # Conteúdo central

        # --- Lógica de Carregamento da Imagem ---
        try:
            # 1. Carrega a imagem e redimensiona (64x64 agora)
            img = Image.open(LOGO_FILEPATH)
            img = img.resize((64, 64), Image.Resampling.LANCZOS)
            self.logo_image = ImageTk.PhotoImage(img)

            # 2. Exibe a imagem
            logo_label = ctk.CTkLabel(center_frame, 
                                      text="", 
                                      image=self.logo_image,
                                      compound="left")
            
        
        except FileNotFoundError:
             # Exibe placeholder se o arquivo não for encontrado
             logo_label = ctk.CTkLabel(center_frame, 
                         text="[LOGO NÃO ENCONTRADO]", 
                         font=ctk.CTkFont(size=10),
                         text_color=CTK_COLOR_DANGER)
        except Exception:
            # Lida com outros erros, como arquivo corrompido
            logo_label = ctk.CTkLabel(center_frame, 
                         text="[ERRO AO CARREGAR LOGO]", 
                         font=ctk.CTkFont(size=10),
                         text_color=CTK_COLOR_DANGER)
            
        
        # Título principal (Centralizado)
        title_label = ctk.CTkLabel(center_frame, 
                     text="GESTOR DE NOTAS DE CRÉDITO", 
                     font=ctk.CTkFont(size=20, weight="bold", family="Arial"),
                     text_color=CTK_COLOR_PRIMARY,
                     anchor="center")
        
        # Centralizando Logo e Título no Frame Interno
        logo_label.grid(row=0, column=1, padx=(0, 10), pady=10, sticky="w")
        title_label.grid(row=0, column=1, padx=(70, 0), pady=10, sticky="e") # Ajuste de padding para logo e texto


    # --- Setup: Gerenciamento de Clientes ---

    def _setup_client_management(self, master):
        master.grid_columnconfigure(0, weight=1)
        master.grid_rowconfigure(5, weight=1)

        # Título da Seção
        ctk.CTkLabel(master, 
                     text="📋 Gestão de Clientes", 
                     font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=CTK_COLOR_PRIMARY
                     ).grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")

        # Campo de busca
        self.client_search_entry = ctk.CTkEntry(master, placeholder_text="Buscar por Código ou Nome...", corner_radius=10, fg_color=("white", "gray25"))
        self.client_search_entry.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="ew")
        self.client_search_entry.bind("<KeyRelease>", self._filter_client_list)

        # Botões de Ação
        button_frame = ctk.CTkFrame(master, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=20, pady=5, sticky="ew")
        button_frame.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkButton(button_frame, text="Novo", command=self._add_client_dialog, fg_color=CTK_COLOR_PRIMARY, hover_color=CTK_COLOR_SECONDARY, corner_radius=10).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(button_frame, text="Editar", command=self._edit_client_dialog, fg_color=CTK_COLOR_PRIMARY, hover_color=CTK_COLOR_SECONDARY, corner_radius=10).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(button_frame, text="Excluir", command=self._delete_client, fg_color=CTK_COLOR_DANGER, hover_color="#C0392B", corner_radius=10).grid(row=0, column=2, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(master, text="Clientes Cadastrados:", anchor="w", font=ctk.CTkFont(weight="bold")).grid(row=3, column=0, padx=20, pady=(10, 5), sticky="ew")

        # Listbox para clientes (Listbox nativa para zebrado)
        list_frame = ctk.CTkFrame(master, corner_radius=10, fg_color=("white", "gray25"))
        list_frame.grid(row=4, column=0, padx=20, pady=(0, 20), sticky="nsew")
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)

        self.client_listbox = Listbox(
            list_frame, 
            height=20, 
            borderwidth=0, 
            highlightthickness=0, 
            selectmode=tk.SINGLE, 
            font=("Inter", 12),
            fg="#202020",
            selectbackground=CTK_COLOR_PRIMARY,
            selectforeground="white"
        )
        self.client_listbox.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)

        scrollbar = ctk.CTkScrollbar(list_frame, command=self.client_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.client_listbox.config(yscrollcommand=scrollbar.set)

        self.client_listbox.bind('<<ListboxSelect>>', self._select_client_from_list)
        self._update_client_list()


    def _update_client_list(self, filter_text=""):
        """
        Atualiza a Listbox de clientes, ordenando por código e aplicando cor zebrada.
        A funcionalidade zebrada é aplicada aqui.
        """
        self.client_listbox.delete(0, tk.END)
        self.filtered_clients = []
        filter_text = filter_text.lower()
        
        # 1. Ordena a lista de clientes por 'codigo' antes de filtrar
        sorted_clients = sorted(self.clientes, key=lambda c: c['codigo'].lower())
        
        listbox_index = 0

        for client in sorted_clients:
            if filter_text in client['codigo'].lower() or filter_text in client['nome'].lower():
                display_text = f"[{client['codigo']}] - {client['nome']}"
                
                # 2. Insere o item e aplica a cor zebrada
                self.client_listbox.insert(tk.END, display_text)
                
                if listbox_index % 2 == 0:
                    # Linha ímpar (index 0, 2, 4...)
                    bg_color = COLOR_LIST_ODD 
                else:
                    # Linha par (index 1, 3, 5...)
                    bg_color = COLOR_LIST_EVEN
                
                # Aplica a cor de fundo
                self.client_listbox.itemconfig(tk.END, {'bg': bg_color})
                
                self.filtered_clients.append(client)
                listbox_index += 1

    def _filter_client_list(self, event):
        """Disparado pela digitação na caixa de busca."""
        self._update_client_list(self.client_search_entry.get())

    def _select_client_from_list(self, event):
        """Seleciona um cliente da Listbox e atualiza o frame da nota."""
        try:
            selection = self.client_listbox.curselection()
            if not selection:
                return

            index = selection[0]
            self.selected_client = self.filtered_clients[index]
            self.client_code_var.set(self.selected_client['codigo'])
            self.client_name_label.configure(text=self.selected_client['nome'])
            
            # Não usar messagebox para seleção, apenas feedback visual ou na barra de status
            # messagebox.showinfo("Cliente Selecionado", f"Cliente {self.selected_client['nome']} selecionado para a nota.")

        except Exception as e:
            print(f"Erro ao selecionar cliente: {e}")
            messagebox.showerror("Erro", "Falha ao selecionar o cliente.")
            self.selected_client = None

    # Gerenciamento de Modal (Cadastrar/Editar)

    def _add_client_dialog(self):
        """Exibe um diálogo para cadastrar um novo cliente."""
        self._show_client_modal("Cadastrar Cliente", None)

    def _edit_client_dialog(self):
        """Exibe um diálogo para editar o cliente selecionado."""
        if not self.selected_client:
            messagebox.showwarning("Atenção", "Selecione um cliente na lista para editar.")
            return

        self._show_client_modal("Editar Cliente", self.selected_client)

    def _show_client_modal(self, title, client_data):
        """Lógica comum do modal de Cadastro/Edição de Cliente."""
        modal = Toplevel(self)
        modal.title(title)
        modal.transient(self)
        modal.grab_set()
        modal.resizable(False, False)
        modal.geometry("350x250")

        self.update_idletasks()
        x = self.winfo_x() + self.winfo_width() // 2 - modal.winfo_width() // 2
        y = self.winfo_y() + self.winfo_height() // 2 - modal.winfo_height() // 2
        modal.geometry(f'+{x}+{y}')

        ctk.CTkLabel(modal, text="Código Único:").pack(pady=(10, 0), padx=10)
        code_entry = ctk.CTkEntry(modal, width=300, corner_radius=10)
        code_entry.pack(padx=10)

        ctk.CTkLabel(modal, text="Nome (Razão Social):").pack(pady=(10, 0), padx=10)
        name_entry = ctk.CTkEntry(modal, width=300, corner_radius=10)
        name_entry.pack(padx=10)

        original_code = None
        if client_data:
            original_code = client_data['codigo']
            code_entry.insert(0, original_code)
            name_entry.insert(0, client_data['nome'])

        def save_client():
            new_code = code_entry.get().strip()
            name = name_entry.get().strip()

            if not new_code or not name:
                messagebox.showerror("Erro", "Código e Nome são obrigatórios.")
                return

            if client_data:
                # Lógica de Edição
                is_duplicate = any(c['codigo'] == new_code and c['codigo'] != original_code for c in self.clientes)
                
                if is_duplicate:
                    messagebox.showerror("Erro", f"O Código de Cliente '{new_code}' já existe para outro cliente.")
                    return
                
                client_data['codigo'] = new_code
                client_data['nome'] = name
                messagebox.showinfo("Sucesso", f"Cliente {new_code} atualizado.")
                
                if self.selected_client and original_code == self.selected_client['codigo']:
                    self.client_code_var.set(new_code)
                    self.client_name_label.configure(text=name)

            else:
                # Lógica de Cadastro
                if any(c['codigo'] == new_code for c in self.clientes):
                    messagebox.showerror("Erro", f"O Código de Cliente '{new_code}' já existe.")
                    return
                self.clientes.append({"codigo": new_code, "nome": name})
                messagebox.showinfo("Sucesso", f"Cliente {new_code} cadastrado.")

            save_clientes(self.clientes)
            self._update_client_list(self.client_search_entry.get())
            modal.destroy()

        ctk.CTkButton(modal, text="Salvar", command=save_client, fg_color=CTK_COLOR_SUCCESS, hover_color="#27AE60", corner_radius=10).pack(pady=20, padx=10)

    def _delete_client(self):
        """Exclui o cliente selecionado."""
        if not self.selected_client:
            messagebox.showwarning("Atenção", "Selecione um cliente para excluir.")
            return

        code = self.selected_client['codigo']
        name = self.selected_client['nome']

        if messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir o cliente:\n[{code}] - {name}?"):
            self.clientes = [c for c in self.clientes if c['codigo'] != code]
            save_clientes(self.clientes)
            self.selected_client = None
            self.client_code_var.set("")
            self.client_name_label.configure(text="Nenhum cliente selecionado")
            self._update_client_list(self.client_search_entry.get())
            messagebox.showinfo("Sucesso", f"Cliente {code} excluído com sucesso.")

    # --- Setup: Geração de Nota de Crédito ---

    def _setup_note_generation(self, master):
        master.grid_columnconfigure(0, weight=1)
        master.grid_rowconfigure(5, weight=1)

        # Título da Seção
        ctk.CTkLabel(master, 
                     text="📝 Geração de Nota de Crédito", 
                     font=ctk.CTkFont(size=18, weight="bold"),
                     text_color=CTK_COLOR_PRIMARY
                     ).grid(row=0, column=0, padx=20, pady=(20, 10), sticky="ew")

        # 1. Dados da Fatura (Data e Número)
        data_frame = ctk.CTkFrame(master, fg_color="transparent")
        data_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        data_frame.grid_columnconfigure((0, 1), weight=1)

        # Data
        ctk.CTkLabel(data_frame, text="Data (DD/MM/AAAA):", anchor="w").grid(row=0, column=0, padx=5, pady=(5, 0), sticky="w")
        self.date_var = tk.StringVar(value=datetime.date.today().strftime("%d/%m/%Y"))
        self.date_var.trace_add("write", self._format_date_input)
        self.date_entry = ctk.CTkEntry(data_frame, textvariable=self.date_var, corner_radius=10, fg_color=("white", "gray25"))
        self.date_entry.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew")

        # Número da Fatura
        self.invoice_label = ctk.CTkLabel(data_frame, text="Número da Fatura (Próx. Sugerido: {}):"
            .format(self.estado['ultima_fatura']), anchor="w")
        self.invoice_label.grid(row=0, column=1, padx=5, pady=(5, 0), sticky="w")
        self.invoice_number_var = ctk.StringVar(value=str(self.estado['ultima_fatura']))
        self.invoice_number_entry = ctk.CTkEntry(data_frame, textvariable=self.invoice_number_var, corner_radius=10, fg_color=("white", "gray25"))
        self.invoice_number_entry.grid(row=1, column=1, padx=5, pady=(0, 5), sticky="ew")

        # 2. Dados do Cliente
        client_info_frame = ctk.CTkFrame(master, fg_color="transparent")
        client_info_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        client_info_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(client_info_frame, text="Cliente Selecionado:", font=ctk.CTkFont(weight="bold"), text_color=CTK_COLOR_PRIMARY).grid(row=0, column=0, padx=5, pady=(5, 0), sticky="w")

        code_name_frame = ctk.CTkFrame(client_info_frame, fg_color="transparent")
        code_name_frame.grid(row=1, column=0, sticky="ew", padx=5, pady=5)
        code_name_frame.grid_columnconfigure((0, 1), weight=1)

        # Código
        ctk.CTkLabel(code_name_frame, text="Código:", anchor="w").grid(row=0, column=0, padx=(0, 5), pady=(0, 0), sticky="w")
        self.client_code_var = ctk.StringVar(value="")
        ctk.CTkEntry(code_name_frame, textvariable=self.client_code_var, state="readonly", corner_radius=10, fg_color=("#ECF0F1", "gray30")).grid(row=1, column=0, padx=(0, 5), pady=(0, 5), sticky="ew")

        # Nome
        ctk.CTkLabel(code_name_frame, text="Razão Social:", anchor="w").grid(row=0, column=1, padx=(5, 0), pady=(0, 0), sticky="w")
        self.client_name_label = ctk.CTkLabel(code_name_frame, text="Nenhum cliente selecionado", anchor="w", fg_color=("#ECF0F1", "gray30"), corner_radius=10, text_color=("#2C3E50", "white"))
        self.client_name_label.grid(row=1, column=1, padx=(5, 0), pady=(0, 5), sticky="ew", ipady=5)


        # 3. Valor e Descrição
        value_desc_frame = ctk.CTkFrame(master, fg_color="transparent")
        value_desc_frame.grid(row=3, column=0, padx=20, pady=10, sticky="ew")
        value_desc_frame.grid_columnconfigure(0, weight=1)

        # Valor da Fatura
        ctk.CTkLabel(value_desc_frame, text="Valor da Fatura (R$):", anchor="w").grid(row=0, column=0, padx=5, pady=(5, 0), sticky="w")
        self.value_var = tk.StringVar(value="0,00")
        self.value_var.trace_add("write", self._format_currency_input)
        self.value_entry = ctk.CTkEntry(value_desc_frame, textvariable=self.value_var, corner_radius=10, fg_color=("white", "gray25"))
        self.value_entry.grid(row=1, column=0, padx=5, pady=(0, 10), sticky="ew")

        # Descrição/Histórico
        ctk.CTkLabel(master, text=f"Histórico/Descrição (Última nota: '{self.estado['ultima_descricao'][:30]}...'):", font=ctk.CTkFont(weight="bold")).grid(row=4, column=0, padx=20, pady=(5, 0), sticky="w")
        
        self.description_textbox = ctk.CTkTextbox(master, height=150, corner_radius=10, fg_color=("white", "gray25"))
        self.description_textbox.grid(row=5, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.description_textbox.insert("1.0", self.estado['ultima_descricao'])

        # 4. Botão de Ação
        action_frame = ctk.CTkFrame(master, fg_color="transparent")
        action_frame.grid(row=6, column=0, padx=20, pady=10, sticky="ew")
        action_frame.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(action_frame, 
                      text="✅ Gerar Nota de Crédito (.xlsx)", 
                      command=self._process_note, 
                      fg_color=CTK_COLOR_PRIMARY, 
                      hover_color=CTK_COLOR_SECONDARY,
                      font=ctk.CTkFont(weight="bold", size=15),
                      corner_radius=10
                      ).grid(row=0, column=0, padx=5, pady=5, sticky="ew", ipady=5)
                      
        ctk.CTkButton(action_frame, 
                      text="🖨️ Imprimir Última Nota", 
                      command=self._print_last_note, 
                      fg_color=CTK_COLOR_ACCENT, 
                      hover_color="#7F8C8D",
                      corner_radius=10
                      ).grid(row=0, column=1, padx=5, pady=5, sticky="ew", ipady=5)

    # --- Funções de Formatação e Validação de Input ---

    def _format_date_input(self, *args):
        """
        Formata o campo de data em DD/MM/AAAA em tempo real.
        """
        current = self.date_var.get()
        cursor_pos = self.date_entry.index(tk.INSERT)
        
        numeric = re.sub(r'[^0-9]', '', current)
        numeric = numeric[:8]

        new_value = ""
        if len(numeric) > 0:
            new_value += numeric[0:2]
        if len(numeric) > 2:
            new_value += "/" + numeric[2:4]
        if len(numeric) > 4:
            new_value += "/" + numeric[4:8]
        
        old_numeric_prefix = re.sub(r'[^0-9]', '', current[:cursor_pos])
        target_numeric_length = len(old_numeric_prefix)
        
        new_cursor_pos = 0
        numeric_count = 0
        
        for char in new_value:
            if char.isdigit():
                numeric_count += 1
            
            new_cursor_pos += 1
            
            if numeric_count == target_numeric_length:
                break
        
        if target_numeric_length == 0:
            new_cursor_pos = 0
        elif cursor_pos >= len(current):
            new_cursor_pos = len(new_value)
        elif len(new_value) > len(current) and new_value[new_cursor_pos - 1] == '/':
             new_cursor_pos += 1

        if new_value != current:
            self.date_var.set(new_value)
            try:
                self.date_entry.icursor(new_cursor_pos)
            except:
                pass 
        else:
            try:
                self.date_entry.icursor(cursor_pos)
            except:
                pass


    def _format_currency_input(self, *args):
        """Formata o campo de valor em R$ X.XXX,XX em tempo real."""
        current = self.value_var.get()
        cursor_pos = self.value_entry.index(tk.INSERT)

        cleaned_value = re.sub(r'[^0-9]', '', current)

        if not cleaned_value:
            new_value = "0,00"
            new_cursor_pos = 3
        else:
            value_cents = int(cleaned_value)
            value_reais = value_cents / 100

            new_value = f"{value_reais:,.2f}".replace('.', '#').replace(',', '.').replace('#', ',')
            
            original_numeric_len = len(re.sub(r'[^0-9]', '', current[:cursor_pos]))
            
            new_cursor_pos = 0
            numeric_count = 0
            for i, char in enumerate(new_value):
                if char.isdigit():
                    numeric_count += 1
                    if numeric_count > original_numeric_len:
                        break
                new_cursor_pos = i + 1
            
            if cursor_pos >= len(current):
                new_cursor_pos = len(new_value)


        if new_value != current:
            self.value_var.set(new_value)
            try:
                self.value_entry.icursor(new_cursor_pos)
            except:
                pass
        
    # --- Funções de Processamento da Nota ---

    def _process_note(self):
        """Valida os dados e chama o backend para processar e salvar, e então pergunta sobre a impressão."""
        # 1. Validação dos Dados (Frontend)
        if not self.selected_client:
            messagebox.showerror("Erro de Validação", "Selecione um cliente para prosseguir.")
            return

        data_input = self.date_var.get()
        invoice_number = self.invoice_number_var.get()
        description_text = self.description_textbox.get("1.0", tk.END).strip()
        value_formatted = self.value_var.get()

        if not re.fullmatch(r'\d{2}/\d{2}/\d{4}', data_input):
            messagebox.showerror("Erro de Validação", "Formato de Data inválido. Use DD/MM/AAAA.")
            return
        
        if not invoice_number.isdigit() or int(invoice_number) <= 0:
            messagebox.showerror("Erro de Validação", "Número da Fatura deve ser um número inteiro positivo.")
            return

        try:
            value_cleaned = value_formatted.replace('.', '').replace(',', '.')
            value_float = float(value_cleaned)
        except ValueError:
            messagebox.showerror("Erro de Validação", "Valor da Fatura inválido.")
            return

        if not description_text:
            messagebox.showerror("Erro de Validação", "A Descrição/Histórico é obrigatória.")
            return
        
        # 2. Chama a função de Backend para processar o XLSX
        success, result_or_path = process_and_save_note(
            data_input, 
            invoice_number, 
            self.selected_client['codigo'], 
            self.selected_client['nome'], 
            description_text, 
            value_float, 
            self.estado # O estado é atualizado dentro do backend
        )

        if success:
            output_path = result_or_path
            self.last_saved_file = output_path
            
            # 3. Atualiza a GUI com o novo estado
            self.invoice_number_var.set(str(self.estado['ultima_fatura']))
            
            # Atualiza o label de sugestão da fatura
            self.invoice_label.configure(text="Número da Fatura (Próx. Sugerido: {}):".format(self.estado['ultima_fatura']))
            
            self.description_textbox.delete("1.0", tk.END)
            self.description_textbox.insert("1.0", self.estado['ultima_descricao'])
            self.value_var.set("0,00")
            
            # 4. Sucesso e Pergunta de Impressão (Mudança aqui)
            
            # Usa tk.messagebox.askyesno, que retorna True para 'Yes' e False para 'No'
            if messagebox.askyesno(
                "Nota Gerada com Sucesso!", 
                f"Nota salva em:\n{output_path}\n\nDeseja enviar o arquivo para IMPRESSÃO agora?"
            ):
                self._print_file(output_path)
            else:
                 messagebox.showinfo("Geração Concluída", "A nota foi gerada e salva com sucesso.")
        else:
            messagebox.showerror("Erro de Processamento", result_or_path)

            
    def _print_file(self, filepath):
        """Tenta abrir o arquivo com o programa padrão (o que geralmente abre a caixa de diálogo de impressão)."""
        try:
            if os.name == 'nt':  # Windows
                # os.startfile("notepad.exe") # Exemplo para abrir o programa
                os.startfile(filepath, "print")
                messagebox.showinfo("Comando de Impressão Enviado", "A caixa de diálogo da impressora deve ter sido aberta.")
            elif os.uname()[0] == 'Darwin':  # macOS
                os.system(f'open -a "Microsoft Excel" -p "{filepath}"')
                messagebox.showinfo("Comando de Impressão Enviado", "Verifique a fila de impressão do sistema.")
            else:  # Linux (genérico)
                os.system(f'xdg-open "{filepath}"')
                messagebox.showinfo("Impressão", "Comando enviado. Verifique a fila de impressão do sistema ou abra o arquivo manualmente.")
        except Exception as e:
            messagebox.showerror("Erro de Impressão/Abertura", f"Não foi possível abrir o arquivo para impressão. Tente abrir o arquivo manualmente: {filepath}\nDetalhe: {e}")

    def _print_last_note(self):
        """Imprime o último arquivo salvo."""
        if not self.last_saved_file or not os.path.exists(self.last_saved_file):
            # Tenta encontrar o último arquivo salvo se a aplicação foi reiniciada
            if os.path.exists(SAIDA_FOLDER):
                files = [os.path.join(SAIDA_FOLDER, f) for f in os.listdir(SAIDA_FOLDER)]
                if files:
                    self.last_saved_file = max(files, key=os.path.getctime)
            
        if not self.last_saved_file or not os.path.exists(self.last_saved_file):
            messagebox.showwarning("Atenção", "Nenhuma nota foi gerada nesta sessão ou o último arquivo salvo não foi encontrado.")
            return

        if messagebox.askyesno("Confirmar Impressão", f"Deseja enviar para impressão o arquivo:\n{os.path.basename(self.last_saved_file)}?"):
            self._print_file(self.last_saved_file)


if __name__ == "__main__":
    app = CreditNoteApp()
    app.mainloop()
