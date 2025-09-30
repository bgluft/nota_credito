import os
import datetime
import re
import tkinter as tk
from tkinter import messagebox, Toplevel, Listbox, Scrollbar
from PIL import Image, ImageTk 
# Importa sys, mas não define _get_resource_path, ele vem do backend

# Importa todas as funções de backend e constantes
try:
    import customtkinter as ctk
    # Importa o utilitário de caminho e a constante do logo
    from backend_data import (
        load_clientes, save_clientes, load_estado, save_estado,
        process_and_save_note, SAIDA_FOLDER, _get_resource_path,
        load_templates, save_templates # NOVAS FUNÇÕES
    )
except ImportError as e:
    print(f"Erro ao importar backend ou customtkinter: {e}")
    print("Verifique se backend_data.py existe e se 'customtkinter' e 'Pillow' estão instalados.")
    exit()

# --- Constantes de Arquivo ---
# O logo precisa do caminho absoluto, obtido via _get_resource_path no setup
LOGO_FILEPATH = "Gemini_Generated_Image_tlt5qdtlt5qdtlt5.png"

# --- Cores Customizadas (Tema Verde Moderno Refinado) ---
CTK_COLOR_BACKGROUND = ("#F4F6F6", "#2C3E50")  # Fundo geral mais neutro (Claro/Chumbo escuro)
CTK_COLOR_PRIMARY = "#27AE60"                   # Verde Esmeralda (Foco, botões principais, títulos)
CTK_COLOR_SECONDARY = "#196F3D"                 # Verde Escuro (Hover e detalhes)
CTK_COLOR_SUCCESS = "#27AE60"                   # Sucesso
CTK_COLOR_DANGER = "#E74C3C"                     # Vermelho (Excluir)
CTK_COLOR_ACCENT = "#95A5A6"                    # Cinza Chumbo (Imprimir)
CTK_COLOR_HEADER = ("#BDC3C7", "#1E2A38")       # Cinza Claro Suave/Azul Escuro para Header 
CTK_COLOR_PANEL = ("#FFFFFF", "#2E4053")        # Painel de Conteúdo (Mais contraste)
CTK_COLOR_BUTTON_GENERATE = "#1ABC9C"           # Verde Água para gerar/imprimir

# --- Cores para Listbox (Zebrado com mais contraste) ---
COLOR_LIST_EVEN = "#E8EAEB"  # Cinza claro mais perceptível
COLOR_LIST_ODD = "#FFFFFF"   # Branco


# --- Classe Principal da Aplicação GUI ---

class CreditNoteApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Setup da Janela Principal ---
        self.title("Sistema de Gestão de Notas de Crédito")
        self.geometry("1000x750")
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue") 

        self.configure(fg_color=CTK_COLOR_BACKGROUND) # Aplica o fundo geral
        
        # --- Dados do Backend ---
        self.clientes = load_clientes()
        self.estado = load_estado()
        self.templates = load_templates() # CARREGA NOVOS TEMPLATES
        self.selected_client = None
        self.last_saved_file = None 
        self.logo_image = None
        self.selected_template = None # NOVO: Template selecionado

        # --- Configuração de Layout (Grid) ---
        self.grid_columnconfigure((0, 1), weight=1)
        self.grid_rowconfigure(2, weight=1) # Row 2 é o conteúdo principal

        # --- Header com Logo e Título ---
        self._setup_header()
        
        # --- Separador Visual ---
        ctk.CTkFrame(self, height=2, fg_color=CTK_COLOR_PRIMARY).grid(
            row=1, column=0, columnspan=2, sticky="ew", padx=0, pady=0
        )

        # --- Frames de Conteúdo Principal ---
        content_frame = ctk.CTkFrame(self, fg_color="transparent")
        content_frame.grid(row=2, column=0, columnspan=2, padx=15, pady=15, sticky="nsew")
        content_frame.grid_columnconfigure((0, 1), weight=1)
        content_frame.grid_rowconfigure(0, weight=1)

        # Frame de Gerenciamento de Clientes (Esquerda)
        self.client_frame = ctk.CTkFrame(content_frame, 
                                         fg_color=CTK_COLOR_PANEL, 
                                         corner_radius=15,
                                         border_width=2,
                                         border_color=("#D5DBDB", "#283747")) 
        self.client_frame.grid(row=0, column=0, padx=10, pady=0, sticky="nsew")

        # Frame de Geração de Nota (Direita)
        self.note_frame = ctk.CTkFrame(content_frame, 
                                       fg_color=CTK_COLOR_PANEL, 
                                       corner_radius=15,
                                       border_width=2,
                                       border_color=("#D5DBDB", "#283747"))
        self.note_frame.grid(row=0, column=1, padx=10, pady=0, sticky="nsew")
        
        # Frame extra para Gerenciamento de Templates na coluna Esquerda
        self.template_management_frame = ctk.CTkFrame(self.client_frame, 
                                            fg_color=("#ECF0F1", "gray25"),
                                            corner_radius=10, 
                                            border_width=1,
                                            border_color=("#D5DBDB", "gray20"))
        self.template_management_frame.grid(row=5, column=0, padx=25, pady=(5, 25), sticky="ew")


        # Configurações internas dos frames
        self._setup_client_management(self.client_frame)
        self._setup_template_management(self.template_management_frame) # NOVO SETUP DE TEMPLATE
        self._setup_note_generation(self.note_frame)
        
    # --- Setup de Headers e UI Geral (sem alteração) ---
    def _setup_header(self):
        """Cria o cabeçalho superior para logo e título principal, centralizando o título."""
        header_frame = ctk.CTkFrame(self, fg_color=CTK_COLOR_HEADER, height=90, corner_radius=0)
        header_frame.grid(row=0, column=0, columnspan=2, sticky="new")
        header_frame.grid_columnconfigure(0, weight=1) 

        # Frame interno para centralizar logo e título
        center_frame = ctk.CTkFrame(header_frame, fg_color="transparent")
        center_frame.grid(row=0, column=0, sticky="nsew", padx=20)
        
        # Configuração para centralizar
        center_frame.grid_columnconfigure(0, weight=1) # Espaço expansível à esquerda
        center_frame.grid_columnconfigure(1, weight=0) # Coluna para Logo
        center_frame.grid_columnconfigure(2, weight=0) # Coluna para Título
        center_frame.grid_columnconfigure(3, weight=1) # Espaço expansível à direita

        # --- Lógica de Carregamento da Imagem ---
        try:
            # *USANDO O UTILITÁRIO DE CAMINHO DO BACKEND PARA COMPATIBILIDADE PYINSTALLER*
            logo_path = _get_resource_path(LOGO_FILEPATH)
            
            # 1. Carrega a imagem e redimensiona (64x64)
            img = Image.open(logo_path)
            img = img.resize((64, 64), Image.Resampling.LANCZOS)
            self.logo_image = ImageTk.PhotoImage(img)

            # 2. Exibe a imagem (Coluna 1)
            logo_label = ctk.CTkLabel(center_frame, 
                                      text="", 
                                      image=self.logo_image,
                                      compound="left")
            logo_label.grid(row=0, column=1, padx=(0, 15), pady=12, sticky="e") 
        
        except FileNotFoundError:
             # Exibe placeholder se o arquivo não for encontrado
             logo_label = ctk.CTkLabel(center_frame, 
                         text="[LOGO NÃO ENCONTRADO]", 
                         font=ctk.CTkFont(size=10),
                         text_color=CTK_COLOR_DANGER)
             logo_label.grid(row=0, column=1, padx=(0, 15), pady=12, sticky="e") 

        except Exception:
            # Lida com outros erros, como arquivo corrompido
            logo_label = ctk.CTkLabel(center_frame, 
                         text="[ERRO AO CARREGAR LOGO]", 
                         font=ctk.CTkFont(size=10),
                         text_color=CTK_COLOR_DANGER)
            logo_label.grid(row=0, column=1, padx=(0, 15), pady=12, sticky="e") 
            
        
        # Título principal (Coluna 2)
        title_label = ctk.CTkLabel(center_frame, 
                     text="GESTOR DE NOTAS DE CRÉDITO", 
                     font=ctk.CTkFont(size=24, weight="bold", family="Arial"), # Fonte maior e mais impactante
                     text_color=CTK_COLOR_PRIMARY,
                     anchor="center")
        
        # Colocando o Título logo ao lado do logo
        title_label.grid(row=0, column=2, padx=(15, 0), pady=12, sticky="w") 


    # --- Setup: Gerenciamento de Clientes (com alteração de row) ---

    def _setup_client_management(self, master):
        master.grid_columnconfigure(0, weight=1)
        master.grid_rowconfigure(4, weight=1) # A linha 5 agora é para templates

        # Título da Seção
        ctk.CTkLabel(master, 
                     text="📋 Gestão de Clientes", 
                     font=ctk.CTkFont(size=19, weight="bold"), # Tamanho ligeiramente maior
                     text_color=CTK_COLOR_SECONDARY
                     ).grid(row=0, column=0, padx=25, pady=(25, 15), sticky="ew")

        # Campo de busca
        self.client_search_entry = ctk.CTkEntry(master, placeholder_text="Buscar por Código ou Nome...", corner_radius=10, fg_color=("#ECF0F1", "gray25"))
        self.client_search_entry.grid(row=1, column=0, padx=25, pady=(0, 15), sticky="ew")
        self.client_search_entry.bind("<KeyRelease>", self._filter_client_list)

        # Botões de Ação
        button_frame = ctk.CTkFrame(master, fg_color="transparent")
        button_frame.grid(row=2, column=0, padx=25, pady=5, sticky="ew")
        button_frame.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkButton(button_frame, text="Novo", command=self._add_client_dialog, fg_color=CTK_COLOR_PRIMARY, hover_color=CTK_COLOR_SECONDARY, corner_radius=8).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(button_frame, text="Editar", command=self._edit_client_dialog, fg_color=CTK_COLOR_PRIMARY, hover_color=CTK_COLOR_SECONDARY, corner_radius=8).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(button_frame, text="Excluir", command=self._delete_client, fg_color=CTK_COLOR_DANGER, hover_color="#C0392B", corner_radius=8).grid(row=0, column=2, padx=5, pady=5, sticky="ew")

        ctk.CTkLabel(master, text="Clientes Cadastrados:", anchor="w", font=ctk.CTkFont(weight="bold")).grid(row=3, column=0, padx=25, pady=(15, 5), sticky="ew")

        # Listbox para clientes (Listbox nativa para zebrado)
        list_frame = ctk.CTkFrame(master, corner_radius=10, fg_color=COLOR_LIST_ODD) # Fundo do frame branco
        list_frame.grid(row=4, column=0, padx=25, pady=(0, 5), sticky="nsew") # Ajuste do pady para encaixar templates
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)

        self.client_listbox = Listbox(
            list_frame, 
            height=12, # Reduzido o height
            borderwidth=0, 
            highlightthickness=0, 
            selectmode=tk.SINGLE, 
            font=("Inter", 12),
            fg="#202020",
            selectbackground=CTK_COLOR_SECONDARY,
            selectforeground="white"
        )
        self.client_listbox.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)

        scrollbar = ctk.CTkScrollbar(list_frame, command=self.client_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.client_listbox.config(yscrollcommand=scrollbar.set)

        self.client_listbox.bind('<<ListboxSelect>>', self._select_client_from_list)
        self._update_client_list()
        
    # Lógica de clientes (inalterada)
    def _update_client_list(self, filter_text=""):
        """
        Atualiza a Listbox de clientes, ordenando por código e aplicando cor zebrada.
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
                    # Linha par (index 0, 2, 4...)
                    bg_color = COLOR_LIST_ODD 
                else:
                    # Linha ímpar (index 1, 3, 5...)
                    bg_color = COLOR_LIST_EVEN
                
                # Aplica a cor de fundo
                self.client_listbox.itemconfig(tk.END, {'bg': bg_color})
                
                self.filtered_clients.append(client)
                listbox_index += 1

    def _filter_client_list(self, event):
        self._update_client_list(self.client_search_entry.get())

    def _select_client_from_list(self, event):
        try:
            selection = self.client_listbox.curselection()
            if not selection:
                return

            index = selection[0]
            self.selected_client = self.filtered_clients[index]
            self.client_code_var.set(self.selected_client['codigo'])
            
            # DESTAQUE APLICADO AQUI: Cor primária no nome do cliente selecionado
            self.client_name_label.configure(
                text=self.selected_client['nome'], 
                text_color=CTK_COLOR_PRIMARY
            )
            
        except Exception as e:
            print(f"Erro ao selecionar cliente: {e}")
            messagebox.showerror("Erro", "Falha ao selecionar o cliente.")
            self.selected_client = None

    # Gerenciamento de Modal (Cadastrar/Editar/Excluir)

    def _add_client_dialog(self):
        self._show_client_modal("Cadastrar Cliente", None)

    def _edit_client_dialog(self):
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
                    self.client_name_label.configure(text=name, text_color=CTK_COLOR_PRIMARY) # Atualiza a cor de destaque

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
            self.client_name_label.configure(text="Nenhum cliente selecionado", text_color=("#2C3E50", "white")) # Retorna à cor padrão
            self._update_client_list(self.client_search_entry.get())
            messagebox.showinfo("Sucesso", f"Cliente {code} excluído com sucesso.")


    # --- NOVO SETUP: Gerenciamento de Templates ---
    
    def _setup_template_management(self, master):
        master.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(master, 
                     text="📄 Templates de Descrição", 
                     font=ctk.CTkFont(size=16, weight="bold"),
                     text_color=CTK_COLOR_PRIMARY # Usa cor primária para destaque
                     ).grid(row=0, column=0, padx=15, pady=(15, 10), sticky="w")
        
        button_frame = ctk.CTkFrame(master, fg_color="transparent")
        button_frame.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="ew")
        button_frame.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkButton(button_frame, text="Novo Template", command=self._add_template_dialog, fg_color=CTK_COLOR_SECONDARY, hover_color="#145A32", corner_radius=8, font=ctk.CTkFont(size=12)).grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(button_frame, text="Editar", command=self._edit_template_dialog, fg_color=CTK_COLOR_ACCENT, hover_color="#7F8C8D", corner_radius=8, font=ctk.CTkFont(size=12)).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkButton(button_frame, text="Excluir", command=self._delete_template, fg_color=CTK_COLOR_DANGER, hover_color="#C0392B", corner_radius=8, font=ctk.CTkFont(size=12)).grid(row=0, column=2, padx=5, pady=5, sticky="ew")
        
    def _add_template_dialog(self):
        self._show_template_modal("Novo Template", None)

    def _edit_template_dialog(self):
        # Obtém o nome selecionado no dropdown
        selected_name = self.template_var.get()
        
        # 1. Verifica se a opção selecionada é válida (não é a opção padrão nem vazia)
        is_valid_selection = selected_name not in ["Selecione um template", "Nenhum template disponível", ""]
        
        # 2. Verifica se o nome selecionado existe na lista real de templates
        template_exists = next((t for t in self.templates if t['nome'] == selected_name), None)

        if not is_valid_selection or not template_exists:
            messagebox.showwarning("Atenção", "Selecione um template válido no menu suspenso para editar.")
            return

        self._show_template_modal("Editar Template", template_exists)
            
    def _delete_template(self):
        # Obtém o nome selecionado no dropdown
        selected_name = self.template_var.get()

        # 1. Verifica se a opção selecionada é válida (não é a opção padrão nem vazia)
        is_valid_selection = selected_name not in ["Selecione um template", "Nenhum template disponível", ""]
        
        # 2. Verifica se o nome selecionado existe na lista real de templates
        template_exists = next((t for t in self.templates if t['nome'] == selected_name), None)

        if not is_valid_selection or not template_exists:
            messagebox.showwarning("Atenção", "Selecione um template válido no menu suspenso para excluir.")
            return
            
        if messagebox.askyesno("Confirmar Exclusão", f"Tem certeza que deseja excluir o template:\n'{selected_name}'?"):
            self.templates = [t for t in self.templates if t['nome'] != selected_name]
            save_templates(self.templates)
            self._update_template_dropdown()
            
            # Limpa o campo de descrição se o template excluído estava sendo usado
            if template_exists and self.description_textbox.get("1.0", tk.END).strip() == template_exists['descricao'].strip():
                 self.description_textbox.delete("1.0", tk.END)
                 self.description_textbox.insert("1.0", self.estado['ultima_descricao'])

            messagebox.showinfo("Sucesso", "Template excluído com sucesso.")

    def _show_template_modal(self, title, template_data):
        modal = Toplevel(self)
        modal.title(title)
        modal.transient(self)
        modal.grab_set()
        modal.resizable(False, False)
        modal.geometry("450x400")
        
        # Centraliza modal
        self.update_idletasks()
        x = self.winfo_x() + self.winfo_width() // 2 - modal.winfo_width() // 2
        y = self.winfo_y() + self.winfo_height() // 2 - modal.winfo_height() // 2
        modal.geometry(f'+{x}+{y}')

        ctk.CTkLabel(modal, text="Nome do Template:").pack(pady=(10, 0), padx=10)
        name_entry = ctk.CTkEntry(modal, width=400, corner_radius=10)
        name_entry.pack(padx=10)
        
        ctk.CTkLabel(modal, text="Texto Completo da Descrição:").pack(pady=(10, 0), padx=10)
        desc_textbox = ctk.CTkTextbox(modal, width=400, height=150, corner_radius=10)
        desc_textbox.pack(padx=10)
        
        original_name = None
        if template_data:
            original_name = template_data['nome']
            name_entry.insert(0, original_name)
            desc_textbox.insert("1.0", template_data['descricao'])
        
        def save_template():
            new_name = name_entry.get().strip()
            description = desc_textbox.get("1.0", tk.END).strip()
            
            if not new_name or not description:
                messagebox.showerror("Erro", "Nome e Descrição são obrigatórios.")
                return
            
            # Checa duplicidade
            is_duplicate = any(t['nome'] == new_name and t.get('nome') != original_name for t in self.templates)
            if is_duplicate:
                messagebox.showerror("Erro", f"O nome de template '{new_name}' já existe.")
                return
                
            if template_data:
                # Edição
                template_data['nome'] = new_name
                template_data['descricao'] = description
                messagebox.showinfo("Sucesso", f"Template '{new_name}' atualizado.")
            else:
                # Cadastro
                self.templates.append({"nome": new_name, "descricao": description})
                messagebox.showinfo("Sucesso", f"Template '{new_name}' cadastrado.")

            save_templates(self.templates)
            self._update_template_dropdown(new_name) # Atualiza e pré-seleciona
            modal.destroy()

        ctk.CTkButton(modal, text="Salvar Template", command=save_template, fg_color=CTK_COLOR_SUCCESS, hover_color="#27AE60", corner_radius=10).pack(pady=20, padx=10)

    def _update_template_dropdown(self, selected_name=None):
        """Atualiza as opções do Combobox de templates."""
        self.template_options = [t['nome'] for t in self.templates]
        
        if not self.template_options:
            self.template_options = ["Nenhum template disponível"]
            self.template_var.set(self.template_options[0])
            self.template_dropdown.configure(values=self.template_options)
            return

        self.template_options.insert(0, "Selecione um template")
        self.template_dropdown.configure(values=self.template_options)
        
        if selected_name and selected_name in self.template_options:
            self.template_var.set(selected_name)
        else:
            self.template_var.set(self.template_options[0])

    def _insert_template_description(self, template_name):
        """Insere o texto do template selecionado no campo de Descrição."""
        if template_name in ["Selecione um template", "Nenhum template disponível", ""]:
            # Não faz nada se for a opção padrão/vazia, mantendo o histórico
            return

        selected_template = next((t for t in self.templates if t['nome'] == template_name), None)
        
        if selected_template:
            self.description_textbox.delete("1.0", tk.END)
            self.description_textbox.insert("1.0", selected_template['descricao'])

    # --- Setup: Geração de Nota de Crédito (com alteração de row) ---

    def _setup_note_generation(self, master):
        master.grid_columnconfigure(0, weight=1)
        master.grid_rowconfigure(6, weight=1) # Row 6 para o textbox

        # Título da Seção
        ctk.CTkLabel(master, 
                     text="📝 Geração de Nota de Crédito", 
                     font=ctk.CTkFont(size=19, weight="bold"), # Tamanho ligeiramente maior
                     text_color=CTK_COLOR_SECONDARY
                     ).grid(row=0, column=0, padx=25, pady=(25, 15), sticky="ew")

        # 1. Dados da Fatura (Data e Número)
        data_frame = ctk.CTkFrame(master, fg_color="transparent")
        data_frame.grid(row=1, column=0, padx=25, pady=10, sticky="ew")
        data_frame.grid_columnconfigure((0, 1), weight=1)

        # Data
        ctk.CTkLabel(data_frame, text="Data (DD/MM/AAAA):", anchor="w").grid(row=0, column=0, padx=5, pady=(5, 0), sticky="w")
        self.date_var = tk.StringVar(value=datetime.date.today().strftime("%d/%m/%Y"))
        # REMOVE A TRACE EM TEMPO REAL
        # self.date_var.trace_add("write", self._format_date_input) 
        self.date_entry = ctk.CTkEntry(data_frame, textvariable=self.date_var, corner_radius=10, fg_color=("#ECF0F1", "gray25"))
        # ADICIONA O BIND PARA FORMATAR AO SAIR DO CAMPO
        self.date_entry.bind("<FocusOut>", self._format_date_input_on_focusout) 
        self.date_entry.grid(row=1, column=0, padx=5, pady=(0, 5), sticky="ew")

        # Número da Fatura
        self.invoice_label = ctk.CTkLabel(data_frame, text="Número da Fatura (Próx. Sugerido: {}):"
            .format(self.estado['ultima_fatura']), anchor="w")
        self.invoice_label.grid(row=0, column=1, padx=5, pady=(5, 0), sticky="w")
        self.invoice_number_var = ctk.StringVar(value=str(self.estado['ultima_fatura']))
        self.invoice_number_entry = ctk.CTkEntry(data_frame, textvariable=self.invoice_number_var, corner_radius=10, fg_color=("#ECF0F1", "gray25"))
        self.invoice_number_entry.grid(row=1, column=1, padx=5, pady=(0, 5), sticky="ew")

        # 2. Dados do Cliente
        client_info_frame = ctk.CTkFrame(master, fg_color="transparent")
        client_info_frame.grid(row=2, column=0, padx=25, pady=10, sticky="ew")
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
        # Ajusta a cor padrão para o label de nome (será alterado para verde no select)
        self.client_name_label = ctk.CTkLabel(code_name_frame, text="Nenhum cliente selecionado", anchor="w", fg_color=("#ECF0F1", "gray30"), corner_radius=10, text_color=("#2C3E50", "white"))
        self.client_name_label.grid(row=1, column=1, padx=(5, 0), pady=(0, 5), sticky="ew", ipady=5)


        # 3. Valor e Descrição
        value_desc_frame = ctk.CTkFrame(master, fg_color="transparent")
        value_desc_frame.grid(row=3, column=0, padx=25, pady=10, sticky="ew")
        value_desc_frame.grid_columnconfigure(0, weight=1)

        # Valor da Fatura
        ctk.CTkLabel(value_desc_frame, text="Valor da Fatura (R$):", anchor="w").grid(row=0, column=0, padx=5, pady=(5, 0), sticky="w")
        self.value_var = tk.StringVar(value="0,00")
        self.value_entry = ctk.CTkEntry(value_desc_frame, textvariable=self.value_var, corner_radius=10, fg_color=("#ECF0F1", "gray25"))
        self.value_entry.bind("<FocusOut>", self._format_currency_input_on_focusout)
        self.value_entry.grid(row=1, column=0, padx=5, pady=(0, 10), sticky="ew")

        # 4. Descrição/Histórico
        desc_header_frame = ctk.CTkFrame(master, fg_color="transparent")
        desc_header_frame.grid(row=4, column=0, padx=25, pady=(5, 0), sticky="ew")
        desc_header_frame.grid_columnconfigure(0, weight=1)
        desc_header_frame.grid_columnconfigure(1, weight=1)
        
        # Label de Histórico
        ctk.CTkLabel(desc_header_frame, 
                     text=f"Histórico/Descrição (Última nota: '{self.estado['ultima_descricao'][:30]}...'):", 
                     font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w")
        
        # Dropdown de Templates
        ctk.CTkLabel(desc_header_frame, text="Ou Carregar Template:", anchor="e", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=1, sticky="e", padx=(10, 5))
        
        self.template_var = ctk.StringVar(value="Selecione um template")
        self.template_dropdown = ctk.CTkOptionMenu(desc_header_frame, 
                                                   variable=self.template_var,
                                                   command=self._insert_template_description,
                                                   fg_color=CTK_COLOR_PANEL,
                                                   button_color=CTK_COLOR_PRIMARY,
                                                   button_hover_color=CTK_COLOR_SECONDARY,
                                                   text_color=CTK_COLOR_PRIMARY,
                                                   dropdown_fg_color=CTK_COLOR_PANEL)
        self.template_dropdown.grid(row=1, column=1, sticky="ew", padx=(10, 5), pady=5)
        self._update_template_dropdown() # Carrega opções iniciais
        
        self.description_textbox = ctk.CTkTextbox(master, height=150, corner_radius=10, fg_color=("#ECF0F1", "gray25"))
        self.description_textbox.grid(row=5, column=0, padx=25, pady=(0, 25), sticky="nsew")
        self.description_textbox.insert("1.0", self.estado['ultima_descricao'])

        # 5. Botão de Ação (Ajuste de Row)
        action_frame = ctk.CTkFrame(master, fg_color="transparent")
        action_frame.grid(row=7, column=0, padx=25, pady=10, sticky="ew")
        action_frame.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(action_frame, 
                      text="✅ Gerar Nota de Crédito (.xlsx)", 
                      command=self._process_note, 
                      fg_color=CTK_COLOR_BUTTON_GENERATE, # Cor mais clara para o botão principal
                      hover_color=CTK_COLOR_PRIMARY, # Hover primário
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

    def _format_date_input_on_focusout(self, event):
        """
        Formata o campo de data em DD/MM/AAAA APÓS o usuário sair do campo.
        """
        current = self.date_var.get()
        
        # 1. Limpa o valor para obter apenas os dígitos
        numeric = re.sub(r'[^0-9]', '', current)
        numeric = numeric[:8] # Limita a 8 dígitos (DDMMYYYY)

        new_value = ""
        
        # 2. Reconstroi a string formatada (DD/MM/AAAA)
        if len(numeric) == 8:
            new_value = f"{numeric[0:2]}/{numeric[2:4]}/{numeric[4:8]}"
        elif len(numeric) > 0:
            # Se for incompleto, tenta manter o máximo de formatação possível (ou deixa como está)
            if len(numeric) > 4:
                new_value = f"{numeric[0:2]}/{numeric[2:4]}/{numeric[4:]}"
            elif len(numeric) > 2:
                new_value = f"{numeric[0:2]}/{numeric[2:]}"
            else:
                new_value = numeric
        
        # 3. Atualiza a variável
        if new_value:
             # A validação final para o formato DD/MM/AAAA será feita no _process_note
            self.date_var.set(new_value)
        else:
             # Caso o usuário tenha apagado tudo, limpa o campo
            self.date_var.set("")


    def _format_currency_input_on_focusout(self, event):
        """
        Formata o campo de valor em R$ X.XXX,XX APÓS o usuário sair do campo.
        """
        current = self.value_var.get()
        
        # 1. Limpa o valor para obter apenas os dígitos
        cleaned_digits = re.sub(r'[^0-9]', '', current)
        
        if not cleaned_digits:
            self.value_var.set("0,00")
            return

        # 2. Converte para o valor inteiro de centavos
        value_cents = int(cleaned_digits)
        
        # 3. Formata o novo valor
        value_reais = value_cents / 100

        # Formatação para o padrão brasileiro X.XXX,XX
        # Usa formatação de ponto flutuante e substitui para a notação correta
        # (ponto como separador de milhar, vírgula como separador decimal)
        new_value = f"{value_reais:,.2f}".replace('.', '#').replace(',', '.').replace('#', ',')
        
        if new_value != current:
            self.value_var.set(new_value)

        
    # --- Funções de Processamento da Nota ---

    def _process_note(self):
        """Valida os dados e chama o backend para processar e salvar, e então pergunta sobre a impressão."""
        # Garante que os campos são formatados ANTES da validação final
        self._format_currency_input_on_focusout(None)
        self._format_date_input_on_focusout(None) 
        
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
            # Converte de formato brasileiro (vírgula decimal) para float (ponto decimal)
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
