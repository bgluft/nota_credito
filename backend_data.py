import os
import json
import re
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

# --- Configurações de Arquivo ---
CLIENTES_FILE = "clientes.json"
ESTADO_FILE = "estado.json"
MODELO_FILE = "modelo.xlsx"
SAIDA_FOLDER = "Notas_de_Credito_Geradas"

# --- Funções de Persistência (JSON) ---

def load_data(filename, default_data):
    """Carrega dados de um arquivo JSON. Se não existir, retorna dados padrão."""
    try:
        if os.path.exists(filename):
            with open(filename, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            return default_data
    except Exception as e:
        print(f"Erro ao carregar {filename}: {e}")
        return default_data

def save_data(filename, data):
    """Salva dados em um arquivo JSON."""
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Erro ao salvar {filename}: {e}")

# Gerenciamento de Clientes
def load_clientes():
    """Carrega a lista de clientes."""
    return load_data(CLIENTES_FILE, [])

def save_clientes(clientes):
    """Salva a lista de clientes."""
    save_data(CLIENTES_FILE, clientes)

# Gerenciamento de Estado (Fatura e Descrição)
def load_estado():
    """Carrega o estado do programa (última fatura e descrição)."""
    return load_data(ESTADO_FILE, {"ultima_fatura": 1, "ultima_descricao": "Devolução de Mercadoria conforme NFe [NUMERO NFE] de [DATA NFE]."})

def save_estado(estado):
    """Salva o estado atual do programa."""
    save_data(ESTADO_FILE, estado)

# --- Funções de Arquivo e Criação de Modelo ---

def create_initial_xlsx_template():
    """Cria um arquivo modelo XLSX mínimo para garantir a funcionalidade."""
    if os.path.exists(MODELO_FILE):
        return

    print(f"Criando {MODELO_FILE} inicial. Por favor, substitua-o pelo seu modelo real.")
    wb = Workbook()
    ws = wb.active
    ws.title = "Modelo_Base"

    # Define as células que o script irá preencher
    ws['H9'] = 'Data'
    ws['K9'] = 'Fatura N° (Topo)'
    ws['A13'] = 'Código Cliente (Topo)'
    ws['A15'] = 'Nome Cliente (Razão Social - Faturar)'
    ws['G15'] = 'Nome Cliente (Razão Social - Enviado)' 
    ws['B28'] = 'Descrição/Histórico'
    ws['K50'] = 'Valor R$'
    ws['J52'] = 'Código Cliente (Rodapé)'
    ws['L52'] = 'Fatura N° (Rodapé)'

    for cell in ['H9', 'K9', 'A13', 'A15', 'G15', 'B28', 'K50', 'J52', 'L52']:
        ws[cell].alignment = Alignment(wrapText=True)
        ws[cell].font = Font(bold=True)

    ws.column_dimensions[get_column_letter(1)].width = 25
    ws.column_dimensions[get_column_letter(2)].width = 35

    wb.save(MODELO_FILE)
    print(f"'{MODELO_FILE}' criado com sucesso.")

# Garante que a pasta de saída e o modelo existam
os.makedirs(SAIDA_FOLDER, exist_ok=True)
create_initial_xlsx_template()

def process_and_save_note(data_input, invoice_number, client_code, client_name, description_text, value_float, estado):
    """
    Carrega o modelo XLSX, preenche os dados, salva o novo arquivo e atualiza o estado.
    Retorna (True, output_path) em caso de sucesso ou (False, mensagem_de_erro).
    """
    try:
        wb = load_workbook(MODELO_FILE)
        ws = wb.active
    except FileNotFoundError:
        return False, f"O arquivo modelo '{MODELO_FILE}' não foi encontrado."
    except Exception as e:
        return False, f"Erro ao abrir o arquivo modelo: {e}"

    # --- GERAÇÃO DO NOME DA PLANILHA (ABA) ---
    name_parts = re.sub(r'[^\w\s]', '', client_name).split()
    sheet_name_base = '_'.join(name_parts[:2]).upper()
    new_sheet_title = f"{sheet_name_base}_{invoice_number}"
    
    if len(new_sheet_title) > 31:
        new_sheet_title = new_sheet_title[:31]

    try:
        ws.title = new_sheet_title
    except Exception as e:
        print(f"Aviso: Não foi possível renomear a planilha para '{new_sheet_title}': {e}. Usando o nome original.")


    # Mapeamento e Preenchimento das Células
    cell_map = {
        'H9': data_input, 'K9': invoice_number, 'L52': invoice_number,
        'A13': client_code, 'J52': client_code,
        'A15': client_name, 'G15': client_name, 
        'B28': description_text,
        'K50': value_float
    }

    try:
        for cell, value in cell_map.items():
            ws[cell] = value
            if cell == 'K50':
                ws[cell].number_format = 'R$ #,##0.00'
    except Exception as e:
        return False, f"Falha ao preencher a célula. Verifique se o '{MODELO_FILE}' está fechado e acessível.\nDetalhe: {e}"

    # 3. Salvar a Nova Nota
    file_name = f"Nota_Credito_{invoice_number}_Cliente_{client_code}.xlsx"
    output_path = os.path.join(SAIDA_FOLDER, file_name)

    try:
        wb.save(output_path)
    except Exception as e:
        return False, f"Não foi possível salvar a nota em '{output_path}'. Verifique as permissões.\nDetalhe: {e}"

    # 4. Atualizar o Estado do Programa
    try:
        next_invoice = int(invoice_number) + 1
        estado['ultima_fatura'] = next_invoice
        estado['ultima_descricao'] = description_text
        save_estado(estado)
    except Exception as e:
        print(f"Aviso: Falha ao atualizar o estado do programa: {e}")

    return True, output_path
