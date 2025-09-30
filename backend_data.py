import os
import json
import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import sys # Importação necessária para PyInstaller

# --- Utilitário de Caminho para PyInstaller ---
def _is_packed():
    """Verifica se o código está rodando como um executável PyInstaller."""
    return hasattr(sys, '_MEIPASS')

def _get_resource_path(relative_path):
    """Obtém o caminho absoluto do recurso, compatível com o PyInstaller."""
    if _is_packed():
        # Ambiente PyInstaller (executável)
        return os.path.join(sys._MEIPASS, relative_path)
    # Ambiente de desenvolvimento padrão
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)

# ---------------------------------------------

# --- Constantes de Arquivos ---
CLIENTES_FILE = "clientes.json"
ESTADO_FILE = "estado.json"
TEMPLATES_FILE = "templates.json"
FORNECEDORES_FILE = "fornecedores.json" # NOVO ARQUIVO DE DADOS
MODELO_FILE = "modelo.xlsx" 
MODELO2_FILE = "modelo2.xlsx" # NOVO MODELO
SAIDA_FOLDER = "Notas_de_Credito_Geradas"

# --- Dados Iniciais ---
INITIAL_ESTADO = {
    "ultima_fatura": 1,
    "ultima_descricao": "DESCONTO COMERCIAL REFERENTE a ACERTO COMERCIAL DE PRODUTOS.",
}
# ATUALIZAÇÃO: Inserindo os fornecedores pré-definidos
INITIAL_FORNECEDORES = [
    {"nome": "PRODUZA COMERCIO DE INSUMOS AGRÍCOLAS LTDA", "modelo": MODELO_FILE},
    {"nome": "BAYER SA", "modelo": MODELO2_FILE},
    {"nome": "DU PONT DO BRASIL SA", "modelo": MODELO2_FILE}
]

# --- Gerenciamento de Dados (JSON) ---

def _load_json_file(filename):
    """Função genérica para carregar dados JSON."""
    if not os.path.exists(filename):
        return []
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Erro ao carregar {filename}: {e}")
        return []

def _save_json_file(data, filename):
    """Função genérica para salvar dados JSON."""
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Erro ao salvar {filename}: {e}")

# Funções de Clientes (inalteradas na lógica)
def load_clientes():
    return _load_json_file(CLIENTES_FILE)
def save_clientes(clientes):
    _save_json_file(clientes, CLIENTES_FILE)

# Funções de Estado (inalteradas na lógica)
def load_estado():
    data = _load_json_file(ESTADO_FILE)
    if not data:
        return INITIAL_ESTADO
    return data
def save_estado(estado):
    _save_json_file(estado, ESTADO_FILE)

# Funções de Templates (inalteradas na lógica)
def load_templates():
    return _load_json_file(TEMPLATES_FILE)
def save_templates(templates):
    _save_json_file(templates, TEMPLATES_FILE)

# NOVAS Funções de Fornecedores
def load_fornecedores():
    """Carrega a lista de fornecedores do arquivo JSON."""
    data = _load_json_file(FORNECEDORES_FILE)
    if not data:
        # Se for o primeiro load, retorna a lista inicial para o usuário
        return INITIAL_FORNECEDORES 
    return data
def save_fornecedores(fornecedores):
    """Salva a lista de fornecedores no arquivo JSON."""
    _save_json_file(fornecedores, FORNECEDORES_FILE)


def _create_initial_model(filename):
    """Cria um arquivo modelo XLSX mínimo se não existir no ambiente de DEV."""
    # Apenas tenta criar se estiver em DEV e o arquivo não existir
    if not _is_packed() and not os.path.exists(filename):
        try:
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.title = filename.replace('.xlsx', '').upper()
            
            # Células de preenchimento obrigatórias
            ws['E2'] = 'NOME_FORNECEDOR' # Célula NOVA/ADICIONAL
            ws['H9'] = 'DATA'
            ws['K9'] = 'FATURA'
            ws['A13'] = 'COD_CLIENTE'
            ws['A15'] = 'NOME_CLIENTE_FATURAR_A'
            ws['G15'] = 'NOME_CLIENTE_ENVIADO_A'
            ws['B28'] = 'DESCRICAO' 
            ws['K50'] = 'VALOR' 
            ws['J52'] = 'COD_CLIENTE_FOOTER' 
            ws['L52'] = 'FATURA_FOOTER'
            
            # Simula mesclagens importantes
            ws.merge_cells('E2:J3') # Mesclagem NOVA
            ws.merge_cells('H9:J9')
            ws.merge_cells('K9:M9')
            ws.merge_cells('A15:F19')
            ws.merge_cells('G15:M19')
            ws.merge_cells('B28:F45')
            ws.merge_cells('K50:M50')
            
            wb.save(filename)
            print(f"AVISO: Arquivo modelo '{filename}' criado. SUBSTITUA este arquivo pelo seu modelo real.")
        except Exception as e:
            print(f"Erro ao criar modelo inicial ({filename}): {e}")

# --- Processamento de XLSX ---

def process_and_save_note(data_input, invoice_number, client_code, client_name, description_text, value_float, estado, model_filename, supplier_name):
    """
    Carrega o modelo, preenche os dados e salva o novo arquivo XLSX, 
    usando o modelo especificado.
    """
    
    # 1. Obter caminho do modelo (MODELO_FILE ou MODELO2_FILE)
    model_path = _get_resource_path(model_filename)
    
    # 1.1 Tenta criar o modelo se não for encontrado e estiver em ambiente DEV
    if not os.path.exists(model_path):
        print(f"DIAGNÓSTICO: Modelo '{model_filename}' não encontrado em {model_path}. Tentando criar fallback.")
        _create_initial_model(model_filename)
        model_path = _get_resource_path(model_filename) # Tenta obter o caminho novamente
        
    if not os.path.exists(model_path):
        return False, f"Erro Fatal: O arquivo modelo '{model_filename}' não foi encontrado nem pôde ser criado."

    try:
        # Carrega o modelo usando o caminho obtido
        wb = load_workbook(model_path)
        ws = wb.active
        
        # 2. Preparação dos Nomes (para Planilha e Arquivo)
        cleaned_name = re.sub(r'[\\/?*\[\]\':]', '', client_name).strip()
        name_parts = [p for p in cleaned_name.split() if p]
        
        if len(name_parts) >= 2:
            base_name = f"{name_parts[0]}_{name_parts[1]}"
        elif len(name_parts) == 1:
            base_name = name_parts[0]
        else:
            base_name = "CLIENTE_SEM_NOME"


        # 2.1 Renomeia a Planilha (Tab)
        sheet_name_raw = f"{base_name}_{invoice_number}"
        new_sheet_name = sheet_name_raw[:31].replace(' ', '_')
        ws.title = new_sheet_name

        # 3. Preenchimento de Células
        
        data_map = {
            'H9': data_input, 
            'K9': invoice_number, 
            'A13': client_code, 
            'J52': client_code, 
            'A15': client_name, 
            'G15': client_name, 
            'B28': description_text, 
            'K50': value_float, 
            'L52': invoice_number, 
        }
        
        # 3.1 Mapeamento Específico do Fornecedor/Modelo (NOVA LÓGICA)
        # O nome do fornecedor é mapeado para E2 apenas se for o modelo2
        if model_filename == "modelo2.xlsx":
             data_map['E2'] = supplier_name 
        # O modelo.xlsx não tem mapeamento especial, usa o PRODUZA.
        elif model_filename == "modelo.xlsx":
             # O mapeamento para E2 não é usado, mas a célula E2 no modelo.xlsx 
             # deve ser preenchida com o nome fixo se a célula não estiver mesclada.
             # Como o mapeamento é igual (exceto a célula extra), verificamos o 
             # mapeamento original do modelo.xlsx na imagem que tem a PRODUZA na linha 2.
             # Para ser fiel ao requisito:
             pass # A célula E2:J3 no modelo.xlsx é ignorada pelo script.

        for cell, value in data_map.items():
            try:
                ws[cell] = value
            except Exception:
                 pass # Ignora se a célula for o meio de uma mesclagem

        # Formata o valor como moeda
        try:
            ws['K50'].number_format = 'R$ #,##0.00'
        except:
            pass 

        # 4. Define o Caminho de Saída (NOVA REGRA DE NOME DE ARQUIVO)
        if not os.path.exists(SAIDA_FOLDER):
            os.makedirs(SAIDA_FOLDER)

        # Usando as duas primeiras palavras do cliente + número da nota
        output_filename = f"{base_name}_{invoice_number}.xlsx"
        output_path = os.path.join(SAIDA_FOLDER, output_filename)
        
        # 5. Salva o Arquivo
        wb.save(output_path)

        # 6. Atualiza o Estado (Próxima Fatura e Descrição)
        try:
            current_invoice = int(invoice_number)
            estado['ultima_fatura'] = current_invoice + 1
            estado['ultima_descricao'] = description_text
            save_estado(estado)
        except ValueError:
            pass 

        return True, output_path

    except Exception as e:
        return False, f"Erro ao processar o arquivo XLSX: {e}"
