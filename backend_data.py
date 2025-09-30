import os
import json
import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import sys # Importação necessária para PyInstaller

# --- Utilitário de Caminho para PyInstaller ---
def _get_resource_path(relative_path):
    """Obtém o caminho absoluto do recurso, compatível com o PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        # Ambiente PyInstaller (executável)
        # O caminho do recurso é relativo ao diretório temporário do PyInstaller
        return os.path.join(sys._MEIPASS, relative_path)
    # Ambiente de desenvolvimento padrão
    return os.path.join(os.path.abspath("."), relative_path)
# ---------------------------------------------


# --- Constantes de Arquivos ---
CLIENTES_FILE = "clientes.json"
ESTADO_FILE = "estado.json"
MODELO_FILE = "modelo.xlsx" # Nome do arquivo que será empacotado
SAIDA_FOLDER = "Notas_de_Credito_Geradas"

# --- Dados Iniciais ---
INITIAL_ESTADO = {
    "ultima_fatura": 1,
    "ultima_descricao": "DESCONTO COMERCIAL REFERENTE A ACERTO COMERCIAL DE PRODUTOS.",
}

# --- Gerenciamento de Dados (JSON) ---

def load_clientes():
    """Carrega a lista de clientes do arquivo JSON."""
    if not os.path.exists(CLIENTES_FILE):
        return []
    try:
        with open(CLIENTES_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Erro ao carregar clientes: {e}")
        return []

def save_clientes(clientes):
    """Salva a lista de clientes no arquivo JSON."""
    try:
        with open(CLIENTES_FILE, 'w', encoding='utf-8') as f:
            json.dump(clientes, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Erro ao salvar clientes: {e}")

def load_estado():
    """Carrega o estado da última fatura e descrição do arquivo JSON."""
    if not os.path.exists(ESTADO_FILE):
        return INITIAL_ESTADO
    try:
        with open(ESTADO_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Erro ao carregar estado: {e}")
        return INITIAL_ESTADO

def save_estado(estado):
    """Salva o estado da última fatura e descrição no arquivo JSON."""
    try:
        with open(ESTADO_FILE, 'w', encoding='utf-8') as f:
            json.dump(estado, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Erro ao salvar estado: {e}")

def _create_initial_model():
    """Cria um arquivo modelo XLSX mínimo se não existir no ambiente de DEV."""
    # NÃO deve ser executado no ambiente PyInstaller, apenas em desenvolvimento.
    if not hasattr(sys, '_MEIPASS') and not os.path.exists(MODELO_FILE):
        try:
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "MODELO" 
            
            # Células de preenchimento obrigatórias
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
            ws.merge_cells('H9:J9')
            ws.merge_cells('K9:M9')
            ws.merge_cells('A15:F19')
            ws.merge_cells('G15:M19')
            ws.merge_cells('B28:F45')
            ws.merge_cells('K50:M50')
            
            wb.save(MODELO_FILE)
            print(f"AVISO: Arquivo modelo '{MODELO_FILE}' criado. SUBSTITUA este arquivo pelo seu modelo real.")
        except Exception as e:
            print(f"Erro ao criar modelo inicial: {e}")

# --- Processamento de XLSX ---

def process_and_save_note(data_input, invoice_number, client_code, client_name, description_text, value_float, estado):
    """
    Carrega o modelo (via PyInstaller ou ambiente local), preenche os dados e salva o novo arquivo XLSX.
    """
    
    # Obtém o caminho correto do modelo (PyInstaller ou local)
    model_path = _get_resource_path(MODELO_FILE)
    
    # Se estiver em desenvolvimento e o modelo não existir, cria o fallback
    if not os.path.exists(model_path):
        _create_initial_model()
        
    # Tenta obter o caminho novamente após a criação do fallback (só em dev)
    model_path = _get_resource_path(MODELO_FILE)

    if not os.path.exists(model_path):
         return False, f"Erro Fatal: O arquivo modelo '{MODELO_FILE}' não foi encontrado nem pôde ser criado."


    try:
        # Carrega o modelo usando o caminho obtido
        wb = load_workbook(model_path)
        ws = wb.active
        
        # 1. Preparação dos Nomes (para Planilha e Arquivo)
        
        # 1.1 Limpa o nome do cliente e pega as duas primeiras palavras
        cleaned_name = re.sub(r'[\\/?*\[\]\':]', '', client_name).strip()
        name_parts = [p for p in cleaned_name.split() if p] # Garante que as partes não sejam vazias
        
        # Pega as duas primeiras palavras, ou apenas a primeira se não houver duas
        if len(name_parts) >= 2:
            base_name = f"{name_parts[0]}_{name_parts[1]}"
        elif len(name_parts) == 1:
            base_name = name_parts[0]
        else:
            base_name = "CLIENTE_SEM_NOME"


        # 1.2 Renomeia a Planilha (Tab)
        sheet_name_raw = f"{base_name}_{invoice_number}"
        new_sheet_name = sheet_name_raw[:31].replace(' ', '_')
        ws.title = new_sheet_name

        # 2. Preenchimento de Células
        
        data_map = {
            'H9': data_input,                  # Data
            'K9': invoice_number,              # Número da Fatura (Topo)
            'A13': client_code,                # Código do Cliente
            'J52': client_code,                # Código do Cliente (Parte inferior)
            'A15': client_name,                # Nome/Razão Social (Faturar a)
            'G15': client_name,                # Nome/Razão Social (Enviado a)
            'B28': description_text,           # Descrição/Histórico
            'K50': value_float,                # Valor da Fatura (Número)
            'L52': invoice_number,             # Número da Fatura (Parte inferior)
        }

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

        # 3. Define o Caminho de Saída (NOVA REGRA DE NOME DE ARQUIVO)
        if not os.path.exists(SAIDA_FOLDER):
            os.makedirs(SAIDA_FOLDER)

        # Usando as duas primeiras palavras do cliente + número da nota
        output_filename = f"{base_name}_{invoice_number}.xlsx"
        output_path = os.path.join(SAIDA_FOLDER, output_filename)
        
        # 4. Salva o Arquivo
        wb.save(output_path)

        # 5. Atualiza o Estado (Próxima Fatura e Descrição)
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
