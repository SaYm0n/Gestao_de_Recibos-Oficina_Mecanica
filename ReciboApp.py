import sys
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QGroupBox, QLabel, QLineEdit, QTextEdit, QPushButton,
    QListWidget, QMessageBox, QFileDialog, QSizePolicy, QComboBox,
    QStyle,
    QScrollArea
)
from PyQt5.QtGui import QFont, QPainter, QPageLayout, QPageSize, QTextOption, QPixmap, QDoubleValidator, QIntValidator
from PyQt5.QtCore import Qt, QDateTime, QRectF, QSizeF, QPointF
import os
import requests
import subprocess
import platform
import base64
from PIL import Image
import tempfile
import atexit
import io

# --- IMPORTS PARA JINJA2 E WEASYPRINT ---
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML, CSS


# --- FIM DOS IMPORTS ---

# --- Configurações Globais ---
def resource_path(relative_path):
    """ Retorna o caminho absoluto para um recurso, funciona para dev e para PyInstaller """
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Determina o diretório base para arquivos de DADOS (Excel, PDFs gerados)
if getattr(sys, 'frozen', False):
    # Se estiver rodando como um executável PyInstaller (.exe)
    application_path = os.path.dirname(sys.executable)
else:
    # Se estiver rodando como um script .py normal
    application_path = os.path.dirname(os.path.abspath(__file__))

# Arquivo Excel para salvar dados do Recibo (sempre ao lado do .exe)
ARQUIVO_EXCEL_RECIBO = os.path.join(application_path, "Recibos_Historico.xlsx")
# Pasta para PDFs (sempre ao lado do .exe)
PASTA_RECIBOS_GERADOS = os.path.join(application_path, "Recibos_Gerados")

# Recursos internos da aplicação (imagens, templates)
ARQUIVO_LOGO = resource_path(os.path.join("resources", "logo.png"))
HTML_TEMPLATE_RECIBO = resource_path("recibo_template.html") # Ajustado para pegar da raiz do bundle

INFO_OFICINA = {
    "nome": "CR Soluções Automotivas",
    "endereco": "Estrada do barro vermelho 341 - Rocha Miranda - RJ",
    "cep": "21540-500",
    "telefone": "(21) 99757-0103 / 97125-0490",
    "email": "thiagosoarescruz01@gmail.com",
    "cnpj": "48.969.894/0001-59"
}

_temp_files_to_clean = []


def _cleanup_temp_files():
    for temp_file_path in _temp_files_to_clean:
        try:
            if os.path.exists(temp_file_path):
                os.remove(temp_file_path)
                print(f"DEBUG: Arquivo temporário removido: {temp_file_path}", file=sys.stderr)
        except Exception as e:
            print(f"ERRO: Não foi possível remover o arquivo temporário {temp_file_path}: {e}", file=sys.stderr)


atexit.register(_cleanup_temp_files)


class ReciboApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gerador de Recibos - Oficina")
        self.setGeometry(100, 100, 1200, 800)
        self.setMinimumSize(1000, 750)

        # CSS moderno para melhorar a aparência
        self.setStyleSheet("""
            QWidget {
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 9pt;
            }
            
            QGroupBox {
                font-weight: bold;
                border: 2px solid #3498db;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
                background-color: #f8f9fa;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 8px 0 8px;
                color: #2c3e50;
                background-color: #f8f9fa;
            }
            
            QLineEdit, QTextEdit, QComboBox {
                border: 2px solid #bdc3c7;
                border-radius: 5px;
                padding: 5px;
                background-color: white;
                selection-background-color: #3498db;
            }
            
            QLineEdit:focus, QTextEdit:focus, QComboBox:focus {
                border-color: #3498db;
                background-color: #f0f8ff;
            }
            
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-weight: bold;
                min-height: 20px;
            }
            
            QPushButton:hover {
                background-color: #2980b9;
            }
            
            QPushButton:pressed {
                background-color: #21618c;
            }
            
            QPushButton#btnSalvar {
                background-color: #27ae60;
            }
            
            QPushButton#btnSalvar:hover {
                background-color: #229954;
            }
            
            QPushButton#btnDeletar {
                background-color: #e74c3c;
            }
            
            QPushButton#btnDeletar:hover {
                background-color: #c0392b;
            }
            
            QPushButton#btnGerarPDF {
                background-color: #f39c12;
            }
            
            QPushButton#btnGerarPDF:hover {
                background-color: #e67e22;
            }
            
            QListWidget {
                border: 2px solid #bdc3c7;
                border-radius: 5px;
                background-color: white;
                alternate-background-color: #f8f9fa;
            }
            
            QListWidget::item {
                padding: 5px;
                border-bottom: 1px solid #ecf0f1;
            }
            
            QListWidget::item:selected {
                background-color: #3498db;
                color: white;
            }
            
            QLabel {
                color: #2c3e50;
            }
            
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            
            QScrollBar:vertical {
                background-color: #ecf0f1;
                width: 12px;
                border-radius: 6px;
            }
            
            QScrollBar::handle:vertical {
                background-color: #bdc3c7;
                border-radius: 6px;
                min-height: 20px;
            }
            
            QScrollBar::handle:vertical:hover {
                background-color: #95a5a6;
            }
        """)

        try:
            os.makedirs(PASTA_RECIBOS_GERADOS, exist_ok=True)
            print(f"Pasta '{PASTA_RECIBOS_GERADOS}' verificada/criada com sucesso.")
        except Exception as e:
            QMessageBox.critical(self, "Erro de Pasta",
                                 f"Não foi possível criar a pasta '{PASTA_RECIBOS_GERADOS}': {e}\nVerifique as permissões.")
            print(f"Erro ao criar pasta: {e}", file=sys.stderr)

        self.df_recibos = self._carregar_dados_recibos()
        self.itens_pecas_servicos_cache = []

        self.env = Environment(loader=FileSystemLoader(resource_path("resources")))
        self.env.filters['format_money'] = self._format_money_filter
        self.env.filters['km_format'] = self._km_format_filter
        self.env.filters['default_if_nan'] = self._default_if_nan_filter

        self._criar_interface()
        self._gerar_novo_id_recibo()

    def _format_money_filter(self, value):
        try:
            if pd.isna(value) or value is None:
                value = 0.0
            val = float(value)
            return f"{val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        except (ValueError, TypeError):
            return f"{0.00:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

    def _km_format_filter(self, value):
        try:
            if pd.isna(value) or value is None:
                value = ""
            clean_text = ''.join(filter(str.isdigit, str(value)))
            if clean_text:
                return f"{int(clean_text):,}".replace(',', '.')
            return ""
        except (ValueError, TypeError):
            return ""

    def _default_if_nan_filter(self, value):
        if pd.isna(value) or value is None:
            return ""
        return str(value)

    def _carregar_dados_recibos(self):
        excel_path = resource_path(ARQUIVO_EXCEL_RECIBO)
        if os.path.exists(excel_path):
            try:
                converters = {
                    'Numero_Recibo': str,
                    'KM_Atual_Veiculo': lambda x: int(str(x).replace('.', '').replace(',', '')) if str(x).replace('.',
                                                                                                                  '').replace(
                        ',', '').isdigit() else pd.NA,
                    'Valor_Total_Final': lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x,
                                                                                                                  str) and ',' in x else float(
                        x) if x is not None else pd.NA,
                    'Deslocamento': lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x,
                                                                                                             str) and ',' in x else float(
                        x) if x is not None else pd.NA,
                    'Desconto_Geral': lambda x: float(str(x).replace('.', '').replace(',', '.')) if isinstance(x,
                                                                                                               str) and ',' in x else float(
                        x) if x is not None else pd.NA,
                }
                df = pd.read_excel(excel_path, converters=converters)
                df['Numero_Recibo'] = df['Numero_Recibo'].astype(str).str.strip()
                print(f"Dados carregados de {excel_path}")
                for col in self._get_expected_columns():
                    if col not in df.columns:
                        df[col] = pd.NA
                return df
            except Exception as e:
                QMessageBox.critical(self, "Erro de Leitura",
                                     f"Erro ao carregar o arquivo Excel: {e}\nUm novo arquivo será criado.")
                return self._criar_dataframe_vazio_e_salvar()
        else:
            print(f"Arquivo {excel_path} não encontrado. Criando e inicializando novo DataFrame.")
            return self._criar_dataframe_vazio_e_salvar()

    def _get_expected_columns(self):
        return [
            "Numero_Recibo", "Data_Recibo", "Hora_Recibo",
            "Nome_Cliente",
            "Rua_Cliente", "Numero_Cliente", "Bairro_Cliente", "Cidade_Cliente", "UF_Cliente", 
            "CEP_Cliente", "Telefone_Cliente", "CPF_CNPJ_Cliente",
            "Placa_Veiculo", "Marca_Veiculo", "Modelo_Veiculo", "Cor_Veiculo", "Ano_Veiculo",
            "KM_Entrada_Veiculo", "KM_Saida_Veiculo",
            "Combustivel_Veiculo", "Box_Veiculo",
            "Problema_Informado", "Problema_Constatado", "Servico_Executado",
            "Detalhes_Itens", "Total_Itens",
            "Deslocamento", "Desconto_Geral", "Valor_Total_Final",
            "Responsavel", "Situacao_Atual", "Condicoes_Pagamento",
            "Email_Cliente", "Observacoes_Gerais", "Prox_Revisao",
            # Coluna antiga mantida para compatibilidade, mas não usada na UI nova
            "Endereco_Cliente"
        ]

    def _criar_dataframe_vazio_e_salvar(self):
        colunas = self._get_expected_columns()
        df = pd.DataFrame(columns=colunas)
        try:
            df.to_excel(resource_path(ARQUIVO_EXCEL_RECIBO), index=False)
            print(f"Arquivo Excel vazio '{ARQUIVO_EXCEL_RECIBO}' criado com sucesso.")
        except Exception as e:
            QMessageBox.critical(self, "Erro de Escritura",
                                 f"Não foi possível criar o arquivo Excel vazio: {e}. Verifique as permissões da pasta.")
        return df

    def _gerar_novo_id_recibo(self):
        if not self.df_recibos.empty:
            numeros_validos = self.df_recibos['Numero_Recibo'].astype(str).apply(
                lambda x: ''.join(filter(str.isdigit, x)))
            ultimos_numeros = [int(n) for n in numeros_validos if n]
            if ultimos_numeros:
                novo_id = max(ultimos_numeros) + 1
            else:
                novo_id = 1
        else:
            novo_id = 1
        self.entry_numero_recibo.setText(str(novo_id).zfill(6))
        self.entry_numero_recibo.setReadOnly(True)

    def _limpar_campos(self):
        for entry in self.findChildren(QLineEdit):
            entry.clear()
        for text_edit in self.findChildren(QTextEdit):
            text_edit.clear()

        self.combo_situacao_atual.setCurrentIndex(0)
        self.combo_condicoes_pagamento.setCurrentIndex(0)
        
        # Limpar comboboxes do veículo
        if isinstance(self.entries_veiculo["combustível"], QComboBox):
            self.entries_veiculo["combustível"].setCurrentIndex(0)
        if isinstance(self.entries_veiculo["box"], QComboBox):
            self.entries_veiculo["box"].setCurrentIndex(0)
            
        # Limpar combobox de tipo de item
        self.combo_item_tipo.setCurrentIndex(0)
        
        # Limpar lista de itens
        self.itens_pecas_servicos_cache = []
        self.listbox_itens.clear()
        
        # Atualizar totais
        self._atualizar_totais()
        
        # Gerar novo ID de recibo
        self._gerar_novo_id_recibo()

    def _criar_interface(self):
        main_layout = QVBoxLayout(self)

        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)

        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_widget.setMinimumSize(800, 1000)

        scroll_area.setWidget(content_widget)
        main_layout.addWidget(scroll_area, 1)

        # --- TOP SECTION: Logo, Info Oficina, Dados Recibo/Busca (Horizontal) ---
        header_layout = QHBoxLayout()
        content_layout.addLayout(header_layout)

        logo_info_layout = QVBoxLayout()
        header_layout.addLayout(logo_info_layout, 0)

        self.logo_label = QLabel()
        if os.path.exists(ARQUIVO_LOGO):
            pixmap = QPixmap(ARQUIVO_LOGO)
            if not pixmap.isNull():
                self.logo_label.setPixmap(pixmap.scaled(80, 80, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            else:
                self.logo_label.setText("LOGO")
        else:
            self.logo_label.setText("LOGO")
        logo_info_layout.addWidget(self.logo_label, alignment=Qt.AlignTop | Qt.AlignLeft)

        oficina_info_label = QLabel(
            f"<b>{INFO_OFICINA['nome']}</b><br>"
            f"{INFO_OFICINA['endereco']}<br>"
            f"CNPJ: {INFO_OFICINA['cnpj']} | Tel: {INFO_OFICINA['telefone']}"
        )
        oficina_info_label.setTextFormat(Qt.RichText)
        oficina_info_label.setFont(QFont("Arial", 8))
        logo_info_layout.addWidget(oficina_info_label, alignment=Qt.AlignTop | Qt.AlignLeft)
        logo_info_layout.addStretch(1)

        recibo_info_group = QGroupBox("Dados do Recibo")
        recibo_info_layout = QGridLayout()
        recibo_info_group.setLayout(recibo_info_layout)
        header_layout.addWidget(recibo_info_group, 1)

        recibo_info_layout.addWidget(QLabel("Número do Recibo:"), 0, 0, Qt.AlignLeft)
        self.entry_numero_recibo = QLineEdit()
        self.entry_numero_recibo.setFixedWidth(80)
        recibo_info_layout.addWidget(self.entry_numero_recibo, 0, 1, Qt.AlignLeft)
        recibo_info_layout.setColumnStretch(1, 0)

        recibo_info_layout.addWidget(QLabel("Data:"), 0, 2, Qt.AlignLeft)
        self.label_data = QLabel(QDateTime.currentDateTime().toString("dd/MM/yyyy hh:mm:ss"))
        recibo_info_layout.addWidget(self.label_data, 0, 3, Qt.AlignLeft)
        recibo_info_layout.setColumnStretch(3, 0)

        recibo_info_layout.addWidget(QLabel("Buscar Recibo por ID:"), 1, 0, Qt.AlignLeft)
        self.entry_busca_recibo = QLineEdit()
        self.entry_busca_recibo.setFixedWidth(80)
        recibo_info_layout.addWidget(self.entry_busca_recibo, 1, 1, Qt.AlignLeft)

        btn_buscar = QPushButton("Buscar")
        btn_buscar.clicked.connect(self._buscar_recibo)
        btn_buscar.setIcon(self.style().standardIcon(QStyle.SP_FileDialogToParent))
        recibo_info_layout.addWidget(btn_buscar, 1, 2, 1, 2, Qt.AlignLeft)

        # --- Layout Horizontal para Dados do Cliente e Dados do Veículo ---
        main_content_top_horizontal_layout = QHBoxLayout()
        content_layout.addLayout(main_content_top_horizontal_layout)
        # Correção: O stretch factor é aplicado no addWidget, não no layout em si.
        # main_content_top_horizontal_layout.setStretchFactor(0, 1)
        # main_content_top_horizontal_layout.setStretchFactor(1, 1)

        # --- Grupo Dados do Cliente (Grid com 3 colunas de campos) ---
        cliente_group = QGroupBox("👤 Dados do Cliente")
        cliente_layout = QGridLayout()
        cliente_group.setLayout(cliente_layout)
        main_content_top_horizontal_layout.addWidget(cliente_group, 1)  # Adiciona com stretch factor

        self.entries_cliente = {}
        client_field_positions = {
            # campo: (linha, coluna_label, col_span_entry)
            "Nome": (0, 0, 5),
            "Telefone": (1, 0, 1), "CPF/CNPJ": (1, 2, 1), "Email": (1, 4, 1),
            "CEP": (2, 0, 1), "Rua": (2, 2, 1), "Número": (2, 4, 1),
            "Bairro": (3, 0, 1), "Cidade": (3, 2, 1), "UF": (3, 4, 1)
        }
        campos_cliente_display_order = ["Nome", "Telefone", "CPF/CNPJ", "Email", "CEP", "Rua", "Número", "Bairro",
                                        "Cidade", "UF"]

        for campo_display_name in campos_cliente_display_order:
            row, col_start, col_span_entry = client_field_positions[campo_display_name]
            
            field_name_internal = campo_display_name.lower().replace('/', '_')

            cliente_layout.addWidget(QLabel(f"{campo_display_name}:"), row, col_start, Qt.AlignLeft)
            entry = QLineEdit()
            entry.setPlaceholderText(f"Digite o {campo_display_name.lower()}...")
            entry.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            
            # Ajuste de largura para campos específicos
            if field_name_internal in ["número", "uf"]:
                entry.setFixedWidth(80)
            elif field_name_internal == "cep":
                 entry.setFixedWidth(100)

            self.entries_cliente[field_name_internal] = entry
            cliente_layout.addWidget(entry, row, col_start + 1, 1, col_span_entry)

            if field_name_internal == "telefone":
                entry.setValidator(QIntValidator())
                entry.textChanged.connect(lambda text, e=entry: self._formatar_telefone_cpf_cnpj(e, "telefone"))
            elif field_name_internal == "cpf_cnpj":
                entry.setValidator(QIntValidator())
                entry.textChanged.connect(lambda text, e=entry: self._formatar_telefone_cpf_cnpj(e, "cpf_cnpj"))
            elif field_name_internal == "cep":
                entry.setValidator(QIntValidator())
                entry.editingFinished.connect(self._autopreencher_cep)
            elif field_name_internal == "número":
                entry.setValidator(QIntValidator())

        cliente_layout.setColumnStretch(1, 1)
        cliente_layout.setColumnStretch(3, 1)
        cliente_layout.setColumnStretch(5, 1)

        # Grupo Dados do Veículo
        veiculo_group = QGroupBox("🚗 Dados do Veículo")
        veiculo_layout = QGridLayout()
        veiculo_group.setLayout(veiculo_layout)
        main_content_top_horizontal_layout.addWidget(veiculo_group, 1)  # Adiciona com stretch factor

        self.entries_veiculo = {}
        vehicle_field_positions = {
            # campo: (linha, coluna_label)
            "Placa": (0, 0), "Ano": (0, 2),
            "Marca": (1, 0), "KM Entrada": (1, 2),
            "Modelo": (2, 0), "KM Saída": (2, 2),
            "Cor": (3, 0), "Combustível": (3, 2),
            "Box": (4, 0)
        }
        campos_veiculo_display_order = ["Placa", "Marca", "Modelo", "Cor", "Ano", "KM Entrada", "KM Saída", "Combustível", "Box"]

        for campo_display_name in campos_veiculo_display_order:
            row, col_start = vehicle_field_positions[campo_display_name]
            field_name_internal = campo_display_name.lower().replace(' ', '_')

            veiculo_layout.addWidget(QLabel(f"{campo_display_name}:"), row, col_start, Qt.AlignLeft)
            widget_col = col_start + 1

            if field_name_internal == "combustível":
                combo = QComboBox()
                combo.setPlaceholderText("Selecione...")
                combo.addItems(["", "Gasolina", "Etanol", "Flex", "Diesel", "GNV", "Elétrico", "Híbrido"])
                self.entries_veiculo[field_name_internal] = combo
                veiculo_layout.addWidget(combo, row, widget_col)
            elif field_name_internal == "box":
                combo = QComboBox()
                combo.setPlaceholderText("Selecione...")
                combo.addItems(["", "Box 1", "Box 2", "Box 3", "Box 4", "Pátio"])
                self.entries_veiculo[field_name_internal] = combo
                veiculo_layout.addWidget(combo, row, widget_col)
            else:
                entry = QLineEdit()
                entry.setPlaceholderText(f"Digite o {campo_display_name.lower()}...")
                entry.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
                entry.setMinimumWidth(100)
                self.entries_veiculo[field_name_internal] = entry
                veiculo_layout.addWidget(entry, row, widget_col)

                if field_name_internal == "ano":
                    entry.setValidator(QIntValidator(1900, QDateTime.currentDateTime().date().year() + 5))
                elif field_name_internal == "km_entrada" or field_name_internal == "km_saída":
                    entry.setValidator(QIntValidator(0, 999999999))
                    entry.textChanged.connect(lambda text, widget=entry: self._formatar_quilometragem(widget))

        veiculo_layout.setColumnStretch(1, 1)
        veiculo_layout.setColumnStretch(3, 1)

        # --- Grupo Itens (Peças e Serviços) - Adicionado Tipo, Código, Desc. (%) ---
        itens_group = QGroupBox("📋 Itens e Serviços")
        itens_layout = QGridLayout()
        itens_group.setLayout(itens_layout)
        content_layout.addWidget(itens_group)
        content_layout.setStretchFactor(itens_group, 2)

        itens_layout.addWidget(QLabel("Tipo:"), 0, 0, Qt.AlignLeft)
        self.combo_item_tipo = QComboBox()
        self.combo_item_tipo.setPlaceholderText("Selecione...")
        self.combo_item_tipo.addItems(["", "Peça", "Serviço"])
        itens_layout.addWidget(self.combo_item_tipo, 0, 1, Qt.AlignLeft)
        itens_layout.setColumnStretch(1, 0)

        itens_layout.addWidget(QLabel("Código:"), 0, 2, Qt.AlignLeft)
        self.entry_item_codigo = QLineEdit()
        self.entry_item_codigo.setPlaceholderText("Código/Ref")
        self.entry_item_codigo.setFixedWidth(80)
        itens_layout.addWidget(self.entry_item_codigo, 0, 3, Qt.AlignLeft)
        itens_layout.setColumnStretch(3, 0)

        itens_layout.addWidget(QLabel("Descrição:"), 0, 4, Qt.AlignLeft)
        self.entry_item_desc = QLineEdit()
        self.entry_item_desc.setPlaceholderText("Descrição do Item/Serviço")
        itens_layout.addWidget(self.entry_item_desc, 0, 5, Qt.AlignLeft)
        itens_layout.setColumnStretch(5, 1)

        itens_layout.addWidget(QLabel("Val Unit:"), 0, 6, Qt.AlignLeft)
        self.entry_item_valor = QLineEdit()
        self.entry_item_valor.setPlaceholderText("0,00")
        itens_layout.addWidget(self.entry_item_valor, 0, 7, Qt.AlignLeft)
        itens_layout.setColumnStretch(7, 0)

        itens_layout.addWidget(QLabel("Qtd:"), 0, 8, Qt.AlignLeft)
        self.entry_item_qtd = QLineEdit()
        self.entry_item_qtd.setPlaceholderText("1")
        itens_layout.addWidget(self.entry_item_qtd, 0, 9, Qt.AlignLeft)
        itens_layout.setColumnStretch(9, 0)

        itens_layout.addWidget(QLabel("Desc. (%):"), 0, 10, Qt.AlignLeft)
        self.entry_item_desc_perc = QLineEdit()
        self.entry_item_desc_perc.setPlaceholderText("0")
        self.entry_item_desc_perc.setValidator(QIntValidator(0, 100))
        itens_layout.addWidget(self.entry_item_desc_perc, 0, 11, Qt.AlignLeft)
        itens_layout.setColumnStretch(11, 0)

        btn_add_item = QPushButton("Adicionar Item")
        btn_add_item.clicked.connect(self._adicionar_item)
        btn_add_item.setIcon(self.style().standardIcon(QStyle.SP_DialogApplyButton))
        itens_layout.addWidget(btn_add_item, 0, 12, 1, 2, Qt.AlignLeft)
        itens_layout.setColumnStretch(12, 0)

        self.listbox_itens = QListWidget()
        itens_layout.addWidget(self.listbox_itens, 1, 0, 1, 14)
        itens_layout.setRowStretch(1, 1)
        btn_rem_item = QPushButton("Remover Item Selecionado")
        btn_rem_item.setIcon(self.style().standardIcon(QStyle.SP_DialogCancelButton))
        btn_rem_item.clicked.connect(self._remover_item)
        itens_layout.addWidget(btn_rem_item, 2, 0, 1, 14)

        # --- Grupo Problemas e Serviços - Não incluído no recibo padrão do cliente ---
        # Removido da parte horizontal superior para ficar em uma linha própria
        # Mantido para compatibilidade de dados no Excel, mas não visível por padrão
        # no recibo, a menos que você adicione de volta.

        # --- Seção de Totais e Finais ---
        finais_group = QGroupBox("💰 Totais e Finalizações")
        finais_layout = QGridLayout()
        finais_group.setLayout(finais_layout)
        content_layout.addWidget(finais_group)
        content_layout.setStretchFactor(finais_group, 0)

        self.entries_finais = {}

        finais_layout.addWidget(QLabel("Responsável:"), 0, 0, Qt.AlignLeft)
        self.entry_responsavel = QLineEdit()
        self.entry_responsavel.setPlaceholderText("Nome do responsável")
        self.entry_responsavel.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.entries_finais["responsável"] = self.entry_responsavel
        finais_layout.addWidget(self.entry_responsavel, 0, 1, Qt.AlignLeft)

        finais_layout.addWidget(QLabel("Observações:"), 1, 0, Qt.AlignLeft)
        self.text_observacoes = QTextEdit()
        self.text_observacoes.setPlaceholderText("Observações gerais...")
        self.text_observacoes.setMinimumHeight(40)
        self.entries_finais["observacoes"] = self.text_observacoes
        finais_layout.addWidget(self.text_observacoes, 1, 1, 2, 1, Qt.AlignLeft)

        finais_layout.addWidget(QLabel("Situação Atual:"), 0, 2, Qt.AlignLeft)
        self.combo_situacao_atual = QComboBox()
        self.combo_situacao_atual.setPlaceholderText("Selecione a situação")
        self.combo_situacao_atual.addItems(
            ["", "Orçamento", "Aprovado", "Em Andamento", "Aguardando Peças", "Finalizado", "Entregue"])
        finais_layout.addWidget(self.combo_situacao_atual, 0, 3, Qt.AlignLeft)
        self.entries_finais["situação_atual"] = self.combo_situacao_atual

        finais_layout.addWidget(QLabel("Condições de Pagamento:"), 1, 2, Qt.AlignLeft)
        self.combo_condicoes_pagamento = QComboBox()
        self.combo_condicoes_pagamento.setPlaceholderText("Selecione a condição")
        self.combo_condicoes_pagamento.addItems(
            ["", "À Vista", "PIX", "Cartão Crédito", "Cartão Débito", "Dinheiro", "Boleto", "Parcelado"])
        finais_layout.addWidget(self.combo_condicoes_pagamento, 1, 3, Qt.AlignLeft)
        self.entries_finais["condições_de_pagamento"] = self.combo_condicoes_pagamento

        finais_layout.addWidget(QLabel("Próxima Revisão:"), 2, 2, Qt.AlignLeft)
        self.entry_prox_revisao = QLineEdit()
        self.entry_prox_revisao.setPlaceholderText("Ex: 3 meses ou 10.000 KM")
        self.entries_finais["prox_revisao"] = self.entry_prox_revisao
        finais_layout.addWidget(self.entry_prox_revisao, 2, 3, Qt.AlignLeft)

        finais_layout.setColumnStretch(1, 1)
        finais_layout.setColumnStretch(3, 1)

        self.label_subtotal_itens = QLabel("Subtotal Itens: R$ 0,00")
        self.label_subtotal_itens.setFont(QFont("Arial", 9, QFont.Normal))
        finais_layout.addWidget(self.label_subtotal_itens, 2, 2, Qt.AlignRight)

        self.label_valor_total = QLabel("Valor Total: R$ 0,00")
        self.label_valor_total.setFont(QFont("Arial", 14, QFont.Bold))
        self.label_valor_total.setStyleSheet("color: #e74c3c; background-color: #fdf2f2; padding: 8px; border-radius: 5px; border: 2px solid #e74c3c;")
        finais_layout.addWidget(self.label_valor_total, 3, 2, Qt.AlignRight)
        finais_layout.setColumnStretch(2, 1)

        # --- Seção de Botões de Ação (fixa na parte inferior, fora do scroll) ---
        button_layout = QHBoxLayout()
        main_layout.addLayout(button_layout)
        main_layout.setStretchFactor(button_layout, 0)

        btn_salvar = QPushButton("Salvar Recibo")
        btn_salvar.clicked.connect(self._salvar_recibo)
        btn_salvar.setObjectName("btnSalvar")
        btn_salvar.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        button_layout.addWidget(btn_salvar)

        btn_limpar = QPushButton("Limpar Campos")
        btn_limpar.clicked.connect(self._limpar_campos)
        btn_limpar.setObjectName("btnLimpar")
        btn_limpar.setIcon(self.style().standardIcon(QStyle.SP_DialogResetButton))
        button_layout.addWidget(btn_limpar)

        btn_deletar = QPushButton("Deletar Recibo Atual")
        btn_deletar.clicked.connect(self._deletar_recibo)
        btn_deletar.setObjectName("btnDeletar")
        btn_deletar.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        button_layout.addWidget(btn_deletar)

        btn_imprimir = QPushButton("Gerar e Visualizar Recibo (PDF)")
        btn_imprimir.clicked.connect(self._imprimir_recibo_pdf)
        btn_imprimir.setObjectName("btnGerarPDF")
        btn_imprimir.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
        button_layout.addWidget(btn_imprimir)

        btn_sair = QPushButton("Sair")
        btn_sair.clicked.connect(self.close)
        btn_sair.setObjectName("btnSair")
        button_layout.addWidget(btn_sair)

    def _formatar_quilometragem(self, widget):
        widget.blockSignals(True)
        try:
            cursor_pos = widget.cursorPosition()
            original_text_len = len(widget.text())

            clean_text = ''.join(filter(str.isdigit, widget.text()))

            if not clean_text:
                widget.setText("")
            else:
                formatted_km = f"{int(clean_text):,}".replace(',', '.')
                widget.setText(formatted_km)
                new_text_len = len(formatted_km)
                len_diff = new_text_len - original_text_len
                widget.setCursorPosition(cursor_pos + len_diff)
        finally:
            widget.blockSignals(False)

    def _formatar_telefone_cpf_cnpj(self, entry_widget, field_type):
        current_text = entry_widget.text()
        clean_text = ''.join(filter(str.isdigit, current_text))
        formatted_text = ""
        cursor_pos = entry_widget.cursorPosition()
        len_diff = 0

        if field_type == "telefone":
            if len(clean_text) > 11:
                clean_text = clean_text[:11]

            if len(clean_text) > 2:
                formatted_text += f"({clean_text[:2]}) "
                if len(clean_text) > 7:
                    formatted_text += f"{clean_text[2:7]}-{clean_text[7:]}"
                else:
                    formatted_text += clean_text[2:]
            else:
                formatted_text = clean_text

            if len(current_text) < len(formatted_text):
                len_diff = len(formatted_text) - len(current_text)
            elif len(current_text) > len(formatted_text):
                len_diff = -(len(current_text) - len(formatted_text))

        elif field_type == "cpf_cnpj":
            if len(clean_text) > 14:
                clean_text = clean_text[:14]

            if len(clean_text) <= 11:  # CPF
                if len(clean_text) > 9:
                    formatted_text = f"{clean_text[:3]}.{clean_text[3:6]}.{clean_text[6:9]}-{clean_text[9:]}"
                elif len(clean_text) > 6:
                    formatted_text = f"{clean_text[:3]}.{clean_text[3:6]}.{clean_text[6:]}"
                elif len(clean_text) > 3:
                    formatted_text = f"{clean_text[:3]}.{clean_text[3:]}"
                else:
                    formatted_text = clean_text
            else:  # CNPJ
                if len(clean_text) > 12:
                    formatted_text = f"{clean_text[:2]}.{clean_text[2:5]}.{clean_text[5:8]}/{clean_text[8:12]}-{clean_text[12:]}"
                elif len(clean_text) > 8:
                    formatted_text = f"{clean_text[:2]}.{clean_text[2:5]}.{clean_text[5:8]}/{clean_text[8:]}"
                elif len(clean_text) > 5:
                    formatted_text = f"{clean_text[:2]}.{clean_text[2:5]}.{clean_text[5:]}"
                elif len(clean_text) > 2:
                    formatted_text = f"{clean_text[:2]}.{clean_text[2:]}"
                else:
                    formatted_text = clean_text

            if len(current_text) < len(formatted_text):
                len_diff = len(formatted_text) - len(current_text)
            elif len(current_text) > len(formatted_text):
                len_diff = -(len(current_text) - len(formatted_text))

        entry_widget.setText(formatted_text)
        if current_text != formatted_text:
            entry_widget.setCursorPosition(cursor_pos + len_diff)

    def _formatar_valor_monetario(self):
        sender = self.sender()
        text = sender.text().strip().replace(',', '.')
        try:
            value = float(text)
            formatted_value = f"{value:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            sender.setText(formatted_value)
        except ValueError:
            sender.setText(f"{0.00:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            QMessageBox.warning(self, "Formato Inválido", "Valor monetário inválido. Use apenas números.")

    def _adicionar_item(self):
        tipo = self.combo_item_tipo.currentText().strip()
        referencia = self.entry_item_codigo.text().strip()
        descricao = self.entry_item_desc.text().strip()
        valor_str = self.entry_item_valor.text().strip().replace(',', '.')
        qtd_str = self.entry_item_qtd.text().strip()
        desc_perc_str = self.entry_item_desc_perc.text().strip()

        if not tipo or not referencia or not descricao or not valor_str or not qtd_str:
            QMessageBox.warning(self, "Entrada Inválida", "Por favor, preencha todos os campos do item.")
            return

        try:
            valor_unitario = float(valor_str)
            quantidade = int(qtd_str)
            desconto_percentual = float(desc_perc_str) if desc_perc_str else 0.0

            if valor_unitario <= 0 or quantidade <= 0:
                QMessageBox.warning(self, "Entrada Inválida", "Valor unitário e quantidade devem ser maiores que zero.")
                return
            if not (0 <= desconto_percentual <= 100):
                QMessageBox.warning(self, "Entrada Inválida", "Desconto percentual deve estar entre 0 e 100.")
                return

        except ValueError:
            QMessageBox.warning(self, "Entrada Inválida", "Valores numéricos inválidos. Use apenas números.")
            return

        valor_total_item_sem_desc = valor_unitario * quantidade
        valor_total_item = valor_total_item_sem_desc * (1 - (desconto_percentual / 100))

        item_data = {
            "tipo": tipo,
            "codigo": referencia,
            "descricao": descricao,
            "uni": "un",
            "valor": valor_unitario,
            "quantia": quantidade,
            "desc": desconto_percentual,
            "valor_total": valor_total_item
        }
        self.itens_pecas_servicos_cache.append(item_data)
        self.listbox_itens.addItem(
            f"Tipo: {tipo} | Código: {referencia} - {descricao} | Qtd: {quantidade} x R${valor_unitario:.2f} | Desc: {desconto_percentual:.0f}% = R${valor_total_item:.2f}")

        self.combo_item_tipo.setCurrentIndex(0)
        self.entry_item_codigo.clear()
        self.entry_item_desc.clear()
        self.entry_item_valor.setText("0.00")
        self.entry_item_qtd.setText("1")
        self.entry_item_desc_perc.setText("0")

        self._atualizar_totais()

    def _remover_item(self):
        try:
            selected_row = self.listbox_itens.currentRow()
            if selected_row != -1:
                self.listbox_itens.takeItem(selected_row)
                del self.itens_pecas_servicos_cache[selected_row]
                self._atualizar_totais()
            else:
                QMessageBox.warning(self, "Seleção Inválida", "Por favor, selecione um item para remover.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Ocorreu um erro ao remover o item: {e}")

    def _atualizar_totais(self):
        try:
            subtotal_itens = sum(item['valor_total'] for item in self.itens_pecas_servicos_cache)
            
            # Agora o valor total é igual ao subtotal dos itens
            valor_total_final = subtotal_itens
            
            self.label_subtotal_itens.setText(f"Subtotal Itens: R$ {subtotal_itens:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            self.label_valor_total.setText(f"Valor Total: R$ {valor_total_final:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            
            # Atualizar o estilo do valor total para destacar
            if valor_total_final > 0:
                self.label_valor_total.setStyleSheet("color: #e74c3c; background-color: #fdf2f2; padding: 8px; border-radius: 5px; border: 2px solid #e74c3c; font-weight: bold;")
            else:
                self.label_valor_total.setStyleSheet("color: #95a5a6; background-color: #f8f9fa; padding: 8px; border-radius: 5px; border: 2px solid #bdc3c7;")
                
        except Exception as e:
            print(f"Erro ao atualizar totais: {e}", file=sys.stderr)
            self.label_subtotal_itens.setText("Subtotal Itens: R$ 0,00")
            self.label_valor_total.setText("Valor Total: R$ 0,00")

    def _autopreencher_cep(self):
        cep = self.entries_cliente["cep"].text().strip().replace('-', '')
        print(f"DEBUG: Autopreencher CEP chamado para: '{cep}'", file=sys.stderr)
        if len(cep) == 8 and cep.isdigit():
            url = f"https://viacep.com.br/ws/{cep}/json/"
            try:
                response = requests.get(url, timeout=5)
                response.raise_for_status()
                data = response.json()
                print(f"DEBUG: Resposta da ViaCEP: {data}", file=sys.stderr)

                if "erro" not in data:
                    self.entries_cliente["rua"].setText(data.get("logradouro", "") or "")
                    self.entries_cliente["bairro"].setText(data.get("bairro", "") or "")
                    self.entries_cliente["cidade"].setText(data.get("localidade", "") or "")
                    self.entries_cliente["uf"].setText(data.get("uf", "") or "")
                    self.entries_cliente["número"].setFocus() # Pula para o campo número
                else:
                    QMessageBox.warning(self, "CEP Inválido", "CEP não encontrado ou inválido.")
                    self.entries_cliente["rua"].clear()
                    self.entries_cliente["bairro"].clear()
                    self.entries_cliente["cidade"].clear()
                    self.entries_cliente["uf"].clear()
            except requests.exceptions.RequestException as e:
                QMessageBox.critical(self, "Erro de Conexão",
                                     f"Não foi possível consultar o CEP: {e}\nVerifique sua conexão com a internet.")
            except Exception as e:
                QMessageBox.critical(self, "Erro Inesperado", f"Ocorreu um erro ao autopreencher o CEP: {e}")
        elif len(cep) > 0 and (len(cep) != 8 or not cep.isdigit()):
            QMessageBox.warning(self, "CEP Inválido", "CEP deve conter 8 dígitos numéricos.")
            self.entries_cliente["rua"].clear()
            self.entries_cliente["bairro"].clear()
            self.entries_cliente["cidade"].clear()
            self.entries_cliente["uf"].clear()

    def _coletar_dados_form(self):
        # Coleta os campos de endereço separados
        rua = self.entries_cliente["rua"].text()
        numero = self.entries_cliente["número"].text()
        bairro = self.entries_cliente["bairro"].text()
        cidade = self.entries_cliente["cidade"].text()
        uf = self.entries_cliente["uf"].text()

        # Cria a string de endereço combinado para compatibilidade
        endereco_completo = ", ".join(filter(None, [rua, numero, bairro, cidade, uf]))

        dados = {
            "Numero_Recibo": self.entry_numero_recibo.text(),
            "Data_Recibo": self.label_data.text().split(' ')[0],
            "Hora_Recibo": self.label_data.text().split(' ')[1],
            "Nome_Cliente": self.entries_cliente["nome"].text(),
            "Telefone_Cliente": self.entries_cliente["telefone"].text(),
            "CPF_CNPJ_Cliente": self.entries_cliente["cpf_cnpj"].text(),
            "Email_Cliente": self.entries_cliente.get("email", QLineEdit()).text(),
            
            # Campos novos
            "Rua_Cliente": rua,
            "Numero_Cliente": numero,
            "Bairro_Cliente": bairro,
            "Cidade_Cliente": cidade,
            "UF_Cliente": uf,
            
            # Campo antigo para compatibilidade
            "Endereco_Cliente": endereco_completo,
            
            "CEP_Cliente": self.entries_cliente["cep"].text(),

            "Placa_Veiculo": self.entries_veiculo["placa"].text(),
            "Marca_Veiculo": self.entries_veiculo["marca"].text(),
            "Modelo_Veiculo": self.entries_veiculo["modelo"].text(),
            "Cor_Veiculo": self.entries_veiculo["cor"].text(),
            "Ano_Veiculo": self.entries_veiculo["ano"].text(),
            "KM_Entrada_Veiculo": self.entries_veiculo["km_entrada"].text().replace('.', ''),
            "KM_Saida_Veiculo": self.entries_veiculo["km_saída"].text().replace('.', ''),
            "Combustivel_Veiculo": self.entries_veiculo["combustível"].currentText() if isinstance(self.entries_veiculo["combustível"], QComboBox) else "",
            "Box_Veiculo": self.entries_veiculo["box"].currentText() if isinstance(self.entries_veiculo["box"], QComboBox) else "",

            "Problema_Informado": "",  # Não usado diretamente no recibo
            "Problema_Constatado": "",  # Não usado diretamente no recibo
            "Servico_Executado": "",  # Não usado diretamente no recibo

            "Detalhes_Itens": "; ".join([
                                            f"Tipo: {item['tipo']} | Código: {item['codigo']} | Descrição: {item['descricao']} | Quantia: {item['quantia']} | Valor Unit: {item['valor']:.2f} | Desc(%): {item['desc']:.0f} | Valor Total: {item['valor_total']:.2f}"
                                            for item in self.itens_pecas_servicos_cache]),
            "Total_Itens": sum(item['valor_total'] for item in self.itens_pecas_servicos_cache),
            "Deslocamento": 0.0,  # Removido da interface, mantido para compatibilidade
            "Desconto_Geral": 0.0,  # Removido da interface, mantido para compatibilidade
            "Valor_Total_Final": sum(item['valor_total'] for item in self.itens_pecas_servicos_cache),
            "Responsavel": self.entry_responsavel.text(),
            "Situacao_Atual": self.combo_situacao_atual.currentText(),
            "Condicoes_Pagamento": self.combo_condicoes_pagamento.currentText(),
            "Observacoes_Gerais": self.text_observacoes.toPlainText(),
            "Prox_Revisao": self.entry_prox_revisao.text()
        }
        dados["Valor_Total_Final"] = dados["Total_Itens"]  # Simplificado - apenas subtotal dos itens

        dados["Itens_Recibo"] = self.itens_pecas_servicos_cache

        return dados

    def _preencher_campos_form(self, dados_recibo_dict):
        self._limpar_campos()

        def get_display_value(key, default_value=""):
            value = dados_recibo_dict.get(key, default_value)
            if pd.isna(value) or value is None:
                return ""
            return str(value)

        self.entry_numero_recibo.setReadOnly(False)
        self.entry_numero_recibo.setText(get_display_value("Numero_Recibo"))
        self.entry_numero_recibo.setReadOnly(True)

        self.label_data.setText(f"{get_display_value('Data_Recibo')} {get_display_value('Hora_Recibo')}")

        self.entries_cliente["nome"].setText(get_display_value("Nome_Cliente"))
        self.entries_cliente["telefone"].setText(get_display_value("Telefone_Cliente"))
        self.entries_cliente["cpf_cnpj"].setText(get_display_value("CPF_CNPJ_Cliente"))
        if "email" in self.entries_cliente:
            self.entries_cliente["email"].setText(get_display_value("Email_Cliente"))
        self.entries_cliente["rua"].setText(get_display_value("Rua_Cliente"))
        self.entries_cliente["número"].setText(str(get_display_value("Numero_Cliente", get_display_value("Numero_Imovel_Cliente"))))
        self.entries_cliente["bairro"].setText(get_display_value("Bairro_Cliente"))
        self.entries_cliente["cidade"].setText(get_display_value("Cidade_Cliente"))
        self.entries_cliente["uf"].setText(get_display_value("UF_Cliente"))
        self.entries_cliente["cep"].setText(get_display_value("CEP_Cliente"))

        self.entries_veiculo["placa"].setText(get_display_value("Placa_Veiculo"))
        self.entries_veiculo["marca"].setText(get_display_value("Marca_Veiculo"))
        self.entries_veiculo["modelo"].setText(get_display_value("Modelo_Veiculo"))
        self.entries_veiculo["cor"].setText(get_display_value("Cor_Veiculo"))
        self.entries_veiculo["ano"].setText(get_display_value("Ano_Veiculo"))
        
        # Compatibilidade com a coluna antiga KM_Atual_Veiculo
        km_entrada = get_display_value("KM_Entrada_Veiculo", get_display_value("KM_Atual_Veiculo"))
        self.entries_veiculo["km_entrada"].setText(km_entrada)
        self.entries_veiculo["km_saída"].setText(get_display_value("KM_Saida_Veiculo"))

        if isinstance(self.entries_veiculo["combustível"], QComboBox):
            self.entries_veiculo["combustível"].setCurrentText(get_display_value("Combustivel_Veiculo"))
        if isinstance(self.entries_veiculo["box"], QComboBox):
            self.entries_veiculo["box"].setCurrentText(get_display_value("Box_Veiculo"))

        self.combo_situacao_atual.setCurrentText(get_display_value("Situacao_Atual"))
        self.combo_condicoes_pagamento.setCurrentText(get_display_value("Condicoes_Pagamento"))
        self.entry_responsavel.setText(get_display_value("Responsavel"))
        self.text_observacoes.setText(get_display_value("Observacoes_Gerais"))
        self.entry_prox_revisao.setText(get_display_value("Prox_Revisao"))

        self.itens_pecas_servicos_cache = []
        self.listbox_itens.clear()

        detalhes_itens_excel_str = get_display_value("Detalhes_Itens")
        if detalhes_itens_excel_str:
            for item_entry_str in detalhes_itens_excel_str.split('; '):
                if item_entry_str.strip():
                    try:
                        parts = {}
                        for part in item_entry_str.split(' | '):
                            if ': ' in part:
                                k, v = part.split(': ', 1)
                                parts[k.strip()] = v.strip()

                        item_data = {
                            "tipo": parts.get("Tipo", "N/A"),
                            "codigo": parts.get("Código", parts.get("Ref", "N/A")),
                            "descricao": parts.get("Descrição", parts.get("Desc", "N/A")),
                            "uni": parts.get("Uni", "un"),
                            "valor": float(parts.get("Valor Unit", "0.0").replace('R$', '').replace(',', '.').strip()),
                            "quantia": int(parts.get("Quantia", "0").strip()),
                            "desc": float(parts.get("Desc(%)", "0.0").replace('%', '').strip()),
                            "valor_total": float(
                                parts.get("Valor Total", "0.0").replace('R$', '').replace(',', '.').strip()),
                        }
                        self.itens_pecas_servicos_cache.append(item_data)
                        self.listbox_itens.addItem(
                            f"Tipo: {item_data['tipo']} | Código: {item_data['codigo']} - {item_data['descricao']} | Qtd: {item_data['quantia']} x R${item_data['valor']:.2f} | Desc: {item_data['desc']:.0f}% = R${item_data['valor_total']:.2f}"
                        )
                    except Exception as e:
                        print(f"Erro ao parsear item do Excel durante o carregamento: '{item_entry_str}' - {e}",
                              file=sys.stderr)
                        self.listbox_itens.addItem(item_entry_str)

        self._atualizar_totais()

    def _buscar_recibo(self):
        try:
            recibo_id_busca = self.entry_busca_recibo.text().strip()

            if not recibo_id_busca:
                QMessageBox.warning(self, "Campo Vazio", "Por favor, digite o número do Recibo para buscar.")
                return

            if recibo_id_busca.isdigit():
                recibo_id_busca = str(recibo_id_busca).zfill(6)
            print(f"DEBUG: Buscando Recibo com ID formatado: '{recibo_id_busca}'", file=sys.stderr)

            recibo_encontrado = self.df_recibos[self.df_recibos['Numero_Recibo'] == recibo_id_busca]

            if not recibo_encontrado.empty:
                dados_recibo_dict = recibo_encontrado.iloc[0].to_dict()
                self._preencher_campos_form(dados_recibo_dict)
                QMessageBox.information(self, "Recibo Encontrado", f"Recibo {recibo_id_busca} carregado com sucesso!")
                self.entry_busca_recibo.clear()
            else:
                QMessageBox.warning(self, "Recibo Não Encontrado",
                                    f"Recibo {recibo_id_busca} não encontrado no arquivo Excel.")
                self._limpar_campos()
        except Exception as e:
            QMessageBox.critical(self, "Erro na Busca", f"Erro ao buscar recibo: {str(e)}")
            print(f"Erro na busca de recibo: {e}", file=sys.stderr)

    def _salvar_recibo(self):
        try:
            dados_recibo_coletados = self._coletar_dados_form()

            # Validação mais robusta dos campos obrigatórios
            campos_obrigatorios = {
                "Nome_Cliente": "Nome do Cliente",
                "Valor_Total_Final": "Valor Total Final"
            }
            
            campos_vazios = []
            for campo, nome_exibicao in campos_obrigatorios.items():
                valor = dados_recibo_coletados.get(campo)
                if valor is None or (isinstance(valor, str) and not valor.strip()) or (isinstance(valor, (int, float)) and valor == 0):
                    if campo == "Valor_Total_Final" and str(valor) in ["0,00", "0.0", "0"]:
                        campos_vazios.append(nome_exibicao)
                    elif campo != "Valor_Total_Final":
                        campos_vazios.append(nome_exibicao)
            
            # Verificar se há pelo menos um item
            if not self.itens_pecas_servicos_cache:
                campos_vazios.append("Pelo menos um item/serviço")
            
            if campos_vazios:
                QMessageBox.warning(self, "Dados Incompletos",
                                    f"Os seguintes campos são obrigatórios:\n\n• " + "\n• ".join(campos_vazios))
                return

            dados_salvar = dados_recibo_coletados.copy()

            # Define colunas que devem ser tratadas como texto
            colunas_texto = [
                "Rua_Cliente", "Numero_Cliente", "Bairro_Cliente", "Cidade_Cliente", "UF_Cliente", 
                "CEP_Cliente", "Telefone_Cliente", "CPF_CNPJ_Cliente", "Nome_Cliente", 
                "Placa_Veiculo", "Marca_Veiculo", "Modelo_Veiculo", "Cor_Veiculo", "Ano_Veiculo"
            ]
            for col in colunas_texto:
                if col in dados_salvar:
                    dados_salvar[col] = str(dados_salvar[col]) if pd.notna(dados_salvar[col]) else ""

            # Conversão e limpeza de dados numéricos para salvar no Excel
            colunas_numericas = ["KM_Entrada_Veiculo", "KM_Saida_Veiculo", "Total_Itens", "Valor_Total_Final"]
            for key in colunas_numericas:
                val = dados_salvar.get(key)
                if isinstance(val, str):
                    clean_val = val.replace('.', '').replace(',', '.')
                    if clean_val.isdigit():
                        dados_salvar[key] = float(clean_val)
                    else:
                        dados_salvar[key] = pd.NA
                elif isinstance(val, (int, float)):
                    pass # já está no formato correto
                else:
                    dados_salvar[key] = pd.NA

            if str(dados_salvar["Numero_Recibo"]).isdigit():
                dados_salvar["Numero_Recibo"] = str(dados_salvar["Numero_Recibo"]).zfill(6)
            
            df_nova_recibo_linha = pd.DataFrame([dados_salvar])

            current_recibo_id = str(dados_recibo_coletados['Numero_Recibo'])
            recibo_existente_idx = self.df_recibos[self.df_recibos['Numero_Recibo'].astype(str) == current_recibo_id].index

            if not recibo_existente_idx.empty:
                idx = recibo_existente_idx[0]
                # Garante que todas as colunas existem no DataFrame antes de atribuir
                for col in df_nova_recibo_linha.columns:
                    if col not in self.df_recibos.columns:
                        self.df_recibos[col] = pd.NA
                    # Tenta converter o tipo da coluna se for incompatível
                    try:
                        self.df_recibos.at[idx, col] = df_nova_recibo_linha.at[0, col]
                    except (ValueError, TypeError):
                        self.df_recibos[col] = self.df_recibos[col].astype(object)
                        self.df_recibos.at[idx, col] = df_nova_recibo_linha.at[0, col]

                QMessageBox.information(self, "Recibo Atualizado", f"Recibo {current_recibo_id} atualizado com sucesso!")
            else:
                self.df_recibos = pd.concat([self.df_recibos, df_nova_recibo_linha], ignore_index=True)
                QMessageBox.information(self, "Recibo Salvo", f"Recibo {current_recibo_id} salvo com sucesso!")

            self.df_recibos.to_excel(resource_path(ARQUIVO_EXCEL_RECIBO), index=False)
            print(f"Dados salvos em {ARQUIVO_EXCEL_RECIBO}")

        except Exception as e:
            QMessageBox.critical(self, "Erro ao Salvar", f"Ocorreu um erro inesperado ao salvar o recibo:\n\n{e}")
            print(f"ERRO CRÍTICO ao salvar: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc()

    def _deletar_recibo(self):
        current_recibo_id = self.entry_numero_recibo.text().strip()

        id_to_delete = current_recibo_id if self.entry_numero_recibo.isReadOnly() and current_recibo_id else None

        if not id_to_delete:
            search_recibo_id = self.entry_busca_recibo.text().strip()
            if search_recibo_id:
                id_to_delete = search_recibo_id
            else:
                QMessageBox.warning(self, "Deletar Recibo", "Nenhum Recibo carregado ou ID de busca para deletar.")
                return

        if id_to_delete.isdigit():
            id_to_delete = str(id_to_delete).zfill(6)

        reply = QMessageBox.question(self, 'Deletar Recibo',
                                     f"Tem certeza que deseja deletar o Recibo {id_to_delete}?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            recibo_existente_idx = self.df_recibos[self.df_recibos['Numero_Recibo'].astype(str) == id_to_delete].index
            if not recibo_existente_idx.empty:
                self.df_recibos = self.df_recibos.drop(recibo_existente_idx).reset_index(drop=True)
                try:
                    self.df_recibos.to_excel(resource_path(ARQUIVO_EXCEL_RECIBO), index=False)
                    QMessageBox.information(self, "Recibo Deletado",
                                            f"Recibo {id_to_delete} deletado com sucesso do Excel!")
                    self._limpar_campos()
                except Exception as e:
                    QMessageBox.critical(self, "Erro ao Deletar", f"Não foi possível deletar o Recibo do Excel: {e}")
                    print(f"Detalles del error al eliminar Excel: {e}", file=sys.stderr)
            else:
                QMessageBox.warning(self, "Deletar Recibo", f"Recibo {id_to_delete} não encontrado para deletar.")

    def _imprimir_recibo_pdf(self):
        try:
            # Primeiro, validamos os dados do formulário sem salvar
            dados_recibo = self._coletar_dados_form()

            campos_obrigatorios_pdf = {
                "Numero_Recibo": "Número do Recibo",
                "Nome_Cliente": "Nome do Cliente",
                "Valor_Total_Final": "Valor Total Final"
            }
            campos_vazios_pdf = []
            for campo, nome_exibicao in campos_obrigatorios_pdf.items():
                valor = dados_recibo.get(campo)
                # Checa se o valor é None, string vazia, ou um número que é zero
                if valor is None or (isinstance(valor, str) and not valor.strip()) or (isinstance(valor, (int, float)) and valor == 0):
                    # Especial para Valor_Total_Final que pode ser string "0,00"
                    if campo == "Valor_Total_Final" and str(valor) in ["0,00", "0.0", "0"]:
                        campos_vazios_pdf.append(nome_exibicao)
                    elif campo != "Valor_Total_Final":
                         campos_vazios_pdf.append(nome_exibicao)

            if not self.itens_pecas_servicos_cache:
                campos_vazios_pdf.append("Pelo menos um item/serviço")

            if campos_vazios_pdf:
                QMessageBox.warning(self, "Dados Mínimos",
                                    f"Os seguintes campos são obrigatórios para gerar o PDF:\n\n• " + "\n• ".join(
                                        campos_vazios_pdf))
                return

            # Se os dados são válidos, agora sim salvamos
            self._salvar_recibo()

            # Gerar o PDF com os dados já coletados
            formatted_date_for_filename = QDateTime.currentDateTime().toString("yyyy-MM-dd")
            safe_recibo_number = "".join(c for c in dados_recibo['Numero_Recibo'] if c.isalnum() or c == '_')
            
            # Usar um arquivo temporário para evitar lixo
            temp_file = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False, prefix=f"recibo_{safe_recibo_number}_", dir=resource_path(PASTA_RECIBOS_GERADOS))
            filename_full_path = temp_file.name
            temp_file.close()
            
            atexit.register(lambda: os.remove(filename_full_path) if os.path.exists(filename_full_path) else None)

            logo_base64_data = None
            logo_absolute_path = resource_path("logo.png") # Simplificado para procurar na raiz ou na pasta de recursos
            if not os.path.exists(logo_absolute_path):
                 logo_absolute_path = resource_path(os.path.join("resources", "logo.png"))

            if os.path.exists(logo_absolute_path):
                with open(logo_absolute_path, "rb") as image_file:
                    logo_base64_data = base64.b64encode(image_file.read()).decode('utf-8')
            else:
                print(f"ALERTA: Arquivo de logo não encontrado em 'logo.png' ou 'resources/logo.png'", file=sys.stderr)

            template = self.env.get_template("recibo_template.html")
            html_content = template.render({
                'dados': dados_recibo,
                'info_oficina': INFO_OFICINA,
                'logo_base64': logo_base64_data,
                'data_atual': QDateTime.currentDateTime().toString("dd/MM/yyyy"),
                 'hora_atual': QDateTime.currentDateTime().toString("hh:mm:ss")
            })

            HTML(string=html_content, base_url=os.getcwd()).write_pdf(filename_full_path)
            
            QMessageBox.information(self, "PDF Gerado", f"Recibo gerado com sucesso!")

            if platform.system() == "Windows":
                os.startfile(filename_full_path)
            elif platform.system() == "Darwin":
                subprocess.run(["open", filename_full_path], check=True)
            else:
                subprocess.run(["xdg-open", filename_full_path], check=True)

        except Exception as e:
            QMessageBox.critical(self, "Erro na Geração do PDF",
                                 f"Ocorreu um erro inesperado ao gerar o PDF:\n\n{e}\n\nPor favor, verifique os dados e tente novamente.")
            print(f"ERRO CRÍTICO ao gerar PDF: {e}", file=sys.stderr)
            import traceback
            traceback.print_exc()


# --- Ejecución de la Aplicación ---
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ReciboApp()
    window.show()
    sys.exit(app.exec_())
