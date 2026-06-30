"""
Comparador NF-e — Excel (C100) vs XMLs
Compara, para cada chave de acesso, os valores do Excel (vl_bc_icms, vl_icms, vl_doc)
com os valores das tags do XML (ICMSTot/vBC, ICMSTot/vICMS, vNF).
"""

import sys
import os
import xml.etree.ElementTree as ET
from pathlib import Path

import pandas as pd

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QTableWidget,
    QTableWidgetItem, QHeaderView, QGroupBox, QStatusBar,
    QAbstractItemView, QMessageBox, QFrame, QTabWidget
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont, QColor


NS = {"n": "http://www.portalfiscal.inf.br/nfe"}

COR_FUNDO    = "#1e1e2e"
COR_PAINEL   = "#2a2a3e"
COR_BORDA    = "#3a3a5e"
COR_PRIMARIA = "#7c3aed"
COR_HOVER    = "#6d28d9"
COR_OK       = "#22c55e"
COR_DIVERG   = "#ef4444"
COR_AUSENTE  = "#f59e0b"
COR_TEXTO    = "#e2e8f0"
COR_SUBTEXTO = "#94a3b8"

TOLERANCIA = 0.01


def parse_valor_br(valor):
    """Converte número para float, detectando automaticamente o formato.

    Aceita padrão BR (ponto=milhar, vírgula=decimal: "1.234,56")
    e padrão EN (ponto=decimal: "19.66" ou "1,234.56").
    """
    if valor is None:
        return None
    s = str(valor).strip()
    if not s or s.lower() == "nan":
        return None

    tem_ponto   = "." in s
    tem_virgula = "," in s

    if tem_virgula and tem_ponto:
        if s.rfind(",") > s.rfind("."):
            # BR: "1.234,56" — vírgula é decimal
            s = s.replace(".", "").replace(",", ".")
        else:
            # EN: "1,234.56" — ponto é decimal
            s = s.replace(",", "")
    elif tem_virgula:
        # Só vírgula → BR decimal: "19,66"
        s = s.replace(",", ".")
    # Só ponto ou nenhum → já é float: "19.66" ou "1966"

    try:
        return float(s)
    except ValueError:
        return None


class ComparadorWorker(QThread):
    resultado = pyqtSignal(list)
    erro      = pyqtSignal(str)

    def __init__(self, caminho_excel, pasta_xml):
        super().__init__()
        self.caminho_excel = caminho_excel
        self.pasta_xml = pasta_xml

    def run(self):
        try:
            df = pd.read_excel(self.caminho_excel, dtype=str)
            df.columns = [c.strip() for c in df.columns]

            obrigatorias = {"chv_nfe", "vl_bc_icms", "vl_icms", "vl_doc"}
            faltando_col = obrigatorias - set(df.columns)
            if faltando_col:
                raise ValueError(f"Colunas ausentes no Excel: {', '.join(sorted(faltando_col))}")

            pasta = Path(self.pasta_xml)
            xmls_disponiveis = {f.stem.strip(): f for f in pasta.iterdir() if f.suffix.lower() == ".xml"}

            linhas = []
            for _, row in df.iterrows():
                chave = str(row.get("chv_nfe", "")).strip()
                if not chave:
                    continue

                vl_bc_icms_xls = parse_valor_br(row.get("vl_bc_icms"))
                vl_icms_xls    = parse_valor_br(row.get("vl_icms"))
                vl_doc_xls     = parse_valor_br(row.get("vl_doc"))

                arq_xml = xmls_disponiveis.get(chave) or xmls_disponiveis.get("NFe" + chave)

                item = {
                    "chave": chave,
                    "num_doc": str(row.get("num_doc", "")).strip(),
                    "vendaunicoid": str(row.get("vendaUnicoId", row.get("vendaunicoid", ""))).strip(),
                    "vl_bc_icms_xls": vl_bc_icms_xls,
                    "vl_icms_xls": vl_icms_xls,
                    "vl_doc_xls": vl_doc_xls,
                    "vl_bc_icms_xml": None,
                    "vl_icms_xml": None,
                    "vl_doc_xml": None,
                    "xml_encontrado": arq_xml is not None,
                    "erro_xml": None,
                    "itens_xml": [],
                }

                if arq_xml is not None:
                    try:
                        vbc, vicms, vnf = self._extrair_valores_xml(arq_xml)
                        item["vl_bc_icms_xml"] = vbc
                        item["vl_icms_xml"] = vicms
                        item["vl_doc_xml"] = vnf
                        item["itens_xml"] = self._extrair_itens_xml(arq_xml)
                    except Exception as e:
                        item["erro_xml"] = str(e)

                item["diverge"] = self._calcular_divergencia(item)
                linhas.append(item)

            self.resultado.emit(linhas)
        except Exception as e:
            self.erro.emit(str(e))

    @staticmethod
    def _extrair_valores_xml(caminho_xml):
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        icms_tot = root.find(".//n:ICMSTot", NS)
        if icms_tot is None:
            raise ValueError("Tag ICMSTot não encontrada no XML")

        def texto(tag):
            el = icms_tot.find(f"n:{tag}", NS)
            return float(el.text) if el is not None and el.text else None

        vbc   = texto("vBC")
        vicms = texto("vICMS")
        vnf   = texto("vNF")
        return vbc, vicms, vnf

    @staticmethod
    def _extrair_itens_xml(caminho_xml):
        """Retorna lista de {cprod, cst, vbc, picms, vicms} para cada <det> do XML."""
        tree = ET.parse(caminho_xml)
        root = tree.getroot()
        itens = []
        for det in root.findall(".//n:det", NS):
            cprod = det.findtext("n:prod/n:cProd", default="", namespaces=NS)
            icms_el = det.find(".//n:ICMS", NS)
            cst = None
            vbc = "0"
            picms = "0"
            vicms = "0"
            if icms_el is not None:
                for child in icms_el:
                    cst_el   = child.find("n:CST",   NS)
                    vbc_el   = child.find("n:vBC",   NS)
                    picms_el = child.find("n:pICMS", NS)
                    vicms_el = child.find("n:vICMS", NS)
                    if cst_el   is not None: cst   = cst_el.text
                    if vbc_el   is not None: vbc   = vbc_el.text
                    if picms_el is not None: picms = picms_el.text
                    if vicms_el is not None: vicms = vicms_el.text
            itens.append({"cprod": cprod, "cst": cst, "vbc": vbc, "picms": picms, "vicms": vicms})
        return itens

    @staticmethod
    def _calcular_divergencia(item):
        if not item["xml_encontrado"] or item["erro_xml"]:
            return None

        pares = [
            (item["vl_bc_icms_xls"], item["vl_bc_icms_xml"]),
            (item["vl_icms_xls"], item["vl_icms_xml"]),
            (item["vl_doc_xls"], item["vl_doc_xml"]),
        ]
        for v_xls, v_xml in pares:
            if v_xls is None or v_xml is None:
                continue
            if abs(v_xls - v_xml) > TOLERANCIA:
                return True
        return False


class JanelaPrincipal(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Comparador NF-e — Excel (C100) vs XMLs")
        self.setMinimumSize(1280, 720)
        self.linhas = []
        self._aplicar_tema()
        self._construir_ui()

    def _aplicar_tema(self):
        self.setStyleSheet(f"""
            QMainWindow, QWidget {{
                background-color: {COR_FUNDO};
                color: {COR_TEXTO};
                font-family: 'Segoe UI', 'Inter', sans-serif;
                font-size: 13px;
            }}
            QGroupBox {{
                border: 1px solid {COR_BORDA};
                border-radius: 8px;
                margin-top: 12px;
                padding: 10px;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 12px;
                color: {COR_SUBTEXTO};
                font-size: 11px;
                text-transform: uppercase;
                letter-spacing: 1px;
            }}
            QLineEdit {{
                background-color: {COR_PAINEL};
                border: 1px solid {COR_BORDA};
                border-radius: 6px;
                padding: 6px 10px;
                color: {COR_TEXTO};
            }}
            QPushButton {{
                background-color: {COR_PRIMARIA};
                color: white;
                border: none;
                border-radius: 6px;
                padding: 7px 18px;
                font-weight: 600;
            }}
            QPushButton:hover  {{ background-color: {COR_HOVER}; }}
            QPushButton:disabled {{ background-color: {COR_BORDA}; color: {COR_SUBTEXTO}; }}
            QPushButton#secundario {{
                background-color: {COR_PAINEL};
                border: 1px solid {COR_BORDA};
                color: {COR_TEXTO};
            }}
            QPushButton#secundario:hover {{ background-color: {COR_BORDA}; }}
            QTableWidget {{
                background-color: {COR_PAINEL};
                border: 1px solid {COR_BORDA};
                border-radius: 6px;
                gridline-color: {COR_BORDA};
                outline: none;
            }}
            QTableWidget::item {{ padding: 6px 8px; }}
            QTableWidget::item:selected {{
                background-color: {COR_PRIMARIA};
                color: white;
            }}
            QHeaderView::section {{
                background-color: {COR_FUNDO};
                color: {COR_SUBTEXTO};
                border: none;
                border-bottom: 1px solid {COR_BORDA};
                padding: 8px 8px;
                font-size: 11px;
                text-transform: uppercase;
                letter-spacing: 0.5px;
            }}
            QStatusBar {{ background-color: {COR_FUNDO}; color: {COR_SUBTEXTO}; }}
            QScrollBar:vertical {{
                background: {COR_FUNDO}; width: 8px; border-radius: 4px;
            }}
            QScrollBar::handle:vertical {{
                background: {COR_BORDA}; border-radius: 4px;
            }}
        """)

    def _construir_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(20, 20, 20, 10)
        layout.setSpacing(14)

        titulo = QLabel("🔍  Comparador NF-e — Excel vs XML / SEFAZ")
        titulo.setFont(QFont("Segoe UI", 18, QFont.Bold))
        layout.addWidget(titulo)

        subtitulo = QLabel("Compara valores do Excel (C100) com XMLs e com extrato do SEFAZ")
        subtitulo.setStyleSheet(f"color: {COR_SUBTEXTO}; font-size: 12px;")
        layout.addWidget(subtitulo)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet(f"color: {COR_BORDA};")
        layout.addWidget(sep)

        # ── Abas ──────────────────────────────────────────────────────────────
        self.abas = QTabWidget()
        self.abas.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid {COR_BORDA};
                border-radius: 6px;
                background: {COR_PAINEL};
            }}
            QTabBar::tab {{
                background: {COR_FUNDO};
                color: {COR_SUBTEXTO};
                border: 1px solid {COR_BORDA};
                border-bottom: none;
                padding: 7px 18px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                font-size: 12px;
            }}
            QTabBar::tab:selected {{
                background: {COR_PAINEL};
                color: {COR_TEXTO};
                font-weight: bold;
            }}
            QTabBar::tab:hover:!selected {{ background: {COR_BORDA}; }}
        """)
        layout.addWidget(self.abas, stretch=1)

        # ── ABA 1: Excel vs XML ───────────────────────────────────────────────
        aba_xml = QWidget()
        lay_xml = QVBoxLayout(aba_xml)
        lay_xml.setContentsMargins(12, 12, 12, 8)
        lay_xml.setSpacing(10)

        grp_entrada = QGroupBox("Arquivos de entrada")
        grp_layout = QVBoxLayout(grp_entrada)
        grp_layout.setSpacing(10)

        row_excel = QHBoxLayout()
        lbl_excel = QLabel("Planilha Excel (C100):")
        lbl_excel.setFixedWidth(180)
        self.campo_excel = QLineEdit()
        self.campo_excel.setPlaceholderText("Caminho do arquivo .xlsx…")
        btn_excel = QPushButton("Procurar")
        btn_excel.setObjectName("secundario")
        btn_excel.setFixedWidth(90)
        btn_excel.clicked.connect(self._selecionar_excel)
        row_excel.addWidget(lbl_excel)
        row_excel.addWidget(self.campo_excel)
        row_excel.addWidget(btn_excel)
        grp_layout.addLayout(row_excel)

        row_pasta = QHBoxLayout()
        lbl_pasta = QLabel("Pasta com XMLs:")
        lbl_pasta.setFixedWidth(180)
        self.campo_pasta = QLineEdit()
        self.campo_pasta.setPlaceholderText("Pasta onde estão os arquivos .xml…")
        btn_pasta = QPushButton("Procurar")
        btn_pasta.setObjectName("secundario")
        btn_pasta.setFixedWidth(90)
        btn_pasta.clicked.connect(self._selecionar_pasta)
        row_pasta.addWidget(lbl_pasta)
        row_pasta.addWidget(self.campo_pasta)
        row_pasta.addWidget(btn_pasta)
        grp_layout.addLayout(row_pasta)

        lay_xml.addWidget(grp_entrada)

        row_botoes = QHBoxLayout()
        self.btn_comparar = QPushButton("▶  Comparar")
        self.btn_comparar.setFixedHeight(38)
        self.btn_comparar.clicked.connect(self._comparar)
        btn_limpar = QPushButton("✕  Limpar")
        btn_limpar.setObjectName("secundario")
        btn_limpar.setFixedHeight(38)
        btn_limpar.clicked.connect(self._limpar)
        btn_exportar = QPushButton("⬇  Exportar divergências")
        btn_exportar.setObjectName("secundario")
        btn_exportar.setFixedHeight(38)
        btn_exportar.clicked.connect(self._exportar)
        btn_sql = QPushButton("⚙  Gerar SQL")
        btn_sql.setObjectName("secundario")
        btn_sql.setFixedHeight(38)
        btn_sql.clicked.connect(self._gerar_sql)
        row_botoes.addWidget(self.btn_comparar)
        row_botoes.addWidget(btn_limpar)
        row_botoes.addStretch()
        row_botoes.addWidget(btn_exportar)
        row_botoes.addWidget(btn_sql)
        lay_xml.addLayout(row_botoes)

        self.lbl_resumo = QLabel("")
        self.lbl_resumo.setTextFormat(Qt.RichText)
        self.lbl_resumo.setStyleSheet(f"color: {COR_SUBTEXTO}; font-size: 12px;")
        lay_xml.addWidget(self.lbl_resumo)

        self.tabela = QTableWidget()
        colunas = [
            "Status", "Nº Doc", "BC ICMS (Excel)", "BC ICMS (XML)",
            "ICMS (Excel)", "ICMS (XML)", "Vl. Doc (Excel)", "Vl. NF (XML)", "Chave de Acesso"
        ]
        self.tabela.setColumnCount(len(colunas))
        self.tabela.setHorizontalHeaderLabels(colunas)
        self.tabela.horizontalHeader().setSectionResizeMode(8, QHeaderView.Stretch)
        self.tabela.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tabela.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tabela.verticalHeader().setVisible(False)
        lay_xml.addWidget(self.tabela, stretch=1)

        self.abas.addTab(aba_xml, "📄  Excel vs XML")

        # ── ABA 2: Excel vs SEFAZ ─────────────────────────────────────────────
        aba_sefaz = QWidget()
        lay_sefaz = QVBoxLayout(aba_sefaz)
        lay_sefaz.setContentsMargins(12, 12, 12, 8)
        lay_sefaz.setSpacing(10)

        grp_sefaz = QGroupBox("Arquivos de entrada — SEFAZ")
        grp_sefaz_lay = QVBoxLayout(grp_sefaz)
        grp_sefaz_lay.setSpacing(10)

        row_c100 = QHBoxLayout()
        lbl_c100 = QLabel("Planilha Excel (C100):")
        lbl_c100.setFixedWidth(180)
        self.campo_sefaz_excel = QLineEdit()
        self.campo_sefaz_excel.setPlaceholderText("Mesmo arquivo C100 com vl_bc_icms / vl_icms…")
        btn_c100 = QPushButton("Procurar")
        btn_c100.setObjectName("secundario")
        btn_c100.setFixedWidth(90)
        btn_c100.clicked.connect(self._selecionar_sefaz_excel)
        row_c100.addWidget(lbl_c100)
        row_c100.addWidget(self.campo_sefaz_excel)
        row_c100.addWidget(btn_c100)
        grp_sefaz_lay.addLayout(row_c100)

        row_sefaz = QHBoxLayout()
        lbl_sefaz = QLabel("Extrato SEFAZ (.xlsx):")
        lbl_sefaz.setFixedWidth(180)
        self.campo_sefaz_arq = QLineEdit()
        self.campo_sefaz_arq.setPlaceholderText("Arquivo exportado do SEFAZ com Total_ICMS / Total_BC_ICMS…")
        btn_sefaz = QPushButton("Procurar")
        btn_sefaz.setObjectName("secundario")
        btn_sefaz.setFixedWidth(90)
        btn_sefaz.clicked.connect(self._selecionar_sefaz_arq)
        row_sefaz.addWidget(lbl_sefaz)
        row_sefaz.addWidget(self.campo_sefaz_arq)
        row_sefaz.addWidget(btn_sefaz)
        grp_sefaz_lay.addLayout(row_sefaz)

        lay_sefaz.addWidget(grp_sefaz)

        row_bot_sefaz = QHBoxLayout()
        self.btn_comparar_sefaz = QPushButton("▶  Comparar com SEFAZ")
        self.btn_comparar_sefaz.setFixedHeight(38)
        self.btn_comparar_sefaz.clicked.connect(self._comparar_sefaz)
        btn_limpar_sefaz = QPushButton("✕  Limpar")
        btn_limpar_sefaz.setObjectName("secundario")
        btn_limpar_sefaz.setFixedHeight(38)
        btn_limpar_sefaz.clicked.connect(self._limpar_sefaz)
        btn_export_sefaz = QPushButton("⬇  Exportar divergências")
        btn_export_sefaz.setObjectName("secundario")
        btn_export_sefaz.setFixedHeight(38)
        btn_export_sefaz.clicked.connect(self._exportar_sefaz)
        row_bot_sefaz.addWidget(self.btn_comparar_sefaz)
        row_bot_sefaz.addWidget(btn_limpar_sefaz)
        row_bot_sefaz.addStretch()
        row_bot_sefaz.addWidget(btn_export_sefaz)
        lay_sefaz.addLayout(row_bot_sefaz)

        self.lbl_resumo_sefaz = QLabel("")
        self.lbl_resumo_sefaz.setTextFormat(Qt.RichText)
        self.lbl_resumo_sefaz.setStyleSheet(f"color: {COR_SUBTEXTO}; font-size: 12px;")
        lay_sefaz.addWidget(self.lbl_resumo_sefaz)

        # Tabela SEFAZ
        self.tabela_sefaz = QTableWidget()
        colunas_sefaz = [
            "Status", "Nº Doc", "Chave de Acesso",
            "BC ICMS (Excel)", "BC ICMS (SEFAZ)", "Dif. BC ICMS",
            "ICMS (Excel)",    "ICMS (SEFAZ)",    "Dif. ICMS",
            "Vl. Doc (Excel)", "Total NF-e (SEFAZ)", "Dif. Vl. Doc",
        ]
        self.tabela_sefaz.setColumnCount(len(colunas_sefaz))
        self.tabela_sefaz.setHorizontalHeaderLabels(colunas_sefaz)
        self.tabela_sefaz.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.tabela_sefaz.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tabela_sefaz.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tabela_sefaz.verticalHeader().setVisible(False)
        lay_sefaz.addWidget(self.tabela_sefaz, stretch=1)

        # Rodapé SEFAZ
        self.tabela_rodape_sefaz = QTableWidget()
        self.tabela_rodape_sefaz.setColumnCount(len(colunas_sefaz))
        self.tabela_rodape_sefaz.setRowCount(1)
        self.tabela_rodape_sefaz.horizontalHeader().setVisible(False)
        self.tabela_rodape_sefaz.verticalHeader().setVisible(False)
        self.tabela_rodape_sefaz.setFixedHeight(38)
        self.tabela_rodape_sefaz.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tabela_rodape_sefaz.setSelectionMode(QAbstractItemView.NoSelection)
        self.tabela_rodape_sefaz.setFocusPolicy(Qt.NoFocus)
        self.tabela_rodape_sefaz.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.tabela_rodape_sefaz.setStyleSheet(f"""
            QTableWidget {{
                background-color: {COR_FUNDO};
                border: none;
                border-top: 2px solid {COR_BORDA};
                gridline-color: {COR_BORDA};
            }}
            QTableWidget::item {{ padding: 6px 8px; font-weight: bold; }}
        """)
        self.tabela_rodape_sefaz.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.tabela_sefaz.horizontalScrollBar().valueChanged.connect(
            self.tabela_rodape_sefaz.horizontalScrollBar().setValue
        )
        self.tabela_sefaz.horizontalHeader().sectionResized.connect(
            lambda col, _, new_size: self.tabela_rodape_sefaz.setColumnWidth(col, new_size)
        )
        lay_sefaz.addWidget(self.tabela_rodape_sefaz)
        self._limpar_rodape_sefaz()

        self.linhas_sefaz = []

        # ── Seção 2: Notas no SEFAZ que NÃO estão no SPED ───────────────────
        sep2 = QFrame()
        sep2.setFrameShape(QFrame.HLine)
        sep2.setStyleSheet(f"color: {COR_BORDA}; margin-top: 6px;")
        lay_sefaz.addWidget(sep2)

        lbl_sec2 = QLabel("⚠  Notas no SEFAZ não encontradas no SPED (C100)")
        lbl_sec2.setFont(QFont("Segoe UI", 11, QFont.Bold))
        lbl_sec2.setStyleSheet(f"color: {COR_AUSENTE}; margin: 4px 0;")
        lay_sefaz.addWidget(lbl_sec2)

        self.lbl_resumo_sefaz2 = QLabel("")
        self.lbl_resumo_sefaz2.setTextFormat(Qt.RichText)
        self.lbl_resumo_sefaz2.setStyleSheet(f"color: {COR_SUBTEXTO}; font-size: 12px;")
        lay_sefaz.addWidget(self.lbl_resumo_sefaz2)

        colunas_sefaz2 = [
            "Nº Doc", "Chave de Acesso (SEFAZ)",
            "Total NF-e", "Total BC ICMS", "Total ICMS",
            "Data Emissão", "Emitente",
        ]
        self.tabela_sefaz2 = QTableWidget()
        self.tabela_sefaz2.setColumnCount(len(colunas_sefaz2))
        self.tabela_sefaz2.setHorizontalHeaderLabels(colunas_sefaz2)
        self.tabela_sefaz2.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.tabela_sefaz2.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tabela_sefaz2.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tabela_sefaz2.verticalHeader().setVisible(False)
        self.tabela_sefaz2.setMaximumHeight(220)
        lay_sefaz.addWidget(self.tabela_sefaz2)

        row_bot2 = QHBoxLayout()
        btn_export2 = QPushButton("⬇  Exportar ausentes no SPED")
        btn_export2.setObjectName("secundario")
        btn_export2.setFixedHeight(34)
        btn_export2.clicked.connect(self._exportar_sefaz2)
        row_bot2.addStretch()
        row_bot2.addWidget(btn_export2)
        lay_sefaz.addLayout(row_bot2)

        self.linhas_sefaz2 = []
        self.abas.addTab(aba_sefaz, "🏛  Excel vs SEFAZ")

        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("Pronto.")

    # ── Métodos aba SEFAZ ─────────────────────────────────────────────────────

    def _selecionar_sefaz_excel(self):
        caminho, _ = QFileDialog.getOpenFileName(
            self, "Selecionar planilha C100", "",
            "Excel (*.xlsx *.xls);;Todos os arquivos (*)"
        )
        if caminho:
            self.campo_sefaz_excel.setText(caminho)

    def _selecionar_sefaz_arq(self):
        caminho, _ = QFileDialog.getOpenFileName(
            self, "Selecionar extrato SEFAZ", "",
            "Excel (*.xlsx *.xls);;Todos os arquivos (*)"
        )
        if caminho:
            self.campo_sefaz_arq.setText(caminho)

    def _comparar_sefaz(self):
        arq_c100  = self.campo_sefaz_excel.text().strip()
        arq_sefaz = self.campo_sefaz_arq.text().strip()

        if not arq_c100 or not os.path.isfile(arq_c100):
            QMessageBox.warning(self, "Atenção", "Selecione a planilha Excel (C100) válida.")
            return
        if not arq_sefaz or not os.path.isfile(arq_sefaz):
            QMessageBox.warning(self, "Atenção", "Selecione o arquivo do extrato SEFAZ válido.")
            return

        self.btn_comparar_sefaz.setEnabled(False)
        self.status.showMessage("Processando comparação SEFAZ…")
        try:
            self._processar_sefaz(arq_c100, arq_sefaz)
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao processar:\n\n{e}")
            self.status.showMessage("Erro no processamento SEFAZ.")
        finally:
            self.btn_comparar_sefaz.setEnabled(True)

    def _processar_sefaz(self, arq_c100, arq_sefaz):
        # Lê C100
        df_c100 = pd.read_excel(arq_c100, dtype=str)
        df_c100.columns = [c.strip().lower() for c in df_c100.columns]
        for col in ("chv_nfe", "vl_bc_icms", "vl_icms", "vl_doc"):
            if col not in df_c100.columns:
                raise ValueError(f"Coluna '{col}' não encontrada na planilha C100.")

        # Lê SEFAZ
        df_sf = pd.read_excel(arq_sefaz, dtype=str)
        df_sf.columns = [c.strip() for c in df_sf.columns]

        # Detecta coluna de chave no SEFAZ — prioriza "chave" no nome,
        # depois tenta "nf-e"/"nfe" mas apenas em colunas com conteúdo de 44 dígitos
        def _e_coluna_chave(col):
            amostra = df_sf[col].dropna().astype(str).head(5)
            return any(len("".join(filter(str.isdigit, v))) >= 40 for v in amostra)

        chave_col = next(
            (c for c in df_sf.columns if "chave" in c.lower() and _e_coluna_chave(c)),
            None
        )
        if chave_col is None:
            # fallback: primeira coluna com conteúdo de 44 dígitos
            chave_col = next(
                (c for c in df_sf.columns if _e_coluna_chave(c)),
                None
            )
        if chave_col is None:
            raise ValueError(
                f"Não foi possível identificar a coluna de chave no SEFAZ.\n"
                f"Colunas disponíveis: {list(df_sf.columns)}"
            )

        # Detecta colunas de valores no SEFAZ
        def _find_col(candidates):
            for c in df_sf.columns:
                for cand in candidates:
                    if cand.lower() in c.lower():
                        return c
            return None

        col_sf_bc   = _find_col(["total_bc_icms", "bc_icms", "vbc"])
        col_sf_icms = _find_col(["total_icms", "vicms", "vl_icms"])
        col_sf_nfe  = _find_col(["total_nf-e", "total_nfe", "vnf", "vl_nfe"])
        if not col_sf_bc:
            raise ValueError("Coluna de BC ICMS não encontrada no extrato SEFAZ.")
        if not col_sf_icms:
            raise ValueError("Coluna de ICMS não encontrada no extrato SEFAZ.")

        # Normaliza chave SEFAZ: mantém apenas dígitos
        def _norm_chave(v):
            return "".join(filter(str.isdigit, str(v))) if pd.notna(v) else ""

        df_sf["_chave_norm"] = df_sf[chave_col].apply(_norm_chave)
        mapa_sefaz = {
            row["_chave_norm"]: {
                "bc":   parse_valor_br(row.get(col_sf_bc)),
                "icms": parse_valor_br(row.get(col_sf_icms)),
                "nfe":  parse_valor_br(row.get(col_sf_nfe)) if col_sf_nfe else None,
                "num":  str(row.get("Numero", row.get("numero", ""))).strip(),
            }
            for _, row in df_sf.iterrows()
        }

        linhas = []
        for _, row in df_c100.iterrows():
            chave_raw = str(row.get("chv_nfe", "")).strip()
            chave_norm = _norm_chave(chave_raw)
            if not chave_norm:
                continue

            bc_c100   = parse_valor_br(row.get("vl_bc_icms"))
            icms_c100 = parse_valor_br(row.get("vl_icms"))
            vldoc_c100= parse_valor_br(row.get("vl_doc"))
            num_doc   = str(row.get("num_doc", "")).strip()

            sf = mapa_sefaz.get(chave_norm)
            bc_sf   = sf["bc"]   if sf else None
            icms_sf = sf["icms"] if sf else None
            nfe_sf  = sf["nfe"]  if sf else None

            # Diferenças
            dif_bc   = None if (bc_c100    is None or bc_sf   is None) else round(bc_c100    - bc_sf,   2)
            dif_icms = None if (icms_c100  is None or icms_sf is None) else round(icms_c100  - icms_sf, 2)
            dif_vldoc= None if (vldoc_c100 is None or nfe_sf  is None) else round(vldoc_c100 - nfe_sf,  2)

            diverge = (
                (dif_bc    is not None and abs(dif_bc)    > TOLERANCIA) or
                (dif_icms  is not None and abs(dif_icms)  > TOLERANCIA) or
                (dif_vldoc is not None and abs(dif_vldoc) > TOLERANCIA)
            )
            ausente = sf is None

            linhas.append({
                "chave":      chave_norm,
                "num_doc":    num_doc,
                "bc_c100":    bc_c100,
                "bc_sf":      bc_sf,
                "icms_c100":  icms_c100,
                "icms_sf":    icms_sf,
                "dif_bc":     dif_bc,
                "dif_icms":   dif_icms,
                "vldoc_c100": vldoc_c100,
                "nfe_sf":     nfe_sf,
                "dif_vldoc":  dif_vldoc,
                "diverge":    diverge,
                "ausente":    ausente,
            })

        # ── Notas no SEFAZ que não estão no C100 ────────────────────────────
        chaves_c100 = set()
        for _, row in df_c100.iterrows():
            ch = _norm_chave(str(row.get("chv_nfe", "")).strip())
            if ch:
                chaves_c100.add(ch)

        col_num  = next((c for c in df_sf.columns if c.lower() == "numero"), None)
        col_dt   = next((c for c in df_sf.columns if "emit" in c.lower() and "dt" in c.lower()), None)
        col_emit = next((c for c in df_sf.columns if "razao" in c.lower() and "emit" in c.lower()), None)

        linhas2 = []
        for _, row in df_sf.iterrows():
            chave = row.get("_chave_norm", "")
            if not chave or chave in chaves_c100:
                continue
            linhas2.append({
                "chave":    chave,
                "num_doc":  str(row.get(col_num, "")).strip()  if col_num  else "",
                "total_nfe":parse_valor_br(row.get(col_sf_nfe)) if col_sf_nfe else None,
                "bc_icms":  parse_valor_br(row.get(col_sf_bc)),
                "icms":     parse_valor_br(row.get(col_sf_icms)),
                "dt_emit":  str(row.get(col_dt,   "")).strip() if col_dt   else "",
                "emitente": str(row.get(col_emit, "")).strip() if col_emit else "",
            })

        self.linhas_sefaz  = linhas
        self.linhas_sefaz2 = linhas2
        self._exibir_resultado_sefaz(linhas)
        self._exibir_resultado_sefaz2(linhas2)

    def _exibir_resultado_sefaz(self, linhas):
        total      = len(linhas)
        ausentes   = sum(1 for l in linhas if l["ausente"])
        divergentes= sum(1 for l in linhas if l["diverge"])
        ok         = total - ausentes - divergentes

        self.lbl_resumo_sefaz.setText(
            f"Total C100: <b>{total}</b>  |  "
            f"<span style='color:{COR_OK}'>OK: {ok}</span>  |  "
            f"<span style='color:{COR_DIVERG}'>Divergentes: {divergentes}</span>  |  "
            f"<span style='color:{COR_AUSENTE}'>Ausente no SEFAZ: {ausentes}</span>"
        )

        self.tabela_sefaz.setRowCount(total)
        for i, item in enumerate(linhas):
            if item["ausente"]:
                status_txt, cor = "⚠ Ausente SEFAZ", COR_AUSENTE
            elif item["diverge"]:
                status_txt, cor = "✘ Divergente", COR_DIVERG
            else:
                status_txt, cor = "✔ OK", COR_OK

            cel_st = QTableWidgetItem(status_txt)
            cel_st.setForeground(QColor(cor))
            cel_st.setFont(QFont("Segoe UI", 9, QFont.Bold))
            self.tabela_sefaz.setItem(i, 0, cel_st)
            self.tabela_sefaz.setItem(i, 1, QTableWidgetItem(item["num_doc"]))
            self.tabela_sefaz.setItem(i, 2, QTableWidgetItem(item["chave"]))

            # BC ICMS
            cel_bc_c  = QTableWidgetItem(self._fmt(item["bc_c100"]))
            cel_bc_sf = QTableWidgetItem(self._fmt(item["bc_sf"]))
            cel_dbc   = QTableWidgetItem(self._fmt(item["dif_bc"]))
            # ICMS
            cel_ic_c  = QTableWidgetItem(self._fmt(item["icms_c100"]))
            cel_ic_sf = QTableWidgetItem(self._fmt(item["icms_sf"]))
            cel_dic   = QTableWidgetItem(self._fmt(item["dif_icms"]))
            # Vl. Doc
            cel_vd_c  = QTableWidgetItem(self._fmt(item["vldoc_c100"]))
            cel_vd_sf = QTableWidgetItem(self._fmt(item["nfe_sf"]))
            cel_dvd   = QTableWidgetItem(self._fmt(item["dif_vldoc"]))

            if item["diverge"]:
                for cel, dif in [
                    (cel_bc_c,  item["dif_bc"]),
                    (cel_bc_sf, item["dif_bc"]),
                    (cel_dbc,   item["dif_bc"]),
                    (cel_ic_c,  item["dif_icms"]),
                    (cel_ic_sf, item["dif_icms"]),
                    (cel_dic,   item["dif_icms"]),
                    (cel_vd_c,  item["dif_vldoc"]),
                    (cel_vd_sf, item["dif_vldoc"]),
                    (cel_dvd,   item["dif_vldoc"]),
                ]:
                    if dif is not None and abs(dif) > TOLERANCIA:
                        cel.setForeground(QColor(COR_DIVERG))

            self.tabela_sefaz.setItem(i, 3, cel_bc_c)
            self.tabela_sefaz.setItem(i, 4, cel_bc_sf)
            self.tabela_sefaz.setItem(i, 5, cel_dbc)
            self.tabela_sefaz.setItem(i, 6, cel_ic_c)
            self.tabela_sefaz.setItem(i, 7, cel_ic_sf)
            self.tabela_sefaz.setItem(i, 8, cel_dic)
            self.tabela_sefaz.setItem(i, 9,  cel_vd_c)
            self.tabela_sefaz.setItem(i, 10, cel_vd_sf)
            self.tabela_sefaz.setItem(i, 11, cel_dvd)

        for c in range(12):
            if c != 2:
                self.tabela_sefaz.resizeColumnToContents(c)

        self._atualizar_rodape_sefaz(linhas)
        self.status.showMessage(
            f"SEFAZ — {divergentes} divergência(s), {ausentes} ausente(s) no SEFAZ."
        )

    def _limpar_rodape_sefaz(self):
        for col in range(self.tabela_rodape_sefaz.columnCount()):
            txt = "TOTAL" if col == 0 else ""
            cel = QTableWidgetItem(txt)
            cel.setForeground(QColor(COR_SUBTEXTO))
            cel.setFont(QFont("Segoe UI", 9, QFont.Bold))
            self.tabela_rodape_sefaz.setItem(0, col, cel)

    def _atualizar_rodape_sefaz(self, linhas):
        tots = {"bc_c100": 0.0, "bc_sf": 0.0, "icms_c100": 0.0, "icms_sf": 0.0,
                "vldoc_c100": 0.0, "nfe_sf": 0.0}
        for item in linhas:
            for k in tots:
                v = item.get(k)
                if v is not None:
                    tots[k] += v

        dif_bc_tot    = round(tots["bc_c100"]    - tots["bc_sf"],   2)
        dif_icms_tot  = round(tots["icms_c100"]  - tots["icms_sf"], 2)
        dif_vldoc_tot = round(tots["vldoc_c100"] - tots["nfe_sf"],  2)

        valores = [
            (0,  "TOTAL"),
            (1,  ""), (2, ""),
            (3,  self._fmt(tots["bc_c100"])),
            (4,  self._fmt(tots["bc_sf"])),
            (5,  self._fmt(dif_bc_tot)),
            (6,  self._fmt(tots["icms_c100"])),
            (7,  self._fmt(tots["icms_sf"])),
            (8,  self._fmt(dif_icms_tot)),
            (9,  self._fmt(tots["vldoc_c100"])),
            (10, self._fmt(tots["nfe_sf"])),
            (11, self._fmt(dif_vldoc_tot)),
        ]
        for col, txt in valores:
            cel = QTableWidgetItem(txt)
            cel.setFont(QFont("Segoe UI", 9, QFont.Bold))
            if col == 0:
                cel.setForeground(QColor(COR_PRIMARIA))
            elif col in (5, 8, 11):
                cor = COR_DIVERG if (
                    (col == 5  and abs(dif_bc_tot)    > TOLERANCIA) or
                    (col == 8  and abs(dif_icms_tot)  > TOLERANCIA) or
                    (col == 11 and abs(dif_vldoc_tot) > TOLERANCIA)
                ) else COR_OK
                cel.setForeground(QColor(cor))
            else:
                cel.setForeground(QColor(COR_TEXTO))
            self.tabela_rodape_sefaz.setItem(0, col, cel)

        for c in range(self.tabela_sefaz.columnCount()):
            self.tabela_rodape_sefaz.setColumnWidth(c, self.tabela_sefaz.columnWidth(c))

    def _exibir_resultado_sefaz2(self, linhas2):
        total = len(linhas2)
        cor = COR_AUSENTE if total > 0 else COR_OK
        self.lbl_resumo_sefaz2.setText(
            f"<span style='color:{cor}'><b>{total}</b> nota(s) encontrada(s) no SEFAZ "
            f"que não constam no SPED (C100)</span>"
        )
        self.tabela_sefaz2.setRowCount(total)
        for i, item in enumerate(linhas2):
            def cel(txt, vermelho=False):
                c = QTableWidgetItem(str(txt) if txt is not None else "")
                if vermelho:
                    c.setForeground(QColor(COR_AUSENTE))
                    c.setFont(QFont("Segoe UI", 9, QFont.Bold))
                return c
            self.tabela_sefaz2.setItem(i, 0, cel(item["num_doc"],  vermelho=True))
            self.tabela_sefaz2.setItem(i, 1, cel(item["chave"]))
            self.tabela_sefaz2.setItem(i, 2, cel(self._fmt(item["total_nfe"])))
            self.tabela_sefaz2.setItem(i, 3, cel(self._fmt(item["bc_icms"])))
            self.tabela_sefaz2.setItem(i, 4, cel(self._fmt(item["icms"])))
            self.tabela_sefaz2.setItem(i, 5, cel(item["dt_emit"]))
            self.tabela_sefaz2.setItem(i, 6, cel(item["emitente"]))
        for c in range(self.tabela_sefaz2.columnCount()):
            if c != 1:
                self.tabela_sefaz2.resizeColumnToContents(c)

    def _exportar_sefaz2(self):
        if not self.linhas_sefaz2:
            QMessageBox.information(self, "Exportar", "Nenhuma nota ausente no SPED encontrada.")
            return
        caminho, _ = QFileDialog.getSaveFileName(
            self, "Salvar notas ausentes no SPED", "sefaz_ausentes_no_sped.xlsx", "Excel (*.xlsx)"
        )
        if not caminho:
            return
        registros = [
            {
                "num_doc":    l["num_doc"],
                "chave":      l["chave"],
                "total_nfe":  l["total_nfe"],
                "bc_icms":    l["bc_icms"],
                "icms":       l["icms"],
                "dt_emit":    l["dt_emit"],
                "emitente":   l["emitente"],
            }
            for l in self.linhas_sefaz2
        ]
        pd.DataFrame(registros).to_excel(caminho, index=False)
        QMessageBox.information(
            self, "Exportar",
            f"{len(registros)} nota(s) exportada(s) para:\n{caminho}"
        )

    def _limpar_sefaz(self):
        self.campo_sefaz_excel.clear()
        self.campo_sefaz_arq.clear()
        self.tabela_sefaz.setRowCount(0)
        self.tabela_sefaz2.setRowCount(0)
        self.lbl_resumo_sefaz.setText("")
        self.lbl_resumo_sefaz2.setText("")
        self.linhas_sefaz  = []
        self.linhas_sefaz2 = []
        self._limpar_rodape_sefaz()
        self.status.showMessage("Pronto.")

    def _exportar_sefaz(self):
        if not self.linhas_sefaz:
            QMessageBox.information(self, "Exportar", "Nenhum resultado para exportar.")
            return
        diverg = [l for l in self.linhas_sefaz if l["diverge"] or l["ausente"]]
        if not diverg:
            QMessageBox.information(self, "Exportar", "Nenhuma divergência encontrada!")
            return
        caminho, _ = QFileDialog.getSaveFileName(
            self, "Salvar divergências SEFAZ", "sefaz_divergencias.xlsx", "Excel (*.xlsx)"
        )
        if not caminho:
            return
        registros = []
        for l in diverg:
            registros.append({
                "status":           "Ausente SEFAZ" if l["ausente"] else "Divergente",
                "num_doc":          l["num_doc"],
                "chave":            l["chave"],
                "bc_icms_excel":    l["bc_c100"],
                "bc_icms_sefaz":    l["bc_sf"],
                "dif_bc_icms":      l["dif_bc"],
                "icms_excel":       l["icms_c100"],
                "icms_sefaz":       l["icms_sf"],
                "dif_icms":         l["dif_icms"],
                "vl_doc_excel":     l["vldoc_c100"],
                "total_nfe_sefaz":  l["nfe_sf"],
                "dif_vl_doc":       l["dif_vldoc"],
            })
        pd.DataFrame(registros).to_excel(caminho, index=False)
        QMessageBox.information(
            self, "Exportar",
            f"{len(diverg)} divergência(s) exportada(s) para:\n{caminho}"
        )

    # ── Métodos aba XML ───────────────────────────────────────────────────────

    def _selecionar_excel(self):
        caminho, _ = QFileDialog.getOpenFileName(
            self, "Selecionar planilha Excel", "",
            "Excel (*.xlsx *.xls);;Todos os arquivos (*)"
        )
        if caminho:
            self.campo_excel.setText(caminho)

    def _selecionar_pasta(self):
        pasta = QFileDialog.getExistingDirectory(self, "Selecionar pasta com XMLs")
        if pasta:
            self.campo_pasta.setText(pasta)

    def _comparar(self):
        excel = self.campo_excel.text().strip()
        pasta = self.campo_pasta.text().strip()

        if not excel or not os.path.isfile(excel):
            QMessageBox.warning(self, "Atenção", "Selecione um arquivo Excel válido.")
            return
        if not pasta or not os.path.isdir(pasta):
            QMessageBox.warning(self, "Atenção", "Selecione uma pasta de XMLs válida.")
            return

        self.btn_comparar.setEnabled(False)
        self.status.showMessage("Processando…")
        self.tabela.setRowCount(0)
        self.lbl_resumo.setText("")

        self.worker = ComparadorWorker(excel, pasta)
        self.worker.resultado.connect(self._exibir_resultado)
        self.worker.erro.connect(self._exibir_erro)
        self.worker.finished.connect(lambda: self.btn_comparar.setEnabled(True))
        self.worker.start()

    @staticmethod
    def _fmt(valor):
        return "" if valor is None else f"{valor:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")

    def _exibir_resultado(self, linhas):
        self.linhas = list(linhas)
        total = len(linhas)
        ausentes = sum(1 for l in linhas if not l["xml_encontrado"])
        divergentes = sum(1 for l in linhas if l["diverge"] is True)
        ok = total - ausentes - divergentes

        self.lbl_resumo.setText(
            f"Total no Excel: <b>{total}</b>  |  "
            f"<span style='color:{COR_OK}'>OK: {ok}</span>  |  "
            f"<span style='color:{COR_DIVERG}'>Divergentes: {divergentes}</span>  |  "
            f"<span style='color:{COR_AUSENTE}'>XML ausente: {ausentes}</span>"
        )

        self.tabela.setRowCount(total)
        for i, item in enumerate(linhas):
            if not item["xml_encontrado"]:
                status_txt, cor = "⚠ XML ausente", COR_AUSENTE
            elif item["erro_xml"]:
                status_txt, cor = "⚠ Erro leitura", COR_AUSENTE
            elif item["diverge"]:
                status_txt, cor = "✘ Divergente", COR_DIVERG
            else:
                status_txt, cor = "✔ OK", COR_OK

            item_status = QTableWidgetItem(status_txt)
            item_status.setForeground(QColor(cor))
            item_status.setFont(QFont("Segoe UI", 9, QFont.Bold))
            self.tabela.setItem(i, 0, item_status)

            self.tabela.setItem(i, 1, QTableWidgetItem(item["num_doc"]))

            valores = [
                (2, item["vl_bc_icms_xls"], item["vl_bc_icms_xml"]),
                (4, item["vl_icms_xls"], item["vl_icms_xml"]),
                (6, item["vl_doc_xls"], item["vl_doc_xml"]),
            ]
            for col_xls, v_xls, v_xml in valores:
                cel_xls = QTableWidgetItem(self._fmt(v_xls))
                cel_xml = QTableWidgetItem(self._fmt(v_xml))
                if item["xml_encontrado"] and not item["erro_xml"] and v_xls is not None and v_xml is not None:
                    if abs(v_xls - v_xml) > TOLERANCIA:
                        cel_xls.setForeground(QColor(COR_DIVERG))
                        cel_xml.setForeground(QColor(COR_DIVERG))
                self.tabela.setItem(i, col_xls, cel_xls)
                self.tabela.setItem(i, col_xls + 1, cel_xml)

            self.tabela.setItem(i, 8, QTableWidgetItem(item["chave"]))

        for c in range(8):
            self.tabela.resizeColumnToContents(c)

        self.status.showMessage(
            f"Concluído — {divergentes} divergência(s), {ausentes} XML(s) ausente(s)."
        )

    def _exibir_erro(self, mensagem):
        QMessageBox.critical(self, "Erro", f"Falha ao processar:\n\n{mensagem}")
        self.status.showMessage("Erro no processamento.")

    def _limpar(self):
        self.campo_excel.clear()
        self.campo_pasta.clear()
        self.tabela.setRowCount(0)
        self.lbl_resumo.setText("")
        self.status.showMessage("Pronto.")
        self.linhas = []

    def _exportar(self):
        if not self.linhas:
            QMessageBox.information(self, "Exportar", "Nenhum resultado para exportar.")
            return

        problemas = [l for l in self.linhas if not l["xml_encontrado"] or l["erro_xml"] or l["diverge"]]
        if not problemas:
            QMessageBox.information(self, "Exportar", "Nenhuma divergência! Tudo certo.")
            return

        caminho, _ = QFileDialog.getSaveFileName(
            self, "Salvar divergências", "nfe_divergencias.xlsx", "Excel (*.xlsx)"
        )
        if not caminho:
            return

        registros = []
        for item in problemas:
            if not item["xml_encontrado"]:
                status = "XML ausente"
            elif item["erro_xml"]:
                status = f"Erro leitura: {item['erro_xml']}"
            else:
                status = "Divergente"
            registros.append({
                "status": status,
                "num_doc": item["num_doc"],
                "chave": item["chave"],
                "vl_bc_icms_excel": item["vl_bc_icms_xls"],
                "vl_bc_icms_xml": item["vl_bc_icms_xml"],
                "vl_icms_excel": item["vl_icms_xls"],
                "vl_icms_xml": item["vl_icms_xml"],
                "vl_doc_excel": item["vl_doc_xls"],
                "vl_nf_xml": item["vl_doc_xml"],
            })

        df_out = pd.DataFrame(registros)
        df_out.to_excel(caminho, index=False)

        QMessageBox.information(
            self, "Exportar",
            f"{len(problemas)} divergência(s) exportada(s) para:\n{caminho}"
        )


    def _gerar_sql(self):
        if not self.linhas:
            QMessageBox.information(self, "Gerar SQL", "Nenhum resultado para gerar SQL.")
            return

        divergentes = [l for l in self.linhas if l["diverge"] is True and l["xml_encontrado"] and not l["erro_xml"]]
        if not divergentes:
            QMessageBox.information(self, "Gerar SQL", "Nenhuma divergência com XML disponível para gerar SQL.")
            return

        caminho, _ = QFileDialog.getSaveFileName(
            self, "Salvar script SQL", "correcao_nfe.sql", "SQL (*.sql);;Todos (*)"
        )
        if not caminho:
            return

        linhas_sql = []
        linhas_sql.append("-- Script gerado automaticamente pelo Comparador NF-e")
        linhas_sql.append("-- Corrige cstcsosn e vbc em vendaproduto a partir dos XMLs")
        linhas_sql.append("-- REVISE antes de executar!\n")
        linhas_sql.append("START TRANSACTION;\n")

        total_updates = 0
        for nota in divergentes:
            vendaunicoid = nota["vendaunicoid"]
            chave = nota["chave"]
            linhas_sql.append(f"-- NF-e: {chave}  |  vendaUnicoId: {vendaunicoid}")

            for item in nota["itens_xml"]:
                cprod   = item["cprod"]
                cst_sql = f"'{item['cst']}'" if item["cst"] is not None else "NULL"
                vbc     = item["vbc"]   or "0"
                picms   = item["picms"] or "0"
                vicms   = item["vicms"] or "0"

                sql = (
                    f"UPDATE vendaproduto "
                    f"SET cstcsosn = {cst_sql}, vbc = {vbc}, picms = {picms}, vicmsop = {vicms} "
                    f"WHERE vendaunicoid = '{vendaunicoid}' "
                    f"AND produtoid = '{cprod}';"
                )
                linhas_sql.append(sql)
                total_updates += 1

            linhas_sql.append("")  # linha em branco entre notas

        linhas_sql.append("COMMIT;")

        with open(caminho, "w", encoding="utf-8") as f:
            f.write("\n".join(linhas_sql))

        QMessageBox.information(
            self, "Gerar SQL",
            f"{total_updates} UPDATE(s) gerados em {len(divergentes)} nota(s) divergente(s).\n\nArquivo: {caminho}"
        )


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    janela = JanelaPrincipal()
    janela.show()
    sys.exit(app.exec_())