import sqlite3
import datetime
import os
import sys
import json
import io
import unicodedata
import requests
import tempfile
import shutil
from urllib.parse import urlsplit
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad

if getattr(sys, 'frozen', False):
    # Executando como .exe
    application_path = sys._MEIPASS
else:
    # Executando como script
    application_path = os.path.dirname(os.path.abspath(__file__))

# Exemplo de uso:
# caminho_do_asset = os.path.join(application_path, "customtkinter", "assets", "themes", "blue", "checkbox_checked.png")

import plotly.graph_objects as go
import plotly.express as px

from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QLabel,
    QLineEdit,
    QTextEdit,
    QPushButton,
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox,
    QFileDialog,
    QCheckBox,
    QInputDialog,
    QHeaderView,
    QToolButton,
    QCalendarWidget,
    QTabWidget,
    QButtonGroup,
    QFrame,
    QScrollBar,
    QSizePolicy,
)
from PySide6.QtCore import QTimer, Qt, QEvent, QLocale, QDate, QPoint
from PySide6.QtGui import QFont, QPixmap, QColor, QIcon, QPalette, QTextCharFormat, QPainter, QPolygon, QFontMetrics, QTextOption

import PySide6.QtSvg

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus import Table, TableStyle, Image, Paragraph
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_LEFT
except Exception:
    print("Biblioteca reportlab não encontrada. Instale com: pip install -r requirements.txt")
    sys.exit(1)

def get_app_data_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(__file__)


def get_resource_dir():
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS
    return os.path.dirname(__file__)


APP_DATA_DIR = get_app_data_dir()
RESOURCE_DIR = get_resource_dir()
SOURCE_PROJECT_ROOT_DIR = os.path.abspath(os.path.dirname(__file__))
JSON_DIR = os.path.join(APP_DATA_DIR, "json")
DATA_JSON_PATH = os.path.join(JSON_DIR, "data.json")
DEFAULT_UPLOAD_URL = "http://localhost:3000/upload-data"

DB_PATH = os.path.join(APP_DATA_DIR, "documents.db")
ICON_PATH = os.path.join(RESOURCE_DIR, "Icone.png")
ESCUDO_PATH = os.path.join(RESOURCE_DIR, "ESCUDO PMMG.png")
CHECK_ICON_PATH = os.path.join(RESOURCE_DIR, "check_blue.svg")
CHECK_ICON_URL = CHECK_ICON_PATH.replace("\\", "/")
CONFIG_JSON_PATH = os.path.join(APP_DATA_DIR, "config.json")

DEFAULT_INSTITUTIONAL_TEXT = "Polícia Militar de Minas Gerais"

# ========================================

def _normalize_config_text(value):
    if value is None:
        return ""
    return str(value).strip()


def load_runtime_config(config_path=CONFIG_JSON_PATH):
    """Carrega config.json com fallback seguro para texto institucional padrão."""
    config_data = {}

    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                loaded_data = json.load(f)
            if isinstance(loaded_data, dict):
                config_data = loaded_data
        except Exception:
            config_data = {}

    subtitle = _normalize_config_text(config_data.get("Sub-titulo"))
    if not subtitle:
        subtitle = DEFAULT_INSTITUTIONAL_TEXT

    header_lines = []
    raw_header = config_data.get("Cabecalho_PDF")
    if isinstance(raw_header, list):
        for line in raw_header:
            parsed_line = _normalize_config_text(line)
            if parsed_line:
                header_lines.append(parsed_line)
    elif isinstance(raw_header, str):
        parsed_line = _normalize_config_text(raw_header)
        if parsed_line:
            header_lines.append(parsed_line)

    if not header_lines:
        header_lines = [DEFAULT_INSTITUTIONAL_TEXT]

    rodape = _normalize_config_text(config_data.get("rodapé"))

    return {
        "Sub-titulo": subtitle,
        "Cabecalho_PDF": header_lines,
        "rodapé": rodape,
    }


RUNTIME_CONFIG = load_runtime_config()
APP_SUBTITLE = RUNTIME_CONFIG["Sub-titulo"]
PDF_HEADER_LINES = RUNTIME_CONFIG["Cabecalho_PDF"]
APP_FOOTER = RUNTIME_CONFIG["rodapé"]

def get_runtime_env_path():
    """Retorna o caminho do .env conforme contexto de execução atual."""
    if getattr(sys, "frozen", False):
        return os.path.join(os.path.dirname(sys.executable), ".env")
    return os.path.join(SOURCE_PROJECT_ROOT_DIR, ".env")


def get_env_value(key, env_path=None):
    """Lê um valor simples de arquivo .env no formato CHAVE=valor."""
    if env_path is None:
        env_path = get_runtime_env_path()

    if not os.path.exists(env_path):
        return None

    try:
        with open(env_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                if k.strip() == key:
                    return v.strip().strip('"').strip("'")
    except Exception:
        return None

    return None


def get_data_upload_url():
    """Retorna a URL de upload configurada no .env (ou padrão local)."""
    configured_url = get_env_value("DATA_UPLOAD_URL")
    if not configured_url:
        configured_url = DEFAULT_UPLOAD_URL

    configured_url = configured_url.strip()
    if not configured_url:
        return DEFAULT_UPLOAD_URL

    if configured_url.startswith("http://") or configured_url.startswith("https://"):
        normalized_url = configured_url.rstrip("/")
        parsed = urlsplit(normalized_url)

        # Se o usuário informou uma rota explícita no .env, respeita exatamente.
        if parsed.path and parsed.path != "/":
            return normalized_url

        return normalized_url + "/upload-data"

    return DEFAULT_UPLOAD_URL


def build_upload_url_candidates():
    """Monta lista de URLs candidatas de upload para lidar com variações de rota."""
    primary_url = get_data_upload_url().rstrip("/")
    candidates = [primary_url]

    try:
        parsed = urlsplit(primary_url)
        if parsed.scheme and parsed.netloc:
            origin = f"{parsed.scheme}://{parsed.netloc}"
            path = (parsed.path or "").rstrip("/")

            if path == "/upload-data":
                candidates.append(origin + "/api/upload-data")
            elif path == "/api/upload-data":
                candidates.append(origin + "/upload-data")
            elif path in ("", "/"):
                candidates.append(origin + "/upload-data")
                candidates.append(origin + "/api/upload-data")
    except Exception:
        pass

    unique_candidates = []
    for url in candidates:
        if url and url not in unique_candidates:
            unique_candidates.append(url)

    return unique_candidates

def normalize_text(text):
    """Remove acentos e cedilhas do texto para normalização"""
    if not text:
        return text
    # Normaliza para NFD (Canonical Decomposition)
    nfd = unicodedata.normalize('NFD', text)
    # Remove acentos (combining characters)
    return ''.join(char for char in nfd if unicodedata.category(char) != 'Mn')


def show_date_picker_dialog(parent, initial_date=None, show_month_year_label=False, dialog_size=None):
    """Exibe um calendário padronizado e retorna a data selecionada (datetime.date) ou None."""
    dlg = QDialog(parent)
    dlg.setWindowTitle("Selecionar Data")
    dlg.setModal(True)
    if dialog_size:
        dlg.resize(*dialog_size)

    layout = QVBoxLayout(dlg)

    month_year_label = None
    if show_month_year_label:
        month_year_label = QLabel()
        month_year_label.setAlignment(Qt.AlignCenter)
        month_year_label.setFont(QFont("Arial", 12, QFont.Bold))
        month_year_label.setStyleSheet("""
            QLabel {
                color: #003366;
                background-color: #e8f4f8;
                padding: 10px;
                border-radius: 5px;
                margin: 5px;
            }
        """)
        layout.addWidget(month_year_label)

    calendar = QCalendarWidget(dlg)
    calendar.setGridVisible(True)
    # Esconde a barra de navegação padrão apenas quando há label customizado
    calendar.setNavigationBarVisible(not show_month_year_label)
    calendar.setHorizontalHeaderFormat(QCalendarWidget.SingleLetterDayNames)
    calendar.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
    calendar.setLocale(QLocale(QLocale.Portuguese, QLocale.Brazil))
    # Configurar cores da paleta para o calendário (ativo e inativo)
    palette = calendar.palette()
    selection_color = QColor("#0078D4")
    palette.setColor(QPalette.Highlight, selection_color)
    palette.setColor(QPalette.HighlightedText, QColor("white"))
    palette.setColor(QPalette.Active, QPalette.Highlight, selection_color)
    palette.setColor(QPalette.Active, QPalette.HighlightedText, QColor("white"))
    palette.setColor(QPalette.Inactive, QPalette.Highlight, selection_color)
    palette.setColor(QPalette.Inactive, QPalette.HighlightedText, QColor("white"))
    calendar.setPalette(palette)
    
    calendar.setStyleSheet("""
        QCalendarWidget QWidget#qt_calendar_navigationbar {
            background-color: #003366;
            min-height: 30px;
        }
        QCalendarWidget QToolButton {
            color: white;
            font-size: 12px;
            font-weight: bold;
            background-color: transparent;
            border: none;
            margin: 2px;
        }
        QCalendarWidget QToolButton:hover {
            background-color: #0056b3;
            border-radius: 4px;
        }
        QCalendarWidget QToolButton#qt_calendar_prevmonth,
        QCalendarWidget QToolButton#qt_calendar_nextmonth {
            min-width: 74px;
        }
        QCalendarWidget QToolButton#qt_calendar_monthbutton {
            padding-right: 8px;
        }
        QCalendarWidget QToolButton#qt_calendar_monthbutton::menu-indicator {
            image: none;
            width: 0px;
        }
        QCalendarWidget QMenu {
            background-color: white;
            color: #003366;
        }
        QCalendarWidget QSpinBox {
            color: white;
            font-size: 12px;
            font-weight: bold;
            background-color: transparent;
            border: none;
        }
        QCalendarWidget QAbstractItemView {
            color: #222222;
            selection-background-color: #0078D4;
            selection-color: white;
        }
        QCalendarWidget QTableView {
            selection-background-color: #0078D4;
        }
    """)

    # Criar botões Anterior/Próximo quando há label customizado
    if show_month_year_label:
        nav_layout = QHBoxLayout()
        nav_layout.addStretch()
        
        btn_prev = QPushButton("< Anterior")
        btn_prev.setMaximumWidth(100)
        
        def go_prev_month():
            y, m = calendar.yearShown(), calendar.monthShown()
            if m == 1:
                calendar.setCurrentPage(y - 1, 12)
            else:
                calendar.setCurrentPage(y, m - 1)
        
        btn_prev.clicked.connect(go_prev_month)
        nav_layout.addWidget(btn_prev)
        
        btn_next = QPushButton("Próximo >")
        btn_next.setMaximumWidth(100)
        
        def go_next_month():
            y, m = calendar.yearShown(), calendar.monthShown()
            if m == 12:
                calendar.setCurrentPage(y + 1, 1)
            else:
                calendar.setCurrentPage(y, m + 1)
        
        btn_next.clicked.connect(go_next_month)
        nav_layout.addWidget(btn_next)
        
        nav_layout.addStretch()
        layout.addLayout(nav_layout)
    else:
        prev_btn = calendar.findChild(QToolButton, "qt_calendar_prevmonth")
        if prev_btn:
            prev_btn.setText("Anterior")
            prev_btn.setToolTip("Mês anterior")

        next_btn = calendar.findChild(QToolButton, "qt_calendar_nextmonth")
        if next_btn:
            next_btn.setText("Próximo")
            next_btn.setToolTip("Próximo mês")

    if isinstance(initial_date, datetime.date):
        calendar.setSelectedDate(QDate(initial_date.year, initial_date.month, initial_date.day))

    # Força a célula da data selecionada com azul, independentemente do estilo do sistema.
    selected_format = QTextCharFormat()
    selected_format.setBackground(selection_color)
    selected_format.setForeground(QColor("white"))

    selected_state = {"last": None}

    def apply_selected_date_format():
        last_date = selected_state["last"]
        if last_date is not None:
            calendar.setDateTextFormat(last_date, QTextCharFormat())
        current_date = calendar.selectedDate()
        calendar.setDateTextFormat(current_date, selected_format)
        selected_state["last"] = QDate(current_date)

    apply_selected_date_format()
    calendar.selectionChanged.connect(apply_selected_date_format)

    if month_year_label is not None:
        def update_month_year_label():
            if month_year_label is None: return
            months_pt = [
                "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
                "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
            ]
            month_name = months_pt[calendar.monthShown() - 1]
            month_year_label.setText(f"{month_name} de {calendar.yearShown()}")

        update_month_year_label()
        calendar.currentPageChanged.connect(update_month_year_label)

    layout.addWidget(calendar)

    btn_row = QHBoxLayout()
    btn_row.addStretch()

    btn_cancel = QPushButton("Cancelar")
    btn_cancel.setCursor(Qt.PointingHandCursor)
    btn_cancel.setStyleSheet("""
        QPushButton {
            background-color: #6c757d;
            color: white;
            border: none;
            border-radius: 6px;
            padding: 8px 16px;
            font-weight: bold;
            min-width: 100px;
        }
        QPushButton:hover {
            background-color: #5a6268;
        }
        QPushButton:pressed {
            background-color: #545b62;
        }
    """)
    btn_cancel.clicked.connect(dlg.reject)
    btn_row.addWidget(btn_cancel)

    btn_ok = QPushButton("Selecionar")
    btn_ok.setCursor(Qt.PointingHandCursor)
    btn_ok.setStyleSheet("""
        QPushButton {
            background-color: #007BFF;
            color: white;
            border: none;
            border-radius: 6px;
            padding: 8px 16px;
            font-weight: bold;
            min-width: 110px;
        }
        QPushButton:hover {
            background-color: #0056b3;
        }
        QPushButton:pressed {
            background-color: #004085;
        }
    """)
    btn_ok.clicked.connect(dlg.accept)
    btn_row.addWidget(btn_ok)

    layout.addLayout(btn_row)

    if dlg.exec() == QDialog.Accepted:
        return calendar.selectedDate().toPython()
    return None

# ========================================

COLUMN_WIDTHS_FILE = os.path.join(APP_DATA_DIR, "table_column_widths.json")
CHECKBOX_CONFIG_FILE = os.path.join(APP_DATA_DIR, "checkbox_config.json")
ALERTS_CONFIG_FILE = os.path.join(APP_DATA_DIR, "alerts_config.json")
RESULT_POS_UNLOCK_PASSWORD = "pmmg"


def get_default_column_widths():
    return [
        COL_INFO_WIDTH,
        COL_TIPO_WIDTH,
        COL_NUMERO_WIDTH,
        COL_NATUREZA_WIDTH,
        COL_LOCAL_WIDTH,
        COL_OBSERVACAO_WIDTH,
        COL_RECEBIMENTO_WIDTH,
        COL_VALIDADE_WIDTH,
        COL_STATUS_WIDTH,
        COL_MILITAR_WIDTH,
        COL_RESULT_POS_WIDTH,
    ]


def load_column_widths_config():
    if not os.path.exists(COLUMN_WIDTHS_FILE):
        return {}

    try:
        with open(COLUMN_WIDTHS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        pass

    return {}


def save_column_widths_config(config):
    try:
        with open(COLUMN_WIDTHS_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def load_checkbox_config():
    """Carrega o estado dos checkboxes salvos"""
    if not os.path.exists(CHECKBOX_CONFIG_FILE):
        return {"show_expired": False, "export_expired": False}

    try:
        with open(CHECKBOX_CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        pass

    return {"show_expired": False, "export_expired": False}


def save_checkbox_config(config):
    """Salva o estado dos checkboxes"""
    try:
        with open(CHECKBOX_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def load_alerts_config():
    """Carrega o histórico de alertas já mostrados"""
    if not os.path.exists(ALERTS_CONFIG_FILE):
        return {}

    try:
        with open(ALERTS_CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return data
    except Exception:
        pass

    return {}


def save_alerts_config(config):
    """Salva o histórico de alertas mostrados"""
    try:
        with open(ALERTS_CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def check_alert_shown_today(alert_type):
    """Verifica se um alerta já foi mostrado hoje"""
    today = datetime.date.today().isoformat()
    config = load_alerts_config()
    return config.get(alert_type) == today


def mark_alert_shown(alert_type):
    """Marca que um alerta foi mostrado hoje"""
    today = datetime.date.today().isoformat()
    config = load_alerts_config()
    config[alert_type] = today
    save_alerts_config(config)


def persist_table_column_widths(table_key, table):
    config = load_column_widths_config()
    config[table_key] = {
        str(col): int(table.columnWidth(col))
        for col in range(table.columnCount())
    }
    save_column_widths_config(config)


def enforce_min_column_width(table, col, width):
    min_width = 90 if col == 10 else 30
    return max(min_width, width)


def handle_table_column_resized(table_key, table, index, new_size):
    if getattr(table, "_column_resize_in_progress", False):
        return

    table._column_resize_in_progress = True
    try:
        adjusted_width = enforce_min_column_width(table, index, int(new_size))
        if adjusted_width != int(new_size):
            header = table.horizontalHeader()
            header_signals_were_blocked = header.blockSignals(True)
            try:
                table.setColumnWidth(index, adjusted_width)
            finally:
                header.blockSignals(header_signals_were_blocked)

        timer = getattr(table, "_column_width_persist_timer", None)
        if timer is None:
            timer = QTimer(table)
            timer.setSingleShot(True)
            timer.timeout.connect(lambda t=table, key=table_key: persist_table_column_widths(key, t))
            table._column_width_persist_timer = timer

        timer.start(250)
    finally:
        table._column_resize_in_progress = False


def confirm_sim_nao(parent, title, message):
    dlg = QMessageBox(parent)
    dlg.setIcon(QMessageBox.Question)
    dlg.setWindowTitle(title)
    dlg.setText(message)
    btn_sim = dlg.addButton("SIM", QMessageBox.YesRole)
    dlg.addButton("NÃO", QMessageBox.NoRole)
    dlg.exec()
    return dlg.clickedButton() == btn_sim


def apply_table_column_behavior(table, table_key, local_elastic=False):
    config = load_column_widths_config()
    saved_widths = config.get(table_key, {})
    default_widths = get_default_column_widths()

    header = table.horizontalHeader()
    header.setStretchLastSection(False)

    for col in range(table.columnCount()):
        header.setSectionResizeMode(col, QHeaderView.Interactive)

        default_width = default_widths[col] if col < len(default_widths) else 100
        width = saved_widths.get(str(col), default_width)

        try:
            width = int(width)
        except Exception:
            width = default_width

        table.setColumnWidth(col, enforce_min_column_width(table, col, width))

    if local_elastic and table.columnCount() > 4:
        header.setSectionResizeMode(4, QHeaderView.Stretch)

    header.sectionResized.connect(
        lambda _index, _old, _new, t=table, key=table_key: handle_table_column_resized(key, t, _index, _new)
    )


class BlueScrollBar(QScrollBar):
    """Scrollbar customizada para garantir handle azul em qualquer tema/estilo do sistema."""
    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        if orientation == Qt.Vertical:
            self.setFixedWidth(16)
        else:
            self.setFixedHeight(16)

    def paintEvent(self, event):
        del event

        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)

        arrow_size = 16

        if self.orientation() == Qt.Vertical:
            top_btn = self.rect().adjusted(0, 0, 0, 0)
            top_btn.setHeight(arrow_size)

            bottom_btn = self.rect().adjusted(0, 0, 0, 0)
            bottom_btn.setTop(max(0, self.height() - arrow_size))

            groove = self.rect().adjusted(1, arrow_size, -1, -arrow_size)

            # Trilha e botões
            painter.fillRect(self.rect(), QColor("#f0f0f0"))
            painter.fillRect(groove, QColor("#e6e6e6"))
            painter.fillRect(top_btn, QColor("#e6e6e6"))
            painter.fillRect(bottom_btn, QColor("#e6e6e6"))
            painter.setPen(QColor("#d0d0d0"))
            painter.drawRect(self.rect().adjusted(0, 0, -1, -1))

            # Cálculo manual do handle
            min_v = self.minimum()
            max_v = self.maximum()
            page = max(1, self.pageStep())
            track_h = max(1, groove.height())

            if max_v > min_v:
                total = (max_v - min_v) + page
                handle_h = max(30, int(track_h * (page / total)))
                handle_h = min(handle_h, track_h)
                usable = max(0, track_h - handle_h)
                ratio = (self.value() - min_v) / max(1, (max_v - min_v))
                handle_y = groove.y() + int(usable * ratio)
            else:
                handle_h = max(36, min(72, track_h // 3))
                handle_y = groove.y() + max(0, (track_h - handle_h) // 2)

            handle_x = groove.x() + 2
            handle_w = max(8, groove.width() - 4)

            painter.setPen(QColor("#002244"))
            painter.setBrush(QColor("#003366"))
            painter.drawRoundedRect(handle_x, handle_y, handle_w, handle_h, 4, 4)

            # Setas
            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor("#003366"))
            up_cx = top_btn.center().x()
            up_cy = top_btn.center().y()
            up = QPolygon([QPoint(up_cx, up_cy - 4), QPoint(up_cx - 4, up_cy + 2), QPoint(up_cx + 4, up_cy + 2)])
            painter.drawPolygon(up)

            dn_cx = bottom_btn.center().x()
            dn_cy = bottom_btn.center().y()
            down = QPolygon([QPoint(dn_cx, dn_cy + 4), QPoint(dn_cx - 4, dn_cy - 2), QPoint(dn_cx + 4, dn_cy - 2)])
            painter.drawPolygon(down)

        else:
            left_btn = self.rect().adjusted(0, 0, 0, 0)
            left_btn.setWidth(arrow_size)

            right_btn = self.rect().adjusted(0, 0, 0, 0)
            right_btn.setLeft(max(0, self.width() - arrow_size))

            groove = self.rect().adjusted(arrow_size, 1, -arrow_size, -1)

            painter.fillRect(self.rect(), QColor("#f0f0f0"))
            painter.fillRect(groove, QColor("#e6e6e6"))
            painter.fillRect(left_btn, QColor("#e6e6e6"))
            painter.fillRect(right_btn, QColor("#e6e6e6"))
            painter.setPen(QColor("#d0d0d0"))
            painter.drawRect(self.rect().adjusted(0, 0, -1, -1))

            min_v = self.minimum()
            max_v = self.maximum()
            page = max(1, self.pageStep())
            track_w = max(1, groove.width())

            if max_v > min_v:
                total = (max_v - min_v) + page
                handle_w = max(30, int(track_w * (page / total)))
                handle_w = min(handle_w, track_w)
                usable = max(0, track_w - handle_w)
                ratio = (self.value() - min_v) / max(1, (max_v - min_v))
                handle_x = groove.x() + int(usable * ratio)
            else:
                handle_w = max(36, min(72, track_w // 3))
                handle_x = groove.x() + max(0, (track_w - handle_w) // 2)

            handle_y = groove.y() + 2
            handle_h = max(8, groove.height() - 4)

            painter.setPen(QColor("#002244"))
            painter.setBrush(QColor("#003366"))
            painter.drawRoundedRect(handle_x, handle_y, handle_w, handle_h, 4, 4)

            painter.setPen(Qt.NoPen)
            painter.setBrush(QColor("#003366"))
            lx = left_btn.center().x()
            ly = left_btn.center().y()
            left = QPolygon([QPoint(lx - 4, ly), QPoint(lx + 2, ly - 4), QPoint(lx + 2, ly + 4)])
            painter.drawPolygon(left)

            rx = right_btn.center().x()
            ry = right_btn.center().y()
            right = QPolygon([QPoint(rx + 4, ry), QPoint(rx - 2, ry - 4), QPoint(rx - 2, ry + 4)])
            painter.drawPolygon(right)




# CONFIGURAÇÃO DE LARGURAS DAS COLUNAS
# Ajuste os valores abaixo para modificar a largura das colunas
# Estas configurações afetam tanto a tabela principal quanto o PDF exportado
# ========================================
# Larguras para a tabela da interface (em pixels)
COL_INFO_WIDTH = 50          # Coluna: Info (botão)
COL_TIPO_WIDTH = 80          # Coluna: Tipo
COL_NUMERO_WIDTH = 150       # Coluna: Número
COL_NATUREZA_WIDTH = 160     # Coluna: Natureza
COL_LOCAL_WIDTH = 220        # Coluna: Local
COL_OBSERVACAO_WIDTH = 270   # Coluna: Observação
COL_RECEBIMENTO_WIDTH = 110   # Coluna: Início
COL_VALIDADE_WIDTH = 95      # Coluna: Validade
COL_MILITAR_WIDTH = 110      # Coluna: Militar
COL_STATUS_WIDTH = 70        # Coluna: Status
COL_RESULT_POS_WIDTH = 90    # Coluna: Resultado Positivo

# Configuração do CheckBox Resultado Positivo
CHECKBOX_SIZE = 20            # Tamanho do checkbox em pixels (largura e altura)
CHECKBOX_MARGIN_LEFT = 0     # Margem esquerda do checkbox em pixels

# Larguras para o PDF (em centímetros) - Total disponível em paisagem: ~27cm
# Proporcionais às larguras da tabela principal
PDF_COL_TIPO_WIDTH = 1.8     # Coluna: Tipo (80px)
PDF_COL_NUMERO_WIDTH = 3.2   # Coluna: Número (150px)
PDF_COL_NATUREZA_WIDTH = 3.0 # Coluna: Natureza (140px)
PDF_COL_LOCAL_WIDTH = 5.0    # Coluna: Local (220px)
PDF_COL_OBSERVACAO_WIDTH = 5.0  # Coluna: Observação (180px)
PDF_COL_RECEBIMENTO_WIDTH = 2.0  # Coluna: Início (95px)
PDF_COL_VALIDADE_WIDTH = 2.0     # Coluna: Validade (95px)
PDF_COL_MILITAR_WIDTH = 2.4      # Coluna: Militar (110px)
PDF_COL_RESULTADO_POS_WIDTH = 2.5 # Coluna: Resultado Positivo
# ========================================


def _wrap_text_lines(text, metrics, max_width_px):
    """Quebra texto em linhas que cabem em max_width_px (por palavras)."""
    normalized = " ".join((text or "").replace("\r", " ").replace("\n", " ").split())
    if not normalized:
        return []

    words = normalized.split(" ")
    lines = []
    current = ""

    for word in words:
        candidate = word if not current else f"{current} {word}"
        if metrics.horizontalAdvance(candidate) <= max_width_px:
            current = candidate
            continue

        if current:
            lines.append(current)
            current = word
        else:
            # Palavra maior que a largura: quebra por caracteres.
            chunk = ""
            for ch in word:
                cand = chunk + ch
                if metrics.horizontalAdvance(cand) <= max_width_px:
                    chunk = cand
                else:
                    if chunk:
                        lines.append(chunk)
                    chunk = ch
            current = chunk

    if current:
        lines.append(current)

    return lines


def _clamp_observation_two_lines(text, metrics, max_width_px):
    """Retorna (texto_2_linhas, is_truncated). Segunda linha recebe '...' se truncar."""
    lines = _wrap_text_lines(text, metrics, max_width_px)
    if not lines:
        return "", False

    if len(lines) <= 2:
        return "\n".join(lines[:2]), False

    first = lines[0]
    second_raw = lines[1]
    ell = "..."
    max_second = max(0, max_width_px - metrics.horizontalAdvance(ell))

    # Garante que a 2ª linha + "..." caiba.
    trimmed = second_raw
    while trimmed and metrics.horizontalAdvance(trimmed) > max_second:
        trimmed = trimmed[:-1]
    trimmed = trimmed.rstrip()
    second = f"{trimmed}{ell}" if trimmed else ell
    return f"{first}\n{second}", True


class ReadMoreObservationDialog(QDialog):
    def __init__(self, ddu_number, natureza, observation_text, parent=None):
        super().__init__(parent)
        self.setWindowTitle("LER MAIS")
        self.setModal(True)
        self.resize(650, 520)

        self.setStyleSheet("""
            QDialog { background-color: #f5f5f5; }
            QLabel { color: #003366; }
        """)

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        # ===== Cabeçalho (fixo) =====
        header = QWidget()
        header.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #003366, stop:1 #004080);
                border-radius: 10px;
                padding: 12px;
            }
        """)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(12, 10, 12, 10)

        title_box = QVBoxLayout()
        title_box.setSpacing(2)

        title = QLabel(f"DDU Nº {ddu_number}")
        title.setStyleSheet("color: white; font-size: 16px; font-weight: bold;")

        subtitle = QLabel(str(natureza or "").strip())
        subtitle.setStyleSheet("color: #e0e0e0; font-size: 13px;")

        title_box.addWidget(title)
        title_box.addWidget(subtitle)


        header_layout.addLayout(title_box)
        header_layout.addStretch()

        root.addWidget(header)

        # ===== Corpo (com scroll somente aqui) =====
        body = QTextEdit()
        body.setReadOnly(True)
        body.setPlainText(observation_text or "")
        body.setStyleSheet("""
            QTextEdit {
                background-color: white;
                border: 1px solid #d0d0d0;
                border-radius: 10px;
                padding: 12px;
                font-size: 13px;
                color: #222222;
            }
        """)
        body.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        root.addWidget(body, 1)

        # ===== Rodapé (fixo) =====
        footer = QWidget()
        footer_layout = QHBoxLayout(footer)
        footer_layout.setContentsMargins(0, 0, 0, 0)

        btn_close = QPushButton("Fechar")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setFixedWidth(140)
        btn_close.setStyleSheet("""
            QPushButton {
                background-color: #007BFF;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 10px 16px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #0056b3; }
            QPushButton:pressed { background-color: #004085; }
        """)
        btn_close.clicked.connect(self.reject)

        footer_layout.addStretch()
        footer_layout.addWidget(btn_close)
        footer_layout.addStretch()

        root.addWidget(footer)


def wrap_text_in_paragraph(text, font_size=8, font_name='Helvetica'):
    """
    Converte texto em um objeto Paragraph do ReportLab para quebra automática de linha.
    Garante que o texto não ultrapasse os limites da coluna na tabela.
    """
    # Criar estilo para o parágrafo
    style = ParagraphStyle(
        'CellStyle',
        fontName=font_name,
        fontSize=font_size,
        leading=font_size + 2,
        alignment=TA_LEFT,
        textColor=colors.black,
        wordWrap='CJK',
    )
    
    # Converter texto vazio ou None em string vazia
    if text is None:
        text = ""
    
    # Escapar caracteres especiais HTML
    text = str(text).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    
    # Criar e retornar o parágrafo
    return Paragraph(text, style)


def criar_cabecalho_pmmg(c, width, height, margin):
    """
    Cria o cabeçalho padrão PMMG no PDF.
    
    Args:
        c: Canvas do PDF
        width: Largura da página
        height: Altura da página
        margin: Margem da página
    
    Returns:
        float: Posição Y após o cabeçalho
    """
    # Tentar carregar a fonte Rawline Black, caso não exista usar Helvetica-Bold
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        # Caso a fonte esteja disponível no sistema
        # pdfmetrics.registerFont(TTFont('RawlineBlack', 'RawlineBlack.ttf'))
        # font_name = 'RawlineBlack'
        font_name = 'Helvetica-Bold'  # Fallback
    except:
        font_name = 'Helvetica-Bold'
    
    # Posição inicial do cabeçalho
    y_start = height - margin
    
    # Caminho para a imagem do escudo
    escudo_path = ESCUDO_PATH
    
    # Criar tabela para o cabeçalho (1 linha, 2 colunas)
    # Primeira coluna: 4cm (imagem), Segunda coluna: resto do espaço
    col1_width = 4 * cm
    col2_width = width - (2 * margin) - col1_width
    
    # Preparar imagem do escudo (2cm x 2cm)
    if os.path.exists(escudo_path):
        img = Image(escudo_path, width=2*cm, height=2*cm)
    else:
        # Se a imagem não existir, criar placeholder
        img = ""
    
    # Texto institucional da segunda coluna (via config.json)
    texto_cabecalho = "\n".join(PDF_HEADER_LINES)
    
    # Criar a tabela do cabeçalho
    from reportlab.platypus import Paragraph
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_LEFT
    
    # Estilo para o texto
    style = ParagraphStyle(
        'CabecalhoPMMG',
        parent=getSampleStyleSheet()['Normal'],
        fontName=font_name,
        fontSize=11,
        leading=13,
        alignment=TA_LEFT,
        textColor=colors.black,
    )
    
    texto_formatado = Paragraph(texto_cabecalho.replace('\n', '<br/>'), style)
    
    # Dados da tabela: [linha1]
    if img:
        table_data = [[img, texto_formatado]]
    else:
        table_data = [["[ESCUDO]", texto_formatado]]
    
    # Criar tabela sem bordas visíveis
    header_table = Table(table_data, colWidths=[col1_width, col2_width])
    
    # Estilo: sem bordas, imagem centralizada
    header_style = TableStyle([
        ('ALIGN', (0, 0), (0, 0), 'CENTER'),      # Centralizar imagem na primeira coluna
        ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),    # Alinhar verticalmente ao meio
        ('LEFTPADDING', (0, 0), (0, 0), 0),       # Sem padding à esquerda na coluna da imagem
        ('LEFTPADDING', (1, 0), (1, 0), 20),      # Padding de 20 pontos à esquerda do texto
        ('RIGHTPADDING', (0, 0), (-1, 0), 0),
        ('TOPPADDING', (0, 0), (-1, 0), 0),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 0),
    ])
    
    header_table.setStyle(header_style)
    
    # Desenhar a tabela do cabeçalho
    table_width, table_height = header_table.wrap(width, height)
    header_table.drawOn(c, margin, y_start - table_height)
    
    # Retornar a posição Y após o cabeçalho (com espaço adicional)
    return y_start - table_height - 20


def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(
        """
    CREATE TABLE IF NOT EXISTS documents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        doc_type TEXT NOT NULL,
        doc_number TEXT NOT NULL,
        expiry_date TEXT NOT NULL,
        received_date TEXT,
        nature TEXT,
        location TEXT,
        military_responsible TEXT,
        alert_5_sent INTEGER DEFAULT 0,
        alert_due_sent INTEGER DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """
    )
    # adicionar coluna início se não existir (migrations simples)
    c.execute("PRAGMA table_info(documents)")
    cols = [r[1] for r in c.fetchall()]
    if "received_date" not in cols:
        try:
            c.execute("ALTER TABLE documents ADD COLUMN received_date TEXT")
        except Exception:
            pass
    # adicionar colunas novas se não existirem
    if "nature" not in cols:
        try:
            c.execute("ALTER TABLE documents ADD COLUMN nature TEXT")
        except Exception:
            pass
    if "location" not in cols:
        try:
            c.execute("ALTER TABLE documents ADD COLUMN location TEXT")
        except Exception:
            pass
    if "military_responsible" not in cols:
        try:
            c.execute("ALTER TABLE documents ADD COLUMN military_responsible TEXT")
        except Exception:
            pass
    if "observation" not in cols:
        try:
            c.execute("ALTER TABLE documents ADD COLUMN observation TEXT")
        except Exception:
            pass
    if "result_pos" not in cols:
        try:
            c.execute("ALTER TABLE documents ADD COLUMN result_pos INTEGER DEFAULT 0")
        except Exception:
            pass
    if "reds_number" not in cols:
        try:
            c.execute("ALTER TABLE documents ADD COLUMN reds_number TEXT")
        except Exception:
            pass
    if "reds_date" not in cols:
        try:
            c.execute("ALTER TABLE documents ADD COLUMN reds_date TEXT")
        except Exception:
            pass
    if "militar_reds" not in cols:
        try:
            c.execute("ALTER TABLE documents ADD COLUMN militar_reds TEXT")
        except Exception:
            pass
    if "resultado_reds" not in cols:
        try:
            c.execute("ALTER TABLE documents ADD COLUMN resultado_reds TEXT")
        except Exception:
            pass
    
    # Criar tabela DDU POSITIVADA
    c.execute(
        """
    CREATE TABLE IF NOT EXISTS ddu_positivada (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        doc_id INTEGER,
        tipo TEXT NOT NULL,
        numero TEXT NOT NULL,
        numero_reds TEXT NOT NULL,
        data TEXT NOT NULL,
        militar TEXT NOT NULL,
        resultado TEXT NOT NULL,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (doc_id) REFERENCES documents(id)
    )
    """
    )

    c.execute(
        """
        DELETE FROM ddu_positivada
        WHERE doc_id IS NULL
           OR doc_id NOT IN (SELECT id FROM documents)
        """
    )
    
    conn.commit()
    conn.close()


def add_document(doc_type, doc_number, expiry_date, received_date=None, nature=None, location=None, observation=None, military_responsible=None):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    # garantir que as colunas existem é responsabilidade do init_db
    c.execute(
        "INSERT INTO documents (doc_type, doc_number, expiry_date, received_date, nature, location, observation, military_responsible) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
        (doc_type, doc_number, expiry_date, received_date, nature, location, observation, military_responsible),
    )
    conn.commit()
    conn.close()


def update_document(doc_id, doc_type, doc_number, expiry_date, received_date=None, nature=None, location=None, observation=None, military_responsible=None):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(
        "UPDATE documents SET doc_type=?, doc_number=?, expiry_date=?, received_date=?, nature=?, location=?, observation=?, military_responsible=? WHERE id=?",
        (doc_type, doc_number, expiry_date, received_date, nature, location, observation, military_responsible, doc_id),
    )
    conn.commit()
    conn.close()


def get_document_by_id(doc_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(
        "SELECT id, doc_type, doc_number, nature, location, observation, received_date, expiry_date, military_responsible FROM documents WHERE id=?",
        (doc_id,)
    )
    row = c.fetchone()
    conn.close()
    return row


def list_documents():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(
        "SELECT id, doc_type, doc_number, nature, location, observation, received_date, expiry_date, alert_5_sent, alert_due_sent, military_responsible, result_pos FROM documents ORDER BY expiry_date"
    )
    rows = c.fetchall()
    conn.close()
    return rows


def _format_date_to_br(value):
    if not value:
        return ""
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            parsed = datetime.datetime.strptime(str(value), fmt)
            return parsed.strftime("%d/%m/%Y")
        except Exception:
            continue
    return str(value)


def export_data_json(path=DATA_JSON_PATH, export_expired=True):
    """Exporta documentos para JSON
    
    Args:
        path: Caminho do arquivo JSON
        export_expired: Se True, exporta todos os documentos. Se False, exporta apenas não vencidos.
    
    Returns:
        Tupla com (total_exportados, documentos_validos, documentos_vencidos)
    """
    docs = list_documents()
    payload = []
    today = datetime.date.today()
    valid_count = 0
    expired_count = 0

    for d in docs:
        _, _, dnum, nature, location, observation, rdate, edate, *_ = d
        
        # Tentar parsear a data de validade
        ed = None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                ed = datetime.datetime.strptime(edate, fmt).date()
                break
            except Exception:
                continue
        
        # Contar se está vencido ou válido
        is_expired = ed is not None and ed < today
        
        if is_expired:
            expired_count += 1
        else:
            valid_count += 1
        
        # Se export_expired for False, filtrar documentos vencidos
        if not export_expired and is_expired:
            continue
        
        payload.append(
            {
                "numero": dnum or "",
                "natureza": nature or "",
                "local": location or "",
                "observacao": observation or "",
                "data_inicio": _format_date_to_br(rdate),
                "data_validade": _format_date_to_br(edate),
            }
        )

    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=4)

    return len(payload), valid_count, expired_count


def delete_document(doc_id):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM ddu_positivada WHERE doc_id = ?", (doc_id,))
    c.execute("DELETE FROM documents WHERE id = ?", (doc_id,))
    conn.commit()
    conn.close()


def delete_expired_documents_by_tab(tab_key):
    """Exclui documentos vencidos da aba específica"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    today = datetime.date.today()
    
    expired_ids = []
    
    if tab_key == "valid":
        # A aba "Doc em Andamento" não deve ter documentos vencidos
        # pois eles migram para outras abas
        pass
    
    elif tab_key == "positive":
        # Deleta documentos vencidos positivos
        c.execute("SELECT id, expiry_date, result_pos FROM documents")
        all_docs = c.fetchall()
        
        for doc_id, edate, result_pos in all_docs:
            ed = None
            for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
                try:
                    ed = datetime.datetime.strptime(edate, fmt).date()
                    break
                except Exception:
                    continue
            
            # Vencido e positivo
            if ed and ed < today and result_pos == 1:
                expired_ids.append(doc_id)
    
    elif tab_key == "negative":
        # Deleta documentos vencidos negativos
        c.execute("SELECT id, expiry_date, result_pos FROM documents")
        all_docs = c.fetchall()
        
        for doc_id, edate, result_pos in all_docs:
            ed = None
            for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
                try:
                    ed = datetime.datetime.strptime(edate, fmt).date()
                    break
                except Exception:
                    continue
            
            # Vencido e negativo
            if ed and ed < today and result_pos == -1:
                expired_ids.append(doc_id)
    
    elif tab_key == "expired_no_response":
        # Deleta todos os documentos da aba "Vencidos sem resposta"
        c.execute("SELECT id, expiry_date, result_pos FROM documents")
        all_docs = c.fetchall()
        
        for doc_id, edate, result_pos in all_docs:
            ed = None
            for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
                try:
                    ed = datetime.datetime.strptime(edate, fmt).date()
                    break
                except Exception:
                    continue
            
            # Vencido e sem resposta
            if ed and ed < today and result_pos != 1 and result_pos != -1:
                expired_ids.append(doc_id)
    
    # Excluir documentos
    if expired_ids:
        placeholders = ",".join(["?"] * len(expired_ids))
        c.execute(f"DELETE FROM ddu_positivada WHERE doc_id IN ({placeholders})", expired_ids)
        c.execute(f"DELETE FROM documents WHERE id IN ({placeholders})", expired_ids)
        conn.commit()
    
    conn.close()
    return len(expired_ids)


def update_result_pos(doc_id, value, reds_number=None, reds_date=None, militar_reds=None, resultado_reds=None):
    """Atualiza o valor de result_pos e dados REDS para um documento"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("UPDATE documents SET result_pos = ?, reds_number = ?, reds_date = ?, militar_reds = ?, resultado_reds = ? WHERE id = ?", 
              (value, reds_number, reds_date, militar_reds, resultado_reds, doc_id))
    conn.commit()
    conn.close()


def add_ddu_positivada(doc_id, tipo, numero, numero_reds, data, militar, resultado):
    """Adiciona um registro na tabela DDU POSITIVADA"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(
        "INSERT INTO ddu_positivada (doc_id, tipo, numero, numero_reds, data, militar, resultado) VALUES (?, ?, ?, ?, ?, ?, ?)",
        (doc_id, tipo, numero, numero_reds, data, militar, resultado)
    )
    conn.commit()
    conn.close()


def remove_ddu_positivada(doc_id):
    """Remove um registro da tabela DDU POSITIVADA"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("DELETE FROM ddu_positivada WHERE doc_id = ?", (doc_id,))
    conn.commit()
    conn.close()


def get_all_ddu_positivada():
    """Obtém todos os dados da tabela DDU POSITIVADA"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute(
        """
        SELECT p.tipo, p.numero, p.numero_reds, p.data, p.militar, p.resultado
        FROM ddu_positivada p
        INNER JOIN documents d ON d.id = p.doc_id
        ORDER BY p.data
        """
    )
    results = c.fetchall()
    conn.close()
    return results


def get_document_info(doc_id):
    """Obtém informações básicas de um documento pelo ID"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT doc_type, doc_number FROM documents WHERE id = ?", (doc_id,))
    result = c.fetchone()
    conn.close()
    return result


def get_document_result_data(doc_id):
    """Obtém os dados de resultado (REDS ou BOS) diretamente da tabela documents"""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT result_pos, reds_number, reds_date, militar_reds, resultado_reds FROM documents WHERE id = ?", (doc_id,))
    result = c.fetchone()
    conn.close()
    return result


def generate_pdf_positivo(path="relatorio_ddu_positivada.pdf"):
    """Gera PDF com todas as DDUs Positivadas"""
    dados = get_all_ddu_positivada()
    
    if not dados:
        return False
    
    today = datetime.date.today()
    
    # Usar paisagem para caber todas as colunas
    c = canvas.Canvas(path, pagesize=landscape(A4))
    width, height = landscape(A4)
    margin = 30
    
    # Adicionar cabeçalho padrão PMMG
    y_position = criar_cabecalho_pmmg(c, width, height, margin)
    y_position -= 20
    
    # Título
    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin, y_position, "Relatório de DDU Positivada")
    c.setFont("Helvetica", 10)
    c.drawString(margin, y_position - 20, f"Gerado em: {today.strftime('%d/%m/%Y')}")
    
    y_position = y_position - 50
    
    # Cabeçalhos
    headers = ["TIPO", "NÚMERO", "NÚMERO DO REDS", "DATA", "MILITAR", "RESULTADO"]
    
    # Preparar dados da tabela convertendo para Paragraph para quebra automática
    processed_data = []
    for row in dados:
        processed_row = [wrap_text_in_paragraph(str(cell).upper() if cell else "", font_size=8) for cell in row]
        processed_data.append(processed_row)
    
    # Criar cabeçalhos como Paragraph também
    header_row = [wrap_text_in_paragraph(h, font_size=9, font_name='Helvetica-Bold') for h in headers]
    table_data = [header_row] + processed_data
    
    # Larguras das colunas
    col_widths = [
        2.5 * cm,  # Tipo
        3.5 * cm,  # Número
        4.0 * cm,  # Número do REDS
        2.5 * cm,  # Data
        4.0 * cm,  # Militar
        7.0 * cm   # Resultado
    ]
    
    # Criar tabela
    table = Table(table_data, colWidths=col_widths)
    
    # Estilo da tabela
    style = TableStyle([
        # Cabeçalho
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#28a745')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        
        # Corpo da tabela
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 1), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 5),
        
        # Linhas alternadas
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0f0f0')]),
    ])
    
    table.setStyle(style)
    
    # Calcular altura da tabela e desenhar
    table_width, table_height = table.wrap(width, height)
    table.drawOn(c, margin, y_position - table_height)
    
    c.save()
    return True


def delete_documents(doc_ids):
    """Exclui múltiplos documentos simultaneamente"""
    if not doc_ids:
        return
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    placeholders = ','.join('?' * len(doc_ids))
    c.execute(f"DELETE FROM ddu_positivada WHERE doc_id IN ({placeholders})", doc_ids)
    c.execute(f"DELETE FROM documents WHERE id IN ({placeholders})", doc_ids)
    conn.commit()
    conn.close()


def mark_alert(doc_id, which):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    if which == "5":
        c.execute("UPDATE documents SET alert_5_sent = 1 WHERE id = ?", (doc_id,))
    elif which == "due":
        c.execute("UPDATE documents SET alert_due_sent = 1 WHERE id = ?", (doc_id,))
    conn.commit()
    conn.close()


def generate_pdf_report(path="relatorio_documentos.pdf", include_expired=True):
    docs = list_documents()
    today = datetime.date.today()
    not_expired = []
    expired = []
    
    for d in docs:
        eid, dtype, dnum, nature, location, observation, rdate, edate, a5, ad, military, result_pos = d
        # aceitar tanto ISO (YYYY-MM-DD) quanto BR (DD/MM/YYYY)
        ed = None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                ed = datetime.datetime.strptime(edate, fmt).date()
                break
            except Exception:
                continue
        # parse received date for display
        r = None
        if rdate:
            for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
                try:
                    r = datetime.datetime.strptime(rdate, fmt).date()
                    break
                except Exception:
                    continue
        
        # Preparar dados para a tabela (sem o ID)
        r_display = r.strftime("%d/%m/%Y") if isinstance(r, datetime.date) else ""
        ed_display = ed.strftime("%d/%m/%Y") if isinstance(ed, datetime.date) else str(edate)
        resultado_positivo = "SIM" if result_pos else "NÃO"
        
        row_data = [
            (dtype or "").upper(),
            (dnum or "").upper(),
            (nature or "").upper(),
            (location or "").upper(),
            (observation or "").upper(),
            r_display,
            ed_display,
            (military or "").upper(),
            resultado_positivo
        ]
        
        if ed and ed >= today:
            not_expired.append(row_data)
        else:
            expired.append(row_data)

    # Usar paisagem para caber todas as colunas
    c = canvas.Canvas(path, pagesize=landscape(A4))
    width, height = landscape(A4)
    margin = 30
    
    # Adicionar cabeçalho padrão PMMG
    y_position = criar_cabecalho_pmmg(c, width, height, margin)
    y_position -= 20
    
    # Título
    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin, y_position, "Relatório de Documentos")
    c.setFont("Helvetica", 10)
    c.drawString(margin, y_position - 20, f"Gerado em: {today.strftime('%d/%m/%Y')}")
    
    y_position = y_position - 50
    
    # Cabeçalhos das tabelas
    headers = ["TIPO", "NÚMERO", "NATUREZA", "LOCAL", "OBSERVAÇÃO", "INÍCIO", "VALIDADE", "MILITAR", "RESULTADO"]
    
    # Função para desenhar tabela
    def draw_table(title, data, y_start, color, header_bg_color='#003366'):
        nonlocal y_position
        
        # Título da seção
        c.setFont("Helvetica-Bold", 12)
        c.setFillColor(colors.HexColor(color))
        c.drawString(margin, y_start, title)
        c.setFillColor(colors.black)
        y_start -= 20
        
        if not data:
            c.setFont("Helvetica", 10)
            c.drawString(margin + 10, y_start, "(nenhum)")
            return y_start - 30
        
        # Converter cabeçalhos para Paragraph
        header_row = [wrap_text_in_paragraph(h, font_size=9, font_name='Helvetica-Bold') for h in headers]
        
        # Converter dados para Paragraph
        processed_data = []
        for row in data:
            processed_row = [wrap_text_in_paragraph(str(cell), font_size=8) for cell in row]
            processed_data.append(processed_row)
        
        # Preparar dados da tabela
        table_data = [header_row] + processed_data
        
        # Usar as larguras configuradas (definidas no início do arquivo)
        col_widths = [
            PDF_COL_TIPO_WIDTH * cm,
            PDF_COL_NUMERO_WIDTH * cm,
            PDF_COL_NATUREZA_WIDTH * cm,
            PDF_COL_LOCAL_WIDTH * cm,
            PDF_COL_OBSERVACAO_WIDTH * cm,
            PDF_COL_RECEBIMENTO_WIDTH * cm,
            PDF_COL_VALIDADE_WIDTH * cm,
            PDF_COL_MILITAR_WIDTH * cm,
            PDF_COL_RESULTADO_POS_WIDTH * cm
        ]
        
        # Criar tabela
        table = Table(table_data, colWidths=col_widths)
        
        # Estilo da tabela
        style = TableStyle([
            # Cabeçalho
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(header_bg_color)),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('TOPPADDING', (0, 0), (-1, 0), 8),
            
            # Corpo da tabela
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 1), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 5),
            
            # Linhas alternadas
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f0f0f0')]),
        ])
        
        table.setStyle(style)
        
        # Calcular altura da tabela e desenhar
        table_width, table_height = table.wrap(width, height)
        table.drawOn(c, margin, y_start - table_height)
        
        return y_start - table_height - 30
    
    # Desenhar tabela de documentos a vencer
    y_position = draw_table("DOCUMENTOS A VENCER OU DENTRO DO PRAZO", not_expired, y_position, "#006400", '#003366')
    
    # Desenhar tabela de documentos vencidos apenas se include_expired for True
    if include_expired:
        # Verificar se precisa de nova página
        if y_position < 200:
            c.showPage()
            y_position = height - margin - 30
        
        # Desenhar tabela de documentos vencidos
        y_position = draw_table("DOCUMENTOS VENCIDOS", expired, y_position, "#8B0000", '#8B0000')
    
    c.save()


class ListDocumentsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Lista de Documentos")
        self.resize(1920, 1080)  #tamanho da tela aberta no máximo
        
        # Estilo geral do diálogo
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QMessageBox {
                background-color: #f5f5f5;
            }
            QMessageBox QLabel {
                color: #003366;
                font-size: 12px;
                padding: 10px;
            }
            QMessageBox QPushButton {
                background-color: #007BFF;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                font-size: 12px;
                font-weight: bold;
                min-width: 80px;
            }
            QMessageBox QPushButton:hover {
                background-color: #0056b3;
            }
            QMessageBox QPushButton:pressed {
                background-color: #004085;
            }

        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Cabeçalho
        header_widget = QWidget()
        header_widget.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #003366, stop:1 #004080);
                border-radius: 8px;
                padding: 15px;
            }
        """)
        header_main_layout = QHBoxLayout(header_widget)
        header_main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Título (centralizado)
        title_label = QLabel("📋 Lista Completa de Documentos")
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont("Arial", 14, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: white;")
        
        header_main_layout.addStretch()
        header_main_layout.addWidget(title_label)
        header_main_layout.addStretch()
        
        layout.addWidget(header_widget)

        search_container = QWidget()
        search_container.setStyleSheet("background-color: transparent;")
        search_layout = QHBoxLayout(search_container)
        search_layout.setContentsMargins(0, 0, 0, 0)

        search_label = QLabel("🔍 Pesquisar:")
        search_label.setStyleSheet("font-weight: bold; font-size: 12px; color: #003366;")

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Digite para filtrar documentos (Número, natureza, local, observação)")
        self.search_input.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                border: 2px solid #e0e0e0;
                border-radius: 4px;
                font-size: 12px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 2px solid #007BFF;
            }
        """)
        self.search_input.textChanged.connect(self.handle_filter_inputs_changed)

        period_label = QLabel("Período (Validade):")
        period_label.setStyleSheet("font-weight: bold; font-size: 12px; color: #003366;")

        self.period_start_input = QLineEdit()
        self.period_start_input.setPlaceholderText("DD/MM/AAAA")
        self.period_start_input.setInputMask("00/00/0000")
        self.period_start_input.setFixedWidth(95)
        self.period_start_input.setStyleSheet("""
            QLineEdit {
                padding: 6px;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                font-size: 12px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 1px solid #007BFF;
            }
        """)
        self.period_start_input.textChanged.connect(self.handle_filter_inputs_changed)

        self.period_start_calendar_btn = QToolButton()
        self.period_start_calendar_btn.setText("📅")
        self.period_start_calendar_btn.setCursor(Qt.PointingHandCursor)
        self.period_start_calendar_btn.setToolTip("Selecionar data inicial")
        self.period_start_calendar_btn.setStyleSheet("""
            QToolButton {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                background-color: white;
                padding: 4px 6px;
                font-size: 12px;
            }
            QToolButton:hover {
                border: 1px solid #007BFF;
                background-color: #f0f7ff;
            }
        """)
        self.period_start_calendar_btn.clicked.connect(lambda: self.open_period_calendar(self.period_start_input))

        period_start_container = QWidget()
        period_start_layout = QHBoxLayout(period_start_container)
        period_start_layout.setContentsMargins(0, 0, 0, 0)
        period_start_layout.setSpacing(4)
        period_start_layout.addWidget(self.period_start_input)
        period_start_layout.addWidget(self.period_start_calendar_btn)

        period_to_label = QLabel("até")
        period_to_label.setStyleSheet("font-size: 12px; color: #333333;")

        self.period_end_input = QLineEdit()
        self.period_end_input.setPlaceholderText("DD/MM/AAAA")
        self.period_end_input.setInputMask("00/00/0000")
        self.period_end_input.setFixedWidth(95)
        self.period_end_input.setStyleSheet("""
            QLineEdit {
                padding: 6px;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                font-size: 12px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 1px solid #007BFF;
            }
        """)
        self.period_end_input.textChanged.connect(self.handle_filter_inputs_changed)

        self.period_end_calendar_btn = QToolButton()
        self.period_end_calendar_btn.setText("📅")
        self.period_end_calendar_btn.setCursor(Qt.PointingHandCursor)
        self.period_end_calendar_btn.setToolTip("Selecionar data final")
        self.period_end_calendar_btn.setStyleSheet("""
            QToolButton {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                background-color: white;
                padding: 4px 6px;
                font-size: 12px;
            }
            QToolButton:hover {
                border: 1px solid #007BFF;
                background-color: #f0f7ff;
            }
        """)
        self.period_end_calendar_btn.clicked.connect(lambda: self.open_period_calendar(self.period_end_input))

        period_end_container = QWidget()
        period_end_layout = QHBoxLayout(period_end_container)
        period_end_layout.setContentsMargins(0, 0, 0, 0)
        period_end_layout.setSpacing(4)
        period_end_layout.addWidget(self.period_end_input)
        period_end_layout.addWidget(self.period_end_calendar_btn)

        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_input, 5)
        search_layout.addSpacing(80)
        search_layout.addWidget(period_label)
        search_layout.addWidget(period_start_container)
        search_layout.addWidget(period_to_label)
        search_layout.addWidget(period_end_container)

        layout.addWidget(search_container)
        
        # Checkbox para visualizar documentos vencidos
        checkbox_container = QWidget()
        checkbox_container.setStyleSheet("background-color: transparent;")
        checkbox_layout = QHBoxLayout(checkbox_container)
        checkbox_layout.setContentsMargins(10, 5, 10, 5)
        
        self.chk_show_expired_list = QCheckBox("Visualizar Documentos Vencidos")
        self.chk_show_expired_list.setChecked(False)
        self.chk_show_expired_list.setCursor(Qt.PointingHandCursor)
        self.chk_show_expired_list.setStyleSheet(f"""
            QCheckBox {{
                font-size: 12px;
                font-weight: bold;
                color: #333333;
            }}
            QCheckBox::indicator {{
                width: 16px;
                height: 16px;
                background-color: #f2f2f2;
                border: 1px solid #8a8a8a;
                border-radius: 3px;
            }}
            QCheckBox::indicator:hover {{
                border: 1px solid #6f6f6f;
            }}
            QCheckBox::indicator:checked {{
                background-color: #eaf3ff;
                border: 1px solid #007BFF;
                image: url("{CHECK_ICON_URL}");
            }}
        """)
        self.chk_show_expired_list.stateChanged.connect(self.handle_show_expired_list_changed)
        
        checkbox_layout.addWidget(self.chk_show_expired_list)
        checkbox_layout.addStretch()
        
        layout.addWidget(checkbox_container)
        
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #d0d0d0;
                border-radius: 8px;
                background-color: white;
            }
            QTabBar::tab {
                background: #a9a9a9;
                color: white;
                border: 1px solid #8c8c8c;
                border-bottom: none;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                padding: 10px 18px;
                margin-right: 4px;
                font-weight: bold;
            }
            QTabBar::tab:selected {
                background: #28a745;
                border-color: #1f7f35;
            }
            QTabBar::tab:hover:!selected {
                background: #9a9a9a;
            }
        """)

        self.btn_delete_expired = QPushButton("🗑️ Excluir Vencidos")
        self.btn_delete_expired.setCursor(Qt.PointingHandCursor)
        self.btn_delete_expired.setStyleSheet("""
            QPushButton {
                background-color: #dc3545;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 12px;
                font-weight: bold;
                min-width: 140px;
            }
            QPushButton:hover {
                background-color: #c82333;
            }
            QPushButton:pressed {
                background-color: #bd2130;
            }
        """)
        self.btn_delete_expired.clicked.connect(self.handle_delete_expired)

        self.btn_pdf = QPushButton("📄 Gerar PDF")
        self.btn_pdf.setCursor(Qt.PointingHandCursor)
        self.btn_pdf.setStyleSheet("""
            QPushButton {
                background-color: #007BFF;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-size: 12px;
                font-weight: bold;
                min-width: 130px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:pressed {
                background-color: #004085;
            }
        """)
        self.btn_pdf.clicked.connect(self.export_current_tab_pdf)
        
        # Container para os botões
        corner_widget = QWidget()
        corner_layout = QHBoxLayout(corner_widget)
        corner_layout.setContentsMargins(0, 0, 0, 0)
        corner_layout.setSpacing(40)
        corner_layout.addWidget(self.btn_delete_expired)
        corner_layout.addWidget(self.btn_pdf)
        
        self.tabs.setCornerWidget(corner_widget, Qt.TopRightCorner)

        self.tab_config = [
            ("valid", "Doc em Andamento", "doc_em_andamento"),
            ("positive", "Respondidos Positivos", "respondidos_positivos"),
            ("negative", "Respondidos Negativos", "respondidos_negativos"),
            ("expired_no_response", "Vencidos sem resposta", "vencidos_sem_resposta"),
        ]

        self.tables = {}
        for key, title, table_key in self.tab_config:
            tab = QWidget()
            tab_layout = QVBoxLayout(tab)
            tab_layout.setContentsMargins(12, 12, 12, 12)
            tab_layout.setSpacing(10)

            table = self._create_list_table(table_key)
            self.tables[key] = table
            tab_layout.addWidget(table)

            self.tabs.addTab(tab, title)

        layout.addWidget(self.tabs)

        self.populate_tables()

    def _create_list_table(self, config_key):
        table = CtrlSelectTableWidget(0, 11)
        table.setHorizontalHeaderLabels(["Info", "Tipo", "Número", "Natureza", "Local", "Observação", "Início", "Validade", "Status", "Militar", "Resultado"])
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setSelectionMode(QTableWidget.MultiSelection)
        table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                gridline-color: #e0e0e0;
                background-color: white;
                font-size: 12px;
            }
            QTableWidget::item {
                padding: 8px;
                font-size: 12px;
            }
            QHeaderView::section {
                background-color: #003366;
                color: white;
                padding: 12px;
                border: none;
                font-weight: bold;
                font-size: 13px;
                min-height: 40px;
            }
            QHeaderView::section:hover {
                background-color: #004a99;
            }
        """)
        table.horizontalHeader().setMinimumHeight(40)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        table.setSortingEnabled(True)
        table.horizontalHeader().setSectionsClickable(True)
        table.horizontalHeader().setSortIndicatorShown(True)
        table.horizontalHeader().setCursor(Qt.PointingHandCursor)
        table.horizontalHeader().sectionClicked.connect(
            lambda col, target_table=table: self.block_sort_for_columns(col, target_table)
        )
        
        # Oculta o cabeçalho vertical para evitar seleção por clique fora da regra de CTRL.
        table.verticalHeader().setVisible(False)
        table.verticalHeader().setDefaultSectionSize(35)

        apply_table_column_behavior(table, config_key)
        return table
    
    def block_sort_for_columns(self, column, table):
        """Bloqueia a ordenação apenas para a coluna Info (0)"""
        if column == 0:
            # Desabilita temporariamente a ordenação
            table.setSortingEnabled(False)
            # Reabilita imediatamente para permitir ordenação em outras colunas
            table.setSortingEnabled(True)
    
    def _parse_document_date(self, date_value):
        if not date_value:
            return None
        for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.datetime.strptime(date_value, fmt).date()
            except Exception:
                continue
        return None

    def _parse_period_filter_date(self, date_text):
        text = (date_text or "").strip()
        if not text:
            return None
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(text, fmt).date()
            except Exception:
                continue
        return None

    def handle_filter_inputs_changed(self, _text):
        QTimer.singleShot(0, self.populate_tables)

    def handle_show_expired_list_changed(self, _state):
        """Handler para o checkbox de visualizar documentos vencidos"""
        QTimer.singleShot(0, self.populate_tables)

    def open_period_calendar(self, target_input):
        current_date = self._parse_period_filter_date(target_input.text().strip())
        selected = show_date_picker_dialog(self, initial_date=current_date)
        if selected:
            target_input.setText(selected.strftime("%d/%m/%Y"))

    def _document_matches_filters(self, expiry_date, dnum, nature, location, observation):
        search_text = (self.search_input.text() if hasattr(self, "search_input") else "").strip().lower()
        if search_text:
            fields = [dnum or "", nature or "", location or "", observation or ""]
            searchable = " ".join(fields).lower()
            if search_text not in searchable:
                return False

        start_filter = self._parse_period_filter_date(self.period_start_input.text()) if hasattr(self, "period_start_input") else None
        end_filter = self._parse_period_filter_date(self.period_end_input.text()) if hasattr(self, "period_end_input") else None
        if start_filter and end_filter and start_filter > end_filter:
            start_filter, end_filter = end_filter, start_filter

        if expiry_date is not None:
            if start_filter and expiry_date < start_filter:
                return False
            if end_filter and expiry_date > end_filter:
                return False

        return True

    def _get_result_label(self, result_pos):
        if result_pos == 1:
            return "POSITIVO"
        if result_pos == -1:
            return "NEGATIVO"
        return "NÃO RESP."

    def _get_document_tab_key(self, expiry_date, result_pos, today):
        if result_pos == 1:
            return "positive"
        if result_pos == -1:
            return "negative"
        if expiry_date is None:
            return None
        if expiry_date < today:
            return "expired_no_response"
        return "valid"

    def _get_status_text(self, expiry_date, today):
        if expiry_date is None:
            return "Data inválida"
        if expiry_date < today:
            return "Vencido"
        days = (expiry_date - today).days
        return f"{days} dias"

    def _get_row_bg_color(self, tab_key, row, expiry_date, today):
        # Verificar se documento está vencido
        is_expired = expiry_date is not None and expiry_date < today
        
        if is_expired:
            # Vermelho alternado para documentos vencidos (em qualquer aba)
            return QColor(255, 100, 100) if row % 2 == 0 else QColor(255, 100, 100)
        
        # Cores para documentos não vencidos
        if tab_key == "positive":
            # Verde alternado para Respondidos Positivos
            return QColor(220, 255, 220) if row % 2 == 0 else QColor(180, 235, 180)
        if tab_key == "negative":
            # Azul alternado para Respondidos Negativos
            return QColor(235, 235, 255) if row % 2 == 0 else QColor(210, 210, 245)
        if tab_key == "expired_no_response":
            return QColor(255, 100, 100)  # Vermelho para vencido sem resposta

        # Doc em Andamento não vencidos
        if expiry_date is not None:
            days_until_expiry = (expiry_date - today).days
            if 0 <= days_until_expiry <= 3:
                return QColor(255, 165, 0)  # Laranja para 0-3 dias
            elif 3 < days_until_expiry <= 7:
                return QColor(255, 255, 100)  # Amarelo para 4-7 dias
        return QColor(245, 245, 245) if row % 2 == 1 else QColor(255, 255, 255)
    
    def populate_tables(self):
        for table in self.tables.values():
            table.setSortingEnabled(False)
            table.setRowCount(0)

        docs = list_documents()
        today = datetime.date.today()
        rows_by_tab = {key: 0 for key in self.tables.keys()}
        
        # Verificar se deve mostrar documentos vencidos
        show_expired = self.chk_show_expired_list.isChecked() if hasattr(self, "chk_show_expired_list") else False
        
        for d in docs:
            eid, dtype, dnum, nature, location, observation, rdate, edate, a5, ad, military, result_pos = d

            ed = self._parse_document_date(edate)
            rd = self._parse_document_date(rdate) if rdate else None
            if not self._document_matches_filters(ed, dnum, nature, location, observation):
                continue
            tab_key = self._get_document_tab_key(ed, result_pos, today)
            if tab_key is None or tab_key not in self.tables:
                continue
            
            # Com checkbox desmarcado, oculta vencidos apenas nas abas de respondidos.
            if not show_expired and tab_key in {"positive", "negative"} and ed is not None and ed < today:
                continue

            status = self._get_status_text(ed, today)
            rd_display = rd.strftime("%d/%m/%Y") if isinstance(rd, datetime.date) else ""
            ed_display = ed.strftime("%d/%m/%Y") if isinstance(ed, datetime.date) else str(edate)

            table = self.tables[tab_key]
            row = rows_by_tab[tab_key]
            rows_by_tab[tab_key] += 1
            bg_color = self._get_row_bg_color(tab_key, row, ed, today)
            
            table.insertRow(row)
            
            # Coluna 0: Botão Info (verde para positivo, cinza para negativo)
            if result_pos == 1:
                btn_info = QPushButton("ℹ")
                btn_info.setStyleSheet("""
                    QPushButton {
                        background-color: #28a745;
                        color: white;
                        border: none;
                        border-radius: 3px;
                        font-size: 14px;
                        font-weight: bold;
                        padding: 2px;
                    }
                    QPushButton:hover {
                        background-color: #218838;
                    }
                """)
                btn_info.setCursor(Qt.PointingHandCursor)
                btn_info.clicked.connect(lambda checked, doc_id=eid: self.show_ddu_info(doc_id))
                table.setCellWidget(row, 0, btn_info)
            elif result_pos == -1:
                btn_info = QPushButton("ℹ")
                btn_info.setStyleSheet("""
                    QPushButton {
                        background-color: #808080;
                        color: white;
                        border: none;
                        border-radius: 3px;
                        font-size: 14px;
                        font-weight: bold;
                        padding: 2px;
                    }
                    QPushButton:hover {
                        background-color: #696969;
                    }
                """)
                btn_info.setCursor(Qt.PointingHandCursor)
                btn_info.clicked.connect(lambda checked, doc_id=eid: self.show_ddu_info(doc_id))
                table.setCellWidget(row, 0, btn_info)
            
            items = [
                QTableWidgetItem((dtype or "").upper()),
                QTableWidgetItem((dnum or "").upper()),
                QTableWidgetItem((nature or "").upper()),
                QTableWidgetItem((location or "").upper()),
                QTableWidgetItem(rd_display),
                QTableWidgetItem(ed_display),
                QTableWidgetItem(status),
                QTableWidgetItem((military or "").upper()),
                QTableWidgetItem(self._get_result_label(result_pos))
            ]
            
            # Colunas 1..4 (Tipo, Número, Natureza, Local)
            for col, item in enumerate(items[:4], start=1):
                item.setBackground(bg_color)
                table.setItem(row, col, item)

            # Coluna 5 (Observação): 2 linhas + "LER MAIS" se truncar
            self._set_observation_cell(
                table=table,
                row=row,
                col=5,
                ddu_number=(dnum or "").strip(),
                natureza=(nature or "").strip(),
                observation_text=(observation or ""),
                bg_color=bg_color,
            )

            # Colunas 6..10 (Início, Validade, Status, Militar, Resultado)
            for col, item in enumerate(items[4:], start=6):
                item.setBackground(bg_color)
                table.setItem(row, col, item)
        
        for table in self.tables.values():
            table.setSortingEnabled(True)

    def _set_observation_cell(self, table, row, col, ddu_number, natureza, observation_text, bg_color):
        full_text = (observation_text or "").strip()
        # Item “invisível” mantém ordenação/clipboard e compatibilidade.
        item = QTableWidgetItem(full_text)
        item.setBackground(bg_color)
        table.setItem(row, col, item)

        col_width = max(80, int(table.columnWidth(col)))
        content_width = max(40, col_width - 20)
        metrics = QFontMetrics(table.font())
        clamped, truncated = _clamp_observation_two_lines(full_text, metrics, content_width)

        container = QWidget()
        container.setStyleSheet(f"background-color: {bg_color.name()};")
        layout = QHBoxLayout(container)
        layout.setContentsMargins(6, 2, 6, 2)
        layout.setSpacing(8)

        preview = QTextEdit()
        preview.setReadOnly(True)
        preview.setFrameShape(QFrame.NoFrame)
        preview.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        preview.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        preview.setWordWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        preview.setPlainText(clamped)
        # Remove padding interno para não “descer” o texto e evitar corte.
        preview.setStyleSheet("""
            QTextEdit {
                background: transparent;
                color: #222222;
                font-size: 12px;
                padding: 0px;
                margin: 0px;
            }
        """)
        preview.setContentsMargins(0, 0, 0, 0)
        preview.setViewportMargins(0, 0, 0, 0)
        try:
            preview.document().setDocumentMargin(0)
        except Exception:
            pass
        preview.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        layout.addWidget(preview, 1)

        if truncated:
            content_width_btn = max(40, col_width - 70)
            clamped_btn, _ = _clamp_observation_two_lines(full_text, metrics, content_width_btn)
            preview.setPlainText(clamped_btn)

            btn = QPushButton("LER MAIS")
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #007BFF;
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 6px 10px;
                    font-size: 11px;
                    font-weight: bold;
                }
                QPushButton:hover { background-color: #0056b3; }
                QPushButton:pressed { background-color: #004085; }
            """)
            btn.clicked.connect(
                lambda _checked=False, num=ddu_number, nat=natureza, obs=full_text: ReadMoreObservationDialog(
                    num, nat, obs, parent=self
                ).exec()
            )
            layout.addWidget(btn, 0, Qt.AlignRight | Qt.AlignVCenter)

        table.setCellWidget(row, col, container)

        needed_h = (metrics.lineSpacing() * 2) + 16
        current_h = table.rowHeight(row)
        if current_h < needed_h:
            table.setRowHeight(row, needed_h)
        preview.setFixedHeight(needed_h - 2)

    def _build_tab_rows_for_pdf(self, tab_key):
        docs = list_documents()
        today = datetime.date.today()
        rows = []
        
        # Verificar estado do checkbox para filtrar vencidos
        show_expired = self.chk_show_expired_list.isChecked() if hasattr(self, "chk_show_expired_list") else False

        for d in docs:
            eid, dtype, dnum, nature, location, observation, rdate, edate, a5, ad, military, result_pos = d
            ed = self._parse_document_date(edate)
            if not self._document_matches_filters(ed, dnum, nature, location, observation):
                continue
            current_tab = self._get_document_tab_key(ed, result_pos, today)
            if current_tab != tab_key:
                continue
            
            # Filtrar vencidos nas abas de respondidos se checkbox desmarcado
            if not show_expired and tab_key in {"positive", "negative"} and ed is not None and ed < today:
                continue

            rd = self._parse_document_date(rdate) if rdate else None
            rd_display = rd.strftime("%d/%m/%Y") if isinstance(rd, datetime.date) else ""
            ed_display = ed.strftime("%d/%m/%Y") if isinstance(ed, datetime.date) else str(edate)
            status = self._get_status_text(ed, today)

            rows.append([
                (dtype or "").upper(),
                (dnum or "").upper(),
                (nature or "").upper(),
                (location or "").upper(),
                (observation or "").upper(),
                rd_display,
                ed_display,
                status,
                (military or "").upper(),
                self._get_result_label(result_pos),
            ])

        return rows

    def export_current_tab_pdf(self):
        current_index = self.tabs.currentIndex()
        if current_index < 0 or current_index >= len(self.tab_config):
            QMessageBox.warning(self, "Atenção", "Nenhuma aba selecionada para exportação.")
            return
        tab_key = self.tab_config[current_index][0]
        self.export_tab_pdf(tab_key)

    def export_tab_pdf(self, tab_key):
        tab_titles = {
            "valid": "Doc em Andamento",
            "positive": "Respondidos Positivos",
            "negative": "Respondidos Negativos",
            "expired_no_response": "Vencidos sem resposta",
        }
        default_names = {
            "valid": "doc_em_andamento.pdf",
            "positive": "respondidos_positivos.pdf",
            "negative": "respondidos_negativos.pdf",
            "expired_no_response": "vencidos_sem_resposta.pdf",
        }

        file_path, _ = QFileDialog.getSaveFileName(
            self,
            f"Salvar PDF - {tab_titles.get(tab_key, 'Lista')}",
            os.path.join(os.path.expanduser("~"), "Desktop", default_names.get(tab_key, "relatorio.pdf")),
            "PDF Files (*.pdf)",
        )
        if not file_path:
            return

        if not file_path.lower().endswith(".pdf"):
            file_path += ".pdf"

        try:
            rows = self._build_tab_rows_for_pdf(tab_key)
            self._create_tab_pdf(file_path, tab_titles.get(tab_key, "Lista"), rows)
            QMessageBox.information(self, "Sucesso", f"PDF gerado com sucesso!\n\nSalvo em:\n{file_path}")
        except Exception as err:
            QMessageBox.critical(self, "Erro", f"Falha ao gerar PDF: {err}")

    def handle_delete_expired(self):
        """Handler para o botão Excluir Vencidos"""
        # Obter a aba atual
        current_index = self.tabs.currentIndex()
        if current_index < 0 or current_index >= len(self.tab_config):
            QMessageBox.warning(self, "Atenção", "Nenhuma aba selecionada.")
            return
        
        tab_key = self.tab_config[current_index][0]
        tab_name = self.tab_config[current_index][1]
        
        # Verificar se a aba é "Doc em Andamento"
        if tab_key == "valid":
            QMessageBox.information(self, "Atenção", 
                "A aba 'Doc em Andamento' não contém documentos vencidos.\n"
                "Os documentos vencidos são migrados automaticamente para outras abas.")
            return
        
        dlg = DeleteExpiredConfirmDialog(self)
        if dlg.exec() == QDialog.Accepted:
            # Excluir documentos vencidos apenas da aba atual
            count = delete_expired_documents_by_tab(tab_key)
            
            if count == 0:
                QMessageBox.information(self, "Atenção", 
                    f"Nenhum documento vencido foi encontrado na aba '{tab_name}'.")
            else:
                QMessageBox.information(self, "Sucesso", 
                    f"{count} documento(s) foi/foram excluído(s) da aba '{tab_name}'!")
            
            # Atualizar a tabela da tela LISTAR
            self.populate_tables()
            
            # Chamar o parent (MainWindow) para atualizar a tela principal
            if self.parent():
                if hasattr(self.parent(), 'refresh_list'):
                    self.parent().refresh_list()

    def _create_tab_pdf(self, path, title, rows):
        today = datetime.date.today()
        c = canvas.Canvas(path, pagesize=landscape(A4))
        width, height = landscape(A4)
        margin = 30

        y_position = criar_cabecalho_pmmg(c, width, height, margin)
        y_position -= 20

        c.setFont("Helvetica-Bold", 16)
        c.drawString(margin, y_position, f"Relatório - {title}")
        c.setFont("Helvetica", 10)
        c.drawString(margin, y_position - 20, f"Gerado em: {today.strftime('%d/%m/%Y')}")
        y_position -= 50

        headers = ["TIPO", "NÚMERO", "NATUREZA", "LOCAL", "OBSERVAÇÃO", "INÍCIO", "VALIDADE", "STATUS", "MILITAR", "RESULTADO"]
        
        # Separar documentos válidos e vencidos
        # O STATUS está no índice 7 da row
        valid_rows = []
        expired_rows = []
        
        if rows:
            for row in rows:
                status = str(row[7]).strip().upper() if len(row) > 7 else ""
                if status == "VENCIDO":
                    expired_rows.append(row)
                else:
                    valid_rows.append(row)
        
        # Função para criar uma tabela
        def draw_table_section(section_title, section_rows, y_start, section_color, header_bg_color):
            nonlocal y_position
            
            # Título da seção
            c.setFont("Helvetica-Bold", 12)
            c.setFillColor(colors.HexColor(section_color))
            c.drawString(margin, y_start, section_title)
            c.setFillColor(colors.black)
            y_start -= 20
            
            if not section_rows:
                c.setFont("Helvetica", 10)
                c.drawString(margin + 10, y_start, "(nenhum)")
                return y_start - 30
            
            # Converter cabeçalhos para Paragraph
            header_row = [wrap_text_in_paragraph(h, font_size=9, font_name='Helvetica-Bold') for h in headers]
            
            # Converter dados para Paragraph
            processed_rows = []
            for row in section_rows:
                processed_row = [wrap_text_in_paragraph(str(cell), font_size=8) for cell in row]
                processed_rows.append(processed_row)
            
            table_data = [header_row] + processed_rows
            
            col_widths = [
                1.8 * cm,
                3.2 * cm,
                2.7 * cm,
                4.2 * cm,
                4.2 * cm,
                1.8 * cm,
                1.8 * cm,
                1.8 * cm,
                2.2 * cm,
                2.2 * cm,
            ]
            
            table = Table(table_data, colWidths=col_widths)
            table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(header_bg_color)),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 9),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("FONTSIZE", (0, 1), (-1, -1), 8),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f0f0f0")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]))
            
            table_width, table_height = table.wrap(width, height)
            table.drawOn(c, margin, y_start - table_height)
            
            return y_start - table_height - 30
        
        # Desenhar seção de documentos válidos
        if valid_rows or not expired_rows:
            y_position = draw_table_section(
                "DOCUMENTOS A VENCER OU DENTRO DO PRAZO", 
                valid_rows, 
                y_position, 
                "#006400", 
                '#003366'
            )
        
        # Desenhar seção de documentos vencidos (se houver)
        if expired_rows:
            # Verificar se precisa de nova página
            if y_position < 200:
                c.showPage()
                y_position = height - margin - 30
            
            y_position = draw_table_section(
                "DOCUMENTOS VENCIDOS", 
                expired_rows, 
                y_position, 
                "#8B0000", 
                '#8B0000'
            )
        
        c.save()
    
    def show_ddu_info(self, doc_id):
        """Exibe as informações de resultado (REDS ou BOS)"""
        # Primeiro tenta buscar dados de resultado do documento
        result_data = get_document_result_data(doc_id)
        if not result_data or result_data[0] == 0:
            QMessageBox.warning(self, "Atenção", "Este documento não possui resultado marcado.")
            return
        
        result_pos, numero_reds, data_reds, militar_reds, resultado_reds = result_data
        
        # Buscar informações do documento (tipo e número)
        doc_info = get_document_info(doc_id)
        if not doc_info:
            QMessageBox.warning(self, "Atenção", "Não foi possível obter informações do documento.")
            return
        
        tipo, numero = doc_info
        
        # Criar tupla compatível com DduPositivadaInfoDialog
        data = (tipo, numero, numero_reds, data_reds, militar_reds, resultado_reds)
        dlg = DduPositivadaInfoDialog(data, result_pos, self)
        dlg.exec()


class DeleteExpiredConfirmDialog(QDialog):
    """Diálogo para confirmar exclusão de documentos vencidos"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Excluir Documentos Vencidos")
        self.setFixedSize(450, 300)
        self.setModal(True)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QLabel {
                color: #003366;
                font-weight: bold;
            }
            QLineEdit {
                padding: 8px;
                border: 2px solid #e0e0e0;
                border-radius: 4px;
                font-size: 12px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 2px solid #dc3545;
            }
            QPushButton {
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        
        # Título
        title = QLabel("⚠️ EXCLUIR DOCUMENTOS VENCIDOS")
        title.setStyleSheet("font-size: 14px; color: #dc3545; font-weight: bold;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Mensagem de aviso
        warning_label = QLabel(
            "Esta ação é IRREVERSÍVEL!\n\n"
            "Todos os documentos vencidos serão\n"
            "permanentemente deletados do banco de dados.\n\n"
            "Para confirmar, digite a palavra:"
        )
        warning_label.setStyleSheet("color: #333333; font-size: 11px; font-weight: normal;")
        warning_label.setWordWrap(True)
        layout.addWidget(warning_label)
        
        # Label da palavra confirmação
        confirm_label = QLabel("LIMPAR")
        confirm_label.setStyleSheet("color: #dc3545; font-size: 13px; font-weight: bold; text-align: center;")
        confirm_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(confirm_label)
        
        # Input field
        self.confirm_input = QLineEdit()
        self.confirm_input.setPlaceholderText("Digite LIMPAR aqui")
        self.confirm_input.textChanged.connect(self.update_button_state)
        layout.addWidget(self.confirm_input)
        
        layout.addStretch()
        
        # Botões
        btn_layout = QHBoxLayout()
        
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #6c757d;
                color: white;
                border: none;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #5a6268;
            }
        """)
        btn_cancel.clicked.connect(self.reject)
        
        self.btn_ok = QPushButton("Excluir")
        self.btn_ok.setStyleSheet("""
            QPushButton {
                background-color: #dc3545;
                color: white;
                border: none;
                min-width: 100px;
            }
            QPushButton:hover:!disabled {
                background-color: #c82333;
            }
            QPushButton:disabled {
                background-color: #a9a9a9;
                color: #c0c0c0;
            }
        """)
        self.btn_ok.setEnabled(False)
        self.btn_ok.clicked.connect(self.accept)
        
        btn_layout.addStretch()
        btn_layout.addWidget(btn_cancel)
        btn_layout.addWidget(self.btn_ok)
        btn_layout.addStretch()
        
        layout.addLayout(btn_layout)
    
    def update_button_state(self):
        """Atualiza o estado do botão OK baseado na entrada"""
        is_valid = self.confirm_input.text().strip() == "LIMPAR"
        self.btn_ok.setEnabled(is_valid)


class DduPositivadaInfoDialog(QDialog):
    """Diálogo para exibir informações de resultado (REDS ou BOS)"""
    def __init__(self, data, result_type=1, parent=None):
        super().__init__(parent)
        title_text = "Informações do Resultado Positivo" if result_type == 1 else "Informações do Resultado Negativo"
        self.setWindowTitle(title_text)
        self.resize(450, 350)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QLabel {
                color: #003366;
                font-size: 12px;
            }
            .field-label {
                font-weight: bold;
                color: #003366;
            }
            .field-value {
                background-color: white;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                padding: 8px;
                color: #333;
            }
            QPushButton {
                background-color: #007BFF;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-size: 12px;
                font-weight: bold;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:pressed {
                background-color: #004085;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        
        # Título
        title_text = "📋 Dados do Resultado Positivo" if result_type == 1 else "📋 Dados do Resultado Negativo"
        title = QLabel(title_text)
        title.setStyleSheet("font-size: 16px; color: #003366; font-weight: bold;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Extrair dados
        tipo, numero, numero_reds, data, militar, resultado = data
        
        # Campos de informação
        number_label = "Número do REDS:" if result_type == 1 else "Número do BOS:"
        fields = [
            ("Tipo:", tipo),
            ("Número:", numero),
            (number_label, numero_reds),
            ("Data:", data),
            ("Militar:", militar if militar else "Não informado")
        ]
        
        # Adicionar campo Resultado apenas para resultados positivos
        if result_type == 1:
            fields.append(("Resultado:", resultado))
        
        for label_text, value_text in fields:
            label = QLabel(label_text)
            label.setProperty("class", "field-label")
            value = QLabel(str(value_text) if value_text else "")
            value.setProperty("class", "field-value")
            value.setWordWrap(True)
            layout.addWidget(label)
            layout.addWidget(value)
        
        # Botão fechar
        btn_close = QPushButton("✖ Fechar")
        btn_close.clicked.connect(self.accept)
        btn_close.setCursor(Qt.PointingHandCursor)
        layout.addWidget(btn_close, alignment=Qt.AlignCenter)


class RedsDialog(QDialog):
    """Diálogo para coletar informações do REDS ou BOS"""
    def __init__(self, parent=None, initial_data=None):
        super().__init__(parent)
        self.setWindowTitle("Informações do Resultado")
        self.setFixedSize(450, 480)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QLabel {
                color: #003366;
                font-size: 12px;
                font-weight: bold;
            }
            QLineEdit {
                padding: 8px;
                border: 2px solid #e0e0e0;
                border-radius: 4px;
                font-size: 12px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 2px solid #007BFF;
            }
            QPushButton {
                background-color: #007BFF;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-size: 12px;
                font-weight: bold;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:pressed {
                background-color: #004085;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        
        # Título
        title = QLabel("📋 Informe o tipo de resultado")
        title.setStyleSheet("font-size: 14px; color: #003366; font-weight: bold;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # Radio buttons para tipo de resultado
        radio_container = QWidget()
        radio_layout = QHBoxLayout(radio_container)
        radio_layout.setContentsMargins(0, 0, 0, 0)
        
        self.radio_group = QButtonGroup(self)
        self.radio_group.setExclusive(True)

        self.radio_positivo = QPushButton("Resultado Positivo")
        self.radio_positivo.setCheckable(True)
        self.radio_positivo.setCursor(Qt.PointingHandCursor)
        self.radio_positivo.setFocusPolicy(Qt.NoFocus)
        self.radio_positivo.setStyleSheet("""
            QPushButton {
                color: #007BFF;
                font-weight: bold;
                font-size: 13px;
                border: none;
                background: transparent;
                padding: 0px;
                text-align: left;
            }
            QPushButton:checked {
                color: #007BFF;
            }
        """)
        self.radio_positivo.toggled.connect(self._on_radio_changed)
        self.radio_group.addButton(self.radio_positivo, 1)
        
        self.radio_negativo = QPushButton("Resultado Negativo")
        self.radio_negativo.setCheckable(True)
        self.radio_negativo.setCursor(Qt.PointingHandCursor)
        self.radio_negativo.setFocusPolicy(Qt.NoFocus)
        self.radio_negativo.setStyleSheet("""
            QPushButton {
                color: #dc3545;
                font-weight: bold;
                font-size: 13px;
                border: none;
                background: transparent;
                padding: 0px;
                text-align: left;
            }
            QPushButton:checked {
                color: #dc3545;
            }
        """)
        self.radio_negativo.toggled.connect(self._on_radio_changed)
        self.radio_group.addButton(self.radio_negativo, -1)

        self.box_positivo = QPushButton("")
        self.box_positivo.setFixedSize(16, 16)
        self.box_positivo.setMinimumSize(16, 16)
        self.box_positivo.setMaximumSize(16, 16)
        self.box_positivo.setCursor(Qt.PointingHandCursor)
        self.box_positivo.setFocusPolicy(Qt.NoFocus)
        self.box_positivo.setFlat(True)
        self.box_positivo.clicked.connect(lambda: self.radio_positivo.setChecked(True))
        self.box_positivo.setStyleSheet("""
            QPushButton {
                border: 1px solid #000000;
                border-radius: 0px;
                background-color: white;
                color: #007BFF;
                font-size: 12px;
                font-weight: bold;
                padding: 0px;
                min-width: 16px;
                max-width: 16px;
                min-height: 16px;
                max-height: 16px;
            }
        """)

        self.box_negativo = QPushButton("")
        self.box_negativo.setFixedSize(16, 16)
        self.box_negativo.setMinimumSize(16, 16)
        self.box_negativo.setMaximumSize(16, 16)
        self.box_negativo.setCursor(Qt.PointingHandCursor)
        self.box_negativo.setFocusPolicy(Qt.NoFocus)
        self.box_negativo.setFlat(True)
        self.box_negativo.clicked.connect(lambda: self.radio_negativo.setChecked(True))
        self.box_negativo.setStyleSheet("""
            QPushButton {
                border: 1px solid #000000;
                border-radius: 0px;
                background-color: white;
                color: #dc3545;
                font-size: 12px;
                font-weight: bold;
                padding: 0px;
                min-width: 16px;
                max-width: 16px;
                min-height: 16px;
                max-height: 16px;
            }
        """)

        positive_container = QWidget()
        positive_container_layout = QHBoxLayout(positive_container)
        positive_container_layout.setContentsMargins(0, 0, 0, 0)
        positive_container_layout.setSpacing(6)
        positive_container_layout.addWidget(self.box_positivo)
        positive_container_layout.addWidget(self.radio_positivo)

        negative_container = QWidget()
        negative_container_layout = QHBoxLayout(negative_container)
        negative_container_layout.setContentsMargins(0, 0, 0, 0)
        negative_container_layout.setSpacing(6)
        negative_container_layout.addWidget(self.box_negativo)
        negative_container_layout.addWidget(self.radio_negativo)

        radio_layout.addWidget(positive_container)
        radio_layout.addSpacing(20)
        radio_layout.addWidget(negative_container)
        radio_layout.addStretch()
        
        layout.addWidget(radio_container)
        
        # Container para campos dinâmicos
        self.fields_container = QWidget()
        self.fields_layout = QVBoxLayout(self.fields_container)
        self.fields_layout.setSpacing(10)
        self.fields_layout.setContentsMargins(0, 0, 0, 0)
        
        # Campos para Resultado Positivo
        self.positive_fields = QWidget()
        positive_layout = QVBoxLayout(self.positive_fields)
        positive_layout.setSpacing(10)
        positive_layout.setContentsMargins(0, 0, 0, 0)
        
        reds_label = QLabel("Número do REDS:")
        self.input_reds = QLineEdit()
        self.input_reds.setInputMask("9999-999999999-999")
        self.input_reds.setPlaceholderText("0000-000000000-000")
        positive_layout.addWidget(reds_label)
        positive_layout.addWidget(self.input_reds)
        
        date_pos_label = QLabel("Data:")
        self.input_date_pos = QLineEdit()
        self.input_date_pos.setInputMask("99/99/9999")
        self.input_date_pos.setPlaceholderText("DD/MM/AAAA")
        positive_layout.addWidget(date_pos_label)
        positive_layout.addWidget(self.input_date_pos)
        
        militar_pos_label = QLabel("Militar:")
        self.input_militar_pos = QLineEdit()
        self.input_militar_pos.setMaxLength(40)
        self.input_militar_pos.setPlaceholderText("Nome do militar responsável")
        self.input_militar_pos.textChanged.connect(lambda text: self._format_militar_title_case(text, self.input_militar_pos))
        positive_layout.addWidget(militar_pos_label)
        positive_layout.addWidget(self.input_militar_pos)
        
        resultado_label = QLabel("Resultado:")
        self.input_resultado = QLineEdit()
        self.input_resultado.setMaxLength(100)
        self.input_resultado.setPlaceholderText("Descrição do resultado")
        positive_layout.addWidget(resultado_label)
        positive_layout.addWidget(self.input_resultado)
        
        # Campos para Resultado Negativo
        self.negative_fields = QWidget()
        negative_layout = QVBoxLayout(self.negative_fields)
        negative_layout.setSpacing(10)
        negative_layout.setContentsMargins(0, 0, 0, 0)
        
        bos_label = QLabel("Número do BOS:")
        self.input_bos = QLineEdit()
        self.input_bos.setInputMask("9999-999999999-999")
        self.input_bos.setPlaceholderText("0000-000000000-000")
        negative_layout.addWidget(bos_label)
        negative_layout.addWidget(self.input_bos)
        
        date_neg_label = QLabel("Data:")
        self.input_date_neg = QLineEdit()
        self.input_date_neg.setInputMask("99/99/9999")
        self.input_date_neg.setPlaceholderText("DD/MM/AAAA")
        negative_layout.addWidget(date_neg_label)
        negative_layout.addWidget(self.input_date_neg)
        
        militar_neg_label = QLabel("Militar (Opcional):")
        self.input_militar_neg = QLineEdit()
        self.input_militar_neg.setMaxLength(40)
        self.input_militar_neg.setPlaceholderText("Nome do militar responsável (opcional)")
        self.input_militar_neg.textChanged.connect(lambda text: self._format_militar_title_case(text, self.input_militar_neg))
        negative_layout.addWidget(militar_neg_label)
        negative_layout.addWidget(self.input_militar_neg)
        
        self.fields_layout.addWidget(self.positive_fields)
        self.fields_layout.addWidget(self.negative_fields)
        
        layout.addWidget(self.fields_container)
        
        # Ocultar ambos os campos inicialmente
        self.positive_fields.hide()
        self.negative_fields.hide()
        
        if initial_data:
            result_type = initial_data.get("result_type", 1)
            if result_type == 1:
                self.radio_positivo.setChecked(True)
                self.input_reds.setText(initial_data.get("reds_number", ""))
                self.input_date_pos.setText(initial_data.get("reds_date", ""))
                self.input_militar_pos.setText(initial_data.get("militar_reds", ""))
                self.input_resultado.setText(initial_data.get("resultado_reds", ""))
            else:
                self.radio_negativo.setChecked(True)
                self.input_bos.setText(initial_data.get("reds_number", ""))
                self.input_date_neg.setText(initial_data.get("reds_date", ""))
                self.input_militar_neg.setText(initial_data.get("militar_reds", ""))

        self._update_result_option_styles()
        
        # Botões
        btn_layout = QHBoxLayout()
        btn_ok = QPushButton("✓ Confirmar")
        btn_ok.clicked.connect(self.validate_and_accept)
        btn_ok.setCursor(Qt.PointingHandCursor)
        
        btn_cancel = QPushButton("✖ Cancelar")
        btn_cancel.clicked.connect(self.reject)
        btn_cancel.setCursor(Qt.PointingHandCursor)
        btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #dc3545;
            }
            QPushButton:hover {
                background-color: #c82333;
            }
            QPushButton:pressed {
                background-color: #bd2130;
            }
        """)
        
        btn_layout.addWidget(btn_ok)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)
    
    def _on_radio_changed(self):
        """Alterna entre os campos conforme a seleção do radio button"""
        if self.radio_positivo.isChecked():
            self.positive_fields.show()
            self.negative_fields.hide()
        elif self.radio_negativo.isChecked():
            self.negative_fields.show()
            self.positive_fields.hide()
        else:
            self.positive_fields.hide()
            self.negative_fields.hide()
        self._update_result_option_styles()

    def _update_result_option_styles(self):
        if self.radio_positivo.isChecked():
            self.box_positivo.setText("✓")
            self.box_negativo.setText("")
            self.radio_positivo.setStyleSheet("""
                QPushButton {
                    color: #007BFF;
                    font-weight: bold;
                    font-size: 13px;
                    border: none;
                    background: transparent;
                    padding: 0px;
                    text-align: left;
                }
            """)
            self.radio_negativo.setStyleSheet("""
                QPushButton {
                    color: #808080;
                    font-weight: bold;
                    font-size: 13px;
                    border: none;
                    background: transparent;
                    padding: 0px;
                    text-align: left;
                }
            """)
        elif self.radio_negativo.isChecked():
            self.box_positivo.setText("")
            self.box_negativo.setText("✓")
            self.radio_positivo.setStyleSheet("""
                QPushButton {
                    color: #808080;
                    font-weight: bold;
                    font-size: 13px;
                    border: none;
                    background: transparent;
                    padding: 0px;
                    text-align: left;
                }
            """)
            self.radio_negativo.setStyleSheet("""
                QPushButton {
                    color: #dc3545;
                    font-weight: bold;
                    font-size: 13px;
                    border: none;
                    background: transparent;
                    padding: 0px;
                    text-align: left;
                }
            """)
        else:
            self.box_positivo.setText("")
            self.box_negativo.setText("")
            self.radio_positivo.setStyleSheet("""
                QPushButton {
                    color: #007BFF;
                    font-weight: bold;
                    font-size: 13px;
                    border: none;
                    background: transparent;
                    padding: 0px;
                    text-align: left;
                }
            """)
            self.radio_negativo.setStyleSheet("""
                QPushButton {
                    color: #dc3545;
                    font-weight: bold;
                    font-size: 13px;
                    border: none;
                    background: transparent;
                    padding: 0px;
                    text-align: left;
                }
            """)

    def _format_militar_title_case(self, text, target_input):
        formatted = text.title()
        if text != formatted:
            cursor_pos = target_input.cursorPosition()
            target_input.blockSignals(True)
            target_input.setText(formatted)
            target_input.setCursorPosition(min(cursor_pos, len(formatted)))
            target_input.blockSignals(False)
    
    def validate_and_accept(self):
        """Valida os dados antes de aceitar"""
        if not self.radio_positivo.isChecked() and not self.radio_negativo.isChecked():
            QMessageBox.warning(self, "Atenção", "Por favor, selecione o tipo de resultado (Positivo ou Negativo).")
            return
        
        if self.radio_positivo.isChecked():
            reds_number = self.input_reds.text().strip()
            reds_date = self.input_date_pos.text().strip()
            militar = self.input_militar_pos.text().strip()
            resultado = self.input_resultado.text().strip()
            
            if not reds_number:
                QMessageBox.warning(self, "Atenção", "Por favor, informe o número do REDS.")
                return
            
            if not reds_date:
                QMessageBox.warning(self, "Atenção", "Por favor, informe a data.")
                return
            
            if not militar:
                QMessageBox.warning(self, "Atenção", "Por favor, informe o militar responsável.")
                return
            
            if not resultado:
                QMessageBox.warning(self, "Atenção", "Por favor, informe o resultado.")
                return
            
            try:
                datetime.datetime.strptime(reds_date, "%d/%m/%Y")
            except ValueError:
                QMessageBox.warning(self, "Atenção", "Data inválida. Use o formato DD/MM/AAAA.")
                return
        
        elif self.radio_negativo.isChecked():
            bos_number = self.input_bos.text().strip()
            bos_date = self.input_date_neg.text().strip()
            
            if not bos_number:
                QMessageBox.warning(self, "Atenção", "Por favor, informe o número do BOS.")
                return
            
            if not bos_date:
                QMessageBox.warning(self, "Atenção", "Por favor, informe a data.")
                return
            
            try:
                datetime.datetime.strptime(bos_date, "%d/%m/%Y")
            except ValueError:
                QMessageBox.warning(self, "Atenção", "Data inválida. Use o formato DD/MM/AAAA.")
                return
        
        self.accept()
    
    def get_data(self):
        """Retorna os dados coletados"""
        if self.radio_positivo.isChecked():
            return {
                "result_type": 1,
                "reds_number": self.input_reds.text().strip(),
                "reds_date": self.input_date_pos.text().strip(),
                "militar_reds": self.input_militar_pos.text().strip(),
                "resultado_reds": self.input_resultado.text().strip()
            }
        elif self.radio_negativo.isChecked():
            return {
                "result_type": -1,
                "reds_number": self.input_bos.text().strip(),
                "reds_date": self.input_date_neg.text().strip(),
                "militar_reds": self.input_militar_neg.text().strip(),
                "resultado_reds": ""
            }
        return None


class AddDocumentDialog(QDialog):
    def __init__(self, parent=None, initial_data=None):
        super().__init__(parent)
        _is_edit = initial_data is not None
        self.setWindowTitle("Editar Documento" if _is_edit else "Cadastrar Novo Documento")
        self.resize(500, 400)
        
        # Estilo geral do diálogo
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QMessageBox {
                background-color: #f5f5f5;
            }
            QMessageBox QLabel {
                color: #003366;
                font-size: 12px;
                padding: 10px;
            }
            QMessageBox QPushButton {
                background-color: #007BFF;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                font-size: 12px;
                font-weight: bold;
                min-width: 80px;
            }
            QMessageBox QPushButton:hover {
                background-color: #0056b3;
            }
            QMessageBox QPushButton:pressed {
                background-color: #004085;
            }

        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(15, 15, 15, 15)
        
        # Cabeçalho
        header_widget = QWidget()
        header_widget.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #003366, stop:1 #004080);
                border-radius: 8px;
                padding: 12px;
            }
        """)
        header_layout = QVBoxLayout(header_widget)
        
        title_label = QLabel("✏️ Editar Documento" if _is_edit else "➕ Cadastrar Novo Documento")
        title_label.setAlignment(Qt.AlignCenter)
        title_font = QFont("Arial", 14, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: white;")
        header_layout.addWidget(title_label)
        
        layout.addWidget(header_widget)
        
        # Formulário
        form_widget = QWidget()
        form_widget.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 8px;
                padding: 15px;
            }
        """)
        form_layout = QVBoxLayout(form_widget)
        form_layout.setSpacing(0)
        form_layout.setContentsMargins(10, 10, 10, 10)
        
        # Estilo para labels e inputs
        label_style = "font-weight: bold; font-size: 13px; color: #003366; margin: 0px; padding: 0px;"
        input_style = """
            QLineEdit {
                padding: 6px;
                border: 2px solid #e0e0e0;
                border-radius: 4px;
                font-size: 13px;
                background-color: white;
                margin: 2px 0px 6px 0px;
            }
            QLineEdit:focus {
                border: 2px solid #007BFF;
            }
        """

        lbl_type = QLabel("Tipo de Documento:")
        lbl_type.setStyleSheet(label_style)
        self.input_type = QLineEdit()
        self.input_type.setMaxLength(10)  # Limite baseado em COL_TIPO_WIDTH
        self.input_type.setStyleSheet(input_style)
        
        lbl_number = QLabel("Número:")
        lbl_number.setStyleSheet(label_style)
        self.input_number = QLineEdit()
        self.input_number.setMaxLength(25)  # Limite baseado em COL_NUMERO_WIDTH
        self.input_number.setStyleSheet(input_style)
        
        lbl_nature = QLabel("Natureza:")
        lbl_nature.setStyleSheet(label_style)
        self.input_nature = QLineEdit()
        self.input_nature.setMaxLength(28)  # Limite baseado em COL_NATUREZA_WIDTH
        self.input_nature.setStyleSheet(input_style)
        
        lbl_location = QLabel("Local:")
        lbl_location.setStyleSheet(label_style)
        self.input_location = QLineEdit()
        self.input_location.setMaxLength(60)  # Limite do campo Local
        self.input_location.setStyleSheet(input_style)
        
        lbl_observation = QLabel("Observação:")
        lbl_observation.setStyleSheet(label_style)
        self.input_observation = QTextEdit()
        self.input_observation.setPlaceholderText("Digite as observações (sem limite de caracteres)")
        self.input_observation.setStyleSheet("""
            QTextEdit {
                padding: 8px;
                border: 2px solid #e0e0e0;
                border-radius: 4px;
                font-size: 13px;
                background-color: white;
                margin: 2px 0px 6px 0px;
            }
            QTextEdit:focus {
                border: 2px solid #007BFF;
            }
        """)
        self.input_observation.setMinimumHeight(90)
        
        lbl_received = QLabel("Data de Início:")
        lbl_received.setStyleSheet(label_style)
        self.input_received = QLineEdit()
        self.input_received.setPlaceholderText("DD/MM/AAAA")
        self.input_received.setStyleSheet(input_style)
        try:
            self.input_received.setInputMask("00/00/0000;_")
        except Exception:
            pass
        
        lbl_expiry = QLabel("Data de Validade:")
        lbl_expiry.setStyleSheet(label_style)
        self.input_expiry = QLineEdit()
        self.input_expiry.setPlaceholderText("DD/MM/AAAA")
        self.input_expiry.setStyleSheet(input_style)
        try:
            self.input_expiry.setInputMask("00/00/0000;_")
        except Exception:
            pass
        
        lbl_military = QLabel("Militar Responsável pelo Cumprimento:")
        lbl_military.setStyleSheet(label_style)
        self.input_military = QLineEdit()
        self.input_military.setMaxLength(18)  # Limite baseado em COL_MILITAR_WIDTH
        self.input_military.setStyleSheet(input_style)

        form_layout.addWidget(lbl_type)
        form_layout.addWidget(self.input_type)
        form_layout.addWidget(lbl_number)
        form_layout.addWidget(self.input_number)
        form_layout.addWidget(lbl_nature)
        form_layout.addWidget(self.input_nature)
        form_layout.addWidget(lbl_location)
        form_layout.addWidget(self.input_location)
        form_layout.addWidget(lbl_observation)
        form_layout.addWidget(self.input_observation)
        form_layout.addWidget(lbl_received)
        form_layout.addWidget(self.input_received)
        form_layout.addWidget(lbl_expiry)
        form_layout.addWidget(self.input_expiry)
        form_layout.addWidget(lbl_military)
        form_layout.addWidget(self.input_military)

        # Preencher campos em modo edição
        if initial_data:
            self.input_type.setText(initial_data.get("doc_type", ""))
            self.input_number.setText(initial_data.get("doc_number", ""))
            self.input_nature.setText(initial_data.get("nature", ""))
            self.input_location.setText(initial_data.get("location", ""))
            self.input_observation.setPlainText(initial_data.get("observation", "") or "")
            self.input_received.setText(initial_data.get("received_date", ""))
            self.input_expiry.setText(initial_data.get("expiry_date", ""))
            self.input_military.setText(initial_data.get("military_responsible", ""))
        
        layout.addWidget(form_widget)

        # Botões
        btns = QHBoxLayout()
        btns.setSpacing(10)
        
        btn_ok = QPushButton("✓ Salvar")
        btn_ok.clicked.connect(self.accept)
        btn_ok.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 12px 20px;
                font-size: 14px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #218838;
            }
            QPushButton:pressed {
                background-color: #1e7e34;
            }
        """)
        btn_ok.setCursor(Qt.PointingHandCursor)
        
        btn_cancel = QPushButton("✖ Cancelar")
        btn_cancel.clicked.connect(self.reject)
        btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #dc3545;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 12px 20px;
                font-size: 14px;
                font-weight: bold;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #c82333;
            }
            QPushButton:pressed {
                background-color: #bd2130;
            }
        """)
        btn_cancel.setCursor(Qt.PointingHandCursor)
        
        btns.addStretch()
        btns.addWidget(btn_ok)
        btns.addWidget(btn_cancel)
        btns.addStretch()
        layout.addLayout(btns)

    def get_data(self):
        return {
            "doc_type": self.input_type.text().strip(),
            "doc_number": self.input_number.text().strip(),
            "nature": self.input_nature.text().strip(),
            "location": self.input_location.text().strip(),
            "observation": self.input_observation.toPlainText().strip(),
            "received_date": self.input_received.text().strip(),
            "expiry_date": self.input_expiry.text().strip(),
            "military_responsible": self.input_military.text().strip(),
        }


class CtrlSelectTableWidget(QTableWidget):
    """Tabela que só permite seleção quando CTRL está pressionado"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Garante que o Qt não substitua as barras customizadas em runtime.
        if not isinstance(self.verticalScrollBar(), BlueScrollBar):
            self.setVerticalScrollBar(BlueScrollBar(Qt.Vertical, self))
        if not isinstance(self.horizontalScrollBar(), BlueScrollBar):
            self.setHorizontalScrollBar(BlueScrollBar(Qt.Horizontal, self))

    def _ctrl_pressed(self):
        return bool(QApplication.keyboardModifiers() & Qt.ControlModifier)
        
    def mousePressEvent(self, event):
        # Só permite seleção se CTRL estiver pressionado
        if event.modifiers() & Qt.ControlModifier:
            super().mousePressEvent(event)
        else:
            # Ignora o clique, não seleciona
            event.ignore()

    def mouseDoubleClickEvent(self, event):
        if event.modifiers() & Qt.ControlModifier:
            super().mouseDoubleClickEvent(event)
        else:
            event.ignore()

    def keyPressEvent(self, event):
        if self._ctrl_pressed():
            super().keyPressEvent(event)
        else:
            event.ignore()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Controle de Documentos")
        self.resize(1100, 700)
        if os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))

        central = QWidget()
        self.setCentralWidget(central)
        
        # Definir estilo geral da janela e MessageBoxes
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QWidget {
                background-color: #f5f5f5;
            }
            QMessageBox {
                background-color: #f5f5f5;
            }
            QMessageBox QLabel {
                color: #003366;
                font-size: 14px;
                padding: 10px;
            }
            QMessageBox QPushButton {
                background-color: #007BFF;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                font-size: 12px;
                font-weight: bold;
                min-width: 80px;
            }
            QMessageBox QPushButton:hover {
                background-color: #0056b3;
            }
            QMessageBox QPushButton:pressed {
                background-color: #004085;
            }

        """)

        # ===== CABEÇALHO INSTITUCIONAL =====
        header_widget = QWidget()
        header_widget.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #003366, stop:1 #004080);
                border-radius: 8px;
                padding: 5px;
            }
        """)
        # Layout principal vertical para a barra azul
        header_main_layout = QVBoxLayout(header_widget)
        header_main_layout.setSpacing(2)
        header_main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Layout horizontal para o conteúdo principal
        header_layout = QHBoxLayout()
        
        # Logo PMMG
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignCenter)
        logo_path = ESCUDO_PATH
        try:
            if os.path.exists(logo_path):
                pix = QPixmap(logo_path)
                if not pix.isNull():
                    pix = pix.scaledToWidth(100, Qt.SmoothTransformation)
                    logo_label.setPixmap(pix)
                else:
                    logo_label.setText("PMMG")
                    logo_label.setStyleSheet("color: white; font-size: 24px; font-weight: bold;")
            else:
                logo_label.setText("PMMG")
                logo_label.setStyleSheet("color: white; font-size: 24px; font-weight: bold;")
        except Exception:
            logo_label.setText("PMMG")
            logo_label.setStyleSheet("color: white; font-size: 24px; font-weight: bold;")
        
        # Informações centrais
        title_layout = QVBoxLayout()
        title_layout.setSpacing(2)
        
        subtitle_label = QLabel("Sistema de Controle de Prazo de Documentos")
        subtitle_label.setAlignment(Qt.AlignCenter)
        subtitle_font = QFont("Arial", 18, QFont.Bold)
        subtitle_label.setFont(subtitle_font)
        subtitle_label.setStyleSheet("color: white;")
        
        unit_label = QLabel(APP_SUBTITLE)
        unit_label.setAlignment(Qt.AlignCenter)
        unit_font = QFont("Arial", 16)
        unit_label.setFont(unit_font)
        unit_label.setStyleSheet("color: #e0e0e0;")
        
        title_layout.addWidget(subtitle_label)
        title_layout.addWidget(unit_label)
        

        # Data no lado direito
        info_layout = QVBoxLayout()
        info_layout.setAlignment(Qt.AlignTop)
        
        current_date = datetime.date.today().strftime("%d/%m/%Y")
        date_label = QLabel(f"📅 {current_date}")
        date_label.setStyleSheet("color: white; font-size: 14px; font-weight: bold;")
        
        info_layout.addWidget(date_label)
        
        header_layout.addWidget(logo_label)
        header_layout.addStretch()
        header_layout.addLayout(title_layout)
        header_layout.addStretch()
        header_layout.addLayout(info_layout)
        
        # Layout horizontal para versão no canto inferior direito
        version_layout = QHBoxLayout()
        version_layout.setContentsMargins(0, 0, 0, 0)
        version_layout.addStretch()
        
        version_label = QLabel("Versão 1.6")
        version_label.setStyleSheet("color: #cccccc; font-size: 14px;")
        version_label.setAlignment(Qt.AlignRight | Qt.AlignBottom)
        version_layout.addWidget(version_label)
        
        # Adicionar layouts ao layout principal
        header_main_layout.addLayout(header_layout)
        header_main_layout.addLayout(version_layout)
        
        # ===== BOTÕES DE AÇÃO =====
        btn_widget = QWidget()
        btn_widget.setStyleSheet("""
            QWidget {
                background-color: #2f3338;
                border: 1px solid #4a4f57;
                border-radius: 8px;
                padding: 10px;
            }
        """)
        btn_layout = QHBoxLayout(btn_widget)
        btn_layout.setSpacing(10)
        
        btn_add = QPushButton("➕ Cadastrar Documento")
        btn_add.clicked.connect(self.handle_add)
        btn_delete = QPushButton("🗑️ Excluir Selecionados")
        btn_delete.clicked.connect(self.handle_delete)
        self.btn_edit = QPushButton("✏️ Editar")
        self.btn_edit.clicked.connect(self.handle_edit)
        self.btn_edit.setEnabled(False)
        btn_list = QPushButton("📋 Listar")
        btn_list.clicked.connect(lambda: self.handle_list())
        btn_pdf = QPushButton("📄 PDF Geral")
        btn_pdf.clicked.connect(self.handle_pdf)
        btn_dashboard = QPushButton("📄 Dashboard")
        btn_dashboard.clicked.connect(self.handle_dashboard)
        btn_update_site = QPushButton("Exportar Dados")
        btn_update_site.clicked.connect(lambda: self.handle_update_site())
        
        # Estilização dos botões
        btn_style = """
            QPushButton {
                background-color: #007BFF;
                color: white;
                border: 2px solid #4da3ff;
                border-radius: 6px;
                padding: 12px 12px;
                font-size: 14px;
                font-weight: bold;
                min-width: 140px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:pressed {
                background-color: #004085;
            }
        """
        btn_delete_style = """
            QPushButton {
                background-color: #dc3545;
                color: white;
                border: 2px solid #ff6b79;
                border-radius: 6px;
                padding: 12px 12px;
                font-size: 14px;
                font-weight: bold;
                min-width: 140px;
            }
            QPushButton:hover {
                background-color: #c82333;
            }
            QPushButton:pressed {
                background-color: #bd2130;
            }
        """
        btn_green_style = """
            QPushButton {
                background-color: #28a745;
                color: white;
                border: 2px solid #59d47a;
                border-radius: 6px;
                padding: 12px 12px;
                font-size: 14px;
                font-weight: bold;
                min-width: 140px;
            }
            QPushButton:hover {
                background-color: #218838;
            }
            QPushButton:pressed {
                background-color: #1e7e34;
            }
        """
        btn_orange_style = """
            QPushButton {
                background-color: #ff8c00;
                color: white;
                border: 2px solid #ffb347;
                border-radius: 6px;
                padding: 12px 12px;
                font-size: 14px;
                font-weight: bold;
                min-width: 140px;
            }
            QPushButton:hover {
                background-color: #e07b00;
            }
            QPushButton:pressed {
                background-color: #cc6f00;
            }
        """
        btn_aqua_style = """
            QPushButton {
                background-color: #00cbbf;
                color: white;
                border: 2px solid #72fff5;
                border-radius: 6px;
                padding: 12px 12px;
                font-size: 14px;
                font-weight: bold;
                min-width: 140px;
            }
            QPushButton:hover {
                background-color: #00b9ad;
            }
            QPushButton:pressed {
                background-color: #00a398;
            }
        """

        btn_edit_style = """
            QPushButton {
                background-color: #007BFF;
                color: white;
                border: 2px solid #4da3ff;
                border-radius: 6px;
                padding: 12px 12px;
                font-size: 14px;
                font-weight: bold;
                min-width: 140px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QPushButton:pressed {
                background-color: #004085;
            }
            QPushButton:disabled {
                background-color: #7fb3e8;
                border: 2px solid #a8ccee;
                color: #d0e8ff;
            }
        """
        btn_add.setStyleSheet(btn_style)
        btn_add.setCursor(Qt.PointingHandCursor)
        btn_delete.setStyleSheet(btn_delete_style)
        btn_delete.setCursor(Qt.PointingHandCursor)
        self.btn_edit.setStyleSheet(btn_edit_style)
        self.btn_edit.setCursor(Qt.PointingHandCursor)
        btn_list.setStyleSheet(btn_style)
        btn_list.setCursor(Qt.PointingHandCursor)
        btn_pdf.setStyleSheet(btn_green_style)
        btn_pdf.setCursor(Qt.PointingHandCursor)
        btn_dashboard.setStyleSheet(btn_orange_style)
        btn_dashboard.setCursor(Qt.PointingHandCursor)
        btn_update_site.setStyleSheet(btn_aqua_style)
        btn_update_site.setCursor(Qt.PointingHandCursor)

        btn_layout.addWidget(btn_add)
        btn_layout.addWidget(btn_delete)
        btn_layout.addWidget(self.btn_edit)
        btn_layout.addWidget(btn_list)
        btn_layout.addWidget(btn_pdf)
        btn_layout.addWidget(btn_dashboard)
        btn_layout.addWidget(btn_update_site)
        
        # ===== TABELA DE DOCUMENTOS =====
        table_widget = QWidget()
        table_widget.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 8px;
                padding: 10px;
            }
        """)
        table_layout = QVBoxLayout(table_widget)
        
        title_row = QHBoxLayout()

        table_title = QLabel("📑 Documentos Cadastrados")
        table_title.setStyleSheet("font-weight: bold; font-size: 13px; color: #003366;")

        # Carregar configurações dos checkboxes
        checkbox_config = load_checkbox_config()

        self.chk_show_expired = QCheckBox("Visualizar Documentos Vencidos")
        self.chk_show_expired.setChecked(checkbox_config.get("show_expired", False))
        self.chk_show_expired.setCursor(Qt.PointingHandCursor)
        self.chk_show_expired.setStyleSheet(f"""
            QCheckBox {{
                font-size: 12px;
                font-weight: bold;
                color: #333333;
            }}
            QCheckBox::indicator {{
                width: 16px;
                height: 16px;
                background-color: #f2f2f2;
                border: 1px solid #8a8a8a;
                border-radius: 3px;
            }}
            QCheckBox::indicator:hover {{
                border: 1px solid #6f6f6f;
            }}
            QCheckBox::indicator:checked {{
                background-color: #eaf3ff;
                border: 1px solid #007BFF;
                image: url("{CHECK_ICON_URL}");
            }}
        """)
        self.chk_show_expired.stateChanged.connect(self.handle_show_expired_changed)

        self.chk_export_expired = QCheckBox("Exportar Documentos Vencidos")
        self.chk_export_expired.setChecked(checkbox_config.get("export_expired", False))
        self.chk_export_expired.setCursor(Qt.PointingHandCursor)
        self.chk_export_expired.setStyleSheet(f"""
            QCheckBox {{
                font-size: 12px;
                font-weight: bold;
                color: #333333;
            }}
            QCheckBox::indicator {{
                width: 16px;
                height: 16px;
                background-color: #f2f2f2;
                border: 1px solid #8a8a8a;
                border-radius: 3px;
            }}
            QCheckBox::indicator:hover {{
                border: 1px solid #6f6f6f;
            }}
            QCheckBox::indicator:checked {{
                background-color: #eaf3ff;
                border: 1px solid #007BFF;
                image: url("{CHECK_ICON_URL}");
            }}
        """)
        self.chk_export_expired.stateChanged.connect(self.handle_export_expired_changed)

        # Botão Suporte
        self.btn_support = QPushButton("Suporte")
        self.btn_support.setCursor(Qt.PointingHandCursor)
        self.btn_support.setStyleSheet("""
            QPushButton {
                color: #000000;
                background-color: white;
                border: 2px solid #007BFF;
                border-radius: 6px;
                padding: 5px 12px;
                font-size: 11px;
                font-weight: bold;
                min-width: 70px;
            }
            QPushButton:hover {
                background-color: #f0f7ff;
                border: 2px solid #0056b3;
                color: #0056b3;
            }
            QPushButton:pressed {
                background-color: #e3f0ff;
            }
        """)
        self.btn_support.clicked.connect(self.show_support_dialog)

        title_row.addWidget(table_title)
        title_row.addWidget(self.chk_show_expired)
        title_row.addSpacing(30)
        title_row.addWidget(self.chk_export_expired)
        title_row.addStretch()
        title_row.addWidget(self.btn_support)

        period_label = QLabel("Período (Validade):")
        period_label.setStyleSheet("font-weight: bold; font-size: 12px; color: #003366;")

        self.period_start_input = QLineEdit()
        self.period_start_input.setPlaceholderText("DD/MM/AAAA")
        self.period_start_input.setInputMask("00/00/0000")
        self.period_start_input.setFixedWidth(95)
        self.period_start_input.setStyleSheet("""
            QLineEdit {
                padding: 6px;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                font-size: 12px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 1px solid #007BFF;
            }
        """)
        self.period_start_input.textChanged.connect(self.handle_period_filter_changed)

        self.period_start_calendar_btn = QToolButton()
        self.period_start_calendar_btn.setText("📅")
        self.period_start_calendar_btn.setCursor(Qt.PointingHandCursor)
        self.period_start_calendar_btn.setToolTip("Selecionar data inicial")
        self.period_start_calendar_btn.setStyleSheet("""
            QToolButton {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                background-color: white;
                padding: 4px 6px;
                font-size: 12px;
            }
            QToolButton:hover {
                border: 1px solid #007BFF;
                background-color: #f0f7ff;
            }
        """)
        self.period_start_calendar_btn.clicked.connect(lambda: self.open_period_calendar(self.period_start_input))

        period_start_container = QWidget()
        period_start_layout = QHBoxLayout(period_start_container)
        period_start_layout.setContentsMargins(0, 0, 0, 0)
        period_start_layout.setSpacing(4)
        period_start_layout.addWidget(self.period_start_input)
        period_start_layout.addWidget(self.period_start_calendar_btn)

        period_to_label = QLabel("até")
        period_to_label.setStyleSheet("font-size: 12px; color: #333333;")

        self.period_end_input = QLineEdit()
        self.period_end_input.setPlaceholderText("DD/MM/AAAA")
        self.period_end_input.setInputMask("00/00/0000")
        self.period_end_input.setFixedWidth(95)
        self.period_end_input.setStyleSheet("""
            QLineEdit {
                padding: 6px;
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                font-size: 12px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 1px solid #007BFF;
            }
        """)
        self.period_end_input.textChanged.connect(self.handle_period_filter_changed)

        self.period_end_calendar_btn = QToolButton()
        self.period_end_calendar_btn.setText("📅")
        self.period_end_calendar_btn.setCursor(Qt.PointingHandCursor)
        self.period_end_calendar_btn.setToolTip("Selecionar data final")
        self.period_end_calendar_btn.setStyleSheet("""
            QToolButton {
                border: 1px solid #d0d0d0;
                border-radius: 4px;
                background-color: white;
                padding: 4px 6px;
                font-size: 12px;
            }
            QToolButton:hover {
                border: 1px solid #007BFF;
                background-color: #f0f7ff;
            }
        """)
        self.period_end_calendar_btn.clicked.connect(lambda: self.open_period_calendar(self.period_end_input))

        period_end_container = QWidget()
        period_end_layout = QHBoxLayout(period_end_container)
        period_end_layout.setContentsMargins(0, 0, 0, 0)
        period_end_layout.setSpacing(4)
        period_end_layout.addWidget(self.period_end_input)
        period_end_layout.addWidget(self.period_end_calendar_btn)

        table_layout.addLayout(title_row)
        
        # Barra de pesquisa
        search_container = QWidget()
        search_container.setStyleSheet("background-color: transparent;")
        search_layout = QHBoxLayout(search_container)
        search_layout.setContentsMargins(0, 5, 0, 5)
        
        search_label = QLabel("🔍 Pesquisar:")
        search_label.setStyleSheet("font-weight: bold; font-size: 12px; color: #003366;")
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Digite para filtrar documentos (Número, natureza, local, observação)")
        self.search_input.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                border: 2px solid #e0e0e0;
                border-radius: 4px;
                font-size: 12px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 2px solid #007BFF;
            }
        """)
        self.search_input.textChanged.connect(self.filter_table)
        
        search_layout.addWidget(search_label)
        search_layout.addWidget(self.search_input, 5)  # Proporção 5
        search_layout.addSpacing(100)  # Redução de 100 pixels
        search_layout.addWidget(period_label)
        search_layout.addWidget(period_start_container)
        search_layout.addWidget(period_to_label)
        search_layout.addWidget(period_end_container)
        search_layout.addSpacing(60)  # Afasta 60 pixels da borda direita
        search_layout.addStretch()  # Espaço livre à direita
        
        table_layout.addWidget(search_container)
        
        self.table = CtrlSelectTableWidget(0, 11)
        self.table.setHorizontalHeaderLabels(["Info", "Tipo", "Número", "Natureza", "Local", "Observação", "Início", "Validade", "Status", "Militar", "Resultado"])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.MultiSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                gridline-color: #e0e0e0;
                background-color: white;
                font-size: 12px;
            }
            QTableWidget::item {
                padding: 8px;
                font-size: 12px;
            }
            QHeaderView::section {
                background-color: #003366;
                color: white;
                padding: 12px 6px;
                border: none;
                font-weight: bold;
                font-size: 13px;
                min-height: 40px;
                text-align: center;
            }
            QHeaderView::section:hover {
                background-color: #0056b3;
            }
        """)
        
        # Oculta o cabeçalho vertical para impedir seleção de linha sem CTRL.
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(35)
        self.table.horizontalHeader().setMinimumHeight(40)
        self.table.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.table.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.table.cellDoubleClicked.connect(self.handle_locked_result_pos_double_click)
        self._locked_result_edit_in_progress = False
        self.table.itemSelectionChanged.connect(self.update_edit_button_state)

        apply_table_column_behavior(self.table, "main_table", local_elastic=True)
        
        table_layout.addWidget(self.table)
        
        # ===== RODAPÉ =====
        footer_label = QLabel(APP_FOOTER)
        footer_label.setAlignment(Qt.AlignCenter)
        footer_label.setStyleSheet("font-size: 12px; color: #222222; margin-top: 5px;")
        
        # ===== LAYOUT PRINCIPAL =====
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(15, 15, 15, 15)
        
        main_layout.addWidget(header_widget)
        main_layout.addWidget(btn_widget)
        main_layout.addWidget(table_widget)
        main_layout.addWidget(footer_label)
        
        self.refresh_list()
        
        # Timer para alertas a cada 60 segundos
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_alerts)
        self.timer.start(60 * 1000)
        
        # checar logo ao iniciar
        QTimer.singleShot(1000, self.check_alerts)

    
    def update_edit_button_state(self):
        selected_rows = {item.row() for item in self.table.selectedItems()}
        self.btn_edit.setEnabled(len(selected_rows) == 1)

    def handle_edit(self):
        selected_rows = {item.row() for item in self.table.selectedItems()}
        if len(selected_rows) != 1:
            QMessageBox.warning(self, "Atenção", "Selecione exatamente um documento para editar.")
            return

        row = next(iter(selected_rows))
        item = self.table.item(row, 1)  # Coluna 1 (Tipo) armazena doc_id em UserRole
        if not item:
            return
        doc_id = item.data(Qt.UserRole)
        if doc_id is None:
            return

        doc = get_document_by_id(doc_id)
        if not doc:
            QMessageBox.warning(self, "Atenção", "Documento não encontrado no banco de dados.")
            return

        # (id, doc_type, doc_number, nature, location, observation, received_date, expiry_date, military_responsible)
        _, doc_type, doc_number, nature, location, observation, received_date, expiry_date, military_responsible = doc

        initial_data = {
            "doc_type": doc_type or "",
            "doc_number": doc_number or "",
            "nature": nature or "",
            "location": location or "",
            "observation": observation or "",
            "received_date": _format_date_to_br(received_date),
            "expiry_date": _format_date_to_br(expiry_date),
            "military_responsible": military_responsible or "",
        }

        dlg = AddDocumentDialog(self, initial_data=initial_data)
        if dlg.exec() != QDialog.Accepted:
            return

        data = dlg.get_data()
        dtype = data.get("doc_type", "")
        dnum = data.get("doc_number", "")
        nature_new = data.get("nature", "")
        location_new = data.get("location", "")
        observation_new = data.get("observation", "")
        rdate = data.get("received_date", "")
        edate = data.get("expiry_date", "")
        military_new = data.get("military_responsible", "")

        if not dtype or not dnum or not edate or not rdate:
            QMessageBox.warning(self, "Atenção", "Preencha todos os campos obrigatórios.")
            return

        try:
            ed = datetime.datetime.strptime(edate, "%d/%m/%Y").date()
            rd = datetime.datetime.strptime(rdate, "%d/%m/%Y").date()
        except Exception:
            QMessageBox.critical(self, "Erro", "Formato de data inválido. Use DD/MM/AAAA")
            return

        update_document(
            doc_id, dtype, dnum,
            ed.strftime("%Y-%m-%d"), rd.strftime("%Y-%m-%d"),
            nature=nature_new or None,
            location=location_new or None,
            observation=observation_new or None,
            military_responsible=military_new or None,
        )
        QMessageBox.information(self, "Sucesso", "Documento atualizado com sucesso.")
        self.refresh_list()

    def handle_add(self):
        dlg = AddDocumentDialog(self)
        if dlg.exec() != QDialog.Accepted:
            return
        data = dlg.get_data()
        dtype = data.get("doc_type", "")
        dnum = data.get("doc_number", "")
        nature = data.get("nature", "")
        location = data.get("location", "")
        observation = data.get("observation", "")
        rdate = data.get("received_date", "")
        edate = data.get("expiry_date", "")
        military = data.get("military_responsible", "")
        if not dtype or not dnum or not edate or not rdate:
            QMessageBox.warning(self, "Atenção", "Preencha todos os campos")
            return
        try:
            ed = datetime.datetime.strptime(edate, "%d/%m/%Y").date()
            rd = datetime.datetime.strptime(rdate, "%d/%m/%Y").date()
        except Exception:
            QMessageBox.critical(self, "Erro", "Formato de data inválido. Use DD/MM/AAAA")
            return
        iso = ed.strftime("%Y-%m-%d")
        r_iso = rd.strftime("%Y-%m-%d")
        add_document(dtype, dnum, iso, r_iso, nature=nature or None, location=location or None, observation=observation or None, military_responsible=military or None)
        QMessageBox.information(self, "Sucesso", "Documento cadastrado")
        self.refresh_list()

    def handle_show_expired_changed(self, _state):
        if not hasattr(self, "table"):
            return
        # Salvar estado do checkbox
        self.save_checkbox_states()
        QTimer.singleShot(0, self.refresh_list)

    def handle_export_expired_changed(self, _state):
        # Salvar estado do checkbox
        self.save_checkbox_states()

    def save_checkbox_states(self):
        """Salva o estado atual dos checkboxes"""
        config = {
            "show_expired": self.chk_show_expired.isChecked() if hasattr(self, "chk_show_expired") else False,
            "export_expired": self.chk_export_expired.isChecked() if hasattr(self, "chk_export_expired") else False,
        }
        save_checkbox_config(config)

    def show_support_dialog(self):
        """Abre janela elegante de suporte técnico"""
        dlg = QDialog(self)
        dlg.setWindowTitle("Suporte Técnico")
        dlg.setModal(True)
        dlg.setFixedSize(420, 420)
        
        dlg.setStyleSheet("""
            QDialog {
                background-color: #f8f9fa;
            }
        """)
        
        layout = QVBoxLayout(dlg)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)
        
        # Título principal
        title = QLabel("Suporte Técnico")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size: 15px; font-weight: bold; color: #1a1a1a;")
        layout.addWidget(title)
        
        # Subtítulo
        subtitle = QLabel("Desenvolvimento de Softwares")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("font-size: 12px; color: #555555;")
        layout.addWidget(subtitle)
        
        # Separador
        separator1 = QFrame()
        separator1.setFrameShape(QFrame.HLine)
        separator1.setStyleSheet("background-color: #ddd;")
        layout.addWidget(separator1)
        
        # Nome do desenvolvedor
        dev_name = QLabel("Lucas Pires Franco")
        dev_name.setAlignment(Qt.AlignCenter)
        dev_name.setStyleSheet("font-size: 14px; font-weight: bold; color: #1a1a1a;")
        layout.addWidget(dev_name)
        
        # Curso
        dev_course = QLabel("Engenharia de Computação")
        dev_course.setAlignment(Qt.AlignCenter)
        dev_course.setStyleSheet("font-size: 12px; color: #555555;")
        layout.addWidget(dev_course)
        
        # Separador
        separator2 = QFrame()
        separator2.setFrameShape(QFrame.HLine)
        separator2.setStyleSheet("background-color: #ddd;")
        layout.addWidget(separator2)
        
        # Telefone
        phone_label = QLabel("Contato: (035) 99746-8037")
        phone_label.setAlignment(Qt.AlignCenter)
        phone_label.setStyleSheet("font-size: 12px; color: #333333;")
        layout.addWidget(phone_label)
        
        # Email
        email_label = QLabel("e-mail: lpfranco09@gmail.com")
        email_label.setAlignment(Qt.AlignCenter)
        email_label.setStyleSheet("font-size: 12px; color: #333333;")
        layout.addWidget(email_label)
        
        # Botão fechar
        layout.addStretch()
        btn_close = QPushButton("Fechar")
        btn_close.setCursor(Qt.PointingHandCursor)
        btn_close.setStyleSheet("""
            QPushButton {
                background-color: #007BFF;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """)
        btn_close.setFixedWidth(100)
        btn_close.clicked.connect(dlg.accept)
        layout.addWidget(btn_close, alignment=Qt.AlignHCenter)
        
        dlg.exec()

    def _parse_period_filter_date(self, date_text):
        text = (date_text or "").strip()
        if not text:
            return None

        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(text, fmt).date()
            except Exception:
                continue
        return None

    def handle_period_filter_changed(self, _text):
        if not hasattr(self, "table"):
            return
        QTimer.singleShot(0, self.refresh_list)

    def open_period_calendar(self, target_input):
        current_text = target_input.text().strip()
        current_date = self._parse_period_filter_date(current_text)
        selected = show_date_picker_dialog(self, initial_date=current_date)
        if selected:
            target_input.setText(selected.strftime("%d/%m/%Y"))

    def refresh_list(self):
        docs = list_documents()
        today = datetime.date.today()
        show_expired = self.chk_show_expired.isChecked() if hasattr(self, "chk_show_expired") else False
        start_filter = self._parse_period_filter_date(self.period_start_input.text()) if hasattr(self, "period_start_input") else None
        end_filter = self._parse_period_filter_date(self.period_end_input.text()) if hasattr(self, "period_end_input") else None

        if start_filter and end_filter and start_filter > end_filter:
            start_filter, end_filter = end_filter, start_filter

        self.table.setRowCount(0)

        if not docs:
            return

        for d in docs:
            eid, dtype, dnum, nature, location, observation, rdate, edate, a5, ad, military, result_pos = d

            ed = None
            rd = None
            for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
                try:
                    ed = datetime.datetime.strptime(edate, fmt).date()
                    break
                except Exception:
                    continue

            if rdate:
                for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
                    try:
                        rd = datetime.datetime.strptime(rdate, fmt).date()
                        break
                    except Exception:
                        continue

            if isinstance(ed, datetime.date) and ed < today and not show_expired:
                continue

            if start_filter or end_filter:
                if not isinstance(ed, datetime.date):
                    continue
                if start_filter and ed < start_filter:
                    continue
                if end_filter and ed > end_filter:
                    continue

            if ed is None:
                status = "Data inválida"
            elif ed < today:
                status = "Vencido"
            else:
                status = f"{(ed - today).days} dias"

            row = self.table.rowCount()
            self.table.insertRow(row)

            if result_pos == 1:
                btn_info = QPushButton("ℹ")
                btn_info.setStyleSheet("""
                    QPushButton {
                        background-color: #28a745;
                        color: white;
                        border: none;
                        border-radius: 3px;
                        font-size: 14px;
                        font-weight: bold;
                        padding: 2px;
                    }
                    QPushButton:hover {
                        background-color: #218838;
                    }
                """)
                btn_info.setCursor(Qt.PointingHandCursor)
                btn_info.setFocusPolicy(Qt.NoFocus)
                btn_info.clicked.connect(lambda checked, doc_id=eid: self.handle_info_button_click(doc_id))
                self.table.setCellWidget(row, 0, btn_info)
            elif result_pos == -1:
                btn_info = QPushButton("ℹ")
                btn_info.setStyleSheet("""
                    QPushButton {
                        background-color: #808080;
                        color: white;
                        border: none;
                        border-radius: 3px;
                        font-size: 14px;
                        font-weight: bold;
                        padding: 2px;
                    }
                    QPushButton:hover {
                        background-color: #696969;
                    }
                """)
                btn_info.setCursor(Qt.PointingHandCursor)
                btn_info.setFocusPolicy(Qt.NoFocus)
                btn_info.clicked.connect(lambda checked, doc_id=eid: self.handle_info_button_click(doc_id))
                self.table.setCellWidget(row, 0, btn_info)

            item_tipo = QTableWidgetItem((dtype or "").upper())
            item_tipo.setData(Qt.UserRole, eid)
            self.table.setItem(row, 1, item_tipo)
            self.table.setItem(row, 2, QTableWidgetItem((dnum or "").upper()))
            self.table.setItem(row, 3, QTableWidgetItem((nature or "").upper()))
            self.table.setItem(row, 4, QTableWidgetItem((location or "").upper()))
            # Observação: 2 linhas na célula + "LER MAIS" se truncar
            self._set_observation_cell(
                table=self.table,
                row=row,
                col=5,
                ddu_number=(dnum or "").strip(),
                natureza=(nature or "").strip(),
                observation_text=(observation or ""),
            )

            rd_display = rd.strftime("%d/%m/%Y") if isinstance(rd, datetime.date) else ""
            ed_display = ed.strftime("%d/%m/%Y") if isinstance(ed, datetime.date) else str(edate)
            self.table.setItem(row, 6, QTableWidgetItem(rd_display))
            self.table.setItem(row, 7, QTableWidgetItem(ed_display))
            self.table.setItem(row, 8, QTableWidgetItem(status))
            self.table.setItem(row, 9, QTableWidgetItem((military or "").upper()))

            for col in range(1, 10):
                cell_item = self.table.item(row, col)
                if cell_item:
                    cell_item.setTextAlignment(Qt.AlignCenter)

            checkbox = QCheckBox()
            checkbox.setChecked(bool(result_pos))
            checkbox.setEnabled(not bool(result_pos))
            if result_pos:
                checkbox.setToolTip("Resultado Positivo confirmado e bloqueado para edição.")
            checkbox.setStyleSheet(f"""
                QCheckBox::indicator {{
                    width: {CHECKBOX_SIZE}px;
                    height: {CHECKBOX_SIZE}px;
                    border: 1px solid #8a8a8a;
                    border-radius: 3px;
                    background-color: #f2f2f2;
                }}
                QCheckBox::indicator:checked {{
                    background-color: #eaf3ff;
                    border: 1px solid #007BFF;
                    image: url("{CHECK_ICON_URL}");
                }}
                QCheckBox::indicator:hover {{
                    border: 1px solid #6f6f6f;
                }}
                QCheckBox::indicator:disabled {{
                    background-color: #eaf3ff;
                    border: 1px solid #007BFF;
                }}
            """)
            checkbox.stateChanged.connect(lambda state, doc_id=eid, cb=checkbox: self.handle_result_pos_change(doc_id, state, cb))

            container = QWidget()
            container_layout = QHBoxLayout(container)
            container_layout.addWidget(checkbox)
            container_layout.setAlignment(Qt.AlignCenter)
            container_layout.setContentsMargins(0, 0, 0, 0)

            if result_pos:
                container.setCursor(Qt.PointingHandCursor)
                checkbox.setCursor(Qt.PointingHandCursor)
                container.setProperty("locked_result_doc_id", int(eid))
                checkbox.setProperty("locked_result_doc_id", int(eid))
                checkbox.setAttribute(Qt.WA_TransparentForMouseEvents, True)
                container.setAttribute(Qt.WA_NoMousePropagation, True)
                checkbox.setAttribute(Qt.WA_NoMousePropagation, True)

            self.table.setCellWidget(row, 10, container)

            try:
                if isinstance(ed, datetime.date) and ed < today:
                    bg_color = QColor(255, 100, 100)  # Vermelho para vencidos
                elif isinstance(ed, datetime.date):
                    days_until_expiry = (ed - today).days
                    if 0 <= days_until_expiry <= 3:
                        bg_color = QColor(255, 165, 0)  # Laranja para 0-3 dias
                    elif 3 < days_until_expiry <= 7:
                        bg_color = QColor(255, 255, 100)  # Amarelo para 4-7 dias
                    else:
                        bg_color = QColor(245, 245, 245) if row % 2 == 1 else QColor(255, 255, 255)
                else:
                    bg_color = QColor(245, 245, 245) if row % 2 == 1 else QColor(255, 255, 255)

                for col in range(1, self.table.columnCount() - 1):
                    it = self.table.item(row, col)
                    if it:
                        it.setBackground(bg_color)
                # Aplica cor de fundo no widget da observação (col 5), se existir
                obs_widget = self.table.cellWidget(row, 5)
                if obs_widget:
                    obs_widget.setStyleSheet(f"QWidget {{ background-color: {bg_color.name()}; }}")
            except Exception:
                pass

        self.filter_table()

    def _set_observation_cell(self, table, row, col, ddu_number, natureza, observation_text):
        """Renderiza Observação como 2 linhas + reticências + botão 'LER MAIS'."""
        full_text = (observation_text or "").strip()
        # Mantém item para filtro/ordenação (mesmo que o widget seja exibido).
        item = QTableWidgetItem(full_text)
        table.setItem(row, col, item)

        # Largura disponível da coluna (tira padding aproximado do widget)
        col_width = max(80, int(table.columnWidth(col)))
        # Reserva espaço pro botão apenas quando houver truncamento; como ainda não sabemos,
        # calcula um preview com uma reserva conservadora pequena.
        content_width = max(40, col_width - 20)

        metrics = QFontMetrics(table.font())
        clamped, truncated = _clamp_observation_two_lines(full_text, metrics, content_width)

        container = QWidget()
        container.setStyleSheet("QWidget { background: transparent; }")
        container_layout = QHBoxLayout(container)
        container_layout.setContentsMargins(6, 2, 6, 2)
        container_layout.setSpacing(8)

        preview = QTextEdit()
        preview.setReadOnly(True)
        preview.setFrameShape(QFrame.NoFrame)
        preview.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        preview.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        preview.setWordWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        preview.setPlainText(clamped)
        # Remove padding interno para não “descer” o texto e evitar corte.
        preview.setStyleSheet("""
            QTextEdit {
                background: transparent;
                color: #222222;
                font-size: 12px;
                padding: 0px;
                margin: 0px;
            }
        """)
        preview.setContentsMargins(0, 0, 0, 0)
        preview.setViewportMargins(0, 0, 0, 0)
        try:
            preview.document().setDocumentMargin(0)
        except Exception:
            pass
        preview.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        container_layout.addWidget(preview, 1)

        if truncated:
            # Recalcula preview reservando espaço do botão para não “empurrar” o texto.
            content_width_btn = max(40, col_width - 70)
            clamped_btn, _ = _clamp_observation_two_lines(full_text, metrics, content_width_btn)
            preview.setPlainText(clamped_btn)

            btn = QPushButton("LER MAIS")
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet("""
                QPushButton {
                    background-color: #007BFF;
                    color: white;
                    border: none;
                    border-radius: 6px;
                    padding: 6px 10px;
                    font-size: 11px;
                    font-weight: bold;
                }
                QPushButton:hover { background-color: #0056b3; }
                QPushButton:pressed { background-color: #004085; }
            """)
            btn.clicked.connect(
                lambda _checked=False, num=ddu_number, nat=natureza, obs=full_text: ReadMoreObservationDialog(
                    num, nat, obs, parent=self
                ).exec()
            )
            container_layout.addWidget(btn, 0, Qt.AlignRight | Qt.AlignVCenter)

        table.setCellWidget(row, col, container)

        # Garante altura para 2 linhas sem estourar a célula
        needed_h = (metrics.lineSpacing() * 2) + 16
        current_h = table.rowHeight(row)
        if current_h < needed_h:
            table.setRowHeight(row, needed_h)
        preview.setFixedHeight(needed_h - 2)
    
    def filter_table(self):
        """Filtra a tabela baseado no texto da pesquisa (Número, Natureza, Local, Observação)"""
        search_text = self.search_input.text().lower()
        
        for row in range(self.table.rowCount()):
            # Se o campo de pesquisa estiver vazio, mostra todas as linhas
            if not search_text:
                self.table.setRowHidden(row, False)
                continue
            
            # Verifica nas colunas: Número (2), Natureza (3), Local (4), Observação (5)
            match_found = False
            search_columns = [2, 3, 4, 5]  # Número, Natureza, Local, Observação
            for col in search_columns:
                item = self.table.item(row, col)
                if item and search_text in item.text().lower():
                    match_found = True
                    break
            
            # Mostra ou oculta a linha baseado na busca
            self.table.setRowHidden(row, not match_found)
    
    def show_ddu_info(self, doc_id):
        """Exibe as informações de resultado (REDS ou BOS)"""
        # Primeiro tenta buscar dados de resultado do documento
        result_data = get_document_result_data(doc_id)
        if not result_data or result_data[0] == 0:
            QMessageBox.warning(self, "Atenção", "Este documento não possui resultado marcado.")
            return
        
        result_pos, numero_reds, data_reds, militar_reds, resultado_reds = result_data
        
        # Buscar informações do documento (tipo e número)
        doc_info = get_document_info(doc_id)
        if not doc_info:
            QMessageBox.warning(self, "Atenção", "Não foi possível obter informações do documento.")
            return
        
        tipo, numero = doc_info
        
        # Criar tupla compatível com DduPositivadaInfoDialog
        data = (tipo, numero, numero_reds, data_reds, militar_reds, resultado_reds)
        dlg = DduPositivadaInfoDialog(data, result_pos, self)
        dlg.exec()

    def handle_info_button_click(self, doc_id):
        if not (QApplication.keyboardModifiers() & Qt.ControlModifier):
            self.table.clearSelection()
            self.table.setCurrentItem(None)

        self.show_ddu_info(doc_id)

        if not (QApplication.keyboardModifiers() & Qt.ControlModifier):
            QTimer.singleShot(0, lambda: (self.table.clearSelection(), self.table.setCurrentItem(None)))

    def get_doc_id_from_row(self, row):
        item = self.table.item(row, 1)
        if not item:
            return None
        return item.data(Qt.UserRole)

    def get_result_checkbox_from_row(self, row):
        container = self.table.cellWidget(row, 10)
        if not container:
            return None
        return container.findChild(QCheckBox)

    def handle_locked_result_pos_double_click(self, row, column):
        if column != 10:
            return

        checkbox = self.get_result_checkbox_from_row(row)
        if not checkbox or checkbox.isEnabled() or not checkbox.isChecked():
            return

        doc_id = self.get_doc_id_from_row(row)
        if doc_id is None:
            return

        self.handle_locked_result_pos_edit(doc_id)

    def eventFilter(self, obj, event):
        try:
            doc_id = obj.property("locked_result_doc_id")
            if doc_id is not None:
                if event.type() in (QEvent.MouseButtonPress, QEvent.MouseButtonRelease):
                    return True
                if event.type() == QEvent.MouseButtonDblClick:
                    self.schedule_locked_result_pos_edit(int(doc_id))
                    return True
            return super().eventFilter(obj, event)
        except KeyboardInterrupt:
            return True
        except Exception:
            return super().eventFilter(obj, event)

    def schedule_locked_result_pos_edit(self, doc_id):
        if self._locked_result_edit_in_progress:
            return

        self._locked_result_edit_in_progress = True

        def _run_edit():
            try:
                self.handle_locked_result_pos_edit(doc_id)
            finally:
                self._locked_result_edit_in_progress = False

        QTimer.singleShot(0, _run_edit)

    def handle_locked_result_pos_edit(self, doc_id):
        self.table.clearSelection()
        self.table.setCurrentItem(None)

        QMessageBox.information(
            self,
            "Coluna Travada",
            "A coluna Resultado está travada.\nInforme a senha para continuar a edição."
        )

        password, ok = QInputDialog.getText(
            self,
            "Desbloquear Resultado",
            "Senha:",
            QLineEdit.Password
        )
        if not ok:
            return

        if password != RESULT_POS_UNLOCK_PASSWORD:
            QMessageBox.warning(self, "Atenção", "Senha incorreta.")
            return

        # Buscar dados de resultado diretamente da tabela documents
        existing_data = get_document_result_data(doc_id)
        if not existing_data:
            QMessageBox.warning(self, "Atenção", "Não foram encontrados dados atuais para edição.")
            return

        result_pos, numero_reds, data_reds, militar_reds, resultado_reds = existing_data
        dlg = RedsDialog(
            self,
            initial_data={
                "result_type": result_pos,
                "reds_number": numero_reds or "",
                "reds_date": data_reds or "",
                "militar_reds": militar_reds or "",
                "resultado_reds": resultado_reds or "",
            }
        )
        if dlg.exec() != QDialog.Accepted:
            return

        data = dlg.get_data()
        if data is None:
            return
        result_type = data.get("result_type", 1)
        
        doc_info = get_document_info(doc_id)
        if not doc_info:
            QMessageBox.warning(self, "Atenção", "Não foi possível obter os dados do documento.")
            return

        tipo, numero = doc_info
        update_result_pos(
            doc_id,
            result_type,
            data["reds_number"],
            data["reds_date"],
            data["militar_reds"],
            data["resultado_reds"]
        )
        
        # Remove da tabela DDU POSITIVADA (caso exista)
        remove_ddu_positivada(doc_id)
        
        # Só adiciona na tabela DDU POSITIVADA se for resultado POSITIVO
        if result_type == 1:
            add_ddu_positivada(
                doc_id,
                tipo,
                numero,
                data["reds_number"],
                data["reds_date"],
                data["militar_reds"],
                data["resultado_reds"]
            )

        QMessageBox.information(self, "Sucesso", "Dados de Resultado atualizados com sucesso.")
        self.refresh_list()
        self.table.clearSelection()
        self.table.setCurrentItem(None)

    def confirm_sim_nao(self, title, message):
        return confirm_sim_nao(self, title, message)

    def _set_checkbox_state_safe(self, checkbox, checked):
        if checkbox is None:
            return
        try:
            checkbox.blockSignals(True)
            checkbox.setChecked(checked)
        except RuntimeError:
            return
        finally:
            try:
                checkbox.blockSignals(False)
            except RuntimeError:
                pass
    
    def handle_result_pos_change(self, doc_id, state, checkbox):
        """Gerencia a mudança do checkbox Result Pos com confirmação"""
        new_value = 1 if state == Qt.CheckState.Checked.value else 0
        old_value = 0 if new_value == 1 else 1
        
        if new_value == 1:
            # Ao marcar, abre o diálogo para coletar dados
            dlg = RedsDialog(self)
            if dlg.exec() == QDialog.Accepted:
                data = dlg.get_data()
                
                # Verificação de segurança: se não há dados válidos, reverte
                if not data:
                    self._set_checkbox_state_safe(checkbox, False)
                    return
                
                result_type = data.get("result_type", 1)  # 1 = positivo, -1 = negativo
                
                # Obter tipo e número do documento
                doc_info = get_document_info(doc_id)
                if doc_info:
                    tipo, numero = doc_info
                    
                    # Atualiza com o result_type apropriado
                    update_result_pos(doc_id, result_type, data["reds_number"], data["reds_date"], 
                                     data["militar_reds"], data["resultado_reds"])
                    
                    # Só adiciona na tabela DDU POSITIVADA se for resultado POSITIVO
                    if result_type == 1:
                        add_ddu_positivada(
                            doc_id,
                            tipo,
                            numero,
                            data["reds_number"],
                            data["reds_date"],
                            data["militar_reds"],
                            data["resultado_reds"]
                        )
                    
                    # Atualizar a lista imediatamente
                    self.refresh_list()
                else:
                    # Se não conseguiu obter info do documento, reverte o checkbox
                    self._set_checkbox_state_safe(checkbox, False)
            else:
                # Usuário cancelou/fechou o diálogo sem confirmar - reverte o checkbox
                # NENHUMA ação é registrada no banco de dados
                self._set_checkbox_state_safe(checkbox, False)
        else:
            # Ao desmarcar, pede confirmação
            confirmado = self.confirm_sim_nao(
                "Confirmar",
                "Deseja desmarcar este documento?"
            )
            
            if confirmado:
                # Confirma a desmarcação
                update_result_pos(doc_id, 0, None, None, None, None)
                # Remove da tabela DDU POSITIVADA (caso tenha sido positivo)
                remove_ddu_positivada(doc_id)
                
                # Atualizar a lista imediatamente
                self.refresh_list()
            else:
                # Reverte a mudança
                self._set_checkbox_state_safe(checkbox, True)

    def handle_list(self, _checked=False):
        dlg = ListDocumentsDialog(self)
        dlg.exec()

    def handle_update_site(self, _checked=False):
        # Determinar se documentos vencidos serão incluídos
        export_expired = self.chk_export_expired.isChecked() if hasattr(self, "chk_export_expired") else False

        if export_expired:
            descricao_exportacao = "Serão enviados todos os documentos (válidos e vencidos)."
        else:
            descricao_exportacao = "Serão enviados somente os documentos com validade em vigor.\nDocumentos vencidos NÃO serão incluídos."

        confirmado = confirm_sim_nao(
            self,
            "Confirmar Exportação",
            f"Deseja gerar e enviar o arquivo data.json para o servidor?\n\n{descricao_exportacao}"
        )

        if not confirmado:
            return
        
        temp_dir = None
        try:
            # Criar pasta temporária
            temp_dir = tempfile.mkdtemp()
            file_path = os.path.join(temp_dir, "data.json")
            
            total, valid_count, expired_count = export_data_json(path=file_path, export_expired=export_expired)

            # Impedir envio de arquivo vazio
            if total == 0:
                QMessageBox.warning(
                    self,
                    "Nenhum documento",
                    "Não há documentos para exportar com os critérios selecionados."
                )
                return
            
            # Ler arquivo e criptografar
            with open(file_path, "rb") as f:
                plaintext = f.read()
            
            # Configurações de criptografia (mesmas do servidor)
            ENCRYPTION_KEY = bytes.fromhex('000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f')
            IV = bytes.fromhex('000102030405060708090a0b0c0d0e0f')
            
            # Criptografar com AES-256-CBC
            cipher = AES.new(ENCRYPTION_KEY, AES.MODE_CBC, IV)
            # Adicionar padding PKCS7
            padded_plaintext = pad(plaintext, AES.block_size)
            encrypted_data = cipher.encrypt(padded_plaintext)
            
            # Enviar arquivo criptografado para servidor
            headers = {
                "x-auth-token": "DDU_CONTROLE_SECRET_TOKEN_12345"
            }
            upload_candidates = build_upload_url_candidates()
            response = None
            used_upload_url = upload_candidates[0] if upload_candidates else get_data_upload_url()

            for candidate_url in upload_candidates:
                used_upload_url = candidate_url
                response = requests.post(
                    candidate_url,
                    data=encrypted_data,
                    headers=headers,
                    timeout=10
                )

                # Se sucesso, encerra. Se 404, tenta próximo candidato.
                if response.status_code in [200, 201] or response.status_code != 404:
                    break
            
            # Verificar se o envio foi bem-sucedido
            if response is not None and response.status_code in [200, 201]:
                if export_expired:
                    detalhe_envio = f"{total} documento(s) enviado(s) ({valid_count} válido(s) e {expired_count} vencido(s))."
                else:
                    detalhe_envio = f"{total} documento(s) válido(s) enviado(s). ({expired_count} vencido(s) não incluído(s))"
                QMessageBox.information(
                    self,
                    "Sucesso",
                    f"Arquivo data.json foi enviado com sucesso!\n\n{detalhe_envio}"
                )
            else:
                response_status = response.status_code if response is not None else "sem resposta"
                response_detail = ""
                if response is not None:
                    detail_text = (response.text or "").strip()
                    if detail_text:
                        response_detail = f"\n\nResposta do servidor: {detail_text[:300]}"
                QMessageBox.critical(
                    self,
                    "Erro",
                    f"Falha ao enviar arquivo. Status: {response_status}\n"
                    f"URL utilizada: {used_upload_url}{response_detail}"
                )
        
        except requests.exceptions.ConnectionError:
            upload_url = get_data_upload_url()
            parsed = urlsplit(upload_url)
            server_base_url = f"{parsed.scheme}://{parsed.netloc}" if parsed.scheme and parsed.netloc else upload_url
            QMessageBox.critical(
                self,
                "Erro de Conexão",
                f"Não foi possível conectar ao servidor em {server_base_url}.\n\n"
                "Verifique se o DDU_Web está rodando."
            )
        except requests.exceptions.Timeout:
            QMessageBox.critical(
                self,
                "Erro de Timeout",
                "O servidor demorou muito para responder. Tente novamente."
            )
        except requests.exceptions.RequestException as e:
            QMessageBox.critical(
                self,
                "Erro na Requisição",
                f"Erro ao enviar arquivo: {str(e)}"
            )
        except Exception as e:
            QMessageBox.critical(
                self,
                "Erro",
                f"Falha ao processar exportação: {str(e)}"
            )
        finally:
            # Limpar pasta temporária
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)

    def handle_pdf(self):
        # Verificar se há documentos cadastrados
        docs = list_documents()
        if not docs:
            QMessageBox.warning(self, "Atenção", "Não há documentos cadastrados. Impossível gerar PDF sem dados.")
            return
        
        confirmado = confirm_sim_nao(
            self,
            "Confirmar",
            "Deseja gerar o relatório PDF de documentos?"
        )
        if not confirmado:
            return
        
        # Perguntar ao usuário onde salvar o arquivo
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar Relatório PDF",
            "relatorio_documentos.pdf",
            "PDF Files (*.pdf)"
        )
        
        if not file_path:
            return  # Usuário cancelou
        
        try:
            # Verificar o estado do checkbox "Visualizar Documentos Vencidos" da MainWindow
            include_expired = self.chk_show_expired.isChecked()
            generate_pdf_report(file_path, include_expired=include_expired)
            QMessageBox.information(self, "Relatório", f"Relatório gerado com sucesso em:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao gerar PDF: {e}")

    def handle_pdf_positivo(self):
        # Verificar se há documentos cadastrados
        docs = list_documents()
        if not docs:
            QMessageBox.warning(self, "Atenção", "Não há documentos cadastrados. Impossível gerar PDF sem dados.")
            return
        
        # Verificar se há pelo menos um documento com Result Pos marcado
        # O result_pos é o último campo (índice 11) de cada documento
        has_positivo = any(doc[11] == 1 for doc in docs)
        if not has_positivo:
            QMessageBox.warning(self, "Atenção", "Não há documentos com Resultado Positivo marcado. Impossível gerar PDF sem dados.")
            return
        
        confirmado = confirm_sim_nao(
            self,
            "Confirmar",
            "Deseja gerar o relatório PDF de DDU Positivada?"
        )
        if not confirmado:
            return
        
        # Perguntar ao usuário onde salvar o arquivo
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar Relatório DDU Positivada",
            "relatorio_ddu_positivada.pdf",
            "PDF Files (*.pdf)"
        )
        
        if not file_path:
            return  # Usuário cancelou
        
        try:
            success = generate_pdf_positivo(file_path)
            if success:
                QMessageBox.information(self, "Relatório", f"Relatório de DDU Positivada gerado com sucesso em:\n{file_path}")
            else:
                QMessageBox.warning(self, "Aviso", "Não há registros de DDU Positivada para gerar o relatório.")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao gerar PDF de DDU Positivada: {e}")

    def handle_delete(self):
        # Obter linhas selecionadas (sem duplicatas)
        selected_rows = set()
        for item in self.table.selectedItems():
            selected_rows.add(item.row())
        
        if not selected_rows:
            QMessageBox.warning(self, "Atenção", "Selecione um ou mais documentos na lista para excluir")
            return
        
        # Coletar IDs dos documentos selecionados
        doc_ids = []
        for row in selected_rows:
            item = self.table.item(row, 1)  # Coluna 1 (Tipo) contém o ID armazenado
            if item:
                doc_id = item.data(Qt.UserRole)
                if doc_id is not None:
                    doc_ids.append(doc_id)
        
        if not doc_ids:
            QMessageBox.warning(self, "Atenção", "Não foi possível determinar os documentos selecionados")
            return
        
        # Confirmar exclusão
        count = len(doc_ids)
        if count == 1:
            message = "Confirma exclusão do documento selecionado?"
        else:
            message = f"Confirma exclusão dos {count} documentos selecionados?"
        
        confirmado = confirm_sim_nao(self, "Confirmar", message)
        if confirmado:
            if count == 1:
                delete_document(doc_ids[0])
            else:
                delete_documents(doc_ids)
            
            success_msg = "Documento excluído com sucesso" if count == 1 else f"{count} documentos excluídos com sucesso"
            QMessageBox.information(self, "Excluído", success_msg)
            self.refresh_list()

    def check_alerts(self):
        docs = list_documents()
        today = datetime.date.today()
        
        # Coletar documentos que vencerão em 5 dias
        docs_five_days = []
        # Coletar documentos que vencerão hoje
        docs_today = []
        
        for d in docs:
            eid, dtype, dnum, nature, location, observation, rdate, edate, a5, ad, military, result_pos = d
            # aceitar tanto ISO quanto BR
            ed = None
            for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
                try:
                    ed = datetime.datetime.strptime(edate, fmt).date()
                    break
                except Exception:
                    continue
            if ed is None:
                continue
            
            days_left = (ed - today).days
            
            # Coletar documentos que vencerão em 5 dias (não ignorar resultado_pos)
            if days_left == 5 and result_pos != 1 and result_pos != -1:
                docs_five_days.append((eid, dtype, dnum, edate))
            
            # Coletar documentos que vencerão hoje (não ignorar resultado_pos)
            if days_left == 0 and result_pos != 1 and result_pos != -1:
                docs_today.append((eid, dtype, dnum, edate))
        
        # Mostrar alerta consolidado para 5 dias
        if docs_five_days and not check_alert_shown_today("alert_5_days"):
            numeros = ", ".join([dnum for _, _, dnum, _ in docs_five_days])
            msg = f"Atenção!\n\n{len(docs_five_days)} documento(s) vencerá/ão em 05 dias\n\nNúmeros: {numeros}"
            QMessageBox.warning(self, "Alerta - 5 Dias para Vencer", msg)
            mark_alert_shown("alert_5_days")
            # Marcar todos os documentos como alertados
            for eid, _, _, _ in docs_five_days:
                mark_alert(eid, "5")
        
        # Mostrar alerta consolidado para hoje
        if docs_today and not check_alert_shown_today("alert_due_today"):
            numeros = ", ".join([dnum for _, _, dnum, _ in docs_today])
            msg = f"Atenção!\n\n{len(docs_today)} documento(s) vence/m HOJE\n\nNúmeros: {numeros}"
            QMessageBox.warning(self, "Alerta - Vencimento Hoje", msg)
            mark_alert_shown("alert_due_today")
            # Marcar todos os documentos como alertados
            for eid, _, _, _ in docs_today:
                mark_alert(eid, "due")
        
        self.refresh_list()

    def handle_dashboard(self):
        """Abre o dashboard com gráficos e análises"""
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            dlg = DashboardDialog(self)
            QApplication.processEvents()
            QApplication.restoreOverrideCursor()
            dlg.showMaximized()  # Maximiza a janela antes de exibir
            dlg.exec()
        finally:
            QApplication.restoreOverrideCursor()

class ComparePeriodDialog(QDialog):
    """Diálogo para configurar comparação de dois períodos"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configurar Comparação de Períodos")
        self.resize(500, 300)
        self.setModal(True)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QLabel {
                color: #003366;
                font-weight: bold;
            }
            QLineEdit {
                padding: 8px;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                background-color: white;
            }
            QPushButton {
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
        """)
        
        layout = QVBoxLayout(self)
        
        # Período 1
        lbl_period1 = QLabel("📅 Período 1:")
        layout.addWidget(lbl_period1)
        
        period1_layout = QHBoxLayout()
        lbl_start1 = QLabel("Data Inicial:")
        self.date_start1 = QLineEdit()
        self.date_start1.setPlaceholderText("DD/MM/AAAA")
        self.date_start1.setInputMask("00/00/0000")
        btn_cal1_start = QPushButton("📆")
        btn_cal1_start.setMaximumWidth(50)
        btn_cal1_start.clicked.connect(lambda: self._open_calendar(self.date_start1))
        period1_layout.addWidget(lbl_start1)
        period1_layout.addWidget(self.date_start1)
        period1_layout.addWidget(btn_cal1_start)
        layout.addLayout(period1_layout)
        
        period1_layout2 = QHBoxLayout()
        lbl_end1 = QLabel("Data Final:")
        self.date_end1 = QLineEdit()
        self.date_end1.setPlaceholderText("DD/MM/AAAA")
        self.date_end1.setInputMask("00/00/0000")
        btn_cal1_end = QPushButton("📆")
        btn_cal1_end.setMaximumWidth(50)
        btn_cal1_end.clicked.connect(lambda: self._open_calendar(self.date_end1))
        period1_layout2.addWidget(lbl_end1)
        period1_layout2.addWidget(self.date_end1)
        period1_layout2.addWidget(btn_cal1_end)
        layout.addLayout(period1_layout2)
        
        layout.addSpacing(15)
        
        # Período 2
        lbl_period2 = QLabel("📅 Período 2:")
        layout.addWidget(lbl_period2)
        
        period2_layout = QHBoxLayout()
        lbl_start2 = QLabel("Data Inicial:")
        self.date_start2 = QLineEdit()
        self.date_start2.setPlaceholderText("DD/MM/AAAA")
        self.date_start2.setInputMask("00/00/0000")
        btn_cal2_start = QPushButton("📆")
        btn_cal2_start.setMaximumWidth(50)
        btn_cal2_start.clicked.connect(lambda: self._open_calendar(self.date_start2))
        period2_layout.addWidget(lbl_start2)
        period2_layout.addWidget(self.date_start2)
        period2_layout.addWidget(btn_cal2_start)
        layout.addLayout(period2_layout)
        
        period2_layout2 = QHBoxLayout()
        lbl_end2 = QLabel("Data Final:")
        self.date_end2 = QLineEdit()
        self.date_end2.setPlaceholderText("DD/MM/AAAA")
        self.date_end2.setInputMask("00/00/0000")
        btn_cal2_end = QPushButton("📆")
        btn_cal2_end.setMaximumWidth(50)
        btn_cal2_end.clicked.connect(lambda: self._open_calendar(self.date_end2))
        period2_layout2.addWidget(lbl_end2)
        period2_layout2.addWidget(self.date_end2)
        period2_layout2.addWidget(btn_cal2_end)
        layout.addLayout(period2_layout2)
        
        layout.addSpacing(20)
        
        # Botões
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        btn_cancel = QPushButton("Cancelar")
        btn_cancel.setStyleSheet("background-color: #6c757d; color: white;")
        btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(btn_cancel)
        
        btn_apply = QPushButton("Aplicar")
        btn_apply.setStyleSheet("background-color: #007BFF; color: white;")
        btn_apply.clicked.connect(self._validate_and_accept)
        btn_layout.addWidget(btn_apply)
        
        layout.addLayout(btn_layout)
    
    def _open_calendar(self, target_input):
        """Abre calendário para seleção de data"""
        # Tentar carregar data atual se houver
        current_text = target_input.text().strip()
        current_date = None
        if current_text:
            for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
                try:
                    current_date = datetime.datetime.strptime(current_text, fmt).date()
                    break
                except:
                    continue

        selected = show_date_picker_dialog(self, initial_date=current_date)
        if selected:
            target_input.setText(selected.strftime("%d/%m/%Y"))
    
    def _validate_and_accept(self):
        """Valida se os períodos não se sobrepõem e aceita"""
        try:
            # Parse dates
            dates = []
            for txt in [self.date_start1.text(), self.date_end1.text(), 
                       self.date_start2.text(), self.date_end2.text()]:
                if not txt.strip():
                    QMessageBox.warning(self, "Erro", "Todos os campos de data são obrigatórios!")
                    return
                
                parsed = None
                for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
                    try:
                        parsed = datetime.datetime.strptime(txt.strip(), fmt).date()
                        break
                    except:
                        continue
                
                if not parsed:
                    QMessageBox.warning(self, "Erro", f"Data inválida: {txt}")
                    return
                dates.append(parsed)
            
            start1, end1, start2, end2 = dates
            
            # Validar que data inicial <= data final em cada período
            if start1 > end1:
                QMessageBox.warning(self, "Erro", "No Período 1: data inicial não pode ser posterior à final!")
                return
            
            if start2 > end2:
                QMessageBox.warning(self, "Erro", "No Período 2: data inicial não pode ser posterior à final!")
                return
            
            # Validar sobreposição de períodos
            # Período 1: [start1, end1]
            # Período 2: [start2, end2]
            # Não se sobrepõem se: end1 < start2 OR end2 < start1
            
            if not (end1 < start2 or end2 < start1):
                QMessageBox.warning(self, "Erro", "Os períodos não podem se sobrepor!\n\n"
                                   f"Período 1: {start1.strftime('%d/%m/%Y')} a {end1.strftime('%d/%m/%Y')}\n"
                                   f"Período 2: {start2.strftime('%d/%m/%Y')} a {end2.strftime('%d/%m/%Y')}")
                return
            
            self.accept()
        
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao validar: {str(e)}")
    
    def get_periods(self):
        """Retorna os dois períodos como tuplas de datas"""
        dates = []
        for txt in [self.date_start1.text(), self.date_end1.text(), 
                   self.date_start2.text(), self.date_end2.text()]:
            for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
                try:
                    parsed = datetime.datetime.strptime(txt.strip(), fmt).date()
                    dates.append(parsed)
                    break
                except:
                    continue
        
        return (dates[0], dates[1]), (dates[2], dates[3])


def plotly_to_pixmap(fig, width=800, height=600):
    """Converte um gráfico Plotly em QPixmap para exibição no PySide6"""
    try:
        # Gerar imagem PNG a partir da figura Plotly
        img_bytes = fig.to_image(format="png", width=width, height=height)
        pixmap = QPixmap()
        pixmap.loadFromData(img_bytes, 'PNG')
        return pixmap
    except Exception as e:
        print(f"Erro ao converter gráfico Plotly: {e}")
        return None


class DashboardDialog(QDialog):
    """Diálogo principal do Dashboard com gráficos e análises"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📊 Dashboard de Análise")
        self.setModal(True)
        # Janela será maximizada após a construção
        
        self.period_start = None
        self.period_end = None
        self.compare_mode = False
        self.period1 = None
        self.period2 = None
        
        self.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QLabel {
                color: #003366;
            }
        """)
        
        self.main_layout = QVBoxLayout(self)
        
        # Header com filtros
        self._create_header()
        
        # Área de conteúdo (será preenchida com gráficos)
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout(self.content_widget)
        self.main_layout.addWidget(self.content_widget)
        
        # Botão Fechar
        btn_close = QPushButton("Fechar")
        btn_close.clicked.connect(self.accept)
        self.main_layout.addWidget(btn_close)
        
        # Carregar dados inicialmente
        self.refresh_dashboard()
    
    def _create_header(self):
        """Cria cabeçalho com filtros"""
        header = QWidget()
        header.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #003366, stop:1 #004080);
                border-radius: 8px;
                padding: 15px;
            }
        """)
        header_layout = QVBoxLayout(header)
        
        # Título
        title = QLabel("📊 Dashboard de Documentos")
        title.setFont(QFont("Arial", 14, QFont.Bold))
        title.setStyleSheet("color: white;")
        header_layout.addWidget(title)
        
        # Filtros
        filter_layout = QHBoxLayout()
        
        # Data Inicial
        lbl_start = QLabel("Data Inicial:")
        lbl_start.setStyleSheet("color: white; font-weight: bold;")
        self.date_start_input = QLineEdit()
        self.date_start_input.setPlaceholderText("DD/MM/AAAA")
        self.date_start_input.setInputMask("00/00/0000")
        self.date_start_input.setMaximumWidth(150)
        self.date_start_input.setStyleSheet("color: white; background-color: rgba(255, 255, 255, 0.2); border: 1px solid white; border-radius: 4px; padding: 5px;")
        btn_start = QPushButton("📆")
        btn_start.setMaximumWidth(50)
        btn_start.clicked.connect(lambda: self._open_calendar(self.date_start_input))
        
        # Data Final
        lbl_end = QLabel("Data Final:")
        lbl_end.setStyleSheet("color: white; font-weight: bold;")
        self.date_end_input = QLineEdit()
        self.date_end_input.setPlaceholderText("DD/MM/AAAA")
        self.date_end_input.setInputMask("00/00/0000")
        self.date_end_input.setMaximumWidth(150)
        self.date_end_input.setStyleSheet("color: white; background-color: rgba(255, 255, 255, 0.2); border: 1px solid white; border-radius: 4px; padding: 5px;")
        btn_end = QPushButton("📆")
        btn_end.setMaximumWidth(50)
        btn_end.clicked.connect(lambda: self._open_calendar(self.date_end_input))
        
        # Botão Comparar Períodos
        btn_compare = QPushButton("⚙️ Configurar Comparação")
        btn_compare.setStyleSheet("""
            QPushButton {
                background-color: #28a745;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #218838;
            }
        """)
        btn_compare.clicked.connect(self._open_compare_dialog)
        
        # Botão Aplicar Filtro
        btn_apply_filter = QPushButton("🔍 Aplicar Filtro")
        btn_apply_filter.setStyleSheet("""
            QPushButton {
                background-color: #FFC107;
                color: #003366;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #FFD54F;
            }
        """)
        btn_apply_filter.clicked.connect(self.refresh_dashboard)
        
        filter_layout.addWidget(lbl_start)
        filter_layout.addWidget(self.date_start_input)
        filter_layout.addWidget(btn_start)
        filter_layout.addSpacing(20)
        filter_layout.addWidget(lbl_end)
        filter_layout.addWidget(self.date_end_input)
        filter_layout.addWidget(btn_end)
        filter_layout.addSpacing(20)
        filter_layout.addWidget(btn_apply_filter)
        filter_layout.addStretch()
        filter_layout.addWidget(btn_compare)
        filter_layout.addSpacing(40)
        
        # Botão Gerar PDF
        btn_generate_pdf = QPushButton("📄 Gerar PDF")
        btn_generate_pdf.setStyleSheet("""
            QPushButton {
                background-color: #007BFF;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
        """)
        btn_generate_pdf.clicked.connect(self._generate_dashboard_pdf)
        filter_layout.addWidget(btn_generate_pdf)
        
        header_layout.addLayout(filter_layout)
        
        # Define altura fixa para o header evitar expansão
        header.setMaximumHeight(150)
        
        self.main_layout.addWidget(header)
    
    def _open_calendar(self, target_input):
        """Abre calendário para seleção de data"""
        current_text = target_input.text().strip()
        current_date = self._parse_date(current_text) if current_text else None

        selected = show_date_picker_dialog(self, initial_date=current_date)
        if selected:
            target_input.setText(selected.strftime("%d/%m/%Y"))
    
    def _open_compare_dialog(self):
        """Abre diálogo para configurar comparação de períodos"""
        dlg = ComparePeriodDialog(self)
        if dlg.exec() == QDialog.Accepted:
            self.period1, self.period2 = dlg.get_periods()
            self.compare_mode = True
            self.refresh_dashboard()
    
    def _generate_dashboard_pdf(self):
        """Gera PDF do dashboard com dados filtrados"""
        all_docs = list_documents()
        
        if not all_docs:
            QMessageBox.warning(self, "Atenção", "Não há documentos cadastrados. Impossível gerar PDF sem dados.")
            return
        
        # Confirmar geração (SIM/NÃO)
        if not confirm_sim_nao(
            self,
            "Confirmar Geração de PDF",
            "Deseja realmente gerar o arquivo PDF do dashboard?"
        ):
            return
        
        # Abrir diálogo para salvar arquivo
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Salvar PDF do Dashboard",
            os.path.join(os.path.expanduser("~"), "Desktop", "dashboard_relatorio.pdf"),
            "PDF Files (*.pdf)"
        )
        
        if not file_path:
            return
        
        # Garantir extensão .pdf
        if not file_path.endswith(".pdf"):
            file_path += ".pdf"
        
        try:
            if self.compare_mode and self.period1 and self.period2:
                start1, end1 = self.period1
                start2, end2 = self.period2
                docs1 = self._filter_documents(all_docs, start1, end1)
                docs2 = self._filter_documents(all_docs, start2, end2)
                self._create_dashboard_pdf(
                    file_path,
                    docs1,
                    start1,
                    end1,
                    docs2=docs2,
                    start_date_2=start2,
                    end_date_2=end2,
                )
            else:
                start_date = self._parse_date(self.date_start_input.text())
                end_date = self._parse_date(self.date_end_input.text())
                docs = self._filter_documents(all_docs, start_date, end_date)
                self._create_dashboard_pdf(file_path, docs, start_date, end_date)
            
            QMessageBox.information(self, "Sucesso", f"PDF do dashboard gerado com sucesso!\n\nSalvo em:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Falha ao gerar PDF: {e}")
    
    def _create_dashboard_pdf(self, path, docs, start_date, end_date, docs2=None, start_date_2=None, end_date_2=None):
        """Cria o arquivo PDF do dashboard com dados e gráficos"""
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib import colors
        from reportlab.pdfgen import canvas
        from reportlab.lib.utils import ImageReader
        from reportlab.platypus import Table, TableStyle
        
        c = canvas.Canvas(path, pagesize=landscape(A4))
        width, height = landscape(A4)
        margin = 50

        def draw_plotly_on_canvas(fig, x, y, w, h):
            try:
                img_bytes = fig.to_image(format="png", width=int(w * 2), height=int(h * 2))
                img_reader = ImageReader(io.BytesIO(img_bytes))
                c.drawImage(img_reader, x, y, width=w, height=h, preserveAspectRatio=True, anchor='c')
            except Exception as err:
                c.setFont("Helvetica", 10)
                c.drawString(x, y + (h / 2), f"Não foi possível renderizar gráfico: {err}")

        def draw_pdf_header(subtitle):
            y_pos = criar_cabecalho_pmmg(c, width, height, margin)
            y_pos -= 30
            c.setFont("Helvetica-Bold", 18)
            c.drawString(margin, y_pos, "Dashboard - Relatório de Documentos")
            y_pos -= 24

            c.setFont("Helvetica-Bold", 12)
            c.drawString(margin, y_pos, subtitle)
            y_pos -= 20

            c.setFont("Helvetica", 11)
            periodo_text = "Período: "
            if start_date and end_date:
                periodo_text += f"{start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')}"
            elif start_date:
                periodo_text += f"A partir de {start_date.strftime('%d/%m/%Y')}"
            elif end_date:
                periodo_text += f"Até {end_date.strftime('%d/%m/%Y')}"
            else:
                periodo_text += "Todos os documentos"
            c.drawString(margin, y_pos, periodo_text)
            y_pos -= 18

            c.drawString(margin, y_pos, f"Gerado em: {datetime.date.today().strftime('%d/%m/%Y')}")
            y_pos -= 24
            return y_pos

        def draw_section_divider(top_y, label="SECAO DOS GRAFICOS"):
            """Desenha divisoria visual entre a area de dados e de graficos."""
            divider_line_y = top_y - 10
            c.setStrokeColor(colors.HexColor("#9aa0a6"))
            c.setLineWidth(1)
            c.line(margin, divider_line_y, width - margin, divider_line_y)

            c.setFillColor(colors.HexColor("#495057"))
            c.setFont("Helvetica-Bold", 9)
            c.drawString(margin, divider_line_y + 3, label)
            c.setFillColor(colors.black)

            # Reserva espaco abaixo da divisoria antes do grafico.
            return divider_line_y - 12

        if docs2 is not None and start_date_2 and end_date_2:
            y_position = draw_pdf_header("Dados Numéricos (base dos gráficos) - Comparação")

            total1 = len(docs)
            positive1 = sum(1 for d in docs if d[11] == 1)
            total2 = len(docs2)
            positive2 = sum(1 for d in docs2 if d[11] == 1)

            compare_data = [
                ["Período", "Total", "Positivas", "Negativas", "% Positivas"],
                [
                    f"1 ({start_date.strftime('%d/%m/%Y')} a {end_date.strftime('%d/%m/%Y')})",
                    str(total1),
                    str(positive1),
                    str(total1 - positive1),
                    f"{(positive1 * 100 / total1):.1f}%" if total1 > 0 else "0.0%",
                ],
                [
                    f"2 ({start_date_2.strftime('%d/%m/%Y')} a {end_date_2.strftime('%d/%m/%Y')})",
                    str(total2),
                    str(positive2),
                    str(total2 - positive2),
                    f"{(positive2 * 100 / total2):.1f}%" if total2 > 0 else "0.0%",
                ],
            ]

            table_compare = Table(compare_data, colWidths=[320, 80, 90, 90, 90])
            table_compare.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003366")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (1, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("GRID", (0, 0), (-1, -1), 1, colors.grey),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#f8f9fa")),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("WORDWRAP", (0, 0), (-1, -1), "CJK"),
            ]))

            tw, th = table_compare.wrapOn(c, width - (2 * margin), y_position)
            table_compare.drawOn(c, margin, y_position - th)
            chart_top = y_position - th
            chart_y = margin
            chart_top = draw_section_divider(chart_top, "SECAO DOS GRAFICOS COMPARATIVOS")
            chart_h = chart_top - chart_y

            # Se não houver altura suficiente para o gráfico sem sobreposição,
            # usa nova página para separar claramente tabela e gráfico.
            if chart_h < 220:
                c.showPage()
                y_position_graph = draw_pdf_header("Gráfico Comparativo")
                chart_top = y_position_graph
                chart_y = margin
                chart_top = draw_section_divider(chart_top, "SECAO DOS GRAFICOS COMPARATIVOS")
                chart_h = chart_top - chart_y

            fig_compare = go.Figure()
            if total1 > 0:
                fig_compare.add_trace(go.Pie(
                    labels=["Positivas", "Negativas"],
                    values=[positive1, total1 - positive1],
                    marker=dict(colors=["#28a745", "#dc3545"]),
                    textinfo='label+percent',
                    hoverinfo='label+value+percent',
                    textfont=dict(size=13, color='white'),
                    domain=dict(x=[0, 0.5]),
                    name="Período 1",
                ))
            if total2 > 0:
                fig_compare.add_trace(go.Pie(
                    labels=["Positivas", "Negativas"],
                    values=[positive2, total2 - positive2],
                    marker=dict(colors=["#28a745", "#dc3545"]),
                    textinfo='label+percent',
                    hoverinfo='label+value+percent',
                    textfont=dict(size=13, color='white'),
                    domain=dict(x=[0.5, 1]),
                    name="Período 2",
                ))
            fig_compare.update_layout(
                title="Comparação de Períodos",
                font=dict(size=13),
                height=320,
                showlegend=False,
                margin=dict(l=0, r=0, t=50, b=0),
                annotations=[
                    dict(text=f"Período 1<br>({total1} docs)", x=0.25, y=-0.1, showarrow=False),
                    dict(text=f"Período 2<br>({total2} docs)", x=0.75, y=-0.1, showarrow=False),
                ],
            )

            draw_plotly_on_canvas(fig_compare, margin, chart_y, width - (2 * margin), chart_h)
        else:
            y_position = draw_pdf_header("Dados Numéricos (base dos gráficos)")

            total_count = len(docs)
            positive_count = sum(1 for d in docs if d[11] == 1)
            negative_count = sum(1 for d in docs if d[11] == -1)
            pending_count = sum(1 for d in docs if d[11] not in (1, -1))

            summary_data = [
                ["Indicador", "Valor"],
                ["Total de DDU", str(total_count)],
                ["DDU Positivas", str(positive_count)],
                ["DDU Negativas", str(negative_count)],
                ["Doc em Andamento", str(pending_count)],
            ]

            table_summary = Table(summary_data, colWidths=[260, 120])
            table_summary.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003366")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (1, 1), (1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("GRID", (0, 0), (-1, -1), 1, colors.grey),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#f8f9fa")),
                ("WORDWRAP", (0, 0), (-1, -1), "CJK"),
            ]))

            left_area_width = (width - (2 * margin) - 20) / 2
            right_area_width = left_area_width

            tw, th = table_summary.wrapOn(c, left_area_width, y_position)
            table_summary.drawOn(c, margin, y_position - th)

            natureza_count = {}
            for doc in docs:
                nature = (doc[3] or "Não informado").strip()
                nature = normalize_text(nature).capitalize()
                natureza_count[nature] = natureza_count.get(nature, 0) + 1

            natureza_data = [["Natureza", "Quantidade"]]
            natureza_sorted = sorted(natureza_count.items(), key=lambda item: item[1], reverse=True)
            top_items = natureza_sorted[:8]
            other_total = sum(count for _, count in natureza_sorted[8:])

            for nature, count in top_items:
                natureza_data.append([nature, str(count)])
            if other_total > 0:
                natureza_data.append(["OUTRAS NATUREZAS", str(other_total)])
            if len(natureza_data) == 1:
                natureza_data.append(["NÃO INFORMADO", "0"])

            table_natureza = Table(natureza_data, colWidths=[right_area_width - 120, 120])
            table_natureza.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#495057")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (1, 1), (1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("GRID", (0, 0), (-1, -1), 1, colors.grey),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#ffffff")),
                ("WORDWRAP", (0, 0), (-1, -1), "CJK"),
            ]))

            tw2, th2 = table_natureza.wrapOn(c, right_area_width, y_position)
            table_natureza.drawOn(c, margin + left_area_width + 20, y_position - th2)

            chart_top = min(y_position - th, y_position - th2)
            chart_y = margin
            chart_top = draw_section_divider(chart_top)
            chart_h = chart_top - chart_y

            # Evita que os gráficos invadam a área da tabela; quando necessário,
            # move os gráficos para a página seguinte.
            if chart_h < 260:
                c.showPage()
                y_position_graph = draw_pdf_header("Gráficos do Dashboard")
                chart_top = y_position_graph
                chart_y = margin
                chart_top = draw_section_divider(chart_top)
                chart_h = chart_top - chart_y

            responded_total = positive_count + negative_count
            if responded_total > 0:
                labels = ["Positivas", "Negativas"]
                sizes = [positive_count, negative_count]
                pie_colors = ["#28a745", "#dc3545"]
                fig_pie = go.Figure(data=[go.Pie(
                    labels=labels,
                    values=sizes,
                    marker=dict(colors=pie_colors),
                    textinfo='label+percent',
                    hoverinfo='label+value+percent',
                    textfont=dict(size=13, color='white'),
                    pull=[0.05, 0],
                )])
                fig_pie.update_layout(
                    title="Respondidas: Positivo vs Negativo",
                    font=dict(size=13),
                    showlegend=True,
                    height=300,
                    margin=dict(l=0, r=0, t=40, b=0),
                )
            else:
                fig_pie = go.Figure()
                fig_pie.add_annotation(text="Sem dados respondidos", showarrow=False, font=dict(size=14), y=0.5)
                fig_pie.update_layout(title="Respondidas: Positivo vs Negativo", height=300, showlegend=False)

            labels_overall = ["Doc em Andamento", "Respondidos Positivos", "Respondidos Negativos"]
            sizes_overall = [pending_count, positive_count, negative_count]
            colors_overall = ["#FFC107", "#28a745", "#dc3545"]
            labels_overall_with_values = [
                f"{labels_overall[0]}: {sizes_overall[0]}",
                f"{labels_overall[1]}: {sizes_overall[1]}",
                f"{labels_overall[2]}: {sizes_overall[2]}",
            ]
            fig_overall = go.Figure(data=[go.Pie(
                labels=labels_overall_with_values,
                values=sizes_overall,
                marker=dict(colors=colors_overall, line=dict(color='white', width=2)),
                textinfo='label+percent',
                hoverinfo='label+value+percent',
                textfont=dict(size=10, color='white'),
                textposition='inside',
                insidetextorientation='radial',
                pull=[0.05, 0.05, 0.05],
            )])
            fig_overall.update_layout(
                title="Distribuição Geral do Sistema",
                font=dict(size=12),
                showlegend=True,
                height=300,
                margin=dict(l=10, r=10, t=45, b=10),
                uniformtext_minsize=8,
                uniformtext_mode='hide',
            )

            if natureza_count:
                natures = list(natureza_count.keys())
                counts = list(natureza_count.values())
                fig_natureza = go.Figure(data=[go.Pie(
                    labels=natures,
                    values=counts,
                    marker=dict(colors=px.colors.qualitative.Set3),
                    textinfo='label+percent+value',
                    textfont=dict(size=13),
                    hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>',
                )])
                fig_natureza.update_layout(
                    title="Documentos por Natureza",
                    font=dict(size=13),
                    height=300,
                    showlegend=True,
                    legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05),
                    margin=dict(l=20, r=120, t=40, b=20),
                )
            else:
                fig_natureza = go.Figure()
                fig_natureza.add_annotation(text="Sem dados disponíveis", showarrow=False, font=dict(size=14), y=0.5)
                fig_natureza.update_layout(height=300, showlegend=False)

            available_width = width - (2 * margin)
            gap = 12

            # Quando houver espaço vertical suficiente, os 3 gráficos ficam na mesma página.
            if chart_h >= 190:
                graph_width = (available_width - (2 * gap)) / 3
                draw_plotly_on_canvas(fig_pie, margin, chart_y, graph_width, chart_h)
                draw_plotly_on_canvas(fig_overall, margin + graph_width + gap, chart_y, graph_width, chart_h)
                draw_plotly_on_canvas(fig_natureza, margin + (2 * (graph_width + gap)), chart_y, graph_width, chart_h)
            else:
                graph_width = (available_width - gap) / 2
                graph_height = chart_h
                draw_plotly_on_canvas(fig_pie, margin, chart_y, graph_width, graph_height)
                draw_plotly_on_canvas(fig_overall, margin + graph_width + gap, chart_y, graph_width, graph_height)

                c.showPage()
                y_position_2 = draw_pdf_header("Gráfico complementar e dados por natureza")

                table_natureza_2 = Table(natureza_data, colWidths=[available_width - 120, 120])
                table_natureza_2.setStyle(TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#495057")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("ALIGN", (1, 1), (1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("GRID", (0, 0), (-1, -1), 1, colors.grey),
                    ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#ffffff")),
                    ("WORDWRAP", (0, 0), (-1, -1), "CJK"),
                ]))

                tw3, th3 = table_natureza_2.wrapOn(c, available_width, y_position_2)
                table_natureza_2.drawOn(c, margin, y_position_2 - th3)

                chart_top_2 = y_position_2 - th3 - 18
                chart_h_2 = max(200, chart_top_2 - margin)
                draw_plotly_on_canvas(fig_natureza, margin, margin, available_width, chart_h_2)
        
        c.save()
    
    def _parse_date(self, text):
        """Converte texto em data"""
        if not text or not text.strip():
            return None
        
        for fmt in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.datetime.strptime(text.strip(), fmt).date()
            except:
                continue
        return None
    
    def _clear_layout(self, layout):
        """Remove todos os widgets de um layout recursivamente"""
        while layout.count():
            item = layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                self._clear_layout(item.layout())
    
    def _filter_documents(self, docs, start_date, end_date):
        """Filtra documentos por período de validade"""
        filtered = []
        for doc in docs:
            edate = doc[7]  # expiry_date
            
            ed = None
            for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
                try:
                    ed = datetime.datetime.strptime(edate, fmt).date()
                    break
                except:
                    continue
            
            if ed is None:
                continue
            
            if start_date and ed < start_date:
                continue
            if end_date and ed > end_date:
                continue
            
            filtered.append(doc)
        
        return filtered
    
    def refresh_dashboard(self):
        """Atualiza os gráficos do dashboard"""
        # Limpar layout anterior
        while self.content_layout.count():
            item = self.content_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                self._clear_layout(item.layout())
        
        all_docs = list_documents()
        
        if not self.compare_mode:
            # Modo de visualização única
            start_date = self._parse_date(self.date_start_input.text())
            end_date = self._parse_date(self.date_end_input.text())
            
            docs = self._filter_documents(all_docs, start_date, end_date)
            self._create_single_view(docs, start_date, end_date)
        else:
            # Modo de comparação
            start1, end1 = self.period1
            start2, end2 = self.period2
            
            docs1 = self._filter_documents(all_docs, start1, end1)
            docs2 = self._filter_documents(all_docs, start2, end2)
            
            self._create_comparison_view(docs1, docs2, start1, end1, start2, end2)
    
    def _create_single_view(self, docs, start_date, end_date):
        """Cria visualização única com todos os gráficos"""
        # Primeira linha - Cards de estatísticas
        stats_layout = QHBoxLayout()
        
        total_count = len(docs)
        positive_count = sum(1 for d in docs if d[11] == 1)  # result_pos == 1
        negative_count = sum(1 for d in docs if d[11] == -1)  # result_pos == -1
        pending_count = sum(1 for d in docs if d[11] == 0)  # result_pos == 0 (não respondidos)
        
        # Card 1: Total de DDU
        card1 = self._create_stat_card("📋 Total de DDU", str(total_count), "#007BFF")
        stats_layout.addWidget(card1)
        
        # Card 2: DDU Positivas
        card2 = self._create_stat_card("✅ DDU Positivas", str(positive_count), "#28a745")
        stats_layout.addWidget(card2)
        
        # Card 3: DDU Negativas
        card3 = self._create_stat_card("❌ DDU Negativas", str(negative_count), "#dc3545")
        stats_layout.addWidget(card3)
        
        self.content_layout.addLayout(stats_layout)
        
        # Segunda linha - Gráficos
        if total_count > 0:
            graphs_layout = QHBoxLayout()
            
            # Gráfico 1 (Esquerda) - Pizza - Resultado Positivo vs Negativo (APENAS respondidos)
            labels_responded = ["Positivas", "Negativas"]
            sizes_responded = [positive_count, negative_count]
            colors_responded = ["#28a745", "#dc3545"]
            
            if positive_count > 0 or negative_count > 0:
                fig_pie_responded = go.Figure(data=[go.Pie(
                    labels=labels_responded,
                    values=sizes_responded,
                    marker=dict(colors=colors_responded),
                    textinfo='label+percent+value',
                    hoverinfo='label+value+percent',
                    textfont=dict(size=12, color='white'),
                    pull=[0.05, 0]
                )])
                fig_pie_responded.update_layout(
                    title="Respondidas: Positivo vs Negativo",
                    font=dict(size=12),
                    showlegend=True,
                    height=350,
                    margin=dict(l=0, r=0, t=40, b=0)
                )
            else:
                fig_pie_responded = go.Figure()
                fig_pie_responded.add_annotation(text="Sem dados disponíveis", showarrow=False, 
                                      font=dict(size=14), y=0.5)
                fig_pie_responded.update_layout(height=350, showlegend=False)
            
            pixmap_pie_responded = plotly_to_pixmap(fig_pie_responded, width=400, height=350)
            label_pie_responded = QLabel()
            if pixmap_pie_responded:
                label_pie_responded.setPixmap(pixmap_pie_responded)
            label_pie_responded.setAlignment(Qt.AlignCenter)
            graphs_layout.addWidget(label_pie_responded)
            
            # Gráfico 2 (Centro) - Pizza - Distribuição Geral (Doc em Andamento + Respondidos Pos + Respondidos Neg)
            labels_overall = ["Doc em Andamento", "Respondidos Positivos", "Respondidos Negativos"]
            sizes_overall = [pending_count, positive_count, negative_count]
            colors_overall = ["#FFC107", "#28a745", "#dc3545"]
            
            # Formatar labels com valores explícitos
            labels_with_values = [
                f"{labels_overall[0]}: {sizes_overall[0]}",
                f"{labels_overall[1]}: {sizes_overall[1]}",
                f"{labels_overall[2]}: {sizes_overall[2]}"
            ]
            
            if total_count > 0:
                fig_pie_overall = go.Figure(data=[go.Pie(
                    labels=labels_with_values,
                    values=sizes_overall,
                    marker=dict(colors=colors_overall, line=dict(color='white', width=2)),
                    textinfo='label+percent',
                    hoverinfo='label+value+percent',
                    textfont=dict(size=10, color='white', family='Arial'),
                    textposition='inside',
                    insidetextorientation='radial',
                    pull=[0.05, 0.05, 0.05]
                )])
                fig_pie_overall.update_layout(
                    title="Distribuição Geral do Sistema",
                    font=dict(size=12),
                    showlegend=True,
                    height=400,
                    margin=dict(l=10, r=10, t=50, b=10),
                    uniformtext_minsize=8,
                    uniformtext_mode='hide'
                )
            else:
                fig_pie_overall = go.Figure()
                fig_pie_overall.add_annotation(text="Sem dados disponíveis", showarrow=False,
                                      font=dict(size=14), y=0.5)
                fig_pie_overall.update_layout(height=350, showlegend=False)
            
            pixmap_pie_overall = plotly_to_pixmap(fig_pie_overall, width=480, height=400)
            label_pie_overall = QLabel()
            if pixmap_pie_overall:
                label_pie_overall.setPixmap(pixmap_pie_overall)
            label_pie_overall.setAlignment(Qt.AlignCenter)
            graphs_layout.addWidget(label_pie_overall)
            
            # Gráfico Pizza - Por Natureza
            natureza_count = {}
            for doc in docs:
                nature = (doc[3] or "Não informado").strip()
                nature = normalize_text(nature).capitalize()
                natureza_count[nature] = natureza_count.get(nature, 0) + 1
            
            if natureza_count:
                natures = list(natureza_count.keys())
                counts = list(natureza_count.values())
                
                fig_natureza = go.Figure(data=[go.Pie(
                    labels=natures,
                    values=counts,
                    marker=dict(colors=px.colors.qualitative.Set3),
                    textinfo='label+percent+value',
                    textfont=dict(size=11),
                    hovertemplate='<b>%{label}</b><br>Quantidade: %{value}<br>Percentual: %{percent}<extra></extra>'
                )])
                fig_natureza.update_layout(
                    title="Documentos por Natureza",
                    font=dict(size=11),
                    height=350,
                    showlegend=True,
                    legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.05),
                    margin=dict(l=20, r=120, t=40, b=20)
                )
            else:
                fig_natureza = go.Figure()
                fig_natureza.add_annotation(text="Sem dados disponíveis", showarrow=False,
                                      font=dict(size=14), y=0.5)
                fig_natureza.update_layout(height=350, showlegend=False)
            
            pixmap_natureza = plotly_to_pixmap(fig_natureza, width=500, height=350)
            label_natureza = QLabel()
            if pixmap_natureza:
                label_natureza.setPixmap(pixmap_natureza)
            label_natureza.setAlignment(Qt.AlignCenter)
            graphs_layout.addWidget(label_natureza)
            
            self.content_layout.addLayout(graphs_layout)
        else:
            lbl_empty = QLabel("Nenhum documento encontrado para o período selecionado.")
            lbl_empty.setAlignment(Qt.AlignCenter)
            lbl_empty.setStyleSheet("color: #666; font-size: 14px; padding: 40px;")
            self.content_layout.addWidget(lbl_empty)
    
    def _create_comparison_view(self, docs1, docs2, start1, end1, start2, end2):
        """Cria visualização comparativa de dois períodos"""
        # Cabeçalho de períodos
        header_layout = QHBoxLayout()
        
        lbl_period1 = QLabel(f"📅 Período 1: {start1.strftime('%d/%m/%Y')} a {end1.strftime('%d/%m/%Y')}")
        lbl_period1.setStyleSheet("font-weight: bold; color: #003366; font-size: 12px; padding: 10px; "
                                 "background-color: #e3f2fd; border-radius: 4px;")
        header_layout.addWidget(lbl_period1)
        
        lbl_period2 = QLabel(f"📅 Período 2: {start2.strftime('%d/%m/%Y')} a {end2.strftime('%d/%m/%Y')}")
        lbl_period2.setStyleSheet("font-weight: bold; color: #003366; font-size: 12px; padding: 10px; "
                                 "background-color: #f3e5f5; border-radius: 4px;")
        header_layout.addWidget(lbl_period2)
        
        header_layout.addStretch()
        self.content_layout.addLayout(header_layout)
        
        # Cards com estatísticas
        stats_layout = QHBoxLayout()
        
        # Período 1
        total1 = len(docs1)
        positive1 = sum(1 for d in docs1 if d[11] == 1)
        negative1 = sum(1 for d in docs1 if d[11] == -1)
        
        period1_card = QWidget()
        period1_layout = QVBoxLayout(period1_card)
        period1_layout.setContentsMargins(10, 10, 10, 10)
        period1_card.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 8px;
                border: 2px solid #2196F3;
            }
        """)
        
        lbl1_total = QLabel(f"Total: {total1}")
        lbl1_total.setStyleSheet("font-weight: bold; color: #003366; font-size: 14px;")
        lbl1_positive = QLabel(f"Positivas: {positive1} ({(positive1*100//total1 if total1 > 0 else 0)}%)")
        lbl1_positive.setStyleSheet("color: #28a745; font-weight: bold;")
        
        period1_layout.addWidget(lbl1_total)
        period1_layout.addWidget(lbl1_positive)
        period1_layout.addStretch()
        
        stats_layout.addWidget(period1_card)
        
        # Período 2
        total2 = len(docs2)
        positive2 = sum(1 for d in docs2 if d[11] == 1)
        negative2 = sum(1 for d in docs2 if d[11] == -1)
        
        period2_card = QWidget()
        period2_layout = QVBoxLayout(period2_card)
        period2_layout.setContentsMargins(10, 10, 10, 10)
        period2_card.setStyleSheet("""
            QWidget {
                background-color: white;
                border-radius: 8px;
                border: 2px solid #9C27B0;
            }
        """)
        
        lbl2_total = QLabel(f"Total: {total2}")
        lbl2_total.setStyleSheet("font-weight: bold; color: #003366; font-size: 14px;")
        lbl2_positive = QLabel(f"Positivas: {positive2} ({(positive2*100//total2 if total2 > 0 else 0)}%)")
        lbl2_positive.setStyleSheet("color: #28a745; font-weight: bold;")
        
        period2_layout.addWidget(lbl2_total)
        period2_layout.addWidget(lbl2_positive)
        period2_layout.addStretch()
        
        stats_layout.addWidget(period2_card)
        stats_layout.addStretch()
        
        self.content_layout.addLayout(stats_layout)
        
        # Gráficos comparativos
        if total1 > 0 or total2 > 0:
            graphs_layout = QHBoxLayout()
            
            # Gráfico Pizza Comparativo usando Plotly
            fig_compare = go.Figure()
            
            # Período 1
            if total1 > 0:
                fig_compare.add_trace(go.Pie(
                    labels=["Positivas", "Negativas"],
                    values=[positive1, negative1],
                    marker=dict(colors=["#28a745", "#dc3545"]),
                    textinfo='label+percent',
                    hoverinfo='label+value+percent',
                    textfont=dict(size=10, color='white'),
                    domain=dict(x=[0, 0.5]),
                    name=f"Período 1"
                ))
            
            # Período 2
            if total2 > 0:
                fig_compare.add_trace(go.Pie(
                    labels=["Positivas", "Negativas"],
                    values=[positive2, negative2],
                    marker=dict(colors=["#28a745", "#dc3545"]),
                    textinfo='label+percent',
                    hoverinfo='label+value+percent',
                    textfont=dict(size=10, color='white'),
                    domain=dict(x=[0.5, 1]),
                    name=f"Período 2"
                ))
            
            fig_compare.update_layout(
                title="Comparação de Períodos",
                font=dict(size=11),
                height=400,
                showlegend=False,
                margin=dict(l=0, r=0, t=50, b=0),
                annotations=[
                    dict(text=f"Período 1<br>({total1} docs)", x=0.25, y=-0.1, showarrow=False),
                    dict(text=f"Período 2<br>({total2} docs)", x=0.75, y=-0.1, showarrow=False)
                ]
            )
            
            pixmap_compare = plotly_to_pixmap(fig_compare, width=800, height=400)
            label_compare = QLabel()
            if pixmap_compare:
                label_compare.setPixmap(pixmap_compare)
            label_compare.setAlignment(Qt.AlignCenter)
            graphs_layout.addWidget(label_compare)
            
            self.content_layout.addLayout(graphs_layout)
    
    def _create_stat_card(self, title, value, color):
        """Cria um card de estatística"""
        card = QWidget()
        card.setStyleSheet(f"""
            QWidget {{
                background-color: white;
                border-radius: 8px;
                border-left: 5px solid {color};
                padding: 20px;
            }}
        """)
        
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 10, 10, 10)
        
        lbl_title = QLabel(title)
        lbl_title.setStyleSheet("color: #666; font-size: 12px; font-weight: bold;")
        
        lbl_value = QLabel(value)
        lbl_value.setStyleSheet(f"color: {color}; font-size: 32px; font-weight: bold;")
        
        layout.addWidget(lbl_title)
        layout.addWidget(lbl_value)
        layout.addStretch()
        
        return card


def main():
    init_db()
    app = QApplication(sys.argv)
    if os.path.exists(ICON_PATH):
        app.setWindowIcon(QIcon(ICON_PATH))
    win = MainWindow()
    win.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
