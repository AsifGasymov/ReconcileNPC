"""Dark navy financial theme — palette, fonts, QSS."""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QPalette
from PySide6.QtWidgets import QApplication


# ─── Palette ──────────────────────────────────────────────────────────────────
BG_BASE = "#0B1220"
PANEL = "#111A2E"
CARD = "#16243B"
CARD_HOVER = "#1B2C4A"
CARD_ACTIVE = "#22365C"
BORDER = "#1E2D4A"
BORDER_STRONG = "#2A3F66"

TEXT = "#E5EBF5"
TEXT_DIM = "#8B9BB4"
TEXT_MUTED = "#5A6B85"

ACCENT = "#3B82F6"
ACCENT_HOVER = "#60A5FA"
ACCENT_PRESSED = "#2563EB"

SUCCESS = "#10B981"
WARNING = "#F59E0B"
DANGER = "#EF4444"
GOLD = "#FCD34D"


# ─── Fonts ────────────────────────────────────────────────────────────────────
def _ui_font_family() -> str:
    from PySide6.QtGui import QFontDatabase
    available = QFontDatabase.families()
    for f in ("Inter", "SF Pro Text", "Segoe UI", "Helvetica Neue", "Arial"):
        if f in available:
            return f
    return "sans-serif"


def _mono_font_family() -> str:
    from PySide6.QtGui import QFontDatabase
    available = QFontDatabase.families()
    for f in ("JetBrains Mono", "SF Mono", "Menlo", "Consolas", "Courier New"):
        if f in available:
            return f
    return "monospace"


# ─── QSS ──────────────────────────────────────────────────────────────────────
def _build_qss() -> str:
    ui = _ui_font_family()
    return f"""
    * {{
        font-family: "{ui}";
        color: {TEXT};
        outline: none;
    }}
    QMainWindow, QWidget#root {{
        background: {BG_BASE};
    }}
    QWidget#sidebar {{
        background: {PANEL};
        border-right: 1px solid {BORDER};
    }}
    QWidget#topbar {{
        background: {PANEL};
        border-bottom: 1px solid {BORDER};
    }}
    QLabel#brand {{
        color: {TEXT};
        font-size: 16px;
        font-weight: 700;
        letter-spacing: 1px;
    }}
    QLabel#brandSub {{
        color: {TEXT_DIM};
        font-size: 11px;
        letter-spacing: 0.5px;
    }}
    QLabel#h1 {{
        color: {TEXT};
        font-size: 20px;
        font-weight: 700;
    }}
    QLabel#h2 {{
        color: {TEXT};
        font-size: 14px;
        font-weight: 600;
    }}
    QLabel#hint {{
        color: {TEXT_DIM};
        font-size: 12px;
    }}
    QLabel#muted {{
        color: {TEXT_MUTED};
        font-size: 11px;
    }}

    /* Sidebar nav buttons */
    QPushButton[nav="true"] {{
        text-align: left;
        background: transparent;
        border: none;
        border-radius: 8px;
        padding: 10px 14px;
        color: {TEXT_DIM};
        font-size: 13px;
        font-weight: 500;
    }}
    QPushButton[nav="true"]:hover {{
        background: {CARD_HOVER};
        color: {TEXT};
    }}
    QPushButton[nav="true"][active="true"] {{
        background: {CARD_ACTIVE};
        color: {TEXT};
        font-weight: 600;
    }}

    /* Cards (file pickers, etc.) */
    QFrame#card {{
        background: {CARD};
        border: 1px solid {BORDER};
        border-radius: 10px;
    }}
    QFrame#card[drop="true"] {{
        border: 1px dashed {ACCENT};
        background: {CARD_HOVER};
    }}
    QFrame#card[filled="true"] {{
        border: 1px solid {BORDER_STRONG};
    }}

    QLabel#cardTitle {{
        color: {TEXT};
        font-size: 13px;
        font-weight: 600;
    }}
    QLabel#cardSubtitle {{
        color: {TEXT_DIM};
        font-size: 11px;
    }}
    QLabel#fileName {{
        color: {TEXT};
        font-size: 12px;
        font-weight: 500;
    }}
    QLabel#fileMeta {{
        color: {TEXT_MUTED};
        font-size: 10px;
    }}

    /* Buttons */
    QPushButton#primary {{
        background: {ACCENT};
        color: white;
        border: none;
        border-radius: 8px;
        padding: 12px 20px;
        font-size: 13px;
        font-weight: 600;
    }}
    QPushButton#primary:hover {{
        background: {ACCENT_HOVER};
    }}
    QPushButton#primary:pressed {{
        background: {ACCENT_PRESSED};
    }}
    QPushButton#primary:disabled {{
        background: {CARD};
        color: {TEXT_MUTED};
    }}

    QPushButton#secondary {{
        background: {CARD};
        color: {TEXT};
        border: 1px solid {BORDER_STRONG};
        border-radius: 8px;
        padding: 8px 14px;
        font-size: 12px;
    }}
    QPushButton#secondary:hover {{
        background: {CARD_HOVER};
        border-color: {ACCENT};
    }}

    QPushButton#ghost {{
        background: transparent;
        color: {TEXT_DIM};
        border: none;
        padding: 4px 8px;
        font-size: 12px;
    }}
    QPushButton#ghost:hover {{
        color: {TEXT};
    }}
    QPushButton#danger {{
        background: transparent;
        color: {DANGER};
        border: none;
        padding: 4px 8px;
        font-size: 11px;
    }}
    QPushButton#danger:hover {{
        color: #FF6B6B;
    }}

    /* Progress */
    QProgressBar {{
        background: {CARD};
        border: 1px solid {BORDER};
        border-radius: 6px;
        height: 8px;
        text-align: center;
        color: {TEXT_DIM};
        font-size: 10px;
    }}
    QProgressBar::chunk {{
        background: {ACCENT};
        border-radius: 5px;
    }}

    /* Inputs */
    QLineEdit, QPlainTextEdit {{
        background: {CARD};
        border: 1px solid {BORDER};
        border-radius: 6px;
        padding: 6px 10px;
        color: {TEXT};
        selection-background-color: {ACCENT};
    }}
    QLineEdit:focus, QPlainTextEdit:focus {{
        border-color: {ACCENT};
    }}
    QLineEdit:read-only {{
        color: {TEXT_DIM};
    }}

    /* Scrollbars */
    QScrollBar:vertical {{
        background: {BG_BASE};
        width: 10px;
        border: none;
    }}
    QScrollBar::handle:vertical {{
        background: {BORDER_STRONG};
        border-radius: 5px;
        min-height: 30px;
    }}
    QScrollBar::handle:vertical:hover {{
        background: {ACCENT};
    }}
    QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
        height: 0;
    }}

    /* List */
    QListWidget {{
        background: transparent;
        border: none;
        font-size: 12px;
    }}
    QListWidget::item {{
        background: {CARD};
        border: 1px solid {BORDER};
        border-radius: 6px;
        padding: 6px 10px;
        margin: 3px 0;
        color: {TEXT};
    }}
    QListWidget::item:hover {{
        background: {CARD_HOVER};
    }}
    QListWidget::item:selected {{
        background: {CARD_ACTIVE};
        border-color: {ACCENT};
    }}

    /* Status pill */
    QLabel#statusOk {{
        color: {SUCCESS};
        font-size: 12px;
        font-weight: 500;
    }}
    QLabel#statusErr {{
        color: {DANGER};
        font-size: 12px;
        font-weight: 500;
    }}
    QLabel#statusWarn {{
        color: {WARNING};
        font-size: 12px;
    }}

    /* Result panel */
    QFrame#resultCard {{
        background: {CARD};
        border: 1px solid {SUCCESS};
        border-radius: 10px;
    }}
    QLabel#resultTitle {{
        color: {SUCCESS};
        font-size: 14px;
        font-weight: 700;
    }}
    QLabel#resultStat {{
        color: {TEXT};
        font-size: 12px;
    }}
    QLabel#resultStatNum {{
        color: {GOLD};
        font-size: 20px;
        font-weight: 700;
        font-family: "{_mono_font_family()}";
    }}

    /* Log stream */
    QPlainTextEdit#logStream {{
        background: {BG_BASE};
        border: 1px solid {BORDER};
        border-radius: 6px;
        font-family: "{_mono_font_family()}";
        font-size: 11px;
        color: {TEXT_DIM};
        padding: 8px;
    }}
    """


def apply_theme(app: QApplication) -> None:
    base_font = QFont(_ui_font_family(), 10)
    app.setFont(base_font)

    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(BG_BASE))
    palette.setColor(QPalette.WindowText, QColor(TEXT))
    palette.setColor(QPalette.Base, QColor(CARD))
    palette.setColor(QPalette.AlternateBase, QColor(CARD_HOVER))
    palette.setColor(QPalette.Text, QColor(TEXT))
    palette.setColor(QPalette.Button, QColor(CARD))
    palette.setColor(QPalette.ButtonText, QColor(TEXT))
    palette.setColor(QPalette.Highlight, QColor(ACCENT))
    palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
    palette.setColor(QPalette.PlaceholderText, QColor(TEXT_MUTED))
    app.setPalette(palette)

    app.setStyleSheet(_build_qss())
