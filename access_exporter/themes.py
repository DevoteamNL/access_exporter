"""Lichte en donkere kleurenschema's en de bijbehorende Qt-stylesheet."""

from pathlib import Path

from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor, QGuiApplication, QPalette

THEMA_SYSTEEM = "system"
THEMA_LICHT = "light"
THEMA_DONKER = "dark"
THEMAS = (THEMA_SYSTEEM, THEMA_LICHT, THEMA_DONKER)

KLEUREN: dict[str, dict[str, str]] = {
    THEMA_DONKER: {
        "achtergrond": "#121212", "vlak": "#1E1E1E", "vlak_alt": "#252525", "knop": "#2D2D2D", "knop_hover": "#3D3D3D",
        "tekst": "#E0E0E0", "kop": "#FFFFFF", "rand": "#333333", "rand_knop": "#444444", "rand_indicator": "#8A8A8A", "uit": "#666666",
        "uit_vlak": "#1E1E1E", "accent": "#0078D7", "accent_hover": "#005a9e", "selectie": "#264F78",
        "ok": "#4CAF50", "fout": "#F44336", "info": "#569cd6", "bezig": "#FFB74D",
        "succes_knop": "#28a745", "succes_hover": "#218838", "waarschuwing_knop": "#fd7e14", "waarschuwing_hover": "#e36209",
    },
    THEMA_LICHT: {
        "achtergrond": "#F5F6F8", "vlak": "#FFFFFF", "vlak_alt": "#F0F2F5", "knop": "#FFFFFF", "knop_hover": "#E8F1FB",
        "tekst": "#1F2328", "kop": "#111111", "rand": "#D0D7DE", "rand_knop": "#B8C0CA", "rand_indicator": "#6E7781", "uit": "#9AA0A6",
        "uit_vlak": "#EEF0F2", "accent": "#0067C0", "accent_hover": "#005299", "selectie": "#CCE4F7",
        "ok": "#1E7E34", "fout": "#C62828", "info": "#0B5CAD", "bezig": "#B45309",
        "succes_knop": "#1E7E34", "succes_hover": "#196C2C", "waarschuwing_knop": "#C2560C", "waarschuwing_hover": "#A34808",
    },
}

# Vinkje voor aangevinkte selectievakjes; ook in de exe staat assets/ naast deze module
VINKJE = (Path(__file__).resolve().parent / "assets" / "check.svg").as_posix()

_actief = THEMA_DONKER


def systeem_is_donker() -> bool:
    hints = QGuiApplication.styleHints()
    return hints is not None and hints.colorScheme() == Qt.ColorScheme.Dark


def los_op(thema: str) -> str:
    """Systeem -> licht of donker, volgens de Windows-instelling."""
    if thema == THEMA_SYSTEEM:
        return THEMA_DONKER if systeem_is_donker() else THEMA_LICHT
    return thema if thema in KLEUREN else THEMA_DONKER


def actief() -> str:
    return _actief


def kleur(naam: str) -> str:
    return KLEUREN[_actief][naam]


def stylesheet(thema: str) -> str:
    k = KLEUREN[thema]
    return f"""
        QMainWindow, QDialog, QWidget {{ background-color: {k['achtergrond']}; color: {k['tekst']}; }}
        QGroupBox {{ font-weight: bold; border: 1px solid {k['rand']}; border-radius: 5px; margin-top: 10px; padding-top: 15px; }}
        QGroupBox::title {{ subcontrol-origin: margin; left: 10px; padding: 0 3px; color: {k['kop']}; }}
        QPushButton {{ background-color: {k['knop']}; border: 1px solid {k['rand_knop']}; border-radius: 4px; padding: 6px; font-size: 13px; color: {k['tekst']}; }}
        QPushButton:hover {{ background-color: {k['knop_hover']}; border-color: {k['accent']}; }}
        QPushButton#accentButton {{ background-color: {k['accent']}; color: white; border: none; font-weight: bold; }}
        QPushButton#accentButton:hover {{ background-color: {k['accent_hover']}; }}
        QPushButton#successButton {{ background-color: {k['succes_knop']}; color: white; border: none; font-weight: bold; }}
        QPushButton#successButton:hover {{ background-color: {k['succes_hover']}; }}
        QPushButton#warningButton {{ background-color: {k['waarschuwing_knop']}; color: white; border: none; font-weight: bold; }}
        QPushButton#warningButton:hover {{ background-color: {k['waarschuwing_hover']}; }}
        QPushButton:disabled, QPushButton#accentButton:disabled, QPushButton#successButton:disabled, QPushButton#warningButton:disabled {{ background-color: {k['uit_vlak']}; color: {k['uit']}; border: 1px solid {k['rand']}; }}
        QTableWidget, QTableView, QListWidget {{ background-color: {k['vlak']}; alternate-background-color: {k['vlak_alt']}; border: 1px solid {k['rand']}; color: {k['tekst']}; gridline-color: {k['rand']}; selection-background-color: {k['selectie']}; selection-color: {k['kop']}; }}
        QHeaderView::section {{ background-color: {k['knop']}; padding: 4px; border: 1px solid {k['rand']}; font-weight: bold; color: {k['kop']}; }}
        QPlainTextEdit, QLineEdit, QComboBox, QSpinBox {{ background-color: {k['vlak']}; border: 1px solid {k['rand']}; color: {k['tekst']}; }}
        QComboBox QAbstractItemView {{ background-color: {k['vlak']}; color: {k['tekst']}; selection-background-color: {k['selectie']}; }}
        QPlainTextEdit {{ font-family: Consolas, monospace; }}
        QLineEdit {{ padding: 4px; border-radius: 4px; }}
        QLineEdit:focus {{ border-color: {k['accent']}; }}
        QProgressBar {{ border: 1px solid {k['rand_knop']}; border-radius: 4px; text-align: center; background-color: {k['vlak']}; color: {k['kop']}; }}
        QProgressBar::chunk {{ background-color: {k['accent']}; width: 20px; }}
        QLabel, QRadioButton, QCheckBox {{ color: {k['tekst']}; background: transparent; }}
        QLabel#infoLabel {{ color: {k['info']}; }}
        QRadioButton:disabled, QCheckBox:disabled, QLabel:disabled {{ color: {k['uit']}; }}
        QCheckBox::indicator, QRadioButton::indicator, QAbstractItemView::indicator {{ width: 14px; height: 14px; border: 1px solid {k['rand_indicator']}; background-color: {k['vlak']}; }}
        QCheckBox::indicator, QAbstractItemView::indicator {{ border-radius: 3px; }}
        QRadioButton::indicator {{ border-radius: 8px; }}
        QCheckBox::indicator:hover, QRadioButton::indicator:hover, QAbstractItemView::indicator:hover {{ border-color: {k['accent']}; }}
        QCheckBox::indicator:checked, QAbstractItemView::indicator:checked {{ background-color: {k['accent']}; border-color: {k['accent']}; image: url("{VINKJE}"); }}
        QRadioButton::indicator:checked {{ background-color: qradialgradient(cx:0.5, cy:0.5, radius:0.5, fx:0.5, fy:0.5, stop:0 white, stop:0.35 white, stop:0.45 {k['accent']}, stop:1 {k['accent']}); border-color: {k['accent']}; }}
        QCheckBox::indicator:disabled, QRadioButton::indicator:disabled {{ border-color: {k['rand']}; background-color: {k['uit_vlak']}; }}
        QTabWidget::pane {{ border: 1px solid {k['rand']}; border-radius: 4px; top: -1px; }}
        QTabBar::tab {{ background: {k['knop']}; color: {k['tekst']}; border: 1px solid {k['rand']}; padding: 7px 16px; margin-right: 2px; border-top-left-radius: 4px; border-top-right-radius: 4px; }}
        QTabBar::tab:selected {{ background: {k['achtergrond']}; color: {k['kop']}; border-bottom-color: {k['achtergrond']}; font-weight: bold; }}
        QTabBar::tab:hover {{ border-color: {k['accent']}; }}
        QToolTip {{ background-color: {k['vlak']}; color: {k['tekst']}; border: 1px solid {k['rand_knop']}; }}
    """


def palet(thema: str) -> QPalette:
    k = KLEUREN[thema]
    p = QPalette()
    rollen = {
        QPalette.ColorRole.Window: "achtergrond", QPalette.ColorRole.WindowText: "tekst",
        QPalette.ColorRole.Base: "vlak", QPalette.ColorRole.AlternateBase: "vlak_alt",
        QPalette.ColorRole.Text: "tekst", QPalette.ColorRole.Button: "knop", QPalette.ColorRole.ButtonText: "tekst",
        QPalette.ColorRole.Highlight: "selectie", QPalette.ColorRole.HighlightedText: "kop",
        QPalette.ColorRole.ToolTipBase: "vlak", QPalette.ColorRole.ToolTipText: "tekst",
        QPalette.ColorRole.PlaceholderText: "uit", QPalette.ColorRole.Link: "info",
    }
    for rol, naam in rollen.items():
        p.setColor(rol, QColor(k[naam]))
    for rol in (QPalette.ColorRole.Text, QPalette.ColorRole.WindowText, QPalette.ColorRole.ButtonText):
        p.setColor(QPalette.ColorGroup.Disabled, rol, QColor(k["uit"]))
    return p


def pas_toe(thema: str) -> str:
    """Past het thema toe op de hele applicatie (ook dialogen) en geeft 'light' of 'dark' terug."""
    global _actief
    _actief = los_op(thema)
    app = QGuiApplication.instance()
    if isinstance(app, QGuiApplication):
        app.setPalette(palet(_actief))
        from PyQt6.QtWidgets import QApplication
        if isinstance(app, QApplication):
            app.setStyleSheet(stylesheet(_actief))
    return _actief
