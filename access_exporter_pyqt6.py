"""Access (.accdb/.mdb) naar Excel/CSV exporter, editor en terugzet-tool met een PyQt6-interface.

Opstartbestand. De code staat in het pakket `access_exporter`.
Vereisten: pip install -r requirements.txt en de "Microsoft Access Driver (*.mdb, *.accdb)"
met dezelfde bitness (32/64-bit) als Python.
"""

import sys
from pathlib import Path

from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import QApplication, QMessageBox

from access_exporter import __version__
from access_exporter.core import STATUS_BEZIG, STATUS_FOUT, STATUS_GESTOPT, STATUS_KLAAR, STATUS_WACHTRIJ
from access_exporter.dialogs import TableSelectionDialog
from access_exporter.main_window import AccessExporterGUI, toon_onverwachte_fout
from access_exporter.worker import ConversionWorker

__all__ = [
    "STATUS_BEZIG", "STATUS_FOUT", "STATUS_GESTOPT", "STATUS_KLAAR", "STATUS_WACHTRIJ", "AccessExporterGUI",
    "ConversionWorker", "QMessageBox", "TableSelectionDialog", "__version__", "main",
]


def icoon_pad() -> Path:
    # In de PyInstaller-exe staan de assets in sys._MEIPASS
    basis = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return basis / "access_exporter" / "assets" / "access-exporter.ico"


def main() -> int:
    sys.excepthook = toon_onverwachte_fout
    if sys.platform == "win32":
        try:
            import ctypes
            # Eigen taakbalkicoon in plaats van het Python-icoon
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(f"DevoteamNL.AccessExporter.{__version__}")
        except (AttributeError, OSError):
            pass
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setApplicationName("Access Exporter")
    app.setApplicationVersion(__version__)
    pad = icoon_pad()
    if pad.exists():
        app.setWindowIcon(QIcon(str(pad)))
    window = AccessExporterGUI()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
