"""Genereert PNG- en Windows-ICO-bestanden uit het originele SVG-icoon.
Generates PNG and Windows ICO files from the original SVG icon."""

import os
import struct
import sys
from pathlib import Path

from PyQt6.QtCore import QBuffer, QByteArray, QIODevice, QRectF, Qt
from PyQt6.QtGui import QImage, QPainter
from PyQt6.QtSvg import QSvgRenderer
from PyQt6.QtWidgets import QApplication

ASSET_DIR = Path(__file__).resolve().parents[1] / "access_exporter" / "assets"
ICON_SIZES = (16, 20, 24, 32, 40, 48, 64, 96, 128, 256)


def render_icon(renderer: QSvgRenderer, size: int) -> QImage:
    """Rendert groter en verkleint zorgvuldig voor scherpe kleine iconen.
    Renders larger and downsamples carefully for crisp small icons."""
    image = QImage(size * 4, size * 4, QImage.Format.Format_ARGB32)
    image.fill(Qt.GlobalColor.transparent)
    painter = QPainter(image)
    renderer.render(painter, QRectF(0, 0, image.width(), image.height()))
    painter.end()
    return image.scaled(size, size, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)


def png_bytes(image: QImage) -> bytes:
    """Codeert een afbeelding als PNG en meldt schrijffouten.
    Encodes an image as PNG and reports write failures."""
    data = QByteArray()
    buffer = QBuffer(data)
    if not buffer.open(QIODevice.OpenModeFlag.WriteOnly) or not image.save(buffer, "PNG"):
        raise OSError("Could not encode icon as PNG")
    buffer.close()
    return data.data()


def main() -> int:
    """Bouwt reproduceerbare icoonbestanden zonder extra afbeeldingsbibliotheken.
    Builds reproducible icon files without additional image libraries."""
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    app = QApplication(sys.argv)
    renderer = QSvgRenderer(str(ASSET_DIR / "access-exporter.svg"))
    if not renderer.isValid():
        raise ValueError("Invalid application SVG icon")
    images = [png_bytes(render_icon(renderer, size)) for size in ICON_SIZES]
    offset = 6 + 16 * len(images)
    directory = bytearray(struct.pack("<HHH", 0, 1, len(images)))
    for size, image in zip(ICON_SIZES, images, strict=True):
        directory.extend(struct.pack("<BBBBHHII", size % 256, size % 256, 0, 0, 1, 32, len(image), offset))
        offset += len(image)
    (ASSET_DIR / "access-exporter.ico").write_bytes(directory + b"".join(images))
    (ASSET_DIR / "access-exporter.png").write_bytes(png_bytes(render_icon(renderer, 1024)))
    print(f"Created ICO ({', '.join(map(str, ICON_SIZES))} px) and 1024 px PNG in {ASSET_DIR}")
    app.quit()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
