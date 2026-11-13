"""Owned host QtSvg raster child. This is not the WebView or a filesystem sandbox."""
import base64
from pathlib import Path
import re
import socket
import sys
import time

if __package__ in (None, ''):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from office_agent.diagram_channel import MAGIC, receive_message, send_message
from office_agent.diagrams import validate_svg


def main(arguments=None):
    arguments = sys.argv[1:] if arguments is None else arguments
    if (len(arguments) != 2 or not re.fullmatch(r'[0-9]{1,5}', arguments[0])
            or not 1 <= int(arguments[0]) <= 65535
            or not re.fullmatch(r'[0-9a-f]{64}', arguments[1])):
        return 2
    outcome = {'ok': False, 'code': 'raster_failed'}
    with socket.create_connection(('127.0.0.1', int(arguments[0])), timeout=1) as connection:
        connection.settimeout(0.1)
        connection.sendall(MAGIC + bytes.fromhex(arguments[1]))
        try:
            request = receive_message(connection, time.monotonic() + 8)
            if set(request) != {'svg'} or type(request['svg']) is not str:
                raise ValueError()
            svg = validate_svg(request['svg'])
            from xml.etree import ElementTree as ET
            root = ET.fromstring(svg)
            width, height = (float(n) for n in root.attrib['viewBox'].split()[2:])
            width, height = int(round(width)), int(round(height))
            if (not 1 <= width <= 2048 or not 1 <= height <= 2048
                    or width * height > 4_194_304):
                raise ValueError()
            from PySide6.QtCore import QByteArray, QBuffer, QIODevice
            from PySide6.QtGui import QImage, QPainter
            from PySide6.QtSvg import QSvgRenderer
            renderer = QSvgRenderer(QByteArray(svg.encode('utf-8')))
            if not renderer.isValid():
                raise ValueError()
            image = QImage(width, height, QImage.Format_ARGB32)
            if image.isNull():
                raise ValueError()
            image.fill(0xFFFFFFFF)
            painter = QPainter(image)
            try:
                renderer.render(painter)
            finally:
                painter.end()
            buffer = QBuffer()
            if not buffer.open(QIODevice.WriteOnly) or not image.save(buffer, 'PNG'):
                raise ValueError()
            png = bytes(buffer.data())
            if not 0 < len(png) <= 1_000_000:
                raise ValueError()
            outcome = {'ok': True, 'png': base64.b64encode(png).decode('ascii')}
        except Exception:
            pass
        send_message(connection, outcome)
    return 0 if outcome['ok'] else 1


if __name__ == '__main__':
    raise SystemExit(main())
