"""Dedicated no-Agent-API WebView2 child; called only by the host supervisor."""
import json
from pathlib import Path
import re
import socket
import sys
import threading
import time

if __package__ in (None, ''):
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from office_agent.diagram_channel import MAGIC, receive_message, send_message
from office_agent.diagrams import DiagramError, validate_source


def _render_script(source):
    return ("OfficeDiagramRenderer.render(" + json.dumps(source) + ")"
            ".then(r=>({ok:true,svg:r.svg,bridge_absent:r.bridgeAbsent===true && Object.keys(window.pywebview?.api||{}).length===0}))"
            ".catch(e=>({ok:false,code:['unsafe_svg','unsupported_source','render_timeout'].includes(e.code)?e.code:'renderer_failed'}))")


def main(arguments=None):
    arguments = sys.argv[1:] if arguments is None else arguments
    if (len(arguments) != 2 or not re.fullmatch(r'[0-9]{1,5}', arguments[0])
            or not 1 <= int(arguments[0]) <= 65535
            or not re.fullmatch(r'[0-9a-f]{64}', arguments[1])):
        return 2
    outcome = {'ok': False, 'code': 'renderer_failed'}
    with socket.create_connection(('127.0.0.1', int(arguments[0])), timeout=1) as connection:
        connection.settimeout(0.1)
        connection.sendall(MAGIC + bytes.fromhex(arguments[1]))
        try:
            request = receive_message(connection, time.monotonic() + 8)
            if set(request) != {'source', 'kind'}:
                raise DiagramError('unsupported_source')
            spec = validate_source(request['source'], request['kind'])
            import webview  # No app/runtime/provider/bootstrap imports or dotenv loading.
            window = webview.create_window('LiX local diagram worker',
                str(Path(__file__).parent / 'webui/static/diagram_host.html'),
                js_api=None, width=1200, height=1000, hidden=True)
            def loaded():
                nonlocal outcome
                ready, values = threading.Event(), []
                def complete(value):
                    values.append(value)
                    ready.set()
                try:
                    window.evaluate_js(_render_script(spec.source), callback=complete)
                    if not ready.wait(10):
                        outcome = {'ok': False, 'code': 'render_timeout'}
                    elif values and type(values[0]) is dict:
                        value = values[0]
                        if value.get('ok') is True and value.get('bridge_absent') is True and isinstance(value.get('svg'), str):
                            outcome = {'ok': True, 'svg': value['svg'], 'bridge_absent': True}
                        else:
                            outcome = {'ok': False, 'code': value.get('code') if value.get('code') in ('unsafe_svg', 'unsupported_source', 'render_timeout') else 'renderer_failed'}
                except Exception:
                    outcome = {'ok': False, 'code': 'renderer_failed'}
                finally:
                    window.destroy()
            window.events.loaded += loaded
            webview.start(gui='edgechromium', private_mode=True, debug=False)
        except DiagramError as exc:
            outcome = {'ok': False, 'code': exc.code if exc.code in ('unsafe_svg', 'unsupported_source') else 'renderer_failed'}
        except Exception:
            outcome = {'ok': False, 'code': 'renderer_failed'}
        send_message(connection, outcome)
    return 0 if outcome['ok'] else 1


if __name__ == '__main__':
    raise SystemExit(main())
