import threading

from analizador import create_app, encolar as _encolar, INBOX

_app = None
_iniciado = False
_lock = threading.Lock()


def iniciar():
    global _app, _iniciado
    with _lock:
        if _iniciado:
            return INBOX
        _app = create_app()
        hilo = threading.Thread(
            target=lambda: _app.run(host='127.0.0.1', port=5000, debug=False,
                                    use_reloader=False, threaded=True),
            daemon=True)
        hilo.start()
        _iniciado = True
    return INBOX


def encolar(path):
    _encolar(path)
