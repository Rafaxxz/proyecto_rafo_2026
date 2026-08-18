"""Utilidades compartidas por los modulos."""
import re
from datetime import datetime

DIAS = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo']


def sello_fecha(momento=None):
    """Sello legible: dia-semana, fecha y hora. Ej: lunes-17-08-2026_08-30-45."""
    m = momento or datetime.now()
    return f'{DIAS[m.weekday()]}-{m.strftime("%d-%m-%Y_%H-%M-%S")}'


def nombre_archivo(base, extension, momento=None):
    """Nombre unico para exportes: base + dia/fecha/hora + extension."""
    base = re.sub(r'[^A-Za-z0-9_-]+', '_', base).strip('_') or 'export'
    return f'{base}_{sello_fecha(momento)}.{extension.lstrip(".")}'
