"""Inicializador com registro de erros visível ao usuário."""
import ctypes
import runpy
import traceback
from pathlib import Path

raiz = Path(__file__).resolve().parent
try:
    runpy.run_path(str(raiz / 'interface.py'), run_name='__main__')
except Exception:
    log = raiz / 'erro_inicializacao.log'
    log.write_text(traceback.format_exc(), encoding='utf-8')
    ctypes.windll.user32.MessageBoxW(None, f'Não foi possível abrir o bot. Detalhes em:\n{log}', 'botBMG', 16)
