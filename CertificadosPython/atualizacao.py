"""
Aviso de versão nova: pergunta ao GitHub qual é a última versão publicada
(página Releases do repositório). Sem internet ou com o GitHub bloqueado,
simplesmente não avisa.
"""

import json
import re
import urllib.request

REPOSITORIO = "diogodimase-ctrl/AHK_Certificados"
PAGINA_DOWNLOADS = f"https://github.com/{REPOSITORIO}/releases/latest"
_API = f"https://api.github.com/repos/{REPOSITORIO}/releases/latest"


def _numeros(versao):
    return tuple(int(n) for n in re.findall(r"\d+", versao or "")[:3])


def versao_mais_nova(atual, tempo_limite=6):
    """(versão, link para baixar) se o GitHub tiver uma versão maior que "atual"; senão None."""
    pedido = urllib.request.Request(_API, headers={"User-Agent": "Certificados", "Accept": "application/vnd.github+json"})
    with urllib.request.urlopen(pedido, timeout=tempo_limite) as resposta:
        dados = json.load(resposta)
    nova = dados.get("tag_name", "").lstrip("vV")
    if not nova or _numeros(nova) <= _numeros(atual):
        return None
    instalador = next((a.get("browser_download_url") for a in dados.get("assets", [])
                       if a.get("name", "").lower().endswith(".exe")), None)
    return nova, instalador or dados.get("html_url") or PAGINA_DOWNLOADS
