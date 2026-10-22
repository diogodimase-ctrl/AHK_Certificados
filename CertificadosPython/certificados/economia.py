"""
Tempo economizado com o programa, só neste computador.

Cada rodada OFICIAL grava quantos certificados saíram e quanto tempo levou, da
saída da tela de Início até o Relatório. O jeito manual leva, em média,
7 min 33 s para montar cada certificado mais 4 min para anexar e enviar.
Economizado = manual - programa, somando todas as rodadas.

Fica em %APPDATA%\\CertificadosPython\\tempo.json: cada pessoa vê o próprio número.
"""

import json
import os
from dataclasses import dataclass
from datetime import datetime

MANUAL_POR_CERTIFICADO = 7 * 60 + 33 + 4 * 60          # 11 min 33 s
ARQUIVO = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "CertificadosPython", "tempo.json")


@dataclass
class Resumo:
    rodadas: int
    certificados: int
    manual: float          # segundos
    programa: float

    @property
    def economizado(self):
        return max(0.0, self.manual - self.programa)


def _ler(arquivo):
    try:
        with open(arquivo, encoding="utf-8") as f:
            dados = json.load(f)
        return dados.get("rodadas", []) if isinstance(dados, dict) else []
    except (OSError, ValueError):
        return []


def registrar(certificados, segundos, arquivo=ARQUIVO):
    """Acrescenta uma rodada. Devolve False se não conseguiu gravar."""
    if certificados <= 0 or segundos <= 0:
        return False
    rodadas = _ler(arquivo)
    rodadas.append({"quando": datetime.now().isoformat(timespec="seconds"), "certificados": int(certificados),
                    "segundos": round(float(segundos), 1)})
    try:
        os.makedirs(os.path.dirname(arquivo), exist_ok=True)
        with open(arquivo, "w", encoding="utf-8") as f:
            json.dump({"rodadas": rodadas}, f, ensure_ascii=False, indent=2)
        return True
    except OSError:
        return False


def resumo(arquivo=ARQUIVO):
    rodadas = _ler(arquivo)
    certificados = sum(int(r.get("certificados", 0)) for r in rodadas)
    programa = sum(float(r.get("segundos", 0)) for r in rodadas)
    return Resumo(len(rodadas), certificados, certificados * MANUAL_POR_CERTIFICADO, programa)


def duracao(segundos):
    """3725 -> "1 h 02 min"; 95 -> "1 min 35 s"."""
    s = int(round(segundos))
    h, resto = divmod(s, 3600)
    m, s = divmod(resto, 60)
    if h:
        return f"{h} h {m:02d} min"
    if m:
        return f"{m} min {s:02d} s"
    return f"{s} s"
