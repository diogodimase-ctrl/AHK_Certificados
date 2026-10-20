"""
Carga horária e datas, com as mesmas regras do AHK (UI_Moderna + v27).
"""

import math
import re
from datetime import date, datetime, time, timedelta


def _arredondar(x):
    """Round do AHK (meio para longe do zero); o round() do Python é diferente."""
    return int(math.floor(x + 0.5)) if x >= 0 else -int(math.floor(-x + 0.5))


def _horas_texto(h):
    return "1 hora" if h == 1 else f"{h} horas"


def formatar_horas(texto_original):
    """FormatarHorasSeguro: "08:00 - 12:00" -> "4 horas"; "2h30" -> "2 horas e 30 minutos"."""
    txt = str(texto_original or "").strip()
    if not txt:
        return "Horario invalido"
    txt = re.sub(r"[   \t]", " ", txt)
    txt = re.sub(r"[‐-―−]", "-", txt)
    txt = re.sub("às", "-", txt, flags=re.I)
    txt = re.sub(r"\s+", " ", txt).strip()

    m = re.match(r"^(\d+)\s*[h:]\s*(\d+)\s*(min)?\s*(horas?|h)?\s*$", txt, re.I)
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
        if h > 0 and mi > 0:
            return f"{h} horas e {mi} minutos"
        if h > 0:
            return _horas_texto(h)
        if mi > 0:
            return f"{mi} minutos"
    m = re.match(r"^(\d+)\s*(horas?|h)\s*$", txt, re.I)
    if m:
        return _horas_texto(int(m.group(1)))
    m = re.match(r"^(\d+)\s*min(utos?)?\s*$", txt, re.I)
    if m:
        return f"{int(m.group(1))} minutos"

    m = re.search(r"(\d{1,2})\s*[:h]?\s*(\d{2})?\s*h?\s*[-–—−]\s*(\d{1,2})\s*[:h]?\s*(\d{2})?\s*h?", txt, re.I)
    if m:
        return calcular_duracao(f"{m.group(1)}:{m.group(2) or '00'}", f"{m.group(3)}:{m.group(4) or '00'}")

    return f"Horario invalido (formato não reconhecido: {texto_original})"


def calcular_duracao(inicio, fim):
    """CalcularDuracao: "08:00", "12:30" -> "4 horas e 30 minutos"."""
    try:
        hi, mi = (int(x) for x in inicio.split(":")[:2])
        hf, mf = (int(x) for x in fim.split(":")[:2])
    except ValueError as e:
        return f"Horario invalido ({e})"
    duracao = (hf * 60 + mf) - (hi * 60 + mi)
    if duracao <= 0:
        return "Horario invalido (fim <= inicio)"
    horas, minutos = divmod(duracao, 60)
    if horas > 0 and minutos > 0:
        return f"{_horas_texto(horas)} e {minutos} minutos"
    if horas > 0:
        return _horas_texto(horas)
    return f"{minutos} minutos"


def formatar_duracao_digitada(entrada):
    """FormatarDuracaoDigitada: "4" -> "4 horas"; "2:30" ou "2,30" -> "2 horas e 30 minutos"."""
    txt = str(entrada or "").strip().replace(",", ":")
    if not txt:
        return ""
    m = re.fullmatch(r"(\d+):(\d+)", txt)
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
        if h > 0 and mi > 0:
            return f"{h} horas e {mi} minutos"
        if h > 0:
            return _horas_texto(h)
        if mi > 0:
            return f"{mi} minutos"
    if re.fullmatch(r"\d+", txt) and int(txt) > 0:
        return _horas_texto(int(txt))
    return ""


def horas_de_valor_excel(valor):
    """FormatarHorasDeValorExcel: fração de dia do Excel -> "4 horas" ou "2h30"."""
    if not isinstance(valor, (int, float)) or isinstance(valor, bool):
        return "Horario invalido"
    total = _arredondar(valor * 24 * 60)
    horas, minutos = divmod(total, 60)
    return f"{horas}h{minutos:02d}" if minutos > 0 else f"{horas} horas"


# ---------------------------------------------------------------------------
# Datas
# ---------------------------------------------------------------------------

_DATA = re.compile(r"^\d{1,2}/\d{1,2}/\d{2,4}$")


def data_valida(txt):
    """EhDataValida: só confere o formato D/M/AA(AA)."""
    return bool(_DATA.match(str(txt or "").strip()))


def serial_para_data(serial):
    """ExcelSerialParaData: 46157 -> "15/05/2026".

    DIFERENÇA CONHECIDA: o AHK conta a partir de 01/01/1900 e por isso dá um
    dia a mais (o Excel considera que 1900 teve 29 de fevereiro). Aqui a conta
    é a do próprio Excel. Só faz diferença quando a célula de data não está
    escrita como dd/mm/aaaa (ex.: aparece "15-mai")."""
    if not isinstance(serial, (int, float)) or serial < 1:
        return ""
    try:
        return (datetime(1899, 12, 30) + timedelta(days=int(serial))).strftime("%d/%m/%Y")
    except OverflowError:
        return ""


def corrigir_data_digitada(valor):
    """CorrigirDataRegistro: "5/3/26" -> "05/03/2026". Devolve "" se inválida."""
    txt = str(valor or "").strip()
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", txt)
    if m:
        ano = "20" + m.group(3) if len(m.group(3)) == 2 else m.group(3)
        txt = f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{ano}"
    return txt if data_valida(txt) else ""


def para_serial(valor):
    """Data/hora do Python -> número do Excel (como o Value2 do AHK)."""
    if isinstance(valor, datetime):
        delta = valor - datetime(1899, 12, 30)
        return delta.days + delta.seconds / 86400
    if isinstance(valor, date):
        return (valor - date(1899, 12, 30)).days
    if isinstance(valor, time):
        return (valor.hour * 3600 + valor.minute * 60 + valor.second) / 86400
    if isinstance(valor, timedelta):
        return valor.total_seconds() / 86400
    return valor
