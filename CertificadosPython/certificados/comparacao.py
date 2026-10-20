"""
Compara os PDFs gerados pelo Python com os que o AHK já gerou: mesmos nomes
de arquivo e mesmo texto dentro de cada PDF.
"""

import difflib
import os
import re
from dataclasses import dataclass

from .gerador import PREFIXO_TESTE

IGUAL, DIFERENTE, SO_AHK, SO_PYTHON = "igual", "diferente", "só no AHK", "só no Python"


@dataclass
class Linha:
    arquivo: str
    situacao: str
    detalhe: str = ""


def _texto_pdf(caminho):
    from pypdf import PdfReader
    leitor = PdfReader(caminho)
    return " ".join(" ".join((p.extract_text() or "").split()) for p in leitor.pages)


def _pdfs(pasta):
    """{nome sem "TESTE - " e sem ".1": caminho} dos PDFs da pasta."""
    saida = {}
    try:
        nomes = sorted(os.listdir(pasta))
    except OSError:
        return saida
    for n in nomes:
        if not n.lower().endswith(".pdf"):
            continue
        chave = n[len(PREFIXO_TESTE):] if n.startswith(PREFIXO_TESTE) else n
        saida.setdefault(chave, os.path.join(pasta, n))
    return saida


def _resumo_diferenca(a, b):
    pa, pb = a.split(" "), b.split(" ")
    trechos = []
    for op, i1, i2, j1, j2 in difflib.SequenceMatcher(None, pa, pb, autojunk=False).get_opcodes():
        if op != "equal":
            trechos.append(f'AHK "{" ".join(pa[i1:i2])}" → Python "{" ".join(pb[j1:j2])}"')
    return "; ".join(trechos[:3]) + (" …" if len(trechos) > 3 else "")


def comparar_pastas(pasta_ahk, pasta_python, ignorar=()):
    """Compara as duas pastas. "ignorar" = textos que podem ser diferentes de
    propósito (ex.: a data de emissão), trocados por "…" antes de comparar."""
    ahk, py = _pdfs(pasta_ahk), _pdfs(pasta_python)
    linhas = []
    for nome in sorted(set(ahk) | set(py), key=str.casefold):
        if nome not in py:
            linhas.append(Linha(nome, SO_AHK))
            continue
        if nome not in ahk:
            linhas.append(Linha(nome, SO_PYTHON))
            continue
        try:
            ta, tp = _texto_pdf(ahk[nome]), _texto_pdf(py[nome])
        except Exception as e:
            linhas.append(Linha(nome, DIFERENTE, f"não deu para ler o PDF: {e}"))
            continue
        for t in ignorar:
            if t:
                ta, tp = ta.replace(t, "…"), tp.replace(t, "…")
        if ta == tp:
            linhas.append(Linha(nome, IGUAL))
        else:
            linhas.append(Linha(nome, DIFERENTE, _resumo_diferenca(ta, tp)))
    return linhas


def datas_de_emissao(pasta):
    """Datas dd/mm/aaaa que aparecem em todos os PDFs da pasta (a de emissão
    costuma ser a única repetida em todos)."""
    comuns = None
    for caminho in list(_pdfs(pasta).values())[:5]:
        try:
            datas = set(re.findall(r"\d{2}/\d{2}/\d{4}", _texto_pdf(caminho)))
        except Exception:
            continue
        comuns = datas if comuns is None else comuns & datas
    return sorted(comuns or [])
