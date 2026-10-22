"""
Números da tela inicial: meses do semestre e certificados previstos por mês.

Os certificados previstos de um curso no mês saem da planilha do curso (as
linhas da aba daquele mês). Abrir dezenas de planilhas do OneDrive demora, então
a contagem fica guardada em %APPDATA%\\CertificadosPython\\contagem.json e só é
refeita quando a planilha muda.
"""

import json
import os
import re
from datetime import date

from . import caminhos, planilhas

ARQUIVO_CONTAGEM = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "CertificadosPython",
                                "contagem.json")
MESES_CURTOS = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MESES_LONGOS = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro",
                "Outubro", "Novembro", "Dezembro"]


def pasta_mes(ano, mes):
    return f"{mes:02d}.{ano % 100:02d}"


def _ano_mes(pasta):
    m = re.match(r"^(\d{2})\.(\d{2})$", pasta or "")
    return (2000 + int(m.group(2)), int(m.group(1))) if m else (0, 0)


def curto(pasta):
    """"10.26" -> "Out"."""
    ano, mes = _ano_mes(pasta)
    return MESES_CURTOS[mes - 1] if mes else pasta


def longo(pasta):
    """"10.26" -> "Outubro/26"."""
    ano, mes = _ano_mes(pasta)
    return f"{MESES_LONGOS[mes - 1]}/{ano % 100:02d}" if mes else pasta


def meses_do_semestre(aba, hoje=None):
    """Aba da mestre "2026.2" -> ["09.26", ..., "02.27"] (mesma regra de
    textos.semestre_do_mes: 1º semestre = março a agosto). Aba sem ano e
    semestre: o semestre de hoje."""
    hoje = hoje or date.today()
    m = re.search(r"(20\d{2}|\d{2})\s*[-.]?\s*([1256])", str(aba or ""))
    if m:
        ano = int(m.group(1)) if len(m.group(1)) == 4 else 2000 + int(m.group(1))
        segundo = m.group(2) in ("2", "6")
    else:
        segundo = not 3 <= hoje.month <= 8
        ano = hoje.year - (1 if hoje.month <= 2 else 0)
    inicio = 9 if segundo else 3
    return [pasta_mes(ano + (inicio + i - 1) // 12, (inicio + i - 1) % 12 + 1) for i in range(6)]


def mes_padrao(meses, hoje=None):
    """O mês de hoje, se for do semestre; senão o último que já passou (ou o primeiro)."""
    hoje = hoje or date.today()
    atual = pasta_mes(hoje.year, hoje.month)
    if atual in meses:
        return atual
    passados = [m for m in meses if _ano_mes(m) < (hoje.year, hoje.month)]
    return passados[-1] if passados else meses[0]


def ja_passou(pasta, hoje=None):
    """O mês é o atual ou já passou."""
    hoje = hoje or date.today()
    return _ano_mes(pasta) <= (hoje.year, hoje.month)


# ---------------------------------------------------------------------------
# Certificados previstos
# ---------------------------------------------------------------------------

def _ler_cache(arquivo):
    try:
        with open(arquivo, encoding="utf-8") as f:
            dados = json.load(f)
        return dados if isinstance(dados, dict) else {}
    except (OSError, ValueError):
        return {}


def _gravar_cache(arquivo, dados):
    try:
        os.makedirs(os.path.dirname(arquivo), exist_ok=True)
        with open(arquivo, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False)
    except OSError:
        pass


def contar_previstos(locais, cursos, aba, pasta, progresso=None, arquivo=ARQUIVO_CONTAGEM):
    """{curso.linha: nº de certificados previstos no mês}. Curso sem planilha
    ou com planilha que não abre fica de fora (None)."""
    cache = _ler_cache(arquivo)
    saida = {}
    for i, c in enumerate(cursos):
        if progresso:
            progresso(i, len(cursos))
        caminho = caminhos.localizar_planilha_do_curso(locais.raiz_onedrive, c.posgrad, c.unidade, aba)
        if not caminho:
            saida[c.linha] = None
            continue
        try:
            chave = f"{os.path.normcase(caminho)}|{os.path.getmtime(caminho):.0f}|{pasta}"
        except OSError:
            saida[c.linha] = None
            continue
        if chave in cache:
            saida[c.linha] = cache[chave]
            continue
        try:
            n = planilhas.contar_certificados_do_mes(caminho, pasta)
        except Exception:
            saida[c.linha] = None
            continue
        cache[chave] = saida[c.linha] = n
    _gravar_cache(arquivo, cache)
    return saida
