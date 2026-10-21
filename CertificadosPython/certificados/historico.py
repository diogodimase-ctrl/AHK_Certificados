"""
Controle dos cursos já processados.

Junta duas fontes:
  - o histórico que o programa grava a cada rodada OFICIAL, num arquivo
    compartilhado na pasta Declarações (vale para qualquer computador e para
    qualquer colega que use o programa);
  - as pastas que o fluxo do Power Automate já arquivou
    (Declarações\\<Semestre>\\<Local>\\<Pós>\\<MM.AA>), o que inclui os envios
    feitos pelo AHK v27.

Cada mês de cada curso fica "ok" (arquivado pelo fluxo, ou gerado e já fora da
Para Enviar) ou "pendente" (gerado, mas a pasta ainda está na Para Enviar).
"""

import getpass
import json
import os
import re
from dataclasses import dataclass
from datetime import datetime

from . import textos

PASTA_HISTORICO = "_Programa Certificados"
ARQUIVO_HISTORICO = "historico.jsonl"
_MESES_CURTOS = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def caminho_historico(locais):
    return os.path.join(locais.declaracoes, PASTA_HISTORICO, ARQUIVO_HISTORICO)


def rotulo_mes(pasta_mes):
    """"10.26" -> "Out26"."""
    m = re.match(r"^(\d{2})\.(\d{2})$", pasta_mes or "")
    if not m or not 1 <= int(m.group(1)) <= 12:
        return pasta_mes
    return f"{_MESES_CURTOS[int(m.group(1)) - 1]}{m.group(2)}"


def _ordem(pasta_mes):
    m = re.match(r"^(\d{2})\.(\d{2})$", pasta_mes or "")
    return (int(m.group(2)), int(m.group(1))) if m else (0, 0)


def chave_curso(sigla, posgrad):
    return textos.normalizar(sigla), textos.normalizar(textos.sanitizar_nome_arquivo(posgrad))


# ---------------------------------------------------------------------------
# Gravar
# ---------------------------------------------------------------------------

def registrar(locais, relatorio, log=print):
    """Depois de uma rodada oficial: uma linha por curso/mês gerado."""
    linhas = []
    agora = datetime.now().isoformat(timespec="seconds")
    try:
        usuario = getpass.getuser()
    except Exception:
        usuario = ""
    for res in relatorio.cursos:
        if res.falhou or not res.gerados:
            continue
        c = res.curso
        for aba in c.abas:
            pasta_mes = textos.mes_para_pasta(aba)
            linhas.append({
                "quando": agora, "usuario": usuario, "curso": c.curso.posgrad, "unidade": c.curso.unidade,
                "sigla": c.curso.sigla_unidade or c.curso.unidade, "aba": aba, "pastaMes": pasta_mes,
                "semestre": textos.semestre_do_mes(pasta_mes), "pastaEnvio": c.nome_com_unidade,
                "pdfsNoCurso": res.gerados,
            })
    if not linhas:
        return True
    arquivo = caminho_historico(locais)
    try:
        os.makedirs(os.path.dirname(arquivo), exist_ok=True)
        with open(arquivo, "a", encoding="utf-8") as f:
            for linha in linhas:
                f.write(json.dumps(linha, ensure_ascii=False) + "\n")
        log(f"Histórico: {len(linhas)} curso/mês registrado(s) em {arquivo}")
        return True
    except OSError as e:
        log(f"AVISO: o histórico não foi gravado em {arquivo}: {e}")
        return False


# ---------------------------------------------------------------------------
# Ler
# ---------------------------------------------------------------------------

def ler_historico(locais):
    arquivo = caminho_historico(locais)
    itens = []
    try:
        with open(arquivo, encoding="utf-8") as f:
            for linha in f:
                linha = linha.strip()
                if linha:
                    try:
                        itens.append(json.loads(linha))
                    except ValueError:
                        continue
    except OSError:
        pass
    return itens


def _listar(pasta):
    try:
        return os.listdir(pasta)
    except OSError:
        return []


def arquivados(locais):
    """{chave_curso: {pastaMes: datetime}} das pastas que o fluxo arquivou."""
    achados = {}
    raiz = locais.declaracoes
    for semestre in _listar(raiz):
        if not re.match(r"^\d{4}-[12]$", semestre):
            continue
        for sigla in _listar(os.path.join(raiz, semestre)):
            for pos in _listar(os.path.join(raiz, semestre, sigla)):
                for mes in _listar(os.path.join(raiz, semestre, sigla, pos)):
                    pasta = os.path.join(raiz, semestre, sigla, pos, mes)
                    if not re.match(r"^\d{2}\.\d{2}$", mes):
                        continue
                    if not any(n.lower().endswith(".pdf") for n in _listar(pasta)):
                        continue
                    try:
                        quando = datetime.fromtimestamp(os.path.getmtime(pasta))
                    except OSError:
                        quando = None
                    achados.setdefault((textos.normalizar(sigla), textos.normalizar(pos)), {})[mes] = quando
    return achados


@dataclass
class MesProcessado:
    pasta_mes: str
    rotulo: str
    data: str          # "dd/mm" ou ""
    pendente: bool     # gerado, mas ainda na Para Enviar


def situacao(locais, cursos):
    """{curso.linha: [MesProcessado, ...]}, do mês mais antigo ao mais novo."""
    from .gerador import CursoNoLote
    arquivo = arquivados(locais)
    por_curso = {}
    for h in ler_historico(locais):
        chave = chave_curso(h.get("sigla", ""), h.get("curso", ""))
        por_curso.setdefault(chave, []).append(h)
    para_enviar = os.path.join(locais.declaracoes, "Para Enviar")
    saida = {}
    for c in cursos:
        chave = chave_curso(c.sigla_unidade or c.unidade, c.posgrad)
        meses = {}
        for mes, quando in arquivo.get(chave, {}).items():
            meses[mes] = MesProcessado(mes, rotulo_mes(mes), quando.strftime("%d/%m") if quando else "", False)
        envio = CursoNoLote(c).nome_com_unidade
        for h in por_curso.get(chave, []):
            mes = h.get("pastaMes", "")
            if not mes:
                continue
            try:
                data = datetime.fromisoformat(h["quando"]).strftime("%d/%m")
            except (KeyError, ValueError):
                data = ""
            pasta = os.path.join(para_enviar, h.get("pastaEnvio") or envio, mes)
            pendente = any(n.lower().endswith(".pdf") for n in _listar(pasta))
            if mes not in meses or pendente:
                meses[mes] = MesProcessado(mes, rotulo_mes(mes), data, pendente)
            elif data:
                meses[mes].data = data           # a data do envio vale mais que a da pasta
        saida[c.linha] = sorted(meses.values(), key=lambda m: _ordem(m.pasta_mes))
    return saida


def meses_conhecidos(situacoes):
    """Pastas de mês que aparecem em algum curso, da mais nova para a mais antiga."""
    todos = {m.pasta_mes for lista in situacoes.values() for m in lista}
    return sorted(todos, key=_ordem, reverse=True)
