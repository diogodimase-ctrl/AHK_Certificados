"""
Regras de texto copiadas do AHK (GerarCertificados_v27 + UI_Moderna).

Cada função aqui tem o nome da função AHK equivalente no comentário, para
dar para conferir lado a lado. A ideia é que o Python dê EXATAMENTE o mesmo
resultado que o AHK; melhorias de regra entram depois, uma de cada vez.
"""

import re

# NormalizarTexto (UI_Moderna) troca só estes acentos.
_ACENTOS_BASICOS = str.maketrans({
    "á": "a", "à": "a", "ã": "a", "â": "a",
    "é": "e", "ê": "e", "í": "i",
    "ó": "o", "ô": "o", "õ": "o",
    "ú": "u", "ü": "u", "ç": "c",
})

# ChaveCoordenador troca mais estes, além dos de cima.
_ACENTOS_EXTRAS = str.maketrans({
    "è": "e", "ë": "e", "ì": "i", "î": "i", "ï": "i",
    "ò": "o", "ö": "o", "ù": "u", "û": "u", "ñ": "n", "ä": "a",
})

_PARTICULAS = {"de", "da", "do", "dos", "das", "e", "di", "du"}
_TITULOS_COORD = re.compile(r"^(dr|dra|prof|profa|professor|professora) ")
_TITULOS_DOCENTE = re.compile(r"^(dra|dr|profa|professora|professor|prof|sra|sr)\.?\s+", re.I)


def normalizar(txt):
    """NormalizarTexto: minúsculo, sem acento, espaços simples."""
    t = str(txt if txt is not None else "").strip().lower()
    t = re.sub(r"[ \t]", " ", t)
    t = re.sub(r"\s+", " ", t)
    return t.translate(_ACENTOS_BASICOS)


def chave_nome(nome):
    """ChaveCoordenador: "Dra. Márcia  Souza" -> "marcia souza"."""
    t = normalizar(nome).translate(_ACENTOS_EXTRAS)
    t = re.sub(r"[^a-z0-9 ]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    while _TITULOS_COORD.match(t):
        t = re.sub(r"^\S+ ", "", t, count=1)
    return t


def palavras_do_nome(chave):
    """PalavrasDoNome: palavras sem as partículas (de, da, do...)."""
    return [p for p in chave.split(" ") if p and p not in _PARTICULAS]


def todas_contidas(menor, maior):
    """TodasPalavrasContidas: todas as palavras de "menor" (2+) estão em "maior"."""
    if len(menor) < 2:
        return False
    return all(p in maior for p in menor)


def sanitizar_nome_arquivo(nome):
    """SanitizarNomeArquivo: tira o que o Windows não aceita em nome de pasta."""
    limpo = re.sub(r'[\\/:*?"<>|]', "", str(nome))
    limpo = re.sub(r"[\x00-\x1f]", "", limpo)
    limpo = limpo.strip(" \t\r\n.")
    return limpo or "Sem_Nome"


def limpar_nome_docente(bruto):
    """Limpeza do nome do docente feita em LerRegistro/CorrigirNomeRegistro.

    Devolve (nome_limpo, nome_original_bruto). O "bruto" é o texto já sem o
    H/M solto no fim (marcação de sexo usada em algumas planilhas)."""
    nome = str(bruto).strip()
    nome = re.sub(r"\s*\b[HMhm]\b\s*$", "", nome).strip()
    original = nome
    nome = re.sub(r"\s*\([^)]*\)", "", nome)
    nome = _TITULOS_DOCENTE.sub("", nome)
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome, original


def separar_coordenadores(texto):
    """ExtrairCoordenadores: "A, B / C" -> ["A", "B", "C"]."""
    return [c.strip() for c in re.split(r"[,/\\]", str(texto or "")) if c.strip()]


def juntar_nomes(lista):
    """JuntarNomes: "A", "A ou B", "A, B ou C"."""
    if len(lista) <= 1:
        return "".join(lista)
    return ", ".join(lista[:-1]) + " ou " + lista[-1]


# ---------------------------------------------------------------------------
# Títulos e cargos de coordenador
# ---------------------------------------------------------------------------

def normalizar_titulo(txt):
    """NormalizarTitulo: "dra"/"Dra." -> "Dra." ; "dr" -> "Dr." ; outro -> ""."""
    t = str(txt or "").strip().lower()
    if t.startswith("dra"):
        return "Dra."
    if t.startswith("dr"):
        return "Dr."
    return ""


def cargo_por_titulo(titulo):
    """CargoPorTitulo."""
    return "Coordenadora do Curso" if titulo == "Dra." else "Coordenador do Curso"


# ---------------------------------------------------------------------------
# Meses e semestres
# ---------------------------------------------------------------------------

_MESES = {
    "janeiro": "01", "jan": "01", "fevereiro": "02", "fev": "02",
    "marco": "03", "mar": "03", "abril": "04", "abr": "04",
    "maio": "05", "mai": "05", "junho": "06", "jun": "06",
    "julho": "07", "jul": "07", "agosto": "08", "ago": "08",
    "setembro": "09", "set": "09", "outubro": "10", "out": "10",
    "novembro": "11", "nov": "11", "dezembro": "12", "dez": "12",
}


def mes_para_pasta(nome_aba):
    """ConverterMesParaPasta: "Abril26" / "Abril/2026" -> "04.26"."""
    nome = str(nome_aba).strip()
    m = re.match(r"^([A-Za-zçÇãÃéÉêÊóÓ]+)\D*(\d{2,4})\s*$", nome)
    if not m:
        return sanitizar_nome_arquivo(nome)
    mes = m.group(1).lower()
    for de, para in (("ç", "c"), ("ã", "a"), ("é", "e"), ("ê", "e"), ("ó", "o")):
        mes = mes.replace(de, para)
    if mes not in _MESES:
        return sanitizar_nome_arquivo(nome)
    ano = m.group(2)
    return _MESES[mes] + "." + (ano[2:4] if len(ano) == 4 else ano)


def semestre_do_mes(pasta_mes):
    """SemestreDoMes: "04.26" -> "2026-1"; "10.26" -> "2026-2"; "02.27" -> "2026-2"."""
    m = re.match(r"^(\d{2})\.(\d{2})$", pasta_mes or "")
    if not m:
        return ""
    mes, ano = int(m.group(1)), 2000 + int(m.group(2))
    if not 1 <= mes <= 12:
        return ""
    if 3 <= mes <= 8:
        return f"{ano}-1"
    if mes >= 9:
        return f"{ano}-2"
    return f"{ano - 1}-2"


# ---------------------------------------------------------------------------
# Unidade e semestre (busca da planilha do curso)
# ---------------------------------------------------------------------------

_SIGLAS = "PTA|CEP|RJ|GOI|VM|MOR|IBI|ALPH|BH|DF|SJC"


def sigla_unidade(codigo):
    """ExtrairSiglaUnidade: "2026-A-PTA-1S" -> "PTA"."""
    txt = str(codigo or "").strip()
    if not txt:
        return ""
    m = re.search(rf"-({_SIGLAS}|[A-Z]{{2,5}})(?:-|$)", txt, re.I)
    if m:
        return m.group(1).upper()
    if len(txt) <= 5 and re.fullmatch(r"[A-Za-z]+", txt):
        return txt.upper()
    return txt


_VARIACOES_UNIDADE = {
    "PTA": ["paulista", "pta", "bela vista"],
    "CEP": ["cerqueira cesar", "cep", "paulista"],
    "RJ": ["rio de janeiro", "rj", "rio"],
    "GOI": ["goiania", "goi", "goias"],
    "VM": ["vila mariana", "vm"],
    "MOR": ["morumbi", "mor"],
    "IBI": ["ibirapuera", "ibi"],
    "ALPH": ["alphaville", "alph", "alpha"],
    "BH": ["belo horizonte", "bh"],
    "DF": ["brasilia", "df"],
    "SJC": ["sao jose dos campos", "sjc", "sao jose"],
}


def variacoes_unidade(unidade):
    """ObterVariacoesUnidade."""
    sigla = sigla_unidade(unidade)
    if not sigla:
        return []
    variacoes = []
    for v in [normalizar(sigla)] + _VARIACOES_UNIDADE.get(sigla.upper(), []):
        v = normalizar(v)
        if v and v not in variacoes:
            variacoes.append(v)
    return variacoes


def variacoes_semestre(semestre):
    """ObterEquivalenciasSemestre: "2026-5" também vale como "2026-1s", "26-1"..."""
    s = str(semestre or "").strip().lower().replace(" ", "").replace("_", "-")
    if not s:
        return []
    variacoes = [s]
    m = re.search(r"(20\d{2}|\d{2})[-.]?([1256])(?:s)?", s)
    if m:
        ano4 = m.group(1) if len(m.group(1)) == 4 else "20" + m.group(1)
        ano2 = ano4[2:4]
        n = "1" if m.group(2) in ("1", "5") else "2"
        alt = "5" if n == "1" else "6"
        variacoes += [f"{ano4}-{n}s", f"{ano4}-{n}", f"{ano4}-{alt}", f"{ano2}-{n}s", f"{ano2}-{n}",
                      f"{ano2}-{alt}", f"{ano2}-{alt}s", f"{ano4}.{n}", f"{ano2}.{n}", f"{n}s{ano4}", f"{n}s{ano2}"]
    return variacoes


# ---------------------------------------------------------------------------
# E-mails
# ---------------------------------------------------------------------------

_PADRAO_EMAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


def email_valido(email):
    """EmailEhValido."""
    email = str(email or "").strip()
    return bool(email) and re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", email) is not None


def extrair_email(bruto):
    """ExtrairEmailValido: o primeiro e-mail do texto; se não houver, o texto."""
    bruto = str(bruto or "").strip()
    m = _PADRAO_EMAIL.search(bruto)
    return m.group(0) if m else bruto


def todos_os_emails(bruto):
    """ExtrairTodosEmails."""
    return _PADRAO_EMAIL.findall(str(bruto or ""))


def email_do_docente(emails, nome):
    """EmailDoDocente: o único e-mail que contém uma palavra (3+ letras) do nome."""
    palavras = palavras_do_nome(chave_nome(nome))
    achados = []
    for e in emails:
        usuario = chave_nome(e.split("@", 1)[0].replace(".", " ").replace("_", " "))
        if any(len(p) >= 3 and p in usuario for p in palavras):
            achados.append(e)
    return achados[0] if len(achados) == 1 else ""
