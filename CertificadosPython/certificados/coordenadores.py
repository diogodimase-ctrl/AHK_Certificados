"""
Coordenadores: cadastro (coordenadores.txt), busca da assinatura e decisão de
quem assina cada curso. Mesmas regras da v25 do AHK.

O cadastro é o MESMO arquivo que o AHK usa, na pasta de assinaturas da nuvem:
    Nome como está na planilha mestre | Dr. ou Dra. | arquivo da assinatura
"""

import os
import re
from dataclasses import dataclass, field

from . import textos

ARQUIVO_CADASTRO = "coordenadores.txt"
_EXTENSOES = (".jpg", ".jpeg", ".png")

_CABECALHO = (
    "# Cadastro de coordenadores usado pelo GerarCertificados (v25 em diante).\r\n"
    "# Uma linha por coordenador:  Nome como está na planilha mestre | Dr. ou Dra. | arquivo da assinatura\r\n"
    "# O arquivo da assinatura é opcional: sem ele, o script procura pelo nome na pasta.\r\n"
    "# Pode corrigir à mão no Bloco de Notas. Linhas começando com # são ignoradas.\r\n\r\n"
)


class Cadastro:
    """CarregarCadastroCoordenadores / SalvarCadastroCoordenadores."""

    def __init__(self, pasta_assinaturas, log=print):
        self.pasta = pasta_assinaturas
        self.arquivo = os.path.join(pasta_assinaturas, ARQUIVO_CADASTRO)
        self.log = log
        self.itens = {}          # chave -> {"nome", "titulo", "assinatura"}
        self.erro_ao_salvar = ""
        self._carregar()

    def _carregar(self):
        if not os.path.isfile(self.arquivo):
            self.log(f"Cadastro de coordenadores ainda não existe: {self.arquivo}")
            return
        try:
            with open(self.arquivo, encoding="utf-8-sig") as f:
                linhas = f.read().splitlines()
        except OSError as e:
            self.log(f"Falha ao ler o cadastro {self.arquivo}: {e}")
            return
        for linha in linhas:
            linha = linha.strip()
            if not linha or linha.startswith("#"):
                continue
            partes = [p.strip(" \t") for p in linha.split("|")]
            chave = textos.chave_nome(partes[0])
            if not chave:
                continue
            self.itens[chave] = {
                "nome": partes[0],
                "titulo": textos.normalizar_titulo(partes[1]) if len(partes) >= 2 else "",
                "assinatura": partes[2] if len(partes) >= 3 else "",
            }
        self.log(f"Cadastro de coordenadores: {len(self.itens)} nome(s) lido(s) de {self.arquivo}")

    def salvar(self):
        linhas = sorted((f"{d['nome']} | {d['titulo']} | {d['assinatura']}" for d in self.itens.values()),
                        key=str.casefold)
        texto = _CABECALHO + "\r\n".join(linhas) + "\r\n"
        try:
            with open(self.arquivo, "w", encoding="utf-8-sig", newline="") as f:
                f.write(texto)
            self.erro_ao_salvar = ""
            return True
        except OSError as e:
            self.erro_ao_salvar = str(e)
            self.log(f"Falha ao salvar o cadastro em {self.arquivo}: {e}")
            return False

    def obter(self, nome):
        return self.itens.get(textos.chave_nome(nome))

    def definir(self, nome, campo, valor):
        chave = textos.chave_nome(nome)
        item = self.itens.setdefault(chave, {"nome": nome, "titulo": "", "assinatura": ""})
        item[campo] = valor
        return self.salvar()

    def caminho_para_cadastro(self, caminho):
        """Dentro da pasta de assinaturas guarda só o nome do arquivo."""
        pasta = os.path.normcase(os.path.join(self.pasta, ""))
        if os.path.normcase(caminho).startswith(pasta):
            return caminho[len(pasta):]
        return caminho

    def assinatura(self, nome):
        """AssinaturaDoCadastro: caminho da assinatura relacionada à mão, se ainda existir."""
        item = self.obter(nome)
        if not item or not item["assinatura"]:
            return ""
        v = item["assinatura"]
        caminho = v if re.match(r"^([A-Za-z]:|\\\\)", v) else os.path.join(self.pasta, v)
        if os.path.isfile(caminho):
            return caminho
        self.log(f"Assinatura do cadastro não existe mais: {caminho} ({nome})")
        return ""


def encontrar_assinatura(pasta, nome_coord, log=print):
    """EncontrarAssinatura: busca a imagem pelo nome, do critério mais forte ao mais fraco.

    1) mesmo nome (sem acento, maiúscula, pontuação e título);
    2) mesmo primeiro e último nome;
    3) todas as palavras de um estão no outro.
    Só aceita quando o critério aponta UM arquivo; ambíguo devolve ""."""
    alvo = textos.chave_nome(nome_coord)
    if not alvo:
        return ""
    palavras_alvo = textos.palavras_do_nome(alvo)
    iguais, primeiro_ultimo, contidos = [], [], []
    try:
        arquivos = sorted(os.listdir(pasta), key=str.casefold)
    except OSError as e:
        log(f"Falha ao listar a pasta de assinaturas {pasta}: {e}")
        return ""
    for arq in arquivos:
        base, ext = os.path.splitext(arq)
        caminho = os.path.join(pasta, arq)
        if ext.lower() not in _EXTENSOES or not os.path.isfile(caminho):
            continue
        chave = textos.chave_nome(base)
        if chave == alvo:
            iguais.append(caminho)
            continue
        palavras = textos.palavras_do_nome(chave)
        if (len(palavras_alvo) >= 2 and len(palavras) >= 2
                and palavras_alvo[0] == palavras[0] and palavras_alvo[-1] == palavras[-1]):
            primeiro_ultimo.append(caminho)
        elif textos.todas_contidas(palavras_alvo, palavras) or textos.todas_contidas(palavras, palavras_alvo):
            contidos.append(caminho)
    if iguais:
        return iguais[0]
    for grupo in (primeiro_ultimo, contidos):
        if len(grupo) == 1:
            return grupo[0]
        if len(grupo) > 1:
            nomes = ", ".join(os.path.basename(c) for c in grupo)
            log(f"Assinatura ambígua para '{nome_coord}' ({nomes}) — será perguntado.")
            return ""
    return ""


@dataclass
class Coordenacao:
    """Quem assina um curso. Se faltar algo, "pendente" diz o que perguntar."""
    nomes_da_planilha: list
    nome: str = ""
    titulo: str = ""
    assinatura: str = ""
    origem: str = ""
    pendente: str = ""        # "", "assinatura" ou "titulo"
    removido: str = ""        # motivo, se o curso saiu do lote

    @property
    def cargo(self):
        return textos.cargo_por_titulo(self.titulo)

    @property
    def assina_como(self):
        return f"{self.titulo} {self.nome}".strip()


def resolver(coordenacao_bruta, cadastro, log=print):
    """ResolverCoordenador, sem perguntar nada: tenta cada coordenador da linha,
    na ordem, e fica com o primeiro que tem assinatura (cadastro, depois pelo
    nome). O que não der para descobrir fica em "pendente" para a tela perguntar."""
    nomes = textos.separar_coordenadores(coordenacao_bruta)
    c = Coordenacao(nomes)
    for nome in nomes:
        caminho = cadastro.assinatura(nome)
        if caminho:
            c.nome, c.assinatura, c.origem = nome, caminho, "cadastro"
            break
        caminho = encontrar_assinatura(cadastro.pasta, nome, log)
        if caminho:
            c.nome, c.assinatura, c.origem = nome, caminho, "pelo nome"
            break
    if not c.assinatura:
        c.pendente = "assinatura"
        return c
    completar_titulo(c, cadastro)
    return c


def completar_titulo(c, cadastro):
    item = cadastro.obter(c.nome)
    c.titulo = item["titulo"] if item else ""
    c.pendente = "" if c.titulo else "titulo"


def definir_assinatura(c, cadastro, caminho, nome=""):
    """Assinatura escolhida à mão. "nome" é obrigatório quando a linha tem
    mais de um coordenador (EscolherCoordenadorUI)."""
    if not nome:
        if len(c.nomes_da_planilha) == 1:
            nome = c.nomes_da_planilha[0]
        elif not c.nomes_da_planilha:
            base = os.path.splitext(os.path.basename(caminho))[0]
            nome = re.sub(r"^\s*(dra?|profa?)\.?\s+", "", base, flags=re.I).strip()
        else:
            raise ValueError("Escolha qual coordenador é o dono da assinatura.")
    c.nome, c.assinatura, c.origem = nome, caminho, "escolhida à mão"
    cadastro.definir(nome, "assinatura", cadastro.caminho_para_cadastro(caminho))
    completar_titulo(c, cadastro)


def definir_titulo(coordenacoes, cadastro, nome, titulo):
    """Dr./Dra. de um coordenador; vale para todos os cursos dele."""
    titulo = textos.normalizar_titulo(titulo)
    chave = textos.chave_nome(nome)
    for c in coordenacoes:
        if c.nome and textos.chave_nome(c.nome) == chave:
            c.titulo = titulo
            if c.pendente == "titulo":
                c.pendente = ""
    cadastro.definir(nome, "titulo", titulo)


def trocar_assinatura(coordenacoes, cadastro, nome, caminho):
    """Troca a assinatura de um coordenador em todos os cursos dele."""
    chave = textos.chave_nome(nome)
    for c in coordenacoes:
        if c.nome and textos.chave_nome(c.nome) == chave:
            c.assinatura, c.origem = caminho, "escolhida à mão"
    cadastro.definir(nome, "assinatura", cadastro.caminho_para_cadastro(caminho))
