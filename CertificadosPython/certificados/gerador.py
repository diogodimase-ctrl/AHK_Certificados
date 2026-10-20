"""
Geração do lote: junta planilhas, coordenadores, modelo, PDF e manifesto.

Três modos (no AHK eram arquivos separados que precisavam andar espelhados):
  oficial   = GerarCertificados_v27: C:\\Certificados + "Para Enviar" (o fluxo envia).
  teste     = GerarCertificados_TESTE_v7: "TESTE - " no nome, C:\\Certificados_TESTE,
              "Para Enviar\\TESTE - <Curso>", e-mails trocados pelo e-mail de teste.
  simulacao = só no computador (C:\\Certificados_SIMULACAO), nada vai para a nuvem
              e nada é enviado. Serve para comparar com o que o AHK gerou.
"""

import os
import re
import shutil
import tempfile
import traceback
from dataclasses import dataclass, field
from datetime import datetime

from . import caminhos, manifesto, modelo, planilhas, textos

PREFIXO_TESTE = "TESTE - "


@dataclass
class Modo:
    nome: str
    pasta_local: str
    prefixo: str = ""
    copia_por_docente: bool = False
    usa_para_enviar: bool = True
    destino_prefixo: str = ""

    @property
    def rotulo(self):
        return {"oficial": "Oficial", "teste": "Teste", "simulacao": "Simulação"}[self.nome]


MODOS = {
    "oficial": Modo("oficial", r"C:\Certificados", copia_por_docente=True),
    "teste": Modo("teste", r"C:\Certificados_TESTE", prefixo=PREFIXO_TESTE, destino_prefixo="TESTE/"),
    "simulacao": Modo("simulacao", r"C:\Certificados_SIMULACAO", usa_para_enviar=False),
}


class Log:
    """RegistrarLog: um arquivo por rodada em C:\\Certificados\\Logs."""

    def __init__(self, modo_nome, pasta=r"C:\Certificados\Logs", ecoar=None):
        os.makedirs(pasta, exist_ok=True)
        sufixo = "" if modo_nome == "oficial" else f"_{modo_nome.upper()}"
        self.arquivo = os.path.join(pasta, f"geracao_python{sufixo}_{datetime.now():%Y-%m-%d_%H-%M-%S}.txt")
        self.ecoar = ecoar

    def __call__(self, msg):
        linha = f"[{datetime.now():%H:%M:%S}] {msg}"
        try:
            with open(self.arquivo, "a", encoding="utf-8") as f:
                f.write(linha + "\n")
        except OSError:
            pass
        if self.ecoar:
            self.ecoar(msg)


@dataclass
class CursoNoLote:
    curso: planilhas.Curso
    coordenacao: object = None              # coordenadores.Coordenacao
    planilha: str = ""
    abas: list = field(default_factory=list)
    dados: object = None                    # planilhas.DadosCurso
    removido: str = ""

    @property
    def nome_com_unidade(self):
        base = textos.sanitizar_nome_arquivo(self.curso.posgrad)
        if self.curso.unidade:
            return f"{base} ({textos.sanitizar_nome_arquivo(self.curso.unidade)})"
        return base


@dataclass
class Problema:
    curso: CursoNoLote
    registro: planilhas.Registro

    def descricao(self):
        r = self.registro
        itens = []
        if r.erro_nome:
            itens.append("Nome em branco")
        if r.erro_curso:
            itens.append("Curso/disciplina em branco")
        if r.erro_data:
            itens.append("Data ausente ou inválida")
        if r.erro_horas:
            itens.append(f"Horário não reconhecido (lido: {r.horario_txt})")
        if r.suspeita_nome_email:
            itens.append("Aviso: coluna de nome parece conter e-mail")
        if r.suspeita_curso_data:
            itens.append("Aviso: coluna de curso parece conter data/hora")
        if r.suspeita_nome_com_extra:
            itens.append(f'Aviso: nome limpo automaticamente — original "{r.nome_original_bruto}" → usado "{r.nome}"')
        return itens


def validar(cursos):
    """Fase 2 (ValidarRegistros): lista os registros com problema."""
    problemas, mescladas, ok = [], 0, 0
    for c in cursos:
        if c.removido or c.dados is None:
            continue
        for aba in c.dados.abas:
            for r in aba.registros:
                mescladas += r.mesclado
                if r.tem_problema:
                    problemas.append(Problema(c, r))
                else:
                    ok += 1
    return problemas, mescladas, ok


@dataclass
class ResultadoCurso:
    curso: CursoNoLote
    gerados: int = 0
    ignorados: int = 0
    falhou: str = ""
    pastas: list = field(default_factory=list)


@dataclass
class Relatorio:
    modo: Modo
    data_emissao: str
    log: str
    cursos: list = field(default_factory=list)
    falhas: list = field(default_factory=list)
    avisos: list = field(default_factory=list)

    @property
    def total_gerados(self):
        return sum(r.gerados for r in self.cursos)


def _pasta_mes_para_enviar(locais, modo, nome_curso, mes):
    if not modo.usa_para_enviar:
        return ""
    caminho = os.path.join(locais.declaracoes, "Para Enviar", modo.prefixo + nome_curso, mes)
    os.makedirs(caminho, exist_ok=True)
    return caminho


def gerar(cursos, locais, modelo_pptx, data_emissao, modo, conversor, log,
          email_teste="", progresso=None, cancelado=lambda: False):
    """Fase 3 (GerarCertificadosCurso para cada curso)."""
    if modo.usa_para_enviar and not (locais.declaracoes and os.path.isdir(locais.declaracoes)):
        raise RuntimeError(f"A pasta \"{caminhos.NOME_DECLARACOES}\" não foi encontrada no OneDrive. "
                           "Sem ela os certificados não chegam ao fluxo de envio.")
    rel = Relatorio(modo, data_emissao, getattr(log, "arquivo", ""))
    ativos = [c for c in cursos if not c.removido and c.dados is not None]
    total = sum(1 for c in ativos for a in c.dados.abas for r in a.registros if not r.pular and not r.tem_erro)
    feitos = 0
    temp = tempfile.mkdtemp(prefix="certificados_")
    try:
        for c in ativos:
            if cancelado():
                break
            res = ResultadoCurso(c)
            rel.cursos.append(res)
            log(f"── Gerando curso: {c.curso.posgrad} ({c.curso.unidade})")
            try:
                avisos = []
                base = modelo.criar_modelo_base(
                    modelo_pptx, c.coordenacao.assina_como, c.coordenacao.cargo, data_emissao,
                    c.coordenacao.assinatura, locais.logo, avisos)
                for a in avisos:
                    log("AVISO: " + a)
                    if a not in rel.avisos:
                        rel.avisos.append(a)
                for aba in c.dados.abas:
                    feitos = _gerar_aba(c, aba, base, locais, modo, conversor, log, email_teste,
                                        res, rel, temp, progresso, feitos, total, cancelado)
            except Exception as e:
                res.falhou = str(e)
                log(f"ERRO no curso {c.curso.posgrad}: {e}\n{traceback.format_exc()}")
            log(f"Curso concluído: {res.gerados} gerados, {res.ignorados} ignorados/com falha.")
    finally:
        shutil.rmtree(temp, ignore_errors=True)
    log(f"Geração encerrada. PDFs gerados: {rel.total_gerados} | Certificados com falha: {len(rel.falhas)}")
    return rel


def _gerar_aba(c, aba, base, locais, modo, conversor, log, email_teste, res, rel, temp,
               progresso, feitos, total, cancelado):
    mes = textos.mes_para_pasta(aba.aba)
    pasta_local = os.path.join(modo.pasta_local, c.nome_com_unidade, mes)
    os.makedirs(pasta_local, exist_ok=True)
    pasta_docentes = os.path.join(modo.pasta_local, c.nome_com_unidade, "Docentes") if modo.copia_por_docente else ""
    pasta_envio = _pasta_mes_para_enviar(locais, modo, c.nome_com_unidade, mes)
    res.pastas.append(pasta_envio or pasta_local)
    log(f"Aba {aba.aba} → local: {pasta_local} | Para Enviar: {pasta_envio or '(simulação: não envia)'}")

    docentes = {}            # chave -> {nome, email, email_valido, arquivos}
    arquivos_do_lote = []
    falhas_seguidas = 0
    for r in aba.registros:
        if cancelado():
            break
        if r.pular:
            continue
        if r.tem_erro:
            res.ignorados += 1
            continue
        feitos += 1
        if progresso:
            progresso(f"{c.curso.posgrad} · {aba.aba} · {r.nome}", feitos, total)

        data_limpa = re.sub(r'[\\/:*?"<>|]', ".", r.data_aula)
        nome_base = f"{modo.prefixo}{textos.sanitizar_nome_arquivo(r.nome)} ({data_limpa})"
        try:
            pptx = os.path.join(temp, f"cert_{feitos}.pptx")
            modelo.preencher_certificado(base, r.nome, c.dados.nome_pos, r.curso, r.data_aula, r.horas_formatadas, pptx)
            pdf_local = caminhos.caminho_unico(pasta_local, nome_base, ".pdf")
            conversor.converter(pptx, pdf_local)
            falhas_seguidas = 0
        except Exception as e:
            falhas_seguidas += 1
            msg = f"{c.curso.posgrad} / {aba.aba} / linha {r.linha} — {r.nome}: {e}"
            log("ERRO certificado: " + msg)
            rel.falhas.append(msg)
            res.ignorados += 1
            if falhas_seguidas >= 3:
                raise RuntimeError(f"O PDF falhou em 3 certificados seguidos. Curso interrompido na aba '{aba.aba}'. "
                                   f"Último erro: {e}") from e
            continue

        if pasta_docentes:
            destino = os.path.join(pasta_docentes, textos.sanitizar_nome_arquivo(r.nome))
            try:
                os.makedirs(destino, exist_ok=True)
                shutil.copy2(pdf_local, caminhos.caminho_unico(destino, nome_base, ".pdf"))
            except OSError as e:
                log(f"Cópia por docente falhou ({nome_base}): {e}")

        arquivo_lote = pdf_local
        if pasta_envio:
            try:
                arquivo_lote = caminhos.caminho_unico(pasta_envio, nome_base, ".pdf")
                shutil.copy2(pdf_local, arquivo_lote)
            except OSError as e:
                log(f"AVISO: {nome_base} não foi copiado para a Para Enviar: {e}")
                rel.falhas.append(f"{c.curso.posgrad} / {aba.aba} / linha {r.linha} — {r.nome}: "
                                  f"gerado só no backup local, não foi para a Para Enviar")
                res.gerados += 1
                continue
        res.gerados += 1
        nome_arquivo = os.path.basename(arquivo_lote)
        arquivos_do_lote.append(nome_arquivo)

        chave = textos.sanitizar_nome_arquivo(r.nome).lower()
        doc = docentes.get(chave)
        if doc is None:
            doc = docentes[chave] = {"nome": r.nome, "email": r.email, "email_valido": r.email_valido, "arquivos": []}
        elif not doc["email_valido"] and r.email_valido:
            doc["email"], doc["email_valido"] = r.email, True
        doc["arquivos"].append({"nomeArquivo": nome_arquivo, "chaveLinha": f"{r.data_aula}|{r.horario_txt}"})

    _gravar_manifesto(c, aba, mes, pasta_envio or pasta_local, docentes, arquivos_do_lote, modo, email_teste, log, rel)
    return feitos


def _gravar_manifesto(c, aba, mes, pasta, docentes, arquivos_do_lote, modo, email_teste, log, rel):
    """GerarManifestoDoMes."""
    if not arquivos_do_lote:
        return
    validos = [d for d in docentes.values() if d["email_valido"]]
    if modo.nome == "teste":
        validos = [dict(d, email=email_teste) for d in validos]
    local = c.curso.sigla_unidade or c.curso.unidade
    semestre = textos.semestre_do_mes(mes)
    pasta_pos = textos.sanitizar_nome_arquivo(c.curso.posgrad)
    pasta_local_sigla = textos.sanitizar_nome_arquivo(local) if local.strip() else ""
    destino = ""
    if semestre and pasta_local_sigla:
        destino = f"{modo.destino_prefixo}{semestre}/{pasta_local_sigla}/{pasta_pos}/{mes}"
    else:
        log(f"AVISO: sem pasta de destino para {c.curso.posgrad} / {aba.aba} (semestre='{semestre}', local='{local}').")
    texto = manifesto.montar(c.curso.posgrad, local, os.path.join(pasta, ""), c.planilha, aba.aba, validos,
                             semestre, mes, pasta_pos, destino, arquivos_do_lote)
    try:
        manifesto.gravar(os.path.join(pasta, "manifesto.txt"), texto)
        log(f"Manifesto gravado ({len(validos)} docentes com e-mail, {len(arquivos_do_lote)} PDFs, destino '{destino}').")
    except OSError as e:
        log(f"ERRO ao gravar o manifesto em {pasta}: {e}")
        rel.falhas.append(f"{c.curso.posgrad} / {aba.aba}: manifesto.txt não foi gravado ({e})")
