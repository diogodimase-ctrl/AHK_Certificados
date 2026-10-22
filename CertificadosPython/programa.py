"""
Certificados — o programa novo (Etapa 2), no lugar do GerarCertificados do AHK.

Dois modos, escolhidos na tela de Início:
  Teste    igual à TESTE_v7: os PDFs ganham "TESTE - " no nome, ficam em
           C:\\Certificados_TESTE e em "Para Enviar\\TESTE - <Curso>\\<Mês>", e no
           manifesto o e-mail de todos os docentes vira o e-mail de teste.
  Oficial  igual à v27: C:\\Certificados (com a cópia por docente) e
           "Para Enviar\\<Curso>\\<Mês>"; o fluxo manda os e-mails para os docentes.

Uma janela só, com as etapas numa faixa no alto (os mesmos passos do AHK):
  1. Início           painel do mês (cursos e certificados feitos), cursos da
                      planilha mestre e data de emissão (Passos 0–3). Modo, e-mail
                      de teste e arquivos ficam no selo "Arquivos e pastas" do alto.
  2. Coordenadores    quem assina cada curso (Passo 4)
  3. Planilhas        planilha de cada curso e meses (Passos 5–7)
  4. Conferir dados   linhas com problema: corrigir, gerar assim ou pular (Fase 2)
  5. Gerar            PDFs, cópia para a Para Enviar e manifesto (Fase 3)
  6. Relatório        (Passo 10)
"""

import ctypes
import json
import os
import shutil
import sys
import time
import webbrowser
from datetime import datetime
from tkinter import filedialog

import customtkinter as ctk

from certificados import (VERSAO, caminhos, coordenadores, economia, gerador, historico, horas, painel, planilhas,
                          textos)
from certificados.pdf import ConversorPowerPoint, powerpoint_aberto
import atualizacao
import tema
from componentes import BarrasTempo, Etapas, PlanilhaCursos, Rosca
from conferencia import em_segundo_plano, linha_de_arquivo
from tema import FONTE, tabela

COR_OK, COR_ERRO, COR_AVISO, COR_APAGADO = tema.OK, tema.ERRO, tema.AVISO, tema.APAGADO

ETAPAS = ["Início", "Coordenadores", "Planilhas e meses", "Conferir dados", "Gerar", "Relatório"]
GERAR, RELATORIO = 4, 5
CONFIG = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "CertificadosPython", "config.json")


def ler_config():
    try:
        with open(CONFIG, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def gravar_config(dados):
    try:
        os.makedirs(os.path.dirname(CONFIG), exist_ok=True)
        with open(CONFIG, "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Diálogos
# ---------------------------------------------------------------------------

class Dialogo(ctk.CTkToplevel):
    """Janela de pergunta: texto + botões. "escolha" fica com o valor do botão
    clicado (None se fechar no X)."""

    def __init__(self, pai, titulo, texto, botoes, largura=560):
        super().__init__(pai)
        self.title(titulo)
        self.escolha = None
        self.transient(pai)
        self.resizable(False, False)
        ctk.CTkLabel(self, text=titulo, font=(FONTE, 18, "bold"), anchor="w").pack(fill="x", padx=20, pady=(16, 6))
        ctk.CTkLabel(self, text=texto, justify="left", anchor="w", wraplength=largura - 40).pack(fill="x", padx=20)
        self.corpo = ctk.CTkFrame(self, fg_color="transparent")
        self.corpo.pack(fill="x", padx=20, pady=(8, 0))
        linha = ctk.CTkFrame(self, fg_color="transparent")
        linha.pack(fill="x", padx=20, pady=16)
        for i, (rotulo, valor) in enumerate(reversed(botoes)):
            principal = i == 0
            ctk.CTkButton(linha, text=rotulo, width=160, height=40, **({} if principal else tema.BOTAO_CONTORNO),
                          command=lambda v=valor: self.fechar(v)).pack(side="right", padx=(8, 0))
        self.protocol("WM_DELETE_WINDOW", lambda: self.fechar(None))
        self.after(150, self._mostrar)

    def _mostrar(self):
        self.lift()
        self.focus_force()
        try:
            self.grab_set()
        except Exception:
            pass

    def fechar(self, valor):
        self.escolha = valor
        self.grab_release()
        self.destroy()


def perguntar(pai, titulo, texto, botoes, largura=560):
    d = Dialogo(pai, titulo, texto, botoes, largura)
    pai.wait_window(d)
    return d.escolha


def avisar(pai, titulo, texto):
    perguntar(pai, titulo, texto, [("OK", True)])


class DialogoCorrecao(Dialogo):
    """Corrigir uma linha da planilha (CorrigirNomeRegistro, CorrigirDataRegistro...)."""

    def __init__(self, pai, problema):
        r = problema.registro
        texto = (f"Curso: {problema.curso.curso.posgrad}\nMês: {r.aba} · linha {r.linha} da planilha\n\n"
                 "A correção vale só para esta rodada. A planilha não é alterada.")
        super().__init__(pai, f"Corrigir a linha {r.linha}", texto, [("Cancelar", None), ("Aplicar", "aplicar")])
        self.campos = {}
        if r.erro_nome or r.suspeita_nome_email or r.suspeita_nome_com_extra:
            self._campo("nome", "Nome do docente", r.nome, "Ex.: Maria da Silva Santos")
        if r.erro_curso or r.suspeita_curso_data:
            self._campo("curso", "Disciplina / aula", r.curso, "Ex.: Gestão Estratégica de Pessoas")
        if r.erro_data:
            self._campo("data", "Data da aula", r.data_aula, "dd/mm/aaaa, ex.: 15/05/2026")
        if r.erro_horas:
            self._campo("horas", f"Duração (lido: {r.horario_txt or 'vazio'})", "", "4  ou  2:30")

    def _campo(self, chave, rotulo, valor, dica):
        ctk.CTkLabel(self.corpo, text=rotulo, anchor="w").pack(fill="x", pady=(6, 0))
        var = ctk.StringVar(value=valor)
        ctk.CTkEntry(self.corpo, textvariable=var, width=500).pack(fill="x")
        ctk.CTkLabel(self.corpo, text=dica, anchor="w", text_color=COR_APAGADO, font=(FONTE, 12)).pack(fill="x")
        self.campos[chave] = var

    def fechar(self, valor):
        self.valores = {k: v.get().strip() for k, v in self.campos.items()}
        super().fechar(valor)


# ---------------------------------------------------------------------------
# Janela
# ---------------------------------------------------------------------------

class App(ctk.CTk):
    def __init__(self, modo="teste"):
        tema.aplicar()
        try:      # barra de tarefas com o ícone do programa (o mesmo id do atalho criado pelo instalador)
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("Certificados.PosGraduacao")
        except (AttributeError, OSError):
            pass
        super().__init__()
        icone = os.path.join(getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__))), "icone.ico")
        if os.path.isfile(icone):
            self.after(300, lambda: self.iconbitmap(icone))     # o CustomTkinter põe o dele aos 200 ms
        self.modo = gerador.MODOS[modo]
        self.geometry("1280x800")
        self.minsize(1100, 700)
        self.after(0, lambda: self.state("zoomed"))
        # CERTIFICADOS_RAIZ aponta para outra "raiz do OneDrive" (usado nos testes).
        self.locais = caminhos.Locais.descobrir(os.environ.get("CERTIFICADOS_RAIZ") or None)
        self.config_usuario = ler_config()
        self.email_teste = self.config_usuario.get("email_teste", "")
        self.mestre = self.locais.planilha_mestre
        self.modelo = self.locais.modelo
        self.previstos = {}                      # (aba, mês) -> {linha da mestre: certificados previstos}
        self._contando = None
        self._estilo_tabelas()
        self._montar_janela()
        self.nova_rodada()

    def _estilo_tabelas(self):
        tema.pintar_tabelas(self)

    def _montar_janela(self):
        self.faixa = ctk.CTkLabel(self, text="", font=(FONTE, 13, "bold"), height=30)
        self.faixa.pack(fill="x")

        topo = ctk.CTkFrame(self, corner_radius=0, fg_color=tema.LATERAL)
        topo.pack(fill="x")
        linha = ctk.CTkFrame(topo, fg_color="transparent")
        linha.pack(fill="x", padx=32, pady=(12, 4))
        ctk.CTkLabel(linha, text="Certificados · Pós-graduação", font=(FONTE, 20, "bold"),
                     text_color=tema.TITULO).pack(side="left")
        ctk.CTkLabel(linha, text=f"versão {VERSAO}", text_color=COR_APAGADO, font=(FONTE, 12)).pack(
            side="left", padx=14)
        self.seletor_modo = ctk.CTkSegmentedButton(linha, values=["Teste", "Oficial"], command=self._trocar_modo,
                                                   width=200, height=34)
        self.seletor_modo.set(self.modo.rotulo)
        self.seletor_modo.pack(side="right")
        self.chip_arquivos = ctk.CTkButton(linha, text="", height=32, corner_radius=16, font=(FONTE, 13, "bold"),
                                           command=self.abrir_configuracoes)
        self.chip_arquivos.pack(side="right", padx=12)
        self.botao_atualizar = ctk.CTkButton(linha, text="", height=32, fg_color=tema.OK_FUNDO,
                                             hover_color="#1f4d40", text_color=tema.OK, font=(FONTE, 13, "bold"))
        # Só depois que a janela abriu (a resposta volta para a janela, que precisa estar rodando).
        self.after(2000, lambda: em_segundo_plano(self, lambda: atualizacao.versao_mais_nova(VERSAO),
                                                  self._versao_nova, lambda e, d: None))
        self.etapas = Etapas(topo, ETAPAS)
        self.etapas.pack(fill="x", padx=24)
        linha = ctk.CTkFrame(topo, fg_color="transparent")
        linha.pack(fill="x", padx=32, pady=(0, 12))
        self.rotulo_etapa = ctk.CTkLabel(linha, text="", text_color=tema.TEXTO_2, font=(FONTE, 13))
        self.rotulo_etapa.pack(side="right", padx=(12, 0))
        self.barra_etapas = ctk.CTkProgressBar(linha, height=6)
        self.barra_etapas.pack(side="left", fill="x", expand=True)

        rodape = ctk.CTkFrame(self, fg_color="transparent")       # antes do conteúdo: em tela baixa não some
        rodape.pack(side="bottom", fill="x", padx=32, pady=12)
        self.conteudo = ctk.CTkFrame(self, fg_color="transparent")
        self.conteudo.pack(fill="both", expand=True, padx=32, pady=(14, 0))
        self.botao_voltar = ctk.CTkButton(rodape, text="← Voltar", width=120, **tema.BOTAO_CONTORNO,
                                           command=self.voltar)
        self.botao_voltar.pack(side="left")
        self.botao_avancar = ctk.CTkButton(rodape, text="Avançar →", width=260, height=44,
                                           font=(FONTE, 16, "bold"), command=self.avancar)
        self.botao_avancar.pack(side="right")
        self.status = ctk.CTkLabel(rodape, text="", anchor="w", text_color=COR_APAGADO)
        self.status.pack(side="left", fill="x", expand=True, padx=16)
        self._aplicar_modo()

    def _versao_nova(self, resultado):
        if not resultado:
            return
        versao, link = resultado
        self.botao_atualizar.configure(text=f"Versão {versao} disponível · baixar",
                                       command=lambda: webbrowser.open(link))
        self.botao_atualizar.pack(side="right", padx=(0, 4))

    def _aplicar_modo(self):
        """Título da janela e faixa do alto: amarela no teste, vermelha no oficial."""
        self.title(f"Certificados — modo {self.modo.rotulo} (versão {VERSAO})")
        if self.modo.nome == "teste":
            destino = self.email_teste or "o e-mail de teste (falta preencher)"
            self.faixa.configure(text=f"MODO TESTE  ·  nada vai para os docentes: todos os e-mails vão para "
                                      f"{destino}", fg_color=tema.FAIXA_TESTE, text_color=tema.FAIXA_TESTE_TEXTO)
        else:
            self.faixa.configure(text="MODO OFICIAL  ·  os e-mails vão para os docentes de verdade",
                                 fg_color=tema.FAIXA_OFICIAL, text_color=tema.FAIXA_OFICIAL_TEXTO)
        self._atualizar_chip()

    @property
    def oficial(self):
        return self.modo.nome == "oficial"

    def _trocar_modo(self, rotulo):
        novo = "oficial" if rotulo == "Oficial" else "teste"
        if novo == self.modo.nome:
            return
        if self.etapa != 0:                    # só se troca de modo na tela inicial
            self.seletor_modo.set(self.modo.rotulo)
            return
        self.modo = gerador.MODOS[novo]
        self._aplicar_modo()
        self.ir_para(0)

    # -- arquivos e configurações ----------------------------------------------

    def _itens_verificacao(self):
        """(rótulo, caminho, ok, obrigatório). Igual à VerificacaoPreRodagem do AHK."""
        l = self.locais
        para_enviar = os.path.join(l.declaracoes, "Para Enviar") if l.declaracoes else ""
        return [
            ("Pasta Declarações (OneDrive)", l.declaracoes or "OneDrive da empresa não encontrado",
             bool(l.declaracoes) and os.path.isdir(l.declaracoes), True),
            ("Pasta Para Enviar", para_enviar, bool(para_enviar) and os.path.isdir(para_enviar), True),
            ("Modelo do certificado", self.modelo, os.path.isfile(self.modelo or ""), True),
            ("Logo", l.logo, os.path.isfile(l.logo), True),
            ("Pasta de assinaturas", l.assinaturas, os.path.isdir(l.assinaturas), True),
            ("Planilha mestre", self.mestre or "—", os.path.isfile(self.mestre or ""), False),
        ]

    def _faltando(self):
        faltam = [f"• {rot}: {cam}" for rot, cam, ok, obrig in self._itens_verificacao() if obrig and not ok]
        if not self.oficial and not textos.email_valido(self.email_teste):
            faltam.append("• E-mail de teste")
        return faltam

    def _atualizar_chip(self):
        if self._faltando():
            self.chip_arquivos.configure(text="✗  Falta configurar · abrir", fg_color=tema.ERRO_FUNDO,
                                         hover_color=tema.SECUNDARIO_HOVER, text_color=tema.ERRO)
        else:
            self.chip_arquivos.configure(text="✓  Arquivos e pastas ok", fg_color=tema.OK_FUNDO,
                                         hover_color="#1f4d40", text_color=tema.OK)

    def abrir_configuracoes(self):
        d = Dialogo(self, "Arquivos e configurações",
                    "O programa acha tudo sozinho no OneDrive. Só troque se algum item estiver em vermelho.",
                    [("Fechar", True)], largura=860)
        var_email = ctk.StringVar(value=self.email_teste)
        if not self.oficial:
            linha = ctk.CTkFrame(d.corpo, fg_color="transparent")
            linha.pack(fill="x", pady=3)
            ctk.CTkLabel(linha, text="E-mail de teste", width=150, anchor="w").pack(side="left")
            ctk.CTkEntry(linha, textvariable=var_email, placeholder_text="seu.nome@einstein.br").pack(
                side="left", fill="x", expand=True)
        quadro = ctk.CTkFrame(d.corpo)

        def conferir():
            self.modelo, mestre = var_modelo.get(), var_mestre.get()
            for w in quadro.winfo_children():
                w.destroy()
            for i, (rotulo, caminho, ok, obrigatorio) in enumerate(self._itens_verificacao()):
                cor = COR_OK if ok else (COR_ERRO if obrigatorio else COR_AVISO)
                ctk.CTkLabel(quadro, text="✓" if ok else ("✗" if obrigatorio else "!"), text_color=cor, width=30,
                             font=(FONTE, 16, "bold")).grid(row=i, column=0, padx=(10, 0), pady=3)
                ctk.CTkLabel(quadro, text=rotulo, anchor="w", width=210).grid(row=i, column=1, sticky="w")
                ctk.CTkLabel(quadro, text=caminho, anchor="w", text_color=COR_APAGADO, wraplength=520,
                             justify="left").grid(row=i, column=2, sticky="w", padx=(0, 10))
            if powerpoint_aberto():
                ctk.CTkLabel(quadro, anchor="w", text_color=COR_AVISO, text="! O PowerPoint está aberto. Salve e "
                             "feche as apresentações antes de gerar.").grid(row=99, column=0, columnspan=3,
                                                                            sticky="w", padx=10, pady=3)

        var_modelo = linha_de_arquivo(d.corpo, "Modelo (.pptx)", self.modelo, conferir, [("PowerPoint", "*.pptx")])
        var_mestre = linha_de_arquivo(d.corpo, "Planilha mestre", self.mestre, conferir, [("Excel", "*.xlsx")])
        ctk.CTkLabel(d.corpo, text="Arquivos e pastas", font=(FONTE, 16, "bold"), anchor="w").pack(fill="x",
                                                                                                pady=(12, 4))
        quadro.pack(fill="x")
        conferir()
        self.wait_window(d)
        self.modelo = var_modelo.get()
        mestre = var_mestre.get()
        if not self.oficial:
            email = var_email.get().strip()
            if email != self.email_teste:
                self.email_teste = email
                self.config_usuario["email_teste"] = email
                gravar_config(self.config_usuario)
        self._aplicar_modo()
        if mestre != self.mestre:
            self.mestre = mestre
            self.abas_mestre, self.aba, self.cursos, self.marcados, self.situacoes = [], "", [], set(), None
            if self.etapa == 0:
                self.ir_para(0)

    # -- estado e navegação ---------------------------------------------------

    def nova_rodada(self):
        self.abas_mestre, self.aba = [], ""
        self.data_emissao = ""                   # começa vazia: tem de ser digitada a cada rodada
        self.cursos, self.marcados = [], set()   # marcados = linhas da mestre
        self.situacoes = None                    # historico.situacao: meses já enviados por curso
        self.mes_painel = ""
        self.lote, self.cadastro = [], None
        self.abas_por_planilha, self.erros_planilha = {}, {}
        self.meses_marcados = set()
        self.problemas, self.mescladas, self.linhas_ok = [], 0, 0
        self.relatorio, self.log, self.modo_do_log, self.cancelar = None, None, "", False
        self.inicio_rodada = None                # relógio do tempo economizado
        self.ir_para(0)

    @property
    def ativos(self):
        return [i for i in self.lote if not i.removido]

    def ir_para(self, n):
        self.etapa, self._ocupado = n, False
        for w in self.conteudo.winfo_children():
            w.destroy()
        self.etapas.marcar(n)
        self.barra_etapas.set((n + 0.5) / len(ETAPAS))
        self.rotulo_etapa.configure(text=f"Etapa {n + 1} de {len(ETAPAS)}")
        self.seletor_modo.configure(state="normal" if n == 0 else "disabled")
        self.status.configure(text="", text_color=COR_APAGADO)
        self.botao_voltar.configure(state="normal" if 0 < n < GERAR else "disabled", text="← Voltar",
                                    command=self.voltar)
        self.botao_avancar.configure(state="normal", text="Avançar →")
        [self.montar_inicio, self.montar_coordenadores, self.montar_planilhas, self.montar_conferir,
         self.montar_gerar, self.montar_relatorio][n]()

    def voltar(self):
        if 0 < self.etapa < GERAR:
            self.ir_para(self.etapa - 1)

    def avancar(self):
        [self.avancar_inicio, self.avancar_coordenadores, self.avancar_planilhas, lambda: self.ir_para(GERAR),
         self.gerar, self.nova_rodada][self.etapa]()

    def ocupado(self, texto):
        self._ocupado = True
        self.botao_avancar.configure(state="disabled")
        self.botao_voltar.configure(state="disabled")
        self.status.configure(text=texto, text_color=COR_APAGADO)

    def livre(self, texto=""):
        self._ocupado = False
        self.botao_avancar.configure(state="normal")
        self.botao_voltar.configure(state="normal" if 0 < self.etapa < GERAR else "disabled")
        self.status.configure(text=texto)
        if self.etapa == 0:
            self._atualizar_continuar()

    def registrar(self, msg):
        if self.log:
            self.log(msg)

    def mostrar_erro(self, erro, detalhe=""):
        self.registrar(f"ERRO: {erro}\n{detalhe}")
        self.livre()
        janela = ctk.CTkToplevel(self)
        janela.title("Erro")
        janela.geometry("660x340")
        janela.transient(self)
        ctk.CTkLabel(janela, text="Algo deu errado", font=(FONTE, 18, "bold")).pack(anchor="w", padx=16, pady=(14, 4))
        caixa = ctk.CTkTextbox(janela, font=("Consolas", 11))
        caixa.pack(fill="both", expand=True, padx=16, pady=4)
        caixa.insert("end", f"{erro}\n\n{detalhe}")
        ctk.CTkButton(janela, text="Fechar", command=janela.destroy).pack(pady=10)

    def cabecalho(self, titulo, texto):
        ctk.CTkLabel(self.conteudo, text=titulo, font=(FONTE, 28, "bold"), text_color=tema.TITULO,
                     anchor="w").pack(fill="x")
        tema.quebrar_na_largura(ctk.CTkLabel(self.conteudo, text=texto, justify="left", anchor="w",
                     text_color=tema.TEXTO_2)).pack(fill="x", pady=(2, 12))

    # -- 1. Início: painel, cursos da mestre e data de emissão -------------------

    def montar_inicio(self):
        cartoes = ctk.CTkFrame(self.conteudo, fg_color="transparent")
        cartoes.pack(fill="x")
        for i in range(4):
            cartoes.grid_columnconfigure(i, weight=1, uniform="cartao")
        self.rosca_cursos = Rosca(cartoes, "Cursos feitos no mês", tema.OK)
        self.rosca_cursos.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self.rosca_certificados = Rosca(cartoes, "Certificados feitos no mês", tema.CERTIFICADOS)
        self.rosca_certificados.grid(row=0, column=1, sticky="nsew", padx=8)

        cartao_mes = ctk.CTkFrame(cartoes)
        cartao_mes.grid(row=0, column=2, sticky="nsew", padx=8)
        ctk.CTkLabel(cartao_mes, text="Mês dos certificados", text_color=tema.TEXTO_2, font=(FONTE, 13),
                     anchor="w").pack(fill="x", padx=16, pady=(16, 2))
        self.combo_mes = ctk.CTkComboBox(cartao_mes, values=[], command=self._trocar_mes, state="readonly")
        self.combo_mes.pack(fill="x", padx=16)
        self.rotulo_mestre = ctk.CTkLabel(cartao_mes, text="", text_color=COR_APAGADO, font=(FONTE, 12),
                                          anchor="w", justify="left", wraplength=260)
        self.rotulo_mestre.pack(fill="x", padx=16, pady=(8, 12))

        self.cartao_extra = ctk.CTkFrame(cartoes)
        self.cartao_extra.grid(row=0, column=3, sticky="nsew", padx=(8, 0))
        self.barras_tempo = BarrasTempo(self.cartao_extra)
        self.barras_tempo.mostrar(economia.resumo())
        self.quadro_data = ctk.CTkFrame(self.cartao_extra, fg_color="transparent")
        ctk.CTkLabel(self.quadro_data, text="Data de emissão", font=(FONTE, 15, "bold"), anchor="w").pack(fill="x")
        # Sem textvariable: com ela o CustomTkinter não mostra o texto de exemplo.
        self.campo_data = ctk.CTkEntry(self.quadro_data, placeholder_text="dd/mm/aaaa", height=40, font=(FONTE, 17))
        if self.data_emissao:
            self.campo_data.insert(0, self.data_emissao)
        self.campo_data.bind("<KeyRelease>", lambda e: self._data_mudou())
        self.campo_data.pack(fill="x", pady=(6, 4))
        self.rotulo_data = ctk.CTkLabel(self.quadro_data, text="", font=(FONTE, 12), anchor="w")
        self.rotulo_data.pack(fill="x")

        quadro = ctk.CTkFrame(self.conteudo)
        quadro.pack(fill="both", expand=True, pady=(16, 0))
        barra = ctk.CTkFrame(quadro, fg_color="transparent")
        barra.pack(fill="x", padx=16, pady=(10, 8))
        ctk.CTkLabel(barra, text="Cursos da planilha mestre", font=(FONTE, 16, "bold")).pack(side="left")
        self.rotulo_marcados = ctk.CTkLabel(barra, text="", text_color=tema.TEXTO_2)
        self.rotulo_marcados.pack(side="left", padx=16)
        self.campo_filtro = ctk.CTkEntry(barra, placeholder_text="procurar curso…", width=240)
        self.campo_filtro.bind("<KeyRelease>", lambda e: self._desenhar_cursos())
        self.campo_filtro.pack(side="right")
        self.var_esconder = ctk.BooleanVar(value=False)
        self.caixa_esconder = ctk.CTkCheckBox(barra, text="Esconder os já feitos no mês", variable=self.var_esconder,
                                              command=self._desenhar_cursos)
        self.caixa_esconder.pack(side="right", padx=16)
        self.tabela_cursos = PlanilhaCursos(quadro, self._marcar)
        self.tabela_cursos.pack(fill="both", expand=True)
        base = ctk.CTkFrame(quadro, fg_color=tema.LATERAL, corner_radius=0, height=38)
        base.pack(fill="x")
        self.quadro_abas = ctk.CTkFrame(base, fg_color="transparent")
        self.quadro_abas.pack(side="left", fill="y", padx=(8, 0))
        ctk.CTkLabel(base, text="Verde = enviado  ·  Amarelo = na Para Enviar  ·  Vermelho = ainda não feito  ·  "
                                "Cinza = mês futuro ou sem aulas", text_color=COR_APAGADO,
                     font=(FONTE, 12)).pack(side="right", padx=14)

        self._atualizar_chip()
        if self.abas_mestre:
            self._desenhar_abas()
            self._preencher_meses()
            self._desenhar_cursos()
            self._contar_previstos()
        else:
            self._carregar_abas()
        self._atualizar_continuar()

    def _na_inicio(self):
        return self.etapa == 0 and self.tabela_cursos.winfo_exists()

    # mestre e semestre

    def _carregar_abas(self):
        self.rotulo_mestre.configure(text=f"Planilha mestre: {os.path.basename(self.mestre or '') or '—'}")
        self.tabela_cursos.mostrar([], [], "", "Lendo a planilha mestre…")
        if not os.path.isfile(self.mestre or ""):
            self.tabela_cursos.mostrar([], [], "", "Planilha mestre não encontrada. Clique no aviso vermelho lá "
                                                   "em cima para escolher.")
            return
        self.ocupado("Lendo a planilha mestre…")
        em_segundo_plano(self, lambda: planilhas.listar_abas(self.mestre), self._abas_carregadas)

    def _abas_carregadas(self, abas):
        self.livre()
        self.abas_mestre = abas
        semestres = self._abas_de_semestre()
        if semestres and self._na_inicio():
            self._trocar_aba(semestres[-1])

    def _abas_de_semestre(self):
        return [a for a in self.abas_mestre if any(ch.isdigit() for ch in a)] or list(self.abas_mestre)

    def _desenhar_abas(self):
        for w in self.quadro_abas.winfo_children():
            w.destroy()
        ctk.CTkLabel(self.quadro_abas, text="Abas da mestre:", text_color=COR_APAGADO, font=(FONTE, 12)).pack(
            side="left", padx=(6, 8))
        for aba in self._abas_de_semestre():
            atual = aba == self.aba
            ctk.CTkButton(self.quadro_abas, text=aba, width=70, height=30, corner_radius=0,
                          font=(FONTE, 13, "bold" if atual else "normal"),
                          fg_color=tema.ACENTO if atual else "transparent", hover_color=tema.SECUNDARIO_HOVER,
                          text_color="#ffffff" if atual else tema.TEXTO_2,
                          command=lambda a=aba: self._trocar_aba(a)).pack(side="left", padx=1, pady=4)

    def _trocar_aba(self, aba):
        if aba == self.aba and self.cursos:
            return
        self.aba, self.cursos, self.marcados, self.situacoes, self.mes_painel = aba, [], set(), None, ""
        self._desenhar_abas()
        self._preencher_meses()
        self.tabela_cursos.mostrar([], self.meses_semestre, self.mes_painel, f"Lendo os cursos da aba {aba}…")
        self.ocupado(f"Lendo os cursos da aba {aba}…")
        em_segundo_plano(self, lambda: planilhas.ler_cursos_da_mestre(self.mestre, aba),
                         lambda cursos: self._cursos_carregados(aba, cursos))

    def _cursos_carregados(self, aba, cursos):
        self.livre()
        if aba != self.aba:
            return
        self.cursos = cursos or []
        if cursos is None:
            self.status.configure(text=f"A aba {aba} não foi encontrada.", text_color=COR_ERRO)
        if self._na_inicio():
            self._desenhar_cursos()
        if self.cursos:
            self.status.configure(text="Lendo o histórico de envios…")
            cursos, locais = self.cursos, self.locais
            em_segundo_plano(self, lambda: historico.situacao(locais, cursos),
                             lambda s: self._situacao_carregada(cursos, s),
                             lambda e, d: (self.registrar(f"Histórico não lido: {e}"),
                                           self.status.configure(text="Não consegui ler o histórico de envios.")))

    def _situacao_carregada(self, cursos, situacoes):
        if cursos is not self.cursos:          # trocou de semestre enquanto lia
            return
        self.situacoes = situacoes
        if self._na_inicio():
            self.status.configure(text="")
            self._desenhar_cursos()
            self._contar_previstos()

    # mês do painel

    @property
    def meses_semestre(self):
        return painel.meses_do_semestre(self.aba)

    def _preencher_meses(self):
        meses = self.meses_semestre
        if self.mes_painel not in meses:
            self.mes_painel = painel.mes_padrao(meses)
        self.combo_mes.configure(values=[painel.longo(m) for m in meses])
        self.combo_mes.set(painel.longo(self.mes_painel))
        self.caixa_esconder.configure(text=f"Esconder os já feitos em {painel.curto(self.mes_painel)}")
        self.rotulo_mestre.configure(text=f"Semestre {self.aba or '—'} · planilha mestre: "
                                          f"{os.path.basename(self.mestre or '') or '—'}")

    def _trocar_mes(self, rotulo):
        for m in self.meses_semestre:
            if painel.longo(m) == rotulo:
                self.mes_painel = m
        self._preencher_meses()
        self._desenhar_cursos()
        self._contar_previstos()

    def _contar_previstos(self):
        """Certificados previstos no mês, lidos das planilhas dos cursos (em segundo plano)."""
        chave = (self.aba, self.mes_painel)
        if not self.cursos or self.situacoes is None or chave in self.previstos or self._contando == chave:
            return
        self._contando = chave
        cursos, locais = self.cursos, self.locais

        def progresso(feitos, total):
            self.after(0, lambda: self._na_inicio() and (self.aba, self.mes_painel) == chave and
                       self.rosca_certificados.carregando(f"lendo as planilhas dos cursos… {feitos} de {total}"))

        def pronto(resultado):
            self._contando = None
            self.previstos[chave] = resultado
            if self._na_inicio() and (self.aba, self.mes_painel) == chave:
                self._desenhar_cursos()

        def falhou(erro, detalhe):
            self._contando = None
            self.registrar(f"Contagem de certificados não feita: {erro}")

        em_segundo_plano(self, lambda: painel.contar_previstos(locais, cursos, chave[0], chave[1], progresso),
                         pronto, falhou)

    # tabela e números

    def _mes_do_curso(self, curso, mes):
        for m in (self.situacoes or {}).get(curso.linha, []):
            if m.pasta_mes == mes:
                return m
        return None

    def _desenhar_cursos(self):
        if not self._na_inicio():
            return
        mes = self.mes_painel
        previstos = self.previstos.get((self.aba, mes))
        filtro = textos.normalizar(self.campo_filtro.get())
        linhas, escondidos = [], 0
        feitos = pendentes = sem_aulas = cert_feitos = cert_previstos = 0
        for c in self.cursos:
            atual = self._mes_do_curso(c, mes)
            previsto = previstos.get(c.linha) if previstos else None
            if atual:
                feitos += 1
                pendentes += atual.pendente
            elif previstos and previsto == 0:
                sem_aulas += 1
            if previstos:
                pdfs = atual.pdfs if atual else 0
                cert_feitos += pdfs
                cert_previstos += max(previsto or 0, pdfs)
            if filtro and filtro not in textos.normalizar(f"{c.posgrad} {c.unidade} {c.coordenacao}"):
                continue
            if self.var_esconder.get() and atual:
                escondidos += 1
                continue
            meses = []
            for m in self.meses_semestre:
                feito = self._mes_do_curso(c, m)
                if feito:
                    meses.append("pendente" if feito.pendente else "ok")
                elif m == mes and previstos and previsto == 0:
                    meses.append("futuro")
                else:
                    meses.append("falta" if painel.ja_passou(m) and self.situacoes is not None else "futuro")
            if self.situacoes is None or (previstos is None and not atual):
                cert, cor = ("…" if self.situacoes is None or self._contando else ""), tema.APAGADO
            elif previstos is not None and previsto is None:
                cert, cor = (f"{atual.pdfs} feitos" if atual else "sem planilha"), tema.APAGADO
            elif previstos is not None and previsto == 0 and not atual:
                cert, cor = "sem aulas", tema.APAGADO
            else:
                pdfs = atual.pdfs if atual else 0
                total = max(previsto or 0, pdfs)
                cert = f"{pdfs} de {total}" if previstos is not None else f"{pdfs} feitos"
                cor = tema.OK if atual and not atual.pendente else (tema.AVISO if atual else tema.TEXTO_2)
            linhas.append({"linha": c.linha, "marcado": c.linha in self.marcados, "pos": c.posgrad,
                           "coord": c.coordenacao, "uni": c.sigla_unidade or c.unidade or "—", "turma": c.turma,
                           "meses": meses, "cert": cert, "cert_cor": cor})
        vazio = "Nenhum curso com esse nome" if filtro else ("Todos os cursos já foram feitos neste mês"
                                                             if escondidos else "Nenhum curso nesta aba")
        if not self.cursos:
            vazio = "Lendo os cursos…" if self.aba else "Escolha a planilha mestre"
        self.tabela_cursos.mostrar(linhas, self.meses_semestre, mes, vazio)

        extra = f"  ·  {escondidos} escondido(s)" if escondidos else ""
        self.rotulo_marcados.configure(text=f"{len(self.cursos)} cursos  ·  {len(self.marcados)} marcado(s){extra}")
        nome_mes = painel.longo(mes).split("/")[0].lower()
        if self.situacoes is None:
            self.rosca_cursos.carregando("lendo o histórico de envios…")
            self.rosca_certificados.carregando("lendo o histórico de envios…")
        else:
            total = len(self.cursos) - sem_aulas
            partes = [f"{pendentes} aguardando o fluxo" if pendentes else "", f"{sem_aulas} sem aulas em {nome_mes}"
                      if sem_aulas else ""]
            self.rosca_cursos.mostrar(feitos, total, "\n".join(p for p in partes if p) or
                                      f"{max(total - feitos, 0)} faltando")
            if previstos is None:
                self.rosca_certificados.carregando("lendo as planilhas dos cursos…")
            else:
                self.rosca_certificados.mostrar(cert_feitos, cert_previstos,
                                                f"{max(cert_previstos - cert_feitos, 0)} faltando")
        self._atualizar_continuar()

    def _marcar(self, linha):
        (self.marcados.discard if linha in self.marcados else self.marcados.add)(linha)
        self._desenhar_cursos()

    # data de emissão e botão continuar

    def _data_mudou(self):
        self.data_emissao = self.campo_data.get().strip()
        self._atualizar_continuar()

    def _atualizar_continuar(self):
        if not (self.etapa == 0 and hasattr(self, "quadro_data") and self.quadro_data.winfo_exists()):
            return
        n = len(self.marcados)
        if n:
            if not self.quadro_data.winfo_ismapped():
                self.barras_tempo.pack_forget()
                self.quadro_data.pack(fill="both", expand=True, padx=16, pady=14)
                self.after(50, self.campo_data.focus_set)
        elif not self.barras_tempo.winfo_ismapped():
            self.quadro_data.pack_forget()
            self.barras_tempo.pack(fill="both", expand=True, padx=12, pady=10)
        data_ok = horas.data_valida(self.data_emissao)
        if not self.data_emissao:
            self.rotulo_data.configure(text="Obrigatória para continuar", text_color=COR_AVISO)
        elif not data_ok:
            self.rotulo_data.configure(text="Use o formato dd/mm/aaaa", text_color=COR_ERRO)
        else:
            self.rotulo_data.configure(text="✓ Data ok", text_color=COR_OK)
        if not n:
            texto = "Marque os cursos desta rodada"
        else:
            texto = f"Continuar com {n} curso{'s' if n > 1 else ''} →"
        pronto = n > 0 and data_ok and not getattr(self, "_ocupado", False)
        self.botao_avancar.configure(text=texto, state="normal" if pronto else "disabled")

    def avancar_inicio(self):
        escolhidos = [c for c in self.cursos if c.linha in self.marcados]
        if not escolhidos or not horas.data_valida(self.data_emissao):
            return
        faltando = self._faltando()
        if faltando:
            avisar(self, "Falta configurar", "Sem estes itens não dá para gerar:\n\n" + "\n".join(faltando) +
                   "\n\nClique no aviso vermelho lá em cima para corrigir.")
            return
        if self.oficial:
            pergunta = ("Modo oficial", "Nesta rodada os e-mails vão para os DOCENTES de verdade.\n\n"
                                        "Antes de gerar, o programa ainda mostra o resumo e pede confirmação.",
                        [("Voltar", None), ("Entendi, continuar", True)])
        else:
            pergunta = ("Confirmar e-mail de teste",
                        f"Todos os e-mails do manifesto irão para:\n\n{self.email_teste}\n\nEstá correto?",
                        [("Corrigir", None), ("Sim, continuar", True)])
        if perguntar(self, *pergunta) is not True:
            return
        data = self.data_emissao
        if self.inicio_rodada is None:
            self.inicio_rodada = time.monotonic()
        if self.log is None or self.modo_do_log != self.modo.nome:
            self.log, self.modo_do_log = gerador.Log(self.modo.nome), self.modo.nome
        self.registrar(f"MODO {self.modo.rotulo.upper()}"
                       + ("" if self.oficial else f" — e-mails do manifesto: {self.email_teste}")
                       + f" | Modelo: {self.modelo}")
        self.registrar(f"Planilha mestre: {self.mestre} | Semestre: {self.aba} | Data de emissão: {data} | "
                       f"Cursos: {len(escolhidos)}")
        anteriores = {i.curso.linha: i for i in self.lote}
        locais, aba, log = self.locais, self.aba, self.registrar

        def trabalho():
            cadastro = coordenadores.Cadastro(locais.assinaturas, log=log)
            lote = []
            for c in escolhidos:
                antigo = anteriores.get(c.linha)
                if antigo is not None and antigo.curso == c:      # voltou e avançou: mantém o que já foi feito
                    lote.append(antigo)
                    continue
                item = gerador.CursoNoLote(c)
                item.coordenacao = coordenadores.resolver(c.coordenacao, cadastro, log=log)
                item.planilha = caminhos.localizar_planilha_do_curso(locais.raiz_onedrive, c.posgrad, c.unidade, aba)
                lote.append(item)
            return cadastro, lote

        self.ocupado("Procurando coordenadores, assinaturas e planilhas…")
        em_segundo_plano(self, trabalho, self._lote_montado)

    def _lote_montado(self, resultado):
        self.cadastro, self.lote = resultado
        self.ir_para(1)

    # -- 3. Coordenadores -------------------------------------------------------

    def montar_coordenadores(self):
        self.cabecalho("Quem assina cada curso",
                       "Confira o coordenador, o Dr./Dra. e a assinatura. O que você corrigir aqui fica guardado no "
                       "cadastro (o mesmo do AHK) e vale para as próximas rodadas. Duplo clique numa linha "
                       "vermelha resolve o que falta.")
        quadro, self.arvore_coord = tabela(self.conteudo, [("curso", "Curso", 260), ("assina", "Assina como", 300),
                                                           ("arquivo", "Assinatura", 200),
                                                           ("situacao", "Situação", 200)], alturas=12)
        quadro.pack(fill="both", expand=True)
        self.arvore_coord.bind("<Double-1>", lambda e: self._resolver_linha())
        botoes = ctk.CTkFrame(self.conteudo, fg_color="transparent")
        botoes.pack(fill="x", pady=8)
        ctk.CTkButton(botoes, text="Escolher assinatura…", command=self._escolher_assinatura).pack(side="left")
        ctk.CTkButton(botoes, text="Dr. / Dra.", command=self._escolher_titulo).pack(side="left", padx=8)
        ctk.CTkButton(botoes, text="Tirar do lote / devolver", **tema.BOTAO_CONTORNO,
                      command=self._tirar_ou_devolver).pack(side="left")
        self._desenhar_coordenadores()

    def _situacao_coord(self, item):
        c = item.coordenacao
        if item.removido:
            return "fora do lote", "apagado"
        if c.pendente == "assinatura":
            return "falta a assinatura", "erro"
        if c.pendente == "titulo":
            return "falta Dr./Dra.", "erro"
        return f"ok ({c.origem})", "ok"

    def _desenhar_coordenadores(self, selecionar=None):
        a = self.arvore_coord
        a.delete(*a.get_children())
        for i, item in enumerate(self.lote):
            c = item.coordenacao
            if c.pendente == "assinatura":
                assina = " ou ".join(c.nomes_da_planilha) or "(coordenação vazia na planilha mestre)"
            else:
                assina = f"{c.assina_como} — {c.cargo}" if c.titulo else c.nome
            situacao, tag = self._situacao_coord(item)
            a.insert("", "end", iid=str(i), values=(item.curso.posgrad, assina,
                                                    os.path.basename(c.assinatura) or "—", situacao), tags=(tag,))
        if selecionar is not None:
            a.selection_set(str(selecionar))
        pendentes = sum(1 for i in self.ativos if i.coordenacao.pendente)
        self.status.configure(text=f"{pendentes} curso(s) com pendência" if pendentes else
                              f"{len(self.ativos)} curso(s) prontos", text_color=COR_ERRO if pendentes else COR_OK)

    def _linha_escolhida(self):
        sel = self.arvore_coord.selection()
        if not sel:
            avisar(self, "Escolha um curso", "Clique primeiro na linha do curso.")
            return None
        return int(sel[0])

    def _resolver_linha(self):
        n = self._linha_escolhida()
        if n is None:
            return
        pendente = self.lote[n].coordenacao.pendente
        if pendente == "assinatura":
            self._escolher_assinatura()
        elif pendente == "titulo":
            self._escolher_titulo()

    def _reresolver_pendentes(self):
        """Depois de uma escolha, outros cursos do mesmo coordenador podem se resolver sozinhos."""
        for item in self.lote:
            if item.coordenacao.pendente == "assinatura":
                item.coordenacao = coordenadores.resolver(item.curso.coordenacao, self.cadastro, log=self.registrar)
            elif item.coordenacao.pendente == "titulo":
                coordenadores.completar_titulo(item.coordenacao, self.cadastro)

    def _escolher_assinatura(self):
        n = self._linha_escolhida()
        if n is None:
            return
        item = self.lote[n]
        c = item.coordenacao
        nomes = " ou ".join(c.nomes_da_planilha) if c.pendente == "assinatura" else c.nome
        caminho = filedialog.askopenfilename(title=f"Assinatura de: {nomes}", initialdir=self.cadastro.pasta,
                                             filetypes=[("Imagens", "*.jpg *.jpeg *.png")])
        if not caminho:
            return
        caminho = os.path.normpath(caminho)
        if c.pendente == "assinatura":
            nome = ""
            if len(c.nomes_da_planilha) > 1:
                nome = perguntar(self, "De quem é esta assinatura?",
                                 f"A linha deste curso tem mais de um coordenador.\n\nArquivo: "
                                 f"{os.path.basename(caminho)}", [("Cancelar", None)] +
                                 [(x, x) for x in c.nomes_da_planilha])
                if not nome:
                    return
            coordenadores.definir_assinatura(c, self.cadastro, caminho, nome)
        else:
            coordenadores.trocar_assinatura([i.coordenacao for i in self.lote], self.cadastro, c.nome, caminho)
        self.registrar(f"Assinatura de {c.nome} definida à mão: {caminho}")
        self._reresolver_pendentes()
        self._avisar_cadastro()
        self._desenhar_coordenadores(n)
        if self.lote[n].coordenacao.pendente == "titulo":
            self._escolher_titulo()

    def _escolher_titulo(self):
        n = self._linha_escolhida()
        if n is None:
            return
        c = self.lote[n].coordenacao
        if c.pendente == "assinatura":
            avisar(self, "Falta a assinatura", "Escolha primeiro a assinatura deste curso.")
            return
        titulo = perguntar(self, "Dr. ou Dra.?",
                           f"Como {c.nome} assina?\n\nA resposta fica guardada e não será perguntada de novo.",
                           [("Cancelar", None), ("Dr.  (Coordenador)", "Dr."), ("Dra.  (Coordenadora)", "Dra.")])
        if not titulo:
            return
        coordenadores.definir_titulo([i.coordenacao for i in self.lote], self.cadastro, c.nome, titulo)
        self.registrar(f"Título de {c.nome}: {titulo}")
        self._reresolver_pendentes()
        self._avisar_cadastro()
        self._desenhar_coordenadores(n)

    def _avisar_cadastro(self):
        if self.cadastro.erro_ao_salvar:
            avisar(self, "Cadastro não foi salvo",
                   f"A escolha vale para esta rodada, mas não foi guardada no cadastro:\n\n"
                   f"{self.cadastro.arquivo}\n\n{self.cadastro.erro_ao_salvar}")

    def _tirar_ou_devolver(self):
        n = self._linha_escolhida()
        if n is None:
            return
        item = self.lote[n]
        item.removido = "" if item.removido else "tirado do lote por você"
        self.registrar(f"{item.curso.posgrad}: {'tirado do lote' if item.removido else 'devolvido ao lote'}")
        self._desenhar_coordenadores(n)

    def avancar_coordenadores(self):
        if not self.ativos:
            avisar(self, "Lote vazio", "Todos os cursos estão fora do lote.")
            return
        pendentes = [i.curso.posgrad for i in self.ativos if i.coordenacao.pendente]
        if pendentes:
            avisar(self, "Falta resolver", "Estes cursos ainda têm pendência (duplo clique na linha para resolver, "
                                           "ou tire do lote):\n\n" + "\n".join(f"• {p}" for p in pendentes))
            return
        for item in self.ativos:
            c = item.coordenacao
            self.registrar(f"Coordenador de {item.curso.posgrad}: {c.assina_como} | Assinatura ({c.origem}): "
                           f"{c.assinatura}")
        self.ocupado("Lendo os meses das planilhas…")
        em_segundo_plano(self, self._listar_abas_das_planilhas, lambda _: self.ir_para(2))

    def _listar_abas_das_planilhas(self):
        for item in self.ativos:
            if item.planilha and item.planilha not in self.abas_por_planilha:
                self._listar_abas(item.planilha)

    def _listar_abas(self, caminho):
        try:
            self.abas_por_planilha[caminho] = planilhas.listar_abas(caminho)
            self.erros_planilha.pop(caminho, None)
        except Exception as e:
            self.abas_por_planilha[caminho] = []
            self.erros_planilha[caminho] = str(e)
            self.registrar(f"Não consegui abrir {caminho}: {e}")

    # -- 4. Planilhas e meses ---------------------------------------------------

    def montar_planilhas(self):
        self.cabecalho("Planilhas e meses",
                       "O programa achou a planilha de cada curso no OneDrive. Se alguma estiver errada ou não foi "
                       "encontrada, dê duplo clique na linha e escolha o arquivo. Depois marque os meses.")
        quadro, self.arvore_plan = tabela(self.conteudo, [("curso", "Curso", 280), ("planilha", "Planilha", 420),
                                                          ("meses", "Meses na planilha", 200)], alturas=9)
        quadro.pack(fill="both", expand=True)
        self.arvore_plan.bind("<Double-1>", self._trocar_planilha)
        ctk.CTkLabel(self.conteudo, text="Meses desta rodada", font=(FONTE, 16, "bold"), anchor="w").pack(
            fill="x", pady=(12, 2))
        self.quadro_meses = ctk.CTkFrame(self.conteudo, fg_color="transparent")
        self.quadro_meses.pack(fill="x")
        self._desenhar_planilhas()

    def _desenhar_planilhas(self):
        a = self.arvore_plan
        a.delete(*a.get_children())
        todas = []
        for i, item in enumerate(self.lote):
            if item.removido and not getattr(item, "tirado_na_etapa4", False):
                continue
            if not item.planilha:
                a.insert("", "end", iid=str(i), values=(item.curso.posgrad, "NÃO ENCONTRADA — duplo clique", "—"),
                         tags=("erro",))
                continue
            abas = self.abas_por_planilha.get(item.planilha, [])
            erro = self.erros_planilha.get(item.planilha)
            for x in abas:
                if x not in todas:
                    todas.append(x)
            meses = f"não abriu: {erro}" if erro else ", ".join(abas)
            a.insert("", "end", iid=str(i), values=(item.curso.posgrad, os.path.basename(item.planilha), meses),
                     tags=("erro" if erro else "ok",))
        for w in self.quadro_meses.winfo_children():
            w.destroy()
        self.vars_meses = {}
        for n, aba in enumerate(todas):
            var = ctk.BooleanVar(value=aba in self.meses_marcados)
            ctk.CTkCheckBox(self.quadro_meses, text=aba, variable=var, width=130,
                            command=lambda m=aba, v=var: (self.meses_marcados.add if v.get()
                                                          else self.meses_marcados.discard)(m)).grid(
                row=n // 7, column=n % 7, sticky="w", pady=3)
            self.vars_meses[aba] = var

    def _trocar_planilha(self, evento):
        linha = self.arvore_plan.identify_row(evento.y)
        if not linha:
            return
        item = self.lote[int(linha)]
        inicial = os.path.dirname(item.planilha) if item.planilha else self.locais.raiz_onedrive
        c = filedialog.askopenfilename(title=f"Planilha de {item.curso.posgrad}", initialdir=inicial or None,
                                       filetypes=[("Excel", "*.xlsx")])
        if not c:
            return
        item.planilha = os.path.normpath(c)
        self.registrar(f"Planilha de {item.curso.posgrad} escolhida à mão: {item.planilha}")
        self.ocupado("Lendo os meses da planilha…")
        em_segundo_plano(self, lambda: self._listar_abas(item.planilha),
                         lambda _: (self.livre(), self._desenhar_planilhas()))

    def avancar_planilhas(self):
        meses = [m for m, v in self.vars_meses.items() if v.get()]
        if not meses:
            avisar(self, "Nenhum mês", "Marque pelo menos um mês.")
            return
        for item in self.lote:                     # voltou e avançou de novo: decide outra vez
            if getattr(item, "tirado_na_etapa4", False):
                item.removido, item.tirado_na_etapa4 = "", False
        fora = []
        for n, item in enumerate(self.lote):
            if item.removido:
                continue
            motivo = ""
            if not item.planilha:
                motivo = "planilha não encontrada"
            elif item.planilha in self.erros_planilha:
                motivo = "a planilha não abriu"
            else:
                item.abas = [m for m in meses if planilhas.aba_na_lista(self.abas_por_planilha[item.planilha], m)]
                if not item.abas:
                    motivo = "a planilha não tem nenhum dos meses marcados"
            if motivo:
                fora.append((n, item, motivo))
        if fora and len(fora) == len(self.ativos):
            avisar(self, "Nenhum curso pronto", "Nenhum curso tem planilha com os meses marcados.")
            return
        if fora:
            texto = "\n".join(f"• {item.curso.posgrad}: {motivo}" for _, item, motivo in fora)
            if perguntar(self, "Cursos que vão ficar de fora", texto + "\n\nContinuar sem eles?",
                         [("Voltar e corrigir", None), ("Continuar sem eles", True)]) is not True:
                return
            for _, item, motivo in fora:
                item.removido, item.tirado_na_etapa4 = motivo, True
                self.registrar(f"{item.curso.posgrad} fora do lote: {motivo}")
        self.registrar("Meses: " + ", ".join(meses))
        lote = self.ativos

        def trabalho():
            for item in lote:
                try:
                    item.dados = planilhas.ler_dados_da_planilha(item.planilha, item.abas)
                    if not item.dados.nome_pos:
                        raise ValueError(f"a célula A2 da aba '{item.abas[0]}' está vazia")
                    self.registrar(f"Planilha de {item.curso.posgrad}: {item.planilha} | Abas: {', '.join(item.abas)}")
                except Exception as e:
                    item.removido, item.tirado_na_etapa4 = f"erro ao ler a planilha: {e}", True
                    self.registrar(f"{item.curso.posgrad} fora do lote: {item.removido}")
            return gerador.validar(self.ativos)

        self.ocupado("Lendo as planilhas…")
        em_segundo_plano(self, trabalho, self._dados_lidos)

    def _dados_lidos(self, resultado):
        self.problemas, self.mescladas, self.linhas_ok = resultado
        for p in self.problemas:              # erro começa como "pular"; aviso começa como "gerar"
            p.registro.pular = p.registro.tem_erro
        lidos_com_erro = [i for i in self.lote if i.removido.startswith("erro ao ler")]
        if lidos_com_erro:
            avisar(self, "Planilha com erro", "\n".join(f"• {i.curso.posgrad}: {i.removido}" for i in lidos_com_erro))
        if not self.ativos:
            self.livre()
            return
        self.registrar(f"Validação: {self.linhas_ok} linhas prontas | {self.mescladas} células mescladas | "
                       f"{len(self.problemas)} linhas com problema")
        self.ir_para(3)

    # -- 5. Conferir dados ------------------------------------------------------

    def montar_conferir(self):
        if not self.problemas:
            self.cabecalho("Conferir dados", "")
            ctk.CTkLabel(self.conteudo, text="✓  Nenhum problema encontrado nas planilhas.", text_color=COR_OK,
                         font=(FONTE, 18, "bold"), anchor="w").pack(fill="x", pady=10)
            ctk.CTkLabel(self.conteudo, anchor="w", text=f"Linhas prontas para gerar: {self.linhas_ok}\n"
                                                         f"Células mescladas lidas: {self.mescladas}",
                         justify="left").pack(fill="x")
            return
        self.cabecalho("Conferir dados",
                       f"{self.linhas_ok} linhas estão prontas. Linhas com ERRO (campo vazio ou inválido) começam como "
                       "PULAR: corrija para gerar. Linhas só com AVISO começam como GERAR ASSIM: confira e pule se "
                       "estiver errado.")
        quadro, self.arvore_prob = tabela(self.conteudo, [("curso", "Curso", 180), ("mes", "Mês", 80),
                                                          ("linha", "Linha", 60), ("docente", "Docente", 170),
                                                          ("problema", "Problema", 280), ("decisao", "Decisão", 110)],
                                          alturas=12)
        quadro.pack(fill="both", expand=True)
        self.arvore_prob.bind("<Double-1>", lambda e: self._corrigir())
        botoes = ctk.CTkFrame(self.conteudo, fg_color="transparent")
        botoes.pack(fill="x", pady=8)
        ctk.CTkButton(botoes, text="Corrigir…", command=self._corrigir).pack(side="left")
        ctk.CTkButton(botoes, text="Gerar assim", command=self._gerar_assim).pack(side="left", padx=8)
        ctk.CTkButton(botoes, text="Pular", **tema.BOTAO_CONTORNO,
                      command=lambda: self._decidir(True)).pack(side="left")
        self._desenhar_problemas()

    def _desenhar_problemas(self, selecionar=None):
        a = self.arvore_prob
        a.delete(*a.get_children())
        for i, p in enumerate(self.problemas):
            r = p.registro
            if r.pular:
                decisao, tag = "pular", "apagado"
            elif r.tem_problema:
                decisao, tag = "gerar assim", "aviso"
            else:
                decisao, tag = "corrigida", "ok"
            a.insert("", "end", iid=str(i), values=(p.curso.curso.posgrad, r.aba, r.linha, r.nome or "(vazio)",
                                                    "; ".join(p.descricao()) or "corrigido", decisao), tags=(tag,))
        if selecionar is not None:
            a.selection_set(str(selecionar))
        pulando = sum(1 for p in self.problemas if p.registro.pular)
        self.status.configure(text=f"{pulando} linha(s) serão puladas", text_color=COR_APAGADO)

    def _problema_escolhido(self):
        sel = self.arvore_prob.selection()
        if not sel:
            avisar(self, "Escolha uma linha", "Clique primeiro na linha.")
            return None, None
        return int(sel[0]), self.problemas[int(sel[0])]

    def _corrigir(self):
        n, p = self._problema_escolhido()
        if p is None:
            return
        r = p.registro
        if not (r.tem_erro or r.suspeita_nome_email or r.suspeita_nome_com_extra or r.suspeita_curso_data):
            avisar(self, "Nada para corrigir", "Esta linha já está corrigida. Use \"Gerar assim\" ou \"Pular\".")
            return
        d = DialogoCorrecao(self, p)
        self.wait_window(d)
        if d.escolha != "aplicar":
            return
        v, erros = d.valores, []
        if v.get("nome"):
            r.corrigir_nome(v["nome"])
        if v.get("curso"):
            r.corrigir_curso(v["curso"])
        if "data" in v and not r.corrigir_data(v["data"]):
            erros.append("a data não está no formato dd/mm/aaaa")
        if "horas" in v and not r.corrigir_horas(v["horas"]):
            erros.append("a duração não foi entendida (use 4 ou 2:30)")
        r.pular = r.tem_erro
        self.registrar(f"Correção manual — {p.curso.curso.posgrad} / {r.aba} / linha {r.linha}"
                       + (f" (ainda com erro: {'; '.join(erros)})" if erros else ""))
        self._desenhar_problemas(n)
        if erros:
            avisar(self, "Ainda falta corrigir", "\n".join(f"• {e}" for e in erros))

    def _gerar_assim(self):
        n, p = self._problema_escolhido()
        if p is None:
            return
        if p.registro.tem_erro:
            avisar(self, "Não dá para gerar assim", "Esta linha tem erro (campo vazio ou inválido). Use \"Corrigir…\".")
            return
        self._decidir(False)

    def _decidir(self, pular):
        n, p = self._problema_escolhido()
        if p is None:
            return
        p.registro.pular = pular
        self._desenhar_problemas(n)

    # -- 6. Gerar ---------------------------------------------------------------

    def _a_gerar(self, item):
        return sum(1 for a in item.dados.abas for r in a.registros if not r.pular and not r.tem_erro)

    def _docentes(self, item):
        """(docentes com e-mail válido, docentes sem e-mail) que vão ter certificado."""
        com, sem = set(), set()
        for a in item.dados.abas:
            for r in a.registros:
                if not r.pular and not r.tem_erro:
                    chave = textos.sanitizar_nome_arquivo(r.nome).lower()
                    (com if r.email_valido else sem).add(chave)
        return com, sem - com

    def _ja_gerados(self):
        """Curso/mês já enviado (histórico ou pasta arquivada) ou com PDF na pasta local deste modo."""
        achados = []
        for item in self.ativos:
            enviados = {m.pasta_mes for m in (self.situacoes or {}).get(item.curso.linha, [])}
            for aba in item.abas:
                mes = textos.mes_para_pasta(aba)
                pasta = os.path.join(self.modo.pasta_local, item.nome_com_unidade, mes)
                if mes in enviados or (os.path.isdir(pasta) and any(n.lower().endswith(".pdf")
                                                                    for n in os.listdir(pasta))):
                    achados.append(f"{item.curso.posgrad} / {aba}")
        return achados

    def montar_gerar(self):
        total = sum(self._a_gerar(i) for i in self.ativos)
        if self.oficial:
            com = sum(len(self._docentes(i)[0]) for i in self.ativos)
            texto = (f"{total} certificado(s) em {len(self.ativos)} curso(s). Data de emissão: {self.data_emissao}. "
                     f"{com} docente(s) vão receber e-mail.")
        else:
            texto = (f"{total} certificado(s) de teste em {len(self.ativos)} curso(s). "
                     f"Data de emissão: {self.data_emissao}. E-mails para: {self.email_teste}")
        self.cabecalho("Gerar", texto)
        colunas = [("curso", "Curso", 270), ("assina", "Assina", 210), ("meses", "Meses", 120),
                   ("n", "Certificados", 90), ("email", "Recebem e-mail", 110), ("sem", "Sem e-mail", 90)]
        quadro, arvore = tabela(self.conteudo, colunas, alturas=6)
        quadro.pack(fill="x")
        for item in self.ativos:
            com, sem = self._docentes(item)
            arvore.insert("", "end", values=(item.curso.posgrad, item.coordenacao.assina_como, ", ".join(item.abas),
                                             self._a_gerar(item), len(com), len(sem)))
        repetidos = self._ja_gerados() if self.oficial else []
        if repetidos:
            tema.quebrar_na_largura(ctk.CTkLabel(self.conteudo, justify="left", anchor="w", text_color=COR_AVISO,
                         text="Atenção: já existem certificados gerados para " + "; ".join(repetidos)
                              + ". Gerar de novo manda os e-mails outra vez.")).pack(fill="x", pady=(8, 0))
        tema.quebrar_na_largura(ctk.CTkLabel(self.conteudo, justify="left", anchor="w", text_color=tema.TEXTO_2, text=(
            f"Os PDFs vão para {self.modo.pasta_local} e para Para Enviar\\{self.modo.prefixo}<Curso>\\<Mês>. "
            "O fluxo do Power Automate percebe o manifesto em alguns minutos e manda os e-mails."))).pack(
            fill="x", pady=(10, 4))
        self.barra = ctk.CTkProgressBar(self.conteudo)
        self.barra.set(0)
        self.barra.pack(fill="x", pady=(6, 4))
        self.caixa_log = ctk.CTkTextbox(self.conteudo, font=("Consolas", 11), height=200)
        self.caixa_log.pack(fill="both", expand=True, pady=(4, 0))
        self.botao_avancar.configure(text="Gerar e enviar aos docentes" if self.oficial
                                     else "Gerar certificados de teste")
        self.botao_voltar.configure(state="normal", command=lambda: self.ir_para(3))

    def _escrever_log(self, msg):
        self.after(0, lambda: (self.caixa_log.insert("end", f"{datetime.now():%H:%M:%S}  {msg}\n"),
                               self.caixa_log.see("end")))

    def gerar(self):
        total = sum(self._a_gerar(i) for i in self.ativos)
        if total == 0:
            avisar(self, "Nada para gerar", "Todas as linhas foram puladas.")
            return
        aviso_ppt = "\n\nO PowerPoint está aberto: salve e feche as apresentações antes." if powerpoint_aberto() else ""
        if self.oficial:
            com = sum(len(self._docentes(i)[0]) for i in self.ativos)
            cursos = "\n".join(f"• {i.curso.posgrad} ({', '.join(i.abas)})" for i in self.ativos)
            repetidos = self._ja_gerados()
            aviso_rep = ("\n\nATENÇÃO: já gerados antes (os e-mails vão de novo):\n"
                         + "\n".join(f"• {r}" for r in repetidos)) if repetidos else ""
            pergunta = ("Enviar para os docentes?",
                        f"{total} certificado(s) e e-mail para {com} docente(s) de verdade:\n\n{cursos}"
                        f"{aviso_rep}{aviso_ppt}",
                        [("Ainda não", None), (f"Enviar para {com} docente(s)", True)])
        else:
            pergunta = ("Gerar agora?", f"Gerar {total} certificado(s) de teste e mandar o manifesto para o "
                                        f"fluxo?\n\nOs e-mails vão para: {self.email_teste}{aviso_ppt}",
                        [("Ainda não", None), ("Gerar", True)])
        if perguntar(self, *pergunta) is not True:
            return
        self.registrar(f"Geração confirmada: {total} certificado(s), modo {self.modo.rotulo}")
        self.ocupado("Gerando…")
        self.cancelar = False
        self.botao_voltar.configure(state="normal", text="Parar", command=self._parar)
        self.log.ecoar = self._escrever_log
        lote = self.ativos
        for item in ([] if self.oficial else lote):   # teste anterior do mesmo curso/mês sai da pasta local
            for aba in item.abas:
                shutil.rmtree(os.path.join(self.modo.pasta_local, item.nome_com_unidade, textos.mes_para_pasta(aba)),
                              ignore_errors=True)

        def progresso(texto, feitos, total):
            self.after(0, lambda: (self.barra.set(feitos / max(total, 1)),
                                   self.status.configure(text=f"{feitos} de {total} · {texto[:70]}")))

        def trabalho():
            with ConversorPowerPoint(log=self.log) as conv:
                return gerador.gerar(lote, self.locais, self.modelo, self.data_emissao, self.modo, conv, self.log,
                                     email_teste="" if self.oficial else self.email_teste, progresso=progresso,
                                     cancelado=lambda: self.cancelar)

        em_segundo_plano(self, trabalho, self._gerado, self._falhou_geracao)

    def _parar(self):
        if perguntar(self, "Parar a geração?", "O que já foi gerado fica. O curso que está no meio não recebe o "
                                               "manifesto completo.", [("Continuar gerando", None),
                                                                       ("Parar", True)]) is True:
            self.cancelar = True
            self.status.configure(text="Parando depois do certificado atual…")

    def _falhou_geracao(self, erro, detalhe):
        self.log.ecoar = None
        self.mostrar_erro(erro, detalhe)
        self.botao_voltar.configure(state="normal", text="← Voltar", command=lambda: self.ir_para(3))

    def _gerado(self, relatorio):
        self.log.ecoar = None
        self.relatorio = relatorio
        if self.oficial:
            historico.registrar(self.locais, relatorio, log=self.registrar)
            self.situacoes = None                 # relê na próxima rodada
            if self.inicio_rodada is not None and relatorio.total_gerados:
                segundos = time.monotonic() - self.inicio_rodada
                economia.registrar(relatorio.total_gerados, segundos)
                self.registrar(f"Tempo da rodada: {economia.duracao(segundos)} para {relatorio.total_gerados} "
                               f"certificado(s)")
        self.ir_para(RELATORIO)

    # -- 7. Relatório -----------------------------------------------------------

    def montar_relatorio(self):
        rel = self.relatorio
        falhou = [r for r in rel.cursos if r.falhou]
        ok = not falhou and not rel.falhas
        ctk.CTkLabel(self.conteudo, text="Geração concluída" if ok else "Geração concluída, com problemas",
                     font=(FONTE, 26, "bold"), text_color=COR_OK if ok else COR_AVISO, anchor="w").pack(fill="x")
        resumo = (f"{rel.total_gerados} PDF(s) gerado(s)  ·  data de emissão {rel.data_emissao}  ·  "
                  f"{len(falhou)} curso(s) com falha  ·  {len(rel.falhas)} certificado(s) com falha")
        ctk.CTkLabel(self.conteudo, text=resumo, anchor="w").pack(fill="x", pady=(2, 10))
        if self.modo.nome == "teste":
            tema.quebrar_na_largura(ctk.CTkLabel(self.conteudo, justify="left", anchor="w", text_color=tema.TEXTO_2, text=(
                f"Agora o fluxo do Power Automate deve mandar os e-mails para {self.email_teste} em alguns minutos. "
                f"Depois do envio ele guarda os PDFs em Declarações\\TESTE\\. Terminado o teste, você pode apagar "
                f"{self.modo.pasta_local} e a pasta Declarações\\TESTE."))).pack(fill="x", pady=(0, 10))
        else:
            tema.quebrar_na_largura(ctk.CTkLabel(self.conteudo, justify="left", anchor="w", text_color=tema.TEXTO_2, text=(
                "Agora o fluxo do Power Automate manda os e-mails para os docentes em alguns minutos e depois "
                "arquiva os PDFs em Declarações\\<Semestre>\\<Local>\\<Pós>\\<Mês>. A cópia de segurança fica em "
                f"{self.modo.pasta_local}."))).pack(fill="x", pady=(0, 10))
        quadro, arvore = tabela(self.conteudo, [("curso", "Curso", 340), ("gerados", "Gerados", 90),
                                                ("ignorados", "Ignorados", 90), ("sit", "Situação", 360)], alturas=8)
        quadro.pack(fill="both", expand=True)
        for r in rel.cursos:
            arvore.insert("", "end", values=(r.curso.curso.posgrad, r.gerados, r.ignorados,
                                             f"FALHOU — {r.falhou}" if r.falhou else "ok"),
                          tags=("erro" if r.falhou else "ok",))
        for item in self.lote:
            if item.removido:
                arvore.insert("", "end", values=(item.curso.posgrad, "—", "—", f"fora do lote: {item.removido}"),
                              tags=("apagado",))
        extras = [f"✗ {f}" for f in rel.falhas] + [f"! {a}" for a in rel.avisos]
        if extras:
            caixa = ctk.CTkTextbox(self.conteudo, height=110, font=("Consolas", 11))
            caixa.pack(fill="x", pady=(8, 0))
            caixa.insert("end", "\n".join(extras))
        botoes = ctk.CTkFrame(self.conteudo, fg_color="transparent")
        botoes.pack(fill="x", pady=(10, 0))
        para_enviar = os.path.join(self.locais.declaracoes, "Para Enviar")
        for rotulo, caminho in (("Abrir pasta dos PDFs", self.modo.pasta_local), ("Abrir Para Enviar", para_enviar),
                                ("Abrir log", rel.log)):
            ctk.CTkButton(botoes, text=rotulo, **tema.BOTAO_CONTORNO,
                          command=lambda c=caminho: self._abrir(c)).pack(side="left", padx=(0, 8))
        self.botao_avancar.configure(text="Nova rodada")
        self.registrar("Relatório mostrado.")

    def _abrir(self, caminho):
        try:
            os.startfile(caminho)
        except OSError as e:
            avisar(self, "Não abriu", f"{caminho}\n\n{e}")


if __name__ == "__main__":
    App().mainloop()
