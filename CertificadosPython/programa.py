"""
Certificados — o programa novo (Etapa 2), no lugar do GerarCertificados do AHK.

Dois modos, escolhidos na tela de Início:
  Teste    igual à TESTE_v7: os PDFs ganham "TESTE - " no nome, ficam em
           C:\\Certificados_TESTE e em "Para Enviar\\TESTE - <Curso>\\<Mês>", e no
           manifesto o e-mail de todos os docentes vira o e-mail de teste.
  Oficial  igual à v27: C:\\Certificados (com a cópia por docente) e
           "Para Enviar\\<Curso>\\<Mês>"; o fluxo manda os e-mails para os docentes.

Uma janela só, com as etapas na lateral (os mesmos passos do AHK):
  1. Início           modo, e-mail de teste e arquivos necessários (Passo 0 + verificação)
  2. Cursos           planilha mestre, semestre, data de emissão e cursos (Passos 1–3)
  3. Coordenadores    quem assina cada curso (Passo 4)
  4. Planilhas        planilha de cada curso e meses (Passos 5–7)
  5. Conferir dados   linhas com problema: corrigir, gerar assim ou pular (Fase 2)
  6. Gerar            PDFs, cópia para a Para Enviar e manifesto (Fase 3)
  7. Relatório        (Passo 10)
"""

import json
import os
import shutil
from datetime import datetime
from tkinter import filedialog, ttk

import customtkinter as ctk

from certificados import VERSAO, caminhos, coordenadores, gerador, horas, planilhas, textos
from certificados.pdf import ConversorPowerPoint, powerpoint_aberto
from conferencia import COR_APAGADO, COR_AVISO, COR_ERRO, COR_OK, FONTE, em_segundo_plano, linha_de_arquivo, tabela

ETAPAS = ["Início", "Cursos", "Coordenadores", "Planilhas e meses", "Conferir dados", "Gerar", "Relatório"]
GERAR, RELATORIO = 5, 6
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
        ctk.CTkLabel(self, text=titulo, font=(FONTE, 16, "bold"), anchor="w").pack(fill="x", padx=20, pady=(16, 6))
        ctk.CTkLabel(self, text=texto, justify="left", anchor="w", wraplength=largura - 40).pack(fill="x", padx=20)
        self.corpo = ctk.CTkFrame(self, fg_color="transparent")
        self.corpo.pack(fill="x", padx=20, pady=(8, 0))
        linha = ctk.CTkFrame(self, fg_color="transparent")
        linha.pack(fill="x", padx=20, pady=16)
        for i, (rotulo, valor) in enumerate(reversed(botoes)):
            principal = i == 0
            ctk.CTkButton(linha, text=rotulo, width=150,
                          fg_color=None if principal else "transparent", border_width=0 if principal else 1,
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
        ctk.CTkLabel(self.corpo, text=dica, anchor="w", text_color=COR_APAGADO, font=(FONTE, 11)).pack(fill="x")
        self.campos[chave] = var

    def fechar(self, valor):
        self.valores = {k: v.get().strip() for k, v in self.campos.items()}
        super().fechar(valor)


# ---------------------------------------------------------------------------
# Janela
# ---------------------------------------------------------------------------

class App(ctk.CTk):
    def __init__(self, modo="teste"):
        super().__init__()
        ctk.set_appearance_mode("dark")
        self.modo = gerador.MODOS[modo]
        self.geometry("1200x760")
        self.minsize(1000, 640)
        self.after(0, lambda: self.state("zoomed"))
        # CERTIFICADOS_RAIZ aponta para outra "raiz do OneDrive" (usado nos testes).
        self.locais = caminhos.Locais.descobrir(os.environ.get("CERTIFICADOS_RAIZ") or None)
        self.config_usuario = ler_config()
        self._estilo_tabelas()
        self._montar_janela()
        self.nova_rodada()

    def _estilo_tabelas(self):
        estilo = ttk.Style(self)
        estilo.theme_use("clam")
        estilo.configure("Treeview", background="#22272e", fieldbackground="#22272e", foreground="#d0d7de",
                         rowheight=26, font=(FONTE, 10), borderwidth=0)
        estilo.configure("Treeview.Heading", background="#2d333b", foreground="#d0d7de", font=(FONTE, 10, "bold"))
        estilo.map("Treeview", background=[("selected", "#316dca")])

    def _montar_janela(self):
        lateral = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color="#1c2128")
        lateral.pack(side="left", fill="y")
        lateral.pack_propagate(False)
        ctk.CTkLabel(lateral, text="Certificados", font=(FONTE, 22, "bold"), anchor="w").pack(fill="x", padx=20,
                                                                                            pady=(22, 0))
        ctk.CTkLabel(lateral, text="Pós-graduação", text_color=COR_APAGADO, anchor="w").pack(fill="x", padx=20)
        self.rotulos_etapas = []
        caixa = ctk.CTkFrame(lateral, fg_color="transparent")
        caixa.pack(fill="x", padx=12, pady=24)
        for i, nome in enumerate(ETAPAS):
            r = ctk.CTkLabel(caixa, text=f"  {i + 1}   {nome}", anchor="w", height=34, corner_radius=6,
                             font=(FONTE, 13))
            r.pack(fill="x", pady=1)
            self.rotulos_etapas.append(r)
        ctk.CTkLabel(lateral, text=f"versão {VERSAO}", text_color=COR_APAGADO, font=(FONTE, 11)).pack(
            side="bottom", pady=12)

        direita = ctk.CTkFrame(self, fg_color="transparent")
        direita.pack(side="left", fill="both", expand=True)
        self.faixa = ctk.CTkLabel(direita, text="", font=(FONTE, 12, "bold"), height=30)
        self.faixa.pack(fill="x")
        self._aplicar_modo()
        self.conteudo = ctk.CTkFrame(direita, fg_color="transparent")
        self.conteudo.pack(fill="both", expand=True, padx=24, pady=(14, 0))
        rodape = ctk.CTkFrame(direita, fg_color="transparent")
        rodape.pack(fill="x", padx=24, pady=14)
        self.botao_voltar = ctk.CTkButton(rodape, text="← Voltar", width=120, fg_color="transparent",
                                          border_width=1, command=self.voltar)
        self.botao_voltar.pack(side="left")
        self.botao_avancar = ctk.CTkButton(rodape, text="Avançar →", width=220, height=38,
                                           font=(FONTE, 14, "bold"), command=self.avancar)
        self.botao_avancar.pack(side="right")
        self.status = ctk.CTkLabel(rodape, text="", anchor="w", text_color=COR_APAGADO)
        self.status.pack(side="left", fill="x", expand=True, padx=16)

    def _aplicar_modo(self):
        """Título da janela e faixa do alto: amarela no teste, vermelha no oficial."""
        self.title(f"Certificados — modo {self.modo.rotulo} (versão {VERSAO})")
        if self.modo.nome == "teste":
            self.faixa.configure(text="MODO TESTE  ·  nada vai para os docentes: todos os e-mails vão para o "
                                      "e-mail de teste", fg_color=COR_AVISO, text_color="#1c1c1c")
        else:
            self.faixa.configure(text="MODO OFICIAL  ·  os e-mails vão para os docentes de verdade",
                                 fg_color=COR_ERRO, text_color="#ffffff")

    @property
    def oficial(self):
        return self.modo.nome == "oficial"

    # -- estado e navegação ---------------------------------------------------

    def nova_rodada(self):
        self.email_teste = self.config_usuario.get("email_teste", "")
        self.mestre = self.locais.planilha_mestre
        self.modelo = self.locais.modelo
        self.abas_mestre, self.aba = [], ""
        self.data_emissao = datetime.now().strftime("%d/%m/%Y")
        self.cursos, self.marcados = [], set()   # marcados = linhas da mestre
        self.lote, self.cadastro = [], None
        self.abas_por_planilha, self.erros_planilha = {}, {}
        self.meses_marcados = set()
        self.problemas, self.mescladas, self.linhas_ok = [], 0, 0
        self.relatorio, self.log, self.modo_do_log, self.cancelar = None, None, "", False
        self.ir_para(0)

    @property
    def ativos(self):
        return [i for i in self.lote if not i.removido]

    def ir_para(self, n):
        self.etapa = n
        for w in self.conteudo.winfo_children():
            w.destroy()
        for i, r in enumerate(self.rotulos_etapas):
            if i == n:
                r.configure(fg_color="#316dca", text_color="#ffffff", font=(FONTE, 13, "bold"))
            else:
                r.configure(fg_color="transparent", text_color="#d0d7de" if i < n else COR_APAGADO,
                            font=(FONTE, 13))
            r.configure(text=f"  {'✓' if i < n else i + 1}   {ETAPAS[i]}")
        self.status.configure(text="", text_color=COR_APAGADO)
        self.botao_voltar.configure(state="normal" if 0 < n < GERAR else "disabled", text="← Voltar",
                                    command=self.voltar)
        self.botao_avancar.configure(state="normal", text="Avançar →")
        [self.montar_inicio, self.montar_cursos, self.montar_coordenadores, self.montar_planilhas,
         self.montar_conferir, self.montar_gerar, self.montar_relatorio][n]()

    def voltar(self):
        if 0 < self.etapa < GERAR:
            self.ir_para(self.etapa - 1)

    def avancar(self):
        [self.avancar_inicio, self.avancar_cursos, self.avancar_coordenadores, self.avancar_planilhas,
         lambda: self.ir_para(GERAR), self.gerar, self.nova_rodada][self.etapa]()

    def ocupado(self, texto):
        self.botao_avancar.configure(state="disabled")
        self.botao_voltar.configure(state="disabled")
        self.status.configure(text=texto, text_color=COR_APAGADO)

    def livre(self, texto=""):
        self.botao_avancar.configure(state="normal")
        self.botao_voltar.configure(state="normal" if 0 < self.etapa < GERAR else "disabled")
        self.status.configure(text=texto)

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
        ctk.CTkLabel(janela, text="Algo deu errado", font=(FONTE, 16, "bold")).pack(anchor="w", padx=16, pady=(14, 4))
        caixa = ctk.CTkTextbox(janela, font=("Consolas", 11))
        caixa.pack(fill="both", expand=True, padx=16, pady=4)
        caixa.insert("end", f"{erro}\n\n{detalhe}")
        ctk.CTkButton(janela, text="Fechar", command=janela.destroy).pack(pady=10)

    def cabecalho(self, titulo, texto):
        ctk.CTkLabel(self.conteudo, text=titulo, font=(FONTE, 20, "bold"), anchor="w").pack(fill="x")
        ctk.CTkLabel(self.conteudo, text=texto, justify="left", anchor="w", wraplength=880,
                     text_color="#adbac7").pack(fill="x", pady=(2, 12))

    # -- 1. Início --------------------------------------------------------------

    def montar_inicio(self):
        if self.oficial:
            self.cabecalho("Antes de começar",
                           "Modo oficial: os certificados vão para a Para Enviar e o Power Automate manda os e-mails "
                           "para os docentes, como na v27 do AHK. Para experimentar sem enviar nada, escolha Teste.")
        else:
            self.cabecalho("Antes de começar",
                           "Modo teste: os certificados são gerados normalmente, mas no manifesto o e-mail de TODOS os "
                           "docentes é trocado pelo e-mail abaixo. Assim, o Power Automate manda tudo só para esse "
                           "endereço.")
        linha = ctk.CTkFrame(self.conteudo, fg_color="transparent")
        linha.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(linha, text="Modo", width=150, anchor="w", font=(FONTE, 13, "bold")).pack(side="left")
        seletor = ctk.CTkSegmentedButton(linha, values=["Teste", "Oficial"], command=self._trocar_modo, width=240)
        seletor.set(self.modo.rotulo)
        seletor.pack(side="left")
        self.var_email = ctk.StringVar(value=self.email_teste)
        if not self.oficial:
            linha = ctk.CTkFrame(self.conteudo, fg_color="transparent")
            linha.pack(fill="x", pady=(0, 10))
            ctk.CTkLabel(linha, text="E-mail de teste", width=150, anchor="w",
                         font=(FONTE, 13, "bold")).pack(side="left")
            ctk.CTkEntry(linha, textvariable=self.var_email, width=360,
                         placeholder_text="seu.nome@einstein.br").pack(side="left")
        self.var_modelo = linha_de_arquivo(self.conteudo, "Modelo (.pptx)", self.modelo, self._conferir_arquivos,
                                           [("PowerPoint", "*.pptx")])
        ctk.CTkLabel(self.conteudo, text="Arquivos e pastas", font=(FONTE, 14, "bold"), anchor="w").pack(
            fill="x", pady=(14, 4))
        self.quadro_arquivos = ctk.CTkFrame(self.conteudo)
        self.quadro_arquivos.pack(fill="x")
        ctk.CTkButton(self.conteudo, text="Conferir de novo", width=150, fg_color="transparent", border_width=1,
                      command=self._conferir_arquivos).pack(anchor="w", pady=8)
        self._conferir_arquivos()

    def _trocar_modo(self, rotulo):
        novo = "oficial" if rotulo == "Oficial" else "teste"
        if novo == self.modo.nome:
            return
        self.email_teste, self.modelo = self.var_email.get().strip(), self.var_modelo.get()
        self.modo = gerador.MODOS[novo]
        self._aplicar_modo()
        self.ir_para(0)

    def _itens_verificacao(self):
        """(rótulo, caminho, ok, obrigatório). Igual à VerificacaoPreRodagem do AHK."""
        l = self.locais
        para_enviar = os.path.join(l.declaracoes, "Para Enviar") if l.declaracoes else ""
        modelo = self.var_modelo.get()
        return [
            ("Pasta Declarações (OneDrive)", l.declaracoes or "OneDrive da empresa não encontrado",
             bool(l.declaracoes) and os.path.isdir(l.declaracoes), True),
            ("Pasta Para Enviar", para_enviar, bool(para_enviar) and os.path.isdir(para_enviar), True),
            ("Modelo do certificado", modelo, os.path.isfile(modelo), True),
            ("Logo", l.logo, os.path.isfile(l.logo), True),
            ("Pasta de assinaturas", l.assinaturas, os.path.isdir(l.assinaturas), True),
            ("Planilha mestre", self.mestre or "—", os.path.isfile(self.mestre or ""), False),
        ]

    def _conferir_arquivos(self):
        for w in self.quadro_arquivos.winfo_children():
            w.destroy()
        for i, (rotulo, caminho, ok, obrigatorio) in enumerate(self._itens_verificacao()):
            cor = COR_OK if ok else (COR_ERRO if obrigatorio else COR_AVISO)
            simbolo = "✓" if ok else ("✗" if obrigatorio else "!")
            ctk.CTkLabel(self.quadro_arquivos, text=simbolo, text_color=cor, width=30,
                         font=(FONTE, 14, "bold")).grid(row=i, column=0, padx=(10, 0), pady=3)
            ctk.CTkLabel(self.quadro_arquivos, text=rotulo, anchor="w", width=210).grid(row=i, column=1, sticky="w")
            ctk.CTkLabel(self.quadro_arquivos, text=caminho, anchor="w", text_color=COR_APAGADO).grid(
                row=i, column=2, sticky="w", padx=(0, 10))
        n = len(self._itens_verificacao())
        if powerpoint_aberto():
            ctk.CTkLabel(self.quadro_arquivos, text="!", text_color=COR_AVISO, width=30,
                         font=(FONTE, 14, "bold")).grid(row=n, column=0, padx=(10, 0), pady=3)
            ctk.CTkLabel(self.quadro_arquivos, anchor="w", text_color=COR_AVISO,
                         text="O PowerPoint está aberto. Salve e feche as apresentações antes de gerar.").grid(
                row=n, column=1, columnspan=2, sticky="w")

    def avancar_inicio(self):
        email = self.var_email.get().strip()
        if not self.oficial and not textos.email_valido(email):
            avisar(self, "E-mail inválido", "Digite um e-mail de teste válido, por exemplo seu.nome@einstein.br.")
            return
        faltando = [f"• {rot}: {cam}" for rot, cam, ok, obrig in self._itens_verificacao() if obrig and not ok]
        if faltando:
            avisar(self, "Falta algum arquivo", "Sem estes itens não dá para gerar:\n\n" + "\n".join(faltando))
            return
        if self.oficial:
            pergunta = ("Modo oficial", "Nesta rodada os e-mails vão para os DOCENTES de verdade.\n\n"
                                        "Antes de gerar, o programa ainda mostra o resumo e pede confirmação.",
                        [("Voltar", None), ("Entendi, continuar", True)])
        else:
            pergunta = ("Confirmar e-mail de teste",
                        f"Todos os e-mails do manifesto irão para:\n\n{email}\n\nEstá correto?",
                        [("Corrigir", None), ("Sim, continuar", True)])
        if perguntar(self, *pergunta) is not True:
            return
        self.email_teste, self.modelo = email, self.var_modelo.get()
        if not self.oficial:
            self.config_usuario["email_teste"] = email
            gravar_config(self.config_usuario)
        if self.log is None or self.modo_do_log != self.modo.nome:
            self.log, self.modo_do_log = gerador.Log(self.modo.nome), self.modo.nome
        self.registrar(f"MODO {self.modo.rotulo.upper()}"
                       + ("" if self.oficial else f" — e-mails do manifesto: {email}") + f" | Modelo: {self.modelo}")
        self.ir_para(1)

    # -- 2. Cursos --------------------------------------------------------------

    def montar_cursos(self):
        self.cabecalho("Cursos", "Escolha o semestre e marque os cursos desta rodada. A data de emissão sai "
                                 "impressa em todos os certificados.")
        self.var_mestre = linha_de_arquivo(self.conteudo, "Planilha mestre", self.mestre, self._trocar_mestre,
                                           [("Excel", "*.xlsx")])
        linha = ctk.CTkFrame(self.conteudo, fg_color="transparent")
        linha.pack(fill="x", pady=3)
        ctk.CTkLabel(linha, text="Semestre (aba)", width=150, anchor="w").pack(side="left")
        self.combo_aba = ctk.CTkComboBox(linha, values=self.abas_mestre, width=160, command=self._trocar_aba)
        self.combo_aba.set(self.aba)
        self.combo_aba.pack(side="left")
        ctk.CTkLabel(linha, text="Data de emissão", anchor="w").pack(side="left", padx=(30, 8))
        self.var_data = ctk.StringVar(value=self.data_emissao)
        ctk.CTkEntry(linha, textvariable=self.var_data, width=110).pack(side="left")

        cab = ctk.CTkFrame(self.conteudo, fg_color="transparent")
        cab.pack(fill="x", pady=(12, 4))
        self.rotulo_marcados = ctk.CTkLabel(cab, text="", font=(FONTE, 13, "bold"), anchor="w")
        self.rotulo_marcados.pack(side="left")
        self.var_filtro = ctk.StringVar()
        self.var_filtro.trace_add("write", lambda *_: self._desenhar_cursos())
        ctk.CTkEntry(cab, textvariable=self.var_filtro, placeholder_text="procurar curso…", width=220).pack(
            side="right")
        self.lista_cursos = ctk.CTkScrollableFrame(self.conteudo)
        self.lista_cursos.pack(fill="both", expand=True, pady=(0, 4))
        if self.abas_mestre:
            self._desenhar_cursos()
        else:
            self._carregar_abas()

    def _trocar_mestre(self):
        self.mestre = self.var_mestre.get()
        self.abas_mestre, self.aba, self.cursos, self.marcados = [], "", [], set()
        self._carregar_abas()

    def _carregar_abas(self):
        if not os.path.isfile(self.mestre or ""):
            self.status.configure(text="Escolha a planilha mestre.")
            return
        self.ocupado("Lendo a planilha mestre…")
        em_segundo_plano(self, lambda: planilhas.listar_abas(self.mestre), self._abas_carregadas)

    def _abas_carregadas(self, abas):
        self.livre()
        self.abas_mestre = abas
        self.combo_aba.configure(values=abas)
        semestres = [a for a in abas if any(ch.isdigit() for ch in a)]
        if semestres:
            self._trocar_aba(semestres[-1])

    def _trocar_aba(self, aba):
        self.combo_aba.set(aba)
        if aba == self.aba and self.cursos:
            return
        self.aba, self.cursos, self.marcados = aba, [], set()
        self.ocupado(f"Lendo os cursos da aba {aba}…")
        em_segundo_plano(self, lambda: planilhas.ler_cursos_da_mestre(self.mestre, aba), self._cursos_carregados)

    def _cursos_carregados(self, cursos):
        self.livre()
        self.cursos = cursos or []
        if cursos is None:
            self.status.configure(text=f"A aba {self.aba} não foi encontrada.", text_color=COR_ERRO)
        self._desenhar_cursos()

    def _desenhar_cursos(self):
        for w in self.lista_cursos.winfo_children():
            w.destroy()
        filtro = textos.normalizar(self.var_filtro.get())
        for c in self.cursos:
            if filtro and filtro not in textos.normalizar(f"{c.posgrad} {c.unidade} {c.coordenacao}"):
                continue
            var = ctk.BooleanVar(value=c.linha in self.marcados)
            texto = f"{c.posgrad}   ·   {c.unidade_exibida}"
            ctk.CTkCheckBox(self.lista_cursos, text=texto, variable=var,
                            command=lambda c=c, v=var: self._marcar(c, v.get())).pack(anchor="w", pady=3, padx=4)
        self._contar_marcados()

    def _marcar(self, curso, marcado):
        (self.marcados.add if marcado else self.marcados.discard)(curso.linha)
        self._contar_marcados()

    def _contar_marcados(self):
        self.rotulo_marcados.configure(text=f"{len(self.cursos)} cursos na aba {self.aba or '—'}  ·  "
                                            f"{len(self.marcados)} marcado(s)")

    def avancar_cursos(self):
        data = self.var_data.get().strip()
        if not horas.data_valida(data):
            avisar(self, "Data de emissão", "Digite a data de emissão no formato dd/mm/aaaa.")
            return
        escolhidos = [c for c in self.cursos if c.linha in self.marcados]
        if not escolhidos:
            avisar(self, "Nenhum curso", "Marque pelo menos um curso.")
            return
        self.data_emissao = data
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
        self.ir_para(2)

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
        ctk.CTkButton(botoes, text="Tirar do lote / devolver", fg_color="transparent", border_width=1,
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
        em_segundo_plano(self, self._listar_abas_das_planilhas, lambda _: self.ir_para(3))

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
        ctk.CTkLabel(self.conteudo, text="Meses desta rodada", font=(FONTE, 14, "bold"), anchor="w").pack(
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
        self.ir_para(4)

    # -- 5. Conferir dados ------------------------------------------------------

    def montar_conferir(self):
        if not self.problemas:
            self.cabecalho("Conferir dados", "")
            ctk.CTkLabel(self.conteudo, text="✓  Nenhum problema encontrado nas planilhas.", text_color=COR_OK,
                         font=(FONTE, 16, "bold"), anchor="w").pack(fill="x", pady=10)
            ctk.CTkLabel(self.conteudo, anchor="w", text=f"Linhas prontas para gerar: {self.linhas_ok}\n"
                                                         f"Células mescladas lidas: {self.mescladas}",
                         justify="left").pack(fill="x")
            return
        self.cabecalho("Conferir dados",
                       f"{self.linhas_ok} linhas estão prontas. Linhas com ERRO (campo vazio ou inválido) começam como "
                       "PULAR: corrija para gerar. Linhas só com AVISO começam como GERAR ASSIM: confira e pule se "
                       "estiver errado.")
        quadro, self.arvore_prob = tabela(self.conteudo, [("curso", "Curso", 200), ("mes", "Mês", 80),
                                                          ("linha", "Linha", 50), ("docente", "Docente", 190),
                                                          ("problema", "Problema", 330), ("decisao", "Decisão", 110)],
                                          alturas=12)
        quadro.pack(fill="both", expand=True)
        self.arvore_prob.bind("<Double-1>", lambda e: self._corrigir())
        botoes = ctk.CTkFrame(self.conteudo, fg_color="transparent")
        botoes.pack(fill="x", pady=8)
        ctk.CTkButton(botoes, text="Corrigir…", command=self._corrigir).pack(side="left")
        ctk.CTkButton(botoes, text="Gerar assim", command=self._gerar_assim).pack(side="left", padx=8)
        ctk.CTkButton(botoes, text="Pular", fg_color="transparent", border_width=1,
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
        """Curso/mês que já tem PDF na pasta local deste modo (rodada anterior)."""
        achados = []
        for item in self.ativos:
            for aba in item.abas:
                pasta = os.path.join(self.modo.pasta_local, item.nome_com_unidade, textos.mes_para_pasta(aba))
                if os.path.isdir(pasta) and any(n.lower().endswith(".pdf") for n in os.listdir(pasta)):
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
            ctk.CTkLabel(self.conteudo, justify="left", anchor="w", text_color=COR_AVISO, wraplength=880,
                         text="Atenção: já existem certificados gerados para " + "; ".join(repetidos)
                              + ". Gerar de novo manda os e-mails outra vez.").pack(fill="x", pady=(8, 0))
        ctk.CTkLabel(self.conteudo, justify="left", anchor="w", text_color="#adbac7", wraplength=880, text=(
            f"Os PDFs vão para {self.modo.pasta_local} e para Para Enviar\\{self.modo.prefixo}<Curso>\\<Mês>. "
            "O fluxo do Power Automate percebe o manifesto em alguns minutos e manda os e-mails.")).pack(
            fill="x", pady=(10, 4))
        self.barra = ctk.CTkProgressBar(self.conteudo)
        self.barra.set(0)
        self.barra.pack(fill="x", pady=(6, 4))
        self.caixa_log = ctk.CTkTextbox(self.conteudo, font=("Consolas", 11), height=200)
        self.caixa_log.pack(fill="both", expand=True, pady=(4, 0))
        self.botao_avancar.configure(text="Gerar e enviar aos docentes" if self.oficial
                                     else "Gerar certificados de teste")
        self.botao_voltar.configure(state="normal", command=lambda: self.ir_para(4))

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
        self.botao_voltar.configure(state="normal", text="← Voltar", command=lambda: self.ir_para(4))

    def _gerado(self, relatorio):
        self.log.ecoar = None
        self.relatorio = relatorio
        self.ir_para(RELATORIO)

    # -- 7. Relatório -----------------------------------------------------------

    def montar_relatorio(self):
        rel = self.relatorio
        falhou = [r for r in rel.cursos if r.falhou]
        ok = not falhou and not rel.falhas
        ctk.CTkLabel(self.conteudo, text="Geração concluída" if ok else "Geração concluída, com problemas",
                     font=(FONTE, 22, "bold"), text_color=COR_OK if ok else COR_AVISO, anchor="w").pack(fill="x")
        resumo = (f"{rel.total_gerados} PDF(s) gerado(s)  ·  data de emissão {rel.data_emissao}  ·  "
                  f"{len(falhou)} curso(s) com falha  ·  {len(rel.falhas)} certificado(s) com falha")
        ctk.CTkLabel(self.conteudo, text=resumo, anchor="w").pack(fill="x", pady=(2, 10))
        if self.modo.nome == "teste":
            ctk.CTkLabel(self.conteudo, justify="left", anchor="w", wraplength=880, text_color="#adbac7", text=(
                f"Agora o fluxo do Power Automate deve mandar os e-mails para {self.email_teste} em alguns minutos. "
                f"Depois do envio ele guarda os PDFs em Declarações\\TESTE\\. Terminado o teste, você pode apagar "
                f"{self.modo.pasta_local} e a pasta Declarações\\TESTE.")).pack(fill="x", pady=(0, 10))
        else:
            ctk.CTkLabel(self.conteudo, justify="left", anchor="w", wraplength=880, text_color="#adbac7", text=(
                "Agora o fluxo do Power Automate manda os e-mails para os docentes em alguns minutos e depois "
                "arquiva os PDFs em Declarações\\<Semestre>\\<Local>\\<Pós>\\<Mês>. A cópia de segurança fica em "
                f"{self.modo.pasta_local}.")).pack(fill="x", pady=(0, 10))
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
            ctk.CTkButton(botoes, text=rotulo, fg_color="transparent", border_width=1,
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
