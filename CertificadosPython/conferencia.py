"""
Conferência do motor Python contra o AHK, com os arquivos reais.

Três abas:
  1. Modelo      monta o modelo-base do jeito do AHK (no PowerPoint) e do jeito
                 do Python e compara a posição de cada forma.
  2. Simulação   gera certificados de cursos reais SÓ neste computador, em
                 C:\\Certificados_SIMULACAO. Não grava na nuvem, não mexe na
                 "Para Enviar", não muda o cadastro de coordenadores e não
                 envia e-mail.
  3. Comparar    compara o texto de cada PDF da simulação com o PDF que o AHK
                 gerou em C:\\Certificados.
"""

import os
import shutil
import sys
import threading
import traceback
from datetime import datetime
from tkinter import filedialog, ttk

import customtkinter as ctk

from certificados import (VERSAO, caminhos, comparacao, coordenadores, gerador, horas, planilhas,
                          referencia_ahk, textos)

FONTE = "Segoe UI"
COR_OK, COR_ERRO, COR_AVISO, COR_APAGADO = "#3fb37f", "#e5534b", "#d4a72c", "#8b949e"


def em_segundo_plano(janela, trabalho, ao_terminar, ao_falhar=None):
    """Roda "trabalho" numa thread e chama ao_terminar(resultado) na janela."""
    def rodar():
        try:
            resultado = trabalho()
        except Exception as e:
            detalhe = traceback.format_exc()
            janela.after(0, lambda: (ao_falhar or janela.mostrar_erro)(e, detalhe))
            return
        janela.after(0, lambda: ao_terminar(resultado))
    threading.Thread(target=rodar, daemon=True).start()


def tabela(pai, colunas, alturas=12):
    """ttk.Treeview com barra de rolagem; colunas = [(id, título, largura)]."""
    quadro = ctk.CTkFrame(pai, fg_color="transparent")
    arvore = ttk.Treeview(quadro, columns=[c[0] for c in colunas], show="headings", height=alturas)
    for cid, titulo, largura in colunas:
        arvore.heading(cid, text=titulo)
        arvore.column(cid, width=largura, stretch=cid == colunas[-1][0])
    for tag, cor in (("ok", COR_OK), ("erro", COR_ERRO), ("aviso", COR_AVISO), ("apagado", COR_APAGADO)):
        arvore.tag_configure(tag, foreground=cor)
    rolagem = ttk.Scrollbar(quadro, orient="vertical", command=arvore.yview)
    arvore.configure(yscrollcommand=rolagem.set)
    arvore.pack(side="left", fill="both", expand=True)
    rolagem.pack(side="right", fill="y")
    return quadro, arvore


def linha_de_arquivo(pai, rotulo, valor, ao_trocar, tipos):
    linha = ctk.CTkFrame(pai, fg_color="transparent")
    ctk.CTkLabel(linha, text=rotulo, width=150, anchor="w").pack(side="left")
    var = ctk.StringVar(value=valor)
    ctk.CTkEntry(linha, textvariable=var).pack(side="left", fill="x", expand=True, padx=(0, 8))

    def escolher():
        inicial = os.path.dirname(var.get()) if var.get() else ""
        c = filedialog.askopenfilename(initialdir=inicial or None, filetypes=tipos) if tipos else \
            filedialog.askdirectory(initialdir=var.get() or None)
        if c:
            var.set(os.path.normpath(c))
            if ao_trocar:
                ao_trocar()
    ctk.CTkButton(linha, text="Escolher…", width=96, command=escolher).pack(side="left")
    linha.pack(fill="x", pady=3)
    return var


# ---------------------------------------------------------------------------
# Aba 1 — Modelo
# ---------------------------------------------------------------------------

class AbaModelo:
    def __init__(self, app, pai):
        self.app = app
        locais = app.locais
        topo = ctk.CTkFrame(pai, fg_color="transparent")
        topo.pack(fill="x", padx=16, pady=(12, 4))
        ctk.CTkLabel(topo, justify="left", anchor="w", wraplength=900, text=(
            "Monta o modelo do certificado de dois jeitos: como o AHK faz (pelo PowerPoint) e como o programa novo faz "
            "(sem o PowerPoint). Depois compara a posição de cada texto, do logo e da assinatura. "
            "Diferença de até 2 pontos (menos de 1 mm) conta como igual.")).pack(fill="x")
        self.modelo = linha_de_arquivo(pai, "Modelo (.pptx)", locais.modelo, None, [("PowerPoint", "*.pptx")])
        self.assinatura = linha_de_arquivo(pai, "Assinatura", self._uma_assinatura(locais.assinaturas), None,
                                           [("Imagens", "*.png *.jpg *.jpeg")])
        self.logo = linha_de_arquivo(pai, "Logo", locais.logo, None, [("Imagens", "*.png *.jpg *.jpeg")])
        for w in pai.winfo_children()[1:4]:
            w.pack_configure(padx=16)
        linha = ctk.CTkFrame(pai, fg_color="transparent")
        linha.pack(fill="x", padx=16, pady=3)
        ctk.CTkLabel(linha, text="Coordenador de teste", width=150, anchor="w").pack(side="left")
        self.coord = ctk.StringVar(value=self._nome_da_assinatura())
        ctk.CTkEntry(linha, textvariable=self.coord).pack(side="left", fill="x", expand=True)

        botoes = ctk.CTkFrame(pai, fg_color="transparent")
        botoes.pack(fill="x", padx=16, pady=8)
        self.botao = ctk.CTkButton(botoes, text="Comparar modelo", command=self.comparar,
                                   font=(FONTE, 14, "bold"), height=38)
        self.botao.pack(side="left")
        self.botao_abrir = ctk.CTkButton(botoes, text="Abrir as duas versões", command=self.abrir, state="disabled")
        self.botao_abrir.pack(side="left", padx=8)
        self.resumo = ctk.CTkLabel(botoes, text="", font=(FONTE, 13, "bold"))
        self.resumo.pack(side="left", padx=12)

        self.diagnostico = ctk.CTkLabel(pai, text="", anchor="w", justify="left", text_color=COR_APAGADO)
        self.diagnostico.pack(fill="x", padx=16)
        quadro, self.arvore = tabela(pai, [("forma", "Forma", 330), ("ahk", "AHK (esq, topo, larg, alt)", 230),
                                           ("py", "Python", 230), ("res", "Resultado", 100)])
        quadro.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        self.arquivos = None

    @staticmethod
    def _uma_assinatura(pasta):
        try:
            for n in sorted(os.listdir(pasta)):
                if n.lower().endswith((".png", ".jpg", ".jpeg")) and not n.startswith("_"):
                    return os.path.join(pasta, n)
        except OSError:
            pass
        return ""

    def _nome_da_assinatura(self):
        base = os.path.splitext(os.path.basename(self.assinatura.get()))[0]
        nome = textos.chave_nome(base).title() if base else "Ana Beatriz Souza"
        return f"Dra. {nome}"

    def comparar(self):
        self.botao.configure(state="disabled", text="Comparando…")
        self.arvore.delete(*self.arvore.get_children())
        coord = self.coord.get().strip()
        titulo = coord.split(" ")[0] if coord.lower().startswith("dr") else ""
        args = (self.modelo.get(), coord, textos.cargo_por_titulo(textos.normalizar_titulo(titulo)),
                datetime.now().strftime("%d/%m/%Y"), self.assinatura.get(), self.logo.get())
        em_segundo_plano(self.app, lambda: referencia_ahk.comparar_modelo(*args), self._mostrar,
                         lambda e, d: (self.app.mostrar_erro(e, d), self.botao.configure(state="normal", text="Comparar modelo")))

    def _mostrar(self, resultado):
        linhas, pior, self.arquivos, diagnostico = resultado
        for forma, va, vp, ok in linhas:
            self.arvore.insert("", "end", values=(forma, va or "—", vp or "—", "igual" if ok else "DIFERENTE"),
                               tags=("ok" if ok else "erro",))
        ok = all(l[3] for l in linhas)
        self.resumo.configure(text=(f"Igual ao AHK (maior diferença {pior:.1f} pt)" if ok
                                    else f"Há diferença (maior: {pior:.1f} pt)"),
                              text_color=COR_OK if ok else COR_ERRO)
        self.botao.configure(state="normal", text="Comparar modelo")
        self.botao_abrir.configure(state="normal")
        self.diagnostico.configure(text="\n".join(diagnostico))
        self.app.registrar_resultado("Modelo", f"{'igual' if ok else 'DIFERENTE'} (maior diferença {pior:.1f} pt)"
                                     + "".join(f" | {d}" for d in diagnostico))

    def abrir(self):
        if self.arquivos:
            for a in self.arquivos:
                os.startfile(a)


# ---------------------------------------------------------------------------
# Aba 2 — Simulação
# ---------------------------------------------------------------------------

class AbaSimulacao:
    def __init__(self, app, pai):
        self.app = app
        self.cursos = []          # planilhas.Curso
        self.marcados = {}        # índice -> BooleanVar
        self.lote = []            # gerador.CursoNoLote
        self.meses = {}           # nome da aba -> BooleanVar

        ctk.CTkLabel(pai, justify="left", anchor="w", wraplength=900, text=(
            "Gera os certificados de cursos reais só neste computador, em C:\\Certificados_SIMULACAO. "
            "Nada vai para a nuvem e nenhum e-mail é enviado. Para comparar com o AHK, escolha cursos e meses "
            "que o AHK já gerou (eles ficam em C:\\Certificados).")).pack(fill="x", padx=16, pady=(12, 4))

        self.mestre = linha_de_arquivo(pai, "Planilha mestre", app.locais.planilha_mestre, self.carregar_abas,
                                       [("Excel", "*.xlsx")])
        pai.winfo_children()[-1].pack_configure(padx=16)

        linha = ctk.CTkFrame(pai, fg_color="transparent")
        linha.pack(fill="x", padx=16, pady=3)
        ctk.CTkLabel(linha, text="Semestre (aba)", width=150, anchor="w").pack(side="left")
        self.aba = ctk.CTkComboBox(linha, values=[], width=160, command=lambda _: self.carregar_cursos())
        self.aba.pack(side="left")
        ctk.CTkLabel(linha, text="Data de emissão", anchor="w").pack(side="left", padx=(24, 8))
        self.data = ctk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        ctk.CTkEntry(linha, textvariable=self.data, width=110).pack(side="left")
        ctk.CTkLabel(linha, text="use a mesma data que o AHK usou, se quiser comparar", text_color=COR_APAGADO,
                     font=(FONTE, 11)).pack(side="left", padx=8)

        corpo = ctk.CTkFrame(pai, fg_color="transparent")
        corpo.pack(fill="both", expand=True, padx=16, pady=6)
        corpo.grid_columnconfigure(0, weight=1)
        corpo.grid_columnconfigure(1, weight=2)
        corpo.grid_rowconfigure(1, weight=1)

        cab = ctk.CTkFrame(corpo, fg_color="transparent")
        cab.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        ctk.CTkLabel(cab, text="Cursos", font=(FONTE, 13, "bold")).pack(side="left")
        self.filtro = ctk.StringVar()
        self.filtro.trace_add("write", lambda *_: self._desenhar_cursos())
        ctk.CTkEntry(cab, textvariable=self.filtro, placeholder_text="filtrar…", width=150).pack(side="right")
        self.lista_cursos = ctk.CTkScrollableFrame(corpo)
        self.lista_cursos.grid(row=1, column=0, sticky="nsew", padx=(0, 8))

        cab2 = ctk.CTkFrame(corpo, fg_color="transparent")
        cab2.grid(row=0, column=1, sticky="ew")
        ctk.CTkLabel(cab2, text="Preparação", font=(FONTE, 13, "bold")).pack(side="left")
        self.botao_preparar = ctk.CTkButton(cab2, text="Preparar cursos marcados", command=self.preparar)
        self.botao_preparar.pack(side="right")
        direita = ctk.CTkFrame(corpo, fg_color="transparent")
        direita.grid(row=1, column=1, sticky="nsew")
        quadro, self.arvore = tabela(direita, [("curso", "Curso", 230), ("coord", "Assina", 200),
                                               ("planilha", "Planilha do curso", 260)], alturas=8)
        quadro.pack(fill="both", expand=True)
        self.arvore.bind("<Double-1>", self._trocar_planilha)
        ctk.CTkLabel(direita, text="Duplo clique numa linha para escolher outra planilha do curso.",
                     text_color=COR_APAGADO, font=(FONTE, 11)).pack(anchor="w")
        self.quadro_meses = ctk.CTkFrame(direita, fg_color="transparent")
        self.quadro_meses.pack(fill="x", pady=(6, 0))

        rodape = ctk.CTkFrame(pai, fg_color="transparent")
        rodape.pack(fill="x", padx=16, pady=(4, 14))
        self.botao_gerar = ctk.CTkButton(rodape, text="Gerar simulação", state="disabled", command=self.gerar,
                                         font=(FONTE, 14, "bold"), height=38)
        self.botao_gerar.pack(side="left")
        self.barra = ctk.CTkProgressBar(rodape)
        self.barra.set(0)
        self.barra.pack(side="left", fill="x", expand=True, padx=12)
        self.status = ctk.CTkLabel(rodape, text="", width=360, anchor="w")
        self.status.pack(side="left")

        self.carregar_abas()

    # -- mestre e cursos ----------------------------------------------------

    def carregar_abas(self):
        caminho = self.mestre.get()
        if not os.path.isfile(caminho):
            self.status.configure(text="Escolha a planilha mestre.")
            return
        em_segundo_plano(self.app, lambda: planilhas.listar_abas(caminho), self._abas_carregadas)

    def _abas_carregadas(self, abas):
        self.aba.configure(values=abas)
        semestres = [a for a in abas if any(ch.isdigit() for ch in a)]
        if semestres:
            self.aba.set(semestres[-1])
            self.carregar_cursos()

    def carregar_cursos(self):
        caminho, aba = self.mestre.get(), self.aba.get()
        em_segundo_plano(self.app, lambda: planilhas.ler_cursos_da_mestre(caminho, aba), self._cursos_carregados)

    def _cursos_carregados(self, cursos):
        self.cursos = cursos or []
        self.marcados = {i: ctk.BooleanVar(value=False) for i in range(len(self.cursos))}
        self._desenhar_cursos()
        self.status.configure(text=f"{len(self.cursos)} cursos na aba {self.aba.get()}.")

    def _desenhar_cursos(self):
        for w in self.lista_cursos.winfo_children():
            w.destroy()
        filtro = textos.normalizar(self.filtro.get())
        for i, c in enumerate(self.cursos):
            if filtro and filtro not in textos.normalizar(f"{c.posgrad} {c.unidade}"):
                continue
            ctk.CTkCheckBox(self.lista_cursos, text=f"{c.posgrad}  ·  {c.unidade_exibida}",
                            variable=self.marcados[i]).pack(anchor="w", pady=2)

    # -- preparação ---------------------------------------------------------

    def preparar(self):
        escolhidos = [c for i, c in enumerate(self.cursos) if self.marcados[i].get()]
        if not escolhidos:
            self.status.configure(text="Marque pelo menos um curso.")
            return
        self.botao_preparar.configure(state="disabled", text="Preparando…")
        aba, locais, log = self.aba.get(), self.app.locais, self.app.log_tela

        def trabalho():
            cad = coordenadores.Cadastro(locais.assinaturas, log=log)
            lote = []
            for c in escolhidos:
                item = gerador.CursoNoLote(c)
                item.coordenacao = coordenadores.resolver(c.coordenacao, cad, log=log)
                if item.coordenacao.pendente == "assinatura":
                    item.removido = "assinatura do coordenador não encontrada (cadastre no AHK primeiro)"
                elif item.coordenacao.pendente == "titulo":
                    item.removido = "Dr./Dra. do coordenador ainda não cadastrado (cadastre no AHK primeiro)"
                item.planilha = caminhos.localizar_planilha_do_curso(locais.raiz_onedrive, c.posgrad, c.unidade, aba)
                lote.append(item)
            abas = {item.planilha: planilhas.listar_abas(item.planilha) for item in lote if item.planilha}
            return lote, abas

        em_segundo_plano(self.app, trabalho, self._preparado)

    def _preparado(self, resultado):
        self.lote, self.abas_por_planilha = resultado
        self.botao_preparar.configure(state="normal", text="Preparar cursos marcados")
        self._desenhar_lote()

    def _desenhar_lote(self):
        self.arvore.delete(*self.arvore.get_children())
        for i, item in enumerate(self.lote):
            coord = item.coordenacao.assina_como if not item.removido else item.removido
            planilha = os.path.basename(item.planilha) if item.planilha else "NÃO ENCONTRADA — duplo clique"
            tag = "erro" if item.removido or not item.planilha else "ok"
            self.arvore.insert("", "end", iid=str(i), values=(item.curso.posgrad, coord, planilha), tags=(tag,))
        for w in self.quadro_meses.winfo_children():
            w.destroy()
        todas = []
        for item in self.lote:
            for a in self.abas_por_planilha.get(item.planilha, []):
                if a not in todas:
                    todas.append(a)
        anteriores = {k: v.get() for k, v in self.meses.items()}
        self.meses = {a: ctk.BooleanVar(value=anteriores.get(a, False)) for a in todas}
        if todas:
            ctk.CTkLabel(self.quadro_meses, text="Meses (abas das planilhas):", anchor="w").pack(anchor="w")
            grade = ctk.CTkFrame(self.quadro_meses, fg_color="transparent")
            grade.pack(fill="x")
            for n, a in enumerate(todas):
                ctk.CTkCheckBox(grade, text=a, variable=self.meses[a], width=120).grid(row=n // 6, column=n % 6,
                                                                                       sticky="w", pady=2)
        prontos = [i for i in self.lote if not i.removido and i.planilha]
        self.botao_gerar.configure(state="normal" if prontos else "disabled")
        self.status.configure(text=f"{len(prontos)} de {len(self.lote)} cursos prontos. Marque os meses e gere.")

    def _trocar_planilha(self, evento):
        linha = self.arvore.identify_row(evento.y)
        if not linha:
            return
        item = self.lote[int(linha)]
        c = filedialog.askopenfilename(title=f"Planilha de {item.curso.posgrad}", filetypes=[("Excel", "*.xlsx")],
                                       initialdir=os.path.dirname(item.planilha) or self.app.locais.raiz_onedrive or None)
        if not c:
            return
        item.planilha = os.path.normpath(c)

        def trabalho():
            return planilhas.listar_abas(item.planilha)
        em_segundo_plano(self.app, trabalho,
                         lambda abas: (self.abas_por_planilha.__setitem__(item.planilha, abas), self._desenhar_lote()))

    # -- geração ------------------------------------------------------------

    def gerar(self):
        meses = [a for a, v in self.meses.items() if v.get()]
        if not meses:
            self.status.configure(text="Marque pelo menos um mês.")
            return
        data = self.data.get().strip()
        if not horas.data_valida(data):
            self.status.configure(text="Data de emissão inválida (use dd/mm/aaaa).")
            return
        self.botao_gerar.configure(state="disabled", text="Gerando…")
        self.barra.set(0)
        locais, modo = self.app.locais, gerador.MODOS["simulacao"]
        lote = [i for i in self.lote if not i.removido and i.planilha]

        def progresso(texto, feitos, total):
            self.app.after(0, lambda: (self.barra.set(feitos / max(total, 1)),
                                       self.status.configure(text=f"{feitos}/{total} · {texto[:60]}")))

        def trabalho():
            from certificados.pdf import ConversorPowerPoint
            log = gerador.Log("simulacao", ecoar=self.app.log_tela)
            for item in lote:
                item.abas = [m for m in meses if planilhas.aba_na_lista(self.abas_por_planilha.get(item.planilha, []), m)]
                if not item.abas:
                    item.removido = "a planilha não tem nenhum dos meses marcados"
                    continue
                item.removido = ""
                item.dados = planilhas.ler_dados_da_planilha(item.planilha, item.abas)
                for aba in item.abas:      # simulação anterior do mesmo curso/mês sai antes
                    pasta = os.path.join(modo.pasta_local, item.nome_com_unidade, textos.mes_para_pasta(aba))
                    shutil.rmtree(pasta, ignore_errors=True)
            problemas, _, _ = gerador.validar(lote)
            for p in problemas:            # igual a escolher "Pular" na tela de problemas do AHK
                p.registro.pular = p.registro.tem_erro
            with ConversorPowerPoint(log=log) as conv:
                rel = gerador.gerar(lote, locais, self.app.aba_modelo.modelo.get(), data, modo, conv, log,
                                    progresso=progresso)
            return rel, problemas

        em_segundo_plano(self.app, trabalho, self._gerado,
                         lambda e, d: (self.app.mostrar_erro(e, d),
                                       self.botao_gerar.configure(state="normal", text="Gerar simulação")))

    def _gerado(self, resultado):
        rel, problemas = resultado
        self.botao_gerar.configure(state="normal", text="Gerar simulação")
        self.barra.set(1)
        falhas = f", {len(rel.falhas)} falha(s)" if rel.falhas else ""
        self.status.configure(text=f"{rel.total_gerados} PDFs gerados{falhas}. Veja a aba Comparar.")
        self.app.registrar_resultado("Simulação", f"{rel.total_gerados} PDFs, {len(problemas)} linhas com problema, "
                                                  f"{len(rel.falhas)} falhas. Log: {rel.log}")
        for p in problemas:
            self.app.log_tela(f"Problema: {p.curso.curso.posgrad} / {p.registro.aba} / linha {p.registro.linha}: "
                              + "; ".join(p.descricao()))
        for f in rel.falhas:
            self.app.log_tela("FALHA: " + f)
        self.app.aba_comparar.comparar_simulacao(rel)


# ---------------------------------------------------------------------------
# Aba 3 — Comparar
# ---------------------------------------------------------------------------

class AbaComparar:
    def __init__(self, app, pai):
        self.app = app
        ctk.CTkLabel(pai, justify="left", anchor="w", wraplength=900, text=(
            "Compara o texto de dentro de cada PDF. Depois de uma simulação, compara sozinho cada curso/mês com a "
            "pasta equivalente em C:\\Certificados. Também dá para escolher duas pastas à mão.")).pack(
            fill="x", padx=16, pady=(12, 4))
        self.pasta_ahk = linha_de_arquivo(pai, "Pasta do AHK", gerador.MODOS["oficial"].pasta_local, None, None)
        pai.winfo_children()[-1].pack_configure(padx=16)
        self.pasta_py = linha_de_arquivo(pai, "Pasta do Python", gerador.MODOS["simulacao"].pasta_local, None, None)
        pai.winfo_children()[-1].pack_configure(padx=16)
        botoes = ctk.CTkFrame(pai, fg_color="transparent")
        botoes.pack(fill="x", padx=16, pady=6)
        self.botao = ctk.CTkButton(botoes, text="Comparar estas duas pastas", command=self.comparar_manual)
        self.botao.pack(side="left")
        self.resumo = ctk.CTkLabel(botoes, text="", font=(FONTE, 13, "bold"))
        self.resumo.pack(side="left", padx=12)
        quadro, self.arvore = tabela(pai, [("grupo", "Curso / mês", 260), ("arquivo", "Arquivo", 260),
                                           ("sit", "Resultado", 110), ("det", "Diferença", 420)], alturas=16)
        quadro.pack(fill="both", expand=True, padx=16, pady=(0, 16))

    def _comparar(self, pares):
        def trabalho():
            saida = []
            for grupo, pa, pp in pares:
                if not os.path.isdir(pa):
                    saida.append((grupo, comparacao.Linha("—", "sem AHK", f"o AHK não tem a pasta {pa}")))
                    continue
                datas = set(comparacao.datas_de_emissao(pa)) ^ set(comparacao.datas_de_emissao(pp))
                for linha in comparacao.comparar_pastas(pa, pp, ignorar=sorted(datas)):
                    saida.append((grupo, linha))
            return saida
        self.arvore.delete(*self.arvore.get_children())
        self.resumo.configure(text="Comparando…", text_color=COR_APAGADO)
        em_segundo_plano(self.app, trabalho, self._mostrar)

    def _mostrar(self, linhas):
        cont = {}
        for grupo, l in linhas:
            cont[l.situacao] = cont.get(l.situacao, 0) + 1
            tag = {"igual": "ok", "diferente": "erro"}.get(l.situacao, "aviso")
            self.arvore.insert("", "end", values=(grupo, l.arquivo, l.situacao, l.detalhe), tags=(tag,))
        iguais = cont.get(comparacao.IGUAL, 0)
        total = sum(cont.values())
        texto = f"{iguais} de {total} iguais" + "".join(f" · {v} {k}" for k, v in cont.items() if k != comparacao.IGUAL)
        self.resumo.configure(text=texto, text_color=COR_OK if iguais == total and total else COR_AVISO)
        self.app.registrar_resultado("Comparação", texto)

    def comparar_manual(self):
        self._comparar([("(escolhidas à mão)", self.pasta_ahk.get(), self.pasta_py.get())])

    def comparar_simulacao(self, rel):
        pares = []
        for r in rel.cursos:
            for pasta_py in r.pastas:
                relativo = os.path.relpath(pasta_py, gerador.MODOS["simulacao"].pasta_local)
                pares.append((relativo, os.path.join(gerador.MODOS["oficial"].pasta_local, relativo), pasta_py))
        self.app.abas.set("3. Comparar")
        self._comparar(pares)


# ---------------------------------------------------------------------------
# Janela
# ---------------------------------------------------------------------------

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        self.title(f"Certificados em Python — Conferência (motor {VERSAO})")
        self.geometry("1180x800")
        self.minsize(980, 680)
        # CERTIFICADOS_RAIZ aponta para outra "raiz do OneDrive" (usado nos testes).
        self.locais = caminhos.Locais.descobrir(os.environ.get("CERTIFICADOS_RAIZ") or None)
        self.resultados = []

        estilo = ttk.Style(self)
        estilo.theme_use("clam")
        estilo.configure("Treeview", background="#22272e", fieldbackground="#22272e", foreground="#d0d7de",
                         rowheight=24, font=(FONTE, 10), borderwidth=0)
        estilo.configure("Treeview.Heading", background="#2d333b", foreground="#d0d7de", font=(FONTE, 10, "bold"))
        estilo.map("Treeview", background=[("selected", "#316dca")])

        cab = ctk.CTkFrame(self, fg_color="transparent")
        cab.pack(fill="x", padx=16, pady=(12, 0))
        ctk.CTkLabel(cab, text="Conferência do programa novo", font=(FONTE, 20, "bold")).pack(side="left")
        onde = self.locais.declaracoes if os.path.isdir(self.locais.declaracoes or "") else "pasta Declarações NÃO encontrada"
        ctk.CTkLabel(cab, text=onde, text_color=COR_APAGADO).pack(side="right")

        self.abas = ctk.CTkTabview(self)
        self.abas.pack(fill="both", expand=True, padx=12, pady=(4, 0))
        for nome in ("1. Modelo", "2. Simulação", "3. Comparar", "Registro"):
            self.abas.add(nome)
        self.aba_modelo = AbaModelo(self, self.abas.tab("1. Modelo"))
        self.texto_log = ctk.CTkTextbox(self.abas.tab("Registro"), font=("Consolas", 11))
        self.texto_log.pack(fill="both", expand=True, padx=12, pady=12)
        botoes_log = ctk.CTkFrame(self.abas.tab("Registro"), fg_color="transparent")
        botoes_log.pack(fill="x", padx=12, pady=(0, 12))
        ctk.CTkButton(botoes_log, text="Copiar resumo dos resultados", command=self.copiar_resumo).pack(side="left")
        self.aba_comparar = AbaComparar(self, self.abas.tab("3. Comparar"))
        self.aba_simulacao = AbaSimulacao(self, self.abas.tab("2. Simulação"))

    def log_tela(self, msg):
        self.after(0, lambda: (self.texto_log.insert("end", f"{datetime.now():%H:%M:%S}  {msg}\n"),
                               self.texto_log.see("end")))

    def registrar_resultado(self, etapa, texto):
        self.resultados.append(f"{etapa}: {texto}")
        self.log_tela(f"== {etapa}: {texto}")

    def copiar_resumo(self):
        texto = "\n".join([f"Conferência motor {VERSAO} — {datetime.now():%d/%m/%Y %H:%M}"] + self.resultados)
        self.clipboard_clear()
        self.clipboard_append(texto)

    def mostrar_erro(self, erro, detalhe=""):
        self.log_tela(f"ERRO: {erro}\n{detalhe}")
        janela = ctk.CTkToplevel(self)
        janela.title("Erro")
        janela.geometry("620x320")
        janela.transient(self)
        ctk.CTkLabel(janela, text="Algo deu errado", font=(FONTE, 16, "bold")).pack(anchor="w", padx=16, pady=(14, 4))
        caixa = ctk.CTkTextbox(janela, font=("Consolas", 11))
        caixa.pack(fill="both", expand=True, padx=16, pady=4)
        caixa.insert("end", f"{erro}\n\n{detalhe}")
        ctk.CTkButton(janela, text="Fechar", command=janela.destroy).pack(pady=10)


if __name__ == "__main__":
    App().mainloop()
