"""
Peças desenhadas à mão (tk.Canvas) da janela: etapas na horizontal, gráficos
de rosca, barras do tempo economizado e a tabela de cursos com cara de Excel.

O Canvas não acompanha sozinho a escala do Windows (125%, 150%...) como os
widgets do CustomTkinter, então toda medida passa por px() e as fontes são em
pixels (tamanho negativo).
"""

import tkinter as tk
from tkinter import font as tkfont

import customtkinter as ctk

import tema
from certificados.economia import duracao
from certificados.painel import curto
from tema import FONTE


class _Desenho(ctk.CTkFrame):
    """Frame com um Canvas que se redesenha quando muda de tamanho."""

    def __init__(self, pai, altura, fundo, **kw):
        super().__init__(pai, fg_color=fundo, **kw)
        self.escala = ctk.ScalingTracker.get_widget_scaling(self)
        self.fundo = fundo
        self.canvas = tk.Canvas(self, height=self.px(altura), bg=fundo, highlightthickness=0, bd=0)
        self.canvas.pack(fill="both", expand=True, padx=self.px(2), pady=self.px(2))
        self.canvas.bind("<Configure>", lambda e: self.desenhar())
        self._fontes = {}

    def px(self, valor):
        return int(round(valor * self.escala))

    def fonte(self, tamanho, negrito=False):
        chave = (tamanho, negrito)
        if chave not in self._fontes:
            self._fontes[chave] = tkfont.Font(family=FONTE, size=-self.px(tamanho),
                                              weight="bold" if negrito else "normal")
        return self._fontes[chave]

    def cortar(self, texto, largura, fonte):
        """Texto com "…" no fim se não couber na largura."""
        texto = str(texto or "")
        if largura <= 0:
            return ""
        if fonte.measure(texto) <= largura:
            return texto
        while texto and fonte.measure(texto + "…") > largura:
            texto = texto[:-1]
        return texto + "…"

    def desenhar(self):
        pass


# ---------------------------------------------------------------------------
# Etapas
# ---------------------------------------------------------------------------

class Etapas(_Desenho):
    """Bolinhas numeradas ligadas por uma linha; a atual em violeta e as feitas com ✓."""

    def __init__(self, pai, nomes):
        self.nomes, self.atual = nomes, 0
        super().__init__(pai, 62, tema.LATERAL)

    def marcar(self, n):
        self.atual = n
        self.desenhar()

    def desenhar(self):
        c = self.canvas
        c.delete("all")
        largura = c.winfo_width()
        if largura < 10:
            return
        passo = largura / len(self.nomes)
        r, y = self.px(15), self.px(18)
        for i in range(len(self.nomes) - 1):
            x1, x2 = passo * (i + 0.5) + r + self.px(6), passo * (i + 1.5) - r - self.px(6)
            c.create_line(x1, y, x2, y, fill=tema.ACENTO if i < self.atual else tema.BORDA, width=self.px(2))
        for i, nome in enumerate(self.nomes):
            x = passo * (i + 0.5)
            if i == self.atual:
                fundo, borda, cor, texto = tema.ACENTO, tema.ACENTO, "#ffffff", str(i + 1)
            elif i < self.atual:
                fundo, borda, cor, texto = tema.UNIDADE_FUNDO, tema.ACENTO, tema.TITULO, "✓"
            else:
                fundo, borda, cor, texto = tema.FUNDO, tema.BORDA, tema.APAGADO, str(i + 1)
            c.create_oval(x - r, y - r, x + r, y + r, fill=fundo, outline=borda, width=self.px(2))
            c.create_text(x, y, text=texto, fill=cor, font=self.fonte(14, True))
            c.create_text(x, y + r + self.px(14), text=nome, font=self.fonte(14, i == self.atual),
                          fill=tema.TEXTO if i <= self.atual else tema.APAGADO)


# ---------------------------------------------------------------------------
# Gráfico de rosca
# ---------------------------------------------------------------------------

class Rosca(ctk.CTkFrame):
    """Cartão com rosca de % à esquerda e os números à direita."""

    def __init__(self, pai, titulo, cor):
        super().__init__(pai)
        self.cor, self.fracao = cor, None
        self.escala = ctk.ScalingTracker.get_widget_scaling(self)
        lado = int(100 * self.escala)
        self.canvas = tk.Canvas(self, width=lado, height=lado, bg=tema.CARTAO, highlightthickness=0, bd=0)
        self.canvas.pack(side="left", padx=(16, 14), pady=14)
        textos = ctk.CTkFrame(self, fg_color="transparent")
        textos.pack(side="left", fill="both", expand=True, pady=14, padx=(0, 12))
        ctk.CTkLabel(textos, text=titulo, text_color=tema.TEXTO_2, font=(FONTE, 13), anchor="w").pack(fill="x")
        self.rotulo_numero = ctk.CTkLabel(textos, text="…", font=(FONTE, 26, "bold"), anchor="w")
        self.rotulo_numero.pack(fill="x")
        self.rotulo_sub = ctk.CTkLabel(textos, text="", text_color=tema.TEXTO_2, font=(FONTE, 13), anchor="w",
                                       justify="left")
        self.rotulo_sub.pack(fill="x")
        self._fonte = tkfont.Font(family=FONTE, size=-int(24 * self.escala), weight="bold")
        self.desenhar()

    def mostrar(self, feitos, total, sub="", cor_sub=None):
        self.fracao = feitos / total if total else None
        self.rotulo_numero.configure(text=f"{feitos} de {total}" if total else "—")
        self.rotulo_sub.configure(text=sub, text_color=cor_sub or tema.TEXTO_2)
        self.desenhar()

    def carregando(self, sub="calculando…"):
        self.fracao = None
        self.rotulo_numero.configure(text="…")
        self.rotulo_sub.configure(text=sub, text_color=tema.APAGADO)
        self.desenhar()

    def desenhar(self):
        c = self.canvas
        c.delete("all")
        lado = int(c.cget("width"))
        grossura = int(13 * self.escala)
        m = grossura / 2 + 2
        caixa = (m, m, lado - m, lado - m)
        c.create_oval(*caixa, outline=tema.BORDA, width=grossura)
        if self.fracao:
            extensao = -359.9 * min(self.fracao, 1)
            c.create_arc(*caixa, start=90, extent=extensao, style="arc", outline=self.cor, width=grossura)
        texto = "…" if self.fracao is None else f"{round(self.fracao * 100)}%"
        c.create_text(lado / 2, lado / 2, text=texto, fill=tema.TEXTO, font=self._fonte)


# ---------------------------------------------------------------------------
# Tempo economizado
# ---------------------------------------------------------------------------

class BarrasTempo(_Desenho):
    """Barras deitadas: manual x programa, e o tempo economizado embaixo."""

    def __init__(self, pai):
        self.resumo = None
        super().__init__(pai, 110, tema.CARTAO)

    def mostrar(self, resumo):
        self.resumo = resumo
        self.desenhar()

    def desenhar(self):
        c = self.canvas
        c.delete("all")
        largura = c.winfo_width()
        if largura < 10:
            return
        x0 = self.px(4)
        c.create_text(x0, self.px(10), text="Tempo: manual × programa", anchor="w", fill=tema.TEXTO_2,
                      font=self.fonte(13))
        r = self.resumo
        if r and r.rodadas:
            c.create_text(largura - self.px(4), self.px(10), anchor="e", fill=tema.APAGADO, font=self.fonte(12),
                          text=f"{r.rodadas} rodada{'s' if r.rodadas > 1 else ''} · {r.certificados} certif.")
        if not r or not r.rodadas:
            c.create_text(x0, self.px(52), anchor="w", fill=tema.APAGADO, font=self.fonte(13), width=largura - x0,
                          text="Aparece depois da primeira rodada oficial neste computador.")
            return
        rotulo = self.px(78)
        valor = self.px(92)
        barra_max = max(self.px(30), largura - rotulo - valor - x0)
        maior = max(r.manual, r.programa, 1)
        for i, (nome, segundos, cor) in enumerate((("Manual", r.manual, tema.ERRO), ("Programa", r.programa, tema.OK))):
            y = self.px(34 + i * 26)
            c.create_text(x0, y, text=nome, anchor="w", fill=tema.TEXTO, font=self.fonte(13))
            comprimento = max(self.px(3), barra_max * segundos / maior)
            c.create_rectangle(x0 + rotulo, y - self.px(8), x0 + rotulo + comprimento, y + self.px(8), fill=cor,
                               outline="")
            c.create_text(x0 + rotulo + comprimento + self.px(8), y, text=duracao(segundos), anchor="w",
                          fill=tema.TEXTO_2, font=self.fonte(13))
        y = self.px(92)
        c.create_text(x0, y, anchor="w", fill=tema.OK, font=self.fonte(16, True),
                      text=f"Economizado: {duracao(r.economizado)}")


# ---------------------------------------------------------------------------
# Tabela de cursos (cara de Excel)
# ---------------------------------------------------------------------------

COR_MES = {"ok": tema.OK, "pendente": tema.AVISO, "falta": tema.ERRO, "futuro": tema.BORDA}


class PlanilhaCursos(ctk.CTkFrame):
    """Letras nas colunas, número da linha da mestre à esquerda, linhas de
    grade e, na coluna de meses, uma barrinha por mês do semestre.

    linhas = [{"linha", "marcado", "pos", "coord", "uni", "turma",
               "meses": [estado...], "cert", "cert_cor"}]
    Clique numa linha chama ao_clicar(linha)."""

    # (chave, título, largura fixa ou None = divide o que sobra, peso)
    COLUNAS = [("marca", "", 40, 0), ("pos", "Pós-graduação", None, 3), ("coord", "Coordenação", None, 2),
               ("uni", "Unidade", 92, 0), ("turma", "Turma", 84, 0), ("meses", "", 236, 0),
               ("cert", "Certificados", 118, 0)]
    NUM, LETRAS, TITULOS, LINHA = 54, 24, 34, 34

    def __init__(self, pai, ao_clicar):
        super().__init__(pai, fg_color=tema.FUNDO, corner_radius=0)
        self.escala = ctk.ScalingTracker.get_widget_scaling(self)
        self.ao_clicar = ao_clicar
        self.linhas, self.meses, self.mes_destaque, self.vazio = [], [], "", "Nenhum curso"
        self._fontes, self._sobre = {}, None
        self.cab = tk.Canvas(self, height=self.px(self.LETRAS + self.TITULOS), bg=tema.LATERAL,
                             highlightthickness=0, bd=0)
        self.cab.grid(row=0, column=0, sticky="ew")
        self.corpo = tk.Canvas(self, bg=tema.FUNDO, highlightthickness=0, bd=0,
                               yscrollincrement=self.px(self.LINHA))
        self.corpo.grid(row=1, column=0, sticky="nsew")
        self.rolagem = ctk.CTkScrollbar(self, command=self.corpo.yview)
        self.rolagem.grid(row=0, column=1, rowspan=2, sticky="ns")
        self.corpo.configure(yscrollcommand=self.rolagem.set)
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.corpo.bind("<Configure>", lambda e: self.desenhar())
        self.corpo.bind("<Button-1>", self._clique)
        self.corpo.bind("<Motion>", self._mover)
        self.corpo.bind("<Leave>", lambda e: self._realcar(None))
        for w in (self.corpo, self.cab):
            w.bind("<Enter>", lambda e: self.corpo.bind_all("<MouseWheel>", self._roda))
            w.bind("<Leave>", lambda e: self.corpo.unbind_all("<MouseWheel>"), add="+")

    def px(self, valor):
        return int(round(valor * self.escala))

    def fonte(self, tamanho, negrito=False):
        chave = (tamanho, negrito)
        if chave not in self._fontes:
            self._fontes[chave] = tkfont.Font(family=FONTE, size=-self.px(tamanho),
                                              weight="bold" if negrito else "normal")
        return self._fontes[chave]

    cortar = _Desenho.cortar

    def mostrar(self, linhas, meses, mes_destaque, vazio="Nenhum curso"):
        self.linhas, self.meses, self.mes_destaque, self.vazio = linhas, meses, mes_destaque, vazio
        self._sobre = None
        self.desenhar()

    # -- medidas ------------------------------------------------------------

    def _larguras(self, total):
        fixas = sum(self.px(l) for _, _, l, _ in self.COLUNAS if l)
        pesos = sum(p for _, _, l, p in self.COLUNAS if not l)
        sobra = max(self.px(120) * pesos, total - self.px(self.NUM) - fixas)
        xs, x = {}, self.px(self.NUM)
        for chave, _, larg, peso in self.COLUNAS:
            w = self.px(larg) if larg else sobra * peso / pesos
            xs[chave] = (x, x + w)
            x += w
        return xs, x

    def _x_mes(self, x0, i):
        return x0 + self.px(12) + i * self.px(36)

    # -- desenho ------------------------------------------------------------

    def desenhar(self):
        largura = self.corpo.winfo_width()
        if largura < 10:
            return
        xs, fim = self._larguras(largura)
        self._xs = xs
        self._cabecalho(xs, fim)
        c = self.corpo
        c.delete("all")
        h = self.px(self.LINHA)
        if not self.linhas:
            c.create_text(largura / 2, self.px(60), text=self.vazio, fill=tema.APAGADO, font=self.fonte(15))
            c.configure(scrollregion=(0, 0, largura, self.px(120)))
            return
        f, fb, fp = self.fonte(14), self.fonte(14, True), self.fonte(12)
        for i, l in enumerate(self.linhas):
            y0, y1 = i * h, (i + 1) * h
            ym = (y0 + y1) / 2
            fundo = tema.MARCADO if l["marcado"] else tema.FUNDO
            c.create_rectangle(0, y0, fim, y1, fill=fundo, outline="", tags=("fundo", f"f{i}"))
            c.create_rectangle(0, y0, self.px(self.NUM), y1, fill=tema.LATERAL, outline="")
            c.create_text(self.px(self.NUM) / 2, ym, text=str(l["linha"]), fill=tema.APAGADO, font=fp)
            # caixinha
            xa, xb = xs["marca"]
            lado = self.px(16)
            cx = (xa + xb) / 2
            if l["marcado"]:
                c.create_rectangle(cx - lado / 2, ym - lado / 2, cx + lado / 2, ym + lado / 2, fill=tema.ACENTO,
                                   outline=tema.ACENTO)
                c.create_text(cx, ym, text="✓", fill="#ffffff", font=self.fonte(12, True))
            else:
                c.create_rectangle(cx - lado / 2, ym - lado / 2, cx + lado / 2, ym + lado / 2, outline="#6c6489",
                                   width=self.px(1.5))
            pad = self.px(10)
            for chave, cor, fonte_ in (("pos", tema.TEXTO, f), ("coord", tema.TEXTO_2, f), ("uni", tema.UNIDADE, fb),
                                       ("turma", tema.TEXTO_2, f)):
                a, b = xs[chave]
                c.create_text(a + pad, ym, anchor="w", fill=cor, font=fonte_,
                              text=self.cortar(l[chave], b - a - 2 * pad, fonte_))
            a, _ = xs["meses"]
            for j, estado in enumerate(l["meses"]):
                x = self._x_mes(a, j)
                c.create_rectangle(x, ym - self.px(5), x + self.px(28), ym + self.px(5),
                                   fill=COR_MES.get(estado, tema.BORDA), outline="")
            a, b = xs["cert"]
            c.create_text(b - pad, ym, anchor="e", fill=l.get("cert_cor") or tema.TEXTO_2, font=f, text=l["cert"])
            # grade
            for chave, _, _, _ in self.COLUNAS:
                c.create_line(xs[chave][1], y0, xs[chave][1], y1, fill=tema.LINHA)
            c.create_line(self.px(self.NUM), y0, self.px(self.NUM), y1, fill=tema.LINHA)
            c.create_line(0, y1 - 1, fim, y1 - 1, fill=tema.LINHA)
        c.configure(scrollregion=(0, 0, largura, len(self.linhas) * h))

    def _cabecalho(self, xs, fim):
        c = self.cab
        c.delete("all")
        hl, ht = self.px(self.LETRAS), self.px(self.TITULOS)
        c.create_rectangle(0, 0, fim, hl, fill=tema.LATERAL, outline="")
        c.create_rectangle(0, hl, fim, hl + ht, fill=tema.SECUNDARIO_HOVER, outline="")
        c.create_rectangle(0, hl, self.px(self.NUM), hl + ht, fill=tema.LATERAL, outline="")
        c.create_text(self.px(self.NUM) / 2, hl + ht / 2, text="1", fill=tema.APAGADO, font=self.fonte(12))
        for i, (chave, titulo, _, _) in enumerate(self.COLUNAS):
            a, b = xs[chave]
            c.create_text((a + b) / 2, hl / 2, text=chr(ord("A") + i), fill=tema.APAGADO, font=self.fonte(12))
            c.create_line(b, 0, b, hl + ht, fill=tema.BORDA)
            if chave == "meses":
                for j, mes in enumerate(self.meses):
                    x = self._x_mes(a, j) + self.px(14)
                    destaque = mes == self.mes_destaque
                    c.create_text(x, hl + ht / 2, text=curto(mes), font=self.fonte(13, destaque),
                                  fill=tema.TEXTO if destaque else tema.TEXTO_2)
            elif chave == "cert":
                c.create_text(b - self.px(10), hl + ht / 2, text=titulo, anchor="e", fill=tema.TITULO,
                              font=self.fonte(13, True))
            elif titulo:
                c.create_text(a + self.px(10), hl + ht / 2, text=titulo, anchor="w", fill=tema.TITULO,
                              font=self.fonte(13, True))
        c.create_line(self.px(self.NUM), 0, self.px(self.NUM), hl + ht, fill=tema.BORDA)
        c.create_line(0, hl, fim, hl, fill=tema.BORDA)
        c.create_line(0, hl + ht - 1, fim, hl + ht - 1, fill=tema.BORDA)

    # -- mouse --------------------------------------------------------------

    def _indice(self, evento):
        y = self.corpo.canvasy(evento.y)
        i = int(y // self.px(self.LINHA))
        return i if 0 <= i < len(self.linhas) else None

    def _clique(self, evento):
        i = self._indice(evento)
        if i is not None:
            self.ao_clicar(self.linhas[i]["linha"])

    def _mover(self, evento):
        self._realcar(self._indice(evento))

    def _realcar(self, i):
        if i == self._sobre:
            return
        for j in (self._sobre, i):
            if j is not None and j < len(self.linhas):
                cor = tema.MARCADO if self.linhas[j]["marcado"] else tema.FUNDO
                if j == i:
                    cor = tema.MARCADO_HOVER if self.linhas[j]["marcado"] else tema.SECUNDARIO_HOVER
                self.corpo.itemconfigure(f"f{j}", fill=cor)
        self._sobre = i
        self.corpo.configure(cursor="hand2" if i is not None else "")

    def _roda(self, evento):
        if self.corpo.winfo_exists():
            self.corpo.yview_scroll(int(-evento.delta / 120) * 3, "units")
            self._realcar(None)
