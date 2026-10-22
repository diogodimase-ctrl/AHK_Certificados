"""
Cores e fontes do programa (estilo "Escuro violeta", escolhido em 21/10/2026).

Todas as cores ficam aqui: as telas usam estes nomes e nunca escrevem cor na
mão. tema.aplicar() roda antes de criar a janela e preenche o tema do
CustomTkinter, então botões, campos e caixinhas já nascem nas cores certas.
"""

from tkinter import ttk

import customtkinter as ctk

FONTE = "Segoe UI"
TAMANHO = 15                      # texto normal (antes era 13)

FUNDO = "#1e1b2e"
LATERAL = "#16131f"
CARTAO = "#262236"                # listas, tabelas e campos
BORDA = "#3d3756"
LINHA = "#332e47"                 # divisória entre linhas de uma lista

TEXTO = "#ece9f6"
TEXTO_2 = "#b3abc9"               # explicações
APAGADO = "#8f88a8"               # etapas futuras, dicas
TITULO = "#c9bbff"

ACENTO = "#6d4cf0"                # botões principais, etapa atual, caixinhas
ACENTO_HOVER = "#5a3bd6"
SECUNDARIO_HOVER = "#2e2942"

OK = "#7ee8bf"
OK_FUNDO = "#183d33"
AVISO = "#f5cc72"
AVISO_FUNDO = "#44361a"
ERRO = "#ff8f87"
ERRO_FUNDO = "#4a2226"
UNIDADE = "#cdc1ff"
UNIDADE_FUNDO = "#3a2f66"
CERTIFICADOS = "#9d85ff"          # rosca dos certificados
MARCADO = "#2e2748"               # linha marcada na tabela de cursos
MARCADO_HOVER = "#3a3260"

FAIXA_TESTE, FAIXA_TESTE_TEXTO = "#e9b949", "#1c1c1c"
FAIXA_OFICIAL, FAIXA_OFICIAL_TEXTO = "#c23a31", "#ffffff"

# Botão com contorno (Voltar, ações secundárias).
BOTAO_CONTORNO = dict(fg_color="transparent", border_width=1, border_color="#4d4669", hover_color=SECUNDARIO_HOVER,
                      text_color=TEXTO)


def fonte(tamanho=TAMANHO, negrito=False):
    return (FONTE, tamanho, "bold" if negrito else "normal")


def aplicar():
    """Preenche o tema do CustomTkinter (cada cor vale para claro e escuro)."""
    ctk.set_appearance_mode("dark")
    t = ctk.ThemeManager.theme

    def cor(widget, **cores):
        for chave, valor in cores.items():
            t[widget][chave] = [valor, valor] if isinstance(valor, str) and valor != "transparent" else valor

    cor("CTk", fg_color=FUNDO)
    cor("CTkToplevel", fg_color=FUNDO)
    cor("CTkFrame", fg_color=CARTAO, top_fg_color=CARTAO, border_color=BORDA)
    t["CTkFrame"]["corner_radius"] = 12
    cor("CTkButton", fg_color=ACENTO, hover_color=ACENTO_HOVER, border_color=BORDA, text_color="#ffffff",
        text_color_disabled=APAGADO)
    t["CTkButton"]["corner_radius"] = 10
    cor("CTkLabel", text_color=TEXTO)
    cor("CTkEntry", fg_color=CARTAO, border_color=BORDA, text_color=TEXTO, placeholder_text_color=APAGADO)
    t["CTkEntry"]["corner_radius"] = 10
    cor("CTkCheckBox", fg_color=ACENTO, border_color="#6c6489", hover_color=ACENTO_HOVER, checkmark_color="#ffffff",
        text_color=TEXTO, text_color_disabled=APAGADO)
    cor("CTkComboBox", fg_color=CARTAO, border_color=BORDA, button_color=BORDA, button_hover_color=ACENTO,
        text_color=TEXTO, text_color_disabled=APAGADO)
    t["CTkComboBox"]["corner_radius"] = 10
    cor("CTkOptionMenu", fg_color=CARTAO, button_color=BORDA, button_hover_color=ACENTO, text_color=TEXTO,
        text_color_disabled=APAGADO)
    cor("CTkSegmentedButton", fg_color=CARTAO, selected_color=ACENTO, selected_hover_color=ACENTO_HOVER,
        unselected_color=CARTAO, unselected_hover_color=SECUNDARIO_HOVER, text_color="#ffffff",
        text_color_disabled=APAGADO)
    cor("CTkProgressBar", fg_color=CARTAO, progress_color=ACENTO, border_color=BORDA)
    cor("CTkScrollbar", fg_color="transparent", button_color=BORDA, button_hover_color=APAGADO)
    cor("CTkTextbox", fg_color=CARTAO, border_color=BORDA, text_color=TEXTO, scrollbar_button_color=BORDA,
        scrollbar_button_hover_color=APAGADO)
    cor("CTkScrollableFrame", label_fg_color=CARTAO)
    cor("DropdownMenu", fg_color=CARTAO, hover_color=ACENTO, text_color=TEXTO)
    t["CTkFont"]["family"] = FONTE
    t["CTkFont"]["size"] = TAMANHO


def pintar_tabelas(janela):
    """Tabelas (ttk.Treeview) nas cores do tema."""
    estilo = ttk.Style(janela)
    estilo.theme_use("clam")
    estilo.configure("Treeview", background=CARTAO, fieldbackground=CARTAO, foreground=TEXTO, rowheight=32,
                     font=(FONTE, 12), borderwidth=0, bordercolor=CARTAO, lightcolor=CARTAO, darkcolor=CARTAO)
    estilo.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])     # sem a moldura branca
    estilo.configure("Treeview.Heading", background=LATERAL, foreground=TEXTO_2, font=(FONTE, 12, "bold"),
                     borderwidth=0, relief="flat", padding=(6, 6))
    estilo.map("Treeview", background=[("selected", ACENTO)], foreground=[("selected", "#ffffff")])
    estilo.map("Treeview.Heading", background=[("active", SECUNDARIO_HOVER)])
    estilo.configure("Vertical.TScrollbar", background=BORDA, troughcolor=CARTAO, borderwidth=0, arrowcolor=TEXTO_2)


def tabela(pai, colunas, alturas=12):
    """ttk.Treeview com barra de rolagem; colunas = [(id, título, largura)]."""
    quadro = ctk.CTkFrame(pai, fg_color="transparent")
    arvore = ttk.Treeview(quadro, columns=[c[0] for c in colunas], show="headings", height=alturas)
    for cid, titulo, largura in colunas:
        arvore.heading(cid, text=titulo, anchor="w")
        arvore.column(cid, width=largura, stretch=cid == colunas[-1][0], anchor="w")
    for tag, cor_ in (("ok", OK), ("erro", ERRO), ("aviso", AVISO), ("apagado", APAGADO)):
        arvore.tag_configure(tag, foreground=cor_)
    rolagem = ttk.Scrollbar(quadro, orient="vertical", command=arvore.yview)
    arvore.configure(yscrollcommand=rolagem.set)
    arvore.pack(side="left", fill="both", expand=True)
    rolagem.pack(side="right", fill="y")
    return quadro, arvore


def quebrar_na_largura(rotulo, margem=24):
    """Texto longo que quebra na largura disponível (em vez de um número fixo).
    Usa a largura de quem contém o rótulo (que não depende do texto); medir o
    próprio rótulo entra em ciclo: o texto muda a largura, que muda o texto."""
    pai = rotulo.master
    escala = ctk.ScalingTracker.get_widget_scaling(rotulo)
    ultimo = {"largura": 0}

    def ajustar(evento):
        largura = max(200, int(evento.width / escala) - margem)
        if abs(largura - ultimo["largura"]) >= 8 and rotulo.winfo_exists():
            ultimo["largura"] = largura
            rotulo.configure(wraplength=largura)
    pai.bind("<Configure>", ajustar, add="+")
    if pai.winfo_width() > 1:                       # a área já tem tamanho: ajusta agora
        ajustar(type("Evento", (), {"width": pai.winfo_width()})())
    return rotulo


def chip(pai, texto, cor_texto, cor_fundo, tamanho=13):
    """Etiqueta colorida (unidade, mês enviado...)."""
    return ctk.CTkLabel(pai, text=texto, text_color=cor_texto, fg_color=cor_fundo, corner_radius=6,
                        font=fonte(tamanho), height=26)
