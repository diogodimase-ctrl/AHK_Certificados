#Requires AutoHotkey v2.0

; Garante a existência da pasta Lib e das dependências UIA locais
if !FileExist(A_ScriptDir "\Lib\UIA.ahk") && FileExist(A_Desktop "\AHK\UIA-v2-main\UIA-v2-main\Lib\UIA.ahk") {
    try DirCreate(A_ScriptDir "\Lib")
    try FileCopy(A_Desktop "\AHK\UIA-v2-main\UIA-v2-main\Lib\UIA.ahk", A_ScriptDir "\Lib\UIA.ahk", true)
}
if !FileExist(A_ScriptDir "\Lib\UIA_Browser.ahk") && FileExist(A_Desktop "\AHK\UIA-v2-main\UIA-v2-main\Lib\UIA_Browser.ahk") {
    try DirCreate(A_ScriptDir "\Lib")
    try FileCopy(A_Desktop "\AHK\UIA-v2-main\UIA-v2-main\Lib\UIA_Browser.ahk", A_ScriptDir "\Lib\UIA_Browser.ahk", true)
}

; ════════════════════════════════════════════════════════════════════════
;   UI_Moderna.ahk — Biblioteca de Componentes Visuais Padronizados
;   Tema: Segoe UI | Fundo #F5F7FA | Destaque #0A1E38
; ════════════════════════════════════════════════════════════════════════

; ────────────────────────────────────────────────────────────────────────
; 1. CAIXA DE MENSAGEM / INFORMAÇÃO / ALERTA (Substitui MsgBox simples)
; ────────────────────────────────────────────────────────────────────────
ExibirMensagem(titulo, cabecalho, texto, tipo := "info", textoBotao := "Continuar ➔", largura := 580) {
    g := Gui("+AlwaysOnTop -MaximizeBox", titulo)
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F5F7FA"

    corCabecalho := "0A1E38"
    iconeEmoji := "ℹ️ "
    if (tipo = "erro" || tipo = "error") {
        corCabecalho := "A61B29"
        iconeEmoji := "⚠️ "
    } else if (tipo = "sucesso" || tipo = "success") {
        corCabecalho := "15633B"
        iconeEmoji := "✅ "
    } else if (tipo = "passo" || tipo = "step") {
        corCabecalho := "0A1E38"
        iconeEmoji := "📌 "
    }

    ; Cabeçalho
    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w" (largura - 40) " c" corCabecalho, iconeEmoji cabecalho)

    ; Divisor
    g.Add("Text", "x20 y48 w" (largura - 40) " h2 0x10")

    ; Corpo do texto
    g.SetFont("s10 norm", "Segoe UI")
    g.Add("Text", "x20 y58 w" (largura - 40) " c1A1A1A", texto)

    ; Botão de Ação
    g.SetFont("s10 bold", "Segoe UI")
    btnOk := g.Add("Button", "Default w160 h36 x" (largura - 180) " y+20", textoBotao)
    btnOk.OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show("w" largura)
    WinWaitClose("ahk_id " g.Hwnd)
}

; ────────────────────────────────────────────────────────────────────────
; 2. CAIXA DE CONFIRMAÇÃO (Substitui MsgBox com Yes/No)
; Retorna: true (Sim/Confirmado) ou false (Não/Cancelado)
; ────────────────────────────────────────────────────────────────────────
Confirmar(titulo, cabecalho, texto, textoSim := "✅ Sim", textoNao := "❌ Não", defaultSim := true, largura := 580) {
    confirmado := false

    g := Gui("+AlwaysOnTop -MaximizeBox", titulo)
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F5F7FA"

    ; Cabeçalho
    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w" (largura - 40) " c0A1E38", "❓ " cabecalho)

    ; Divisor
    g.Add("Text", "x20 y48 w" (largura - 40) " h2 0x10")

    ; Corpo do texto
    g.SetFont("s10 norm", "Segoe UI")
    g.Add("Text", "x20 y58 w" (largura - 40) " c1A1A1A", texto)

    ; Botões
    g.SetFont("s10 bold", "Segoe UI")
    optSim := defaultSim ? "Default " : ""
    optNao := !defaultSim ? "Default " : ""

    btnSim := g.Add("Button", optSim "w140 h36 x" (largura - 310) " y+20", textoSim)
    btnNao := g.Add("Button", optNao "w140 h36 x+10", textoNao)

    btnSim.OnEvent("Click", (*) => (confirmado := true, g.Destroy()))
    btnNao.OnEvent("Click", (*) => (confirmado := false, g.Destroy()))
    g.OnEvent("Close", (*) => (confirmado := false, g.Destroy()))
    g.OnEvent("Escape", (*) => (confirmado := false, g.Destroy()))

    g.Show("w" largura)
    WinWaitClose("ahk_id " g.Hwnd)

    return confirmado
}

; ────────────────────────────────────────────────────────────────────────
; 3. CAIXA DE ENTRADA DE TEXTO (Substitui InputBox)
; Retorna objeto compatível com InputBox: Map("Result", "OK"/"Cancel", "Value", texto)
; ────────────────────────────────────────────────────────────────────────
PedirTexto(titulo, cabecalho, texto, valorPadrao := "", dica := "", largura := 580) {
    resultadoValor := ""
    resultadoStatus := "Cancel"

    g := Gui("+AlwaysOnTop -MaximizeBox", titulo)
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F5F7FA"

    ; Cabeçalho
    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w" (largura - 40) " c0A1E38", "✏️ " cabecalho)

    ; Divisor
    g.Add("Text", "x20 y48 w" (largura - 40) " h2 0x10")

    ; Instrução / Descrição
    g.SetFont("s10 norm", "Segoe UI")
    g.Add("Text", "x20 y58 w" (largura - 40) " c1A1A1A", texto)

    ; Campo de texto
    g.SetFont("s11", "Segoe UI")
    edInput := g.Add("Edit", "x20 y+12 w" (largura - 40) " h30", valorPadrao)

    if (dica != "") {
        g.SetFont("s9 italic", "Segoe UI")
        g.Add("Text", "x20 y+6 w" (largura - 40) " c666666", dica)
    }

    ; Botões
    g.SetFont("s10 bold", "Segoe UI")
    btnOk := g.Add("Button", "Default w140 h36 x" (largura - 310) " y+16", "Confirmar ➔")
    btnCancelar := g.Add("Button", "w140 h36 x+10", "Cancelar")

    btnOk.OnEvent("Click", (*) => (
        resultadoValor := edInput.Text,
        resultadoStatus := "OK",
        g.Destroy()
    ))
    btnCancelar.OnEvent("Click", (*) => (resultadoStatus := "Cancel", g.Destroy()))
    g.OnEvent("Close", (*) => (resultadoStatus := "Cancel", g.Destroy()))
    g.OnEvent("Escape", (*) => (resultadoStatus := "Cancel", g.Destroy()))

    g.Show("w" largura)
    WinWaitClose("ahk_id " g.Hwnd)

    return Map("Result", resultadoStatus, "Value", resultadoValor)
}

; ────────────────────────────────────────────────────────────────────────
; 4. CAIXA DE SELEÇÃO DE OPÇÕES (Rádio ou Checkbox)
; Retorna: Array com os itens selecionados (vazio se cancelado)
; ────────────────────────────────────────────────────────────────────────
SelecionarOpcoesUI(titulo, cabecalho, opcoes, unica := false, largura := 580) {
    resultado := []
    confirmado := false
    ctrls := []

    g := Gui("+AlwaysOnTop -MaximizeBox", titulo)
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F5F7FA"

    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w" (largura - 40) " c0A1E38", "📋 " cabecalho)

    g.Add("Text", "x20 y48 w" (largura - 40) " h2 0x10")

    g.SetFont("s10 norm", "Segoe UI")
    for i, opt in opcoes {
        posY := (i = 1) ? "y+12" : "y+6"
        if unica
            ctrl := g.Add("Radio", "x25 " posY " w" (largura - 50) (i = 1 ? " Checked" : ""), opt)
        else
            ctrl := g.Add("CheckBox", "x25 " posY " w" (largura - 50), opt)
        ctrls.Push(ctrl)
    }

    g.SetFont("s10 bold", "Segoe UI")
    
    if (!unica && opcoes.Length > 1) {
        btnTodos := g.Add("Button", "w130 h36 x20 y+20", "Marcar Todos")
        btnTodos.OnEvent("Click", (*) => MarcarTodosControles(ctrls, true))
        btnOk := g.Add("Button", "Default w140 h36 x" (largura - 310) " yp", "Confirmar ➔")
        btnCancelar := g.Add("Button", "w140 h36 x+10", "Cancelar")
    } else {
        btnOk := g.Add("Button", "Default w140 h36 x" (largura - 310) " y+20", "Confirmar ➔")
        btnCancelar := g.Add("Button", "w140 h36 x+10", "Cancelar")
    }

    ConfirmarClick(*) {
        marcados := []
        for i, ctrl in ctrls {
            if ctrl.Value
                marcados.Push(opcoes[i])
        }
        resultado := marcados
        confirmado := true
        g.Destroy()
    }

    btnOk.OnEvent("Click", ConfirmarClick)
    btnCancelar.OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show("w" largura)
    WinWaitClose("ahk_id " g.Hwnd)

    return confirmado ? resultado : []
}

MarcarTodosControles(ctrls, estado) {
    for ctrl in ctrls
        ctrl.Value := estado
}

; ────────────────────────────────────────────────────────────────────────
; 5. SELEÇÃO DE CURSOS EM LOTE (ListView Estilizado)
; Retorna: Array de objetos com os cursos selecionados
; ────────────────────────────────────────────────────────────────────────
SelecionarCursosUI(tituloJanela, cabecalho, cursos, largura := 680) {
    resultado := []
    confirmado := false

    g := Gui("+AlwaysOnTop -MaximizeBox", tituloJanela)
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F5F7FA"

    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w" (largura - 40) " c0A1E38", "🎓 " cabecalho)

    g.Add("Text", "x20 y48 w" (largura - 40) " h2 0x10")

    g.SetFont("s9", "Segoe UI")
    g.Add("Text", "x20 y58 w" (largura - 40) " c555555", "Marque as caixas dos cursos que deseja incluir nesta rodada de geração:")

    lv := g.Add("ListView", "x20 y+10 w" (largura - 40) " r16 Checked Grid", ["Curso de Pós-Graduação", "Unidade / Código", "Turma"])
    for c in cursos {
        txtUnidade := c.Has("unidadeExib") ? c["unidadeExib"] : c["unidade"]
        lv.Add(, c["posgrad"], txtUnidade, c["turma"])
    }
    lv.ModifyCol(1, 380)
    lv.ModifyCol(2, 170)
    lv.ModifyCol(3, 70)

    g.SetFont("s10 bold", "Segoe UI")
    btnMarcar := g.Add("Button", "w130 h36 x20 y+16", "Marcar Todos")
    btnDesmarcar := g.Add("Button", "w130 h36 x+10", "Desmarcar")

    btnOk := g.Add("Button", "Default w140 h36 x" (largura - 310) " yp", "Confirmar ➔")
    btnCancelar := g.Add("Button", "w140 h36 x+10", "Cancelar")

    MarcarTodosLV(*) {
        Loop lv.GetCount()
            lv.Modify(A_Index, "Check")
    }
    DesmarcarTodosLV(*) {
        Loop lv.GetCount()
            lv.Modify(A_Index, "-Check")
    }
    btnMarcar.OnEvent("Click", MarcarTodosLV)
    btnDesmarcar.OnEvent("Click", DesmarcarTodosLV)

    ConfirmarClick(*) {
        marcados := []
        linha := 0
        while (linha := lv.GetNext(linha, "C"))
            marcados.Push(cursos[linha])
        resultado := marcados
        confirmado := true
        g.Destroy()
    }

    btnOk.OnEvent("Click", ConfirmarClick)
    btnCancelar.OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show("w" largura)
    WinWaitClose("ahk_id " g.Hwnd)

    return confirmado ? resultado : []
}

; ────────────────────────────────────────────────────────────────────────
; 6. ESCOLHER COORDENADOR (Quando há mais de um nome na planilha)
; ────────────────────────────────────────────────────────────────────────
EscolherCoordenadorUI(lista, infoContexto, largura := 580) {
    if (lista.Length = 0)
        return ""
    if (lista.Length = 1)
        return lista[1]

    resultado := ""
    g := Gui("+AlwaysOnTop -MaximizeBox", "Selecionar Coordenador")
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F5F7FA"

    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w" (largura - 40) " c0A1E38", "✍️ Selecionar Coordenador")

    g.Add("Text", "x20 y48 w" (largura - 40) " h2 0x10")

    g.SetFont("s10 norm", "Segoe UI")
    g.Add("Text", "x20 y58 w" (largura - 40) " c1A1A1A", infoContexto "`n`nEsta linha possui mais de um coordenador. Selecione quem deve assinar o certificado:")

    lb := g.Add("ListBox", "x20 y+12 w" (largura - 40) " r" Min(lista.Length + 1, 6), lista)
    lb.Choose(1)

    g.SetFont("s10 bold", "Segoe UI")
    btnOk := g.Add("Button", "Default w140 h36 x" (largura - 310) " y+16", "Confirmar ➔")
    btnCancelar := g.Add("Button", "w140 h36 x+10", "Cancelar")

    btnOk.OnEvent("Click", (*) => (resultado := lb.Text, g.Destroy()))
    btnCancelar.OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show("w" largura)
    WinWaitClose("ahk_id " g.Hwnd)

    return resultado
}

; ────────────────────────────────────────────────────────────────────────
; 7. ESCOLHER TÍTULO (Dr./Dra.) E CARGO DO COORDENADOR
; ────────────────────────────────────────────────────────────────────────
EscolherTituloECargoUI(coordRaw, infoContexto, largura := 580) {
    prefixoResult := ""
    cargoResult   := ""
    confirmado    := false

    g := Gui("+AlwaysOnTop -MaximizeBox", "Título e Cargo do Coordenador")
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F5F7FA"

    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w" (largura - 40) " c0A1E38", "🎓 Título e Cargo do Coordenador")

    g.Add("Text", "x20 y48 w" (largura - 40) " h2 0x10")

    g.SetFont("s10 norm", "Segoe UI")
    g.Add("Text", "x20 y58 w" (largura - 40) " c1A1A1A", infoContexto)
    
    g.SetFont("s10 bold", "Segoe UI")
    g.Add("Text", "x20 y+10 w" (largura - 40) " c0A1E38", "Coordenador identificado: " coordRaw)

    g.SetFont("s10 bold", "Segoe UI")
    g.Add("Text", "x20 y+14 w200 c444444", "Título:")
    g.SetFont("s10 norm", "Segoe UI")
    rbDr  := g.Add("Radio", "x25 y+6 w150 Checked", "Dr.")
    rbDra := g.Add("Radio", "x+10 yp w150", "Dra.")

    g.SetFont("s10 bold", "Segoe UI")
    g.Add("Text", "x20 y+14 w200 c444444", "Cargo:")
    g.SetFont("s10 norm", "Segoe UI")
    rbCoord  := g.Add("Radio", "x25 y+6 w220 Checked", "Coordenador do Curso")
    rbCoorda := g.Add("Radio", "x+10 yp w220", "Coordenadora do Curso")

    rbDr.OnEvent("Click", (*) => (rbCoord.Value := 1, rbCoorda.Value := 0))
    rbDra.OnEvent("Click", (*) => (rbCoorda.Value := 1, rbCoord.Value := 0))

    g.SetFont("s10 bold", "Segoe UI")
    btnOk := g.Add("Button", "Default w140 h36 x" (largura - 310) " y+20", "Confirmar ➔")
    btnCancelar := g.Add("Button", "w140 h36 x+10", "Cancelar")

    btnOk.OnEvent("Click", (*) => (
        prefixoResult := rbDra.Value ? "Dra. " : "Dr. ",
        cargoResult   := rbCoorda.Value ? "Coordenadora do Curso" : "Coordenador do Curso",
        confirmado    := true,
        g.Destroy()
    ))
    btnCancelar.OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show("w" largura)
    WinWaitClose("ahk_id " g.Hwnd)

    if !confirmado
        return ""
    return Map("prefixo", prefixoResult, "cargo", cargoResult)
}

; ────────────────────────────────────────────────────────────────────────
; 8. TELA DE PROBLEMAS / CORREÇÕES (Com rolagem estilizada)
; ────────────────────────────────────────────────────────────────────────
TelaProblemasUI(resumo, detalhe, largura := 680) {
    resultado := "cancelar"

    g := Gui("+AlwaysOnTop -MaximizeBox", "Passo 9 — Validação de Dados")
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F5F7FA"

    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w" (largura - 40) " cA61B29", "⚠️ Atenção: Registros com Inconsistências")

    g.Add("Text", "x20 y48 w" (largura - 40) " h2 0x10")

    g.SetFont("s10 norm", "Segoe UI")
    g.Add("Text", "x20 y58 w" (largura - 40) " c333333", resumo)

    g.SetFont("s9", "Consolas")
    g.Add("Edit", "x20 y+10 w" (largura - 40) " r14 ReadOnly -Wrap VScroll", detalhe)

    g.SetFont("s10 bold", "Segoe UI")
    g.Add("Text", "x20 y+12 w" (largura - 40) " c0A1E38", "Como deseja proceder?")

    btnCorrigir := g.Add("Button", "Default w200 h36 x20 y+8", "🔧 Corrigir um a um")
    btnPular    := g.Add("Button", "w190 h36 x+10", "⏭️ Pular esses itens")
    btnCancelar := g.Add("Button", "w190 h36 x+10", "✖ Cancelar tudo")

    btnCorrigir.OnEvent("Click", (*) => (resultado := "corrigir", g.Destroy()))
    btnPular.OnEvent("Click", (*) => (resultado := "pular", g.Destroy()))
    btnCancelar.OnEvent("Click", (*) => (resultado := "cancelar", g.Destroy()))
    g.OnEvent("Close", (*) => (resultado := "cancelar", g.Destroy()))
    g.OnEvent("Escape", (*) => (resultado := "cancelar", g.Destroy()))

    g.Show("w" largura)
    WinWaitClose("ahk_id " g.Hwnd)

    return resultado
}

; ────────────────────────────────────────────────────────────────────────
; 9. RELATÓRIO FINAL DO LOTE (Com botão para abrir pasta de certificados)
; ────────────────────────────────────────────────────────────────────────
ExibirRelatorioFinalUI(titulo, cabecalho, textoResumo, detalheCursos, pastaSaida := "C:\Certificados\", largura := 700) {
    g := Gui("+AlwaysOnTop -MaximizeBox", titulo)
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F5F7FA"

    g.SetFont("s14 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w" (largura - 40) " c15633B", "🎉 " cabecalho)

    g.Add("Text", "x20 y48 w" (largura - 40) " h2 0x10")

    g.SetFont("s10 norm", "Segoe UI")
    g.Add("Text", "x20 y58 w" (largura - 40) " c1A1A1A", textoResumo)

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x20 y+10 w" (largura - 40) " c0A1E38", "📊 Detalhamento por Curso:")

    g.SetFont("s9", "Consolas")
    g.Add("Edit", "x20 y+6 w" (largura - 40) " r10 ReadOnly -Wrap VScroll", detalheCursos)

    g.SetFont("s9 italic", "Segoe UI")
    g.Add("Text", "x20 y+8 w" (largura - 40) " c555555", "Local de gravação: " pastaSaida)

    g.SetFont("s10 bold", "Segoe UI")
    btnAbrirPasta := g.Add("Button", "w220 h38 x20 y+12", "📂 Abrir Pasta dos Certificados")
    btnConcluir   := g.Add("Button", "Default w160 h38 x" (largura - 180) " yp", "Concluir ➔")

    btnAbrirPasta.OnEvent("Click", (*) => Run('explorer.exe "' pastaSaida '"'))
    btnConcluir.OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show("w" largura)
    WinWaitClose("ahk_id " g.Hwnd)
}

; ────────────────────────────────────────────────────────────────────────
; UTILITÁRIOS: CONVERSÃO DE COLUNAS E NORMALIZAÇÃO DE TEXTO
; ────────────────────────────────────────────────────────────────────────

; Converte índice numérico (1, 2, ..., 26, 27) para letra da coluna do Excel (A, B, ..., Z, AA)
IndiceParaColuna(colIndex) {
    coluna := ""
    while (colIndex > 0) {
        modVal := Mod(colIndex - 1, 26)
        coluna := Chr(65 + modVal) . coluna
        colIndex := Floor((colIndex - 1) / 26)
    }
    return coluna
}

; Converte letra da coluna do Excel (A, B, ..., Z, AA) para índice numérico (1, 2, ..., 26, 27)
ColunaParaIndice(colLetra) {
    colLetra := StrUpper(Trim(colLetra))
    total := 0
    Loop StrLen(colLetra) {
        c := SubStr(colLetra, A_Index, 1)
        cod := Ord(c) - 64
        if (cod < 1 || cod > 26)
            return 0
        total := (total * 26) + cod
    }
    return total
}

; Normaliza texto removendo acentuações e espaços extras para comparações robustas
NormalizarTexto(txt) {
    t := StrLower(Trim(String(txt)))
    t := RegExReplace(t, "[\x{00A0}\t]", " ")
    t := RegExReplace(t, "\s+", " ")
    
    t := StrReplace(t, "á", "a")
    t := StrReplace(t, "à", "a")
    t := StrReplace(t, "ã", "a")
    t := StrReplace(t, "â", "a")
    t := StrReplace(t, "é", "e")
    t := StrReplace(t, "ê", "e")
    t := StrReplace(t, "í", "i")
    t := StrReplace(t, "ó", "o")
    t := StrReplace(t, "ô", "o")
    t := StrReplace(t, "õ", "o")
    t := StrReplace(t, "ú", "u")
    t := StrReplace(t, "ü", "u")
    t := StrReplace(t, "ç", "c")
    return t
}

; ────────────────────────────────────────────────────────────────────────
; INTEGRAÇÃO DINÂMICA ONEDRIVE / SHAREPOINT — HOSPITAL ALBERT EINSTEIN
; ────────────────────────────────────────────────────────────────────────

; Localiza a raiz do OneDrive Corporativo do Einstein na máquina
ObterOneDriveEinstein() {
    candidatos := [
        "C:\Users\" A_UserName "\OneDrive - Hospital Albert Einstein",
        EnvGet("OneDriveCommercial"),
        EnvGet("OneDrive"),
        "C:\Users\" A_UserName "\OneDrive - Sociedade Beneficente Israelita Brasileira Hospital Albert Einstein"
    ]
    
    for c in candidatos {
        if (c != "" && DirExist(c))
            return RTrim(c, "\") "\"
    }
    
    ; Busca varrendo a pasta do usuário por qualquer pasta OneDrive com o nome Einstein
    Loop Files, "C:\Users\" A_UserName "\OneDrive*", "D" {
        if InStr(A_LoopFileName, "Einstein") || InStr(A_LoopFileName, "Hospital")
            return A_LoopFilePath "\"
    }
    
    Loop Files, "C:\Users\" A_UserName "\OneDrive*", "D" {
        return A_LoopFilePath "\"
    }
    
    return ""
}

; Retorna um Map com os principais caminhos da estrutura compartilhada
ObterCaminhosCompartilhados() {
    raizOneDrive := ObterOneDriveEinstein()
    
    pastaDeclaracoes := raizOneDrive != "" ? raizOneDrive "Planilha compartilhada - Declarações\" : ""
    planilhaMestre   := pastaDeclaracoes != "" ? pastaDeclaracoes "Controle de Declaração de Aula.xlsx" : ""
    modeloPptx       := pastaDeclaracoes != "" ? pastaDeclaracoes "Certificado_Modelo_AHK.pptx" : ""
    pastaAssinaturas := pastaDeclaracoes != "" ? pastaDeclaracoes "Coordenadores e suas assinaturas\" : ""
    
    return Map(
        "raizOneDrive",     raizOneDrive,
        "pastaDeclaracoes", pastaDeclaracoes,
        "planilhaMestre",   planilhaMestre,
        "modeloPptx",       modeloPptx,
        "pastaAssinaturas", pastaAssinaturas,
        "pastaLocalCerts",  "C:\Certificados\",
        "pastaLocalAssin",  "C:\CAssinaturas\"
    )
}

; ────────────────────────────────────────────────────────────────────────
; RECONHECIMENTO DE UNIDADE E EQUIVALÊNCIAS DE SEMESTRE (2026-5 = 2026-1S)
; ────────────────────────────────────────────────────────────────────────

ExtrairDadosCursoDeCaminho(caminhoPasta) {
    caminhoLimpo := RTrim(caminhoPasta, "\")
    SplitPath(caminhoLimpo, &nomePasta)
    
    ; Extrai unidade se estiver entre parênteses (ex: "Psiquiatria Adulto (PTA)")
    unidade := ""
    if RegExMatch(nomePasta, "i)\((PTA|CEP|RJ|GOI|VM|MOR|IBI|ALPH|BH|DF|SJC|[A-Z]{2,5})\)", &m)
        unidade := m[1]
    else if RegExMatch(caminhoLimpo, "i)[\\/](PTA|CEP|RJ|GOI|VM|MOR|IBI|ALPH|BH|DF|SJC)[\\/]", &m)
        unidade := m[1]
        
    ; Extrai nome limpo do curso
    nomeCurso := RegExReplace(nomePasta, "\s*\([^)]*\)", "")
    nomeCurso := Trim(nomeCurso)
    
    ; Extrai semestre do caminho (ex: \2026-1S\ ou \2026-5\ ou \26-5\)
    semestre := ""
    if RegExMatch(caminhoLimpo, "i)[\\/](\d{2,4}[-.]?[1256](?:S)?)[\\/]", &m)
        semestre := m[1]
        
    return Map("nomeCurso", nomeCurso, "unidade", unidade, "semestre", semestre, "nomePasta", nomePasta)
}

ExtrairSiglaUnidade(codigoOuUnidade) {
    txt := Trim(codigoOuUnidade)
    if (txt = "")
        return ""
    
    ; Se contém padrão como "2026-A-PTA-1S", extrai a sigla da unidade "PTA"
    if RegExMatch(txt, "i)-(PTA|CEP|RJ|GOI|VM|MOR|IBI|ALPH|BH|DF|SJC|[A-Z]{2,5})(?:-|$)", &m)
        return StrUpper(m[1])
    
    ; Se já for sigla curta direta (ex: "PTA", "CEP")
    if (StrLen(txt) <= 5 && RegExMatch(txt, "^[A-Za-z]+$"))
        return StrUpper(txt)
    
    return txt
}

ObterVariacoesUnidade(unidadeRaw) {
    sigla := ExtrairSiglaUnidade(unidadeRaw)
    variacoes := []
    if (sigla = "")
        return variacoes
    
    siglaUpper := StrUpper(sigla)
    siglaNorm  := NormalizarTexto(sigla)
    if (siglaNorm != "")
        variacoes.Push(siglaNorm)
    
    mapaUnidades := Map(
        "PTA",  ["paulista", "pta", "bela vista"],
        "CEP",  ["cerqueira cesar", "cep", "paulista"],
        "RJ",   ["rio de janeiro", "rj", "rio"],
        "GOI",  ["goiania", "goi", "goias"],
        "VM",   ["vila mariana", "vm"],
        "MOR",  ["morumbi", "mor"],
        "IBI",  ["ibirapuera", "ibi"],
        "ALPH", ["alphaville", "alph", "alpha"],
        "BH",   ["belo horizonte", "bh"],
        "DF",   ["brasilia", "df"],
        "SJC",  ["sao jose dos campos", "sjc", "sao jose"]
    )
    
    if mapaUnidades.Has(siglaUpper) {
        for v in mapaUnidades[siglaUpper] {
            vNorm := NormalizarTexto(v)
            if (vNorm = "")
                continue
            jaTem := false
            for item in variacoes {
                if (item = vNorm) {
                    jaTem := true
                    break
                }
            }
            if !jaTem
                variacoes.Push(vNorm)
        }
    }
    
    return variacoes
}

ObterEquivalenciasSemestre(semestreRaw) {
    semNorm := StrLower(Trim(semestreRaw))
    semNorm := StrReplace(semNorm, " ", "")
    semNorm := StrReplace(semNorm, "_", "-")
    
    if (semNorm = "")
        return []
    
    variacoes := [semNorm]
    
    ; Detecta ano e sufixo de semestre (1S/5 ou 2S/6)
    if RegExMatch(semNorm, "(20\d{2}|\d{2})[-.]?([1256])(?:s)?", &m) {
        anoStr := m[1]
        semNum := m[2]
        
        ano4 := (StrLen(anoStr) = 2) ? ("20" anoStr) : anoStr
        ano2 := SubStr(ano4, 3, 2)
        
        ehPrimeiroSem := (semNum = "1" || semNum = "5")
        
        if ehPrimeiroSem {
            variacoes.Push(ano4 "-1s", ano4 "-1", ano4 "-5", ano2 "-1s", ano2 "-1", ano2 "-5", ano2 "-5s", ano4 ".1", ano2 ".1", "1s" ano4, "1s" ano2)
        } else {
            variacoes.Push(ano4 "-2s", ano4 "-2", ano4 "-6", ano2 "-2s", ano2 "-2", ano2 "-6", ano2 "-6s", ano4 ".2", ano2 ".2", "2s" ano4, "2s" ano2)
        }
    }
    
    return variacoes
}

; Localiza dinamicamente a planilha do curso (Planejamento Mensal)
; Busca em todas as pastas de planejamento presentes no OneDrive com inteligência de Unidade e Semestre
LocalizarPlanilhaCursoDinamica(nomeCurso, unidade := "", semestre := "") {
    raizOneDrive := ObterOneDriveEinstein()
    if (raizOneDrive = "")
        return ""
    
    nomeCursoNorm := NormalizarTexto(nomeCurso)
    variacoesUnidade := ObterVariacoesUnidade(unidade)
    variacoesSemestre := ObterEquivalenciasSemestre(semestre)
    
    ; Extrai palavras-chave do curso (ex: psiquiatria, adulto)
    palavras := StrSplit(nomeCursoNorm, " ")
    palavrasCurso := []
    for w in palavras {
        w := Trim(w)
        if (StrLen(w) >= 3 && w != "para" && w != "com" && w != "pos" && w != "graduacao" && w != "especializacao" && w != "em" && w != "de" && w != "da" && w != "do" && w != "dos" && w != "das")
            palavrasCurso.Push(w)
    }
    
    pastasPlanejamento := []
    
    ; Pastas do OneDrive (prioriza pastas do semestre atual e equivalentes)
    for semVar in variacoesSemestre {
        if (semVar = "")
            continue
        pSem := raizOneDrive "Planilha compartilhada - Pós-graduação " semVar "\"
        if DirExist(pSem)
            pastasPlanejamento.Push(pSem)
    }
    
    Loop Files, raizOneDrive "Planilha compartilhada - Pós-graduação*", "D" {
        caminhoDir := A_LoopFilePath "\"
        jaTem := false
        for p in pastasPlanejamento {
            if (p = caminhoDir) {
                jaTem := true
                break
            }
        }
        if !jaTem
            pastasPlanejamento.Push(caminhoDir)
    }
    
    pastaDecl := raizOneDrive "Planilha compartilhada - Declarações\"
    if DirExist(pastaDecl)
        pastasPlanejamento.Push(pastaDecl)
    
    ; Coleta todos os arquivos .xlsx das pastas
    todosArquivos := []
    for pastaBase in pastasPlanejamento {
        Loop Files, pastaBase "*.xlsx", "R" {
            todosArquivos.Push(A_LoopFilePath)
        }
    }
    
    ; ─── TIER 1 (Perfeito): TODAS as palavras do curso + UNIDADE + SEMESTRE ───
    if (palavrasCurso.Length > 0 && variacoesUnidade.Length > 0 && variacoesSemestre.Length > 0) {
        for arq in todosArquivos {
            nomeArqNorm := NormalizarTexto(arq)
            
            temCurso := true
            for kw in palavrasCurso {
                if (kw = "" || !InStr(nomeArqNorm, kw)) {
                    temCurso := false
                    break
                }
            }
            if !temCurso
                continue
                
            temUnidade := false
            for u in variacoesUnidade {
                if (u != "" && InStr(nomeArqNorm, u)) {
                    temUnidade := true
                    break
                }
            }
            if !temUnidade
                continue
                
            temSem := false
            for s in variacoesSemestre {
                if (s != "" && InStr(nomeArqNorm, s)) {
                    temSem := true
                    break
                }
            }
            if temSem
                return arq
        }
    }
    
    ; ─── TIER 2: TODAS as palavras do curso + UNIDADE ───
    if (palavrasCurso.Length > 0 && variacoesUnidade.Length > 0) {
        for arq in todosArquivos {
            nomeArqNorm := NormalizarTexto(arq)
            
            temCurso := true
            for kw in palavrasCurso {
                if (kw = "" || !InStr(nomeArqNorm, kw)) {
                    temCurso := false
                    break
                }
            }
            if !temCurso
                continue
                
            temUnidade := false
            for u in variacoesUnidade {
                if (u != "" && InStr(nomeArqNorm, u)) {
                    temUnidade := true
                    break
                }
            }
            if temUnidade
                return arq
        }
    }
    
    ; ─── TIER 3: TODAS as palavras do curso + SEMESTRE ───
    if (palavrasCurso.Length > 0 && variacoesSemestre.Length > 0) {
        for arq in todosArquivos {
            nomeArqNorm := NormalizarTexto(arq)
            
            temCurso := true
            for kw in palavrasCurso {
                if (kw = "" || !InStr(nomeArqNorm, kw)) {
                    temCurso := false
                    break
                }
            }
            if !temCurso
                continue
                
            temSem := false
            for s in variacoesSemestre {
                if (s != "" && InStr(nomeArqNorm, s)) {
                    temSem := true
                    break
                }
            }
            if temSem
                return arq
        }
    }
    
    ; ─── TIER 4: TODAS as palavras do curso ───
    if (palavrasCurso.Length > 0) {
        for arq in todosArquivos {
            nomeArqNorm := NormalizarTexto(arq)
            
            temCurso := true
            for kw in palavrasCurso {
                if (kw = "" || !InStr(nomeArqNorm, kw)) {
                    temCurso := false
                    break
                }
            }
            if temCurso
                return arq
        }
    }
    
    return ""
}

; Garante a criação de toda a estrutura de pastas na Nuvem e no Backup Local
GarantirPastasSaida(nomeCurso, mesFormatado, semestre := "2026-5") {
    info := ObterCaminhosCompartilhados()
    
    ; 1. Caminho na Nuvem (SharePoint / OneDrive)
    caminhoNuvem := ""
    if (info["pastaDeclaracoes"] != "") {
        caminhoNuvem := info["pastaDeclaracoes"] semestre "\" nomeCurso "\" mesFormatado "\"
        try DirCreate(caminhoNuvem)
    }
    
    ; 2. Caminho Local (Backup em C:\Certificados\)
    caminhoLocal := info["pastaLocalCerts"] nomeCurso "\" mesFormatado "\"
    try DirCreate(caminhoLocal)
    
    ; 3. Caminho Local por Docente
    caminhoLocalDocentes := info["pastaLocalCerts"] nomeCurso "\Docentes\"
    try DirCreate(caminhoLocalDocentes)
    
    return Map(
        "nuvem", caminhoNuvem,
        "local", caminhoLocal,
        "localDocentes", caminhoLocalDocentes
    )
}

; Garante a pasta "Para Enviar\<Curso>\<Mês>\" usada pelo manifesto do Power Automate.
; Profundidade FIXA em 2 níveis (Curso → Mês), sem nível de semestre — o flow depende
; dessa estrutura fixa para não precisar de recursão genérica. Não mudar essa profundidade.
GarantirPastaParaEnviar(nomeCurso, mesFormatado) {
    info := ObterCaminhosCompartilhados()
    if (info["pastaDeclaracoes"] = "")
        return ""

    caminho := info["pastaDeclaracoes"] "Para Enviar\" nomeCurso "\" mesFormatado "\"
    try DirCreate(caminho)
    return caminho
}

; =========================================================
; FUNÇÃO: Formatar horas de texto livre para leitura humana
; Aceita: "10 horas", "08h - 18h", "2:30", "4h30", "30min"
; =========================================================

FormatarHorasSeguro(txtOriginal) {
    txt := Trim(txtOriginal)
    if (txt = "")
        return "Horario invalido"

    txt := RegExReplace(txt, "[\x{00A0}\x{2007}\x{202F}\t]", " ")
    txt := RegExReplace(txt, "[\x{2010}-\x{2015}\x{2212}]", "-")
    txt := StrReplace(txt, "às", "-")
    txt := StrReplace(txt, "Às", "-")
    txt := RegExReplace(txt, "\s+", " ")
    txt := Trim(txt)

    ; ── Carga horária direta: "4h", "4 h", "4 horas", "4h30", "4:30" ──
    if RegExMatch(txt, "i)^(\d+)\s*[h:]\s*(\d+)\s*(min)?\s*(horas?|h)?\s*$", &m) {
        h := Integer(m[1])
        min := Integer(m[2])
        if (h > 0 && min > 0)
            return h " horas e " min " minutos"
        if (h > 0)
            return (h = 1) ? "1 hora" : h " horas"
        if (min > 0)
            return min " minutos"
    }
    if RegExMatch(txt, "i)^(\d+)\s*(horas?|h)\s*$", &m) {
        h := Integer(m[1])
        return (h = 1) ? "1 hora" : h " horas"
    }
    if RegExMatch(txt, "i)^(\d+)\s*min(utos?)?\s*$", &m) {
        return Integer(m[1]) " minutos"
    }

    ; ── Intervalo HH:MM - HH:MM (formato original) ───────────────────
    if RegExMatch(txt, "i)(\d{1,2})\s*[:h]?\s*(\d{2})?\s*h?\s*[-–—−]\s*(\d{1,2})\s*[:h]?\s*(\d{2})?\s*h?", &m) {
        horaIni := m[1]
        minIni  := (m[2] != "") ? m[2] : "00"
        horaFim := m[3]
        minFim  := (m[4] != "") ? m[4] : "00"
        return CalcularDuracao(horaIni ":" minIni, horaFim ":" minFim)
    }

    return "Horario invalido (formato não reconhecido: " txtOriginal ")"
}

; =========================================================
; FUNÇÃO: Calcular duração entre HH:MM e HH:MM
; =========================================================

CalcularDuracao(inicio, fim) {
    try {
        inicioPartes := StrSplit(inicio, ":")
        fimPartes := StrSplit(fim, ":")
        if (inicioPartes.Length < 2 || fimPartes.Length < 2)
            return "Horario invalido"
        inicioMin := (Integer(inicioPartes[1]) * 60) + Integer(inicioPartes[2])
        fimMin := (Integer(fimPartes[1]) * 60) + Integer(fimPartes[2])
        duracao := fimMin - inicioMin
        if (duracao <= 0)
            return "Horario invalido (fim <= inicio)"
        horas := Floor(duracao / 60)
        minutos := Mod(duracao, 60)
        if (horas > 0 && minutos > 0) {
            horaLabel := (horas = 1) ? "1 hora" : horas " horas"
            return horaLabel " e " minutos " minutos"
        }
        if (horas > 0)
            return (horas = 1) ? "1 hora" : horas " horas"
        return minutos " minutos"
    } catch as err {
        return "Horario invalido (" err.Message ")"
    }
}

; =========================================================
; FUNÇÃO: Interpretar duração digitada manualmente
; Aceita: "4" → 4 horas | "2:30" → 2h e 30min | "2,30" → 2h e 30min
; =========================================================

FormatarDuracaoDigitada(entrada) {
    txt := Trim(entrada)
    if (txt = "")
        return ""

    txt := StrReplace(txt, ",", ":")

    if RegExMatch(txt, "^(\d+):(\d+)$", &m) {
        h := Integer(m[1])
        min := Integer(m[2])
        if (h > 0 && min > 0)
            return h " horas e " min " minutos"
        if (h > 0)
            return (h = 1) ? "1 hora" : h " horas"
        if (min > 0)
            return min " minutos"
    }

    if RegExMatch(txt, "^\d+$") {
        h := Integer(txt)
        if (h > 0)
            return (h = 1) ? "1 hora" : h " horas"
    }

    return ""
}

; =========================================================
; FUNÇÃO: Formatar horas a partir de valor decimal Excel
; =========================================================

FormatarHorasDeValorExcel(val) {
    if (val = "" || !IsNumber(val))
        return "Horario invalido"
    totalMin := Round(val * 24 * 60)
    horas := Floor(totalMin / 60)
    minutos := Mod(totalMin, 60)
    if (minutos > 0)
        return horas "h" Format("{:02}", minutos)
    else
        return horas " horas"
}
