#Requires AutoHotkey v2.0
#SingleInstance Force
#Include UI_Moderna.ahk
#Include AHK\UIA-v2-main\UIA-v2-main\Lib\UIA.ahk
#Include AHK\UIA-v2-main\UIA-v2-main\Lib\UIA_Browser.ahk

; ════════════════════════════════════════════════════════════════════════
;   ENVIAR E-MAILS — PÓS-GRADUAÇÃO (v11 — AUTOMATIZADO)
;   Versão com Anexação Automática no Outlook Web via UI Automation
; ════════════════════════════════════════════════════════════════════════
;
;   DESTAQUES DESTA VERSÃO (v11):
;   - ⚡ Anexação 100% Automática de Arquivos no Outlook Web:
;     O script conecta ao navegador (Edge/Chrome) via UI Automation, aguarda
;     o carregamento da página da web, aciona o anexo e insere os PDFs do docente,
;     eliminando completamente a necessidade de arrastar arquivos com o mouse.
;   - 🚀 Dois Modos de Envio de Alta Eficiência:
;     1. Semi-Automático (Supervisionado): Anexa automaticamente, você confere
;        na tela e pressiona ENTER para enviar e avançar para o próximo.
;     2. 100% Automático (Turbo): Anexa, aguarda upload e clica em Enviar sozinho.
;   - 📂 Suporte a Múltiplos Certificados:
;     Se um docente tiver aulas em múltiplos meses da rodada, todos os PDFs
;     são localizados e anexados juntos em uma única operação.
;   - 🧹 Área de Trabalho Limpa:
;     Não abre mais janelas acumuladas do Windows Explorer.
;   - 📊 Leitura Otimizada da Planilha e Detecção de Colunas na Linha 3.
;   - 💾 Controle de Progresso Persistente (_EnvioControle.txt).
; ════════════════════════════════════════════════════════════════════════

; =========================================================
; CONFIGURAÇÕES GERAIS E INTEGRAÇÃO NUVEM
; =========================================================

global INFO_NUVEM        := ObterCaminhosCompartilhados()
global PASTA_PADRAO_CERTS := (INFO_NUVEM["pastaDeclaracoes"] != "" && DirExist(INFO_NUVEM["pastaDeclaracoes"])) ? INFO_NUVEM["pastaDeclaracoes"] : "C:\Certificados\"
global CONTROLE_ARQUIVO   := "_EnvioControle.txt"

; =========================================================
; FUNÇÕES AUXILIARES — SISTEMA DE ARQUIVOS E VALIDAÇÃO
; =========================================================

ListarSubpastas(pasta) {
    lista := []
    if !DirExist(pasta)
        return lista
    Loop Files, pasta "*", "D" {
        if (A_LoopFileName = "_EnvioTemp" || A_LoopFileName = "Docentes")
            continue
        lista.Push(A_LoopFileName)
    }
    return lista
}

EmailEhValido(email) {
    email := Trim(email)
    if (email = "")
        return false
    return RegExMatch(email, "^[^@\s]+@[^@\s]+\.[^@\s]+$") > 0
}

ExtrairEmailValido(bruto) {
    bruto := Trim(bruto)
    if RegExMatch(bruto, "[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", &m)
        return m[0]
    return bruto
}

SanitizarNomeDocente(nomeBruto) {
    nome := Trim(String(nomeBruto))
    nome := RegExReplace(nome, "[\x{00A0}\t]", " ")
    nome := RegExReplace(nome, "\s*\b[HMhm]\b\s*$")
    nome := Trim(nome)
    nome := RegExReplace(nome, "\s*\([^)]*\)", "")
    nome := RegExReplace(nome, "i)^(dra|dr|profa|professora|professor|prof|sra|sr)\.?\s+", "")
    nome := RegExReplace(nome, '[\\/:*?"<>|]', "")
    nome := Trim(RegExReplace(nome, "\s+", " "))
    return nome
}

NormalizarNome(s) {
    s := StrLower(Trim(String(s)))
    s := RegExReplace(s, "[\x{00A0}\t]", " ")
    s := RegExReplace(s, "\s+", " ")
    return s
}

global MESES_NUMERO_NOME := Map(
    "01", "janeiro", "02", "fevereiro", "03", "marco", "04", "abril",
    "05", "maio", "06", "junho", "07", "julho", "08", "agosto",
    "09", "setembro", "10", "outubro", "11", "novembro", "12", "dezembro"
)

RemoverAcentos(s) {
    s := StrLower(String(s))
    subs := Map(
        "á", "a", "à", "a", "ã", "a", "â", "a", "ä", "a",
        "é", "e", "ê", "e", "è", "e", "ë", "e",
        "í", "i", "ì", "i", "î", "i", "ï", "i",
        "ó", "o", "ò", "o", "õ", "o", "ô", "o", "ö", "o",
        "ú", "u", "ù", "u", "û", "u", "ü", "u",
        "ç", "c"
    )
    for de, para in subs
        s := StrReplace(s, de, para)
    return s
}

NormalizarParaComparar(s) {
    s := RemoverAcentos(Trim(s))
    return RegExReplace(s, "[^a-z0-9]", "")
}

ExtrairMesAnoDaPasta(nomePasta) {
    if RegExMatch(Trim(nomePasta), "^(\d{1,2})\.(\d{2,4})$", &m) {
        numMes := Format("{:02}", Integer(m[1]))
        if MESES_NUMERO_NOME.Has(numMes)
            return Map("numMes", numMes, "nomeMes", MESES_NUMERO_NOME[numMes], "ano", m[2])
    }
    return ""
}

AbaCorrespondeAoMes(abaNome, nomePasta) {
    abaNorm := NormalizarParaComparar(abaNome)
    pastaNorm := NormalizarParaComparar(nomePasta)
    if (abaNorm = pastaNorm)
        return true

    infoMes := ExtrairMesAnoDaPasta(nomePasta)
    if (infoMes = "")
        return false

    nomeMesNorm := infoMes["nomeMes"]
    numMesNorm  := infoMes["numMes"]
    anoCurto    := (StrLen(infoMes["ano"]) = 4) ? SubStr(infoMes["ano"], 3, 2) : infoMes["ano"]
    anoLongo    := (StrLen(infoMes["ano"]) = 2) ? "20" infoMes["ano"] : infoMes["ano"]

    if InStr(abaNorm, nomeMesNorm) {
        resto := Trim(StrReplace(abaNorm, nomeMesNorm, "", , , 1))
        if (resto = "" || resto = anoCurto || resto = anoLongo)
            return true
        if !RegExMatch(resto, "^\d+$")
            return true
        return false
    }

    if RegExMatch(abaNorm, "^\d+$")
        return (Integer(abaNorm) = Integer(numMesNorm))

    return false
}

LocalizarArquivosDocente(pastaPosgrad, mesesSelecionados, nomeDocente, enviadosMap) {
    arquivos := []
    nomeAlvoNorm := NormalizarNome(nomeDocente)
    chaveDocente := StrLower(nomeDocente)

    for mes in mesesSelecionados {
        if (enviadosMap.Has(chaveDocente) && enviadosMap[chaveDocente].Has(mes))
            continue

        pastaMes := pastaPosgrad mes "\"
        if !DirExist(pastaMes)
            continue

        Loop Files, pastaMes "*.pdf" {
            nomeArquivoSemExt := RegExReplace(A_LoopFileName, "\.pdf$", "")
            nomeDocenteArquivo := RegExReplace(nomeArquivoSemExt, "\s*\([^)]*\)(\.\d+)?$", "")
            nomeDocenteArquivo := SanitizarNomeDocente(nomeDocenteArquivo)

            if (NormalizarNome(nomeDocenteArquivo) = nomeAlvoNorm || InStr(NormalizarNome(nomeDocenteArquivo), nomeAlvoNorm) || InStr(nomeAlvoNorm, NormalizarNome(nomeDocenteArquivo))) {
                arquivos.Push(A_LoopFileFullPath)
            }
        }
    }

    if (arquivos.Length = 0) {
        pastaDocenteDirect := pastaPosgrad "Docentes\" nomeDocente "\"
        if DirExist(pastaDocenteDirect) {
            Loop Files, pastaDocenteDirect "*.pdf"
                arquivos.Push(A_LoopFileFullPath)
        }
    }

    return arquivos
}

JoinArr(arr, sep) {
    resultado := ""
    for idx, item in arr {
        resultado .= (idx = 1 ? "" : sep) item
    }
    return resultado
}

; =========================================================
; DETECÇÃO AUTOMÁTICA DE COLUNAS NO EXCEL
; =========================================================

IdentificarColunaDocentes(wsAba) {
    maxCols := 50
    try {
        usedCols := wsAba.UsedRange.Columns.Count
        if (usedCols > maxCols)
            maxCols := usedCols
    }

    termos := ["docente", "professor", "professora", "nome"]
    for linhaCheck in [3, 4, 2, 1] {
        Loop maxCols {
            colIdx := A_Index
            celVal := ""
            try celVal := String(wsAba.Cells(linhaCheck, colIdx).Value)
            if (celVal = "")
                try celVal := String(wsAba.Cells(linhaCheck, colIdx).Text)

            celValNorm := NormalizarTexto(celVal)
            for t in termos {
                if InStr(celValNorm, t)
                    return IndiceParaColuna(colIdx)
            }
        }
    }
    return ""
}

IdentificarColunaEmails(wsAba) {
    maxCols := 50
    try {
        usedCols := wsAba.UsedRange.Columns.Count
        if (usedCols > maxCols)
            maxCols := usedCols
    }

    termos := ["email", "e-mail", "contato", "correio", "eletronico"]
    for linhaCheck in [3, 4, 2, 1] {
        Loop maxCols {
            colIdx := A_Index
            celVal := ""
            try celVal := String(wsAba.Cells(linhaCheck, colIdx).Value)
            if (celVal = "")
                try celVal := String(wsAba.Cells(linhaCheck, colIdx).Text)

            celValNorm := NormalizarTexto(celVal)
            for t in termos {
                if InStr(celValNorm, t)
                    return IndiceParaColuna(colIdx)
            }
        }
    }
    return ""
}

; =========================================================
; CONTROLE DE PROGRESSO (PERSISTÊNCIA _EnvioControle.txt)
; =========================================================

CarregarEnviados(pastaControle) {
    enviados := Map()
    if !FileExist(pastaControle)
        return enviados
    Loop Read, pastaControle {
        campos := StrSplit(A_LoopReadLine, "`t")
        if (campos.Length < 2)
            continue
        nomeDocente := Trim(campos[1])
        mes := Trim(campos[2])
        if (nomeDocente = "" || mes = "")
            continue
        chave := StrLower(nomeDocente)
        if !enviados.Has(chave)
            enviados[chave] := Map()
        enviados[chave][mes] := true
    }
    return enviados
}

MarcarEnviado(pastaControle, nomeDocente, mes) {
    linha := nomeDocente "`t" mes "`t" FormatTime(A_Now, "yyyy-MM-dd HH:mm") "`n"
    try FileAppend(linha, pastaControle, "UTF-8")
}

EstaTotalmenteEnviado(enviadosMap, chave, mesesSelecionados) {
    if !enviadosMap.Has(chave)
        return false
    for mes in mesesSelecionados {
        if !enviadosMap[chave].Has(mes)
            return false
    }
    return true
}

MesesPendentes(enviadosMap, chave, mesesSelecionados) {
    pendentes := []
    for mes in mesesSelecionados {
        jaEnviado := enviadosMap.Has(chave) && enviadosMap[chave].Has(mes)
        if !jaEnviado
            pendentes.Push(mes)
    }
    return pendentes
}

MarcarPendentesComoEnviados(pastaControle, nomeDocente, chave, mesesSelecionados, enviadosMap) {
    pendentes := MesesPendentes(enviadosMap, chave, mesesSelecionados)
    for mes in pendentes {
        MarcarEnviado(pastaControle, nomeDocente, mes)
        if !enviadosMap.Has(chave)
            enviadosMap[chave] := Map()
        enviadosMap[chave][mes] := true
    }
}

; =========================================================
; NAVEGAÇÃO E AUTOMAÇÃO NO OUTLOOK WEB VIA UI AUTOMATION
; =========================================================

ObterJanelaNavegadorOutlook(timeoutSegundos := 20) {
    inicio := A_TickCount
    while ((A_TickCount - inicio) < timeoutSegundos * 1000) {
        if WinExist("ahk_exe msedge.exe") {
            WinActivate("ahk_exe msedge.exe")
            return WinGetID("ahk_exe msedge.exe")
        }
        if WinExist("ahk_exe chrome.exe") {
            WinActivate("ahk_exe chrome.exe")
            return WinGetID("ahk_exe chrome.exe")
        }
        if WinExist("Novo email") {
            WinActivate("Novo email")
            return WinGetID("Novo email")
        }
        Sleep(350)
    }
    return WinGetID("A")
}

AnexarArquivosOutlookWeb(hwndNavegador, arquivosArr, maxTimeoutSegundos := 25) {
    if (arquivosArr.Length = 0)
        return true

    listaFormatada := ""
    for arq in arquivosArr {
        listaFormatada .= '"' arq '" '
    }
    listaFormatada := Trim(listaFormatada)

    WinActivate(hwndNavegador)

    ; 1. Espera ativa pelo botão "Anexar arquivo" aparecer (aguarda página carregar)
    inicio := A_TickCount
    cUIA := ""
    btnAnexar := ""

    ToolTip("⏳ Aguardando Outlook Web carregar...")

    while ((A_TickCount - inicio) < maxTimeoutSegundos * 1000) {
        try {
            WinActivate(hwndNavegador)
            if !cUIA
                cUIA := UIA_Browser("ahk_id " hwndNavegador)

            doc := cUIA.GetCurrentDocumentElement()
            if doc {
                try btnAnexar := doc.FindElement({Name: "Anexar arquivo", Type: "Button"})
                if !btnAnexar
                    try btnAnexar := doc.FindElement({Name: "Anexar", Type: "Button"})
                if !btnAnexar
                    try btnAnexar := doc.FindElement({Name: "Anexar arquivo"})
                if btnAnexar
                    break
            }
        } catch {
            cUIA := ""
        }
        Sleep(600)
    }

    ToolTip()

    if !btnAnexar
        return false

    ; 2. Clica em "Anexar arquivo"
    ToolTip("📎 Abrindo menu de anexo...")
    Sleep(350)
    btnAnexar.Click()
    Sleep(900)

    ; 3. Localiza "Navegar neste computador" via polling (doc.WaitElement não existe no contexto do UIA_Browser document)
    ToolTip("🔍 Localizando 'Navegar neste computador'...")
    itemNavegar := ""
    inicioMenu := A_TickCount
    while ((A_TickCount - inicioMenu) < 5000) {
        try {
            docAtual := cUIA.GetCurrentDocumentElement()
            try itemNavegar := docAtual.FindElement({Name: "Navegar neste computador"})
            if !itemNavegar
                try itemNavegar := docAtual.FindElement({Name: "Navegar no computador"})
            if itemNavegar
                break
        } catch {
        }
        Sleep(300)
    }

    if itemNavegar {
        ToolTip("🖱️ Clicando em 'Navegar neste computador'...")
        itemNavegar.Click()
    } else {
        ; Fallback via teclado: Home+Enter navega até o 1º item do menu ("Navegar neste computador")
        ToolTip("⌨️ Acionando via teclado (Home + Enter)...")
        Send("{Home}")
        Sleep(150)
        Send("{Enter}")
    }
    ToolTip()

    ; 4. Aguarda o diálogo nativo do Windows "Abrir"
    if WinWaitActive("Abrir ahk_class #32770", , 10) || WinWaitActive("Open ahk_class #32770", , 3) {
        dialogoHwnd := WinGetID("A")
        Sleep(300)

        ControlSetText(listaFormatada, "Edit1", dialogoHwnd)
        Sleep(400)

        ControlSend("{Enter}", "Edit1", dialogoHwnd)
        Sleep(500)

        if WinExist("ahk_id " dialogoHwnd) {
            try ControlClick("&Abrir", dialogoHwnd)
            try ControlClick("Abrir", dialogoHwnd)
            try ControlClick("Button2", dialogoHwnd)
        }

        WinWaitClose(dialogoHwnd, , 8)
        Sleep(2200)
        return true
    }

    return false
}

DispararEnvioOutlookWeb(hwndNavegador) {
    WinActivate(hwndNavegador)
    Sleep(400)

    try {
        cUIA := UIA_Browser("ahk_id " hwndNavegador)
        doc := cUIA.GetCurrentDocumentElement()
        botaoEnviar := doc.FindElement({Name: "Enviar", Type: "Button"})
        if botaoEnviar {
            botaoEnviar.Click()
            Sleep(800)
            return true
        }
    } catch {
    }

    Send("^{Enter}")
    Sleep(800)
    return true
}

; =========================================================
; COMPONENTES VISUAIS MODERNOS
; =========================================================

CorrigirAbasUI(tituloJanela, cabecalho, meses, abasDisponiveis, largura := 600) {
    resultado  := ""
    confirmado := false
    combos     := []

    g := Gui("+AlwaysOnTop -MaximizeBox", tituloJanela)
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F7F9FC"

    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w" (largura - 40) " c003366", "📑 " cabecalho)

    g.Add("Text", "x20 y48 w" (largura - 40) " h2 0x10")

    g.SetFont("s10 norm", "Segoe UI")
    g.Add("Text", "x20 y58 w" (largura - 40) " c222222", 
        "Os nomes das subpastas não bateram de forma idêntica com as abas da planilha.`n"
        "Selecione para cada mês a aba correspondente no Excel:"
    )

    posY := 105
    for m in meses {
        g.SetFont("s10 bold", "Segoe UI")
        g.Add("Text", "x25 y" posY " w140 h26 c003366", "Mês: " m)

        indicePalpite := 1
        for i, aba in abasDisponiveis {
            if AbaCorrespondeAoMes(aba, m) {
                indicePalpite := i
                break
            }
        }

        g.SetFont("s10 norm", "Segoe UI")
        cb := g.Add("DDL", "x170 y" (posY - 2) " w" (largura - 195) " Choose" indicePalpite, abasDisponiveis)
        combos.Push(cb)
        posY += 38
    }

    g.SetFont("s10 bold", "Segoe UI")
    btnOk := g.Add("Button", "Default w140 h36 x" (largura - 310) " y" (posY + 10), "Confirmar ➔")
    btnCancelar := g.Add("Button", "w140 h36 x+10", "Cancelar")

    ConfirmarClick(*) {
        escolhidas := []
        for cb in combos
            escolhidas.Push(cb.Text)
        resultado := escolhidas
        confirmado := true
        g.Destroy()
    }

    btnOk.OnEvent("Click", ConfirmarClick)
    btnCancelar.OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show("w" largura)
    WinWaitClose("ahk_id " g.Hwnd)

    return confirmado ? resultado : ""
}

EscolherModoEnvioUI(largura := 640) {
    resultado := "semi"

    g := Gui("+AlwaysOnTop -MaximizeBox", "Passo 8 — Modo de Envio")
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F7F9FC"

    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w" (largura - 40) " c003366", "⚡ Escolha o Modo de Envio")

    g.Add("Text", "x20 y48 w" (largura - 40) " h2 0x10")

    g.SetFont("s10 norm", "Segoe UI")
    g.Add("Text", "x20 y58 w" (largura - 40) " c222222", 
        "O sistema anexa os PDFs automaticamente no Outlook Web. Como você deseja conduzir os envios?"
    )

    g.SetFont("s10 bold", "Segoe UI")
    g.Add("GroupBox", "x20 y110 w" (largura - 40) " h95 c003366", " 👤 1. Semi-Automático / Supervisionado (Recomendado) ")
    g.SetFont("s9 norm", "Segoe UI")
    g.Add("Text", "x35 y135 w" (largura - 70) " c333333", 
        "O script abre o Outlook Web e anexa todos os certificados do docente automaticamente.`n"
        "Você apenas visualiza a tela e aperta ENTER (ou clica no botão) para disparar o e-mail e avançar.`n"
        "Máxima velocidade com total controle visual."
    )

    g.SetFont("s10 bold", "Segoe UI")
    g.Add("GroupBox", "x20 y215 w" (largura - 40) " h95 c003366", " 🚀 2. Modo Turbo (100% Automático) ")
    g.SetFont("s9 norm", "Segoe UI")
    g.Add("Text", "x35 y240 w" (largura - 70) " c333333", 
        "O script abre, anexa os PDFs, aguarda 2s e clica em 'Enviar' sozinho para toda a lista de docentes.`n"
        "Execução autônoma contínua. Você pode pausar a qualquer momento pressionando a tecla ESC."
    )

    g.SetFont("s10 bold", "Segoe UI")
    btnSemi := g.Add("Button", "Default w260 h42 x40 y325", "👤 1. Modo Supervisionado (ENTER)")
    btnAuto := g.Add("Button", "w220 h42 x+15", "🚀 2. Modo Turbo (100% Auto)")
    btnCancelar := g.Add("Button", "w70 h42 x+15", "✖")

    btnSemi.OnEvent("Click", (*) => (resultado := "semi", g.Destroy()))
    btnAuto.OnEvent("Click", (*) => (resultado := "turbo", g.Destroy()))
    btnCancelar.OnEvent("Click", (*) => (resultado := "semi", g.Destroy()))
    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show("w" largura)
    WinWaitClose("ahk_id " g.Hwnd)

    return resultado
}

ExibirPainelSupervisionadoUI(idxDocente, totalDocentes, nomeDocente, emailDocente, totalArquivos, nomeCurso, largura := 580) {
    acao := "enviar"

    g := Gui("+AlwaysOnTop -MaximizeBox", "Envio Supervisionado — Outlook Web")
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F7F9FC"

    g.SetFont("s12 bold", "Segoe UI")
    g.Add("Text", "x20 y14 w" (largura - 40) " c003366", "📬 Docente " idxDocente " de " totalDocentes " — " nomeDocente)

    g.Add("Text", "x20 y42 w" (largura - 40) " h2 0x10")

    g.SetFont("s9 norm", "Segoe UI")
    g.Add("Text", "x20 y50 w" (largura - 40) " c444444", "Curso: " nomeCurso "`nE-mail: " emailDocente)

    g.SetFont("s10 bold", "Segoe UI")
    g.Add("Text", "x20 y90 w" (largura - 40) " c1B5E20", "✅ " totalArquivos " certificado(s) anexado(s) automaticamente!")

    g.SetFont("s9 italic", "Segoe UI")
    g.Add("Text", "x20 y112 w" (largura - 40) " c666666", "Confira o anexo no Outlook e pressione ENTER (ou clique abaixo) para enviar:")

    g.SetFont("s10 bold", "Segoe UI")
    btnEnviar := g.Add("Button", "Default w240 h40 x20 y140", "🚀 Enviar e Próximo (ENTER) ➔")
    btnPular  := g.Add("Button", "w140 h40 x+10", "⏭️ Pular Docente")
    btnParar  := g.Add("Button", "w140 h40 x+10", "🛑 Pausar Rodada")

    btnEnviar.OnEvent("Click", (*) => (acao := "enviar", g.Destroy()))
    btnPular.OnEvent("Click", (*) => (acao := "pular", g.Destroy()))
    btnParar.OnEvent("Click", (*) => (acao := "parar", g.Destroy()))

    g.OnEvent("Close", (*) => (acao := "parar", g.Destroy()))
    g.OnEvent("Escape", (*) => (acao := "parar", g.Destroy()))

    g.Show("w" largura " y50")
    WinWaitClose("ahk_id " g.Hwnd)

    return acao
}

; ════════════════════════════════════════════════════════════════════════
;   FASE 1 — MONTAGEM E CONFIGURAÇÃO DA RODADA DE ENVIO
; ════════════════════════════════════════════════════════════════════════

cursosRodada := []
primeiraVez  := true

Loop {
    if primeiraVez {
        ExibirMensagem(
            "Passo 1 — Central de Envio",
            "Envio Automático de Certificados por E-mail",
            "Este assistente prepara e envia os certificados em PDF com ANEXAÇÃO AUTOMÁTICA no Outlook Web.`n`n"
            "Você pode montar uma RODADA DE ENVIO com um ou vários cursos de uma vez só.`n`n"
            "Para cada curso, o sistema irá:`n"
            "  1. Selecionar a pasta dos certificados gerados;`n"
            "  2. Selecionar o período/meses desejados;`n"
            "  3. Ler a planilha do curso com detecção automática de colunas;`n"
            "  4. Anexar e enviar os e-mails com controle automático de progresso.`n`n"
            "Clique abaixo para começar selecionando o primeiro curso.",
            "passo",
            "Iniciar Rodada ➔",
            640
        )
        primeiraVez := false
    } else {
        respAdd := Confirmar(
            "Adicionar Outro Curso?",
            "Adicionar Mais Cursos à Rodada",
            "Curso adicionado com sucesso à rodada de envio!`n`n"
            "Cursos já inclusos nesta rodada: " cursosRodada.Length "`n`n"
            "Deseja adicionar mais um curso antes de iniciar os envios?",
            "➕ Adicionar Outro Curso",
            "🚀 Seguir para o Envio",
            true,
            600
        )
        if !respAdd
            break
    }

    numCursoAtual := cursosRodada.Length + 1

    ; ─── SELEÇÃO DA PASTA DO CURSO ──────────────────────────────────────
    pastaPosgrad := DirSelect(PASTA_PADRAO_CERTS, 3, "Passo 1 — Selecione a pasta da pós-graduação [Curso " numCursoAtual "]")
    if !pastaPosgrad {
        if (cursosRodada.Length = 0) {
            ExibirMensagem("Operação Cancelada", "Nenhuma Pasta Selecionada", "Nenhuma pasta de curso foi selecionada. O programa será encerrado.", "erro", "Fechar")
            ExitApp
        } else {
            break
        }
    }
    pastaPosgrad := RTrim(pastaPosgrad, "\") "\"

    mesesDisponiveis := ListarSubpastas(pastaPosgrad)
    if (mesesDisponiveis.Length = 0) {
        ExibirMensagem(
            "Sem Subpastas de Mês",
            "Nenhum Mês Encontrado",
            "Nenhuma subpasta de mês (ex: 04.26, 05.26) foi encontrada dentro de:`n`n" pastaPosgrad "`n`n"
            "Este curso não pôde ser adicionado.",
            "erro",
            "Continuar"
        )
        continue
    }

    ; ─── SELEÇÃO DOS MESES / PERÍODO ────────────────────────────────────
    mesesSelecionados := SelecionarOpcoesUI(
        "Passo 2 — Selecionar Período (Meses)",
        "Meses Disponíveis — Curso " numCursoAtual,
        mesesDisponiveis,
        false,
        600
    )

    if (mesesSelecionados.Length = 0) {
        ExibirMensagem("Seleção Vazia", "Nenhum Mês Selecionado", "Nenhum mês foi selecionado. Este curso não será adicionado à rodada.", "erro", "Continuar")
        continue
    }

    ; ─── SELEÇÃO AUTOMÁTICA DA PLANILHA DO CURSO (ONEDRIVE) ───────────────
    infoCurso := ExtrairDadosCursoDeCaminho(pastaPosgrad)
    excelPath := LocalizarPlanilhaCursoDinamica(infoCurso["nomeCurso"], infoCurso["unidade"], infoCurso["semestre"])

    if (excelPath = "" || !FileExist(excelPath))
        excelPath := LocalizarPlanilhaCursoDinamica(infoCurso["nomePasta"])

    if (excelPath != "" && FileExist(excelPath)) {
        SplitPath(excelPath, &nomeArquivoPlanilha, &dirPlanilha)
        usarAuto := Confirmar(
            "Passo 3 — Planilha do Curso",
            "Planilha Localizada na Nuvem",
            "Curso: " infoCurso["nomeCurso"] (infoCurso["unidade"] != "" ? " (" infoCurso["unidade"] ")" : "") "`n`n"
            "O sistema localizou automaticamente no OneDrive:`n"
            "📁 " nomeArquivoPlanilha "`n`n"
            "Caminho:`n" excelPath "`n`n"
            "Deseja utilizar esta planilha para envio?",
            "✅ Usar Esta Planilha",
            "📁 Escolher Outra Manualmente",
            true,
            620
        )
        if !usarAuto {
            pastaInicio := dirPlanilha != "" ? dirPlanilha : (INFO_NUVEM["raizOneDrive"] != "" ? INFO_NUVEM["raizOneDrive"] : "")
            excelPath := FileSelect(1, pastaInicio, "Passo 3 — Selecione a planilha (.xlsx) com docentes e e-mails [Curso " numCursoAtual "]", "*.xlsx")
            if !excelPath {
                ExibirMensagem("Operação Cancelada", "Nenhuma Planilha Selecionada", "Nenhuma planilha foi selecionada. Este curso não será adicionado.", "erro", "Continuar")
                continue
            }
        }
    } else {
        pastaInicioExcel := INFO_NUVEM["raizOneDrive"] != "" ? INFO_NUVEM["raizOneDrive"] : ""
        excelPath := FileSelect(1, pastaInicioExcel, "Passo 3 — Selecione a planilha (.xlsx) com docentes e e-mails [Curso " numCursoAtual "]", "*.xlsx")
        if !excelPath {
            ExibirMensagem("Operação Cancelada", "Nenhuma Planilha Selecionada", "Nenhuma planilha foi selecionada. Este curso não será adicionado.", "erro", "Continuar")
            continue
        }
    }

    excelTemp := ""
    wbTemp := ""
    try {
        excelTemp := ComObject("Excel.Application")
        excelTemp.Visible := false
        wbTemp := excelTemp.Workbooks.Open(excelPath)
    } catch as err {
        ExibirMensagem("Erro no Excel", "Falha ao Abrir Planilha", "Não foi possível abrir a planilha selecionada:`n`n" err.Message, "erro", "Continuar")
        if IsObject(excelTemp)
            try excelTemp.Quit()
        continue
    }

    abasWb := []
    Loop wbTemp.Sheets.Count
        abasWb.Push(Trim(wbTemp.Sheets.Item(A_Index).Name))

    ; ─── CASAMENTO AUTOMÁTICO DE ABAS ──────────────────────────────────
    abasCorrespondentes := []
    todasCasaram := true
    for mes in mesesSelecionados {
        achou := ""
        for aba in abasWb {
            if AbaCorrespondeAoMes(aba, mes) {
                achou := aba
                break
            }
        }
        if (achou = "") {
            todasCasaram := false
            break
        }
        abasCorrespondentes.Push(achou)
    }

    if !todasCasaram {
        abasCorrespondentes := CorrigirAbasUI(
            "Passo 4 — Corresponder Abas",
            "Ajuste de Abas da Planilha [Curso " numCursoAtual "]",
            mesesSelecionados,
            abasWb,
            600
        )
        if (abasCorrespondentes = "" || abasCorrespondentes.Length = 0) {
            try wbTemp.Close(false)
            try excelTemp.Quit()
            ExibirMensagem("Ajuste Cancelado", "Abas Não Confirmadas", "A correspondência de abas foi cancelada. Este curso não será adicionado.", "erro", "Continuar")
            continue
        }
    }

    ; ─── IDENTIFICAÇÃO DO NOME DA PÓS-GRADUAÇÃO (CÉLULA A2) ────────────
    wsPrimeira := ""
    Loop wbTemp.Sheets.Count {
        if (Trim(wbTemp.Sheets.Item(A_Index).Name) = abasCorrespondentes[1]) {
            wsPrimeira := wbTemp.Sheets.Item(A_Index)
            break
        }
    }

    nomePosGraduacao := ""
    if wsPrimeira {
        try nomePosGraduacao := Trim(String(wsPrimeira.Range("A2").Value))
    }

    if (nomePosGraduacao = "") {
        try wbTemp.Close(false)
        try excelTemp.Quit()
        ExibirMensagem("Planilha Incompleta", "Célula A2 Vazia", "A célula A2 da primeira aba está vazia. O nome da pós-graduação é obrigatório.", "erro", "Continuar")
        continue
    }

    ; ─── DETECÇÃO AUTOMÁTICA DE COLUNAS (DOCENTES E E-MAILS) ───────────
    colNomeLetra := IdentificarColunaDocentes(wsPrimeira)
    colEmailLetra := IdentificarColunaEmails(wsPrimeira)

    if (colNomeLetra = "") {
        cxNome := PedirTexto(
            "Passo 6 — Coluna dos Docentes",
            "Coluna de Nomes não Detectada",
            "[Curso " numCursoAtual "] Digite a letra da coluna onde constam os NOMES dos docentes na planilha:",
            "E",
            "Exemplo: E"
        )
        if (cxNome["Result"] != "OK" || Trim(cxNome["Value"]) = "") {
            try wbTemp.Close(false)
            try excelTemp.Quit()
            continue
        }
        colNomeLetra := StrUpper(Trim(cxNome["Value"]))
    }

    if (colEmailLetra = "") {
        cxEmail := PedirTexto(
            "Passo 6 — Coluna dos E-mails",
            "Coluna de E-mails não Detectada",
            "[Curso " numCursoAtual "] Digite a letra da coluna onde constam os E-MAILS dos docentes na planilha:",
            "F",
            "Exemplo: F"
        )
        if (cxEmail["Result"] != "OK" || Trim(cxEmail["Value"]) = "") {
            try wbTemp.Close(false)
            try excelTemp.Quit()
            continue
        }
        colEmailLetra := StrUpper(Trim(cxEmail["Value"]))
    }

    ; ─── LEITURA E PROCESSAMENTO DOS DOCENTES ───────────────────────────
    mapaDocentesCurso := Map()
    ordemDocentesCurso := []

    for idxAba, abaNome in abasCorrespondentes {
        mesRotulo := mesesSelecionados[idxAba]
        wsAba := ""
        Loop wbTemp.Sheets.Count {
            if (Trim(wbTemp.Sheets.Item(A_Index).Name) = abaNome) {
                wsAba := wbTemp.Sheets.Item(A_Index)
                break
            }
        }
        if !wsAba
            continue

        colNomeAba := IdentificarColunaDocentes(wsAba)
        colEmailAba := IdentificarColunaEmails(wsAba)
        if (colNomeAba = "")
            colNomeAba := colNomeLetra
        if (colEmailAba = "")
            colEmailAba := colEmailLetra

        ultimaLinha := wsAba.Cells(wsAba.Rows.Count, colNomeAba).End(-4162).Row

        Loop ultimaLinha - 4 {
            linha := A_Index + 4
            nomeBruto := Trim(String(wsAba.Cells(linha, colNomeAba).Value))
            emailBruto := Trim(String(wsAba.Cells(linha, colEmailAba).Value))
            if (nomeBruto = "")
                continue

            nomeLimpo := SanitizarNomeDocente(nomeBruto)
            if (nomeLimpo = "")
                continue

            chave := StrLower(nomeLimpo)
            emailExtraido := ExtrairEmailValido(emailBruto)

            if !mapaDocentesCurso.Has(chave) {
                emailValido := EmailEhValido(emailExtraido)
                mapaDocentesCurso[chave] := Map(
                    "nome", nomeLimpo,
                    "email", emailExtraido,
                    "emailValido", emailValido,
                    "origemMes", mesRotulo,
                    "origemAba", abaNome,
                    "origemLinha", linha
                )
                ordemDocentesCurso.Push(chave)
            } else {
                if (!mapaDocentesCurso[chave]["emailValido"] && EmailEhValido(emailExtraido)) {
                    mapaDocentesCurso[chave]["email"] := emailExtraido
                    mapaDocentesCurso[chave]["emailValido"] := true
                }
            }
        }
    }

    try wbTemp.Close(false)
    try excelTemp.Quit()

    if (ordemDocentesCurso.Length = 0) {
        ExibirMensagem("Sem Docentes", "Nenhum Docente Encontrado", "Nenhum docente válido foi encontrado nas abas selecionadas deste curso.", "erro", "Continuar")
        continue
    }

    cursosRodada.Push(Map(
        "pastaPosgrad", pastaPosgrad,
        "nomePosGraduacao", nomePosGraduacao,
        "mesesSelecionados", mesesSelecionados,
        "mapaDocentes", mapaDocentesCurso,
        "ordemDocentes", ordemDocentesCurso
    ))
}

if (cursosRodada.Length = 0) {
    ExibirMensagem("Encerrando", "Nenhum Curso Configurado", "Nenhum curso foi configurado para envio. O programa será encerrado.", "erro", "Fechar")
    ExitApp
}

; ════════════════════════════════════════════════════════════════════════
;   FASE 2 — VALIDAÇÃO E CORREÇÃO DE E-MAILS PENDENTES
; ════════════════════════════════════════════════════════════════════════

for curso in cursosRodada {
    for chave in curso["ordemDocentes"] {
        dados := curso["mapaDocentes"][chave]
        if !dados["emailValido"] {
            valorLido := (dados["email"] = "") ? "(em branco)" : dados["email"]

            cxCorr := PedirTexto(
                "Passo 7 — Corrigir E-mail",
                "E-mail Ausente ou Inválido",
                "Docente: " dados["nome"] "`n"
                "Curso: " curso["nomePosGraduacao"] "`n"
                "Mês: " dados["origemMes"] " (Aba: '" dados["origemAba"] "' — Linha " dados["origemLinha"] ")`n"
                "Valor lido na planilha: " valorLido "`n`n"
                "Digite o e-mail correto abaixo (ou deixe em branco para pular este docente nesta rodada):",
                (EmailEhValido(dados["email"]) ? dados["email"] : ""),
                "Exemplo: docente@einstein.br",
                600
            )

            if (cxCorr["Result"] = "OK" && EmailEhValido(Trim(cxCorr["Value"]))) {
                dados["email"] := Trim(cxCorr["Value"])
                dados["emailValido"] := true
            }
        }
    }
}

; ════════════════════════════════════════════════════════════════════════
;   FASE 3 — SELEÇÃO DO MODO E CONFIRMAÇÃO DA RODADA
; ════════════════════════════════════════════════════════════════════════

modoEnvioEscolhido := EscolherModoEnvioUI(640)

resumoRodada := ""
totalDocentesGeral := 0

for idx, curso in cursosRodada {
    resumoRodada .= idx ". " curso["nomePosGraduacao"] "`n"
    resumoRodada .= "   Meses: " JoinArr(curso["mesesSelecionados"], ", ") "`n"
    resumoRodada .= "   Docentes mapeados: " curso["ordemDocentes"].Length "`n`n"
    totalDocentesGeral += curso["ordemDocentes"].Length
}

modoTextoDesc := (modoEnvioEscolhido = "turbo") ? "⚡ Modo Turbo (100% Automático)" : "👤 Modo Semi-Automático (Supervisionado com ENTER)"

confirmFinal := Confirmar(
    "Passo 9 — Confirmar Rodada de Envio",
    "Resumo da Rodada de Envio",
    "Cursos configurados na rodada: " cursosRodada.Length "`n"
    "Total de docentes mapeados: " totalDocentesGeral "`n"
    "Modo de envio: " modoTextoDesc "`n`n"
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n"
    resumoRodada
    "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n"
    "O controle automático (_EnvioControle.txt) identificará quem já foi enviado antes e pulará automaticamente.`n`n"
    "Deseja iniciar os envios agora?",
    "🚀 Iniciar Envio",
    "❌ Cancelar",
    true,
    680
)

if !confirmFinal
    ExitApp

; ════════════════════════════════════════════════════════════════════════
;   FASE 4 — EXECUÇÃO DO ENVIO (AUTOMATIZADO)
; ════════════════════════════════════════════════════════════════════════

assuntoPadrao := "Envio de certificados Pós Graduação"

totalGeralAbertos    := 0
totalGeralPulados    := 0
totalGeralJaEnviados := 0
relatorioFinalCursos := ""
interromperGeral     := false

for idxCurso, curso in cursosRodada {

    if interromperGeral
        break

    pastaPosgrad      := curso["pastaPosgrad"]
    mesesSelecionados := curso["mesesSelecionados"]
    nomePosGraduacao  := curso["nomePosGraduacao"]
    mapaDocentes      := curso["mapaDocentes"]
    ordemDocentes     := curso["ordemDocentes"]

    pastaControle := pastaPosgrad CONTROLE_ARQUIVO
    enviadosMap   := CarregarEnviados(pastaControle)

    ordemPendente := []
    totalJaEnviadoCurso := 0

    for chave in ordemDocentes {
        if EstaTotalmenteEnviado(enviadosMap, chave, mesesSelecionados) {
            totalJaEnviadoCurso++
            continue
        }
        ordemPendente.Push(chave)
    }
    totalGeralJaEnviados += totalJaEnviadoCurso

    if (ordemPendente.Length = 0) {
        ExibirMensagem(
            "Curso Já Concluído",
            "Todos os Docentes Enviados",
            "Curso " idxCurso "/" cursosRodada.Length ": " nomePosGraduacao "`n`n"
            "Todos os docentes deste curso já constam como enviados em rodadas anteriores.`n"
            "Avançando para o próximo curso...",
            "sucesso",
            "Continuar ➔"
        )
        relatorioFinalCursos .= idxCurso ". " nomePosGraduacao "`n     Status: 100% concluído anteriormente (" totalJaEnviadoCurso " docentes)`n`n"
        continue
    }

    corpoEmail := "Boa tarde!%0D%0A%0D%0ASegue certificado(s) referente às aulas da pós-graduação em " nomePosGraduacao "."

    totalAbertosCurso := 0
    totalPuladosCurso := 0

    for idx, chave in ordemPendente {

        dados := mapaDocentes[chave]
        nomeDocente := dados["nome"]
        emailReal := dados["email"]

        if !dados["emailValido"] {
            totalPuladosCurso++
            continue
        }

        ; 1. Localiza todos os certificados em PDF do docente para os meses selecionados
        arquivosPdf := LocalizarArquivosDocente(pastaPosgrad, mesesSelecionados, nomeDocente, enviadosMap)

        if (arquivosPdf.Length = 0) {
            ExibirMensagem(
                "Certificado Não Localizado",
                "Sem PDFs Pendentes",
                "Não foram localizados certificados PDF para:`n`n"
                "  " nomeDocente " <" emailReal ">`n`n"
                "Este docente será pulado.",
                "erro",
                "Continuar ➔"
            )
            totalPuladosCurso++
            continue
        }

        ; 2. Abre o Outlook Web via Deeplink
        link := "https://outlook.office.com/mail/deeplink/compose?to=" emailReal
             . "&subject=" assuntoPadrao
             . "&body=" corpoEmail
        Run(link)

        ; 3. Obtém o handle da janela do navegador
        hwndNavegador := ObterJanelaNavegadorOutlook(18)

        ; 4. Anexa automaticamente os PDFs no e-mail (com espera ativa de carregamento da página)
        sucessoAnexo := AnexarArquivosOutlookWeb(hwndNavegador, arquivosPdf, 25)

        if !sucessoAnexo {
            respErro := Confirmar(
                "Aviso de Anexação",
                "Anexo Não Confirmado Automaticamente",
                "Não foi possível confirmar a anexação automática para:`n`n"
                "  Docente: " nomeDocente "`n"
                "  E-mail: " emailReal "`n`n"
                "Deseja tentar anexar manualmente ou pular este docente?",
                "Tentei Manualmente (Avançar)",
                "⏭️ Pular Docente",
                true,
                600
            )
            if !respErro {
                totalPuladosCurso++
                continue
            }
        }

        totalAbertosCurso++

        ; 5. Fluxo de Envio conforme o modo escolhido
        if (modoEnvioEscolhido = "turbo") {
            Sleep(2000)
            DispararEnvioOutlookWeb(hwndNavegador)
            MarcarPendentesComoEnviados(pastaControle, nomeDocente, chave, mesesSelecionados, enviadosMap)
            Sleep(1500)

        } else {
            acaoSupervisor := ExibirPainelSupervisionadoUI(
                idx, ordemPendente.Length, nomeDocente, emailReal,
                arquivosPdf.Length, nomePosGraduacao, 580
            )

            if (acaoSupervisor = "enviar") {
                DispararEnvioOutlookWeb(hwndNavegador)
                MarcarPendentesComoEnviados(pastaControle, nomeDocente, chave, mesesSelecionados, enviadosMap)
                Sleep(800)
            } else if (acaoSupervisor = "pular") {
                totalPuladosCurso++
            } else if (acaoSupervisor = "parar") {
                interromperGeral := true
                break
            }
        }
    }

    totalGeralAbertos += totalAbertosCurso
    totalGeralPulados += totalPuladosCurso

    relatorioFinalCursos .= idxCurso ". " nomePosGraduacao "`n"
    relatorioFinalCursos .= "     Enviados: " totalAbertosCurso " | Pulados: " totalPuladosCurso " | Já enviados antes: " totalJaEnviadoCurso "`n`n"

    if interromperGeral {
        ExibirMensagem("Rodada Interrompida", "Envios Pausados pelo Usuário", "A rodada foi pausada conforme solicitado. O progresso foi salvo com segurança no arquivo de controle.", "info", "Ver Relatório")
        break
    }
}

; ════════════════════════════════════════════════════════════════════════
;   FASE 5 — RELATÓRIO FINAL DA RODADA (UI_MODERNA)
; ════════════════════════════════════════════════════════════════════════

resumoFinalEnvio := "Cursos processados:              " cursosRodada.Length "`n"
resumoFinalEnvio .= "E-mails enviados:                " totalGeralAbertos "`n"
resumoFinalEnvio .= "Pulados nesta rodada:            " totalGeralPulados "`n"
resumoFinalEnvio .= "Já enviados em rodadas passadas: " totalGeralJaEnviados "`n"
resumoFinalEnvio .= "Modo utilizado:                  " ((modoEnvioEscolhido = "turbo") ? "Modo Turbo (100% Auto)" : "Modo Supervisionado")

ExibirRelatorioFinalUI(
    "Passo 10 — Relatório Final da Rodada",
    "Rodada de Envio Concluída!",
    resumoFinalEnvio,
    Trim(relatorioFinalCursos, "`n"),
    PASTA_PADRAO_CERTS,
    700
)

ExitApp
