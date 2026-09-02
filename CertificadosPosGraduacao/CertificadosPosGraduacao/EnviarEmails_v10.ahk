#Requires AutoHotkey v2.0
#SingleInstance Force
#Include UI_Moderna.ahk

; ════════════════════════════════════════════════════════════════════════
;   ENVIAR E-MAILS — PÓS-GRADUAÇÃO (v10)
;   Versão Padronizada: Detecção Automática de Colunas & UI Moderna
; ════════════════════════════════════════════════════════════════════════
;
;   DESTAQUES DESTA VERSÃO (v10):
;   - Interface 100% Padronizada com UI_Moderna.ahk:
;     Todas as caixas de mensagem, formulários de seleção em lote, confirmações,
;     guias de envio no Outlook e relatório final agora utilizam a biblioteca
;     UI_Moderna.ahk (Segoe UI, tons #F7F9FC e #003366).
;   - Detecção Automática de Colunas (Docente e E-mail):
;     O script analisa o cabeçalho (Linha 3) da planilha do curso e identifica
;     automaticamente onde estão as colunas de "Docente" e "E-mail", eliminando
;     a necessidade de digitar letras a cada curso.
;   - Limpeza e Normalização Uniforme de Nomes:
;     Tratamento idêntico ao de GerarCertificados_v16.ahk (remoção de títulos
;     Dr./Dra./Profa., sufixos [H/M] e parênteses), assegurando que o nome
;     lido na planilha coincida perfeitamente com os PDFs gerados.
;   - Controle Automático de Progresso (_EnvioControle.txt):
;     Registra cada envio efetuado por docente e mês, permitindo executar
;     o envio em partes sem risco de duplicação.
;   - Fechamento Seguro de Processos COM (Excel):
;     Protegido com try em todos os pontos de saída para evitar processos
;     órfãos na memória.
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
        ; Ignora pastas internas do sistema como _EnvioTemp e Docentes se não forem meses
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

; Limpeza robusta do nome do docente — 100% compatível com GerarCertificados_v16.ahk
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

; =========================================================
; MESES E CORRESPONDÊNCIA TOLERANTE COM ABAS
; =========================================================

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

EncontrarPastaDocente(pastaBase, nomeAlvo) {
    alvoNorm := NormalizarNome(nomeAlvo)
    melhorMatch := ""
    if !DirExist(pastaBase)
        return ""
    Loop Files, pastaBase "*", "D" {
        nomeNorm := NormalizarNome(A_LoopFileName)
        if (nomeNorm = alvoNorm)
            return A_LoopFilePath "\"
        if (melhorMatch = "") && (InStr(nomeNorm, alvoNorm) || InStr(alvoNorm, nomeNorm))
            melhorMatch := A_LoopFilePath "\"
    }
    return melhorMatch
}

ContarArquivosPdf(pasta) {
    if (!pasta || !DirExist(pasta))
        return 0
    total := 0
    Loop Files, pasta "*.pdf"
        total++
    return total
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
    ; Prioridade 1: Linha 3 (Padrão das planilhas)
    Loop maxCols {
        colIdx := A_Index
        celVal := ""
        try celVal := String(wsAba.Cells(3, colIdx).Value)
        if (celVal = "")
            try celVal := String(wsAba.Cells(3, colIdx).Text)

        celValNorm := NormalizarTexto(celVal)
        for t in termos {
            if InStr(celValNorm, t)
                return IndiceParaColuna(colIdx)
        }
    }

    ; Prioridade 2: Linhas 4, 2, 1 como fallback
    for linhaCheck in [4, 2, 1] {
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
    ; Prioridade 1: Linha 3 (Padrão das planilhas)
    Loop maxCols {
        colIdx := A_Index
        celVal := ""
        try celVal := String(wsAba.Cells(3, colIdx).Value)
        if (celVal = "")
            try celVal := String(wsAba.Cells(3, colIdx).Text)

        celValNorm := NormalizarTexto(celVal)
        for t in termos {
            if InStr(celValNorm, t)
                return IndiceParaColuna(colIdx)
        }
    }

    ; Prioridade 2: Linhas 4, 2, 1 como fallback
    for linhaCheck in [4, 2, 1] {
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
; COMPONENTES VISUAIS MODERNOS (CUSTOMIZADOS PARA O ENVIO)
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

EscolherModoEnvioUI(largura := 620) {
    resultado := true

    g := Gui("+AlwaysOnTop -MaximizeBox", "Passo 8 — Modo de Envio")
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F7F9FC"

    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w" (largura - 40) " c003366", "✉️ Escolha o Modo de Envio")

    g.Add("Text", "x20 y48 w" (largura - 40) " h2 0x10")

    g.SetFont("s10 norm", "Segoe UI")
    g.Add("Text", "x20 y58 w" (largura - 40) " c222222", 
        "Defina como você deseja que as abas do Outlook sejam abertas durante o processo:`n"
        "(Em ambos os modos, você anexa os arquivos e confirma o envio com total controle.)"
    )

    g.SetFont("s10 bold", "Segoe UI")
    g.Add("GroupBox", "x20 y115 w" (largura - 40) " h85 c003366", " Opção 1: Um a um (Recomendado) ")
    g.SetFont("s9 norm", "Segoe UI")
    g.Add("Text", "x35 y140 w" (largura - 70) " c444444", 
        "Abre o e-mail no Outlook e a pasta de PDFs do docente atual. Você anexa, clica em enviar e avança para o próximo. Mais organizado e seguro para grandes volumes."
    )

    g.SetFont("s10 bold", "Segoe UI")
    g.Add("GroupBox", "x20 y210 w" (largura - 40) " h85 c003366", " Opção 2: Todos de uma vez ")
    g.SetFont("s9 norm", "Segoe UI")
    g.Add("Text", "x35 y235 w" (largura - 70) " c444444", 
        "Abre todas as abas de e-mail do curso de uma vez só no navegador e depois abre as pastas de PDFs sequencialmente para que você faça os anexos."
    )

    g.SetFont("s10 bold", "Segoe UI")
    btnUmAUm := g.Add("Button", "Default w220 h40 x40 y310", "👤 1. Enviar Um a um")
    btnTodos := g.Add("Button", "w220 h40 x+20", "⚡ 2. Todos de uma vez")
    btnCancelar := g.Add("Button", "w80 h40 x+20", "✖ Fechar")

    btnUmAUm.OnEvent("Click", (*) => (resultado := true, g.Destroy()))
    btnTodos.OnEvent("Click", (*) => (resultado := false, g.Destroy()))
    btnCancelar.OnEvent("Click", (*) => (resultado := true, g.Destroy()))
    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show("w" largura)
    WinWaitClose("ahk_id " g.Hwnd)

    return resultado
}

ExibirGuiaEnvioDocenteUI(idxCurso, totalCursos, nomeCurso, idxDocente, totalDocentes, nomeDocente, emailDocente, pastaDocente, largura := 640) {
    acao := "enviado"
    totalPdfs := ContarArquivosPdf(pastaDocente)

    g := Gui("+AlwaysOnTop -MaximizeBox", "Envio de Certificados — Outlook")
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F7F9FC"

    ; Cabeçalho
    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w" (largura - 40) " c003366", "📬 Envio de Certificado (" idxDocente " de " totalDocentes ")")

    g.Add("Text", "x20 y48 w" (largura - 40) " h2 0x10")

    ; Resumo do Docente e Curso
    g.SetFont("s10 bold", "Segoe UI")
    g.Add("Text", "x20 y58 w" (largura - 40) " c003366", "Curso " idxCurso "/" totalCursos ": " nomeCurso)
    
    g.SetFont("s10 norm", "Segoe UI")
    g.Add("Text", "x20 y82 w" (largura - 40) " c222222", 
        "Destinatário: " nomeDocente " <" emailDocente ">`n"
        "Arquivos PDF localizados: " totalPdfs " certificado(s) na pasta."
    )

    ; Caixa de instruções passo a passo
    g.SetFont("s10 bold", "Segoe UI")
    g.Add("GroupBox", "x20 y135 w" (largura - 40) " h150 c003366", " Siga os passos para concluir o envio: ")

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x35 y162 w" (largura - 70) " c003366", "1. Na janela do Explorer que acabou de abrir:")
    g.SetFont("s9 norm", "Segoe UI")
    g.Add("Text", "x50 y180 w" (largura - 90) " c444444", "Pressione Ctrl+A para selecionar todos os arquivos PDF do docente.")

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x35 y202 w" (largura - 70) " c003366", "2. No navegador (Outlook Web):")
    g.SetFont("s9 norm", "Segoe UI")
    g.Add("Text", "x50 y220 w" (largura - 90) " c444444", "Arraste os arquivos selecionados para a área de anexos do e-mail.")

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x35 y242 w" (largura - 70) " c003366", "3. Finalizar:")
    g.SetFont("s9 norm", "Segoe UI")
    g.Add("Text", "x50 y260 w" (largura - 90) " c444444", "Confira o destinatário/anexos, clique em 'Enviar' no Outlook e confirme abaixo.")

    ; Botões de Ação
    g.SetFont("s10 bold", "Segoe UI")
    btnOk := g.Add("Button", "Default w230 h40 x20 y300", "✅ E-mail Enviado ➔")
    btnPular := g.Add("Button", "w170 h40 x+10", "⏭️ Pular este Docente")
    btnParar := g.Add("Button", "w160 h40 x+10", "🛑 Interromper Rodada")

    btnOk.OnEvent("Click", (*) => (acao := "enviado", g.Destroy()))
    btnPular.OnEvent("Click", (*) => (acao := "pular", g.Destroy()))
    btnParar.OnEvent("Click", (*) => (acao := "parar", g.Destroy()))

    g.OnEvent("Close", (*) => (acao := "parar", g.Destroy()))
    g.OnEvent("Escape", (*) => (acao := "parar", g.Destroy()))

    g.Show("w" largura)
    WinWaitClose("ahk_id " g.Hwnd)

    return acao
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

ConstruirPastaTempCurso(pastaPosgrad, mesesSelecionados, enviadosMap) {
    pastaTemp := pastaPosgrad "_EnvioTemp\"
    if DirExist(pastaTemp)
        try DirDelete(pastaTemp, true)
    DirCreate(pastaTemp)

    totalCopiados := 0
    totalPulados := 0

    for mes in mesesSelecionados {
        pastaMes := pastaPosgrad mes "\"
        if !DirExist(pastaMes)
            continue

        Loop Files, pastaMes "*.pdf" {
            nomeDocente := RegExReplace(A_LoopFileName, "\s*\([^)]*\)(\.\d+)?\.pdf$", "")
            nomeDocente := SanitizarNomeDocente(nomeDocente)
            if (nomeDocente = "")
                continue

            chave := StrLower(nomeDocente)
            if enviadosMap.Has(chave) && enviadosMap[chave].Has(mes) {
                totalPulados++
                continue
            }

            pastaDestino := pastaTemp nomeDocente "\"
            DirCreate(pastaDestino)
            destinoArquivo := pastaDestino A_LoopFileName

            if FileExist(destinoArquivo) {
                base := RegExReplace(A_LoopFileName, "\.pdf$", "")
                contador := 1
                Loop {
                    destinoArquivo := pastaDestino base "_" contador ".pdf"
                    if !FileExist(destinoArquivo)
                        break
                    contador++
                }
            }

            try {
                FileCopy(A_LoopFilePath, destinoArquivo)
                totalCopiados++
            }
        }
    }

    return Map("pasta", pastaTemp, "totalCopiados", totalCopiados, "totalPulados", totalPulados)
}

; ════════════════════════════════════════════════════════════════════════
;   FASE 1 — MONTAGEM E CONFIGURAÇÃO DA RODADA DE ENVIO
; ════════════════════════════════════════════════════════════════════════

cursosRodada  := []
primeiraVez   := true

Loop {
    if primeiraVez {
        ExibirMensagem(
            "Passo 1 — Central de Envio",
            "Envio Automático de Certificados por E-mail",
            "Este assistente prepara e envia os certificados em PDF gerados para cada docente.`n`n"
            "Você pode montar uma RODADA DE ENVIO com um ou vários cursos de uma vez só.`n`n"
            "Para cada curso, o sistema irá:`n"
            "  1. Localizar a pasta do curso com os certificados em PDF;`n"
            "  2. Selecionar o período/meses desejados;`n"
            "  3. Ler a planilha do curso e identificar automaticamente as colunas de Docente e E-mail;`n"
            "  4. Abrir os e-mails prontos no Outlook com controle automático de progresso.`n`n"
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
    ExibirMensagem(
        "Passo 1 — Selecionar Pasta do Curso",
        "Pasta dos Certificados Gerados",
        "Selecione a pasta da Pós-Graduação dentro de " PASTA_PADRAO_CERTS " que contém as subpastas dos meses com os certificados já gerados.`n`n"
        "[Curso " numCursoAtual "]",
        "passo",
        "Selecionar Pasta ➔",
        600
    )

    pastaPosgrad := DirSelect(PASTA_PADRAO_CERTS, 3, "Selecione a pasta da pós-graduação [Curso " numCursoAtual "]")
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
            excelPath := FileSelect(1, pastaInicio, "Selecione a planilha (.xlsx) com docentes e e-mails [Curso " numCursoAtual "]", "*.xlsx")
            if !excelPath {
                ExibirMensagem("Operação Cancelada", "Nenhuma Planilha Selecionada", "Nenhuma planilha foi selecionada. Este curso não será adicionado.", "erro", "Continuar")
                continue
            }
        }
    } else {
        ExibirMensagem(
            "Passo 3 — Planilha do Curso",
            "Planilha de Docentes e E-mails",
            "Não foi possível localizar automaticamente a planilha na nuvem.`n`nSelecione a planilha Excel (.xlsx) contendo nome e e-mail dos docentes deste curso.`n`n"
            "[Curso " numCursoAtual "]",
            "passo",
            "Selecionar Planilha ➔",
            600
        )
        pastaInicioExcel := INFO_NUVEM["raizOneDrive"] != "" ? INFO_NUVEM["raizOneDrive"] : ""
        excelPath := FileSelect(1, pastaInicioExcel, "Selecione a planilha (.xlsx) com docentes e e-mails [Curso " numCursoAtual "]", "*.xlsx")
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

    confirmPos := Confirmar(
        "Passo 5 — Confirmar Curso",
        "Nome da Pós-Graduação Identificado",
        "[Curso " numCursoAtual "] O sistema identificou o seguinte nome de curso na célula A2:`n`n"
        "  🎓 " nomePosGraduacao "`n`n"
        "Esse nome será utilizado no assunto e no texto padrão dos e-mails.`n`n"
        "Deseja confirmar?",
        "✅ Confirmar e Avançar",
        "❌ Cancelar Curso",
        true,
        620
    )

    if !confirmPos {
        try wbTemp.Close(false)
        try excelTemp.Quit()
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

        ; Detecta colunas específicas por aba caso variem
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

modoUmAUm := EscolherModoEnvioUI(620)

resumoRodada := ""
totalDocentesGeral := 0

for idx, curso in cursosRodada {
    resumoRodada .= idx ". " curso["nomePosGraduacao"] "`n"
    resumoRodada .= "   Meses: " JoinArr(curso["mesesSelecionados"], ", ") "`n"
    resumoRodada .= "   Docentes mapeados: " curso["ordemDocentes"].Length "`n`n"
    totalDocentesGeral += curso["ordemDocentes"].Length
}

confirmFinal := Confirmar(
    "Passo 9 — Confirmar Rodada de Envio",
    "Resumo da Rodada de Envio",
    "Cursos configurados na rodada: " cursosRodada.Length "`n"
    "Total de docentes mapeados: " totalDocentesGeral "`n"
    "Modo de envio: " (modoUmAUm ? "👤 Um a um" : "⚡ Todos de uma vez") "`n`n"
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
;   FASE 4 — EXECUÇÃO DO ENVIO (CURSO POR CURSO)
; ════════════════════════════════════════════════════════════════════════

assuntoPadrao := "Envio de certificados Pós Graduação"

totalGeralAbertos    := 0
totalGeralPulados    := 0
totalGeralJaEnviados := 0
relatorioFinalCursos := ""

for idxCurso, curso in cursosRodada {

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

    resultadoTemp := ConstruirPastaTempCurso(pastaPosgrad, mesesSelecionados, enviadosMap)
    PASTA_CERTS_ENVIO := resultadoTemp["pasta"]

    corpoEmail := "Boa tarde!%0D%0A%0D%0ASegue certificado(s) referente às aulas da pós-graduação em " nomePosGraduacao "."

    totalAbertosCurso := 0
    totalPuladosCurso := 0
    interromperRodada := false

    if modoUmAUm {
        for idx, chave in ordemPendente {
            dados := mapaDocentes[chave]
            nomeDocente := dados["nome"]
            emailReal := dados["email"]

            if !dados["emailValido"] {
                totalPuladosCurso++
                continue
            }

            ; Abre o Outlook Web com destinatário, assunto e corpo
            link := "https://outlook.office.com/mail/deeplink/compose?to=" emailReal
                 . "&subject=" assuntoPadrao
                 . "&body=" corpoEmail
            Run(link)
            totalAbertosCurso++
            Sleep(2200)

            pastaDocente := EncontrarPastaDocente(PASTA_CERTS_ENVIO, nomeDocente)
            if (pastaDocente != "") {
                Run('explorer.exe "' pastaDocente '"')
                Sleep(800)

                acaoEnvio := ExibirGuiaEnvioDocenteUI(
                    idxCurso, cursosRodada.Length, nomePosGraduacao,
                    idx, ordemPendente.Length, nomeDocente, emailReal,
                    pastaDocente, 640
                )

                if (acaoEnvio = "enviado") {
                    MarcarPendentesComoEnviados(pastaControle, nomeDocente, chave, mesesSelecionados, enviadosMap)
                } else if (acaoEnvio = "pular") {
                    totalPuladosCurso++
                } else if (acaoEnvio = "parar") {
                    interromperRodada := true
                    break
                }
            } else {
                ExibirMensagem(
                    "Certificado Não Localizado",
                    "Pasta do Docente Não Encontrada",
                    "Não foram localizados arquivos PDF pendentes para:`n`n"
                    "  " nomeDocente " <" emailReal ">`n`n"
                    "Este docente será pulado.",
                    "erro",
                    "Continuar ➔"
                )
                totalPuladosCurso++
            }
        }
    } else {
        ; Modo Todos de uma vez: abre todas as abas primeiro
        for chave in ordemPendente {
            dados := mapaDocentes[chave]
            emailReal := dados["email"]
            if !dados["emailValido"] {
                totalPuladosCurso++
                continue
            }
            link := "https://outlook.office.com/mail/deeplink/compose?to=" emailReal
                 . "&subject=" assuntoPadrao
                 . "&body=" corpoEmail
            Run(link)
            totalAbertosCurso++
            Sleep(600)
        }

        ExibirMensagem(
            "Abas Abertas no Navegador",
            "Outlook Pronto para Anexos",
            totalAbertosCurso " aba(s) do Outlook foram abertas para o curso:`n`n"
            "  🎓 " nomePosGraduacao "`n`n"
            "Clique abaixo para iniciar a esteira de confirmação dos anexos docente por docente.",
            "sucesso",
            "Iniciar Anexos ➔"
        )

        for idx, chave in ordemPendente {
            dados := mapaDocentes[chave]
            nomeDocente := dados["nome"]
            emailReal := dados["email"]
            if !dados["emailValido"]
                continue

            pastaDocente := EncontrarPastaDocente(PASTA_CERTS_ENVIO, nomeDocente)
            if (pastaDocente != "") {
                Run('explorer.exe "' pastaDocente '"')
                Sleep(800)

                acaoEnvio := ExibirGuiaEnvioDocenteUI(
                    idxCurso, cursosRodada.Length, nomePosGraduacao,
                    idx, ordemPendente.Length, nomeDocente, emailReal,
                    pastaDocente, 640
                )

                if (acaoEnvio = "enviado") {
                    MarcarPendentesComoEnviados(pastaControle, nomeDocente, chave, mesesSelecionados, enviadosMap)
                } else if (acaoEnvio = "pular") {
                    totalPuladosCurso++
                } else if (acaoEnvio = "parar") {
                    interromperRodada := true
                    break
                }
            }
        }
    }

    ; Pergunta sobre limpeza da pasta temporária
    respLimpeza := Confirmar(
        "Limpeza de Pasta Temporária",
        "Curso Concluído: " nomePosGraduacao,
        "Deseja remover agora a pasta temporária de envio deste curso?`n`n"
        PASTA_CERTS_ENVIO "`n`n"
        "(Os certificados originais organizados por mês e o arquivo de controle NÃO são afetados).",
        "🗑️ Sim, Limpar",
        "Manter Pasta",
        true,
        600
    )
    if respLimpeza {
        try DirDelete(PASTA_CERTS_ENVIO, true)
    }

    totalGeralAbertos += totalAbertosCurso
    totalGeralPulados += totalPuladosCurso

    relatorioFinalCursos .= idxCurso ". " nomePosGraduacao "`n"
    relatorioFinalCursos .= "     Enviados/Abertos: " totalAbertosCurso " | Pulados: " totalPuladosCurso " | Já enviados antes: " totalJaEnviadoCurso "`n`n"

    if interromperRodada {
        ExibirMensagem("Rodada Interrompida", "Envios Pausados pelo Usuário", "A rodada foi interrompida conforme solicitado. O progresso realizado foi devidamente salvo no controle de envio.", "info", "Ver Relatório")
        break
    }
}

; ════════════════════════════════════════════════════════════════════════
;   FASE 5 — RELATÓRIO FINAL DA RODADA (UI_MODERNA)
; ════════════════════════════════════════════════════════════════════════

resumoFinalEnvio := "Cursos processados:              " cursosRodada.Length "`n"
resumoFinalEnvio .= "E-mails abertos/enviados:        " totalGeralAbertos "`n"
resumoFinalEnvio .= "Pulados nesta rodada:            " totalGeralPulados "`n"
resumoFinalEnvio .= "Já enviados em rodadas passadas: " totalGeralJaEnviados "`n"
resumoFinalEnvio .= "Modo de envio utilizado:         " (modoUmAUm ? "Um a um" : "Todos de uma vez")

ExibirRelatorioFinalUI(
    "Passo 10 — Relatório Final da Rodada",
    "Rodada de Envio Concluída!",
    resumoFinalEnvio,
    Trim(relatorioFinalCursos, "`n"),
    PASTA_PADRAO_CERTS,
    700
)

ExitApp
