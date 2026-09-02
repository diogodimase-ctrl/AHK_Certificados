#Requires AutoHotkey v2.0
#SingleInstance Force
#Include UI_Moderna.ahk

; ════════════════════════════════════════════════════════════════════════
;   ENVIAR EMAILS — v9 — RODADAS DE ENVIO (Parte I / Parte II / ...)
;   Interface visual padronizada com tema moderno (Segoe UI / #F7F9FC)
; ════════════════════════════════════════════════════════════════════════
;
;   Base: v7. NOVIDADES DESTA VERSÃO (v9):
;   - Passo 7 (Corrigir E-mail): a mensagem agora é específica — mostra
;     o docente, a pós-graduação, o mês e a aba da planilha, e a LINHA
;     exata onde o e-mail foi lido, para você achar rápido o que corrigir
;     na planilha (antes só mostrava nome do docente e curso).
;   - Revisão geral dos textos de todas as telas (Passo 1 a Passo 10):
;     explicações mais completas e diretas, pensadas para quem nunca usou
;     o programa — cada etapa agora deixa claro o que é esperado, dá
;     exemplos e explica os detalhes que podem gerar dúvida (ex: como
;     achar a letra da coluna no Excel, o que fazer exatamente na hora
;     de anexar e enviar o e-mail).
;
;   Herdado da v7 — mesma filosofia da polida feita em GerarCertificados v16:
;   - Passo 2 (Período/Meses): menu clicável com caixinhas e botão
;     "Marcar Todos", em vez de digitar os meses separados por vírgula.
;   - Passo 4 (Abas não bateram com os meses): lista suspensa por mês,
;     já pré-selecionada com o palpite mais provável (reconhece "04.26"
;     como "Abril", por exemplo) — sem digitar nada.
;   - Passo 8 (Modo de Envio): botões com o texto real ("Um a um" /
;     "Todos de uma vez"), em vez de Sim/Não.
;   - Todas as telas numeradas de Passo 1 a Passo 10.
;   - Fechamento do Excel (wb.Close / excel.Quit) protegido com try em
;     todos os pontos de saída, para não travar o script se o Excel
;     estiver ocupado no momento do fechamento.
;
;   O que já existia desde v5 continua igual:
;   Assim como a geração de certificados permite selecionar vários cursos
;   numa "rodada de geração", este script permite montar uma
;   "RODADA DE ENVIO" com vários cursos de uma vez. Dentro da rodada, os
;   cursos são processados um por um, e dentro de cada curso, os docentes
;   também são processados um por um (ou todos de uma vez, à sua escolha).
;
;   CONTROLE AUTOMÁTICO DE PROGRESSO:
;   Cada pasta de pós-graduação recebe um arquivo de controle
;   "_EnvioControle.txt". Sempre que você confirma o envio de um docente,
;   o script marca ali quais meses já foram enviados para aquele docente.
;   Assim, se você fizer a "Parte I" hoje (ex: 15 dos 30 docentes) e
;   fechar o script, ao rodar de novo depois para a "Parte II", quem já
;   foi enviado é detectado e pulado automaticamente — sem duplicar envio.
;
; ════════════════════════════════════════════════════════════════════════

CONTROLE_ARQUIVO := "_EnvioControle.txt"

; ────────────────────────────────────────────────────────────────────────
;   FUNÇÕES AUXILIARES
; ────────────────────────────────────────────────────────────────────────

ListarSubpastas(pasta) {
    lista := []
    Loop Files, pasta "*", "D" {
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

; Extrai um e-mail válido de dentro de uma célula que pode ter outras
; informações junto (telefone, observações, etc.), separadas por ; , ou espaço.
; Ex: "11 99999-9999; fulano@dominio.com" -> "fulano@dominio.com"
ExtrairEmailValido(bruto) {
    bruto := Trim(bruto)
    if RegExMatch(bruto, "[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", &m)
        return m[0]
    return bruto
}

; Remove espaços duplicados/extremos e normaliza para minúsculas, para
; permitir comparar nomes de pasta com pequenas diferenças de formatação.
NormalizarNome(s) {
    s := StrLower(Trim(s))
    s := RegExReplace(s, "\s+", " ")
    return s
}

; =========================================================
; MESES: código numérico -> nome por extenso (sem acento), usado
; para reconhecer as pastas no formato "MM.AA" (ex.: "04.26") e
; casá-las com abas de planilha chamadas por nome (ex.: "Abril").
; =========================================================

MESES_NUMERO_NOME := Map(
    "01", "janeiro", "02", "fevereiro", "03", "marco", "04", "abril",
    "05", "maio", "06", "junho", "07", "julho", "08", "agosto",
    "09", "setembro", "10", "outubro", "11", "novembro", "12", "dezembro"
)

; Remove acentos comuns do português para comparação tolerante.
RemoverAcentos(s) {
    s := StrLower(s)
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

; Normaliza para comparação: sem acento, minúsculo, só letras/números.
NormalizarParaComparar(s) {
    s := RemoverAcentos(Trim(s))
    return RegExReplace(s, "[^a-z0-9]", "")
}

; Se o nome da pasta estiver no formato "MM.AA" (ex.: "04.26"),
; devolve Map com o número do mês, o nome por extenso e o ano.
; Caso contrário, devolve "".
ExtrairMesAnoDaPasta(nomePasta) {
    if RegExMatch(Trim(nomePasta), "^(\d{1,2})\.(\d{2,4})$", &m) {
        numMes := Format("{:02}", Integer(m[1]))
        if MESES_NUMERO_NOME.Has(numMes)
            return Map("numMes", numMes, "nomeMes", MESES_NUMERO_NOME[numMes], "ano", m[2])
    }
    return ""
}

; Verifica se o nome de uma aba da planilha corresponde ao mês de
; uma pasta no formato "MM.AA" — reconhecendo tanto o nome por
; extenso ("Abril", "Abril/26", "Abril 2026") quanto o número puro
; ("04", "4"), e conferindo o ano quando a aba também tiver um.
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

    ; 1) casar pelo NOME do mês (ex.: "abril", "abril26", "abril2026")
    if InStr(abaNorm, nomeMesNorm) {
        resto := Trim(StrReplace(abaNorm, nomeMesNorm, "", , , 1))
        if (resto = "" || resto = anoCurto || resto = anoLongo)
            return true
        if !RegExMatch(resto, "^\d+$")
            return true    ; sobrou texto (turma, unidade etc.) — tolera
        return false        ; sobrou um número que não bate com o ano da pasta
    }

    ; 2) casar pelo NÚMERO do mês, quando a aba é só um número (ex.: "04", "4")
    if RegExMatch(abaNorm, "^\d+$")
        return (Integer(abaNorm) = Integer(numMesNorm))

    return false
}

; Procura, dentro de pastaBase, a subpasta que corresponde ao docente,
; mesmo que o nome não bata 100% caractere a caractere (acentuação,
; espaços extras, sufixo de gênero etc.). Devolve o caminho completo
; com "\" no final, ou "" se não encontrar nada parecido.
EncontrarPastaDocente(pastaBase, nomeAlvo) {
    alvoNorm := NormalizarNome(nomeAlvo)
    melhorMatch := ""
    Loop Files, pastaBase "*", "D" {
        nomeNorm := NormalizarNome(A_LoopFileName)
        if (nomeNorm = alvoNorm)
            return A_LoopFilePath "\"
        if (melhorMatch = "") && (InStr(nomeNorm, alvoNorm) || InStr(alvoNorm, nomeNorm))
            melhorMatch := A_LoopFilePath "\"
    }
    return melhorMatch
}

JoinArr(arr, sep) {
    resultado := ""
    for idx, item in arr {
        resultado .= (idx = 1 ? "" : sep) item
    }
    return resultado
}

; =========================================================
; FUNÇÃO: Selecionar meses por menu clicável (ListView com
; caixinhas), com botão "Marcar Todos" — mesmo padrão usado para
; selecionar cursos na geração de certificados. Retorna array com
; os meses marcados (vazio se cancelado ou nada marcado).
; =========================================================

SelecionarMesesClicavel(tituloJanela, cabecalho, meses) {
    resultado  := []
    confirmado := false

    g := Gui("+AlwaysOnTop -MaximizeBox", tituloJanela)
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F7F9FC"

    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w540 c003366", "📅 " cabecalho)
    g.Add("Text", "x20 y48 w540 h2 0x10")

    g.SetFont("s10 normal", "Segoe UI")
    lv := g.Add("ListView", "x20 y58 w540 r12 Checked", ["Mês / Pasta"])
    for m in meses
        lv.Add(, m)
    lv.ModifyCol(1, 510)

    MarcarTodos(*) {
        Loop lv.GetCount()
            lv.Modify(A_Index, "Check")
    }

    DesmarcarTodos(*) {
        Loop lv.GetCount()
            lv.Modify(A_Index, "-Check")
    }

    ConfirmarClick(*) {
        marcados := []
        linha := 0
        while (linha := lv.GetNext(linha, "C"))
            marcados.Push(meses[linha])
        resultado := marcados
        confirmado := true
        g.Destroy()
    }

    g.SetFont("s9", "Segoe UI")
    btnTodos := g.Add("Button", "x20 y+12 w120 h32", "Marcar Todos")
    btnNenhum := g.Add("Button", "x+8 w120 h32", "Desmarcar Todos")
    btnTodos.OnEvent("Click", MarcarTodos)
    btnNenhum.OnEvent("Click", DesmarcarTodos)

    g.SetFont("s10 bold", "Segoe UI")
    btnOk := g.Add("Button", "Default w140 h36 x290 yp-2", "Confirmar ➔")
    btnOk.OnEvent("Click", ConfirmarClick)
    btnCancelar := g.Add("Button", "w120 h36 x+10", "Cancelar")
    btnCancelar.OnEvent("Click", (*) => g.Destroy())

    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show("w580")
    WinWaitClose("ahk_id " g.Hwnd)

    return confirmado ? resultado : []
}

; =========================================================
; FUNÇÃO: Corrigir a correspondência mês → aba clicando numa
; lista suspensa por mês
; =========================================================

CorrigirAbasClicavel(tituloJanela, meses, abasDisponiveis) {
    resultado  := ""
    confirmado := false
    combos     := []

    g := Gui("+AlwaysOnTop -MaximizeBox", tituloJanela)
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F7F9FC"

    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w540 c003366", "🔍 Passo 4 — Correspondência de Abas")
    g.Add("Text", "x20 y48 w540 h2 0x10")

    g.SetFont("s10 normal", "Segoe UI")
    g.Add("Text", "x20 y58 w540 c333333", "Os nomes das pastas não bateram exatamente com as abas da planilha.`nPara cada mês, confira ou selecione a aba correspondente:")

    for m in meses {
        g.SetFont("s10 bold", "Segoe UI")
        g.Add("Text", "x20 y+12 w540 c003366", "Mês da pasta: " m)

        ; Tenta um "chute" automático
        indicePalpite := 1
        for i, aba in abasDisponiveis {
            if AbaCorrespondeAoMes(aba, m) {
                indicePalpite := i
                break
            }
        }

        g.SetFont("s10 normal", "Segoe UI")
        cb := g.Add("DDL", "x20 y+4 w540 Choose" indicePalpite, abasDisponiveis)
        combos.Push(cb)
    }

    ConfirmarClick(*) {
        escolhidas := []
        for cb in combos
            escolhidas.Push(cb.Text)
        resultado := escolhidas
        confirmado := true
        g.Destroy()
    }

    g.SetFont("s10 bold", "Segoe UI")
    btnOk := g.Add("Button", "Default w140 h36 x290 y+20", "Confirmar ➔")
    btnOk.OnEvent("Click", ConfirmarClick)
    btnCancelar := g.Add("Button", "w120 h36 x+10", "Cancelar")
    btnCancelar.OnEvent("Click", (*) => g.Destroy())

    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show("w580")
    WinWaitClose("ahk_id " g.Hwnd)

    return confirmado ? resultado : ""
}

; =========================================================
; FUNÇÃO: Escolher o modo de envio
; =========================================================

EscolherModoEnvio() {
    resultado := true

    g := Gui("+AlwaysOnTop -MaximizeBox", "Passo 8 — Modo de Envio")
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F7F9FC"

    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w540 c003366", "⚙️ Passo 8 — Escolha o Modo de Envio")
    g.Add("Text", "x20 y48 w540 h2 0x10")

    g.SetFont("s10 normal", "Segoe UI")
    g.Add("Text", "x20 y58 w540 c333333",
        "Como deseja processar o envio dos e-mails desta rodada?`n`n"
        "• Um a um — Abre um e-mail por vez no Outlook. Você anexa os PDFs, envia e confirma para abrir o próximo.`n`n"
        "• Todos de uma vez — Abre todas as abas do Outlook de uma vez só para este curso e depois confirma docente por docente."
    )

    g.SetFont("s10 bold", "Segoe UI")
    btnUmAUm := g.Add("Button", "Default w200 h40 x120 y+20", "📩 Um a um (Recomendado)")
    btnTodos := g.Add("Button", "w200 h40 x+15", "⚡ Todos de uma vez")

    btnUmAUm.OnEvent("Click", (*) => (resultado := true, g.Destroy()))
    btnTodos.OnEvent("Click", (*) => (resultado := false, g.Destroy()))
    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show("w580")
    WinWaitClose("ahk_id " g.Hwnd)

    return resultado
}

; Lê o arquivo de controle e devolve um Map: chave(docente lower) -> Map(mes -> true)
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

; Grava uma linha no arquivo de controle marcando docente+mes como enviado
MarcarEnviado(pastaControle, nomeDocente, mes) {
    linha := nomeDocente "`t" mes "`t" FormatTime(A_Now, "yyyy-MM-dd HH:mm") "`n"
    try FileAppend(linha, pastaControle, "UTF-8")
}

; Verifica se um docente já está 100% enviado para TODOS os meses selecionados
EstaTotalmenteEnviado(enviadosMap, chave, mesesSelecionados) {
    if !enviadosMap.Has(chave)
        return false
    for mes in mesesSelecionados {
        if !enviadosMap[chave].Has(mes)
            return false
    }
    return true
}

; Devolve a lista de meses ainda NÃO enviados para um docente
MesesPendentes(enviadosMap, chave, mesesSelecionados) {
    pendentes := []
    for mes in mesesSelecionados {
        jaEnviado := enviadosMap.Has(chave) && enviadosMap[chave].Has(mes)
        if !jaEnviado
            pendentes.Push(mes)
    }
    return pendentes
}

; Marca como enviados todos os meses que ainda estavam pendentes para o docente
MarcarPendentesComoEnviados(pastaControle, nomeDocente, chave, mesesSelecionados, enviadosMap) {
    pendentes := MesesPendentes(enviadosMap, chave, mesesSelecionados)
    for mes in pendentes {
        MarcarEnviado(pastaControle, nomeDocente, mes)
        if !enviadosMap.Has(chave)
            enviadosMap[chave] := Map()
        enviadosMap[chave][mes] := true
    }
}

; Monta a pasta temporária de um curso, pulando arquivos de docente+mês
; que já constam como enviados no arquivo de controle
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
            nomeDocente := Trim(nomeDocente)
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
;   FASE 1 — CONFIGURAÇÃO DA RODADA (adicionar um ou mais cursos)
; ════════════════════════════════════════════════════════════════════════

cursosRodada := []
ultimoColNome := ""
ultimoColEmail := ""
primeiraVez := true

Loop {
    if primeiraVez {
        ExibirMensagem(
            "Passo 1 — Como funciona este programa",
            "Envio de Certificados por E-mail aos Docentes",
            "Este programa envia por e-mail, para cada docente, os certificados em PDF "
            "que já foram gerados anteriormente.`n`n"
            "Você vai montar uma RODADA DE ENVIO: pode incluir um ou vários cursos de "
            "uma vez só. Para cada curso, o programa vai pedir:`n"
            "  1) a pasta da pós-graduação com os certificados já gerados;`n"
            "  2) os meses que você quer enviar;`n"
            "  3) a planilha Excel com nome e e-mail de cada docente.`n`n"
            "No final, o envio acontece curso por curso, e dentro de cada curso, "
            "docente por docente — o programa abre o e-mail já com o destinatário e o "
            "anexo prontos, e é só você conferir e clicar em enviar no Outlook.`n`n"
            "IMPORTANTE: quem já recebeu o e-mail é identificado automaticamente e pulado "
            "pelo arquivo _EnvioControle.txt — sem risco de envio duplicado.",
            "passo",
            "Começar Configuração ➔",
            600
        )
        primeiraVez := false
    } else {
        respAdd := Confirmar(
            "Adicionar Outro Curso?",
            "Curso Adicionado com Sucesso!",
            "Cursos já incluídos nesta rodada: " cursosRodada.Length "`n`n"
            "Deseja adicionar mais um curso a esta mesma rodada de envio?",
            "➕ Adicionar Outro Curso",
            "🚀 Seguir para o Envio",
            false,
            560
        )
        if !respAdd
            break
    }

    numCursoAtual := cursosRodada.Length + 1

    ; --- seleção da pasta do curso ---
    pastaPosgrad := DirSelect("C:\Certificados\", 3, "Passo 1 - selecione a PASTA DA PÓS-GRADUAÇÃO (a pasta que contém as subpastas de cada mês com os certificados já gerados) [Curso " numCursoAtual "]")
    if !pastaPosgrad {
        if (cursosRodada.Length = 0) {
            ExibirMensagem("Programa Encerrado", "Nenhuma Pasta Selecionada", "Nenhuma pasta de certificados foi selecionada. O programa será encerrado.", "erro", "Fechar")
            ExitApp
        } else {
            break
        }
    }
    pastaPosgrad := RTrim(pastaPosgrad, "\") "\"

    mesesDisponiveis := ListarSubpastas(pastaPosgrad)
    if (mesesDisponiveis.Length = 0) {
        ExibirMensagem("Pasta Sem Subpastas", "Nenhum Mês Encontrado", "Nenhuma subpasta de mês encontrada dentro de:`n" pastaPosgrad "`n`nEste curso não será adicionado.", "erro", "Continuar")
        continue
    }

    mesesSelecionados := SelecionarMesesClicavel(
        "Passo 2 — Selecionar Período (Meses)",
        "[Curso " numCursoAtual "] Selecione os meses a enviar:",
        mesesDisponiveis
    )
    if (mesesSelecionados.Length = 0) {
        ExibirMensagem("Seleção Cancelada", "Nenhum Mês Selecionado", "Nenhum mês foi selecionado. Este curso não será adicionado.", "erro", "Continuar")
        continue
    }

    ; --- planilha do curso ---
    ExibirMensagem(
        "Passo 3 — Planilha do Curso",
        "Planilha de Docentes — Curso " numCursoAtual,
        "Selecione a planilha Excel (.xlsx) que contém os NOMES e E-MAILS dos docentes deste curso.`n`n"
        "Os nomes precisam bater com os nomes dos certificados em PDF para anexar corretamente.",
        "passo",
        "Selecionar Planilha ➔"
    )
    excelPath := FileSelect(1, "", "Passo 3 - selecione a planilha (.xlsx) com nome e e-mail dos docentes [Curso " numCursoAtual "]", "*.xlsx")
    if !excelPath {
        ExibirMensagem("Seleção Cancelada", "Planilha Não Selecionada", "Nenhuma planilha selecionada. Este curso não será adicionado.", "erro", "Continuar")
        continue
    }

    excelTemp := ""
    wbTemp := ""
    try {
        excelTemp := ComObject("Excel.Application")
        excelTemp.Visible := false
        wbTemp := excelTemp.Workbooks.Open(excelPath)
    } catch as err {
        ExibirMensagem("Erro no Excel", "Falha ao Abrir Planilha", "Erro ao abrir a planilha:`n`n" err.Message "`n`nEste curso não será adicionado.", "erro", "Continuar")
        if IsObject(excelTemp)
            try excelTemp.Quit()
        continue
    }

    abasWb := []
    Loop wbTemp.Sheets.Count
        abasWb.Push(Trim(wbTemp.Sheets.Item(A_Index).Name))

    ; --- tentar casar automaticamente os meses da pasta com as abas da planilha ---
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
        abasCorrespondentes := CorrigirAbasClicavel(
            "Passo 4 — Corrigir Abas [Curso " numCursoAtual "]",
            mesesSelecionados,
            abasWb
        )
        if (abasCorrespondentes = "") {
            try wbTemp.Close(false)
            try excelTemp.Quit()
            ExibirMensagem("Seleção Cancelada", "Nenhuma Aba Selecionada", "Nenhuma aba selecionada. Este curso não será adicionado.", "erro", "Continuar")
            continue
        }
    }

    ; --- nome da pós-graduação (célula A2 da primeira aba correspondente) ---
    wsPrimeira := ""
    Loop wbTemp.Sheets.Count {
        if (Trim(wbTemp.Sheets.Item(A_Index).Name) = abasCorrespondentes[1]) {
            wsPrimeira := wbTemp.Sheets.Item(A_Index)
            break
        }
    }
    nomePosGraduacao := Trim(wsPrimeira.Range("A2").Value)
    if (nomePosGraduacao = "") {
        try wbTemp.Close(false)
        try excelTemp.Quit()
        ExibirMensagem("Planilha Incompleta", "Célula A2 Vazia", "A célula A2 está vazia nesta planilha. Este curso não será adicionado.", "erro", "Continuar")
        continue
    }

    confirmPos := Confirmar(
        "Passo 5 — Confirmar Pós-Graduação",
        "Nome da Pós-Graduação [Curso " numCursoAtual "]",
        "O programa identificou na célula A2 da planilha:`n`n`"" nomePosGraduacao "`"`n`n"
        "Esse é o nome que vai aparecer no corpo do e-mail enviado aos docentes. Está correto?",
        "✅ Está Correto",
        "❌ Cancelar Curso",
        true,
        600
    )
    if !confirmPos {
        try wbTemp.Close(false)
        try excelTemp.Quit()
        continue
    }

    ; --- colunas de nome e e-mail ---
    colNomeBox := PedirTexto(
        "Passo 6 — Coluna dos Nomes",
        "Coluna dos Docentes [Curso " numCursoAtual "]",
        "Abra a planilha e informe em qual COLUNA (letra) estão os NOMES dos docentes.`n`nExemplo: E",
        ultimoColNome,
        "Digite apenas a letra da coluna (ex: E)"
    )
    if (colNomeBox["Result"] != "OK" || Trim(colNomeBox["Value"]) = "") {
        try wbTemp.Close(false)
        try excelTemp.Quit()
        ExibirMensagem("Cancelado", "Coluna Não Informada", "Coluna de nomes não informada. Este curso não será adicionado.", "erro", "Continuar")
        continue
    }
    colEmailBox := PedirTexto(
        "Passo 6 — Coluna dos E-mails",
        "Coluna dos E-mails [Curso " numCursoAtual "]",
        "Agora informe em qual COLUNA estão os E-MAILS dos docentes:`n`nExemplo: F",
        ultimoColEmail,
        "Digite apenas a letra da coluna (ex: F)"
    )
    if (colEmailBox["Result"] != "OK" || Trim(colEmailBox["Value"]) = "") {
        try wbTemp.Close(false)
        try excelTemp.Quit()
        ExibirMensagem("Cancelado", "Coluna Não Informada", "Coluna de e-mails não informada. Este curso não será adicionado.", "erro", "Continuar")
        continue
    }
    colNomeLetra := Trim(colNomeBox["Value"])
    colEmailLetra := Trim(colEmailBox["Value"])
    ultimoColNome := colNomeLetra
    ultimoColEmail := colEmailLetra

    ; --- ler docentes/e-mails deste curso, em todas as abas correspondentes ---
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

        ultimaLinha := wsAba.Cells(wsAba.Rows.Count, colNomeLetra).End(-4162).Row

        Loop ultimaLinha - 4 {
            linha := A_Index + 4
            nomeBruto := Trim(wsAba.Cells(linha, colNomeLetra).Value)
            emailBruto := Trim(wsAba.Cells(linha, colEmailLetra).Value)
            if (nomeBruto = "")
                continue

            nomeLimpo := RegExReplace(nomeBruto, "\s*\b[HMhm]\b\s*$")
            nomeLimpo := Trim(nomeLimpo)
            nomeLimpo := RegExReplace(nomeLimpo, '[\\/:*?"<>|]')
            chave := StrLower(nomeLimpo)

            ; extrai só o e-mail de fato
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
        ExibirMensagem("Planilha Sem Docentes", "Nenhum Docente Encontrado", "Nenhum docente encontrado nas abas selecionadas. Este curso não será adicionado.", "erro", "Continuar")
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
    ExibirMensagem("Processamento Encerrado", "Nenhum Curso Válido", "Nenhum curso foi adicionado. Programa encerrado.", "erro", "Fechar")
    ExitApp
}

; ────────────────────────────────────────────────────────────────────────
;   Correção de e-mails inválidos/ausentes, em todos os cursos da rodada
; ────────────────────────────────────────────────────────────────────────

for curso in cursosRodada {
    for chave in curso["ordemDocentes"] {
        dados := curso["mapaDocentes"][chave]
        if !dados["emailValido"] {
            valorLido := (dados["email"] = "") ? "(célula vazia — nenhum e-mail encontrado)" : "`"" dados["email"] "`""

            correcao := PedirTexto(
                "Passo 7 — Corrigir E-mail — " dados["nome"],
                "Docente Sem E-mail Válido",
                "Docente:        " dados["nome"] "`n"
                "Pós-graduação:  " curso["nomePosGraduacao"] "`n"
                "Mês / Aba:      " dados["origemMes"] " (aba `"" dados["origemAba"] "`")`n"
                "Linha:          " dados["origemLinha"] "`n"
                "E-mail lido:    " valorLido "`n`n"
                "Digite o e-mail correto abaixo (ou deixe em branco para pular este docente nesta rodada):",
                dados["email"],
                "Exemplo: docente@dominio.com.br",
                600
            )
            if (correcao["Result"] = "OK" && EmailEhValido(Trim(correcao["Value"]))) {
                dados["email"] := Trim(correcao["Value"])
                dados["emailValido"] := true
            }
        }
    }
}

; ────────────────────────────────────────────────────────────────────────
;   Modo de envio + confirmação final da rodada
; ────────────────────────────────────────────────────────────────────────

resumoRodada := ""
for idx, curso in cursosRodada {
    resumoRodada .= idx ". " curso["nomePosGraduacao"] "`n"
    resumoRodada .= "   Meses: " JoinArr(curso["mesesSelecionados"], ", ") "`n"
    resumoRodada .= "   Docentes na planilha: " curso["ordemDocentes"].Length "`n`n"
}

modoUmAUm := EscolherModoEnvio()

confirmFinal := Confirmar(
    "Passo 9 — Confirmar Rodada de Envio",
    "Resumo da Rodada de Envio (" cursosRodada.Length " cursos)",
    resumoRodada
    "Modo de envio escolhido: " (modoUmAUm ? "Um a um" : "Todos de uma vez") "`n`n"
    "O sistema verifica quem já foi enviado anteriormente e pula automaticamente.`n`n"
    "Deseja iniciar o envio agora?",
    "🚀 Iniciar Envio",
    "❌ Cancelar",
    true,
    620
)
if !confirmFinal
    ExitApp

; ════════════════════════════════════════════════════════════════════════
;   FASE 2 — ENVIO, CURSO POR CURSO
; ════════════════════════════════════════════════════════════════════════

assunto := "Envio de certificados Pós Graduação"

totalGeralAbertos := 0
totalGeralPulados := 0
totalGeralJaEnviados := 0

for idxCurso, curso in cursosRodada {

    pastaPosgrad := curso["pastaPosgrad"]
    mesesSelecionados := curso["mesesSelecionados"]
    nomePosGraduacao := curso["nomePosGraduacao"]
    mapaDocentes := curso["mapaDocentes"]
    ordemDocentes := curso["ordemDocentes"]

    pastaControle := pastaPosgrad CONTROLE_ARQUIVO
    enviadosMap := CarregarEnviados(pastaControle)

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

    ExibirMensagem(
        "Passo 9 — Enviando",
        "Iniciando Curso " idxCurso " de " cursosRodada.Length,
        nomePosGraduacao "`n`n"
        "Docentes pendentes de envio: " ordemPendente.Length "`n"
        "Já enviados em rodadas anteriores (serão pulados): " totalJaEnviadoCurso "`n`n"
        "Clique abaixo para montar a pasta de anexos e começar o envio.",
        "passo",
        "Montar Anexos e Iniciar ➔",
        600
    )

    if (ordemPendente.Length = 0) {
        ExibirMensagem("Curso Concluído", "Todos Enviados", "Todos os docentes deste curso já foram enviados anteriormente. Avançando para o próximo.", "sucesso", "Próximo Curso ➔")
        continue
    }

    resultadoTemp := ConstruirPastaTempCurso(pastaPosgrad, mesesSelecionados, enviadosMap)
    PASTA_CERTS_ENVIO := resultadoTemp["pasta"]

    ExibirMensagem(
        "Organização Concluída",
        "Pasta Temporária Pronta — " nomePosGraduacao,
        "Certificados novos copiados: " resultadoTemp["totalCopiados"] "`n"
        "Já enviados antes (não duplicados): " resultadoTemp["totalPulados"] "`n`n"
        "Local: " PASTA_CERTS_ENVIO,
        "sucesso",
        "Avançar para Envio ➔",
        600
    )

    corpo := "Boa tarde!%0D%0A%0D%0ASegue certificado(s) referente às aulas da pós-graduação em " nomePosGraduacao "."

    totalAbertosCurso := 0
    totalPuladosCurso := 0

    if modoUmAUm {
        for idx, chave in ordemPendente {
            dados := mapaDocentes[chave]
            nomeDocente := dados["nome"]
            emailReal := dados["email"]

            if !dados["emailValido"] {
                totalPuladosCurso++
                continue
            }

            link := "https://outlook.office.com/mail/deeplink/compose?to=" emailReal
                 . "&subject=" assunto
                 . "&body=" corpo
            Run(link)
            totalAbertosCurso++
            Sleep(2500)

            pastaDocente := EncontrarPastaDocente(PASTA_CERTS_ENVIO, nomeDocente)
            if (pastaDocente != "") {
                Run("explorer.exe `"" pastaDocente "`"")
                Sleep(800)
                ExibirMensagem(
                    "Anexar e Enviar — Docente " idx " de " ordemPendente.Length,
                    "Docente: " nomeDocente,
                    "E-mail aberto no navegador para: " emailReal "`n`n"
                    "O Windows Explorer abriu com a pasta do docente.`n`n"
                    "Instruções:`n"
                    "  1) Na janela do Explorer, selecione todos os PDFs (Ctrl+A);`n"
                    "  2) Arraste para dentro do e-mail aberto no Outlook;`n"
                    "  3) Confira o(s) anexo(s) e clique em Enviar no Outlook.`n`n"
                    "⚠️ Clique em 'Confirmar Envio' somente DEPOIS de ter enviado no Outlook.",
                    "info",
                    "✅ Confirmar Envio e Próximo",
                    620
                )
                MarcarPendentesComoEnviados(pastaControle, nomeDocente, chave, mesesSelecionados, enviadosMap)
            } else {
                ExibirMensagem(
                    "Sem Certificado",
                    "Docente: " nomeDocente,
                    "E-mail aberto para: " emailReal "`n`n⚠️ Pasta não encontrada (sem certificados pendentes para este docente).",
                    "erro",
                    "Continuar ➔"
                )
            }
        }
    } else {
        for chave in ordemPendente {
            dados := mapaDocentes[chave]
            emailReal := dados["email"]
            if !dados["emailValido"] {
                totalPuladosCurso++
                continue
            }
            link := "https://outlook.office.com/mail/deeplink/compose?to=" emailReal
                 . "&subject=" assunto
                 . "&body=" corpo
            Run(link)
            totalAbertosCurso++
            Sleep(800)
        }

        ExibirMensagem(
            "Abas Abertas",
            "Abas do Outlook Prontas!",
            totalAbertosCurso " aba(s) do Outlook foram abertas para o curso `"" nomePosGraduacao "`"!`n`n"
            "Clique abaixo para iniciar a confirmação de anexos, docente por docente.",
            "sucesso",
            "Iniciar Confirmação ➔",
            580
        )

        for idx, chave in ordemPendente {
            dados := mapaDocentes[chave]
            nomeDocente := dados["nome"]
            emailReal := dados["email"]
            if !dados["emailValido"]
                continue

            pastaDocente := EncontrarPastaDocente(PASTA_CERTS_ENVIO, nomeDocente)
            if (pastaDocente != "") {
                Run("explorer.exe `"" pastaDocente "`"")
                Sleep(800)
                ExibirMensagem(
                    "Anexar e Enviar — Docente " idx " de " ordemPendente.Length,
                    "Docente: " nomeDocente,
                    "Docente: " nomeDocente " <" emailReal ">`n`n"
                    "O Windows Explorer abriu com a pasta deste docente.`n`n"
                    "Instruções:`n"
                    "  1) No Explorer, selecione todos os PDFs (Ctrl+A);`n"
                    "  2) Arraste para o e-mail correspondente no Outlook;`n"
                    "  3) Confira os anexos e clique em Enviar.`n`n"
                    "⚠️ Clique abaixo somente DEPOIS de enviar.",
                    "info",
                    "✅ Confirmar Envio e Próximo",
                    620
                )
                MarcarPendentesComoEnviados(pastaControle, nomeDocente, chave, mesesSelecionados, enviadosMap)
            } else {
                ExibirMensagem(
                    "Sem Certificado",
                    "Docente: " nomeDocente,
                    "Docente: " nomeDocente " <" emailReal ">`n`n⚠️ Pasta de certificados não encontrada.",
                    "erro",
                    "Continuar ➔"
                )
            }
        }
    }

    respLimpeza := Confirmar(
        "Limpar Pasta Temporária",
        "Curso Concluído — " nomePosGraduacao,
        "Deseja apagar agora a pasta temporária de envio deste curso?`n`n"
        PASTA_CERTS_ENVIO "`n`n"
        "(Os certificados originais organizados por mês e o controle de envio NÃO são afetados.)",
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
}

ExibirMensagem(
    "Passo 10 — Relatório Final da Rodada",
    "Rodada de Envio Concluída com Sucesso!",
    "Cursos processados: " cursosRodada.Length "`n"
    "E-mails abertos: " totalGeralAbertos "`n"
    "Pulados nesta rodada (sem e-mail válido): " totalGeralPulados "`n"
    "Já enviados em rodadas anteriores (pulados automaticamente): " totalGeralJaEnviados "`n`n"
    "Modo usado: " (modoUmAUm ? "Um a um" : "Todos de uma vez") "`n`n"
    "Se ainda faltar gente, basta rodar novamente depois — o histórico de envio é preservado.",
    "sucesso",
    "Concluir e Fechar",
    620
)

ExitApp
