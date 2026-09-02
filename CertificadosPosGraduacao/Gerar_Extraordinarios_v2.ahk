#Requires AutoHotkey v2.0
#SingleInstance Force
#Include UI_Moderna.ahk

; =========================================================
; GERADOR DE CERTIFICADOS EXTRAORDINÁRIOS
; Ferramenta configurável via popups para planilhas
; com estrutura variável (formato muda a cada solicitação)
; =========================================================

; =========================================================
; CONFIGURAÇÃO DINÂMICA DE NUVEM
; =========================================================

global INFO_NUVEM := ObterCaminhosCompartilhados()

; =========================================================
; SELECIONAR PLANILHA EXCEL
; =========================================================

ExibirMensagem(
    "Certificados Extraordinários",
    "Geração de Certificados Extraordinários",
    "Este assistente permite gerar certificados a partir de planilhas avulsas e customizadas.`n`n"
    "Você configurará passo a passo a aba, células, colunas e dados do coordenador.`n`n"
    "Clique no botão abaixo para selecionar a planilha Excel (.xlsx).",
    "passo",
    "Selecionar Planilha ➔",
    580
)

pastaInicioExcel := INFO_NUVEM["raizOneDrive"] != "" ? INFO_NUVEM["raizOneDrive"] : ""
excelPath := FileSelect(1, pastaInicioExcel, "Selecione a planilha Excel", "*.xlsx")

if !excelPath {
    ExibirMensagem("Operação Cancelada", "Nenhuma Planilha Selecionada", "Nenhuma planilha Excel foi selecionada. O programa será encerrado.", "erro", "Fechar")
    ExitApp
}

; =========================================================
; ABRIR EXCEL
; =========================================================

try {
    excel := ComObject("Excel.Application")
    excel.Visible := false
    wb := excel.Workbooks.Open(excelPath)
}
catch as err {
    ExibirMensagem("Erro no Excel", "Falha ao Abrir Planilha", "Não foi possível abrir a planilha:`n`n" err.Message, "erro", "Fechar")
    ExitApp
}

; =========================================================
; INFORMAR ABA
; =========================================================

listaAbas := []
Loop wb.Sheets.Count {
    listaAbas.Push(wb.Sheets.Item(A_Index).Name)
}

abaEscolhidaArr := SelecionarOpcoesUI(
    "PASSO 1 — Selecionar Aba",
    "Escolha a Aba da Planilha",
    "Selecione a aba que contém o cronograma de aulas e docentes:",
    listaAbas,
    false,
    580
)

if (abaEscolhidaArr.Length = 0) {
    wb.Close(false)
    excel.Quit()
    ExitApp
}

nomeAbaTexto := abaEscolhidaArr[1]

ws := ""
Loop wb.Sheets.Count {
    if (Trim(wb.Sheets.Item(A_Index).Name) = nomeAbaTexto) {
        ws := wb.Sheets.Item(A_Index)
        break
    }
}

if !ws {
    wb.Close(false)
    excel.Quit()
    ExibirMensagem("Erro de Aba", "Aba Não Encontrada", "Aba '" nomeAbaTexto "' não foi encontrada na planilha.", "erro", "Fechar")
    ExitApp
}

; =========================================================
; CÉLULA DA PÓS-GRADUAÇÃO
; =========================================================

inputCelulaPos := PedirTexto(
    "PASSO 2 — Nome da Pós-Graduação",
    "Célula da Pós-Graduação",
    "Digite a referência da célula que contém o NOME da pós-graduação`n"
    "(texto a ser usado no placeholder {POSGRADUACAO}).`n`n"
    "Exemplo: B5 ou E2",
    "B5",
    "Exemplo: B5"
)

if (inputCelulaPos["Result"] != "OK" || Trim(inputCelulaPos["Value"]) = "") {
    wb.Close(false)
    excel.Quit()
    ExitApp
}

celulaPos := Trim(inputCelulaPos["Value"])

nomePos := ""
try {
    nomePos := Trim(ws.Range(celulaPos).Value)
} catch as err {
    wb.Close(false)
    excel.Quit()
    ExibirMensagem("Célula Inválida", "Erro de Referência", "Referência de célula inválida: '" celulaPos "'`n`n" err.Message, "erro", "Fechar")
    ExitApp
}

if (nomePos = "") {
    wb.Close(false)
    excel.Quit()
    ExibirMensagem("Célula Vazia", "Valor Não Encontrado", "A célula '" celulaPos "' está vazia. Verifique a referência na planilha.", "erro", "Fechar")
    ExitApp
}

confirmPos := Confirmar(
    "PASSO 2 — Confirmar Pós-Graduação",
    "Pós-Graduação Identificada",
    "Na célula " celulaPos " foi encontrado o seguinte nome:`n`n`"" nomePos "`"`n`n"
    "Este é o nome correto da Pós-Graduação para o certificado?",
    "✅ Confirmar Nome",
    "❌ Cancelar",
    true,
    580
)

if !confirmPos {
    wb.Close(false)
    excel.Quit()
    ExitApp
}

; =========================================================
; LINHA INICIAL DOS DADOS
; =========================================================

inputLinhaInicial := PedirTexto(
    "PASSO 3 — Linha Inicial dos Dados",
    "Primeira Linha de Aulas",
    "Digite o NÚMERO da primeira linha que contém um registro de aula real`n"
    "(pulando cabeçalhos e linhas em branco no topo).`n`n"
    "Exemplo: 10",
    "10",
    "Digite o número da linha (ex: 10)"
)

if (inputLinhaInicial["Result"] != "OK" || Trim(inputLinhaInicial["Value"]) = "") {
    wb.Close(false)
    excel.Quit()
    ExitApp
}

linhaInicial := Integer(Trim(inputLinhaInicial["Value"]))

if (linhaInicial < 1) {
    wb.Close(false)
    excel.Quit()
    ExibirMensagem("Linha Inválida", "Número Incorreto", "A linha inicial informada é inválida.", "erro", "Fechar")
    ExitApp
}

; =========================================================
; COLUNA DO NOME DO DOCENTE
; =========================================================

inputColNome := PedirTexto(
    "PASSO 4 — Coluna do Nome do Docente",
    "Coluna dos Docentes",
    "Digite a LETRA da coluna que contém o NOME do docente/professor.`n`n"
    "Exemplo: G ou I",
    "G",
    "Digite a letra da coluna (ex: G)"
)

if (inputColNome["Result"] != "OK" || Trim(inputColNome["Value"]) = "") {
    wb.Close(false)
    excel.Quit()
    ExitApp
}

colNomeLetra := Trim(inputColNome["Value"])

; =========================================================
; COLUNA DA DISCIPLINA / AULA
; =========================================================

inputColCurso := PedirTexto(
    "PASSO 5 — Coluna da Disciplina / Aula",
    "Coluna da Disciplina / Tema",
    "Digite a LETRA da coluna que contém o NOME DA AULA ou DISCIPLINA`n"
    "(placeholder {CURSO}).`n`n"
    "Exemplo: F ou H",
    "F",
    "Digite a letra da coluna (ex: F)"
)

if (inputColCurso["Result"] != "OK" || Trim(inputColCurso["Value"]) = "") {
    wb.Close(false)
    excel.Quit()
    ExitApp
}

colCursoLetra := Trim(inputColCurso["Value"])

; =========================================================
; FORMATO DA DATA: ÚNICA OU SEPARADA (DIA/MÊS/ANO)
; =========================================================

dataEmColunaUnica := Confirmar(
    "PASSO 6 — Formato da Data",
    "Organização das Datas na Planilha",
    "Como a data da aula está organizada nesta planilha?",
    "📅 Em uma ÚNICA coluna (ex: 15/05/2026)",
    "📑 Em colunas SEPARADAS (Dia, Mês, Ano)",
    true,
    600
)

if dataEmColunaUnica {

    inputColData := PedirTexto(
        "PASSO 6.1 — Coluna da Data Única",
        "Coluna da Data Completa",
        "Digite a LETRA da coluna que contém a DATA COMPLETA da aula.`n`n"
        "Exemplo: A",
        "A",
        "Digite a letra da coluna (ex: A)"
    )

    if (inputColData["Result"] != "OK" || Trim(inputColData["Value"]) = "") {
        wb.Close(false)
        excel.Quit()
        ExitApp
    }

    colDataLetra := Trim(inputColData["Value"])

} else {

    inputColDia := PedirTexto(
        "PASSO 6.1 — Coluna do Dia",
        "Coluna do DIA da Aula",
        "Digite a LETRA da coluna que contém o número do DIA da aula.`n`n"
        "Exemplo: A ou B",
        "A",
        "Digite a letra da coluna (ex: A)"
    )
    if (inputColDia["Result"] != "OK" || Trim(inputColDia["Value"]) = "") {
        wb.Close(false)
        excel.Quit()
        ExitApp
    }
    colDiaLetra := Trim(inputColDia["Value"])

    inputColMes := PedirTexto(
        "PASSO 6.2 — Coluna do Mês",
        "Coluna do MÊS da Aula",
        "Digite a LETRA da coluna que contém o MÊS da aula (aceita número ou nome do mês).`n`n"
        "Exemplo: B ou C",
        "B",
        "Digite a letra da coluna (ex: B)"
    )
    if (inputColMes["Result"] != "OK" || Trim(inputColMes["Value"]) = "") {
        wb.Close(false)
        excel.Quit()
        ExitApp
    }
    colMesLetra := Trim(inputColMes["Value"])

    inputColAno := PedirTexto(
        "PASSO 6.3 — Coluna do Ano",
        "Coluna do ANO da Aula",
        "Digite a LETRA da coluna que contém o ANO da aula (ex: 2026).`n`n"
        "Exemplo: C ou D",
        "C",
        "Digite a letra da coluna (ex: C)"
    )
    if (inputColAno["Result"] != "OK" || Trim(inputColAno["Value"]) = "") {
        wb.Close(false)
        excel.Quit()
        ExitApp
    }
    colAnoLetra := Trim(inputColAno["Value"])
}

; =========================================================
; COLUNA DO HORÁRIO
; =========================================================

inputColHorario := PedirTexto(
    "PASSO 7 — Coluna do Horário",
    "Coluna do Horário da Aula",
    "Digite a LETRA da coluna que contém o HORÁRIO da aula (ex: '20h-22h', '08:00-10:00').`n`n"
    "Exemplo: D ou E",
    "D",
    "Digite a letra da coluna (ex: D)"
)

if (inputColHorario["Result"] != "OK" || Trim(inputColHorario["Value"]) = "") {
    wb.Close(false)
    excel.Quit()
    ExitApp
}

colHorarioLetra := Trim(inputColHorario["Value"])

; =========================================================
; PASSO 8 — COORDENADOR
; =========================================================

inputCoordenador := PedirTexto(
    "PASSO 8 — Coordenador do Curso",
    "Nome do Coordenador",
    "Digite o NOME COMPLETO do coordenador responsável pelo curso:`n`n"
    "Exemplo: Marcelo Bettega",
    "Marcelo Bettega",
    "Nome completo do(a) coordenador(a)"
)

if (inputCoordenador["Result"] != "OK" || Trim(inputCoordenador["Value"]) = "") {
    wb.Close(false)
    excel.Quit()
    ExitApp
}

coordUsado := Trim(inputCoordenador["Value"])

; =========================================================
; PASSO 9 — CARGO
; =========================================================

respCargo := Confirmar(
    "PASSO 9 — Cargo do Coordenador",
    "Definir Cargo — " coordUsado,
    "Selecione o cargo que deve constar no certificado:",
    "Coordenadora do Curso",
    "Coordenador do Curso",
    false,
    550
)

cargo := respCargo ? "Coordenadora do Curso" : "Coordenador do Curso"

; =========================================================
; PASSO 10 — ASSINATURA (SELEÇÃO MANUAL — PADRÃO NESTE FLUXO)
; =========================================================

ExibirMensagem(
    "PASSO 10 — Assinatura do Coordenador",
    "Selecionar Imagem da Assinatura",
    "Selecione na próxima janela o arquivo de imagem (.jpg ou .png) com a assinatura de " coordUsado ".",
    "passo",
    "Selecionar Arquivo ➔",
    580
)

pastaInicioAssin := (INFO_NUVEM["pastaAssinaturas"] != "" && DirExist(INFO_NUVEM["pastaAssinaturas"])) ? INFO_NUVEM["pastaAssinaturas"] : "C:\CAssinaturas\"
assinaturaPath := FileSelect(1, pastaInicioAssin, "Selecione a assinatura de " coordUsado, "*.jpg; *.png")

if !assinaturaPath {
    wb.Close(false)
    excel.Quit()
    ExibirMensagem("Operação Cancelada", "Nenhuma Assinatura Selecionada", "Nenhuma assinatura foi selecionada. Operação cancelada.", "erro", "Fechar")
    ExitApp
}

; =========================================================
; PASSO 11 — DATA DE EMISSÃO
; =========================================================

hoje := FormatTime(A_Now, "dd/MM/yyyy")

respDataEmissao := Confirmar(
    "PASSO 11 — Data de Emissão",
    "Data de Emissão do Certificado",
    "Data de hoje identificada:`n`n" hoje "`n`nDeseja utilizar esta data no certificado?",
    "✅ Usar Hoje (" hoje ")",
    "✏️ Digitar Outra Data",
    true,
    580
)

if respDataEmissao {
    dataEmissao := hoje
} else {
    inputDataEmissao := PedirTexto(
        "Data de Emissão Personalizada",
        "Informar Data de Emissão",
        "Digite a data de emissão no formato DD/MM/AAAA:`n`nExemplo: 09/06/2026",
        hoje,
        "DD/MM/AAAA"
    )
    if (inputDataEmissao["Result"] != "OK" || Trim(inputDataEmissao["Value"]) = "") {
        wb.Close(false)
        excel.Quit()
        ExitApp
    }
    dataEmissao := Trim(inputDataEmissao["Value"])
}

; =========================================================
; DIAGNÓSTICO FINAL
; =========================================================

resumoData := dataEmColunaUnica
    ? "  COLUNA DATA (única): " colDataLetra "`n"
    : "  COLUNA DIA: " colDiaLetra "  |  COLUNA MÊS: " colMesLetra "  |  COLUNA ANO: " colAnoLetra "`n"

confirmFinal := Confirmar(
    "Diagnóstico Final",
    "Confira as Configurações da Rodada",
    "ABA:                " nomeAbaTexto "`n"
    "PÓS-GRADUAÇÃO:      " nomePos " (célula " celulaPos ")`n"
    "LINHA INICIAL:      " linhaInicial "`n"
    "COLUNA NOME:        " colNomeLetra "`n"
    "COLUNA DISCIPLINA:  " colCursoLetra "`n"
    resumoData
    "COLUNA HORÁRIO:     " colHorarioLetra "`n`n"
    "COORDENADOR:        " coordUsado "`n"
    "CARGO:              " cargo "`n"
    "DATA DE EMISSÃO:    " dataEmissao "`n`n"
    "Deseja iniciar a geração dos certificados?",
    "🚀 Iniciar Geração",
    "❌ Cancelar",
    true,
    620
)

if !confirmFinal {
    wb.Close(false)
    excel.Quit()
    ExitApp
}

; =========================================================
; ESCOLHER MODELO POWERPOINT
; =========================================================

pptModelo := ""
if (INFO_NUVEM["modeloPptx"] != "" && FileExist(INFO_NUVEM["modeloPptx"])) {
    pptModelo := INFO_NUVEM["modeloPptx"]
} else {
    ExibirMensagem(
        "Modelo PowerPoint",
        "Selecionar Modelo do Certificado",
        "Selecione na próxima janela o arquivo PowerPoint (.pptx) com o layout do certificado.",
        "passo",
        "Selecionar Modelo ➔",
        580
    )
    pastaInicioPpt := INFO_NUVEM["pastaDeclaracoes"] != "" ? INFO_NUVEM["pastaDeclaracoes"] : ""
    pptModelo := FileSelect(1, pastaInicioPpt, "Selecione o modelo PowerPoint", "*.pptx")
}

if !pptModelo {
    wb.Close(false)
    excel.Quit()
    ExitApp
}

; =========================================================
; PASTA DE SAÍDA
; =========================================================

pastaSaida := "C:\Certificados\"
DirCreate(pastaSaida)

; =========================================================
; ABRIR POWERPOINT
; =========================================================

try {
    ppt := ComObject("PowerPoint.Application")
    ppt.Visible := true
}
catch as err {
    MsgBox("Erro ao abrir PowerPoint.`n`n" err.Message)
    wb.Close(false)
    excel.Quit()
    ExitApp
}

; =========================================================
; CRIAR MODELO-BASE COM COORDENADOR + CARGO + DATA + ASSINATURA
; (mesma lógica do script Master — feito uma única vez)
; =========================================================

modeloBase := A_Temp "\certificado_base_extra_" A_TickCount ".pptx"

try {
    presBase  := ppt.Presentations.Open(pptModelo)
    slideBase := presBase.Slides(1)

    ; --- Substituir placeholders de texto ---
    Loop slideBase.Shapes.Count {
        shape := slideBase.Shapes.Item(A_Index)
        try {
            txt := shape.TextFrame.TextRange.Text
            if InStr(txt, "{COORDENADOR}") || InStr(txt, "{CARGO}") || InStr(txt, "{DATAEMISSAO}") {
                SubstituirPlaceholder(shape, "{COORDENADOR}", coordUsado)
                SubstituirPlaceholder(shape, "{CARGO}",       cargo)
                SubstituirPlaceholder(shape, "{DATAEMISSAO}", dataEmissao)
            }
        }
        catch {
        }
    }

    ; --- Localizar imagem de assinatura no slide ---
    primeiroNomeCoord := StrSplit(coordUsado, " ")[1]

    shapeSelecionado := EncontrarImagemPorNome(slideBase.Shapes, coordUsado, primeiroNomeCoord)

    if !shapeSelecionado {
        ultimaImagem := EncontrarUltimaImagem(slideBase.Shapes)
        if ultimaImagem {
            shapeSelecionado := ultimaImagem
            MsgBox(
                "Atenção: não foi possível identificar a assinatura pelo nome da forma.`n`n"
                "Será utilizada a última imagem encontrada no slide como referência de posição.`n`n"
                "Verifique o resultado após a geração.",
                "Aviso — Identificação da Assinatura",
                "Icon!"
            )
        }
    }

    if shapeSelecionado {
        posLeft   := shapeSelecionado.Left
        posTop    := shapeSelecionado.Top
        posWidth  := shapeSelecionado.Width
        posHeight := shapeSelecionado.Height

        shapeSelecionado.Delete()

        picAssinatura := slideBase.Shapes.AddPicture(
            assinaturaPath,
            false,
            true,
            posLeft,
            posTop,
            posWidth,
            posHeight
        )
        picAssinatura.Name := "ASSINATURA_REAL"
    } else {
        shapeCoord := EncontrarShapeComTexto(slideBase.Shapes, "{COORDENADOR}")

        if !shapeCoord
            shapeCoord := EncontrarShapeComTexto(slideBase.Shapes, coordUsado)

        if shapeCoord {
            larguraPadrao := 120
            alturaPadrao  := 45

            posLeft := shapeCoord.Left + (shapeCoord.Width - larguraPadrao) / 2
            posTop  := shapeCoord.Top - alturaPadrao - 5

            picAssinatura := slideBase.Shapes.AddPicture(
                assinaturaPath,
                false,
                true,
                posLeft,
                posTop,
                larguraPadrao,
                alturaPadrao
            )
            picAssinatura.Name := "ASSINATURA_REAL"

            MsgBox(
                "Nenhuma imagem foi encontrada no modelo para usar como referência.`n`n"
                "A assinatura foi posicionada automaticamente logo acima do texto`n"
                "do coordenador, com tamanho padrão (120 x 45 pontos).`n`n"
                "Ajuste a posição/tamanho manualmente no PowerPoint se necessário,`n"
                "ou adicione uma imagem de referência no slide-modelo para`n"
                "posicionamento automático mais preciso da próxima vez.",
                "Aviso — Assinatura Posicionada Automaticamente",
                "Icon!"
            )
        } else {
            MsgBox(
                "Nenhuma imagem encontrada no slide para substituir, e também`n"
                "não foi possível localizar o texto do coordenador para`n"
                "posicionar a assinatura automaticamente.`n`n"
                "A assinatura NÃO será inserida.",
                "Erro — Assinatura",
                "Icon!"
            )
        }
    }

    presBase.SaveAs(modeloBase, 24)
    presBase.Close()

} catch as err {
    try ppt.Quit()
    wb.Close(false)
    excel.Quit()
    MsgBox("Erro ao criar modelo-base.`n`n" err.Message)
    ExitApp
}

; =========================================================
; ÚLTIMA LINHA (baseado na coluna do nome)
; =========================================================

colNomeIndice := ColLetraParaIndice(colNomeLetra)
ultimaLinha   := ws.Cells(ws.Rows.Count, colNomeIndice).End(-4162).Row

if (ultimaLinha < linhaInicial) {
    try ppt.Quit()
    wb.Close(false)
    excel.Quit()
    MsgBox("Nenhum dado encontrado a partir da linha " linhaInicial " na coluna " colNomeLetra ".")
    ExitApp
}

; =========================================================
; PRIMEIRA PASSAGEM — SEPARAR OK x INVÁLIDOS
; =========================================================

registrosOk       := []
registrosInvalido := []

Loop ultimaLinha - linhaInicial + 1 {

    linha := linhaInicial + A_Index - 1

    nome  := Trim(ws.Range(colNomeLetra  linha).Value)
    curso := Trim(ws.Range(colCursoLetra linha).Value)

    nome := RegExReplace(nome, "\s*\b[HMhm]\b\s*$")
    nome := Trim(nome)

    ; --- Montar data ---
    dataAula := ""
    erroData := false

    if dataEmColunaUnica {
        dataRaw := Trim(ws.Range(colDataLetra linha).Text)
        if (dataRaw = "")
            erroData := true
        else
            dataAula := dataRaw
    } else {
        diaRaw := Trim(ws.Range(colDiaLetra linha).Text)
        mesRaw := Trim(ws.Range(colMesLetra linha).Text)
        anoRaw := Trim(ws.Range(colAnoLetra linha).Text)

        if (diaRaw = "" || mesRaw = "" || anoRaw = "") {
            erroData := true
        } else {
            mesNumero := NormalizarMes(mesRaw)
            if (mesNumero = "")
                erroData := true
            else
                dataAula := diaRaw "/" mesNumero "/" anoRaw
        }
    }

    ; --- Horário ---
    horarioCell := ws.Range(colHorarioLetra linha)
    horarioTxt  := String(horarioCell.Text)

    if (horarioTxt = "" || RegExMatch(horarioTxt, "^\d[\.,]\d")) {
        horarioVal      := horarioCell.Value
        horasFormatadas := FormatarHorasDeValorExcel(horarioVal)
    } else {
        horasFormatadas := FormatarHorasSeguro(horarioTxt)
    }

    erroNome  := (nome = "")
    erroCurso := (curso = "")
    erroHoras := InStr(horasFormatadas, "invalido")

    if (erroNome && erroCurso && erroData)
        continue

    registro := Map(
        "linha",           linha,
        "nome",            nome,
        "curso",           curso,
        "dataAula",        dataAula,
        "horasFormatadas", horasFormatadas,
        "horarioTxt",      horarioTxt,
        "erroNome",        erroNome,
        "erroCurso",       erroCurso,
        "erroData",        erroData,
        "erroHoras",       erroHoras
    )

    if (erroNome || erroCurso || erroData || erroHoras)
        registrosInvalido.Push(registro)
    else
        registrosOk.Push(registro)
}

; =========================================================
; SEGUNDA PASSAGEM — GERAR CERTIFICADOS OK
; =========================================================

totalGerados := 0
totalErros   := 0

for registro in registrosOk {
    resultado := GerarCertificado(registro, ppt, modeloBase, pastaSaida, nomePos)
    if resultado
        totalGerados++
    else
        totalErros++
}

; =========================================================
; TERCEIRA PASSAGEM — CORRIGIR INVÁLIDOS E GERAR
; =========================================================

if (registrosInvalido.Length > 0) {

    ExibirMensagem(
        "Certificados com Erro",
        "Inconsistências Encontradas",
        "Foram encontrados " registrosInvalido.Length " registro(s) com dados incompletos ou formatos inválidos.`n`n"
        "Clique abaixo para iniciar a correção manual passo a passo.",
        "aviso",
        "Iniciar Correção ➔",
        580
    )

    for registro in registrosInvalido {

        linha           := registro["linha"]
        nome            := registro["nome"]
        curso           := registro["curso"]
        dataAula        := registro["dataAula"]
        horasFormatadas := registro["horasFormatadas"]
        horarioTxt      := registro["horarioTxt"]
        erroNome        := registro["erroNome"]
        erroCurso       := registro["erroCurso"]
        erroData        := registro["erroData"]
        erroHoras       := registro["erroHoras"]

        listaErros  := ""
        ordemCampos := ""
        totalErrosCampo := 0

        if erroNome {
            totalErrosCampo++
            listaErros  .= "  • Nome do docente: Digite o nome completo (Ex: Maria da Silva Santos)`n"
            ordemCampos .= (ordemCampos = "") ? "NOME" : ", NOME"
        }
        if erroData {
            totalErrosCampo++
            listaErros  .= "  • Data da aula: Digite no formato DD/MM/AA (Ex: 15/05/26)`n"
            ordemCampos .= (ordemCampos = "") ? "DATA" : ", DATA"
        }
        if erroCurso {
            totalErrosCampo++
            listaErros  .= "  • Disciplina / Tema: Digite o tema completo da aula`n"
            ordemCampos .= (ordemCampos = "") ? "CURSO" : ", CURSO"
        }
        if erroHoras {
            totalErrosCampo++
            listaErros  .= "  • Carga horária: Digite o total em MINUTOS (60 = 1h, 90 = 1h30, 120 = 2h)`n"
            ordemCampos .= (ordemCampos = "") ? "MINUTOS" : ", MINUTOS"
        }

        instrucao := (totalErrosCampo = 1)
            ? "Insira o valor correto no campo abaixo:"
            : "Insira os valores na ordem (" ordemCampos "), separados por vírgula:"

        msg := "O registro da linha " linha " apresentou as seguintes pendências:`n`n"
        msg .= listaErros "`n"
        msg .= instrucao

        correcao := PedirTexto(
            "Correção Manual — Linha " linha,
            "Corrigir Registro (Linha " linha ")",
            msg,
            "",
            (totalErrosCampo = 1 ? "Digite o valor correto" : "Exemplo: " ordemCampos),
            620
        )

        if (correcao["Result"] != "OK" || Trim(correcao["Value"]) = "") {
            ExibirMensagem(
                "Certificado Pulado",
                "Registro Ignorado (Linha " linha ")",
                "Docente: " (nome = "" ? "(vazio)" : nome) "`nData: " (dataAula = "" ? "(vazio)" : dataAula) "`n`nEste registro foi ignorado.",
                "erro",
                "Continuar ➔"
            )
            totalErros++
            continue
        }

        partes := StrSplit(correcao["Value"], ",")
        idx    := 1

        if erroNome {
            if (partes.Length >= idx && Trim(partes[idx]) != "")
                nome := Trim(partes[idx])
            idx++
        }
        if erroData {
            if (partes.Length >= idx && Trim(partes[idx]) != "")
                dataAula := Trim(partes[idx])
            idx++
        }
        if erroCurso {
            if (partes.Length >= idx && Trim(partes[idx]) != "")
                curso := Trim(partes[idx])
            idx++
        }
        if erroHoras {
            if (partes.Length >= idx && Trim(partes[idx]) != "") {
                minutosInformados := Integer(Trim(partes[idx]))
                horasFormatadas   := MinutosParaExtenso(minutosInformados)
            }
            idx++
        }

        if (nome = "" || curso = "" || dataAula = "") {
            ExibirMensagem(
                "Campos Obrigatórios Vazios",
                "Registro Incompleto (Linha " linha ")",
                "O registro continua sem nome, curso ou data preenchidos e foi ignorado.",
                "erro",
                "Continuar ➔"
            )
            totalErros++
            continue
        }

        registro["nome"]            := nome
        registro["curso"]           := curso
        registro["dataAula"]        := dataAula
        registro["horasFormatadas"] := horasFormatadas

        resultado := GerarCertificado(registro, ppt, modeloBase, pastaSaida, nomePos)
        if resultado
            totalGerados++
        else
            totalErros++
    }
}

; =========================================================
; FINALIZAÇÃO
; =========================================================

try ppt.Quit()

wb.Close(false)
excel.Quit()

try FileDelete(modeloBase)

; =========================================================
; RELATÓRIO FINAL
; =========================================================

ExibirMensagem(
    "Relatório Final",
    "Geração Extraordinária Concluída!",
    "PÓS-GRADUAÇÃO:         " nomePos "`n"
    "COORDENADOR:           " coordUsado "`n"
    "CARGO:                 " cargo "`n"
    "DATA DE EMISSÃO:       " dataEmissao "`n"
    "ASSINATURA USADA:      " assinaturaPath "`n`n"
    "CERTIFICADOS GERADOS:  " totalGerados "`n"
    "ERROS / PULADOS:       " totalErros "`n`n"
    "PASTA DE SAÍDA:`n" pastaSaida,
    "sucesso",
    "Concluir e Fechar",
    620
)

ExitApp

; =========================================================
; FUNÇÃO: Converte letra de coluna (A, B, ... AA) em índice numérico
; =========================================================

ColLetraParaIndice(letra) {
    letra := StrUpper(Trim(letra))
    indice := 0
    Loop Parse, letra {
        indice := indice * 26 + (Ord(A_LoopField) - Ord("A") + 1)
    }
    return indice
}

; =========================================================
; FUNÇÃO: Normaliza nome do mês (com ou sem acento, abreviado
; ou completo, qualquer caixa) para número (1 a 12).
; Retorna "" se não conseguir reconhecer.
; =========================================================

NormalizarMes(valor) {
    valor := Trim(valor)

    ; Se já é número válido entre 1 e 12, retorna direto
    if RegExMatch(valor, "^\d{1,2}$") {
        n := Integer(valor)
        if (n >= 1 && n <= 12)
            return String(n)
    }

    ; Normalizar: minúsculo e sem acentos comuns em português
    texto := StrLower(valor)
    texto := StrReplace(texto, "á", "a")
    texto := StrReplace(texto, "ã", "a")
    texto := StrReplace(texto, "â", "a")
    texto := StrReplace(texto, "é", "e")
    texto := StrReplace(texto, "ê", "e")
    texto := StrReplace(texto, "í", "i")
    texto := StrReplace(texto, "ó", "o")
    texto := StrReplace(texto, "õ", "o")
    texto := StrReplace(texto, "ô", "o")
    texto := StrReplace(texto, "ú", "u")
    texto := StrReplace(texto, "ç", "c")
    texto := Trim(texto)

    mapaMeses := Map(
        "janeiro", "1",  "jan", "1",
        "fevereiro", "2", "fev", "2",
        "marco", "3",    "mar", "3",
        "abril", "4",    "abr", "4",
        "maio", "5",     "mai", "5",
        "junho", "6",    "jun", "6",
        "julho", "7",    "jul", "7",
        "agosto", "8",   "ago", "8",
        "setembro", "9", "set", "9",
        "outubro", "10", "out", "10",
        "novembro", "11","nov", "11",
        "dezembro", "12","dez", "12"
    )

    if mapaMeses.Has(texto)
        return mapaMeses[texto]

    return ""
}

; =========================================================
; FUNÇÃO: Verifica se uma shape é uma imagem (foto solta OU
; placeholder de imagem preenchido). Usada pelas buscas abaixo.
; =========================================================

EhImagem(s) {
    try {
        if (s.Type = 13)  ; 13 = msoPicture
            return true
    } catch {
    }
    try {
        if (s.Type = 14 && s.PlaceholderFormat.Type = 18)
            return true
    } catch {
    }
    return false
}

; =========================================================
; FUNÇÃO: Busca recursiva (entra em grupos) por uma imagem cujo
; nome contenha o nome do coordenador ou seu primeiro nome.
; =========================================================

EncontrarImagemPorNome(shapes, coordUsado, primeiroNomeCoord) {
    Loop shapes.Count {
        s := shapes.Item(A_Index)
        try {
            if (s.Type = 6) {  ; 6 = msoGroup -> busca dentro do grupo
                achado := EncontrarImagemPorNome(s.GroupItems, coordUsado, primeiroNomeCoord)
                if achado
                    return achado
                continue
            }
        } catch {
        }
        if EhImagem(s) {
            try {
                nomeForma := s.Name
                if InStr(nomeForma, coordUsado) || InStr(nomeForma, primeiroNomeCoord)
                    return s
            } catch {
            }
        }
    }
    return ""
}

; =========================================================
; FUNÇÃO: Busca recursiva (entra em grupos) pela última imagem
; encontrada no slide, na ordem das shapes.
; =========================================================

EncontrarUltimaImagem(shapes) {
    ultima := ""
    Loop shapes.Count {
        s := shapes.Item(A_Index)
        try {
            if (s.Type = 6) {  ; 6 = msoGroup -> busca dentro do grupo
                achadoGrupo := EncontrarUltimaImagem(s.GroupItems)
                if achadoGrupo
                    ultima := achadoGrupo
                continue
            }
        } catch {
        }
        if EhImagem(s)
            ultima := s
    }
    return ultima
}

; =========================================================
; FUNÇÃO: Busca recursiva (entra em grupos) pela primeira shape
; de texto que contenha um determinado trecho.
; =========================================================

EncontrarShapeComTexto(shapes, trecho) {
    Loop shapes.Count {
        s := shapes.Item(A_Index)
        try {
            if (s.Type = 6) {  ; 6 = msoGroup -> busca dentro do grupo
                achado := EncontrarShapeComTexto(s.GroupItems, trecho)
                if achado
                    return achado
                continue
            }
        } catch {
        }
        try {
            if s.HasTextFrame && InStr(s.TextFrame.TextRange.Text, trecho)
                return s
        } catch {
        }
    }
    return ""
}

; =========================================================
; FUNÇÃO: Substituir placeholder preservando formatação
; =========================================================

SubstituirPlaceholder(shape, placeholder, valor) {
    try {
        tr := shape.TextFrame.TextRange
        Loop {
            textoAtual := tr.Text
            pos := InStr(textoAtual, placeholder)
            if !pos
                break
            chars := tr.Characters(pos, StrLen(placeholder))
            fNome  := chars.Characters(1, 1).Font.Name
            fSize  := chars.Characters(1, 1).Font.Size
            fBold  := chars.Characters(1, 1).Font.Bold
            fColor := chars.Characters(1, 1).Font.Color
            chars.Text := valor
            tr.Characters(pos, StrLen(valor)).Font.Name  := fNome
            tr.Characters(pos, StrLen(valor)).Font.Size  := fSize
            tr.Characters(pos, StrLen(valor)).Font.Bold  := fBold
            tr.Characters(pos, StrLen(valor)).Font.Color := fColor
        }
    }
    catch {
    }
}

; =========================================================
; FUNÇÃO: Gerar certificado PDF a partir de um registro
; =========================================================

GerarCertificado(registro, ppt, pptModelo, pastaSaida, nomePos) {

    nome            := registro["nome"]
    curso           := registro["curso"]
    dataAula        := registro["dataAula"]
    horasFormatadas := registro["horasFormatadas"]
    linha           := registro["linha"]

    try {

        pres  := ppt.Presentations.Open(pptModelo)
        slide := pres.Slides(1)

        shapeCount := slide.Shapes.Count

        Loop shapeCount {
            shape := slide.Shapes.Item(A_Index)
            try {
                if !InStr(shape.TextFrame.TextRange.Text, "{NOME}")
                    continue
                SubstituirPlaceholder(shape, "{NOME}",         nome)
                SubstituirPlaceholder(shape, "{POSGRADUACAO}", nomePos)
                SubstituirPlaceholder(shape, "{CURSO}",        curso)
                SubstituirPlaceholder(shape, "{DATA}",         dataAula)
                SubstituirPlaceholder(shape, "{HORAS}",        horasFormatadas)
            }
            catch {
            }
        }

        nomeLimpo := RegExReplace(nome, '[\\/:*?"<>|]')
        dataLimpa := RegExReplace(dataAula, '[\\/:*?"<>|]', ".")

        nomeBase   := nomeLimpo " (" dataLimpa ")"
        caminhoPDF := pastaSaida nomeBase ".pdf"

        if FileExist(caminhoPDF) {
            contador := 1
            Loop {
                caminhoPDF := pastaSaida nomeBase "." contador ".pdf"
                if !FileExist(caminhoPDF)
                    break
                contador++
            }
        }

        pres.SaveAs(caminhoPDF, 32)
        pres.Close()
        return true

    }
    catch as err {
        MsgBox("Erro na linha " linha ".`n`n" err.Message)
        return false
    }
}

; =========================================================
; FUNÇÃO: Converte minutos informados manualmente em extenso
; =========================================================

MinutosParaExtenso(minutos) {
    if (minutos <= 0)
        return "Horario invalido"
    horas := Floor(minutos / 60)
    resto := Mod(minutos, 60)
    if (horas > 0 && resto > 0) {
        horaLabel := (horas = 1) ? "1 hora" : horas " horas"
        return horaLabel " e " resto " minutos"
    }
    if (horas > 0)
        return (horas = 1) ? "1 hora" : horas " horas"
    return minutos " minutos"
}
