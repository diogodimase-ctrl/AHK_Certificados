#Requires AutoHotkey v2.0
#SingleInstance Force
#Include UI_Moderna.ahk

; ════════════════════════════════════════════════════════════════════════
;   GERAR CERTIFICADOS — PÓS-GRADUAÇÃO
;   Versão Atualizada: Detecção Automática de Colunas & UI Moderna
; ════════════════════════════════════════════════════════════════════════
;
;   DESTAQUES DESTA VERSÃO:
;   - Detecção Automática de Colunas (Passo 8):
;     O script agora analisa a Linha 3 da planilha do curso e identifica
;     automaticamente a coluna onde está escrito "Docentes" (e a coluna de
;     Aulas/Disciplinas), eliminando a necessidade de digitar letras a cada curso.;     Salva os PDFs na pasta compartilhada do OneDrive/SharePoint e mantém
;     cópia espelho de backup em C:\Certificados\<Curso>\ e C:\Certificados\<Curso>\Docentes\
;   - v18: abre as planilhas do Excel suprimindo os diálogos de "Atualizar
;     Links" / "Recomendado Somente Leitura". Com o Excel invisível esses
;     diálogos ficavam travados sem ninguém pra clicar, e a próxima chamada
;     COM falhava com 0x80010001 (RPC_E_CALL_REJECTED), aparecendo como o
;     erro "This value of type string has no property named Count" na
;     listagem de abas.
;   - v19/v20 tentaram mexer no Workbooks.Open pra deixar essa abertura mais
;     robusta, mas isso causou um erro NOVO (0x800A03EC) em pelo menos um
;     computador — testado pelo usuário, que confirmou que a lógica de
;     abertura da v18 (sem as mudanças da v19/v20) é a que funciona
;     corretamente na prática. Por isso a v21 parte da v18 (não da v20) e só
;     adiciona as correções de SALVAMENTO abaixo, sem tocar em
;     AbrirPastaDeTrabalho.
;   - v21: duas correções no salvamento dos PDFs (mesma lógica que tinha sido
;     colocada na v19/v20) —
;     (1) GarantirPastasSaida (em UI_Moderna.ahk) procura uma pasta já
;     existente em cada nível (Semestre → Pós → Mês) dentro de "Declarações"
;     no SharePoint/OneDrive antes de criar uma nova, evitando pastas
;     duplicadas por causa de acento/maiúscula/espaço diferente.
;     (2) Se salvar na nuvem falhar (ex.: OneDrive fora do ar/sincronizando),
;     o certificado agora cai para o backup local em vez de não salvar em
;     lugar nenhum — antes, uma falha ao salvar na nuvem abortava o
;     certificado inteiro sem nunca tentar o C:\Certificados.
;   - v22: o erro 0x80010001 ao listar as abas voltou a acontecer rodando a
;     v21 (que usa a MESMA lógica de abertura da v18) — ou seja, não é causado
;     por nenhuma diferença de código entre as versões, é intermitente: o
;     Excel às vezes ainda está processando internamente logo após abrir o
;     arquivo e rejeita a chamada COM seguinte, funcionando numa hora e
;     falhando noutra com o código idêntico. Nova função ListarNomesAbas()
;     tenta de novo automaticamente (até 5x, com pequena espera entre
;     tentativas) antes de mostrar erro pro usuário, tanto pra planilha
;     mestre quanto pra planilha de cada curso. AbrirPastaDeTrabalho() em si
;     continua exatamente igual à v18 — nada mudou na abertura do arquivo.
; ════════════════════════════════════════════════════════════════════════

; =========================================================
; CONFIGURAÇÕES DINÂMICAS DE PASTAS, LOGO E NUVEM
; =========================================================

global INFO_NUVEM        := ObterCaminhosCompartilhados()
global PASTA_ASSINATURAS := (INFO_NUVEM["pastaAssinaturas"] != "" && DirExist(INFO_NUVEM["pastaAssinaturas"])) ? INFO_NUVEM["pastaAssinaturas"] : "C:\CAssinaturas\"
global PASTA_SAIDA       := "C:\Certificados\"
global ARQUIVO_LOGO      := "C:\CAssinaturas\_logo_einstein.jpg"
global SEMESTRE_ATUAL    := "2026-5"

; =========================================================
; PASSO 1 — SELECIONAR PLANILHA MESTRE
; =========================================================

masterPath := ""

if (INFO_NUVEM["planilhaMestre"] != "" && FileExist(INFO_NUVEM["planilhaMestre"])) {
    masterPath := INFO_NUVEM["planilhaMestre"]
    usarMestreAuto := Confirmar(
        "Passo 1 — Planilha Mestre",
        "Planilha Mestre Localizada no SharePoint / OneDrive",
        "O sistema localizou a planilha mestre mais recente na nuvem:`n`n"
        "📁 " masterPath "`n`n"
        "Deseja utilizar esta planilha automaticamente?",
        "✅ Usar Planilha da Nuvem",
        "📁 Escolher Outro Arquivo",
        true,
        600
    )
    if !usarMestreAuto {
        masterPath := FileSelect(1, INFO_NUVEM["pastaDeclaracoes"], "Passo 1 - Selecione a planilha Controle de Declaração de Aula", "*.xlsx")
    }
} else {
    ExibirMensagem(
        "Passo 1 — Planilha Mestre",
        "Planilha Mestre de Controle",
        "Selecione o arquivo Excel que contém a LISTA GERAL de todos os cursos e seus respectivos coordenadores.`n`n"
        "Esse é o arquivo mestre chamado 'Controle de Declaração de Aula'.",
        "passo",
        "Selecionar Planilha ➔",
        580
    )
    masterPath := FileSelect(1, "", "Passo 1 - Selecione a planilha Controle de Declaração de Aula", "*.xlsx")
}

if !masterPath {
    ExibirMensagem("Operação Cancelada", "Nenhuma Planilha Selecionada", "Nenhuma planilha Excel foi selecionada. O programa será encerrado.", "erro", "Fechar")
    ExitApp
}

try {
    excel := ComObject("Excel.Application")
    excel.Visible := false
    excel.DisplayAlerts := false
    excel.AskToUpdateLinks := false
    wb := AbrirPastaDeTrabalho(excel, masterPath)
} catch as err {
    ExibirMensagem("Erro no Excel", "Falha ao Abrir Planilha Mestre", "Não foi possível abrir o arquivo Excel:`n`n" err.Message, "erro", "Fechar")
    ExitApp
}

; =========================================================
; PASSO 2 — LISTAR ABAS E PEDIR SELEÇÃO DO SEMESTRE
; =========================================================

abasMestreArr := []
try {
    abasMestreArr := ListarNomesAbas(wb)
} catch as err {
    try wb.Close(false)
    try excel.Quit()
    ExibirMensagem(
        "Erro no Excel", "Falha ao Ler Abas da Planilha Mestre",
        "O Excel abriu o arquivo mas não respondeu ao listar as abas`n"
        "(costuma acontecer quando um diálogo do Excel — atualizar links,`n"
        "recuperação de arquivo, etc. — fica travado em segundo plano).`n`n"
        "Feche o Excel manualmente (verifique a barra de tarefas/gerenciador`n"
        "de tarefas) e tente novamente.`n`nDetalhe técnico: " err.Message,
        "erro", "Fechar"
    )
    ExitApp
}

selecaoAba := SelecionarOpcoesUI(
    "Passo 2 — Selecionar Semestre / Aba",
    "Aba da Planilha Mestre (Semestre)",
    abasMestreArr,
    true,
    580
)

if (selecaoAba.Length = 0) {
    try wb.Close(false)
    try excel.Quit()
    ExitApp
}

nomeAba := selecaoAba[1]
SEMESTRE_ATUAL := nomeAba ; ex: "2026-5", "2027-1", etc.

ws := ObterAba(wb, nomeAba)
if !ws {
    try wb.Close(false)
    try excel.Quit()
    ExibirMensagem("Erro de Aba", "Aba Não Encontrada", "Aba '" nomeAba "' não foi localizada na planilha.", "erro", "Fechar")
    ExitApp
}

; =========================================================
; PASSO 3 — LER CURSOS E SELEÇÃO EM LOTE
; =========================================================

ultimaLinha := ws.Cells(ws.Rows.Count, 2).End(-4162).Row
cursos := []

Loop ultimaLinha - 1 {
    linha := A_Index + 1
    coordenacao := Trim(ws.Cells(linha, 1).Value)
    posgrad     := Trim(ws.Cells(linha, 2).Value)
    codigoTurma := Trim(ws.Cells(linha, 3).Value)
    siglaUnid   := ExtrairSiglaUnidade(codigoTurma)
    unidadeExib := (siglaUnid != "" && siglaUnid != codigoTurma) ? siglaUnid " (" codigoTurma ")" : codigoTurma
    turma       := Trim(ws.Cells(linha, 4).Value)

    if (posgrad = "")
        continue

    cursos.Push(Map(
        "linha", linha,
        "coordenacao", coordenacao,
        "posgrad", posgrad,
        "unidade", codigoTurma,
        "siglaUnidade", siglaUnid,
        "unidadeExib", unidadeExib,
        "turma", turma
    ))
}

if (cursos.Length = 0) {
    try wb.Close(false)
    try excel.Quit()
    ExibirMensagem("Sem Cursos", "Nenhum Curso Encontrado", "Nenhum curso foi encontrado na aba '" nomeAba "'.", "erro", "Fechar")
    ExitApp
}

selecionados := SelecionarCursosUI(
    "Passo 3 — Selecionar Cursos",
    "Cursos Disponíveis na Aba " nomeAba,
    cursos,
    680
)

if (selecionados.Length = 0) {
    try wb.Close(false)
    try excel.Quit()
    ExitApp
}

; Diagnóstico da seleção de cursos
diagSelecao := ""
for idx, c in selecionados {
    coordLista := ExtrairCoordenadores(c["coordenacao"])
    coordTxt := ""
    for i, nomeCoord in coordLista
        coordTxt .= (i = 1 ? "" : " / ") nomeCoord
    if (coordTxt = "")
        coordTxt := "(nenhum coordenador listado)"
    diagSelecao .= idx ". " c["posgrad"]
    if (c["unidade"] != "")
        diagSelecao .= " (" c["unidade"] ")"
    diagSelecao .= "`n     Coordenador(es): " coordTxt "`n`n"
}

confirmSelecao := Confirmar(
    "Passo 3 — Confirmar Lote",
    "Cursos Selecionados (" selecionados.Length ")",
    diagSelecao "Deseja confirmar o processamento destes cursos?",
    "✅ Confirmar e Avançar",
    "❌ Cancelar",
    true,
    660
)

if !confirmSelecao {
    try wb.Close(false)
    try excel.Quit()
    ExitApp
}

; =========================================================
; DATA DE EMISSÃO (ÚNICA PARA O LOTE)
; =========================================================

hoje := FormatTime(A_Now, "dd/MM/yyyy")

respData := Confirmar(
    "Data de Emissão",
    "Data de Emissão dos Certificados",
    "Data sugerida para emissão: " hoje "`n`n"
    "Esta data será impressa em TODOS os certificados gerados nesta rodada.`n`n"
    "Deseja utilizar esta data?",
    "✅ Usar " hoje,
    "✏️ Alterar Data",
    true,
    580
)

if respData {
    dataEmissao := hoje
} else {
    inputData := PedirTexto(
        "Data de Emissão",
        "Informar Data de Emissão",
        "Digite a data de emissão a ser impressa nos certificados:",
        hoje,
        "Formato: DD/MM/AAAA",
        580
    )
    if (inputData["Result"] != "OK" || Trim(inputData["Value"]) = "") {
        try wb.Close(false)
        try excel.Quit()
        ExitApp
    }
    dataEmissao := Trim(inputData["Value"])
}

CriarPastaSegura(PASTA_SAIDA)

try {
    ppt := ComObject("PowerPoint.Application")
    ppt.Visible := true
} catch as err {
    try wb.Close(false)
    try excel.Quit()
    ExibirMensagem("Erro no PowerPoint", "Falha ao Inicializar PowerPoint", "Não foi possível abrir o PowerPoint:`n`n" err.Message, "erro", "Fechar")
    ExitApp
}

; =========================================================
; ================  FASE 1 — CARREGAMENTO  ================
; =========================================================

cacheCoord     := Map()
ultimoTemplate := ""
mesesGlobais   := []

cursoConfigs := []

for numCursoAtual, cursoSelecionado in selecionados {

    try {
        coordLista := ExtrairCoordenadores(cursoSelecionado["coordenacao"])
        assinaturaPath := ""
        coordRaw := ""

        if (coordLista.Length = 1) {
            coordRaw := coordLista[1]
            encontrado := EncontrarAssinatura(PASTA_ASSINATURAS, coordRaw)
            if (encontrado != "")
                assinaturaPath := encontrado
        } else {
            candidatosEncontrados := []
            for nomeCand in coordLista {
                enc := EncontrarAssinatura(PASTA_ASSINATURAS, nomeCand)
                if (enc != "")
                    candidatosEncontrados.Push(Map("nome", nomeCand, "caminho", enc))
            }

            if (candidatosEncontrados.Length = 1) {
                coordRaw       := candidatosEncontrados[1]["nome"]
                assinaturaPath := candidatosEncontrados[1]["caminho"]
            } else if (candidatosEncontrados.Length > 1) {
                nomesRestantes := []
                for cand in candidatosEncontrados
                    nomesRestantes.Push(cand["nome"])

                nomeEscolhido := EscolherCoordenadorUI(
                    nomesRestantes,
                    "[Curso " numCursoAtual "/" selecionados.Length "] " cursoSelecionado["posgrad"]
                )
                if (nomeEscolhido = "")
                    throw Error("Nenhum coordenador selecionado para este curso.")

                for cand in candidatosEncontrados {
                    if (cand["nome"] = nomeEscolhido) {
                        coordRaw       := cand["nome"]
                        assinaturaPath := cand["caminho"]
                        break
                    }
                }
            }
        }

        if (assinaturaPath = "") {
            nomesTexto := ""
            if (coordLista.Length > 1) {
                Loop coordLista.Length {
                    nomesTexto .= coordLista[A_Index]
                    if (A_Index < coordLista.Length - 1)
                        nomesTexto .= ", "
                    else if (A_Index = coordLista.Length - 1)
                        nomesTexto .= " ou "
                }
            } else {
                nomesTexto := coordLista[1]
            }

            respManual := Confirmar(
                "Passo 4 — Assinatura Não Encontrada",
                "Assinatura Não Localizada Automaticamente",
                "[Curso " numCursoAtual "/" selecionados.Length "] " cursoSelecionado["posgrad"] "`n`n"
                "Assinatura não encontrada na pasta " PASTA_ASSINATURAS " para:`n`n  " nomesTexto "`n`n"
                "Deseja selecionar o arquivo de imagem da assinatura manualmente?",
                "📁 Selecionar Imagem",
                "❌ Cancelar",
                true,
                600
            )
            if !respManual
                throw Error("Assinatura não encontrada e seleção manual cancelada.")

            tituloSelecao := (coordLista.Length > 1)
                ? "Selecione a assinatura do coordenador: " nomesTexto
                : "Selecione a assinatura de: " nomesTexto

            assinaturaPath := FileSelect(1, PASTA_ASSINATURAS, tituloSelecao, "*.jpg; *.jpeg; *.png")
            if !assinaturaPath
                throw Error("Nenhuma assinatura selecionada.")

            SplitPath(assinaturaPath, , , , &nomeArquivoSemExt)
            coordRaw := (coordLista.Length > 1) ? nomeArquivoSemExt : coordLista[1]
        }

        usarCache := cacheCoord.Has(coordRaw)
        if usarCache {
            dadosCache := cacheCoord[coordRaw]
            respReusar := Confirmar(
                "Passo 4 — Reaproveitar Coordenador",
                "Coordenador Já Configurado",
                "[Curso " numCursoAtual "/" selecionados.Length "] " cursoSelecionado["posgrad"] "`n`n"
                "Já configuramos este coordenador nesta rodada:`n`n"
                "  " dadosCache["prefixo"] coordRaw "`n  " dadosCache["cargo"] "`n`n"
                "Deseja reaproveitar os mesmos dados para este curso?",
                "✅ Sim, Reaproveitar",
                "✏️ Não, Reconfigurar",
                true,
                580
            )
            usarCache := respReusar
        }

        if usarCache {
            prefixo := cacheCoord[coordRaw]["prefixo"]
            cargo   := cacheCoord[coordRaw]["cargo"]
        } else {
            escolhaTC := EscolherTituloECargoUI(
                coordRaw,
                "[Curso " numCursoAtual "/" selecionados.Length "] " cursoSelecionado["posgrad"]
            )
            if (escolhaTC = "")
                throw Error("Título e cargo do coordenador não confirmados.")

            prefixo := escolhaTC["prefixo"]
            cargo   := escolhaTC["cargo"]
            cacheCoord[coordRaw] := Map("prefixo", prefixo, "cargo", cargo)
        }

        coordUsado := prefixo coordRaw

        ; ─── Localização automática da planilha do curso no OneDrive / SharePoint ───
        cursoPath := LocalizarPlanilhaCursoDinamica(cursoSelecionado["posgrad"], cursoSelecionado["unidade"], SEMESTRE_ATUAL)

        if (cursoPath != "" && FileExist(cursoPath)) {
            SplitPath(cursoPath, &nomeArquivoPlanilha, &dirPlanilha)
            usarAuto := Confirmar(
                "Passo 5 — Planilha do Curso",
                "Planilha Localizada na Nuvem",
                "Curso: " cursoSelecionado["posgrad"] (cursoSelecionado["unidade"] != "" ? " (" cursoSelecionado["unidade"] ")" : "") "`n`n"
                "O sistema localizou automaticamente no OneDrive:`n"
                "📁 " nomeArquivoPlanilha "`n`n"
                "Caminho:`n" cursoPath "`n`n"
                "Deseja utilizar esta planilha?",
                "✅ Usar Esta Planilha",
                "📁 Escolher Outra Manualmente",
                true,
                620
            )
            if !usarAuto {
                pastaInicio := dirPlanilha != "" ? dirPlanilha : (INFO_NUVEM["raizOneDrive"] != "" ? INFO_NUVEM["raizOneDrive"] : "")
                cursoPath := FileSelect(1, pastaInicio, "Passo 5 — Selecione a planilha de: " cursoSelecionado["posgrad"], "*.xlsx")
                if !cursoPath
                    throw Error("Nenhuma planilha de curso selecionada.")
            }
        } else {
            ExibirMensagem(
                "Passo 5 — Planilha do Curso",
                "Selecionar Planilha do Curso",
                "Não foi possível localizar automaticamente a planilha na nuvem.`n`nSelecione a planilha ESPECÍFICA de:`n  " cursoSelecionado["posgrad"]
                (cursoSelecionado["unidade"] != "" ? " (" cursoSelecionado["unidade"] ")" : "")
                "`n`n[Curso " numCursoAtual "/" selecionados.Length "]",
                "passo",
                "Selecionar Planilha ➔",
                580
            )
            pastaInicio := INFO_NUVEM["raizOneDrive"] != "" ? INFO_NUVEM["raizOneDrive"] : ""
            cursoPath := FileSelect(1, pastaInicio, "Passo 5 — Selecione a planilha do curso: " cursoSelecionado["posgrad"], "*.xlsx")
            if !cursoPath
                throw Error("Nenhuma planilha de curso selecionada.")
        }

        ; ─── Template PowerPoint do Certificado (Certificado_Modelo_AHK.pptx) ───
        pptModelo := ""
        if (ultimoTemplate != "") {
            pptModelo := ultimoTemplate
        } else if (INFO_NUVEM["modeloPptx"] != "" && FileExist(INFO_NUVEM["modeloPptx"])) {
            pptModelo := INFO_NUVEM["modeloPptx"]
            ultimoTemplate := pptModelo
        } else {
            ExibirMensagem(
                "Passo 6 — Template PowerPoint",
                "Selecionar Modelo do Certificado",
                "Selecione o arquivo de modelo PowerPoint (.pptx) que será utilizado em TODOS os cursos desta rodada.",
                "passo",
                "Selecionar Modelo ➔",
                580
            )
            pastaInicioPpt := INFO_NUVEM["pastaDeclaracoes"] != "" ? INFO_NUVEM["pastaDeclaracoes"] : ""
            pptModelo := FileSelect(1, pastaInicioPpt, "Template do Certificado (.pptx)", "*.pptx")
            if !pptModelo
                throw Error("Nenhum template selecionado.")
            ultimoTemplate := pptModelo
        }

        wbCurso := AbrirPastaDeTrabalho(excel, cursoPath)

        abasCursoArr := ListarNomesAbas(wbCurso)

        if (mesesGlobais.Length = 0) {
            selecaoAbasCurso := SelecionarOpcoesUI(
                "Passo 7 — Selecionar Período da Rodada",
                "Mês(es) a Processar",
                abasCursoArr,
                false,
                580
            )
            if (selecaoAbasCurso.Length = 0) {
                wbCurso.Close(false)
                throw Error("Nenhum mês selecionado para este curso.")
            }
            mesesGlobais  := selecaoAbasCurso
            arrAbasCurso  := selecaoAbasCurso
        } else {
            arrAbasCurso := []
            for m in mesesGlobais {
                if ObterAba(wbCurso, m)
                    arrAbasCurso.Push(m)
            }
            if (arrAbasCurso.Length = 0) {
                nomesPeriodo := ""
                for i, m in mesesGlobais
                    nomesPeriodo .= (i = 1 ? "" : ", ") m

                selecaoAbasCurso := SelecionarOpcoesUI(
                    "Passo 7 — Selecionar Meses — " cursoSelecionado["posgrad"],
                    "Meses Específicos para este Curso",
                    abasCursoArr,
                    false,
                    580
                )
                if (selecaoAbasCurso.Length = 0) {
                    wbCurso.Close(false)
                    throw Error("Nenhum mês selecionado para este curso.")
                }
                arrAbasCurso := selecaoAbasCurso
            }
        }

        abasInvalidas := ""
        for idx, nomeAbaCurso in arrAbasCurso {
            if !ObterAba(wbCurso, nomeAbaCurso)
                abasInvalidas .= "  • " nomeAbaCurso "`n"
        }
        if (abasInvalidas != "") {
            wbCurso.Close(false)
            throw Error("Abas não encontradas:`n" abasInvalidas)
        }

        wsPrimeiraCurso := ObterAba(wbCurso, arrAbasCurso[1])
        nomePos := Trim(wsPrimeiraCurso.Range("A2").Value)
        if (nomePos = "") {
            wbCurso.Close(false)
            throw Error("Célula A2 da aba '" arrAbasCurso[1] "' está vazia.")
        }

        modeloBase := CriarModeloBase(ppt, pptModelo, coordUsado, cargo, dataEmissao, assinaturaPath, ARQUIVO_LOGO)

        cursoConfigs.Push(Map(
            "posgrad", cursoSelecionado["posgrad"],
            "unidade", cursoSelecionado["unidade"],
            "siglaUnidade", cursoSelecionado["siglaUnidade"],
            "cursoPath", cursoPath,
            "wbCurso", wbCurso,
            "abas", arrAbasCurso,
            "nomePos", nomePos,
            "modeloBase", modeloBase,
            "falhouCarregamento", false
        ))

    } catch as errCarregamento {
        ExibirMensagem(
            "Erro no Carregamento",
            "Curso Removido do Lote",
            "O curso a seguir não pôde ser carregado e foi removido desta rodada:`n`n"
            cursoSelecionado["posgrad"] "`n`nMotivo: " errCarregamento.Message,
            "erro",
            "Continuar"
        )
        continue
    }
}

if (cursoConfigs.Length = 0) {
    try ppt.Quit()
    try wb.Close(false)
    try excel.Quit()
    ExibirMensagem("Encerrando", "Nenhum Curso Carregado", "Nenhum curso foi carregado com sucesso. O programa será encerrado.", "erro", "Fechar")
    ExitApp
}

; =========================================================
; ================  FASE 2 — VALIDAÇÃO  ====================
; =========================================================

todosProblemas      := []
totalMescladasGeral := 0
totalOkGeral         := 0

for cfg in cursoConfigs {
    for abaNome in cfg["abas"] {
        wsAba := ObterAba(cfg["wbCurso"], abaNome)

        colsAba := DetectarColunasPorAba(wsAba)
        colNomeAba  := colsAba["colNome"]
        colCursoAba := colsAba["colCurso"]

        resultado := ValidarAba(wsAba, abaNome, colNomeAba, colCursoAba)
        totalMescladasGeral += resultado["mescladasCount"]
        totalOkGeral += resultado["totalOk"]
        for issue in resultado["issues"] {
            issue["curso"]    := cfg["posgrad"]
            issue["ws"]       := wsAba
            issue["colNome"]  := colNomeAba
            issue["colCurso"] := colCursoAba
            todosProblemas.Push(issue)
        }
    }
}

resumo := "Linhas prontas para geração: " totalOkGeral "`n"
resumo .= "Células mescladas lidas automaticamente: " totalMescladasGeral "`n"
resumo .= "Linhas com inconsistências encontradas: " todosProblemas.Length "`n"

linhasParaPular := Map()

if (todosProblemas.Length = 0) {
    ExibirMensagem(
        "Passo 9 — Validação OK",
        "Validação Concluída com Sucesso",
        "Nenhuma inconsistência foi encontrada nos dados!`n`n" resumo,
        "sucesso",
        "Gerar Certificados ➔",
        580
    )
} else {
    detalheProblemas := ""
    for idx, p in todosProblemas {
        r := p["registro"]
        detalheProblemas .= idx ". " p["curso"] " / " p["aba"] " / linha " p["linha"] ":`n"
        if r["erroNome"]
            detalheProblemas .= "    - Nome em branco`n"
        if r["erroCurso"]
            detalheProblemas .= "    - Curso/Disciplina em branco`n"
        if r["erroData"]
            detalheProblemas .= "    - Data ausente ou inválida`n"
        if r["erroHoras"]
            detalheProblemas .= "    - Horário não reconhecido (lido: " r["horarioTxt"] ")`n"
        if r["suspeitaNomeEmail"]
            detalheProblemas .= "    - AVISO: coluna de nome parece conter e-mail`n"
        if r["suspeitaCursoData"]
            detalheProblemas .= "    - AVISO: coluna de curso parece conter data/hora`n"
        if r["suspeitaNomeComExtra"]
            detalheProblemas .= "    - AVISO: nome limpo automaticamente — original: `"" r["nomeOriginalBruto"] "`"  →  usado: `"" r["nome"] "`"`n"
        detalheProblemas .= "`n"
    }

    escolhaValidacao := TelaProblemasUI(resumo, Trim(detalheProblemas, "`n"))

    if (escolhaValidacao = "cancelar") {
        for cfg in cursoConfigs {
            try cfg["wbCurso"].Close(false)
            try FileDelete(cfg["modeloBase"])
        }
        try ppt.Quit()
        try wb.Close(false)
        try excel.Quit()
        ExitApp
    }

    if (escolhaValidacao = "corrigir") {
        for p in todosProblemas {
            r := p["registro"]
            wsX := p["ws"]
            linha := p["linha"]

            listaErros := ""
            if r["erroNome"]
                listaErros .= "  • Nome do docente`n"
            if r["erroCurso"]
                listaErros .= "  • Curso / disciplina`n"
            if r["erroData"]
                listaErros .= "  • Data da aula (formato DD/MM/AAAA)`n"
            if r["erroHoras"]
                listaErros .= "  • Horário (lido: " r["horarioTxt"] ") — formato HH:MM-HH:MM`n"
            if r["suspeitaNomeEmail"]
                listaErros .= "  • (aviso) coluna de nome parece ter e-mail`n"
            if r["suspeitaCursoData"]
                listaErros .= "  • (aviso) coluna de curso parece ter data/hora`n"
            if r["suspeitaNomeComExtra"]
                listaErros .= "  • (aviso) nome limpo automaticamente — original: `"" r["nomeOriginalBruto"] "`"  →  usado: `"" r["nome"] "`"`n"

            resp := Confirmar(
                "Corrigir Linha " linha,
                "Inconsistência na Linha " linha,
                "CURSO: " p["curso"] "`nABA: " p["aba"] "`nLINHA: " linha "`n`n"
                "Campos com inconsistência:`n" listaErros "`n"
                "Deseja corrigir esta linha agora?",
                "🔧 Corrigir",
                "⏭️ Pular Linha",
                true,
                580
            )

            if !resp {
                linhasParaPular[p["aba"] "|" linha] := true
                continue
            }

            if r["erroNome"] {
                cx := PedirTexto(
                    "Corrigir Nome — Linha " linha,
                    "Nome do Docente",
                    "CURSO: " p["curso"] "`nABA: " p["aba"] " — LINHA: " linha "`n`n"
                    "Valor atual: '" r["nome"] "'`n`n"
                    "Digite o nome completo e correto do docente:",
                    r["nome"],
                    "Exemplo: Maria da Silva Santos"
                )
                if (cx["Result"] = "OK" && Trim(cx["Value"]) != "")
                    wsX.Range(p["colNome"] linha).Value := Trim(cx["Value"])
            }
            if r["erroCurso"] {
                cx := PedirTexto(
                    "Corrigir Disciplina — Linha " linha,
                    "Curso / Disciplina",
                    "CURSO: " p["curso"] "`nABA: " p["aba"] " — LINHA: " linha "`n`n"
                    "Valor atual: '" r["curso"] "'`n`n"
                    "Digite o nome correto da disciplina/aula:",
                    r["curso"],
                    "Exemplo: Gestão Estratégica de Pessoas"
                )
                if (cx["Result"] = "OK" && Trim(cx["Value"]) != "")
                    wsX.Range(p["colCurso"] linha).Value := Trim(cx["Value"])
            }
            if r["erroData"] {
                cx := PedirTexto(
                    "Corrigir Data — Linha " linha,
                    "Data da Aula",
                    "CURSO: " p["curso"] "`nABA: " p["aba"] " — LINHA: " linha "`n`n"
                    "Valor atual: '" r["dataAula"] "'`n`n"
                    "Digite a data correta no formato DD/MM/AAAA:",
                    r["dataAula"],
                    "Exemplo: 15/05/2026"
                )
                if (cx["Result"] = "OK" && Trim(cx["Value"]) != "")
                    wsX.Cells(linha, 1).Value := Trim(cx["Value"])
            }
            if r["erroHoras"] {
                cx := PedirTexto(
                    "Corrigir Carga Horária — Linha " linha,
                    "Carga Horária da Aula",
                    "CURSO: " p["curso"] "`nABA: " p["aba"] " — LINHA: " linha "`n`n"
                    "Valor atual: '" r["horarioTxt"] "' (não reconhecido)`n`n"
                    "Digite a duração da aula:`n"
                    "  • Apenas horas inteiras → escreva o número  (ex: 4)`n"
                    "  • Horas e minutos → use H:MM              (ex: 2:30)`n`n"
                    "O sistema converterá automaticamente para o texto do certificado.",
                    "",
                    "Ex: 4   ou   2:30   ou   1:30"
                )
                if (cx["Result"] = "OK" && Trim(cx["Value"]) != "") {
                    duracaoFormatada := FormatarDuracaoDigitada(Trim(cx["Value"]))
                    if (duracaoFormatada != "")
                        wsX.Cells(linha, 2).Value := duracaoFormatada
                }
            }
        }
    } else {
        for p in todosProblemas
            linhasParaPular[p["aba"] "|" p["linha"]] := true
    }
}

; =========================================================
; ================  FASE 3 — GERAÇÃO  ======================
; =========================================================

relatorioPorCurso   := []
totalGeralGerados   := 0
totalGeralErros     := 0
totalCursosComFalha := 0

for cfg in cursoConfigs {
    try {
        resultadoCurso := GerarCertificadosCurso(cfg, ppt, PASTA_SAIDA, linhasParaPular)
        relatorioPorCurso.Push(Map(
            "posgrad", cfg["posgrad"], "unidade", cfg["unidade"],
            "gerados", resultadoCurso["gerados"], "erros", resultadoCurso["erros"], "falhou", false
        ))
        totalGeralGerados += resultadoCurso["gerados"]
        totalGeralErros   += resultadoCurso["erros"]
    } catch as errCurso {
        totalCursosComFalha++
        relatorioPorCurso.Push(Map(
            "posgrad", cfg["posgrad"], "unidade", cfg["unidade"],
            "gerados", 0, "erros", 0, "falhou", true, "motivo", errCurso.Message
        ))
    }
    try cfg["wbCurso"].Close(false)
    try FileDelete(cfg["modeloBase"])
}

try ppt.Quit()
try wb.Close(false)
try excel.Quit()

; =========================================================
; RELATÓRIO FINAL DO LOTE
; =========================================================

detalheRelatorio := ""
for idx, r in relatorioPorCurso {
    detalheRelatorio .= idx ". " r["posgrad"]
    if (r["unidade"] != "")
        detalheRelatorio .= " (" r["unidade"] ")"
    if (r["falhou"])
        detalheRelatorio .= "`n     ⚠️ FALHOU — " r["motivo"] "`n`n"
    else
        detalheRelatorio .= "`n     Gerados: " r["gerados"] " | Linhas ignoradas: " r["erros"] "`n`n"
}

resumoFinal := "Cursos processados no lote: " cursoConfigs.Length "`n"
resumoFinal .= "Data de emissão impressa:   " dataEmissao "`n"
resumoFinal .= "Total de PDFs gerados:      " totalGeralGerados "`n"
resumoFinal .= "Cursos que falharam:        " totalCursosComFalha "`n"
resumoFinal .= "Células mescladas tratadas: " totalMescladasGeral

ExibirRelatorioFinalUI(
    "Passo 10 — Relatório Final do Lote",
    "Geração de Certificados Concluída!",
    resumoFinal,
    Trim(detalheRelatorio, "`n"),
    PASTA_SAIDA,
    700
)

ExitApp

; =========================================================
; IDENTIFICAÇÃO AUTOMÁTICA DE COLUNAS NA LINHA 3
; =========================================================

IdentificarColunaDocentes(wsAba) {
    maxCols := 50
    try {
        usedCols := wsAba.UsedRange.Columns.Count
        if (usedCols > maxCols)
            maxCols := usedCols
    }

    Loop maxCols {
        colIdx := A_Index
        celVal := ""
        try celVal := String(wsAba.Cells(3, colIdx).Value)
        if (celVal = "")
            try celVal := String(wsAba.Cells(3, colIdx).Text)

        celValNorm := NormalizarTexto(celVal)
        if InStr(celValNorm, "docente") || InStr(celValNorm, "professor")
            return IndiceParaColuna(colIdx)
    }
    return ""
}

IdentificarColunaAulas(wsAba) {
    maxCols := 50
    try {
        usedCols := wsAba.UsedRange.Columns.Count
        if (usedCols > maxCols)
            maxCols := usedCols
    }

    termos := ["disciplina", "aula", "modulo", "conteudo", "tema", "materia", "topico"]
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
    return ""
}

; =========================================================
; FUNÇÃO: Detectar colunas de Docente e Aula para uma aba
; Chamada individualmente para cada aba — garante que cada
; aba usa as suas próprias colunas, mesmo que variem.
; =========================================================

DetectarColunasPorAba(wsAba) {
    colNome  := IdentificarColunaDocentes(wsAba)
    colCurso := IdentificarColunaAulas(wsAba)
    colEmail := IdentificarColunaEmails(wsAba)

    if (colNome = "")
        colNome := "E"
    if (colCurso = "")
        colCurso := "D"
    if (colEmail = "")
        colEmail := "F"

    return Map("colNome", colNome, "colCurso", colCurso, "colEmail", colEmail)
}

; =========================================================
; FUNÇÃO: Identificar coluna de e-mails na Linha 3 (mesmo padrão
; usado por IdentificarColunaDocentes / IdentificarColunaAulas acima)
; =========================================================

IdentificarColunaEmails(wsAba) {
    maxCols := 50
    try {
        usedCols := wsAba.UsedRange.Columns.Count
        if (usedCols > maxCols)
            maxCols := usedCols
    }

    termos := ["email", "e-mail", "contato", "correio", "eletronico"]
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
    return ""
}

; =========================================================
; FUNÇÕES: Validação e extração de e-mail (portadas do EnviarEmails_v11.ahk
; para permitir a montagem do manifesto sem depender de outro script)
; =========================================================

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

; =========================================================
; FUNÇÃO: Abrir uma pasta de trabalho suprimindo os diálogos que travam
; o Excel quando ele está invisível (atualizar links, recomendação de
; somente leitura). Sem isso o Excel fica esperando um clique numa janela
; que ninguém vê, e a próxima chamada COM falha com 0x80010001.
; =========================================================

AbrirPastaDeTrabalho(excel, caminho) {
    ; Parâmetros do Workbooks.Open: (Filename, UpdateLinks, ReadOnly, Format,
    ; Password, WriteResPassword, IgnoreReadOnlyRecommended, ...)
    return excel.Workbooks.Open(caminho, 0, false, , , , true)
}

; =========================================================
; FUNÇÃO: Listar os nomes das abas de uma pasta de trabalho, com repetição
; automática. Logo depois de abrir o arquivo, o Excel às vezes ainda está
; processando internamente (mesmo já tendo retornado o Workbook pro
; AutoHotkey) e rejeita a PRÓXIMA chamada COM com 0x80010001
; (RPC_E_CALL_REJECTED) — um erro intermitente, não uma falha de lógica:
; o mesmo código pode funcionar numa hora e falhar noutra, dependendo de
; quão ocupado o Excel está naquele instante. Em vez de desistir na
; primeira falha, tenta de novo algumas vezes com uma pequena espera.
; =========================================================

ListarNomesAbas(wbX, tentativas := 5, esperaMs := 500) {
    Loop tentativas {
        try {
            nomes := []
            Loop wbX.Sheets.Count
                nomes.Push(wbX.Sheets.Item(A_Index).Name)
            return nomes
        } catch as errTentativa {
            if (A_Index = tentativas)
                throw errTentativa
            Sleep esperaMs
        }
    }
}

; =========================================================
; FUNÇÃO: Obter aba pelo nome
; =========================================================

ObterAba(wbX, nomeAba) {
    Loop wbX.Sheets.Count {
        if (Trim(wbX.Sheets.Item(A_Index).Name) = nomeAba)
            return wbX.Sheets.Item(A_Index)
    }
    return ""
}

; =========================================================
; FUNÇÃO: Extrair lista de coordenadores
; =========================================================

ExtrairCoordenadores(rawCoord) {
    arrCoord := StrSplit(rawCoord, [",", "/", "\"])
    coordLista := []
    for idx, c in arrCoord {
        c := Trim(c)
        if (c != "")
            coordLista.Push(c)
    }
    return coordLista
}

; =========================================================
; FUNÇÃO: Normalizar nome (espaços invisíveis, maiúsculas)
; =========================================================

NormalizarNome(nome) {
    nome := Trim(nome)
    nome := RegExReplace(nome, "[\x{00A0}\t]", " ")
    nome := RegExReplace(nome, "\s+", " ")
    return StrLower(nome)
}

; =========================================================
; FUNÇÃO: Encontrar assinatura de forma tolerante
; =========================================================

EncontrarAssinatura(pastaAssinaturas, nomeCoord) {
    for ext in ["jpg", "jpeg", "png", "JPG", "JPEG", "PNG"] {
        caminho := pastaAssinaturas nomeCoord "." ext
        if FileExist(caminho)
            return caminho
    }

    nomeAlvo := NormalizarNome(nomeCoord)
    try {
        Loop Files, pastaAssinaturas "*.*" {
            nomeSemExt := SubStr(A_LoopFileName, 1, StrLen(A_LoopFileName) - StrLen(A_LoopFileExt) - 1)
            if (NormalizarNome(nomeSemExt) = nomeAlvo)
                return A_LoopFileFullPath
        }
    } catch {
    }

    return ""
}

; =========================================================
; FUNÇÃO: Sanitizar nome para uso como pasta/arquivo
; =========================================================

SanitizarNomeArquivo(nome) {
    limpo := RegExReplace(nome, '[\\/:*?"<>|]', "")
    limpo := RegExReplace(limpo, "[\x00-\x1F]", "")
    limpo := Trim(limpo, " `t`r`n.")
    if (limpo = "")
        limpo := "Sem_Nome"
    return limpo
}

; =========================================================
; FUNÇÃO: Converter nome de mês/ano (ex.: "Abril26", "Abril/2026")
; para o formato de pasta "MM.AA" (ex.: "04.26").
; =========================================================

ConverterMesParaPasta(abaNome) {
    nome := Trim(abaNome)

    if !RegExMatch(nome, "^([A-Za-zçÇãÃéÉêÊóÓ]+)\D*(\d{2,4})\s*$", &m)
        return SanitizarNomeArquivo(nome)

    mesTexto := StrLower(m[1])
    mesTexto := StrReplace(mesTexto, "ç", "c")
    mesTexto := StrReplace(mesTexto, "ã", "a")
    mesTexto := StrReplace(mesTexto, "é", "e")
    mesTexto := StrReplace(mesTexto, "ê", "e")
    mesTexto := StrReplace(mesTexto, "ó", "o")

    mapaMeses := Map(
        "janeiro", "01", "jan", "01",
        "fevereiro", "02", "fev", "02",
        "marco", "03", "mar", "03",
        "abril", "04", "abr", "04",
        "maio", "05", "mai", "05",
        "junho", "06", "jun", "06",
        "julho", "07", "jul", "07",
        "agosto", "08", "ago", "08",
        "setembro", "09", "set", "09",
        "outubro", "10", "out", "10",
        "novembro", "11", "nov", "11",
        "dezembro", "12", "dez", "12"
    )

    if !mapaMeses.Has(mesTexto)
        return SanitizarNomeArquivo(nome)

    numMes   := mapaMeses[mesTexto]
    anoTexto := m[2]
    anoCurto := (StrLen(anoTexto) = 4) ? SubStr(anoTexto, 3, 2) : anoTexto

    return numMes "." anoCurto
}

; =========================================================
; FUNÇÃO: Criar pasta com diagnóstico claro em caso de falha
; =========================================================

CriarPastaSegura(caminho) {
    try {
        DirCreate(caminho)
    } catch as errDir {
        throw Error("Falha ao criar a pasta:`n" caminho "`n`n" errDir.Message)
    }
}

; =========================================================
; FUNÇÃO: Validar se um texto é uma data real
; =========================================================

EhDataValida(str) {
    str := Trim(str)
    if (str = "")
        return false
    return RegExMatch(str, "^\d{1,2}/\d{1,2}/\d{2,4}$") ? true : false
}

; =========================================================
; FUNÇÃO: Ler célula resolvendo mescla vertical com segurança
; =========================================================

ResolverCelulaMesclada(cel) {
    try {
        if cel.MergeCells
            return cel.MergeArea.Cells(1, 1)
    }
    return cel
}

; =========================================================
; FUNÇÃO: Converter serial de data do Excel para DD/MM/AAAA
; =========================================================

ExcelSerialParaData(valorSerial) {
    if (valorSerial = "" || !IsNumber(valorSerial) || valorSerial < 1)
        return ""
    try {
        serial := Integer(valorSerial)
        dtExcel := DateAdd("19000101000000", serial - 1, "Days")
        return FormatTime(dtExcel, "dd/MM/yyyy")
    }
    return ""
}

; =========================================================
; FUNÇÃO: Ler data de uma célula de forma robusta
; =========================================================

LerDataCelula(cel) {
    txt := ""
    try txt := Trim(String(cel.Text))
    if EhDataValida(txt)
        return txt

    try {
        v2 := cel.Value2
        if (v2 != "" && IsNumber(v2) && v2 > 1) {
            dtConv := ExcelSerialParaData(v2)
            if EhDataValida(dtConv)
                return dtConv
        }
    }

    try {
        vStr := Trim(String(cel.Value))
        if EhDataValida(vStr)
            return vStr
        if RegExMatch(vStr, "\d{1,2}/\d{1,2}/\d{2,4}", &m)
            return m[]
    }

    return ""
}

; =========================================================
; FUNÇÃO: Ler um registro (linha) de forma robusta
; =========================================================

LerRegistro(wsCurso, linha, colNome, colCurso, colEmail := "") {

    celData    := ResolverCelulaMesclada(wsCurso.Cells(linha, 1))
    celHorario := ResolverCelulaMesclada(wsCurso.Cells(linha, 2))
    celNome    := ResolverCelulaMesclada(wsCurso.Range(colNome linha))
    celCurso   := ResolverCelulaMesclada(wsCurso.Range(colCurso linha))

    mesclado := false
    try mesclado := wsCurso.Cells(linha, 1).MergeCells
        || wsCurso.Cells(linha, 2).MergeCells
        || wsCurso.Range(colNome linha).MergeCells
        || wsCurso.Range(colCurso linha).MergeCells

    dataAulaBruta := LerDataCelula(celData)
    dataAula := EhDataValida(dataAulaBruta) ? dataAulaBruta : ""

    nome := Trim(String(celNome.Value))
    nome := RegExReplace(nome, "\s*\b[HMhm]\b\s*$")
    nome := Trim(nome)

    nomeOriginalBruto := nome

    nome := RegExReplace(nome, "\s*\([^)]*\)", "")
    nome := RegExReplace(nome, "i)^(dra|dr|profa|professora|professor|prof|sra|sr)\.?\s+", "")
    nome := Trim(RegExReplace(nome, "\s+", " "))

    suspeitaNomeComExtra := (nome != Trim(RegExReplace(nomeOriginalBruto, "\s+", " ")))

    curso := Trim(String(celCurso.Value))

    horarioTxt := String(celHorario.Text)
    if (horarioTxt = "" || RegExMatch(horarioTxt, "^\d[\.,]\d")) {
        horarioVal := celHorario.Value
        horasFormatadas := FormatarHorasDeValorExcel(horarioVal)
    } else {
        horasFormatadas := FormatarHorasSeguro(horarioTxt)
    }

    erroNome  := (nome = "")
    erroCurso := (curso = "")
    erroData  := (dataAula = "")
    erroHoras := InStr(horasFormatadas, "invalido") ? true : false

    camposPreenchidos := 0
    if !erroNome
        camposPreenchidos++
    if !erroCurso
        camposPreenchidos++
    if !erroData
        camposPreenchidos++
    if (horarioTxt != "")
        camposPreenchidos++

    nomeNorm    := NormalizarTexto(nome)
    cursoNorm   := NormalizarTexto(curso)
    horarioNorm := NormalizarTexto(horarioTxt)

    termosIgnorar := ["almoco", "intervalo", "coffee break", "coffeebreak", "lanche", "refeicao", "pausa", "descanso"]
    ehIntervaloOuInutil := false
    for t in termosIgnorar {
        if (InStr(nomeNorm, t) || InStr(cursoNorm, t) || InStr(horarioNorm, t)) {
            ehIntervaloOuInutil := true
            break
        }
    }

    ehRodapeIgnoravel := (camposPreenchidos <= 1) || (nome = "" && curso = "") || ehIntervaloOuInutil

    suspeitaNomeEmail := InStr(nome, "@") ? true : false
    suspeitaCursoData := (curso != "" && RegExMatch(curso, "^\d{1,2}[/:h]\d")) ? true : false

    ; Leitura do e-mail do docente (usada só na Fase 3, para montar o manifesto).
    ; Se colEmail não for informado (ex.: chamada da Fase 2/validação), fica em branco.
    email := ""
    emailValido := false
    if (colEmail != "") {
        emailBruto := ""
        try emailBruto := Trim(String(wsCurso.Range(colEmail linha).Value))
        email := ExtrairEmailValido(emailBruto)
        emailValido := EmailEhValido(email)
    }

    return Map(
        "linha", linha, "nome", nome, "curso", curso, "dataAula", dataAula,
        "horarioTxt", horarioTxt, "horasFormatadas", horasFormatadas,
        "erroNome", erroNome, "erroCurso", erroCurso, "erroData", erroData, "erroHoras", erroHoras,
        "ehRodapeIgnoravel", ehRodapeIgnoravel, "mesclado", mesclado,
        "suspeitaNomeEmail", suspeitaNomeEmail, "suspeitaCursoData", suspeitaCursoData,
        "suspeitaNomeComExtra", suspeitaNomeComExtra, "nomeOriginalBruto", nomeOriginalBruto,
        "email", email, "emailValido", emailValido
    )
}

; =========================================================
; FUNÇÃO: Validar uma aba inteira (Fase 2)
; =========================================================

ValidarAba(wsCurso, abaNome, colNome, colCurso) {
    ultimaLinhaCurso := wsCurso.Cells(wsCurso.Rows.Count, 1).End(-4162).Row
    issues := []
    mescladasCount := 0
    totalOk := 0

    Loop ultimaLinhaCurso - 4 {
        linha := A_Index + 4
        reg := LerRegistro(wsCurso, linha, colNome, colCurso)

        if reg["ehRodapeIgnoravel"]
            continue

        if reg["mesclado"]
            mescladasCount++

        temProblema := reg["erroNome"] || reg["erroCurso"] || reg["erroData"] || reg["erroHoras"] || reg["suspeitaNomeEmail"] || reg["suspeitaCursoData"] || reg["suspeitaNomeComExtra"]

        if !temProblema {
            totalOk++
            continue
        }

        issues.Push(Map("aba", abaNome, "linha", linha, "registro", reg))
    }

    return Map("issues", issues, "mescladasCount", mescladasCount, "totalOk", totalOk)
}

; =========================================================
; FUNÇÃO: Gerar certificados de um curso (Fase 3 — Salvamento Duplo)
; =========================================================

GerarCertificadosCurso(cfg, ppt, pastaSaidaBase, linhasParaPular) {
    totalGerados := 0
    totalErros := 0

    posgradSanit := SanitizarNomeArquivo(cfg["posgrad"])
    nomeCursoComUnidade := (cfg["unidade"] != "") ? posgradSanit " (" SanitizarNomeArquivo(cfg["unidade"]) ")" : posgradSanit

    for abaNome in cfg["abas"] {
        mesPasta := ConverterMesParaPasta(abaNome)

        pastas := GarantirPastasSaida(nomeCursoComUnidade, mesPasta, SEMESTRE_ATUAL)
        pastaNuvemMes := pastas["nuvem"]
        pastaLocalMes := pastas["local"]
        pastaDocentesRaiz := pastas["localDocentes"]
        pastaParaEnviarMes := GarantirPastaParaEnviar(nomeCursoComUnidade, mesPasta)

        wsAba := ObterAba(cfg["wbCurso"], abaNome)
        ultimaLinhaCurso := wsAba.Cells(wsAba.Rows.Count, 1).End(-4162).Row

        colsAba := DetectarColunasPorAba(wsAba)
        colNomeAba  := colsAba["colNome"]
        colCursoAba := colsAba["colCurso"]
        colEmailAba := colsAba["colEmail"]

        ; Agrupamento por docente para o manifesto do Power Automate (Fase 3).
        manifestoDocentes := Map()
        ordemManifesto := []

        Loop ultimaLinhaCurso - 4 {
            linha := A_Index + 4

            if linhasParaPular.Has(abaNome "|" linha)
                continue

            reg := LerRegistro(wsAba, linha, colNomeAba, colCursoAba, colEmailAba)

            if reg["ehRodapeIgnoravel"]
                continue

            if (reg["erroNome"] || reg["erroCurso"] || reg["erroData"] || reg["erroHoras"]) {
                totalErros++
                continue
            }

            pastaDocente := pastaDocentesRaiz SanitizarNomeArquivo(reg["nome"]) "\"
            CriarPastaSegura(pastaDocente)

            registro := Map(
                "aba", abaNome, "linha", linha, "nome", reg["nome"], "curso", reg["curso"],
                "dataAula", reg["dataAula"], "horasFormatadas", reg["horasFormatadas"]
            )

            caminhoGerado := GerarCertificado(registro, ppt, cfg["modeloBase"], pastaNuvemMes, pastaLocalMes, pastaDocente, pastaParaEnviarMes, cfg["nomePos"])
            if (caminhoGerado != "") {
                totalGerados++

                SplitPath(caminhoGerado, &nomeArquivoGerado)
                chaveDoc := StrLower(SanitizarNomeArquivo(reg["nome"]))

                if !manifestoDocentes.Has(chaveDoc) {
                    manifestoDocentes[chaveDoc] := Map(
                        "nome", reg["nome"], "email", reg["email"], "emailValido", reg["emailValido"], "arquivos", []
                    )
                    ordemManifesto.Push(chaveDoc)
                } else if (!manifestoDocentes[chaveDoc]["emailValido"] && reg["emailValido"]) {
                    ; Docente já apareceu antes nesta aba sem e-mail válido — aproveita
                    ; o e-mail válido encontrado agora numa linha posterior do mesmo docente.
                    manifestoDocentes[chaveDoc]["email"] := reg["email"]
                    manifestoDocentes[chaveDoc]["emailValido"] := true
                }

                chaveLinha := reg["dataAula"] "|" reg["horarioTxt"]
                manifestoDocentes[chaveDoc]["arquivos"].Push(Map("nomeArquivo", nomeArquivoGerado, "chaveLinha", chaveLinha))
            } else {
                totalErros++
            }
        }

        GerarManifestoDoMes(pastaParaEnviarMes, manifestoDocentes, ordemManifesto, cfg, abaNome)
    }

    return Map("gerados", totalGerados, "erros", totalErros)
}

; =========================================================
; FUNÇÃO: Gerar o manifesto.txt (conteúdo JSON) do mês, se houver
; ao menos um docente com e-mail válido. Regra de negócio: se
; NENHUM certificado do lote tiver e-mail válido, o manifesto NÃO
; é gerado (a pasta fica só com os PDFs).
; =========================================================

GerarManifestoDoMes(pastaParaEnviarMes, manifestoDocentes, ordemManifesto, cfg, abaNome) {
    if (pastaParaEnviarMes = "")
        return

    listaDocentesValidos := []
    for chaveDoc in ordemManifesto {
        doc := manifestoDocentes[chaveDoc]
        if doc["emailValido"]
            listaDocentesValidos.Push(doc)
    }

    if (listaDocentesValidos.Length = 0)
        return

    loteId := FormatTime(A_Now, "yyyyMMddHHmmss") "-" Random(1000, 9999)
    dataGeracao := FormatTime(A_Now, "yyyy-MM-ddTHH:mm:ss")
    localSigla := (cfg["siglaUnidade"] != "") ? cfg["siglaUnidade"] : cfg["unidade"]

    jsonTexto := ConstruirManifestoJson(
        loteId, dataGeracao, cfg["posgrad"], localSigla,
        pastaParaEnviarMes, cfg["cursoPath"], abaNome, listaDocentesValidos
    )

    try {
        EscreverArquivoAtomico(pastaParaEnviarMes "manifesto.txt", jsonTexto)
    } catch as errManifesto {
        ExibirMensagem(
            "Aviso — Manifesto",
            "Falha ao Gravar Manifesto",
            "Os certificados foram gerados normalmente, mas houve falha ao gravar o manifesto.txt em:`n`n"
            pastaParaEnviarMes "`n`nMotivo: " errManifesto.Message,
            "erro", "Continuar"
        )
    }
}

; =========================================================
; FUNÇÃO: Montar o texto JSON do manifesto (schema fixo — não mudar
; os nomes de campo, o flow do Power Automate já está validado neles)
; =========================================================

ConstruirManifestoJson(loteId, dataGeracao, curso, localUnidade, caminhoPastaPDF, caminhoArquivoExcel, nomeAba, listaDocentes) {
    json := "{`n"
    json .= '  "versaoManifesto": 1,' "`n"
    json .= '  "loteId": "' JsonEscape(loteId) '",' "`n"
    json .= '  "dataGeracao": "' JsonEscape(dataGeracao) '",' "`n"
    json .= '  "curso": "' JsonEscape(curso) '",' "`n"
    json .= '  "local": "' JsonEscape(localUnidade) '",' "`n"
    json .= '  "caminhoPastaPDF": "' JsonEscape(caminhoPastaPDF) '",' "`n"
    json .= '  "caminhoArquivoExcel": "' JsonEscape(caminhoArquivoExcel) '",' "`n"
    json .= '  "nomeAba": "' JsonEscape(nomeAba) '",' "`n"
    json .= '  "certificados": [' "`n"

    for idx, doc in listaDocentes {
        json .= "    {`n"
        json .= '      "docente": "' JsonEscape(doc["nome"]) '",' "`n"
        json .= '      "email": "' JsonEscape(doc["email"]) '",' "`n"
        json .= '      "arquivos": [' "`n"

        for idxA, arq in doc["arquivos"] {
            json .= '        { "nomeArquivo": "' JsonEscape(arq["nomeArquivo"]) '", "chaveLinha": "' JsonEscape(arq["chaveLinha"]) '" }'
            json .= (idxA < doc["arquivos"].Length) ? ",`n" : "`n"
        }

        json .= "      ]`n"
        json .= "    }"
        json .= (idx < listaDocentes.Length) ? ",`n" : "`n"
    }

    json .= "  ]`n"
    json .= "}"
    return json
}

JsonEscape(valor) {
    valor := String(valor)
    valor := StrReplace(valor, "\", "\\")
    valor := StrReplace(valor, '"', '\"')
    valor := StrReplace(valor, "`r`n", "\n")
    valor := StrReplace(valor, "`n", "\n")
    valor := StrReplace(valor, "`r", "\n")
    valor := StrReplace(valor, "`t", "\t")
    return valor
}

; Escrita atômica: grava em .tmp e só então renomeia para o nome final,
; para o Power Automate nunca ler um arquivo pela metade.
EscreverArquivoAtomico(caminhoFinal, conteudo) {
    caminhoTmp := caminhoFinal ".tmp"
    arq := ""
    try {
        arq := FileOpen(caminhoTmp, "w", "UTF-8-RAW")
        if !IsObject(arq)
            throw Error("Não foi possível criar o arquivo temporário: " caminhoTmp)
        arq.Write(conteudo)
        arq.Close()
        FileMove(caminhoTmp, caminhoFinal, 1)
    } catch as err {
        try arq.Close()
        try FileDelete(caminhoTmp)
        throw err
    }
}

; =========================================================
; FUNÇÃO: Substituir placeholder preservando formatação
; =========================================================

SubstituirPlaceholder(shape, placeholder, valor, forcaSize := 0, forcaBold := -1) {
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
            tr.Characters(pos, StrLen(valor)).Font.Color := fColor
            tr.Characters(pos, StrLen(valor)).Font.Size  := (forcaSize > 0) ? forcaSize : fSize
            tr.Characters(pos, StrLen(valor)).Font.Bold  := (forcaBold >= 0) ? forcaBold : fBold
        }
    } catch {
    }
}

; =========================================================
; FUNÇÃO: Criar o modelo-base de um curso
; =========================================================

CriarModeloBase(ppt, pptModelo, coordUsado, cargo, dataEmissao, assinaturaPath, ARQUIVO_LOGO) {

    modeloBase := A_Temp "\certificado_base_" A_TickCount "_" Random(1000, 9999) ".pptx"
    presBase  := ppt.Presentations.Open(pptModelo)
    slideBase := presBase.Slides(1)

    Loop slideBase.Shapes.Count {
        shape := slideBase.Shapes.Item(A_Index)
        try {
            txt := shape.TextFrame.TextRange.Text
            if InStr(txt, "{COORDENADOR}") || InStr(txt, "{CARGO}") || InStr(txt, "{DATAEMISSAO}") {
                SubstituirPlaceholder(shape, "{COORDENADOR}", coordUsado, 10, 0)
                SubstituirPlaceholder(shape, "{CARGO}",       cargo,      10, 1)
                SubstituirPlaceholder(shape, "{DATAEMISSAO}", dataEmissao)
            }
            if InStr(txt, "{COORDENADOR}") || InStr(txt, "{CARGO}") {
                try {
                    shape.TextFrame.WordWrap  := 0
                    shape.TextFrame2.AutoSize := 1
                } catch {
                }
            }
        } catch {
        }
    }

    todasImagens := []
    Loop slideBase.Shapes.Count {
        s := slideBase.Shapes.Item(A_Index)
        try {
            tipo := s.Type
            if (tipo = 13 || tipo = 7 || tipo = 3)
                todasImagens.Push(Map("shape", s, "nome", s.Name, "tipo", tipo, "left", s.Left, "top", s.Top, "w", s.Width, "h", s.Height))
        } catch {
        }
    }

    if (todasImagens.Length = 0) {
        Loop slideBase.Shapes.Count {
            s := slideBase.Shapes.Item(A_Index)
            try {
                temTexto := false
                try temTexto := (Trim(s.TextFrame.TextRange.Text) != "")
                if !temTexto
                    todasImagens.Push(Map("shape", s, "nome", s.Name, "tipo", s.Type, "left", s.Left, "top", s.Top, "w", s.Width, "h", s.Height))
            } catch {
            }
        }
    }

    imgLogo := ""
    imgAssina := ""
    if (todasImagens.Length = 1) {
        imgAssina := todasImagens[1]
    } else if (todasImagens.Length >= 2) {
        menorTop := 999999
        maiorTop := -1
        for i, img in todasImagens {
            if (img["top"] < menorTop) {
                menorTop := img["top"]
                imgLogo := img
            }
            if (img["top"] > maiorTop) {
                maiorTop := img["top"]
                imgAssina := img
            }
        }
    }

    logoTop := 0
    if imgLogo {
        logoTop := imgLogo["top"]
        imgLogo["shape"].Delete()
    }

    larguraPadrao := 130
    alturaPadrao := 50

    if imgAssina {
        posLeft := imgAssina["left"]
        posTop := imgAssina["top"]
        imgAssina["shape"].Delete()

        coordLeft := ""
        coordTop := ""
        Loop slideBase.Shapes.Count {
            sh := slideBase.Shapes.Item(A_Index)
            try {
                txt := sh.TextFrame.TextRange.Text
                if InStr(txt, coordUsado) {
                    coordLeft := sh.Left
                    coordTop := sh.Top
                    break
                }
            } catch {
            }
        }

        if (coordLeft != "") {
            margem := 8
            novoLeft := coordLeft
            novoTop := coordTop - alturaPadrao - margem
        } else {
            novoLeft := posLeft
            novoTop := posTop
        }

        picAssinatura := slideBase.Shapes.AddPicture(assinaturaPath, false, true, novoLeft, novoTop, larguraPadrao, alturaPadrao)
        picAssinatura.Name := "ASSINATURA_REAL"
    } else {
        ExibirMensagem("Aviso — Assinatura", "Assinatura Não Inserida", "Nenhuma imagem de assinatura encontrada no modelo PowerPoint.`n`nO certificado será gerado sem a imagem da assinatura.", "erro", "Continuar")
    }

    LOGO_LARGURA := 320
    LOGO_ALTURA := 200
    if FileExist(ARQUIVO_LOGO) {
        slideW := presBase.PageSetup.SlideWidth
        novoLogoLeft := (slideW - LOGO_LARGURA) / 2
        novoLogoTop := (logoTop > 0) ? logoTop : 20
        picLogo := slideBase.Shapes.AddPicture(ARQUIVO_LOGO, false, true, novoLogoLeft, novoLogoTop, LOGO_LARGURA, LOGO_ALTURA)
        picLogo.Name := "LOGO_EINSTEIN"
    } else {
        ExibirMensagem("Aviso — Logo", "Logotipo Não Localizado", "Arquivo do logo não encontrado em:`n" ARQUIVO_LOGO "`n`nO certificado será gerado sem o logo.", "erro", "Continuar")
    }

    presBase.SaveAs(modeloBase, 24)
    presBase.Close()
    return modeloBase
}

; =========================================================
; FUNÇÃO: Gerar certificado PDF individual (Nuvem + Backup Local)
; =========================================================

; Retorna o caminho completo do PDF gerado (referência canônica para o manifesto:
; a cópia em "Para Enviar" quando existir, senão o caminho principal) ou "" em caso de erro.
GerarCertificado(registro, ppt, pptModelo, pastaNuvemMes, pastaLocalMes, pastaDocente, pastaParaEnviar, nomePos) {
    nome := registro["nome"]
    curso := registro["curso"]
    dataAula := registro["dataAula"]
    horasFormatadas := registro["horasFormatadas"]
    linha := registro["linha"]
    aba := registro["aba"]

    try {
        pres := ppt.Presentations.Open(pptModelo)
        slide := pres.Slides(1)

        Loop slide.Shapes.Count {
            shape := slide.Shapes.Item(A_Index)
            try {
                if !InStr(shape.TextFrame.TextRange.Text, "{NOME}")
                    continue
                SubstituirPlaceholder(shape, "{NOME}", nome)
                SubstituirPlaceholder(shape, "{POSGRADUACAO}", nomePos)
                SubstituirPlaceholder(shape, "{CURSO}", curso)
                SubstituirPlaceholder(shape, "{DATA}", dataAula)
                SubstituirPlaceholder(shape, "{HORAS}", horasFormatadas)
            } catch {
            }
        }

        nomeLimpo := SanitizarNomeArquivo(nome)
        dataLimpa := RegExReplace(dataAula, '[\\/:*?"<>|]', ".")
        nomeBase := nomeLimpo " (" dataLimpa ")"

        ; Tenta salvar primeiro na nuvem (se a pasta existir). Se a gravação na
        ; nuvem falhar (ex.: OneDrive fora do ar, sincronizando, arquivo
        ; bloqueado), cai para o backup local em vez de perder o certificado —
        ; antes, uma falha aqui abortava o certificado sem salvar em lugar
        ; nenhum, nem localmente.
        caminhoPrincipal := ""
        salvouNaNuvem := false
        if (pastaNuvemMes != "" && DirExist(pastaNuvemMes)) {
            try {
                caminhoPrincipal := CaminhoUnico(pastaNuvemMes, nomeBase, ".pdf")
                pres.SaveAs(caminhoPrincipal, 32)
                salvouNaNuvem := true
            } catch as errNuvem {
                caminhoPrincipal := ""
            }
        }
        if !salvouNaNuvem {
            caminhoPrincipal := CaminhoUnico(pastaLocalMes, nomeBase, ".pdf")
            pres.SaveAs(caminhoPrincipal, 32)
        }
        pres.Close()

        ; Espelha no backup local só quando o principal foi salvo na nuvem —
        ; quando caiu no fallback acima, caminhoPrincipal JÁ É o arquivo local
        ; (copiar de novo criaria um ".1.pdf" duplicado na mesma pasta).
        if (salvouNaNuvem && pastaLocalMes != "" && caminhoPrincipal != "") {
            caminhoLocalMes := CaminhoUnico(pastaLocalMes, nomeBase, ".pdf")
            if (caminhoLocalMes != caminhoPrincipal)
                try FileCopy(caminhoPrincipal, caminhoLocalMes, true)
        }

        if (pastaDocente != "" && caminhoPrincipal != "") {
            caminhoDocente := CaminhoUnico(pastaDocente, nomeBase, ".pdf")
            try FileCopy(caminhoPrincipal, caminhoDocente, true)
        }

        ; Cópia para "Para Enviar\<Curso>\<Mês>\" (usada pelo Power Automate).
        ; Vira a referência canônica do arquivo no manifesto, já que é a pasta
        ; que o flow efetivamente vasculha.
        caminhoParaManifesto := caminhoPrincipal
        if (pastaParaEnviar != "" && caminhoPrincipal != "") {
            caminhoParaEnviarFinal := CaminhoUnico(pastaParaEnviar, nomeBase, ".pdf")
            try FileCopy(caminhoPrincipal, caminhoParaEnviarFinal, true)
            caminhoParaManifesto := caminhoParaEnviarFinal
        }

        return caminhoParaManifesto
    } catch as err {
        ExibirMensagem("Erro na Geração", "Falha no Certificado", "Erro — Aba: " aba " / Linha: " linha "`n`n" err.Message, "erro", "Continuar")
        return ""
    }
}

; =========================================================
; FUNÇÃO: Caminho único dentro de uma pasta
; =========================================================

CaminhoUnico(pasta, nomeBase, extensao) {
    caminho := pasta nomeBase extensao
    if !FileExist(caminho)
        return caminho
    contador := 1
    Loop {
        caminho := pasta nomeBase "." contador extensao
        if !FileExist(caminho)
            return caminho
        contador++
    }
}
