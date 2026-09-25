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
;   - v23: o manifesto.txt agora é gravado DIRETO com o nome final. Antes ele
;     era gravado como manifesto.txt.tmp e depois renomeado, e o gatilho do
;     Power Automate ("Quando um arquivo é criado ou modificado") não
;     disparava: via a criação do .tmp (ignorada pela condição de gatilho) e
;     depois só uma renomeação. Testado em 24/09/2026: editar e salvar o
;     manifesto disparava, gerar um novo não. O risco de o fluxo ler o
;     arquivo pela metade é coberto pelo Atraso de 3 minutos no início dele.
;   - v24: robustez contra o Excel/PowerPoint "caindo" no meio da rodada
;     (erros 0x800706BE = o programa morreu durante a chamada; 0x800706BA =
;     o programa já não existe mais). Em 25/09/2026 isso derrubava o lote
;     inteiro logo depois da escolha da assinatura, mesmo após reiniciar o PC.
;     (1) LOG: cada rodada grava C:\Certificados\Logs\geracao_<data>.txt com
;         os passos e o erro completo (linha do script, pilha), pra achar a
;         causa exata em vez de só o código 0x8007....
;     (2) VERIFICAÇÃO PRÉ-RODAGEM: antes de começar, pede pra fechar Excel e
;         PowerPoint abertos (forçando só com confirmação) e confere pastas,
;         logo e permissão de gravação em C:\Certificados.
;     (3) LER TUDO ANTES DE GERAR: planilhas, modelo, logo e assinaturas são
;         COPIADOS pra uma pasta temporária local (o Office nunca abre arquivo
;         direto do OneDrive). Cada planilha é aberta, lida pra memória e
;         fechada na hora; o Excel é encerrado antes da geração. A Fase 2
;         (validação/correções) trabalha só na memória. Os PDFs são salvos
;         primeiro no disco local e só depois copiados pra nuvem.
;     (4) REINÍCIO AUTOMÁTICO: se o Excel ou o PowerPoint cair, o script abre
;         um novo e tenta de novo a mesma etapa (até 3x). Certificados que
;         falharem mesmo assim aparecem listados no relatório final, e o
;         Excel/PowerPoint do script é sempre encerrado ao sair (sem sobrar
;         processo "zumbi" pra próxima rodada).
;     Efeito colateral corrigido: linhas "puladas" na validação eram marcadas
;     por aba+linha e podiam pular a mesma linha de OUTRO curso com o mesmo
;     mês no lote; agora a marcação fica no próprio registro. E a correção
;     manual de carga horária com minutos (ex.: 2:30) agora vale de fato.
; ════════════════════════════════════════════════════════════════════════

; =========================================================
; CONFIGURAÇÕES DINÂMICAS DE PASTAS, LOGO E NUVEM
; =========================================================

global INFO_NUVEM        := ObterCaminhosCompartilhados()
global PASTA_ASSINATURAS := (INFO_NUVEM["pastaAssinaturas"] != "" && DirExist(INFO_NUVEM["pastaAssinaturas"])) ? INFO_NUVEM["pastaAssinaturas"] : "C:\CAssinaturas\"
global PASTA_SAIDA       := "C:\Certificados\"
global ARQUIVO_LOGO      := "C:\CAssinaturas\_logo_einstein.jpg"
global SEMESTRE_ATUAL    := "2026-5"

; ─── Estado do Excel/PowerPoint controlados por este script (v24) ───
global excel := "", ppt := "", wb := ""
global EXCEL_PID := 0, PPT_PID := 0

; ─── Log, pasta de trabalho local e listas para o relatório final (v24) ───
global LOG_ARQUIVO        := ""
global PASTA_TRABALHO     := A_Temp "\CertificadosAHK\" FormatTime(A_Now, "yyyyMMdd_HHmmss") "\"
global COPIAS_LOCAIS      := Map()
global CERTS_COM_FALHA    := []
global CURSOS_REMOVIDOS   := []

IniciarLog()
OnError(RegistrarErroNaoTratado)
OnExit(EncerrarOfficeAoSair)

; =========================================================
; VERIFICAÇÃO PRÉ-RODAGEM (v24)
; =========================================================

VerificacaoPreRodagem()

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

RegistrarLog("Planilha mestre: " masterPath)
masterPathLocal := CopiarParaPastaTrabalho(masterPath)

; =========================================================
; PASSO 2 — LISTAR ABAS E PEDIR SELEÇÃO DO SEMESTRE
; =========================================================

; v24: a planilha mestre (cópia local) é aberta só pelo tempo de ler as abas
; e fechada logo em seguida — nada fica aberto enquanto você responde às telas.
abasMestreArr := []
try {
    abasMestreArr := ExecutarComRetentativa("Listar abas da planilha mestre", "excel", LerAbasDaPlanilha.Bind(masterPathLocal))
} catch as err {
    LogErro(err, "Listar abas da planilha mestre")
    ExibirMensagem(
        "Erro no Excel", "Falha ao Ler Abas da Planilha Mestre",
        "O Excel não conseguiu abrir ou ler a planilha mestre, mesmo tentando de novo.`n`n"
        "Detalhe técnico: " err.Message "`n`n"
        "O registro completo do erro está no log:`n" LOG_ARQUIVO,
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
    ExitApp
}

nomeAba := selecaoAba[1]
SEMESTRE_ATUAL := nomeAba ; ex: "2026-5", "2027-1", etc.
RegistrarLog("Semestre / aba da planilha mestre: " nomeAba)

; =========================================================
; PASSO 3 — LER CURSOS E SELEÇÃO EM LOTE
; =========================================================

try {
    cursos := ExecutarComRetentativa("Ler cursos da planilha mestre", "excel", LerCursosDaMestre.Bind(masterPathLocal, nomeAba))
} catch as err {
    LogErro(err, "Ler cursos da planilha mestre")
    ExibirMensagem("Erro no Excel", "Falha ao Ler a Planilha Mestre", "Não foi possível ler os cursos da aba '" nomeAba "'.`n`nDetalhe técnico: " err.Message "`n`nLog completo:`n" LOG_ARQUIVO, "erro", "Fechar")
    ExitApp
}

if !IsObject(cursos) {
    ExibirMensagem("Erro de Aba", "Aba Não Encontrada", "Aba '" nomeAba "' não foi localizada na planilha.", "erro", "Fechar")
    ExitApp
}
RegistrarLog("Cursos encontrados na aba: " cursos.Length)

if (cursos.Length = 0) {
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
        ExitApp
    }
    dataEmissao := Trim(inputData["Value"])
}

CriarPastaSegura(PASTA_SAIDA)
RegistrarLog("Data de emissão: " dataEmissao " | Cursos no lote: " selecionados.Length)

; v24: o PowerPoint não é mais aberto aqui. Ele só é iniciado quando for
; realmente usado (criação do modelo-base e geração), por GarantirPowerPoint(),
; que também o reabre sozinho se ele tiver caído.

; =========================================================
; ================  FASE 1 — CARREGAMENTO  ================
; =========================================================

cacheCoord     := Map()
ultimoTemplate := ""
mesesGlobais   := []

cursoConfigs := []

for numCursoAtual, cursoSelecionado in selecionados {

    RegistrarLog("── Carregando curso " numCursoAtual "/" selecionados.Length ": " cursoSelecionado["posgrad"] " (" cursoSelecionado["unidade"] ")")
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

        ; v24: copia a assinatura pra pasta local e confere se é uma imagem de
        ; verdade (PNG/JPG/GIF/BMP) antes de entregar ao PowerPoint.
        RegistrarLog("Coordenador: " coordRaw " | Assinatura: " assinaturaPath)
        assinaturaLocal := PrepararImagem(assinaturaPath, "assinatura de " coordRaw)

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

        ; v24: a planilha do curso é copiada pra pasta local e só é aberta pelo
        ; tempo de listar as abas (fechada antes da tela de escolha do mês).
        RegistrarLog("Planilha do curso: " cursoPath)
        RegistrarLog("Modelo PowerPoint: " pptModelo)
        cursoPathLocal := CopiarParaPastaTrabalho(cursoPath)
        abasCursoArr := ExecutarComRetentativa("Listar abas — " cursoSelecionado["posgrad"], "excel", LerAbasDaPlanilha.Bind(cursoPathLocal))

        if (mesesGlobais.Length = 0) {
            selecaoAbasCurso := SelecionarOpcoesUI(
                "Passo 7 — Selecionar Período da Rodada",
                "Mês(es) a Processar",
                abasCursoArr,
                false,
                580
            )
            if (selecaoAbasCurso.Length = 0)
                throw Error("Nenhum mês selecionado para este curso.")
            mesesGlobais  := selecaoAbasCurso
            arrAbasCurso  := selecaoAbasCurso
        } else {
            arrAbasCurso := []
            for m in mesesGlobais {
                if AbaNaLista(abasCursoArr, m)
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
                if (selecaoAbasCurso.Length = 0)
                    throw Error("Nenhum mês selecionado para este curso.")
                arrAbasCurso := selecaoAbasCurso
            }
        }

        abasInvalidas := ""
        for idx, nomeAbaCurso in arrAbasCurso {
            if !AbaNaLista(abasCursoArr, nomeAbaCurso)
                abasInvalidas .= "  • " nomeAbaCurso "`n"
        }
        if (abasInvalidas != "")
            throw Error("Abas não encontradas:`n" abasInvalidas)

        ; v24: lê TODAS as linhas das abas escolhidas pra memória e fecha a
        ; planilha. Daqui pra frente (validação e geração) o Excel não é usado.
        dadosCurso := ExecutarComRetentativa("Ler dados — " cursoSelecionado["posgrad"], "excel", LerDadosDaPlanilha.Bind(cursoPathLocal, arrAbasCurso))
        nomePos := dadosCurso["nomePos"]
        if (nomePos = "")
            throw Error("Célula A2 da aba '" arrAbasCurso[1] "' está vazia.")

        pptModeloLocal := CopiarParaPastaTrabalho(pptModelo)
        logoLocal := FileExist(ARQUIVO_LOGO) ? CopiarParaPastaTrabalho(ARQUIVO_LOGO) : ARQUIVO_LOGO

        modeloBase := ExecutarComRetentativa(
            "Criar modelo-base — " cursoSelecionado["posgrad"], "ppt",
            CriarModeloBase.Bind(pptModeloLocal, coordUsado, cargo, dataEmissao, assinaturaLocal, logoLocal)
        )
        RegistrarLog("Modelo-base criado: " modeloBase)

        cursoConfigs.Push(Map(
            "posgrad", cursoSelecionado["posgrad"],
            "unidade", cursoSelecionado["unidade"],
            "siglaUnidade", cursoSelecionado["siglaUnidade"],
            "cursoPath", cursoPath,
            "abas", arrAbasCurso,
            "dadosAbas", dadosCurso["dadosAbas"],
            "nomePos", nomePos,
            "modeloBase", modeloBase,
            "falhouCarregamento", false
        ))
        RegistrarLog("Curso carregado com sucesso.")

    } catch as errCarregamento {
        LogErro(errCarregamento, "Carregamento do curso " cursoSelecionado["posgrad"])
        CURSOS_REMOVIDOS.Push(cursoSelecionado["posgrad"] " — " errCarregamento.Message)
        ExibirMensagem(
            "Erro no Carregamento",
            "Curso Removido do Lote",
            "O curso a seguir não pôde ser carregado e foi removido desta rodada:`n`n"
            cursoSelecionado["posgrad"] "`n`nMotivo: " errCarregamento.Message
            "`n`nDetalhes completos no log:`n" LOG_ARQUIVO,
            "erro",
            "Continuar"
        )
        continue
    }
}

; v24: todos os dados já estão na memória — o Excel não é mais necessário.
FinalizarExcel()

if (cursoConfigs.Length = 0) {
    ExibirMensagem("Encerrando", "Nenhum Curso Carregado", "Nenhum curso foi carregado com sucesso. O programa será encerrado.", "erro", "Fechar")
    ExitApp
}

; =========================================================
; ================  FASE 2 — VALIDAÇÃO  ====================
; =========================================================

todosProblemas      := []
totalMescladasGeral := 0
totalOkGeral         := 0

; v24: a validação trabalha nos registros já lidos pra memória na Fase 1.
; Cada "issue" aponta pro PRÓPRIO registro, então as correções e o "pular"
; abaixo alteram direto o que a Fase 3 vai usar.
for cfg in cursoConfigs {
    for dadosAba in cfg["dadosAbas"] {
        resultado := ValidarRegistros(dadosAba["registros"], dadosAba["aba"])
        totalMescladasGeral += resultado["mescladasCount"]
        totalOkGeral += resultado["totalOk"]
        for issue in resultado["issues"] {
            issue["curso"] := cfg["posgrad"]
            todosProblemas.Push(issue)
        }
    }
}

resumo := "Linhas prontas para geração: " totalOkGeral "`n"
resumo .= "Células mescladas lidas automaticamente: " totalMescladasGeral "`n"
resumo .= "Linhas com inconsistências encontradas: " todosProblemas.Length "`n"
RegistrarLog("Validação: " StrReplace(Trim(resumo, "`n"), "`n", " | "))

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
        RegistrarLog("Usuário cancelou na tela de validação.")
        for cfg in cursoConfigs
            try FileDelete(cfg["modeloBase"])
        ExitApp
    }

    if (escolhaValidacao = "corrigir") {
        for p in todosProblemas {
            r := p["registro"]
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
                r["pular"] := true
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
                    CorrigirNomeRegistro(r, Trim(cx["Value"]))
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
                    CorrigirCursoRegistro(r, Trim(cx["Value"]))
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
                    CorrigirDataRegistro(r, Trim(cx["Value"]))
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
                        CorrigirHorasRegistro(r, duracaoFormatada)
                }
            }
            RegistrarLog("Correção manual — " p["curso"] " / " p["aba"] " / linha " linha)
        }
    } else {
        for p in todosProblemas
            p["registro"]["pular"] := true
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
    RegistrarLog("── Gerando curso: " cfg["posgrad"] " (" cfg["unidade"] ")")
    try {
        resultadoCurso := GerarCertificadosCurso(cfg, PASTA_SAIDA)
        relatorioPorCurso.Push(Map(
            "posgrad", cfg["posgrad"], "unidade", cfg["unidade"],
            "gerados", resultadoCurso["gerados"], "erros", resultadoCurso["erros"], "falhou", false
        ))
        totalGeralGerados += resultadoCurso["gerados"]
        totalGeralErros   += resultadoCurso["erros"]
        RegistrarLog("Curso concluído: " resultadoCurso["gerados"] " gerados, " resultadoCurso["erros"] " ignorados/com falha.")
    } catch as errCurso {
        LogErro(errCurso, "Geração do curso " cfg["posgrad"])
        totalCursosComFalha++
        relatorioPorCurso.Push(Map(
            "posgrad", cfg["posgrad"], "unidade", cfg["unidade"],
            "gerados", 0, "erros", 0, "falhou", true, "motivo", errCurso.Message
        ))
    }
    try FileDelete(cfg["modeloBase"])
}

FinalizarPowerPoint()
FinalizarExcel()
RegistrarLog("Geração encerrada. PDFs gerados: " totalGeralGerados " | Cursos com falha: " totalCursosComFalha " | Certificados com falha: " CERTS_COM_FALHA.Length)

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

; v24: cursos que nem chegaram a ser carregados e certificados que falharam
; mesmo depois das novas tentativas (antes não apareciam no relatório).
if (CURSOS_REMOVIDOS.Length > 0) {
    detalheRelatorio .= "⚠️ CURSOS REMOVIDOS NO CARREGAMENTO:`n"
    for txt in CURSOS_REMOVIDOS
        detalheRelatorio .= "  • " txt "`n"
    detalheRelatorio .= "`n"
}
if (CERTS_COM_FALHA.Length > 0) {
    detalheRelatorio .= "⚠️ CERTIFICADOS QUE NÃO FORAM GERADOS:`n"
    for txt in CERTS_COM_FALHA
        detalheRelatorio .= "  • " txt "`n"
    detalheRelatorio .= "`n"
}

resumoFinal := "Cursos processados no lote: " cursoConfigs.Length "`n"
resumoFinal .= "Data de emissão impressa:   " dataEmissao "`n"
resumoFinal .= "Total de PDFs gerados:      " totalGeralGerados "`n"
resumoFinal .= "Cursos que falharam:        " totalCursosComFalha "`n"
resumoFinal .= "Células mescladas tratadas: " totalMescladasGeral "`n"
resumoFinal .= "Log desta rodada: " LOG_ARQUIVO

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
; FUNÇÃO: Validar os registros de uma aba já lidos pra memória (Fase 2)
; v24: mesma regra da antiga ValidarAba, mas sem tocar no Excel.
; =========================================================

ValidarRegistros(registros, abaNome) {
    issues := []
    mescladasCount := 0
    totalOk := 0

    for reg in registros {
        linha := reg["linha"]

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

GerarCertificadosCurso(cfg, pastaSaidaBase) {
    totalGerados := 0
    totalErros := 0
    falhasSeguidas := 0

    posgradSanit := SanitizarNomeArquivo(cfg["posgrad"])
    nomeCursoComUnidade := (cfg["unidade"] != "") ? posgradSanit " (" SanitizarNomeArquivo(cfg["unidade"]) ")" : posgradSanit

    ; v24: percorre os registros já lidos na Fase 1 (o Excel já foi fechado).
    for dadosAba in cfg["dadosAbas"] {
        abaNome := dadosAba["aba"]
        mesPasta := ConverterMesParaPasta(abaNome)

        pastas := GarantirPastasSaida(nomeCursoComUnidade, mesPasta, SEMESTRE_ATUAL)
        pastaNuvemMes := pastas["nuvem"]
        pastaLocalMes := pastas["local"]
        pastaDocentesRaiz := pastas["localDocentes"]
        pastaParaEnviarMes := GarantirPastaParaEnviar(nomeCursoComUnidade, mesPasta)
        RegistrarLog("Aba " abaNome " → nuvem: " pastaNuvemMes " | local: " pastaLocalMes " | Para Enviar: " pastaParaEnviarMes)

        ; Agrupamento por docente para o manifesto do Power Automate (Fase 3).
        manifestoDocentes := Map()
        ordemManifesto := []

        for reg in dadosAba["registros"] {
            linha := reg["linha"]

            if reg.Has("pular") && reg["pular"]
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

            caminhoGerado := ""
            try {
                caminhoGerado := GerarCertificadoComRetentativa(registro, cfg["modeloBase"], pastaNuvemMes, pastaLocalMes, pastaDocente, pastaParaEnviarMes, cfg["nomePos"])
                falhasSeguidas := 0
            } catch as errCert {
                falhasSeguidas++
                LogErro(errCert, "Certificado — " cfg["posgrad"] " / " abaNome " / linha " linha " / " reg["nome"])
                CERTS_COM_FALHA.Push(cfg["posgrad"] " / " abaNome " / linha " linha " — " reg["nome"] ": " errCert.Message)
                ; Se o PowerPoint falha em 3 certificados seguidos (mesmo com as
                ; novas tentativas), o problema não é daquela linha — interrompe
                ; o curso em vez de repetir o erro em todas as linhas restantes.
                if (falhasSeguidas >= 3)
                    throw Error("O PowerPoint falhou em 3 certificados seguidos, mesmo reiniciando. Curso interrompido na aba '" abaNome "'. Último erro: " errCert.Message)
            }
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
        EscreverArquivoDireto(pastaParaEnviarMes "manifesto.txt", jsonTexto)
        RegistrarLog("Manifesto gravado (" listaDocentesValidos.Length " docentes): " pastaParaEnviarMes "manifesto.txt")
    } catch as errManifesto {
        LogErro(errManifesto, "Gravar manifesto em " pastaParaEnviarMes)
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

; Grava direto no nome final (sem .tmp + renomear): o gatilho do Power
; Automate não dispara com renomeação, só com criação/modificação do
; próprio manifesto.txt. Se já existir, é sobrescrito (conta como modificação).
EscreverArquivoDireto(caminhoFinal, conteudo) {
    arq := ""
    try {
        arq := FileOpen(caminhoFinal, "w", "UTF-8-RAW")
        if !IsObject(arq)
            throw Error("Não foi possível criar o arquivo: " caminhoFinal)
        arq.Write(conteudo)
        arq.Close()
    } catch as err {
        try arq.Close()
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
; v24: usa o PowerPoint controlado pelo script (reabre se tiver caído) e, se
; der erro no meio, fecha a apresentação aberta antes de repassar o erro —
; ExecutarComRetentativa decide se tenta de novo.
; =========================================================

CriarModeloBase(pptModelo, coordUsado, cargo, dataEmissao, assinaturaPath, ARQUIVO_LOGO) {
    pptApp := GarantirPowerPoint()
    presBase := ""
    try {
        return CriarModeloBaseInterno(pptApp, pptModelo, coordUsado, cargo, dataEmissao, assinaturaPath, ARQUIVO_LOGO, &presBase)
    } catch as err {
        if IsObject(presBase)
            try presBase.Close()
        throw err
    }
}

CriarModeloBaseInterno(ppt, pptModelo, coordUsado, cargo, dataEmissao, assinaturaPath, ARQUIVO_LOGO, &presBase) {

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
; FUNÇÃO: Gerar certificado PDF individual (Local + Nuvem)
; v24: com novas tentativas. Se o PowerPoint cair (0x800706BA/BE) ou estiver
; ocupado, apaga o que a tentativa que falhou chegou a gravar (pra não sobrar
; PDF duplicado ".1.pdf"), reabre o PowerPoint e tenta de novo o MESMO
; certificado, até 3 vezes. Se ainda assim falhar, lança o erro — quem chama
; registra no log e lista no relatório final.
; =========================================================

GerarCertificadoComRetentativa(registro, pptModelo, pastaNuvemMes, pastaLocalMes, pastaDocente, pastaParaEnviar, nomePos, tentativas := 3) {
    Loop tentativas {
        criados := []
        try {
            return GerarCertificado(registro, pptModelo, pastaNuvemMes, pastaLocalMes, pastaDocente, pastaParaEnviar, nomePos, criados)
        } catch as err {
            for arquivoCriado in criados
                try FileDelete(arquivoCriado)
            if (A_Index >= tentativas || !EhErroDeOffice(err))
                throw err
            LogErro(err, "Certificado de " registro["nome"] " (linha " registro["linha"] ") — tentativa " A_Index " de " tentativas ", tentando de novo")
            if EhErroQuedaOffice(err)
                ReiniciarPowerPoint()
            else
                Sleep 1500
        }
    }
}

; Retorna o caminho completo do PDF gerado (referência canônica para o manifesto:
; a cópia em "Para Enviar" quando existir, senão o caminho principal).
; v24: o PowerPoint grava o PDF SEMPRE no disco local (C:\Certificados\...) e
; só depois o arquivo é copiado pra nuvem — o PowerPoint nunca escreve direto
; numa pasta do OneDrive. Se a cópia pra nuvem falhar, o local vira o principal
; (mesmo resultado do fallback da v21). Cada arquivo gravado entra em "criados".
GerarCertificado(registro, pptModelo, pastaNuvemMes, pastaLocalMes, pastaDocente, pastaParaEnviar, nomePos, criados) {
    nome := registro["nome"]
    curso := registro["curso"]
    dataAula := registro["dataAula"]
    horasFormatadas := registro["horasFormatadas"]

    pptApp := GarantirPowerPoint()
    pres := ""
    try {
        pres := pptApp.Presentations.Open(pptModelo)
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

        caminhoLocal := CaminhoUnico(pastaLocalMes, nomeBase, ".pdf")
        pres.SaveAs(caminhoLocal, 32)
        criados.Push(caminhoLocal)
        pres.Close()
        pres := ""
    } catch as err {
        if IsObject(pres)
            try pres.Close()
        throw err
    }

    ; Cópia pra nuvem (Declarações\Semestre\Pós\Mês). Vira o caminho principal
    ; quando dá certo, como nas versões anteriores.
    caminhoPrincipal := caminhoLocal
    if (pastaNuvemMes != "" && DirExist(pastaNuvemMes)) {
        try {
            caminhoNuvem := CaminhoUnico(pastaNuvemMes, nomeBase, ".pdf")
            FileCopy(caminhoLocal, caminhoNuvem, true)
            criados.Push(caminhoNuvem)
            caminhoPrincipal := caminhoNuvem
        } catch as errNuvem {
            LogErro(errNuvem, "Cópia pra nuvem de " nomeBase " (ficou só no backup local)")
        }
    }

    if (pastaDocente != "") {
        caminhoDocente := CaminhoUnico(pastaDocente, nomeBase, ".pdf")
        try {
            FileCopy(caminhoLocal, caminhoDocente, true)
            criados.Push(caminhoDocente)
        }
    }

    ; Cópia para "Para Enviar\<Curso>\<Mês>\" (usada pelo Power Automate).
    ; Vira a referência canônica do arquivo no manifesto, já que é a pasta
    ; que o flow efetivamente vasculha.
    caminhoParaManifesto := caminhoPrincipal
    if (pastaParaEnviar != "") {
        caminhoParaEnviarFinal := CaminhoUnico(pastaParaEnviar, nomeBase, ".pdf")
        try {
            FileCopy(caminhoLocal, caminhoParaEnviarFinal, true)
            criados.Push(caminhoParaEnviarFinal)
        }
        caminhoParaManifesto := caminhoParaEnviarFinal
    }

    return caminhoParaManifesto
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

; ════════════════════════════════════════════════════════════════════════
;   v24 — LOG, VERIFICAÇÃO PRÉ-RODAGEM, CÓPIAS LOCAIS E CONTROLE DO OFFICE
; ════════════════════════════════════════════════════════════════════════

; ─── LOG ────────────────────────────────────────────────────────────────

IniciarLog() {
    global LOG_ARQUIVO
    pastaLogs := PASTA_SAIDA "Logs\"
    try DirCreate(pastaLogs)
    if !DirExist(pastaLogs)
        pastaLogs := A_Temp "\"
    LOG_ARQUIVO := pastaLogs "geracao_" FormatTime(A_Now, "yyyy-MM-dd_HH-mm-ss") ".txt"
    RegistrarLog("════ " A_ScriptName " ════")
    RegistrarLog("Computador: " A_ComputerName " | Usuário: " A_UserName " | Windows " A_OSVersion " | AutoHotkey " A_AhkVersion)
    RegistrarLog("OneDrive: " INFO_NUVEM["raizOneDrive"])
    RegistrarLog("Pasta de trabalho local: " PASTA_TRABALHO)
}

; (O nome não pode ser só "Log": já existe uma função RegistrarLog() do AutoHotkey.)
RegistrarLog(msg) {
    if (LOG_ARQUIVO = "")
        return
    try FileAppend(FormatTime(A_Now, "HH:mm:ss") "  " msg "`r`n", LOG_ARQUIVO, "UTF-8")
}

DescreverErro(err) {
    if !IsObject(err)
        return String(err)
    txt := ""
    try txt := err.Message
    try {
        if (err.Extra != "")
            txt .= " | Extra: " err.Extra
    }
    try txt .= " | Onde: " err.What
    try txt .= " | Linha do script: " err.Line
    return txt
}

LogErro(err, contexto) {
    RegistrarLog("ERRO — " contexto)
    RegistrarLog("   " DescreverErro(err))
    try {
        pilha := Trim(err.Stack, "`r`n")
        if (pilha != "")
            RegistrarLog("   Pilha:`r`n      " StrReplace(pilha, "`n", "`n      "))
    }
}

RegistrarErroNaoTratado(err, modo) {
    LogErro(err, "ERRO NÃO TRATADO (o AutoHotkey vai mostrar a mensagem padrão dele)")
    return 0
}

; Roda em QUALQUER saída do script (fim normal, Cancelar, erro não tratado):
; fecha o Excel/PowerPoint abertos por este script, pra não sobrar processo
; invisível travando a próxima rodada, e apaga a pasta de trabalho local.
EncerrarOfficeAoSair(motivo, codigo) {
    RegistrarLog("Encerrando o script (" motivo ").")
    FinalizarPowerPoint()
    FinalizarExcel()
    try DirDelete(PASTA_TRABALHO, true)
    RegistrarLog("Fim.")
    return 0
}

; ─── VERIFICAÇÃO PRÉ-RODAGEM ────────────────────────────────────────────

VerificacaoPreRodagem() {
    RegistrarLog("Verificação pré-rodagem...")
    if !FecharOfficeAberto() {
        RegistrarLog("Rodada cancelada na verificação pré-rodagem (Excel/PowerPoint abertos).")
        ExitApp
    }

    avisos := []
    if (INFO_NUVEM["raizOneDrive"] = "") {
        avisos.Push("Pasta do OneDrive da empresa não encontrada. Você vai precisar escolher as planilhas e o modelo manualmente, e os PDFs ficarão só em " PASTA_SAIDA ".")
    } else if !DirExist(INFO_NUVEM["pastaDeclaracoes"]) {
        avisos.Push("Pasta 'Planilha compartilhada - Declarações' não encontrada no OneDrive:`n   " INFO_NUVEM["pastaDeclaracoes"])
    }
    if !DirExist(PASTA_ASSINATURAS)
        avisos.Push("Pasta de assinaturas não encontrada:`n   " PASTA_ASSINATURAS)
    if !FileExist(ARQUIVO_LOGO) {
        avisos.Push("Logo não encontrado (os certificados sairão sem logo):`n   " ARQUIVO_LOGO)
    } else {
        tipoLogo := TipoRealImagem(ARQUIVO_LOGO)
        if !ImagemAceita(tipoLogo)
            avisos.Push("O arquivo do logo não parece ser uma imagem válida (" tipoLogo "):`n   " ARQUIVO_LOGO)
    }
    try {
        DirCreate(PASTA_SAIDA)
        arqTeste := PASTA_SAIDA "~teste_gravacao.tmp"
        FileAppend("ok", arqTeste)
        FileDelete(arqTeste)
    } catch as err {
        avisos.Push("Não foi possível gravar em " PASTA_SAIDA " (" err.Message ").")
    }
    try {
        DirCreate(PASTA_TRABALHO)
    } catch as err {
        avisos.Push("Não foi possível criar a pasta de trabalho temporária (" err.Message "). As planilhas serão abertas direto do OneDrive.")
    }

    if (avisos.Length = 0) {
        RegistrarLog("Pré-rodagem OK.")
        return
    }

    txt := ""
    for aviso in avisos {
        RegistrarLog("AVISO pré-rodagem: " StrReplace(aviso, "`n", " "))
        txt .= "• " aviso "`n`n"
    }
    continuar := Confirmar(
        "Verificação Pré-Rodagem",
        "Pontos de Atenção Antes de Começar",
        txt "Deseja continuar mesmo assim?",
        "▶️ Continuar",
        "❌ Cancelar",
        true,
        660
    )
    if !continuar {
        RegistrarLog("Rodada cancelada pelo usuário na verificação pré-rodagem.")
        ExitApp
    }
}

; Garante que não há Excel/PowerPoint abertos antes de começar. Primeiro pede
; pra fechar normalmente (o Office pergunta se quer salvar o que estiver
; aberto); só força o fechamento com confirmação explícita.
FecharOfficeAberto() {
    programas := Map("EXCEL.EXE", "Excel", "POWERPNT.EXE", "PowerPoint")
    Loop {
        abertos := ""
        for exe, nomeProg in programas {
            if ProcessExist(exe)
                abertos .= (abertos = "" ? "" : " e o ") nomeProg
        }
        if (abertos = "")
            return true

        RegistrarLog("Pré-rodagem: encontrado aberto — " abertos)
        fecharAgora := Confirmar(
            "Verificação Pré-Rodagem",
            "Feche o " abertos " Antes de Começar",
            "O " abertos " está aberto neste computador (às vezes invisível, sobrando de uma rodada anterior que deu erro).`n`n"
            "Para a geração funcionar, o script precisa usar o Excel e o PowerPoint sozinho.`n`n"
            "👉 SALVE qualquer planilha ou apresentação sua que esteja aberta e clique em 'Fechar Agora'.`n`n"
            "Dica: enquanto os certificados estiverem sendo gerados, não abra arquivos do Excel nem do PowerPoint.",
            "🔒 Fechar Agora",
            "❌ Cancelar",
            true,
            640
        )
        if !fecharAgora
            return false

        for exe in programas {
            for hwnd in WinGetList("ahk_exe " exe)
                try WinClose("ahk_id " hwnd)
        }
        Loop 20 {
            if (!ProcessExist("EXCEL.EXE") && !ProcessExist("POWERPNT.EXE"))
                break
            Sleep 500
        }
        if (!ProcessExist("EXCEL.EXE") && !ProcessExist("POWERPNT.EXE")) {
            RegistrarLog("Pré-rodagem: Excel/PowerPoint fechados normalmente.")
            return true
        }

        forcar := Confirmar(
            "Verificação Pré-Rodagem",
            "Ainda Há Excel ou PowerPoint Aberto",
            "Algum Excel ou PowerPoint não fechou.`n`n"
            "• Se apareceu uma janela perguntando se você quer SALVAR, responda a ela e clique em 'Verificar de Novo'.`n`n"
            "• Se não apareceu nada, provavelmente é um processo travado e invisível: clique em 'Forçar Fechamento'.`n`n"
            "⚠️ Forçar o fechamento descarta alterações não salvas.",
            "⚠️ Forçar Fechamento",
            "🔁 Verificar de Novo",
            false,
            640
        )
        if forcar {
            for exe in programas {
                Loop 20 {
                    if !ProcessExist(exe)
                        break
                    try ProcessClose(exe)
                    Sleep 250
                }
            }
            RegistrarLog("Pré-rodagem: Excel/PowerPoint encerrados à força.")
        }
    }
}

; ─── CÓPIAS LOCAIS E CONFERÊNCIA DE IMAGENS ─────────────────────────────

; Copia o arquivo pra pasta de trabalho local e devolve o caminho da cópia.
; Assim o Excel/PowerPoint nunca abre arquivo direto da pasta do OneDrive
; (sincronização/coautoria em andamento podem travar ou derrubar o Office).
; Se a cópia falhar, devolve o caminho original (e registra no log).
CopiarParaPastaTrabalho(caminhoOriginal) {
    if (COPIAS_LOCAIS.Has(caminhoOriginal) && FileExist(COPIAS_LOCAIS[caminhoOriginal]))
        return COPIAS_LOCAIS[caminhoOriginal]

    SplitPath(caminhoOriginal, &nomeArquivo)
    destino := PASTA_TRABALHO (COPIAS_LOCAIS.Count + 1) "_" nomeArquivo
    try {
        DirCreate(PASTA_TRABALHO)
        FileCopy(caminhoOriginal, destino, true)
        COPIAS_LOCAIS[caminhoOriginal] := destino
        RegistrarLog("Cópia local: " caminhoOriginal " → " destino)
        return destino
    } catch as err {
        LogErro(err, "Não deu pra copiar " caminhoOriginal " pra pasta local — usando o arquivo original")
        return caminhoOriginal
    }
}

; Lê os primeiros bytes do arquivo pra saber o formato REAL da imagem
; (uma imagem pode ter extensão .png e ser outra coisa por dentro).
TipoRealImagem(caminho) {
    try {
        arq := FileOpen(caminho, "r")
        if !IsObject(arq)
            return "arquivo ilegível"
        buf := Buffer(8, 0)
        lidos := arq.RawRead(buf, 8)
        arq.Close()
        if (lidos < 4)
            return "arquivo vazio"
        b1 := NumGet(buf, 0, "UChar")
        b2 := NumGet(buf, 1, "UChar")
        b3 := NumGet(buf, 2, "UChar")
        b4 := NumGet(buf, 3, "UChar")
        if (b1 = 0x89 && b2 = 0x50 && b3 = 0x4E && b4 = 0x47)
            return "PNG"
        if (b1 = 0xFF && b2 = 0xD8 && b3 = 0xFF)
            return "JPEG"
        if (b1 = 0x47 && b2 = 0x49 && b3 = 0x46)
            return "GIF"
        if (b1 = 0x42 && b2 = 0x4D)
            return "BMP"
        return "formato desconhecido"
    } catch as err {
        return "arquivo ilegível: " err.Message
    }
}

ImagemAceita(tipo) {
    return (tipo = "PNG" || tipo = "JPEG" || tipo = "GIF" || tipo = "BMP")
}

; Copia a imagem pra pasta local e confere se é imagem de verdade. Lança um
; erro com mensagem clara se não for (em vez de deixar o PowerPoint travar).
PrepararImagem(caminhoOriginal, descricao) {
    caminhoLocal := CopiarParaPastaTrabalho(caminhoOriginal)
    tipo := TipoRealImagem(caminhoLocal)
    tamanhoKB := 0
    try tamanhoKB := Round(FileGetSize(caminhoLocal) / 1024)
    RegistrarLog("Imagem (" descricao "): " tipo ", " tamanhoKB " KB")
    if !ImagemAceita(tipo)
        throw Error("A imagem da " descricao " não é um arquivo de imagem válido (" tipo "):`n" caminhoOriginal "`n`nAbra a imagem, salve de novo como PNG ou JPG e tente outra vez.")
    return caminhoLocal
}

; ─── EXCEL E POWERPOINT CONTROLADOS PELO SCRIPT ─────────────────────────

; 0x800706BA = o programa não existe mais | 0x800706BE = morreu durante a
; chamada | 0x800706BF = chamada falhou sem executar | 0x80010108 = objeto
; desconectado | 0x80010114 = objeto inválido.
EhErroQuedaOffice(err) {
    try return RegExMatch(err.Message, "i)0x800706B[AEF]|0x80010108|0x80010114") > 0
    return false
}

; 0x80010001 = chamada rejeitada | 0x8001010A = programa ocupado. O programa
; está vivo, só precisa de um instante.
EhErroOcupadoOffice(err) {
    try return RegExMatch(err.Message, "i)0x80010001|0x8001010A") > 0
    return false
}

EhErroDeOffice(err) {
    return EhErroQuedaOffice(err) || EhErroOcupadoOffice(err)
}

; Executa fn (sem argumentos) e, se o Excel/PowerPoint cair ou estiver
; ocupado, reabre o programa (app = "excel" ou "ppt") e tenta de novo.
; Qualquer outro tipo de erro é repassado na hora, sem novas tentativas.
ExecutarComRetentativa(descricao, app, fn, tentativas := 3) {
    Loop tentativas {
        try {
            return fn.Call()
        } catch as err {
            if (A_Index >= tentativas || !EhErroDeOffice(err))
                throw err
            LogErro(err, descricao " — tentativa " A_Index " de " tentativas ", tentando de novo")
            if EhErroQuedaOffice(err) {
                if (app = "excel")
                    ReiniciarExcel()
                else
                    ReiniciarPowerPoint()
            } else {
                Sleep 1500
            }
        }
    }
}

OfficeRespondendo(app) {
    if !IsObject(app)
        return false
    Loop 5 {
        try {
            versao := app.Version
            return true
        } catch as err {
            if !EhErroOcupadoOffice(err)
                return false
            Sleep 500
        }
    }
    return true  ; ocupado, mas vivo
}

PidDoOffice(app, nomeExe) {
    hwnd := 0
    try hwnd := app.Hwnd
    if hwnd {
        anterior := A_DetectHiddenWindows
        DetectHiddenWindows true
        pid := 0
        try pid := WinGetPID("ahk_id " hwnd)
        DetectHiddenWindows anterior
        if pid
            return pid
    }
    ; Plano B: a verificação pré-rodagem garantiu que não havia outro aberto.
    return ProcessExist(nomeExe)
}

; Espera o processo do script terminar sozinho; se não terminar, encerra à
; força — mas só se o PID ainda for mesmo do programa esperado.
EncerrarProcessoDoScript(pid, nomeExe) {
    if !pid
        return
    Loop 10 {
        if !ProcessExist(pid)
            return
        Sleep 500
    }
    nomeAtual := ""
    try nomeAtual := ProcessGetName(pid)
    if (StrUpper(nomeAtual) = nomeExe) {
        RegistrarLog(nomeExe " (processo " pid ") não fechou sozinho — encerrando à força.")
        try ProcessClose(pid)
    }
}

GarantirExcel() {
    global excel, EXCEL_PID
    if OfficeRespondendo(excel)
        return excel
    if IsObject(excel) {
        RegistrarLog("O Excel parou de responder — abrindo um novo.")
        ReiniciarExcel()
    }
    excel := ComObject("Excel.Application")
    excel.Visible := false
    excel.DisplayAlerts := false
    excel.AskToUpdateLinks := false
    EXCEL_PID := PidDoOffice(excel, "EXCEL.EXE")
    versao := ""
    try versao := excel.Version
    RegistrarLog("Excel iniciado (versão " versao ", processo " EXCEL_PID ").")
    return excel
}

ReiniciarExcel() {
    global excel, EXCEL_PID
    if IsObject(excel)
        try excel.Quit()
    excel := ""
    EncerrarProcessoDoScript(EXCEL_PID, "EXCEL.EXE")
    EXCEL_PID := 0
}

FinalizarExcel() {
    if (!IsObject(excel) && !EXCEL_PID)
        return
    RegistrarLog("Fechando o Excel do script.")
    ReiniciarExcel()
}

GarantirPowerPoint() {
    global ppt, PPT_PID
    if OfficeRespondendo(ppt)
        return ppt
    if IsObject(ppt) {
        RegistrarLog("O PowerPoint parou de responder — abrindo um novo.")
        ReiniciarPowerPoint()
    }
    ppt := ComObject("PowerPoint.Application")
    ppt.Visible := true
    PPT_PID := PidDoOffice(ppt, "POWERPNT.EXE")
    versao := ""
    try versao := ppt.Version
    RegistrarLog("PowerPoint iniciado (versão " versao ", processo " PPT_PID ").")
    return ppt
}

ReiniciarPowerPoint() {
    global ppt, PPT_PID
    if IsObject(ppt)
        try ppt.Quit()
    ppt := ""
    EncerrarProcessoDoScript(PPT_PID, "POWERPNT.EXE")
    PPT_PID := 0
}

FinalizarPowerPoint() {
    if (!IsObject(ppt) && !PPT_PID)
        return
    RegistrarLog("Fechando o PowerPoint do script.")
    ReiniciarPowerPoint()
}

; ─── LEITURA DAS PLANILHAS (abre, lê pra memória, fecha) ────────────────

LerAbasDaPlanilha(caminho) {
    xl := GarantirExcel()
    wbX := ""
    try {
        wbX := AbrirPastaDeTrabalho(xl, caminho)
        return ListarNomesAbas(wbX)
    } finally {
        if IsObject(wbX)
            try wbX.Close(false)
    }
}

; Devolve a lista de cursos da aba do semestre, ou "" se a aba não existir.
LerCursosDaMestre(caminho, nomeAba) {
    xl := GarantirExcel()
    wbX := ""
    try {
        wbX := AbrirPastaDeTrabalho(xl, caminho)
        wsM := ObterAba(wbX, nomeAba)
        if !wsM
            return ""

        lista := []
        ultimaLinha := wsM.Cells(wsM.Rows.Count, 2).End(-4162).Row
        Loop ultimaLinha - 1 {
            linha := A_Index + 1
            coordenacao := Trim(wsM.Cells(linha, 1).Value)
            posgrad     := Trim(wsM.Cells(linha, 2).Value)
            codigoTurma := Trim(wsM.Cells(linha, 3).Value)
            siglaUnid   := ExtrairSiglaUnidade(codigoTurma)
            unidadeExib := (siglaUnid != "" && siglaUnid != codigoTurma) ? siglaUnid " (" codigoTurma ")" : codigoTurma
            turma       := Trim(wsM.Cells(linha, 4).Value)

            if (posgrad = "")
                continue

            lista.Push(Map(
                "linha", linha,
                "coordenacao", coordenacao,
                "posgrad", posgrad,
                "unidade", codigoTurma,
                "siglaUnidade", siglaUnid,
                "unidadeExib", unidadeExib,
                "turma", turma
            ))
        }
        return lista
    } finally {
        if IsObject(wbX)
            try wbX.Close(false)
    }
}

; Lê a A2 (nome da pós) e TODAS as linhas das abas escolhidas, com a mesma
; LerRegistro de antes, e devolve tudo em memória:
;   Map("nomePos", ..., "dadosAbas", [Map("aba", nome, "registros", [...]), ...])
; Linhas de rodapé/intervalo já ficam de fora (mesma regra da Fase 3 antiga).
LerDadosDaPlanilha(caminho, abas) {
    xl := GarantirExcel()
    wbX := ""
    try {
        wbX := AbrirPastaDeTrabalho(xl, caminho)

        nomePos := ""
        wsPrimeira := ObterAba(wbX, abas[1])
        if wsPrimeira
            nomePos := Trim(wsPrimeira.Range("A2").Value)

        dadosAbas := []
        for abaNome in abas {
            wsAba := ObterAba(wbX, abaNome)
            if !wsAba
                throw Error("Aba '" abaNome "' não encontrada na planilha.")

            cols := DetectarColunasPorAba(wsAba)
            ultimaLinhaCurso := wsAba.Cells(wsAba.Rows.Count, 1).End(-4162).Row
            registros := []
            Loop ultimaLinhaCurso - 4 {
                linha := A_Index + 4
                reg := LerRegistro(wsAba, linha, cols["colNome"], cols["colCurso"], cols["colEmail"])
                if reg["ehRodapeIgnoravel"]
                    continue
                reg["aba"] := abaNome
                registros.Push(reg)
            }
            RegistrarLog("Aba " abaNome ": " registros.Length " linhas lidas (colunas: nome=" cols["colNome"] ", aula=" cols["colCurso"] ", e-mail=" cols["colEmail"] ").")
            dadosAbas.Push(Map("aba", abaNome, "registros", registros))
        }
        return Map("nomePos", nomePos, "dadosAbas", dadosAbas)
    } finally {
        if IsObject(wbX)
            try wbX.Close(false)
    }
}

AbaNaLista(listaAbas, nomeAba) {
    for nome in listaAbas {
        if (Trim(nome) = nomeAba)
            return true
    }
    return false
}

; ─── CORREÇÕES MANUAIS DA FASE 2 (aplicadas no registro em memória) ─────

; Mesma limpeza de nome que a LerRegistro faz ao ler a célula.
CorrigirNomeRegistro(reg, valor) {
    nome := Trim(valor)
    nome := RegExReplace(nome, "\s*\b[HMhm]\b\s*$")
    nome := Trim(nome)
    reg["nomeOriginalBruto"] := nome
    nome := RegExReplace(nome, "\s*\([^)]*\)", "")
    nome := RegExReplace(nome, "i)^(dra|dr|profa|professora|professor|prof|sra|sr)\.?\s+", "")
    nome := Trim(RegExReplace(nome, "\s+", " "))
    reg["nome"] := nome
    reg["erroNome"] := (nome = "")
    reg["suspeitaNomeEmail"] := InStr(nome, "@") ? true : false
    reg["suspeitaNomeComExtra"] := false
}

CorrigirCursoRegistro(reg, valor) {
    curso := Trim(valor)
    reg["curso"] := curso
    reg["erroCurso"] := (curso = "")
    reg["suspeitaCursoData"] := (curso != "" && RegExMatch(curso, "^\d{1,2}[/:h]\d")) ? true : false
}

CorrigirDataRegistro(reg, valor) {
    dataTxt := Trim(valor)
    if RegExMatch(dataTxt, "^(\d{1,2})/(\d{1,2})/(\d{2,4})$", &m) {
        ano := (StrLen(m[3]) = 2) ? "20" m[3] : m[3]
        dataTxt := Format("{:02}/{:02}/{}", Integer(m[1]), Integer(m[2]), ano)
    }
    if EhDataValida(dataTxt) {
        reg["dataAula"] := dataTxt
        reg["erroData"] := false
    } else {
        RegistrarLog("Data digitada inválida ('" valor "') — linha " reg["linha"] " continua com erro e será ignorada.")
    }
}

CorrigirHorasRegistro(reg, duracaoFormatada) {
    reg["horasFormatadas"] := duracaoFormatada
    reg["horarioTxt"] := duracaoFormatada
    reg["erroHoras"] := false
}
