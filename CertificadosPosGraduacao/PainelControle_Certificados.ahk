#Requires AutoHotkey v2.0
#SingleInstance Force
#Include UI_Moderna.ahk

; ════════════════════════════════════════════════════════════════════════
;   PAINEL DE CONTROLE — CERTIFICADOS PÓS-GRADUAÇÃO
;   Central de Execução, Diagnóstico e Acesso Rápido com Integração Nuvem
; ════════════════════════════════════════════════════════════════════════

; Configuração dinâmica de caminhos (Nuvem e Local)
global INFO_NUVEM        := ObterCaminhosCompartilhados()
global PASTA_CERTS       := "C:\Certificados\"
global PASTA_ASSINATURAS := (INFO_NUVEM["pastaAssinaturas"] != "" && DirExist(INFO_NUVEM["pastaAssinaturas"])) ? INFO_NUVEM["pastaAssinaturas"] : "C:\CAssinaturas\"
global PASTA_SHAREPOINT  := (INFO_NUVEM["pastaDeclaracoes"] != "") ? INFO_NUVEM["pastaDeclaracoes"] : INFO_NUVEM["raizOneDrive"]

; Nomes padrão dos scripts
global SCRIPT_GERAR_PADRAO  := "GerarCertificados_v16.ahk"
global SCRIPT_ENVIAR_PADRAO := "EnviarEmails_v11.ahk"
global SCRIPT_MANUAL_PADRAO := "Gerar_Manual_v1.ahk"
global SCRIPT_EXTRA_PADRAO  := "Gerar_Extraordinarios_v2.ahk"
global SCRIPT_AMBOS_PADRAO  := "GerarEEnviar_v1.ahk"

; Inicialização e construção da Interface
CriarInterfacePrincipal()

CriarInterfacePrincipal() {
    g := Gui("+AlwaysOnTop -MaximizeBox", "Central de Certificados — Pós-Graduação")
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F7F9FC"

    ; ─── CABEÇALHO / HEADER ─────────────────────────────────────────────
    g.SetFont("s15 bold", "Segoe UI")
    g.Add("Text", "x20 y15 w640 c003366", "🎓 Central de Certificados de Pós-Graduação")
    
    g.SetFont("s9", "Segoe UI")
    g.Add("Text", "x20 y45 w640 c555555", "Painel de controle para geração em lote de certificados e envio automático de e-mails aos docentes.")
    
    g.Add("Text", "x20 y65 w640 h2 0x10") ; Linha separadora horizontal

    ; ─── GRUPO 1: MÓDULOS DE EXECUÇÃO ──────────────────────────────────
    g.SetFont("s10 bold", "Segoe UI")
    g.Add("GroupBox", "x20 y80 w640 h315 c003366", " ⚡ Módulos de Execução ")

    ; Botão 1: Fluxo Completo (Destaque)
    g.SetFont("s10 bold", "Segoe UI")
    btnAmbos := g.Add("Button", "x40 y105 w280 h50 Default", "🚀 Fluxo Completo`n(Gerar + Enviar)")
    g.SetFont("s9 norm", "Segoe UI")
    g.Add("Text", "x335 y110 w300 h45 c444444", "Gera os certificados selecionados na nuvem/local e já abre a esteira de envio de e-mails.")

    ; Botão 2: Gerar Certificados (Lote)
    g.SetFont("s10 bold", "Segoe UI")
    btnGerar := g.Add("Button", "x40 y165 w280 h42", "🎓 1. Gerar Certificados (Lote)")
    g.SetFont("s9 norm", "Segoe UI")
    g.Add("Text", "x335 y170 w300 h38 c444444", "Lê a planilha mestre do SharePoint, aplica assinaturas e gera os PDFs na nuvem e local.")

    ; Botão 3: Enviar E-mails
    g.SetFont("s10 bold", "Segoe UI")
    btnEnviar := g.Add("Button", "x40 y215 w280 h42", "✉️ 2. Enviar E-mails aos Docentes")
    g.SetFont("s9 norm", "Segoe UI")
    g.Add("Text", "x335 y220 w300 h38 c444444", "Abre os e-mails no Outlook e prepara os anexos com controle de envio (_EnvioControle.txt).")

    ; Botão 4: Certificado Manual / Correção
    g.SetFont("s10 bold", "Segoe UI")
    btnManual := g.Add("Button", "x40 y265 w280 h42", "✏️ 3. Certificado Manual (Correção)")
    g.SetFont("s9 norm", "Segoe UI")
    g.Add("Text", "x335 y270 w300 h38 c444444", "Emite ou altera certificados individuais preenchendo os dados diretamente.")

    ; Botão 5: Certificados Extraordinários
    g.SetFont("s10 bold", "Segoe UI")
    btnExtra := g.Add("Button", "x40 y315 w280 h40", "📋 4. Certificados Extraordinários")
    g.SetFont("s9 norm", "Segoe UI")
    g.Add("Text", "x335 y320 w300 h38 c444444", "Assistente para planilhas avulsas ou fora do padrão comum.")

    ; ─── GRUPO 2: ATALHOS RÁPIDOS ───────────────────────────────────────
    g.SetFont("s10 bold", "Segoe UI")
    g.Add("GroupBox", "x20 y405 w640 h90 c003366", " 📁 Acesso Rápido a Pastas e Nuvem ")

    g.SetFont("s9", "Segoe UI")
    btnPastaNuvem := g.Add("Button", "x35 y435 w145 h45", "☁️ SharePoint Einstein`n(Declarações)")
    btnPastaCerts := g.Add("Button", "x190 y435 w145 h45", "📂 Backup Local`n(C:\Certificados)")
    btnPastaAssin := g.Add("Button", "x345 y435 w145 h45", "✍️ Assinaturas`n(Coordenadores)")
    btnPastaScripts := g.Add("Button", "x500 y435 w145 h45", "📁 Pasta Scripts`n(Arquivos .ahk)")

    ; ─── BARRA DE STATUS / INFORMAÇÕES ──────────────────────────────────
    g.SetFont("s9 italic", "Segoe UI")
    statusTxt := g.Add("Text", "x20 y510 w480 h35 c333333", "Verificando ambiente...")

    g.SetFont("s9 norm", "Segoe UI")
    btnAjuda := g.Add("Button", "x510 y505 w70 h30", "❓ Ajuda")
    btnSair  := g.Add("Button", "x590 y505 w70 h30", "✖ Fechar")

    ; Lista de botões para controle de bloqueio
    todosBotoes := [btnAmbos, btnGerar, btnEnviar, btnManual, btnExtra]

    ; ─── EVENTOS DOS BOTÕES ─────────────────────────────────────────────
    btnAmbos.OnEvent("Click", (*) => ExecutarModulo("ambos", g, statusTxt, todosBotoes))
    btnGerar.OnEvent("Click", (*) => ExecutarModulo("gerar", g, statusTxt, todosBotoes))
    btnEnviar.OnEvent("Click", (*) => ExecutarModulo("enviar", g, statusTxt, todosBotoes))
    btnManual.OnEvent("Click", (*) => ExecutarModulo("manual", g, statusTxt, todosBotoes))
    btnExtra.OnEvent("Click", (*) => ExecutarModulo("extra", g, statusTxt, todosBotoes))

    btnPastaNuvem.OnEvent("Click", (*) => AbrirDiretorio(PASTA_SHAREPOINT))
    btnPastaCerts.OnEvent("Click", (*) => AbrirDiretorio(PASTA_CERTS))
    btnPastaAssin.OnEvent("Click", (*) => AbrirDiretorio(PASTA_ASSINATURAS))
    btnPastaScripts.OnEvent("Click", (*) => AbrirDiretorio(A_ScriptDir))

    btnAjuda.OnEvent("Click", (*) => ExibirGuiaAjuda())
    btnSair.OnEvent("Click", (*) => ExitApp())
    g.OnEvent("Close", (*) => ExitApp())

    ; Atualiza status inicial
    AtualizarStatusAmbiente(statusTxt)

    g.Show("w680 h550")
}

; ────────────────────────────────────────────────────────────────────────
;   FUNÇÃO PARA LOCALIZAR E EXECUTAR OS SCRIPTS COM TRATAMENTO DE ERRO
; ────────────────────────────────────────────────────────────────────────

LocalizarScript(nomeArquivo) {
    ; Locais de busca prioritários
    candidatos := [
        A_ScriptDir "\" nomeArquivo,
        A_ScriptDir "\CertificadosPosGraduacao\" nomeArquivo,
        A_ScriptDir "\CertificadosPosGraduacao\CertificadosPosGraduacao\" nomeArquivo,
        A_Desktop "\" nomeArquivo,
        A_Desktop "\CertificadosPosGraduacao\CertificadosPosGraduacao\" nomeArquivo,
        A_Desktop "\AHK\" nomeArquivo,
        A_Desktop "\Gerar Certificados\Curso normal\" nomeArquivo,
        A_Desktop "\Gerar Certificados\Envio\" nomeArquivo,
        A_Desktop "\Gerar Certificados\Certificado extraordinário\" nomeArquivo
    ]

    for caminho in candidatos {
        if FileExist(caminho)
            return caminho
    }

    ; Fallback se for extraordinários
    if (nomeArquivo = SCRIPT_EXTRA_PADRAO) {
        fallbackExtra := [
            A_Desktop "\GerarCertificadosExtraordinarios (1).ahk",
            A_Desktop "\AHK\GerarCertificadosExtraordinarios (1).ahk",
            A_Desktop "\Gerar Certificados\Certificado extraordinário\GerarCertificadosExtraordinarios (1).ahk",
            A_Desktop "\teste.ahk"
        ]
        for f in fallbackExtra {
            if FileExist(f)
                return f
        }
    }

    return ""
}

ExecutarModulo(tipo, guiObj, statusCtrl, botoes) {
    nomeScript := ""
    tituloModulo := ""

    switch tipo {
        case "gerar":
            nomeScript := SCRIPT_GERAR_PADRAO
            tituloModulo := "Geração de Certificados"
        case "enviar":
            nomeScript := SCRIPT_ENVIAR_PADRAO
            tituloModulo := "Envio de E-mails"
        case "manual":
            nomeScript := SCRIPT_MANUAL_PADRAO
            tituloModulo := "Certificado Manual (Correção)"
        case "extra":
            nomeScript := SCRIPT_EXTRA_PADRAO
            tituloModulo := "Certificados Extraordinários"
        case "ambos":
            nomeScript := SCRIPT_AMBOS_PADRAO
            tituloModulo := "Fluxo Completo (Gerar + Enviar)"
    }

    caminho := LocalizarScript(nomeScript)

    if (caminho = "") {
        MsgBox(
            "⚠️ Não foi possível localizar o arquivo:`n`n" nomeScript "`n`n"
            "Na próxima tela, você poderá selecionar o arquivo manualmente.",
            "Arquivo Não Encontrado", "Icon!"
        )
        caminho := FileSelect(1, A_ScriptDir, "Selecione o arquivo " nomeScript, "*.ahk")
        if !caminho
            return
    }

    ; Desabilita botões temporariamente para evitar duplo clique acidental
    for b in botoes
        b.Enabled := false

    statusCtrl.Text := "⏳ Executando: " tituloModulo "..."
    
    ; Minimiza temporariamente a interface para dar foco ao módulo executado
    guiObj.Minimize()

    try {
        RunWait('"' A_AhkPath '" "' caminho '"')
    } catch as err {
        MsgBox("Erro ao iniciar o script:`n`n" err.Message, "Erro de Execução", "IconX")
    }

    ; Restaura a interface ao término
    guiObj.Restore()
    for b in botoes
        b.Enabled := true

    AtualizarStatusAmbiente(statusCtrl)
}

; ────────────────────────────────────────────────────────────────────────
;   UTILITÁRIOS: STATUS DO AMBIENTE E ABERTURA DE PASTAS
; ────────────────────────────────────────────────────────────────────────

AtualizarStatusAmbiente(statusCtrl) {
    nuvemOk := (INFO_NUVEM["raizOneDrive"] != "") ? "☁️ OneDrive Conectado" : "⚠️ OneDrive não localizado"
    certsOk := DirExist(PASTA_CERTS) ? "✅ C:\Certificados" : "⚠️ C:\Certificados"
    
    statusCtrl.Text := "Ambiente: " nuvemOk "  |  " certsOk
}

AbrirDiretorio(caminho) {
    if !DirExist(caminho) {
        resposta := MsgBox(
            "A pasta não existe atualmente:`n`n" caminho "`n`n"
            "Deseja criá-la agora?",
            "Pasta Não Encontrada", "YesNo Icon?"
        )
        if (resposta = "Yes") {
            try {
                DirCreate(caminho)
            } catch as err {
                MsgBox("Erro ao criar pasta: " err.Message, "Erro", "IconX")
                return
            }
        } else {
            return
        }
    }
    Run('explorer.exe "' caminho '"')
}

; ────────────────────────────────────────────────────────────────────────
;   GUIA DE AJUDA / MANUAL RÁPIDO
; ────────────────────────────────────────────────────────────────────────

ExibirGuiaAjuda() {
    ajudaTexto := 
    (
        "GUIA RÁPIDO — FLUXO DE TRABALHO`n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━`n`n"
        "1. GERAR CERTIFICADOS (LOTE):`n"
        "   • Abre a planilha Mestre ('Controle de Declaração de Aula').`n"
        "   • Escolha os cursos e o período desejado.`n"
        "   • Os certificados serão gerados em PDF na pasta C:\Certificados\`n`n"
        "2. ENVIAR E-MAILS:`n"
        "   • Abre o assistente de envio para os cursos gerados.`n"
        "   • Valida e-mails dos docentes e permite correções.`n"
        "   • Abre o Outlook Web com assunto/corpo prontos e exibe os PDFs.`n"
        "   • Salva o progresso em _EnvioControle.txt para nunca duplicar envios.`n`n"
        "3. FLUXO COMPLETO:`n"
        "   • Executa o Gerador e, ao terminar, segue direto para o Envio.`n`n"
        "4. CERTIFICADOS EXTRAORDINÁRIOS:`n"
        "   • Para planilhas avulsas que não seguem o formato padrão.`n"
        "   • Permite configurar colunas de docente, data, hora e assinatura na hora."
    )
    MsgBox(ajudaTexto, "Ajuda — Central de Certificados", "Iconi")
}
