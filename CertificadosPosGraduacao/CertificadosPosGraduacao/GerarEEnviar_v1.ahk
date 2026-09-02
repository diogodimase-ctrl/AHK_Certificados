#Requires AutoHotkey v2.0
#SingleInstance Force
#Include UI_Moderna.ahk

; ════════════════════════════════════════════════════════════════════════
;   GERAR E ENVIAR — LANÇADOR SEQUENCIAL
;   Executa a esteira completa: Geração na Nuvem/Local + Envio por E-mail
; ════════════════════════════════════════════════════════════════════════

global NOME_SCRIPT_GERAR  := "GerarCertificados_v16.ahk"
global NOME_SCRIPT_ENVIAR := "EnviarEmails_v11.ahk"

pastaScripts  := A_ScriptDir
caminhoGerar  := pastaScripts "\" NOME_SCRIPT_GERAR
caminhoEnviar := pastaScripts "\" NOME_SCRIPT_ENVIAR

if !FileExist(caminhoGerar) {
    if FileExist(A_Desktop "\" NOME_SCRIPT_GERAR)
        caminhoGerar := A_Desktop "\" NOME_SCRIPT_GERAR
    else if FileExist(A_Desktop "\CertificadosPosGraduacao\CertificadosPosGraduacao\" NOME_SCRIPT_GERAR)
        caminhoGerar := A_Desktop "\CertificadosPosGraduacao\CertificadosPosGraduacao\" NOME_SCRIPT_GERAR
    else {
        ExibirMensagem(
            "Arquivo de Geração Não Encontrado",
            "Localizar GerarCertificados",
            "Não foi possível localizar o arquivo '" NOME_SCRIPT_GERAR "'.`n`nLocalize-o na próxima janela.",
            "erro",
            "Localizar Arquivo ➔"
        )
        caminhoGerar := FileSelect(1, pastaScripts, "Selecione o arquivo de GERAÇÃO (.ahk)", "*.ahk")
        if !caminhoGerar
            ExitApp
    }
}

if !FileExist(caminhoEnviar) {
    if FileExist(A_Desktop "\" NOME_SCRIPT_ENVIAR)
        caminhoEnviar := A_Desktop "\" NOME_SCRIPT_ENVIAR
    else if FileExist(A_Desktop "\CertificadosPosGraduacao\CertificadosPosGraduacao\" NOME_SCRIPT_ENVIAR)
        caminhoEnviar := A_Desktop "\CertificadosPosGraduacao\CertificadosPosGraduacao\" NOME_SCRIPT_ENVIAR
    else {
        ExibirMensagem(
            "Arquivo de Envio Não Encontrado",
            "Localizar EnviarEmails",
            "Não foi possível localizar o arquivo '" NOME_SCRIPT_ENVIAR "'.`n`nLocalize-o na próxima janela.",
            "erro",
            "Localizar Arquivo ➔"
        )
        caminhoEnviar := FileSelect(1, pastaScripts, "Selecione o arquivo de ENVIO (.ahk)", "*.ahk")
        if !caminhoEnviar
            ExitApp
    }
}

opcao := EscolherOpcaoInicialUI()
if (opcao = "cancelar")
    ExitApp

; ─── ETAPA 1: GERAÇÃO DE CERTIFICADOS ───
if (opcao = "gerar_e_enviar" || opcao = "so_gerar") {
    ExibirMensagem(
        "Passo 1 de 2 — Gerar Certificados",
        "Iniciando Módulo de Geração",
        "O programa de geração de certificados será iniciado agora.`n`n"
        "Siga as etapas até o Relatório Final. Quando ele concluir, esta esteira continuará automaticamente.",
        "passo",
        "Iniciar Geração ➔",
        580
    )
    RunWait('"' A_AhkPath '" "' caminhoGerar '"')

    if (opcao = "so_gerar") {
        ExibirMensagem("Geração Concluída", "Processo Concluído", "Geração finalizada com sucesso!", "sucesso", "Fechar")
        ExitApp
    }

    respContinuar := Confirmar(
        "Continuar para o Envio?",
        "Geração de Certificados Concluída",
        "A geração de certificados foi finalizada!`n`n"
        "Deseja iniciar agora o ENVIO DOS E-MAILS para os docentes desta mesma rodada?",
        "✉️ Sim, Iniciar Envio",
        "⏹️ Não, Enviar Depois",
        true,
        580
    )
    if !respContinuar {
        ExibirMensagem("Processo Finalizado", "Envio Pausado", "Quando desejar enviar, basta acionar o botão de envio no Painel de Controle.", "info", "Fechar")
        ExitApp
    }
}

; ─── ETAPA 2: ENVIO DE E-MAILS ───
ExibirMensagem(
    "Passo 2 de 2 — Enviar E-mails",
    "Iniciando Módulo de Envio",
    "O assistente de envio de e-mails via Outlook será iniciado agora.",
    "passo",
    "Iniciar Envio ➔",
    580
)
RunWait('"' A_AhkPath '" "' caminhoEnviar '"')

ExibirMensagem(
    "Fluxo Completo Concluído",
    "Esteira Integrada Finalizada",
    "Todas as etapas da esteira integrada foram executadas com sucesso!",
    "sucesso",
    "Concluir",
    580
)
ExitApp

EscolherOpcaoInicialUI(largura := 580) {
    resultado := "cancelar"

    g := Gui("+AlwaysOnTop -MaximizeBox", "Esteira Integrada — Gerar e Enviar")
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F7F9FC"

    g.SetFont("s13 bold", "Segoe UI")
    g.Add("Text", "x20 y16 w" (largura - 40) " c003366", "🚀 Esteira Integrada de Certificados")

    g.Add("Text", "x20 y48 w" (largura - 40) " h2 0x10")

    g.SetFont("s10 norm", "Segoe UI")
    g.Add("Text", "x20 y58 w" (largura - 40) " c222222",
        "Selecione o fluxo desejado para a operação de hoje:`n`n"
        "• Gerar e Enviar: Executa a geração em lote na nuvem/local e já abre o envio.`n"
        "• Só Gerar: Executa exclusivamente a geração dos certificados em PDF.`n"
        "• Só Enviar: Executa exclusivamente o assistente de disparo de e-mails."
    )

    g.SetFont("s10 bold", "Segoe UI")
    btnAmbos := g.Add("Button", "Default w200 h40 x20 y160", "🚀 Gerar e Enviar (Tudo)")
    btnGerar := g.Add("Button", "w160 h40 x+10", "🎓 Só Gerar")
    btnEnviar := g.Add("Button", "w160 h40 x+10", "✉️ Só Enviar")
    
    g.SetFont("s9 norm", "Segoe UI")
    btnCancelar := g.Add("Button", "w" (largura - 40) " h30 x20 y215", "Cancelar")

    btnAmbos.OnEvent("Click", (*) => (resultado := "gerar_e_enviar", g.Destroy()))
    btnGerar.OnEvent("Click", (*) => (resultado := "so_gerar", g.Destroy()))
    btnEnviar.OnEvent("Click", (*) => (resultado := "so_enviar", g.Destroy()))
    btnCancelar.OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show("w" largura)
    WinWaitClose("ahk_id " g.Hwnd)

    return resultado
}
