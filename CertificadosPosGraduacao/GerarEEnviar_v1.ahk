#Requires AutoHotkey v2.0
#SingleInstance Force

; ════════════════════════════════════════════════════════════════════════
;   GERAR E ENVIAR — v1 — LANÇADOR (roda os dois programas em sequência)
; ════════════════════════════════════════════════════════════════════════
;
;   Este script NÃO substitui os outros dois — ele só chama um e depois o
;   outro, na ordem certa, para você não precisar abrir cada um na mão.
;
;   IMPORTANTE: este arquivo precisa ficar na MESMA PASTA que:
;     - GerarCertificados_v16.ahk
;     - EnviarEmails_v11.ahk
; ════════════════════════════════════════════════════════════════════════

NOME_SCRIPT_GERAR  := "GerarCertificados_v16.ahk"
NOME_SCRIPT_ENVIAR := "EnviarEmails_v11.ahk"

pastaScripts  := A_ScriptDir
caminhoGerar  := pastaScripts "\" NOME_SCRIPT_GERAR
caminhoEnviar := pastaScripts "\" NOME_SCRIPT_ENVIAR

; --- confere se os dois scripts estão na mesma pasta; se não, deixa você localizar ---
if !FileExist(caminhoGerar) {
    MsgBox(
        "Não encontrei o arquivo `"" NOME_SCRIPT_GERAR "`" nesta pasta:`n`n" pastaScripts "`n`n"
        "Na próxima tela, localize esse arquivo manualmente.",
        "Arquivo de Geração Não Encontrado", "Iconi"
    )
    caminhoGerar := FileSelect(1, pastaScripts, "Selecione o arquivo de GERAÇÃO de certificados (.ahk)", "*.ahk")
    if !caminhoGerar {
        MsgBox("Nenhum arquivo selecionado. Programa encerrado.")
        ExitApp
    }
}
if !FileExist(caminhoEnviar) {
    MsgBox(
        "Não encontrei o arquivo `"" NOME_SCRIPT_ENVIAR "`" nesta pasta:`n`n" pastaScripts "`n`n"
        "Na próxima tela, localize esse arquivo manualmente.",
        "Arquivo de Envio Não Encontrado", "Iconi"
    )
    caminhoEnviar := FileSelect(1, pastaScripts, "Selecione o arquivo de ENVIO de e-mails (.ahk)", "*.ahk")
    if !caminhoEnviar {
        MsgBox("Nenhum arquivo selecionado. Programa encerrado.")
        ExitApp
    }
}

; --- tela inicial: o que vai acontecer ---
opcao := EscolherOpcaoInicial()
if (opcao = "cancelar")
    ExitApp

; --- executa a geração (se aplicável) ---
if (opcao = "gerar_e_enviar" || opcao = "so_gerar") {
    MsgBox(
        "Passo 1 de 2 — Gerar Certificados`n`n"
        "Vou abrir agora o programa de geração de certificados, normalmente, "
        "como se você tivesse aberto ele sozinho.`n`n"
        "Siga as telas dele até o fim (Relatório Final). Quando ele terminar "
        "e a janela dele fechar, este lançador continua sozinho.",
        "Passo 1 de 2 — Gerar Certificados", "Iconi"
    )
    RunWait('"' A_AhkPath '" "' caminhoGerar '"')

    if (opcao = "so_gerar") {
        MsgBox("Geração concluída (ou encerrada). Este lançador não vai abrir o envio, como você pediu.", "Concluído")
        ExitApp
    }

    respContinuar := MsgBox(
        "Geração de certificados concluída (ou encerrada).`n`n"
        "Deseja seguir agora para o ENVIO dos e-mails com os certificados "
        "que acabaram de ser gerados?`n`n"
        "(Clicando em Não, o envio não é aberto agora — você pode rodar o "
        "EnviarEmails manualmente depois, quando quiser.)",
        "Continuar para o Envio?", "YesNo Iconi"
    )
    if (respContinuar = "No") {
        MsgBox("Certo! Quando quiser enviar, é só rodar este lançador de novo ou abrir o EnviarEmails diretamente.", "Encerrado")
        ExitApp
    }
}

; --- executa o envio ---
MsgBox(
    "Passo 2 de 2 — Enviar E-mails`n`n"
    "Vou abrir agora o programa de envio de e-mails, normalmente. Siga as "
    "telas dele até o fim (Relatório Final da Rodada).",
    "Passo 2 de 2 — Enviar E-mails", "Iconi"
)
RunWait('"' A_AhkPath '" "' caminhoEnviar '"')

MsgBox(
    "Processo concluído!`n`n"
    "Os certificados foram gerados e o envio dos e-mails foi executado "
    "(ou encerrado) em seguida.",
    "Processo Concluído", "Iconi"
)
ExitApp

; ════════════════════════════════════════════════════════════════════════
;   Tela inicial com 3 botões claros, em vez de Sim/Não genérico
; ════════════════════════════════════════════════════════════════════════

EscolherOpcaoInicial() {
    resultado := "cancelar"

    g := Gui("+AlwaysOnTop", "Gerar e Enviar Certificados")
    g.SetFont("s10")
    g.Add("Text", "w480",
        "GERAR E ENVIAR CERTIFICADOS`n`n"
        "O que você quer fazer agora?`n`n"
        "• Gerar e Enviar — faz os dois passos em sequência: primeiro gera "
        "os certificados, depois já pergunta se você quer seguir direto "
        "para o envio dos e-mails.`n`n"
        "• Só Gerar — abre somente o programa de geração de certificados.`n`n"
        "• Só Enviar — abre somente o programa de envio de e-mails (use "
        "quando os certificados já foram gerados antes)."
    )

    btnAmbos := g.Add("Button", "Default w220 y+15", "Gerar e Enviar (tudo)")
    btnGerar := g.Add("Button", "x+10 w130", "Só Gerar")
    btnEnviar := g.Add("Button", "x+10 w130", "Só Enviar")
    btnCancelar := g.Add("Button", "w480 y+10", "Cancelar")

    btnAmbos.OnEvent("Click", (*) => (resultado := "gerar_e_enviar", g.Destroy()))
    btnGerar.OnEvent("Click", (*) => (resultado := "so_gerar", g.Destroy()))
    btnEnviar.OnEvent("Click", (*) => (resultado := "so_enviar", g.Destroy()))
    btnCancelar.OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    g.Show()
    WinWaitClose("ahk_id " g.Hwnd)

    return resultado
}
