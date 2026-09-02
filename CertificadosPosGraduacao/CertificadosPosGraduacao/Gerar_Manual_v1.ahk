#Requires AutoHotkey v2.0
#SingleInstance Force
#Include UI_Moderna.ahk

; ════════════════════════════════════════════════════════════════════════
;   GERADOR MANUAL DE CERTIFICADOS / CORREÇÃO AVULSA
;   Permite emitir ou corrigir qualquer certificado de forma direta e rápida
; ════════════════════════════════════════════════════════════════════════

global INFO_NUVEM       := ObterCaminhosCompartilhados()
global PASTA_SAIDA_BASE := "C:\Certificados\"
global ARQUIVO_LOGO     := ""

; Tenta localizar o logo automaticamente
for logoCandidate in [
    "C:\CAssinaturas\_logo_einstein.jpg",
    "C:\CAssinaturas\_logo_einstein.png",
    INFO_NUVEM["pastaDeclaracoes"] . "_logo_einstein.jpg"
] {
    if (logoCandidate != "" && FileExist(logoCandidate)) {
        ARQUIVO_LOGO := logoCandidate
        break
    }
}

ExibirFormularioManual()

ExibirFormularioManual() {
    largura := 660
    hoje := FormatTime(A_Now, "dd/MM/yyyy")

    g := Gui("+AlwaysOnTop -MaximizeBox", "Emissão Manual de Certificado — Correção")
    g.SetFont("s10", "Segoe UI")
    g.BackColor := "F7F9FC"

    ; ─── HEADER ────────────────────────────────────────────────────────
    g.SetFont("s14 bold", "Segoe UI")
    g.Add("Text", "x20 y15 w" (largura - 40) " c003366", "✍️ Emissão Manual de Certificado / Correção")

    g.SetFont("s9 norm", "Segoe UI")
    g.Add("Text", "x20 y45 w" (largura - 40) " c555555", "Preencha os campos abaixo para gerar ou corrigir um certificado individualmente.")

    g.Add("Text", "x20 y68 w" (largura - 40) " h2 0x10")

    ; ─── SEÇÃO 1: DADOS DO DOCENTE E AULA ──────────────────────────────
    g.SetFont("s10 bold", "Segoe UI")
    g.Add("GroupBox", "x20 y80 w" (largura - 40) " h210 c003366", " 🎓 Dados do Certificado e Aula ")

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x40 y105 w130", "Nome do Docente:")
    g.SetFont("s10 norm", "Segoe UI")
    edtNome := g.Add("Edit", "x180 y102 w420", "")

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x40 y145 w130", "Pós-Graduação:")
    g.SetFont("s10 norm", "Segoe UI")
    edtPos := g.Add("Edit", "x180 y142 w420", "Pós-Graduação em ")

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x40 y185 w130", "Aula / Disciplina:")
    g.SetFont("s10 norm", "Segoe UI")
    edtAula := g.Add("Edit", "x180 y182 w420", "")

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x40 y225 w90", "Data da Aula:")
    g.SetFont("s10 norm", "Segoe UI")
    edtDataAula := g.Add("Edit", "x135 y222 w105", hoje)

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x250 y225 w65", "Semestre:")
    g.SetFont("s10 norm", "Segoe UI")
    ddlSemestre := g.Add("DropDownList", "x320 y222 w105 Choose1", ["2026-5", "2026-6", "2025-5", "2025-6", "2027-1", "2026-1S", "2026-2S"])

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x435 y225 w50", "Horas:")
    g.SetFont("s10 norm", "Segoe UI")
    edtHoras := g.Add("Edit", "x490 y222 w110", "10 horas")

    g.SetFont("s8 italic", "Segoe UI")
    g.Add("Text", "x135 y255 w465 c777777", "Exemplos: '10 horas', '08h - 18h', '2 horas e 30 minutos'")

    ; ─── SEÇÃO 2: COORDENADOR E ASSINATURA ─────────────────────────────
    g.SetFont("s10 bold", "Segoe UI")
    g.Add("GroupBox", "x20 y300 w" (largura - 40) " h190 c003366", " ✍️ Dados do Coordenador e Assinatura ")

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x40 y325 w130", "Coordenador(a):")
    g.SetFont("s10 norm", "Segoe UI")
    edtCoord := g.Add("Edit", "x180 y322 w420", "")

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x40 y365 w130", "Título / Prefixo:")
    g.SetFont("s9 norm", "Segoe UI")
    ddlTitulo := g.Add("DropDownList", "x180 y362 w130 Choose1", ["(Nenhum)", "Dr. ", "Dra. ", "Prof. Dr. ", "Profa. Dra. ", "Prof. ", "Profa. "])

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x330 y365 w60", "Cargo:")
    g.SetFont("s9 norm", "Segoe UI")
    ddlCargo := g.Add("DropDownList", "x390 y362 w210 Choose1", ["Coordenador do Curso", "Coordenadora do Curso", "Coordenador(a) do Curso"])

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x40 y405 w130", "Assinatura (.png/.jpg):")
    g.SetFont("s9 norm", "Segoe UI")
    edtAssin := g.Add("Edit", "x180 y402 w310 ReadOnly", "")
    btnBuscarAssin := g.Add("Button", "x500 y401 w100 h28", "📁 Localizar")

    g.SetFont("s9 bold", "Segoe UI")
    g.Add("Text", "x40 y445 w130", "Data de Emissão:")
    g.SetFont("s10 norm", "Segoe UI")
    edtDataEmissao := g.Add("Edit", "x180 y442 w130", hoje)

    ; ─── SEÇÃO 3: OPÇÕES E BOTÕES ──────────────────────────────────────
    g.SetFont("s9 norm", "Segoe UI")
    chkAbrirPdf := g.Add("Checkbox", "x40 y505 w280 Checked", "👁️ Abrir o PDF gerado automaticamente")
    chkSalvarNuvem := g.Add("Checkbox", "x330 y505 w280 Checked", "☁️ Salvar no SharePoint na pasta oficial do Curso")

    g.SetFont("s10 bold", "Segoe UI")
    btnGerar := g.Add("Button", "Default w200 h42 x20 y545", "🚀 Gerar Certificado")
    btnLimpar := g.Add("Button", "w120 h42 x+10", "🔄 Limpar")
    btnCancelar := g.Add("Button", "w120 h42 x+10", "❌ Fechar")

    ; ─── EVENTOS ───────────────────────────────────────────────────────
    btnBuscarAssin.OnEvent("Click", (*) => EscolherAssinatura(edtAssin))

    btnLimpar.OnEvent("Click", (*) => LimparCampos(edtNome, edtPos, edtAula, edtCoord, edtAssin))

    btnCancelar.OnEvent("Click", (*) => g.Destroy())
    g.OnEvent("Close", (*) => g.Destroy())
    g.OnEvent("Escape", (*) => g.Destroy())

    btnGerar.OnEvent("Click", (*) => ProcessarGeracaoManual(
        g,
        edtNome.Text,
        edtPos.Text,
        edtAula.Text,
        edtDataAula.Text,
        ddlSemestre.Text,
        edtHoras.Text,
        edtCoord.Text,
        ddlTitulo.Text,
        ddlCargo.Text,
        edtAssin.Text,
        edtDataEmissao.Text,
        chkAbrirPdf.Value,
        chkSalvarNuvem.Value
    ))

    g.Show("w" largura " h600")
}

EscolherAssinatura(ctrlAssin) {
    pastaIni := ""
    if (INFO_NUVEM["pastaAssinaturas"] != "" && DirExist(INFO_NUVEM["pastaAssinaturas"]))
        pastaIni := INFO_NUVEM["pastaAssinaturas"]
    else if DirExist("C:\CAssinaturas\")
        pastaIni := "C:\CAssinaturas\"
    arq := FileSelect(1, pastaIni, "Selecione a Imagem da Assinatura", "Imagens (*.png; *.jpg; *.jpeg)")
    if arq
        ctrlAssin.Text := arq
}

LimparCampos(edtNome, edtPos, edtAula, edtCoord, edtAssin) {
    edtNome.Text  := ""
    edtPos.Text   := "Pós-Graduação em "
    edtAula.Text  := ""
    edtCoord.Text := ""
    edtAssin.Text := ""
}

; ─── Sanitiza nome para uso como pasta/arquivo ──────────────────────────
SanitNome(nome) {
    limpo := RegExReplace(nome, '[\\/:*?"<>|]', "")
    limpo := RegExReplace(limpo, "[\x00-\x1F]", "")
    limpo := Trim(limpo, " `t`r`n.")
    return (limpo = "") ? "Sem_Nome" : limpo
}

; ─── Cria pasta sem erro se já existir ──────────────────────────────────
CriarPasta(caminho) {
    try DirCreate(caminho)
}

; ─── Localiza ou cria a pasta do curso e do mês no SharePoint ─────────
LocalizarPastaDestinoSharePoint(pastaDeclaracoes, semestre, nomeCurso, mesPasta) {
    pastaSemestre := pastaDeclaracoes . semestre . "\"
    if !DirExist(pastaSemestre) {
        variacoes := ObterEquivalenciasSemestre(semestre)
        for varSem in variacoes {
            if DirExist(pastaDeclaracoes . varSem . "\") {
                pastaSemestre := pastaDeclaracoes . varSem . "\"
                break
            }
        }
    }
    CriarPasta(pastaSemestre)

    nomeCursoNorm := NormalizarTexto(nomeCurso)
    nomeCursoSemPrefixo := RegExReplace(nomeCursoNorm, "i)^pos\s*graduacao\s*(em|de)?\s*", "")
    pastaCurso := ""

    Loop Files, pastaSemestre . "*", "D" {
        pastaNorm := NormalizarTexto(A_LoopFileName)
        pastaSemPrefixo := RegExReplace(pastaNorm, "i)^pos\s*graduacao\s*(em|de)?\s*", "")
        if (pastaNorm = nomeCursoNorm || (nomeCursoSemPrefixo != "" && pastaSemPrefixo = nomeCursoSemPrefixo)) {
            pastaCurso := A_LoopFilePath . "\"
            break
        }
    }

    if (pastaCurso = "") {
        pastaCurso := pastaSemestre . SanitNome(nomeCurso) . "\"
        CriarPasta(pastaCurso)
    }

    pastaMes := pastaCurso . mesPasta . "\"
    CriarPasta(pastaMes)

    return pastaMes
}

ProcessarGeracaoManual(guiObj, nome, posgrad, aula, dataAula, semestreEscolhido, horas, coord, titulo, cargo, assinaturaPath, dataEmissao, abrirPdf, salvarNuvem) {
    nome          := Trim(nome)
    posgrad       := Trim(posgrad)
    aula          := Trim(aula)
    dataAula      := Trim(dataAula)
    horas         := Trim(horas)
    coord         := Trim(coord)
    dataEmissao   := Trim(dataEmissao)
    assinaturaPath := Trim(assinaturaPath)

    ; ── Validações básicas ──────────────────────────────────────────────
    erros := ""
    if (nome = "")
        erros .= "  • Nome do Docente`n"
    if (posgrad = "" || posgrad = "Pós-Graduação em " || posgrad = "Pós-Graduação em")
        erros .= "  • Nome da Pós-Graduação`n"
    if (aula = "")
        erros .= "  • Aula / Disciplina`n"
    if (dataAula = "")
        erros .= "  • Data da Aula`n"
    if (coord = "")
        erros .= "  • Nome do Coordenador`n"

    if (erros != "") {
        ExibirMensagem("Campos Obrigatórios", "Preencha os Campos Obrigatórios", "Os seguintes campos estão em branco:`n`n" . erros, "erro")
        return
    }

    ; ── Formatação do Coordenador ───────────────────────────────────────
    prefixoFinal  := (titulo = "(Nenhum)") ? "" : titulo
    coordCompleto := prefixoFinal . coord

    ; ── Formatação de horas ─────────────────────────────────────────────
    horasFormatadas := (horas != "") ? FormatarHorasSeguro(horas) : "10 horas"
    if InStr(horasFormatadas, "invalido")
        horasFormatadas := horas  ; usa o texto original se não reconhecido

    ; ── Localização do Modelo PowerPoint ────────────────────────────────
    pptModelo := ""
    if (INFO_NUVEM["modeloPptx"] != "" && FileExist(INFO_NUVEM["modeloPptx"]))
        pptModelo := INFO_NUVEM["modeloPptx"]

    if (pptModelo = "") {
        pastaIni := (INFO_NUVEM["pastaDeclaracoes"] != "") ? INFO_NUVEM["pastaDeclaracoes"] : ""
        pptModelo := FileSelect(1, pastaIni, "Selecione o Modelo PowerPoint (.pptx)", "*.pptx")
        if !pptModelo {
            ExibirMensagem("Modelo Não Encontrado", "Selecione o Modelo PPTX", "Nenhum modelo PowerPoint foi selecionado. Operação cancelada.", "erro")
            return
        }
    }

    ; ── Extração do Mês da Pasta (ex: "04.26") ──────────────────────────
    mesPasta := "01.26"
    if RegExMatch(dataAula, "^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$", &mDt) {
        numMes := Format("{:02}", Integer(mDt[2]))
        anoStr := mDt[3]
        anoCurto := (StrLen(anoStr) = 4) ? SubStr(anoStr, 3, 2) : anoStr
        mesPasta := numMes . "." . anoCurto
    }

    ; ── Definição de pastas de saída (Nuvem e Local) ────────────────────
    nomePosSanit    := SanitNome(posgrad)
    nomeDocSanit    := SanitNome(nome)
    dataSanit       := RegExReplace(dataAula, '[\\/:*?"<>|]', ".")
    nomeBaseArquivo := nomeDocSanit . " (" . dataSanit . ").pdf"

    ; 1) Pasta no SharePoint (Semestre\Curso\Mês\)
    caminhoNuvemPdf := ""
    if (salvarNuvem && INFO_NUVEM["pastaDeclaracoes"] != "") {
        pastaNuvemCursoMes := LocalizarPastaDestinoSharePoint(INFO_NUVEM["pastaDeclaracoes"], semestreEscolhido, posgrad, mesPasta)
        caminhoNuvemPdf := pastaNuvemCursoMes . nomeBaseArquivo
    }

    ; 2) Pastas de Backup Local (C:\Certificados\<Curso>\<Mês>\ e \Docentes\<Nome>\)
    pastaLocalCursoMes := PASTA_SAIDA_BASE . nomePosSanit . "\" . mesPasta . "\"
    pastaLocalDoc      := PASTA_SAIDA_BASE . nomePosSanit . "\Docentes\" . nomeDocSanit . "\"
    CriarPasta(pastaLocalCursoMes)
    CriarPasta(pastaLocalDoc)

    caminhoLocalPdf := pastaLocalCursoMes . nomeBaseArquivo
    caminhoDocPdf   := pastaLocalDoc   . nomeBaseArquivo

    ; ── GERAÇÃO VIA POWERPOINT COM ──────────────────────────────────────
    ppt := ""
    try {
        ppt := ComObject("PowerPoint.Application")
        ppt.Visible := true
        try ppt.WindowState := 2  ; Minimiza a janela do PowerPoint para não atrapalhar

        ; 1. Cria o modelo temporário com remoção da assinatura âncora e do logo antigo
        modeloTemp := CriarModeloBaseManual(ppt, pptModelo, coordCompleto, cargo, dataEmissao, assinaturaPath, ARQUIVO_LOGO)

        ; 2. Abre o modelo temp e substitui dados do docente
        pres  := ppt.Presentations.Open(modeloTemp)
        slide := pres.Slides(1)

        Loop slide.Shapes.Count {
            shape := slide.Shapes.Item(A_Index)
            try {
                txt := shape.TextFrame.TextRange.Text
                if !InStr(txt, "{NOME}") && !InStr(txt, "{POSGRADUACAO}") && !InStr(txt, "{CURSO}") && !InStr(txt, "{DATA}") && !InStr(txt, "{HORAS}")
                    continue
                SubstituirPlaceholderManual(shape, "{NOME}",         nome)
                SubstituirPlaceholderManual(shape, "{POSGRADUACAO}", posgrad)
                SubstituirPlaceholderManual(shape, "{CURSO}",        aula)
                SubstituirPlaceholderManual(shape, "{DATA}",         dataAula)
                SubstituirPlaceholderManual(shape, "{HORAS}",        horasFormatadas)
            } catch {
            }
        }

        ; 3. Exporta PDF
        pres.SaveAs(caminhoLocalPdf, 32)
        pres.Close()

        ; 4. Cópias
        try FileCopy(caminhoLocalPdf, caminhoDocPdf, true)
        if (caminhoNuvemPdf != "")
            try FileCopy(caminhoLocalPdf, caminhoNuvemPdf, true)

        ; 5. Limpa temporário
        try FileDelete(modeloTemp)
        try ppt.Quit()

        ; ─── SUCESSO ───────────────────────────────────────────────────
        msg := "Certificado gerado com sucesso!`n`n"
            . "👤 Docente: " . nome . "`n"
            . "🎓 Curso: " . posgrad . "`n"
            . "📖 Aula: " . aula . "`n"
            . "📅 Data: " . dataAula . "  |  ⏰ Horas: " . horasFormatadas . "`n"
            . "✍️ Coordenador: " . coordCompleto . " (" . cargo . ")`n`n"
            . "📁 Salvo em:`n" . caminhoLocalPdf
            . (caminhoNuvemPdf != "" ? "`n`n☁️ Cópia no SharePoint:`n" . caminhoNuvemPdf : "")

        ExibirMensagem("Certificado Gerado", "Certificado Gerado com Sucesso!", msg, "sucesso", "OK", 620)

        if abrirPdf && FileExist(caminhoLocalPdf)
            Run('"' . caminhoLocalPdf . '"')

    } catch as err {
        try ppt.Quit()
        ExibirMensagem("Erro na Geração", "Falha ao Gerar o Certificado", "Ocorreu um erro durante o processo:`n`n" . err.Message, "erro")
    }
}

; ─── Substitui placeholder preservando fonte/cor/tamanho ────────────────
SubstituirPlaceholderManual(shape, placeholder, valor, forcaSize := 0, forcaBold := -1) {
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

; ─── Cria modelo base removendo explicitamente imagem de assinatura antiga e logo ───
CriarModeloBaseManual(ppt, pptModelo, coordUsado, cargo, dataEmissao, assinaturaPath, logoPath) {
    modeloBase := A_Temp "\certificado_manual_base_" A_TickCount "_" Random(1000, 9999) ".pptx"
    presBase  := ppt.Presentations.Open(pptModelo)
    slideBase := presBase.Slides(1)

    ; 1. Substitui textos do coordenador e data de emissão
    Loop slideBase.Shapes.Count {
        shape := slideBase.Shapes.Item(A_Index)
        try {
            txt := shape.TextFrame.TextRange.Text
            if InStr(txt, "{COORDENADOR}") || InStr(txt, "{CARGO}") || InStr(txt, "{DATAEMISSAO}") {
                SubstituirPlaceholderManual(shape, "{COORDENADOR}", coordUsado, 10, 0)
                SubstituirPlaceholderManual(shape, "{CARGO}",       cargo,      10, 1)
                SubstituirPlaceholderManual(shape, "{DATAEMISSAO}", dataEmissao)
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

    ; 2. Coleta todas as imagens no slide
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

    ; 3. Deleta o logo antigo
    logoTop := 0
    if imgLogo {
        logoTop := imgLogo["top"]
        try imgLogo["shape"].Delete()
    }

    larguraPadrao := 130
    alturaPadrao  := 50

    ; 4. Deleta a assinatura âncora antiga e insere a nova
    if imgAssina {
        posLeft := imgAssina["left"]
        posTop  := imgAssina["top"]
        try imgAssina["shape"].Delete()

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

        if (assinaturaPath != "" && FileExist(assinaturaPath)) {
            picAssinatura := slideBase.Shapes.AddPicture(assinaturaPath, false, true, novoLeft, novoTop, larguraPadrao, alturaPadrao)
            picAssinatura.Name := "ASSINATURA_REAL"
        }
    } else if (assinaturaPath != "" && FileExist(assinaturaPath)) {
        picAssinatura := slideBase.Shapes.AddPicture(assinaturaPath, false, true, 200, 380, larguraPadrao, alturaPadrao)
        picAssinatura.Name := "ASSINATURA_REAL"
    }

    ; 5. Insere o logotipo oficial Einstein
    LOGO_LARGURA := 320
    LOGO_ALTURA  := 200
    if (logoPath != "" && FileExist(logoPath)) {
        slideW := presBase.PageSetup.SlideWidth
        novoLogoLeft := (slideW - LOGO_LARGURA) / 2
        novoLogoTop  := (logoTop > 0) ? logoTop : 20
        picLogo := slideBase.Shapes.AddPicture(logoPath, false, true, novoLogoLeft, novoLogoTop, LOGO_LARGURA, LOGO_ALTURA)
        picLogo.Name := "LOGO_EINSTEIN"
    }

    presBase.SaveAs(modeloBase, 24)
    presBase.Close()
    return modeloBase
}
