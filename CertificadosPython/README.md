# Certificados em Python

Versão em Python do `GerarCertificados` (AHK v27). Em construção, por etapas:

1. **Motor** (pronto, em conferência): lê as planilhas, aplica as mesmas regras do AHK,
   preenche o modelo `.pptx` sem abrir o PowerPoint, gera o PDF e o `manifesto.txt`
   igual ao da v27 (o fluxo do Power Automate não muda).
2. **Programa novo**: `programa.py`. Uma janela com as etapas numa faixa no alto
   (Início, Coordenadores, Planilhas e meses, Conferir dados, Gerar e Relatório), no
   estilo "Escuro violeta" (`tema.py` tem todas as cores e fontes; as peças desenhadas
   à mão ficam em `componentes.py`). A tela de Início junta o painel do mês (% de
   cursos e de certificados feitos), a tabela de cursos da mestre com cara de Excel
   (uma barrinha por mês do semestre: verde enviado, amarela na Para Enviar, vermelha
   não feito) e a data de emissão, que só aparece depois de marcar um curso. Sem curso
   marcado, o lugar da data mostra o tempo economizado (`certificados/economia.py`,
   11 min 33 s por certificado no jeito manual, guardado só neste computador).
   Modo (Teste/Oficial), e-mail de teste e arquivos ficam no alto da janela.
   Meses já enviados: `certificados/historico.py` (histórico em
   `Declarações\_Programa Certificados\historico.jsonl` mais as pastas arquivadas pelo
   fluxo). Certificados previstos no mês: `certificados/painel.py` (lê a planilha de
   cada curso; contagem guardada em `%APPDATA%\CertificadosPython\contagem.json`).
   O e-mail de teste fica em `%APPDATA%\CertificadosPython\config.json`.
3. **Instalador** (`instalador/`): instala sem senha de administrador, com atalho na
   Área de Trabalho e no menu Iniciar. Cada versão é publicada na página Releases do
   GitHub, e o programa avisa quando há uma versão nova (`atualizacao.py`).

## Publicar uma versão nova

1. Aumentar `VERSAO` em `certificados/__init__.py`.
2. `python instalador\montar.py` (precisa do Inno Setup 6) — gera
   `dist_instalador\Instalar-Certificados-<versão>.exe`.
3. `gh release create v<versão> dist_instalador\Instalar-Certificados-<versão>.exe --title "Certificados <versão>" --notes "..."`.

## Pastas

- `certificados/` — o motor. Cada função cita a função AHK equivalente.
- `conferencia.py` — programa que compara o motor com o AHK usando os arquivos reais
  (modelo, simulação e comparação dos PDFs). Empacotado em `ConferenciaCertificados.zip`.
- `testes/` — testes com arquivos de mentira (`python testes/testar_motor.py --pdf`)
  e comparação do modelo com o PowerPoint (`python testes/comparar_com_powerpoint.py`).

## Modos

Oficial e teste são o mesmo programa (no AHK eram arquivos espelhados):

- `oficial`: `C:\Certificados` + `Para Enviar\<Curso>\<Mês>`.
- `teste`: `TESTE - ` no nome, `C:\Certificados_TESTE`, e-mails trocados pelo de teste.
- `simulacao`: só `C:\Certificados_SIMULACAO`, nada vai para a nuvem.

## Diferenças conhecidas em relação ao AHK

- Coordenador com título na planilha mestre ("Dra. Ana"): o AHK imprime
  "Dra. Dra. Ana"; aqui sai "Dra. Ana" (`textos.sem_titulo`).
- Conferir dados: linhas só com aviso começam como "gerar" (no AHK, "pular
  todas" pulava também os avisos); linhas com erro começam como "pular".

- Data da aula escrita num formato que não é dd/mm/aaaa: o AHK soma um dia
  (`ExcelSerialParaData` conta a partir de 01/01/1900). Aqui a data é a do Excel.
