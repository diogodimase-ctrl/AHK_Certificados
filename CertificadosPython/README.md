# Certificados em Python

Versão em Python do `GerarCertificados` (AHK v27). Em construção, por etapas:

1. **Motor** (pronto, em conferência): lê as planilhas, aplica as mesmas regras do AHK,
   preenche o modelo `.pptx` sem abrir o PowerPoint, gera o PDF e o `manifesto.txt`
   igual ao da v27 (o fluxo do Power Automate não muda).
2. Interface nova (próxima etapa).
3. Programa para distribuir para a equipe.

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

## Diferença conhecida em relação ao AHK

- Data da aula escrita num formato que não é dd/mm/aaaa: o AHK soma um dia
  (`ExcelSerialParaData` conta a partir de 01/01/1900). Aqui a data é a do Excel.
