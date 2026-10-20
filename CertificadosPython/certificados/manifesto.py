"""
manifesto.txt (conteúdo JSON) lido pelo fluxo do Power Automate.

Os nomes e a ordem dos campos são os mesmos da v27 do AHK: o fluxo já está
validado com eles. Não mudar sem mudar o fluxo junto.
"""

import json
import random
from datetime import datetime


def montar(curso, local, caminho_pasta_pdf, caminho_excel, nome_aba, docentes,
           semestre, pasta_mes, pasta_pos, pasta_destino, arquivos_do_lote, agora=None):
    """ConstruirManifestoJson. "docentes" = lista de dicts {nome, email, arquivos}."""
    agora = agora or datetime.now()
    dados = {
        "versaoManifesto": 1,
        "loteId": f"{agora:%Y%m%d%H%M%S}-{random.randint(1000, 9999)}",
        "dataGeracao": f"{agora:%Y-%m-%dT%H:%M:%S}",
        "curso": curso,
        "local": local,
        "caminhoPastaPDF": caminho_pasta_pdf,
        "caminhoArquivoExcel": caminho_excel,
        "nomeAba": nome_aba,
        "semestre": semestre,
        "pastaMes": pasta_mes,
        "pastaPos": pasta_pos,
        "pastaDestino": pasta_destino,
        "arquivosDoLote": list(arquivos_do_lote),
        "certificados": [
            {
                "docente": d["nome"],
                "email": d["email"],
                "arquivos": [{"nomeArquivo": a["nomeArquivo"], "chaveLinha": a["chaveLinha"]} for a in d["arquivos"]],
            }
            for d in docentes
        ],
    }
    return json.dumps(dados, ensure_ascii=False, indent=2)


def gravar(caminho, texto):
    """EscreverArquivoDireto: grava direto com o nome final (sem .tmp), em
    UTF-8 sem BOM. O gatilho do fluxo não dispara com renomeação."""
    with open(caminho, "w", encoding="utf-8", newline="\n") as f:
        f.write(texto)
