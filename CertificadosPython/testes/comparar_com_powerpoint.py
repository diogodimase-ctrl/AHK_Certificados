"""
Compara o modelo-base feito pelo Python com o feito do jeito do AHK no
PowerPoint (certificados/referencia_ahk.py). Precisa do PowerPoint.

Rodar: python testes/comparar_com_powerpoint.py [modelo.pptx assinatura logo]
Sem argumentos usa o modelo de mentira de testes/amostras.py.
"""

import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from certificados import referencia_ahk  # noqa: E402


def main():
    if len(sys.argv) >= 4:
        caminho_modelo, assinatura, logo = sys.argv[1:4]
    else:
        from testes import amostras
        locais = amostras.criar_ambiente(tempfile.mkdtemp())
        caminho_modelo, logo = locais.modelo, locais.logo
        assinatura = os.path.join(locais.assinaturas, "Ana Beatriz Souza.png")
    linhas, pior, _ = referencia_ahk.comparar_modelo(
        caminho_modelo, "Dra. Ana Beatriz Souza", "Coordenadora do Curso", "20/10/2026", assinatura, logo)
    print(f"{'forma':42} {'AHK (esq, topo, larg, alt)':34} Python")
    for chave, va, vp, ok in linhas:
        print(f"{chave[:42]:42} {str(va):34} {vp}   {'ok' if ok else 'DIFERENTE'}")
    print(f"\nMaior diferença: {pior:.1f} pt")
    return 0 if pior <= referencia_ahk.TOLERANCIA_PT else 1


if __name__ == "__main__":
    sys.exit(main())
