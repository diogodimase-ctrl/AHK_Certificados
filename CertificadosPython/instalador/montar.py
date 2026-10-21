"""
Monta o instalador do Certificados.

Rodar (de qualquer pasta), com o Python que tem as bibliotecas do programa:
    python instalador\\montar.py

1. PyInstaller monta o programa numa pasta (dist_programa\\Certificados), que
   abre mais rápido que o .exe único de antes;
2. o Inno Setup (instalador\\certificados.iss) empacota essa pasta em
   dist_instalador\\Instalar-Certificados-<versão>.exe.

Depois, para publicar a versão no GitHub (página Releases):
    gh release create v<versão> dist_instalador\\Instalar-Certificados-<versão>.exe
"""

import os
import shutil
import subprocess
import sys

AQUI = os.path.dirname(os.path.abspath(__file__))
PROJETO = os.path.dirname(AQUI)
sys.path.insert(0, PROJETO)

from certificados import VERSAO  # noqa: E402


def iscc():
    candidatos = [os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Inno Setup 6", "ISCC.exe"),
                  os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Inno Setup 6", "ISCC.exe"),
                  os.path.join(os.environ.get("ProgramFiles", ""), "Inno Setup 6", "ISCC.exe")]
    for c in candidatos:
        if os.path.isfile(c):
            return c
    sys.exit("Inno Setup 6 não encontrado (winget install JRSoftware.InnoSetup).")


def main():
    trabalho = os.path.join(PROJETO, "build_programa")
    saida = os.path.join(PROJETO, "dist_programa")
    shutil.rmtree(os.path.join(saida, "Certificados"), ignore_errors=True)
    subprocess.run([sys.executable, "-m", "PyInstaller", "--noconfirm", "--onedir", "--windowed",
                    "--name", "Certificados", "--icon", os.path.join(PROJETO, "icone.ico"),
                    "--add-data", f"{os.path.join(PROJETO, 'icone.ico')};.",
                    "--collect-all", "customtkinter", "--hidden-import", "comtypes.client",
                    "--paths", PROJETO, "--distpath", saida, "--workpath", trabalho, "--specpath", trabalho,
                    os.path.join(PROJETO, "programa.py")], check=True, cwd=PROJETO)
    subprocess.run([iscc(), f"/DVersao={VERSAO}", os.path.join(AQUI, "certificados.iss")], check=True)
    print("\nInstalador:", os.path.join(PROJETO, "dist_instalador", f"Instalar-Certificados-{VERSAO}.exe"))


if __name__ == "__main__":
    main()
