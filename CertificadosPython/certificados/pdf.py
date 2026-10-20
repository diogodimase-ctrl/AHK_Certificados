"""
.pptx -> .pdf pelo PowerPoint instalado. O PowerPoint só abre o arquivo já
pronto e salva como PDF (o preenchimento é todo feito pelo python-pptx).

Mesmo cuidado da v24 do AHK: se o PowerPoint cair ou estiver ocupado, tenta
de novo (até 3 vezes), reabrindo-o quando preciso. Só fecha o PowerPoint no
fim se ele não estava aberto antes do programa começar.
"""

import os
import subprocess
import time

# 0x800706BA/BE/BF e 0x80010108/0114: o PowerPoint caiu ou sumiu.
_QUEDA = {0x800706BA, 0x800706BE, 0x800706BF, 0x80010108, 0x80010114}
# 0x80010001 e 0x8001010A: está ocupado, vale esperar um pouco.
_OCUPADO = {0x80010001, 0x8001010A}


def _codigo(erro):
    hr = getattr(erro, "hresult", None)
    if hr is None and getattr(erro, "args", None) and isinstance(erro.args[0], int):
        hr = erro.args[0]
    return (hr or 0) & 0xFFFFFFFF


def descrever_erro(erro):
    codigo = _codigo(erro)
    return f"0x{codigo:08X} {erro}" if codigo else str(erro)


def powerpoint_aberto():
    try:
        saida = subprocess.run(["tasklist", "/FI", "IMAGENAME eq POWERPNT.EXE"], capture_output=True, text=True,
                               creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return "POWERPNT.EXE" in saida.stdout.upper()
    except OSError:
        return False


class ConversorPowerPoint:
    nome = "PowerPoint deste computador"

    def __init__(self, log=print, tentativas=3):
        self.log = log
        self.tentativas = tentativas
        self.app = None
        self._estava_aberto = False

    def __enter__(self):
        import comtypes
        comtypes.CoInitialize()
        self._estava_aberto = powerpoint_aberto()
        self._abrir()
        return self

    def __exit__(self, *_):
        try:
            if self.app is not None and not self._estava_aberto and self.app.Presentations.Count == 0:
                self.app.Quit()
        except Exception:
            pass
        self.app = None
        import comtypes
        comtypes.CoUninitialize()

    def _abrir(self):
        import comtypes.client
        self.app = comtypes.client.CreateObject("PowerPoint.Application")

    def converter(self, pptx, pdf):
        for tentativa in range(1, self.tentativas + 1):
            pres = None
            try:
                pres = self.app.Presentations.Open(os.path.abspath(pptx), True, False, False)
                pres.SaveAs(os.path.abspath(pdf), 32)      # 32 = PDF
                pres.Close()
                return
            except Exception as e:
                if pres is not None:
                    try:
                        pres.Close()
                    except Exception:
                        pass
                if os.path.exists(pdf):
                    try:
                        os.remove(pdf)
                    except OSError:
                        pass
                codigo = _codigo(e)
                if tentativa >= self.tentativas or codigo not in _QUEDA | _OCUPADO:
                    raise
                self.log(f"PowerPoint: {descrever_erro(e)} — tentativa {tentativa} de {self.tentativas}, tentando de novo")
                if codigo in _QUEDA:
                    self._abrir()
                else:
                    time.sleep(1.5)
