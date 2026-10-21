; Instalador do Certificados (Inno Setup 6).
; Não rodar direto: use "python instalador\montar.py", que monta o programa
; e chama o Inno Setup passando a versão (/DVersao=...).
;
; Instala só para o usuário (sem senha de administrador) em
; %LOCALAPPDATA%\Programs\Certificados, cria o atalho no menu Iniciar e,
; se a pessoa deixar marcado, na Área de Trabalho. Aparece em
; Configurações > Aplicativos para desinstalar. Instalar uma versão nova por
; cima atualiza; o e-mail de teste e o cadastro de coordenadores ficam.

#ifndef Versao
  #define Versao "0.0.0"
#endif

[Setup]
AppId={{8C1F6E2A-4B1D-4C55-9E3A-2F7B5D0C9A11}
AppName=Certificados
AppVersion={#Versao}
AppVerName=Certificados {#Versao}
AppPublisher=Certificados da Pós-graduação
DefaultDirName={localappdata}\Programs\Certificados
DefaultGroupName=Certificados
DisableProgramGroupPage=yes
DisableDirPage=yes
PrivilegesRequired=lowest
OutputDir=..\dist_instalador
OutputBaseFilename=Instalar-Certificados-{#Versao}
SetupIconFile=..\icone.ico
UninstallDisplayIcon={app}\Certificados.exe
UninstallDisplayName=Certificados
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
CloseApplications=yes

[Languages]
Name: "pt"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"

[Tasks]
Name: "atalho"; Description: "Criar atalho na Área de Trabalho"; GroupDescription: "Atalhos:"

[InstallDelete]
; Versão nova por cima: tira as bibliotecas da versão anterior antes de copiar.
Type: filesandordirs; Name: "{app}\_internal"

[Files]
Source: "..\dist_programa\Certificados\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\Certificados"; Filename: "{app}\Certificados.exe"; AppUserModelID: "Certificados.PosGraduacao"
Name: "{autodesktop}\Certificados"; Filename: "{app}\Certificados.exe"; Tasks: atalho; AppUserModelID: "Certificados.PosGraduacao"

[Run]
Filename: "{app}\Certificados.exe"; Description: "Abrir o Certificados agora"; Flags: nowait postinstall skipifsilent
