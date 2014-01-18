; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "tbrEmergencyGroup"
#define MyAppVerName "tbrEmergencyGroup 2.3"
#define MyAppPublisher "tbrSoft Internacional"
#define MyAppURL "http://www.tbrsoft.com"
#define MyAppExeName "tbrEmergencyGroup.EXE"

#define fullInst true
;incluye base de datos MDAC, framework, etc
;con false es solo una actualizacion

[Setup]
AppName={#MyAppName}
AppVerName={#MyAppVerName}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={pf}\{#MyAppName}
;DisableDirPage=yes
DefaultGroupName=tbrSoft
AllowNoIcons=no
#if fullInst
  OutputBaseFilename=Instalar_tbrEmergencyGroup
#else
  OutputBaseFilename=Update_tbrEmergencyGroup
#endif
SetupIconFile=estrella.ico
Compression=lzma
SolidCompression=yes

;lenguajes del instalador
[Languages]
Name: "eng"; MessagesFile: "compiler:Default.isl"
Name: "bra"; MessagesFile: "compiler:Languages\BrazilianPortuguese.isl"
Name: "cat"; MessagesFile: "compiler:Languages\Catalan.isl"
Name: "cze"; MessagesFile: "compiler:Languages\Czech.isl"
Name: "dan"; MessagesFile: "compiler:Languages\Danish.isl"
Name: "dut"; MessagesFile: "compiler:Languages\Dutch.isl"
Name: "fin"; MessagesFile: "compiler:Languages\Finnish.isl"
Name: "fre"; MessagesFile: "compiler:Languages\French.isl"
Name: "ger"; MessagesFile: "compiler:Languages\German.isl"
Name: "hun"; MessagesFile: "compiler:Languages\Hungarian.isl"
Name: "ita"; MessagesFile: "compiler:Languages\Italian.isl"
Name: "nor"; MessagesFile: "compiler:Languages\Norwegian.isl"
Name: "pol"; MessagesFile: "compiler:Languages\Polish.isl"
Name: "por"; MessagesFile: "compiler:Languages\Portuguese.isl"
Name: "rus"; MessagesFile: "compiler:Languages\Russian.isl"
Name: "slo"; MessagesFile: "compiler:Languages\Slovenian.isl"
Name: "spa"; MessagesFile: "compiler:Languages\Spanish.isl"
;Name: "spa"; MessagesFile: "compiler:Languages\SpanishMex.isl"
;Name: "spa"; MessagesFile: "compiler:Languages\SpanishEsp.isl"

;accesos directos
[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "quicklaunchicon"; Description: "{cm:CreateQuickLaunchIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
; cosas que el artime hizo ... se entiende ...
#if fullInst
  Name: MDAC; Description: Microsoft Data Access Components; Flags: checkedonce restart; MinVersion: 4.0,0
  Name: JET; Description: Microsoft JET Components; Flags: checkedonce restart; OnlyBelowVersion: 5,4.1
#endif
[Files]
; esto son los runtimes de visual basic y cosas basicas. Mirando los ultimos flags ver�n que se instala condicionalmente si esta en 98 xp o vista
; NOTE: Don't use "Flags: ignoreversion" on any shared system files
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\stdole2.tlb";  DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regtypelib; minversion: 5.0,5.0; onlybelowversion: 5.2,5.2
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\asycfilt.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall; minversion: 5.0,5.0
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\olepro32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver; minversion: 5.0,5.0
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\oleaut32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver; minversion: 5.0,5.0; onlybelowversion: 5.2,5.2
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\msvbvm60.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\mswinsck.ocx"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver  sharedfile
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\richtx32.ocx"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver  sharedfile

Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\hhctrl.ocx"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall regserver  sharedfile

; archivos sueltos que se necesiten
; "..\" hace que el path sea relativo a la carpeta anterior
;la segunda parte crea la carpeta "ayuda" dentro de la carpeta de la alpicaci�n si no existiera y coloca alli el archivo indicado
; el flag ignoreversi�n indica que se reemplazara sin mirar nada en el mismo archivo
;Source: "..\datos\ayuda.txt"; DestDir: "{app}\datos\ayuda"; Flags: ignoreversion
;Source: "..\imagenes\logo1.jpg"; DestDir: "{app}\datos\imagenes"; Flags: ignoreversion
;Source: "..\manual.doc"; DestDir: "{app}"; Flags: ignoreversion
;pueden ir ejecutables adicionales por ejemplo (mas abajo se les crean accesos directorios opcionalmente)
;Source: "..\repair.exe"; DestDir: "{app}"; Flags: ignoreversion
;Source: "..\Manual\Manual de Usuario.doc"; DestDir: "{app}\Manual"; Flags: ignoreversion

#if fullInst
  Source: "..\ALCemi\ayuda.chm"; DestDir: "{app}\bin\";
  Source: "..\ALCemi\formularios\fondorecibo.bmp"; DestDir: "{app}\bin\formularios";
  Source: "..\fondos\Emergencias\fondo01-800x600.jpg"; DestDir: "{app}\fondos\Emergencias";
  Source: "..\fondos\Emergencias\fondo01-1024x768.jpg"; DestDir: "{app}\fondos\Emergencias";
  Source: "..\fondos\Emergencias\fondo01-1280x960.jpg"; DestDir: "{app}\fondos\Emergencias";

  Source: "..\fondos\Bomberos\fondo01-800x600.jpg"; DestDir: "{app}\fondos\Bomberos";
  Source: "..\fondos\Bomberos\fondo01-1024x768.jpg"; DestDir: "{app}\fondos\Bomberos";
  Source: "..\fondos\Bomberos\fondo01-1280x960.jpg"; DestDir: "{app}\fondos\Bomberos";
#endif

Source: "..\ALCemi\formularios\recibo.form"; DestDir: "{app}\bin\formularios";
Source: "..\ALCemi\formularios\reciboAfiliados.form"; DestDir: "{app}\bin\formularios";
Source: "..\ALCemi\formularios\reciboArea.form"; DestDir: "{app}\bin\formularios";




;archivos de sistemas de codificacion
Source: "..\ALCEMI\Codificaciones\Bomberos\Sistema Bomberos Mendiolaza.xml"; DestDir: "{app}\bin\Codificaciones\Bomberos";
Source: "..\ALCEMI\Codificaciones\Emergencias\Sistema Actual.xml"; DestDir: "{app}\bin\Codificaciones\Emergencias";
Source: "..\ALCEMI\Codificaciones\Emergencias\Sistema Actual (Con codigo Celeste).xml"; DestDir: "{app}\bin\Codificaciones\Emergencias";

; restartreplace quiere decir que si el archivo existe se copiara una copia en temporales para ser reemplazado la proxima que inicie windows. Solo para archivos de sistema
; uninsneveruninstall quiere decir que no es una librer�a m�a por lo que no se desinstalar� cuando se desinstale el sistema
; sharedfile quiere decir que es una librer�a que puede usar mas de un programa (para que no la desinstale)
; regserver es para que ejecute el comando regsvr32 xxx.dll y registre la dll

Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\scrrnes.dll";  DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\scrrun.dll";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
; estos dos de recien son para usar p�r ejemplo el FileSystemObject. Es muy comun usarlos
; ejemplo de tlb, es igual a dll
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\wbemdisp.tlb"; DestDir: "{sys}\Wbem"; Flags: restartreplace uninsneveruninstall sharedfile
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\tabctl32.ocx"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile
; ejemplo de dll mia
; ojo que si la dll instalada es versi�n >= que la actual no la instalar�!!!!!
;Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrerr.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
; ejemplo de ocx. Mismo cuidado con la versi�n
;Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrFrame.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver

#if fullInst
  ;Base de datos
  ; onlyifdoesntexist instala la base SOLO si no existe. Sirve perfecto para actualizaciones
  Source: "..\BD\database.inst.mdb"; DestDir: {app}\BD; Destname:database.mdb; Flags: onlyifdoesntexist
  ; esta vacia se pone en oct 2010 para que lo yo toco no salga al publico
  ; ver que tenga datos basicos para empezar
  ;Source: "..\BD\databaseDemo.mdb"; DestDir: {app}\BD; Flags: onlyifdoesntexist
#endif

;**************************************************************************************
;access segun artime .... (increible pero sin access funciona !!!!)


#if fullInst
  ; (Unsafe file) Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\comdlg32.dll"; DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile regserver
  Source: C:\Archivos de programa\Inno Setup 5\vbFiles\mscomctl.ocx; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
  Source: C:\Archivos de programa\Inno Setup 5\vbFiles\mscomct2.ocx; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
  ;Source: C:\Archivos de programa\Inno Setup 5\vbFiles\msado21.tlb; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regtypelib

  ;me da error en vista!!!!!!!!!!
  Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\msado27.tlb"; DestDir: "{cf}\system\ado"; Flags: restartreplace uninsneveruninstall regtypelib; onlybelowversion: 5.2,5.2
  Source: C:\Archivos de programa\Inno Setup 5\vbFiles\MSDATGRD.OCX; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver


  ;me da error en vista y me parece q no lo uso VER!!!!!!!!!!!!
  Source: C:\Archivos de programa\Inno Setup 5\vbFiles\msxml3.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver; onlybelowversion: 5.2,5.2
  ;no estoy seguro del onlybelowversion, preguntar...
  Source: C:\Archivos de programa\Inno Setup 5\vbFiles\msxml.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver; onlybelowversion: 5.2,5.2

  ;Access
  ;pruebo sin esto Source: C:\Archivos de programa\Inno Setup 5\vbFiles\Other\MSACC9.OLB; DestDir: {sys}; Flags: restartreplace sharedfile
  Source: C:\Archivos de programa\Archivos comunes\Microsoft Shared\DAO\dao360.dll; DestDir: {dao}; Flags: restartreplace uninsneveruninstall regserver; MinVersion: 4.0,0
  Source: C:\Archivos de programa\Archivos comunes\System\ADO\msjro.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall regserver; MinVersion: 4.0,0


  ;MDAC
  Source: "MDAC - JET\mdac28.exe"; DestDir: {tmp}; OnlyBelowVersion: 4.2,5.0; Flags: Deleteafterinstall; Tasks: MDAC
  Source: "MDAC - JET\mdac28 (ME).exe"; DestDir: {tmp}; MinVersion: 4.2,0; Flags: Deleteafterinstall; Tasks: MDAC
  Source: "MDAC - JET\Jet40SP8_9xNT.exe"; DestDir: {tmp}; OnlyBelowVersion: 4.2,5.0; Flags: Deleteafterinstall; Tasks: JET
  Source: "MDAC - JET\Jet40SP7_WMe.exe"; DestDir: {tmp}; MinVersion: 4.2,0; Flags: Deleteafterinstall; Tasks: JET

;**************************************************************************************
#endif

;--------------------------------------------------------------------------------------
;**************************************************************************************
;*******************aqui pongo todas mis dependencias personales***********************
;**************************************************************************************
; libreria del sistema de licencias
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrcaescrypto.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrerr.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrnfo.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrjuse.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrjuse2.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "C:\Archivos de programa\Inno Setup 5\vbFiles\other\tbrtimer.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
; esta mas arriba Source: "..\dlls\wbemdisp.tlb"; DestDir: "{sys}\Wbem"; Flags: restartreplace uninsneveruninstall sharedfile

Source: "..\BLCemi\BLCemi.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "..\DataBaseLayer\DataBaseLayer.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "..\tbrCamposImpresion\tbrCamposImpresion.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
;/////////////////////////////////////////
; se instala aqui por el path de la base de datos
; pateeeeeeeeeeeeeeeeetico
Source: "..\Config\tbrConfig.dll"; DestDir: "{app}/bin"; Flags: restartreplace sharedfile regserver
;/////////////////////////////////////////
Source: "..\ControlesUsuario\ListViewConsultaCtl2.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "..\tbrNet\tbrNet.dll"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver



;**************************************************************************************
;**************************************************************************************

; main ejecutable
Source: "..\AlCemi\tbrEmergencyGroup.EXE"; DestDir: "{app}\bin"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}\tbrEmergencyGroup"; Filename: "{app}\bin\{#MyAppExeName}"
Name: "{group}\{#MyAppName}\Manual"; Filename: "{app}\Manual\Manual de usuario.doc"
;Name: "{group}\{#MyAppName}\Reparar 3PM"; Filename: "{app}\repair.exe"
; lo que sogue es uin acceso directo a desinstalar
Name: "{group}\{#MyAppName}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
; acceso directo en el escritotio
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\bin\{#MyAppExeName}"; Tasks: desktopicon
; acceso en inicio r�pido
Name: "{userappdata}\Microsoft\Internet Explorer\Quick Launch\{#MyAppName}"; Filename: "{app}\bin\{#MyAppExeName}"; Tasks: quicklaunchicon

;le pongo por defecto la base de datos de pruebas, despues la pueden cambiar por la vacia
[Registry]
;Root: HKCU; SubKey: Software\VB and VBA Program Settings\TbrEmergencyGroup\DBLayer; ValueType: string; ValueName: PathDB; ValueData:{app}\bd\databasedemo.mdb; Flags: uninsdeletekey;
[Run]
#if fullInst
  ; opciones para ejecutar al terminar la inatalaci�n
  Filename: {tmp}\mdac28.exe; Parameters: "/Q /C:""setup /QNT"""; WorkingDir: {tmp}; Flags: skipifdoesntexist; Tasks: MDAC; OnlyBelowVersion: 4.2,5.0
  Filename: {tmp}\mdac28 (ME).exe; Parameters: "/Q /C:""setup /QNT"""; WorkingDir: {tmp}; Flags: skipifdoesntexist; Tasks: MDAC; MinVersion: 4.2,0
  Filename: {tmp}\Jet40SP8_9xNT.exe; Parameters: "/Q /C:""setup /QNT"""; WorkingDir: {tmp}; Flags: skipifdoesntexist; Tasks: JET; OnlyBelowVersion: 4.2,5.0
  Filename: {tmp}\Jet40SP7_WMe.exe; Parameters: "/Q /C:""setup /QNT"""; WorkingDir: {tmp}; Flags: skipifdoesntexist; Tasks: JET; MinVersion: 4.2,0
#endif

Filename: "{app}\bin\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#MyAppName}}"; Flags: nowait postinstall skipifsilent