;NSIS script for StealthBot v2.6 Revision 3 created by Atomic GUI for NSIS

!include MUI2.nsh

!define MUI_HEADERIMAGE 
!define MUI_HEADERIMAGE_BITMAP "Assets\Banner2.bmp"
!define MUI_INSTFILESPAGE_COLORS "0099CC 000000" ;Two colors

; Title of this installation
Name "StealthBot v2.7"

; Do a CRC check when initializing setup
CRCCheck On

XPStyle On

BrandingText "StealthBot v2.7"

; Output filename
Outfile "Build\StealthBotSetup.exe"

; pages
!insertmacro MUI_PAGE_LICENSE "PF\eula.txt"
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

; The default installation folder
InstallDir "$PROGRAMFILES\StealthBot"

; Uninstall info registry location
InstallDirRegKey HKLM "SOFTWARE\StealthBot" "Install_Dir"

; Folder selection prompt
DirText "Please select an installation folder."

; Section Default
; Section Default
Section "PF" PF
     SetCompress Auto
     SetDateSave On
     
     SetOutPath "$INSTDIR"
     File "PF\eula.txt"
     File "PF\Commands.xsd"
     File "PF\BNCSutil.dll"
     File "PF\libeay32.dll"
     File "PF\zlib1.dll"
     File "PF\Warden.dll"
     File "PF\Launcher.exe"
     File "PF\StealthBot v2.7.exe"
     SetOutPath "$INSTDIR\Default"
     File "PF\Default\Commands.xml"
     File "PF\Default\CheckRevision.ini"
     File "PF\Default\Warden.ini"
     SetOutPath "$INSTDIR\Default\scripts"
     File "PF\Default\scripts\demo.txt"
     File "PF\Default\scripts\PluginSystem.txt"
     SetOutPath "$INSTDIR\Default\scripts\demo"
     File "PF\Default\scripts\demo\frm.txt"
     File "PF\Default\scripts\demo\sck.txt"
     SetOutPath "$INSTDIR\Default\scripts\lib\StealthBot"
     File "PF\Default\scripts\lib\StealthBot\frmDataGrid.vbs"
     File "PF\Default\scripts\lib\StealthBot\XMLTextWriter.vbs"
SectionEnd

Section "Components" Components
     SetOutPath "$SYSDIR"
     SetOverwrite off
     SetCompress Auto
     SetDateSave On
     
     ;Do not overwrite these files if they exist or are in use
     ; (Dependencies)
     File "Components\COMDLG32.OCX"
     RegDLL COMDLG32.OCX
     File "Components\MSCOMCTL.OCX"
     RegDLL MSCOMCTL.OCX
     File "Components\MSINET.OCX"
     RegDLL MSINET.OCX
     File "Components\MSSCRIPT.OCX"
     RegDLL MSSCRIPT.OCX
     File "Components\MSWINSCK.OCX"
     RegDLL MSWINSCK.OCX
     File "Components\RICHTX32.OCX"
     RegDLL RICHTX32.OCX
     File "Components\SSubTmr6.dll"
     RegDLL SSubTmr6.dll
     File "Components\TABCTL32.OCX"
     RegDLL TABCTL32.OCX
     File "Components\TLBINF32.DLL"
     RegDLL TLBINF32.DLL
     File "Components\vbalTreeView6.ocx"
     RegDLL vbalTreeView6.ocx
SectionEnd

Section -post
     SetOutPath "$INSTDIR"
     CreateDirectory "$DESKTOP"
     CreateShortCut "$DESKTOP\StealthBot.lnk" "$INSTDIR\Launcher.exe" "" "$INSTDIR\Launcher.exe" 0
SectionEnd

; This emptily named section will always run
Section ""
     WriteRegStr HKLM "SOFTWARE\StealthBot" "Install_Dir" "$INSTDIR"
     WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\StealthBot" "DisplayName" "StealthBot v2.7 (remove only)"
     WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\StealthBot" "UninstallString" '"$INSTDIR\uninst.exe"'

     SetOutPath $INSTDIR
     WriteUninstaller "uninst.exe"
SectionEnd

; Uninstall section here...
UninstallText "This will uninstall StealthBot v2.7. Press NEXT to continue."
Section "Uninstall"
     DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\StealthBot"
     
     ;SetOutPath "$INSTDIR\Default\scripts\lib\StealthBot"
     Delete "$INSTDIR\Default\scripts\lib\StealthBot\frmDataGrid.vbs"
     Delete "$INSTDIR\Default\scripts\lib\StealthBot\XMLTextWriter.vbs"
     RmDir "$INSTDIR\Default\scripts\lib\StealthBot"
     RmDir "$INSTDIR\Default\scripts\lib"
     
     ;SetOutPath "$INSTDIR\Default\scripts\demo"
     Delete "$INSTDIR\Default\scripts\demo\frm.txt"
     Delete "$INSTDIR\Default\scripts\demo\sck.txt"
     RmDir "$INSTDIR\Default\scripts\demo"
     
     ;SetOutPath "$INSTDIR\Default\scripts"
     Delete "$INSTDIR\Default\scripts\demo.txt"
     Delete "$INSTDIR\Default\scripts\PluginSystem.txt"
     RmDir "$INSTDIR\Default\scripts"
     
     ;SetOutPath "$INSTDIR\Default"
     Delete "$INSTDIR\Default\Commands.xml"
     Delete "$INSTDIR\Default\CheckRevision.ini"
     Delete "$INSTDIR\Default\Warden.ini"
     RmDir "$INSTDIR\Default"
     
     ;SetOutPath "$INSTDIR"
     Delete "$INSTDIR\eula.txt"
     Delete "$INSTDIR\Commands.xsd"
     Delete "$INSTDIR\BNCSutil.dll"
     Delete "$INSTDIR\libeay32.dll"
     Delete "$INSTDIR\zlib1.dll"
     Delete "$INSTDIR\Warden.dll"
     Delete "$INSTDIR\Launcher.exe"
     Delete "$INSTDIR\StealthBot v2.7.exe"
     Delete "$INSTDIR\uninst.exe"
     RmDir "$INSTDIR"
SectionEnd


