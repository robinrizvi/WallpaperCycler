;NSIS Modern User Interface
;Start Menu Folder Selection Example Script
;Written by Joost Verburg

;--------------------------------
;Include Modern UI

  !include "mui2.nsh"
  !include "StrFunc.nsh"
  !include "Library.nsh"
  !include "VB6RunTime.nsh"
;--------------------------------
;General

  ;Name and file
Name "Wallpaper Cycler"
OutFile "WCsetup.exe"
RequestExecutionLevel user
SetCompress auto
SetCompressor /SOLID /FINAL Lzma
WindowIcon on
AutoCloseWindow false
SetDatablockOptimize on
CRCCheck on
InstProgressFlags smooth
SetDateSave on
AllowSkipFiles on
SetOverwrite IfNewer
  
  ;Default installation folder
  InstallDir "$PROGRAMFILES\SoftLogic"

  ;Get installation folder from registry if available
  ;InstallDirRegKey HKCU "Software\Modern UI Test" ""

  ;Request application privileges for Windows Vista
  RequestExecutionLevel user

;--------------------------------
;Variables

  Var StartMenuFolder
  Var ALREADY_INSTALLED
  var Version
;--------------------------------
;Interface Settings

  !define MUI_ABORTWARNING

;--------------------------------
;Pages

  !insertmacro MUI_PAGE_LICENSE "${NSISDIR}\Docs\Modern UI\License.txt"
  !insertmacro MUI_PAGE_COMPONENTS
  !insertmacro MUI_PAGE_DIRECTORY

  ;Start Menu Folder Page Configuration
  !define MUI_STARTMENUPAGE_REGISTRY_ROOT "HKCU"
  !define MUI_STARTMENUPAGE_REGISTRY_KEY "Software\Wallpaper Cycler"
  !define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "Start Menu Folder"

  !insertmacro MUI_PAGE_STARTMENU Application $StartMenuFolder

  !insertmacro MUI_PAGE_INSTFILES

  !insertmacro MUI_UNPAGE_CONFIRM
  !insertmacro MUI_UNPAGE_INSTFILES

;--------------------------------
;Languages

  !insertmacro MUI_LANGUAGE "English"

;--------------------------------
;Installer Sections

Section "Wallpaper Cycler" SecDummy
  ReadRegDWord $Version HKLM "Software\Wallpaper Cycler" Version
  IfErrors new_installation
	StrCpy $ALREADY_INSTALLED 1
	new_installation:

  SetOutPath "$INSTDIR"

  ;ADD YOUR OWN FILES HERE...
  File /a "Support\Wallpaper Cycler.exe"
  
# Put shortcuts in the start menu
	CreateDirectory "$SMPrograms\Wallpaper Cycler"        
	CreateShortCut "$SMPrograms\Wallpaper Cycler\Wallpaper Cycler.lnk" "$InstDir\Wallpaper Cycler.EXE"
	CreateShortCut "$DESKTOP\Wallpaper Cycler.lnk" "$InstDir\Wallpaper Cycler.EXE"
	CreateShortCut "$SMPrograms\Wallpaper Cycler\uninstaller.lnk" "$InstDir\uninstaller.EXE"
	WriteRegDWord HKLM "Software\Wallpaper Cycler" Version 1
	WriteUninstaller "$InstDir\uninstaller.exe"
	# Add uninstaller information to registry.
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Wallpaper Cycler" "DisplayName" "Wallpaper Cycler"
	WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Wallpaper Cycler" "UninstallString" '"$InstDir\uninstaller.exe"'
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Wallpaper Cycler" "NoModify" 1
	WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Wallpaper Cycler" "NoRepair" 1
	
  ;Store installation folder


  ;Create uninstaller


  !insertmacro MUI_STARTMENU_WRITE_BEGIN Application

    ;Create shortcuts
    CreateDirectory "$SMPROGRAMS\$StartMenuFolder"
    ;CreateShortCut "$SMPROGRAMS\$StartMenuFolder\Uninstall.lnk" "$INSTDIR\Uninstaller.exe"

  !insertmacro MUI_STARTMENU_WRITE_END

SectionEnd

;--------------------------------
;Descriptions

  ;Language strings
  LangString DESC_SecDummy ${LANG_ENGLISH} "Install Wallpaper Cycler"

  ;Assign language strings to sections
  !insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
    !insertmacro MUI_DESCRIPTION_TEXT ${SecDummy} $(DESC_SecDummy)
  !insertmacro MUI_FUNCTION_DESCRIPTION_END

;--------------------------------
;Uninstaller Section

Section "Uninstall"

  ;ADD YOUR OWN FILES HERE...
        MessageBox MB_OK "Please close any of the instances of Wallpaper Cycler before uninstall else the application will be rendered unfunctional."
again:
	DeleteRegKey HKCU "Software\WallpaperCycler"
	DeleteRegKey HKCU "Software\Wallpaper Cycler"
	DeleteRegValue HKCU "Software\Microsoft\Windows\CurrentVersion\Run" "WallpaperCycler"
	Delete /RebootOK "$instDir\count.inf"
	Delete /RebootOK "$instDir\main.bmp"
  	Delete /RebootOk "$SMPrograms\Wallpaper Cycler\Wallpaper Cycler.lnk"
  	Delete /RebootOk "$DESKTOP\Wallpaper Cycler.lnk"
	UnRegDLL "$InstDir\Wallpaper Cycler.exe"
	Delete /RebootOk "$InstDir\Wallpaper Cycler.exe"

	Delete /RebootOk "$SMPrograms\Wallpaper Cycler\uninstaller.lnk"
	Delete /RebootOk "$InstDir\uninstaller.exe"

	RMDir "$SMPrograms\"
	;DeleteRegValue HKCU "Software\WallpaperCycler"
	
	SetOutPath "\"	;Change to root directory to allow the next step to succeed.
	RmDir /RebootOk "$INSTDIR"
  Delete "$INSTDIR\Uninstall.exe"

  RMDir "$INSTDIR"

  !insertmacro MUI_STARTMENU_GETFOLDER Application $StartMenuFolder

  Delete "$SMPROGRAMS\$StartMenuFolder\Uninstall.lnk"
  RMDir "$SMPROGRAMS\$StartMenuFolder"
IfFileExists "$INSTDIR\Wallpaper Cycler.exe" 0 new_ins
   MessageBOx MB_RetryCancel "An instance Wallpaper Cycler is open. Please close it and press Retry. Pressing Cancel will restart the computer to complete the process." IDRETRY again IDCANCEL reb
reb:
Reboot 
 new_ins:
SectionEnd

Section "-Install VB6 runtimes"
	;Add code here that sets $ALREADY_INSTALLED to a non-zero value if"
	;the application is already installed. For example:
         IfFileExists "$INSTDIR\Wallpaper Cycler.exe" 0 new_installation2 ;Replace MyApp.exe with your application filename
   StrCpy $ALREADY_INSTALLED 1
 new_installation2:
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_NOTPROTECTED \
		"Support\msvbvm60.dll" "$SYSDIR\msvbvm60.dll" "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
		"Support\oleaut32.dll" "$SYSDIR\oleaut32.dll" "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
		"Support\olepro32.dll" "$SYSDIR\olepro32.dll" "$SYSDIR"
	!insertmacro InstallLib REGDLL $ALREADY_INSTALLED REBOOT_PROTECTED \
		"Support\comcat.dll"   "$SYSDIR\comcat.dll"   "$SYSDIR"
	!insertmacro InstallLib DLL    $ALREADY_INSTALLED REBOOT_PROTECTED \
		"Support\asycfilt.dll" "$SYSDIR\asycfilt.dll" "$SYSDIR"
	!insertmacro InstallLib TLB    $ALREADY_INSTALLED REBOOT_PROTECTED \
		"Support\stdole2.tlb"  "$SYSDIR\stdole2.tlb"  "$SYSDIR"
SectionEnd

Section "-un.Uninstall VB6 runtimes"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\msvbvm60.dll"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\oleaut32.dll"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\olepro32.dll"
	!insertmacro UnInstallLib REGDLL SHARED NOREMOVE "$SYSDIR\comcat.dll"
	!insertmacro UnInstallLib DLL    SHARED NOREMOVE "$SYSDIR\asycfilt.dll"
	!insertmacro UnInstallLib TLB    SHARED NOREMOVE "$SYSDIR\stdole2.tlb"
SectionEnd
