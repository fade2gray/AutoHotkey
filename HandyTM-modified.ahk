/*
    The Handy TNA Menu for AHK - by SKAN // Ver:Pre 1.0  CD: 19-Nov-2008 / LM: 29-Nov-2008
	[Modified by fade2gray to include use with 64Bit OS 23/07/2018 - AutoHotkey v1.1.29.01]
	[Requires to be run/compiled as ANSI 32 (A32)]

        ooooo ooooo                              oooo            ooooooooooo oooo     oooo
         888   888   ooooooo   oo oooooo    ooooo888 oooo   oooo 88  888  88  8888o   888 
         888ooo888   ooooo888   888   888 888    888  888   888      888      88 888o8 88 
         888   888 888    888   888   888 888    888   888 888       888      88  888  88 
        o888o o888o 88ooo88 8o o888o o888o  88ooo888o    8888       o888o    o88o  8  o88o
                                                      o8o888
					____________________________________________________________
						C·o·n·t·r·o·l  a·l·l  r·u·n·n·i·n·g  s·c·r·i·p·t·s
							w·i·t·h  o·n·e  m·a·s·t·e·r  m·e·n·u
*/

#Persistent
#NoTrayIcon
#NoEnv

SetBatchLines -1
#SingleInstance, Force
SetWorkingDir, %A_ScriptDir%
DetectHiddenWindows, On
OnMessage( 0x404, "AHK_NotifyIcon" )
InitGlobalVars()

Menu, Tray, Add
Menu, Tray, Add, Delete %DAT% , DeleteSettings
Menu, Tray, Default
Menu, Tray, Tip, [ HandyTM - Pre 1.0 ]

Gui0 := WinExist( A_ScriptFullPath " ahk_class AutoHotkey" ), SetTrayIcon( Gui0 )
VarSetCapacity( nbs, 40, 32 ) ; this affects the Menu width
OnExit, QuitScript
OnMessage( DllCall( "RegisterWindowMessage", Str,"AHKHOOK" ), "AHK_Hook" )

Return ;                                                   // end of auto-execute section //

AHK_Hook( PID, State ) {
 SetTimer, CreateLeftClickMenu, -1
Return 0
}

AHK_NotifyIcon( wParam, lParam ) {
          If ( lParam = 0x202 )  {               ; WM_LBUTTONUP
            SetTimer, ShowLeftClickMenu, -1
            Return 0
   } Else If ( lParam = 0x208 )  {               ; WM_MBUTTONUP
            SetTimer, RunFromClipboard, -1
            Return 0
   } Else If ( lParam = 0x205 )  {               ; WM_RBUTTONUP
            SetTimer, ShowRightClickMenu, -1
            Return 0
   }
}

CreateLeftClickMenu:
 CreateLeftClickMenu()
Return

ShowRightClickMenu:
 CreateRightClickMenu()
Return

CreateRightClickMenu() {
 Global Gui0, DAT, InstallExe, nbs
 MenuList := "", VarSetCapacity( nbs,50,32)

 Menu, RightClickMenu,      UseErrorLevel
 Menu, RightClickMenu,      DeleteAll

 LaunchList := IniReadSection( DAT, "LaunchList" )
 Loop, Parse, LaunchList, `n
 {
   StringSplit,P, A_LoopField, =
   IfNotExist, %P2%, Continue
   LastRun    := IniGet( DAT, P1, "LastRun" )
   IfEqual,LastRun,0000000000000000, Continue
   MenuName   := IniGet( DAT, P1, "MenuName" )
   RunParam   := IniGet( DAT, P1, "RunParam" )
   MenuList   .=  ( (MenuList<>"") ? "`n" : "" )  LastRun "|" SubStr(MenuName nbs,1,100)
              . "|" RunParam
 }

 Sort, MenuList, R D`n
 Loop, Parse, MenuList, `n
 {
   Loop, Parse, A_LoopField, |
      F%A_Index% := RegExReplace( A_LoopField, "(^\s*|\s*$)" )
   MenuName := F2,  RunParam := F3
   Menu, %MenuName%, UseErrorLevel
   Menu, %MenuName%, DeleteAll
   Menu, %MenuName%, Add, %RunParam%, LaunchScript
   If MenuName not in %MenuArray%
      MenuArray .= (( MenuArray<>"" ) ? "," : "" ) MenuName
 }

 Loop, Parse, MenuArray, `,
 {
   Menu, RightClickMenu, Add, %A_LoopField%, :%A_LoopField%
   IfGreaterOrEqual, A_Index, 25, Break
 }

 IfNotEqual,MenuList,,Menu, RightClickMenu, Add
 Menu, RightClickMenu, Add, Run Clipboard Code, RunFromClipBoard
 Menu, RightClickMenu, Add, Exit all Temp Scripts, ExitAllTemp
 Menu, RightClickMenu, Show
Return
}

LaunchScript:
 Hash := HashStr( A_ThisMenuItem ), WorkingDir := IniGet( DAT, Hash, "WorkingDir" )
 Run, %A_ThisMenuItem%, %WorkingDir%, UseErrorLevel
Return

RunFromClipboard:
 Clipboard := StrReplace(Clipboard, "`n", "`n", ErrorLevel)
 IfLess, ErrorLevel, 1
 {
	MsgBox,0,HandyTM,Run Clipboard Code`n`nAt least 1 line of code with a`nterminating line feed required.
	Return
 }
 Code := "#ErrorStdOut `; Temporary ahk script created by HandyTM`n" Clipboard
 Code := StrReplace(Code, Chr(160), Chr(32), ErrorLevel)
 FileAppend, %Code%, % ( TmpAhkFile := A_Temp "\~AHK" HashStr( A_Now,2 ) ".ahk" )
 Loop %TmpAhkFile%
    TmpAhkFile := A_LoopFileLongPath
 Run, %Installexe% "%TmpAhkFile%", %A_Temp%, UseErrorLevel
 WinWait, %TmpAhkFile% ahk_class AutoHotkey
 SendMessage, 0x111, 65401, 0,, %TmpAhkFile%
Return

ExitAllTemp:
 Loop %A_Temp%,1
    TmpTitle := A_LoopFileLongPath "\~AHK"
 WinGet, TAHK, List, %TmpTitle% ahk_class AutoHotkey
 Loop %TAHK%
   PostMessage, 0x111, 65405, 0,, % "ahk_id " TAHK%A_Index%
Return

CreateLeftClickMenu() {
 Static Q := """" , C := "," , LF := "`n"
 Static ProcessPriorities := "Low|BelowNormal|Normal|AboveNormal|High|Realtime"
 Static LinkFolders := "Startup|Desktop|QuickLaunch"
 Global Gui0, DAT, InstallExe, A_QuickLaunch, nbs

 WinGet, AhkW, List, ahk_class AutoHotkey
 Loop %AhkW% {

   hWnd := uHex( AhkW%A_Index% )
   AhkF := ScriptFullPath( hWnd )
   SplitPath, AhkF, ScriptName, OD, OE, NNE
   WinGet, PID, PID, ahk_id %hWnd%

   ProcessClose( nProc )
   nProc := ProcessOpen( PID )
   If ( ProcessGetOwner( nProc,0 )<> A_UserName )
        Continue
   peFile := ProcessGetModuleFileName( nProc,A_Is64BitOS ? PID : 0 )
   pTime := ( hWnd = Gui0 ) ? "0000000000000000" : ProcessGetCreationTime( nProc,0 )
   CmdL  := ProcessGetCommandLine( nProc,A_Is64BitOS ? PID : 0 )
   Hash := HashStr(CmdL)
   CmdL := RegExReplace( CmdL, "^"".*""$", """$0""" )

   If ( SubStr( NNE,1,4) <> "~AHK" ) {
   IniPut( AhkF  , DAT, "LaunchList", Hash         )
   IniPut( NNE   , DAT, Hash,         "MenuName"   )
   IniPut( CmdL  , DAT, Hash,         "RunParam"   )
   IniPut( OD    , DAT, Hash,         "WorkingDir" )
   IniPut( pTime , DAT, Hash,         "LastRun"    )
   }
   If IniGet( DAT, "MenuFlag", AhkF, 0 )
      Continue

   MenuList .= pTime C hWnd C  PID C Q peFile Q C NNE C OE C Q ScriptName Q C Q AhkF Q LF
 } StringTrimRight, MenuList, MenuList, 1
 Sort MenuList, R D`n

 MenuFlag := IniReadSection( DAT, "MenuFlag" ), RemE := 0, MI := 1
 Loop, Parse, MenuFlag, `n
 {  P1 := P2 := 0
    StringSplit, P, A_LoopField, =
    If ! FileExist( P1 ) {
       IniDelete, %DAT%, MenuFlag, %P1%
       Continue
 }  If ( P2=1 ) {
       Menu, Restore, Add, %P1%, MenuRestore
       RemE := RemE+1
 }}
    If ( RemE ) {
    RestoreMenu := "Restore Menu" nbs "`t" chr(0x00AB) " " RemE " script" (RemE>1 ? "s" : "" ) " " chr(0x00BB)
    Menu, LeftClickMenu, Add,    %RestoreMenu%, :Restore
    Menu, LeftClickMenu, Add
 }

 Loop, Parse, MenuList, `n
 {
   Loop, Parse, A_LoopField, CSV
         F%A_Index% := A_LoopField
   pTime:=F1, hWnd:=F2, PID:=F3, peFile:=F4, NNE:=F5, OE:=F6, ScriptName:=F7,  AhkF:=F8

   Vers := FileGetVersionInfo( peFile, "ProductName" ) " v"
        .  RegExReplace( FileGetVersionInfo( peFile, "ProductVersion" ), ", ", "." )

   If ( SubStr( NNE,1,4) <> "~AHK" )
   IniPut( False, DAT, "MenuFlag", AhkF )

   MainMenu       := ScriptName "`t{" hWnd  "}"
   ProcessMenu    := AhkF "|" PID
   PriorityMenu   := PID
   SwapExeMenu    := PID "|" hWnd "|" peFile "|" Vers "|"
   ShortcutMenu   := AhkF "|" hWnd "|" PID
   DriveMenu      := AhkF

   Menu, %MainMenu%,      UseErrorLevel
   Menu, %MainMenu%,      DeleteAll
   Menu, %PriorityMenu%,  UseErrorLevel
   Menu, %PriorityMenu%,  DeleteAll
   Menu, %ShortcutMenu%,  UseErrorLevel
   Menu, %ShortcutMenu%,  DeleteAll
   Menu, %SwapExeMenu%,   UseErrorLevel
   Menu, %SwapExeMenu%,   DeleteAll
   Menu, %DriveMenu%,     UseErrorLevel
   Menu, %DriveMenu%,     DeleteAll


   Loop, Parse, ProcessPriorities, |
   Menu, %PriorityMenu%,   Add,          %A_LoopField%,               SetPriority
   Menu, %PriorityMenu%,   Disable,      Realtime

   Menu, %PriorityMenu%,   Check,        % ProcessGetPriority( nProc,0 )
   Menu, %ProcessMenu%,    UseErrorLevel
   Menu, %ProcessMenu%,    DeleteAll
   Menu, %ProcessMenu%,    Add,          Priority,                    :%PriorityMenu%

   IfNotEqual,hWnd,%Gui0%
   Menu, %ProcessMenu%,    Add,          Close/Kill,                  ProcessClose
   Menu, %ProcessMenu%,    Add,
   Menu, %ProcessMenu%,    Add,          File Properties Dialog (EXE), ShellOptions
   Menu, %MainMenu%,       Add,          Process %nbs%`t{ PID:%PID% },:%ProcessMenu%

   IfNotEqual,hWnd,%Gui0%, IfEqual,OE,AHK
{
   Menu, %SwapExeMenu%,    Add,      Select AHK Executable`t[%PeFile%],SelectAhkExecutable
   AhkVersions := IniReadSection( DAT, "AhkVersions" ), Chk := 0
   Loop, Parse, AhkVersions, `n
   {
   StringSplit, K, A_LoopField, =
   IfEqual,K0,0,Continue
   If ! FileExist( K2 ) {
      IniDelete, %DAT%, AhkVersions, %K1%
      Continue
   }
   Chk +=1
   IfEqual, Chk,1, Menu, %SwapExeMenu%,    Add
   Menu, %SwapExeMenu%,    Add,          %K1%,                        SwapAhkExecutable
 }
   Menu, %MainMenu%,       Add,          %Vers%`tSwap EXE w/,         :%SwapExeMenu%
}
   Menu, %MainMenu%, Add

   Menu, %MainMenu%,       Add,          Open,                        DefaultTrayOptions
   Menu, %MainMenu%,       Add,          Reload,                      DefaultTrayOptions

   IfNotEqual,OE,EXE
   Menu, %MainMenu%,       Add,          Edit Script,                 DefaultTrayOptions

   Menu, %MainMenu%,       Add,          Suspend,                     DefaultTrayOptions
   Menu, %MainMenu%,       Add,          Pause,                       DefaultTrayOptions
   Menu, %MainMenu%,       Add,          Exit,                        DefaultTrayOptions
   Menu, %MainMenu%,       Add

   IfNotEqual,hWnd,%Gui0%
{
   Menu, %MainMenu%,       Add,          Reload`t[ w/ Parameters ],   ReloadWithParameters
   Menu, %MainMenu%,       Add,          Exit`t[ all Instances ],  ExitAllSimilarInstances
   Menu, %MainMenu%,       Add
}

   If ( SubStr( NNE,1,4) <> "~AHK" )
{
   Loop, Parse, LinkFolders, |
   Menu, %ShortcutMenu%,   Add,          %A_LoopField%,               CreateShortcut
   Menu, %MainMenu%,       Add,          Create Shortcut,             :%ShortcutMenu%
   Menu, %MainMenu%,       Add
}

   IfNotEqual,OE,Exe
   Menu, %MainMenu%,       Add,          File Properties Dialog,      ShellOptions

   Menu, %MainMenu%,       Add,          Explore Script Folder,       ShellOptions
   Menu, %MainMenu%,       Add,          Select Script Folder,        ShellOptions

   Menu, %MainMenu%,       Add
   Menu, %MainMenu%,       Add,          Remove Menu,                 MenuRemove
   Menu, LeftClickMenu,    Add,          %MainMenu%,                  :%MainMenu%

   MI += 1
  }
Return MI
}

DefaultTrayOptions:
 RegExMatch(A_ThisMenu,"\{(.*)\}",ID),hWnd:=ID1, AhkF:=ScriptFullPath(hWnd), wParam:=0
 IfEqual, A_ThisMenuItem, Reload     , SetEnv, wParam, 65400
 IfEqual, A_ThisMenuItem, Edit Script, SetEnv, wParam, 65401
 IfEqual, A_ThisMenuItem, Pause      , SetEnv, wParam, 65403
 IfEqual, A_ThisMenuItem, Suspend    , SetEnv, wParam, 65404
 IfEqual, A_ThisMenuItem, Exit       , SetEnv, wParam, 65405
 IfEqual, A_ThisMenuItem, Open       , SetEnv, wParam, 65407
 IfNotEqual,wParam,0, PostMessage, 0x111, wParam, 0,, ahk_id %hWnd%
Return

ReloadWithParameters:
 RegExMatch(A_ThisMenu,"\{(.*)\}",ID),hWnd:=ID1, AhkF:=ScriptFullPath(hWnd)
 WinGet, PID, PID, ahk_id %hWnd%
 CmdL := ProcessGetCommandLine( PID ), AhkF := ScriptFullPath( hWnd )
 SplitPath, AhkF,,WorkingDir
 PostMessage, 0x111, 65405,0,,ahk_id %hWnd%
 Loop 30 {
     Sleep 100
     Process, Exist, %PID%
     IfEqual, ErrorLevel,0, Break
 }   If ( ErrorLevel ) {
     TrayTip, HandyTM, Unable to Exit:`n%AhkF%`n`nTry killing the Process: %PID%,,2
     Return
 }   Run, %CmdL%, %WorkingDir% , UseErrorLevel
 IfEqual,ErrorLevel,ERROR, TrayTip, HandyTM, Unable to Launch:`n%CmdL%,,3
Return

ExitAllSimilarInstances:
 RegExMatch(A_ThisMenu,"\{(.*)\}",ID),hWnd:=ID1, AhkF:=ScriptFullPath(hWnd)
 WinGet, Sim, List, %AhkF% ahk_class AutoHotkey
 Loop %Sim%
   PostMessage, 0x111, 65405, 0,, % "ahk_id " Sim%A_Index%
Return

CreateShortcut:
 cPath := "A_" A_ThisMenuItem
 cPath := %cPath%
 StringSplit, P, A_ThisMenu, |
 AhkF := P1, hWnd := P2, PID := P3,
 SplitPath, AhkF,,Work,,Script
 CmdL := ProcessGetCommandLine( PID )
 Targ := PathRemoveArgs( CmdL )
 Args := StrReplace(CmdL, Targ)
 Args := PathRemoveBlanks( Args )
 Link  := cPath "\" Script ".lnk"
 FileCreateShortcut, %Targ%, %Link%, %Work%, %Args%, Link auto-created with HandyTM
 IfNotEqual, ErrorLevel, 1
{
 TrayTip, HandyTM, Shortcut Created in %A_ThisMenuItem%:`n%Link% ,,1
 Run properties %link%
}
Return

SelectAhkExecutable:
SwapAhkExecutable:
 StringSplit, F, A_ThisMenu, |
 PID := F1, hWnd := F2,  PeFile := F3,  Vers := F4, ,  AhkF := ScriptFullPath(F2)
 SplitPath, AhkF,,Work,,NNE

 If ( A_ThisLabel="SelectAhkExecutable" )  {
      FileSelectFile, SelFile, 3, %peFile%, Select AutoHotkey Executable for :: %NNE%
                    , AutoHotkey (*.exe)
      If ( ErrorLevel || SelFile="" )
         Return
      InternalName := SubStr( FileGetVersionInfo( SelFile, "InternalName" ),1,10)
      If ( InternalName <> "AutoHotkey" ) {
         TrayTip, HandyTM, %SelFile%`nis not an AutoHotkey Executable,,3
         Return
      }
      SelVers := FileGetVersionInfo( SelFile, "ProductName" ) " v"
              .  RegExReplace( FileGetVersionInfo( SelFile, "ProductVersion" ), ", ", "." )
      IniWrite,%SelFile%,%DAT%,AhkVersions, %SelVers%`t{%SelFile%}
 }  Else {
      RegExMatch(A_ThisMenuItem,"\{(.*)\}",ID), SelFile:=ID1
      SelVers := FileGetVersionInfo( SelFile, "ProductName" ) " v"
              .  RegExReplace( FileGetVersionInfo( SelFile, "ProductVersion" ), ", ", "." )
 }

 If ( SelVers . SelFile = Vers . peFile ) {
    TrayTip, HandyTM, Source and Target are of same version,, 3
    Return
 }
 CmdL := ProcessGetCommandLine( PID )
 CmdL := StrReplace(CmdL, PeFile, SelFile)
 PostMessage, 0x111, 65405,0,, ahk_id %hWnd%
 Run, %CmdL%, %Work%, UseErrorLevel
Return

ShellOptions:
 RegExMatch(A_ThisMenu,"\{(.*)\}",ID),  hWnd:=ID1, AhkF:=ScriptFullPath( hWnd )
 IfEqual, A_ThisMenuItem, File Properties Dialog , Run, Properties "%AhkF%"
 IfEqual, A_ThisMenuItem, File Properties Dialog (EXE)
   {
     StringSplit, P, A_ThisMenu, |
     PEFile := ProcessGetModuleFileName( P2 )
     Run, Properties "%PEFile%"
   }
 IfEqual, A_ThisMenuItem, Explore Script Folder  , Run Explorer /select`, "%AhkF%",, MAX
 IfEqual, A_ThisMenuItem, Select Script Folder
   {
     SplitPath,AhkF,,OD
     Run Explorer /select`, "%OD%",, MAX
   }
Return

MenuRestore:
 IfEqual,A_ThisMenu,Restore, IniWrite, 0, %DAT%, MenuFlag, %A_ThisMenuItem%
Return

MenuRemove:
 RegExMatch(A_ThisMenu,"\{(.*)\}",ID),hWnd:=ID1, AhkF:=ScriptFullPath(hWnd)
 Menu, LeftClickMenu, Delete, %A_ThisMenu%
 IniPut( True, DAT, "MenuFlag", AhkF )
Return

DeleteSettings:
  FileDelete, %DAT%
  FileAppend, [MenuFlag]`r`n`r`n[LaunchList]`r`n`r`n[AhkVersions]`r`n,%DAT%
Return

ShowLeftClickMenu:
  Menu, ConslidatedMenu, UseErrorLevel
  Menu, Restore, DeleteAll
  Menu, LeftClickMenu, DeleteAll
  Menu, ConslidatedMenu, UseErrorLevel, Off
  If ! ( A:=CreateLeftClickMenu( ) )
       TrayTip, HandyTM, No items to display
  else Menu, LeftClickMenu, Show
Return

QuitScript:
 FileDelete, %A_Temp%\~AHK*.ahk
 ExitApp
Return

SetPriority:
 Process, Priority, %A_ThisMenu%, %A_ThisMenuItem%
Return

ProcessClose:
 StringSplit, P, A_ThisMenu, |
 PID:=P2, AhkF:=P1
 Loop 10 {
           Process, Close, %PID%
           Process, Exist, %PID%
           If ! ( ErrorLevel ) {
           TrayTip, HandyTM (PID:%PID%), Process Killed:`n%AhkF%,5,2
             Break
         }                     }
Return

SetTrayIcon( Gui0=0 ) {
 Hex := "W1Y1Z10101X1Y4Z2801X16V28V1U2T1Y4RCKMFF9C9CZFFA2A2ZFFAAAAZFFB0BYFFB4B4ZFFBABAZFF"
 . "BFBFZFFC4C4ZFFCACAZFFD1D1ZFFD6D6ZFFDDDDZFFE4E4ZFFEBEBZFFFFFFKWBDVCEVADVCEV9DVCEV7BVCE"
 . "V5AVCDV37ABA988ACV246765458AV13V58V12V37V12V25V12V24V22V23KYFFFFXC3C3XC3C3XC3C3XC3C3X"
 . "C3C3XCZ3XCZ3XCZ3XCZ3XC3C3XC3C3XC3C3XC3C3XC3C3XFFFFX"
 VarSetCapacity(Z,512,48), Nums:="512|256|128|64|32|16|15|14|13|12|11|10|9|8|7|6|5|4|3|2"
 Loop, Parse, Nums, |                                  ;  uncompressing nulls in hex data
  hex := StrReplace(hex, Chr(70+A_Index), SubStr(Z,1,A_LoopField))
 VarSetCapacity( IconData,( nSize:=StrLen(Hex)//2) )
 Loop %nSize%
   NumPut( "0x" . SubStr(Hex,2*A_Index-1,2), IconData, A_Index-1, "Char" )
 hICon := DllCall( "CreateIconFromResourceEx", UInt,&IconData+22, UInt,0, Int,1
                   ,UInt,196608, Int,16, Int,16, UInt,0 )
 PID := DllCall("GetCurrentProcessId"), VarSetCapacity( NID,444,0 ), NumPut( 444,NID )
 NumPut( Gui0,NID,4 ), NumPut( 1028,NID,8 ), NumPut( 2,NID,12 ), NumPut( hIcon,NID,20 )
 Menu, Tray, Icon,,, 1
 DllCall( "shell32\Shell_NotifyIcon", UInt,0x1, UInt,&NID )
}

ScriptFullPath( hWnd ) {
 WinGetTitle, T, ahk_id %hWnd%
 Return ( sPos := InStr(T," - AutoHotkey v" ) ) ? SubStr( T,1,sPos-1 ) : T
}

;--  --------  --------  --------  --------  --------  --------     .INI Related Functions

IniReadSection( iniFile, Section ) {
 FileRead, DAT, %iniFile%
 DAT := StrReplace(DAT, "`r`n", "`n")
 sStr:="[" Section "]", sPos:=InStr(DAT,sStr)+(L:=StrLen(sStr)), End:=StrLen(DAT)-sPos
 rStr := SubStr( DAT, sPos, ((E:=InStr(DAT,"`n[",0,sPos)) ? E-sPos : End) )
Return RegExReplace( rStr, "^\n*(.*[^\n])\n*$", "$1", "m" )
}

IniGet( IniFile, Section, Key, def="", t=0 ) {
 IniRead, Value, %IniFile%, %Section%, %Key%, %def%
 IfNotEqual,t,0,Transform, Value, Deref, %Value%
Return Value
}

IniPut( Value, IniFile, Section, Key ) {
 IniWrite, %Value%, %IniFile%, %Section%, %Key%
Return errorLevel
}

;--  --------  --------  --------  --------  --------  --------    Miscellaneous Functions

InitGlobalVars() {
 Global
 SetRegView, % A_Is64bitOS = true ? 64 : 32
 RegRead, InstallDir, HKLM, SOFTWARE\AutoHotkey, InstallDir
 InstallExe := InstallDir "\AutoHotkey.exe"
 A_QuickLaunch := A_AppData . "\Microsoft\Internet Explorer\Quick Launch"
 SplitPath,A_ScriptFullPath,,,,A_Script
 DAT := A_Script ".data"
 IfNotExist, %DAT%, GoSub, DeleteSettings
}

HashStr(Str, L=4, nSz=0) { ; L Bytes, 2*L HEX digits output, 1<=L<=8
  ; by Laszlo Hars, Post: www.autohotkey.com/forum/viewtopic.php?p=231847#231847
  Static HashData, sprintf, S=1234567890123456
  If (HashData="")
  HashData := DllCall( "GetProcAddress", UInt, DllCall( "LoadLibrary", Str,"shlwapi.dll" )
  , Str,"HashData" ), sprintf := DllCall( "GetProcAddress", UInt, DllCall( "LoadLibrary"
  , Str,"msvcrt.dll" ),  Str,"sprintf" )
  DllCall(HashData, UInt,&Str, UInt,nSz<1 ? StrLen(Str):nSz, Int64P,Hash, UInt,8 )
  DllCall(sprintf, Str,S, Str,"%016I64X", Int64,Hash) ; 0-padded (left), 16 HEX digits
  Return SubStr(S, 1-2*L) ; only 2L digits
}

FileGetVersionInfo( peFile="", StringFileInfo="" ) {
 FSz := DllCall( "Version\GetFileVersionInfoSizeA",Str,peFile,UInt,0 )
 IfLess, FSz,1, Return -1
 VarSetCapacity( FVI, FSz, 0 ), VarSetCapacity( Trans,8 )
 DllCall( "Version\GetFileVersionInfoA", Str,peFile, Int,0, Int,FSz, UInt,&FVI )
 If ! DllCall( "Version\VerQueryValueA", UInt,&FVI, Str,"\VarFileInfo\Translation"
                                       , UIntP,Translation, UInt,0 )
   Return -2
 If ! DllCall( "msvcrt.dll\sprintf", Str,Trans, Str,"%08X", UInt,NumGet(Translation+0) )
   Return -4
 subBlock := "\StringFileInfo\" SubStr(Trans,-3) SubStr(Trans,1,4) "\" StringFileInfo
 If ! DllCall( "Version\VerQueryValueA", UInt,&FVI, Str,SubBlock, UIntP,InfoPtr, UInt,0 )
   Return
 VarSetCapacity( Info, DllCall( "lstrlen", UInt,InfoPtr ) )
 DllCall( "lstrcpy", Str,Info, UInt,InfoPtr )
Return Info
}

uHex( U=0, L=8 ) {
  Static sprintf, Hex
   If ( sprintf="" ) {
     If ! ( hModule := DllCall( "GetModuleHandle", Str,"msvcrt.dll" ) )
            hModule := DllCall( "LoadLibrary", Str,"msvcrt.dll" )
     sprintf := DllCall( "GetProcAddress", UInt,hModule, Str,"sprintf" )
     VarSetCapacity( hex,10 )
} DllCall( sprintf, Str,Hex, Str,"0x%0" L "X", UInt,U )
Return Hex
}

;--  --------  --------  --------  --------  --------  --------  Path Related Functions

PathRemoveArgs( sPath ) {
  DllCall( "shlwapi.dll\PathRemoveArgsA", Str,sPath )
Return sPath
}

PathRemoveBlanks( sPath ) {
  DllCall( "shlwapi.dll\PathRemoveBlanksA", Str,sPath )
Return sPath
}

PathUnquoteSpaces( sPath ) {
  Return DllCall( "shlwapi.dll\PathUnquoteSpacesA", Str,sPath, Str )
}

PathQuoteSpaces( sPath ) {
  DllCall( "shlwapi.dll\PathQuoteSpacesA", Str,sPath )
Return sPath
}

;--  --------  --------  --------  --------  --------  --------  Process Related Functions

SetDebugPrivilege() {
 ;PROCESS_QUERY_INFORMATION=0x400, TOKEN_ADJUST_PRIVILEGES=0x20, SE_PRIVILEGE_ENABLED:=0x2
 hProcess := DllCall( "OpenProcess", UInt,0x400,Int,0,UInt,DllCall("GetCurrentProcessId"))
 DllCall( "Advapi32.dll\LookupPrivilegeValueA", UInt,0, Str,"SeDebugPrivilege", UIntP,lu )
  ; TOKEN_PRIVILEGES Structure : www.msdn.microsoft.com/en-us/library/aa379630(VS.85).aspx
 VarSetCapacity( TP,16,0), NumPut( 1,TP,0,4 ),  NumPut( lu,TP,4,8 ), NumPut( 0x2,TP,12,4 )
 DllCall( "Advapi32.dll\OpenProcessToken", UInt,hProcess, UInt,0x20, UIntP,hToken )
 Result :=  DllCall( "Advapi32.dll\AdjustTokenPrivileges"
                 , UInt,hToken, UInt,0, UInt,&TP, UInt,0, UInt,0, UInt,0 )
 DllCall( "CloseHandle", UInt,hProcess ), DllCall( "CloseHandle", UInt,hToken )
Return Result
}

ProcessOpen( PID, da=0x43a, ihnd=0 ) {
;  Desired Access 0x43a := PROCESS_QUERY_INFORMATION=0x400 + PROCESS_CREATE_THREAD=0x2
;     + PROCESS_VM_READ=0x10 + PROCESS_VM_OPERATION=0x8 + PROCESS_VM_WRITE=0x20
 Return DllCall( "OpenProcess", UInt,da, UInt,ihnd, UInt,PID )
}

ProcessClose( Byref nProc ) {
 Return DllCall( "CloseHandle", UInt,nProc ) + VarSetCapacity( nProc,0 )
}

;       :: Note :: If PID is False nProc is a open Process Handle else nProc is PID itself

ProcessGetOwner( nProc, PID=1 ) {
 ; PROCESS_QUERY_INFORMATION=0x400, TOKEN_READ:=0x20008 , TokenUser:=0x1
 IfNotEqual,PID,0, SetEnv, nProc, % DllCall( "OpenProcess", UInt,0x400,Int,0,UInt,nProc )
 DllCall( "Advapi32.dll\OpenProcessToken", UInt,nProc, UInt,0x20008, UIntP,Tok )
 DllCall( "Advapi32.dll\GetTokenInformation", UInt,Tok, UInt,0x1, Int,0, Int,0, UIntP,RL )
 VarSetCapacity( TI,RL,0 )
 DllCall( "Advapi32.dll\GetTokenInformation"
          , UInt,Tok, UInt,0x1, UInt,&TI, Int,RL, UIntP,RL ),           pSid := NumGet(TI)
 IfNotEqual,PID,0, SetEnv, nProc, % DllCall( "CloseHandle", UInt,nProc )
 DllCall( "CloseHandle", UInt,Tok )
 ; following code taken from www.autohotkey.com/forum/viewtopic.php?p=116487 - Author Sean
 DllCall( "Advapi32\LookupAccountSidA"
         , Str,"", UInt,pSid, UInt,0, UIntP,nSizeNM, UInt,0, UIntP,nSizeRD, UIntP,eUser )
 VarSetCapacity( sName,nSizeNM,0 ), VarSetCapacity( sRDmn,nSizeRD,0 )
 DllCall( "Advapi32\LookupAccountSidA"
    , Str,"", UInt,pSid, Str,sName, UIntP,nSizeNM, Str,sRDmn, UIntP,nSizeRD, UIntP,eUser )
 ;DllCall( "LocalFree", UInt,pSid ) ; Crashes AHK on both 32Bit and 64Bit Systems
Return sName
}

ProcessGetPriority( nProc, PID=1 ) {
 IfNotEqual,PID,0, SetEnv, nProc, % DllCall( "OpenProcess", UInt,0x400,Int,0,UInt,nProc )
 P := DllCall( "GetPriorityClass", Int,nProc )
 IfNotEqual,PID,0, SetEnv, nProc, % DllCall( "CloseHandle", UInt,nProc )
 Return P=64 ? "Low" : P=16384 ? "BelowNormal" : P=32 ? "Normal" : P=32768 ? "AboveNormal"
       : P=128 ? "High" : P=256 ? "Realtime" : ""
}

ProcessGetModuleFileName( nProc, PID=1 ) {
	if(A_Is64BitOS){ ; www.autohotkey.com/docs/commands/ComObjGet.htm
		queryEnum := ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process where ProcessId=" PID)._NewEnum()
		if queryEnum[process]
			fname := process.ExecutablePath
		wmi := queryEnum := process := ""
	}else{
		IfNotEqual,PID,0, SetEnv,nProc,% DllCall("OpenProcess", UInt,0x410,Int,0,UInt,nProc)
		nSz := VarSetCapacity( fName,260,0 )
		DllCall( "psapi.dll\GetModuleFileNameExA", UInt,nProc, UInt,0, Str,fName, UInt,nSz )
		IfNotEqual,PID,0, SetEnv,nProc,% DllCall( "CloseHandle", UInt,nProc )
	}
	Return, fname
}

ProcessGetCommandLine( nProc, PID=1 ) {
	if(A_Is64BitOS){ ; www.autohotkey.com/docs/commands/ComObjGet.htm
		If (PID = "1")
			PID := nProc
		queryEnum := ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process where ProcessId=" PID)._NewEnum()
		if queryEnum[process]
			sCmdLine := process.CommandLine
		wmi := queryEnum := process := ""
	}else{ ; www.autohotkey.com/forum/viewtopic.php?t=16575 Author Sean
		Static pFunc
		If pFunc=
			pFunc := DllCall( "GetProcAddress", UInt
				, DllCall( "GetModuleHandle", Str,"kernel32.dll" ), Str,"GetCommandLineA" )
		IfNotEqual,PID,0, SetEnv, nProc, % DllCall( "OpenProcess", UInt,0x43a,Int,0,UInt,nProc )
		hThrd := DllCall( "CreateRemoteThread", UInt,nProc, UInt,0, UInt,0, UInt,pFunc, UInt,0, UInt,0, UInt,0 ),  DllCall( "WaitForSingleObject", UInt,hThrd, UInt,0xFFFFFFFF )
		DllCall( "GetExitCodeThread", UInt,hThrd, UIntP,pcl ), VarSetCapacity( sCmdLine,512 )
		DllCall( "ReadProcessMemory", UInt,nProc, UInt,pcl, Str,sCmdLine, UInt,512, UInt,0 )
		DllCall( "CloseHandle", UInt,hThrd )
		IfNotEqual,PID,0, SetEnv, nProc, % DllCall( "CloseHandle", UInt,nProc )
	}
 Return RegExReplace( sCmdLine, "(^\s*|\s*$)" )
}

ProcessGetCreationTime( nProc, PID=1 ) {
 VarSetCapacity( PT,16 ), VarSetCapacity( ST,16 ), u:="UShort" z:=0
 IfNotEqual,PID,0, SetEnv,nProc,% DllCall( "OpenProcess", UInt,0x400,Int,0,UInt,nProc )
 DllCall( "GetProcessTimes" , UInt,nProc, UInt,&PT, UInt,0, UInt,0, UInt,0 )
 IfNotEqual,PID,0, SetEnv,nProc,% DllCall( "CloseHandle", UInt,nProc )
 DllCall( "FileTimeToLocalFileTime", UInt,&PT, UInt,&PT )
 DllCall( "FileTimeToSystemTime"   , UInt,&PT, UInt,&ST )
 Return NumGet(ST,0,U) . SubStr("0" NumGet(ST, 2,U),-1) . SubStr("0" NumGet(ST, 6,U),-1)
                       . SubStr("0" NumGet(ST, 8,U),-1) . SubStr("0" NumGet(ST,10,U),-1)
                       . SubStr("0" NumGet(ST,12,U),-1) . SubStr("0" NumGet(ST,14,U),-1)
}
