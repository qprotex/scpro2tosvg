#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         Miguel Febres <mfebres@q-protex.com>

 Script Function:
	Converts SignCut scpro2 files to SVG using SignCut DRAW

#ce ----------------------------------------------------------------------------

#include <Array.au3>
#include <File.au3>
#include <WinAPI.au3>
#include <WindowsConstants.au3>

Opt("WinTitleMatchMode", 2)

$sSignCut2Path = IniRead('config.ini', "General", "SignCut2Path", "")
$sSourceFolder = IniRead('config.ini', "General", "SourceFolder", "")


ConsoleWrite("scpro2 to svg converter v0.2" & @LF)
ConsoleWrite("============================" & @LF & @LF)

 $aArray = _FileListToArrayRec($sSourceFolder, "*scpro2;*.sc", $FLTAR_FILES, $FLTAR_RECUR, $FLTAR_SORT)
 ;_ArrayDisplay($aArray, ".EXE Files")
 $filesCount=0
 For $i = 1 to UBound($aArray) - 2
    $fileName = $sSourceFolder & $aArray[$i]
   $newFileName = StringReplace($fileName, ".scpro2", ".svg")
   $newFileName = StringReplace($newFileName, ".sc", ".svg")
   if not FileExists($newFileName) then
	  ConsoleWrite("Converting file: " &$fileName & @LF)
	  Run($sSignCut2Path & "\SignCut.exe /FILE " & """" & $fileName & """") ;
	  $hWnd = WinWait("SignCut Productivity Pro 2 - Premium")

	  Sleep(2000)
	  $hToolBar = ControlGetHandle($hWnd, "", "[CLASS:wxWindowNR; INSTANCE:2]")
	  WinActivate($hToolBar)
	  ControlClick($hWnd, "", $hToolBar, "left",1, 150, 20)

	  $hWnd2 = WinWait(" - SignCut DRAW")
	  Sleep(2000)

	  $menu = ControlGetHandle($hWnd2, "", "[CLASS:SCDRAWMenu]")
	  WinActivate($menu)
	  $aWindowsPos = WinGetPos($hWnd2)
	  $aControlPos = ControlGetPos($hWnd2, "", $menu)
	  MouseClick($MOUSE_CLICK_LEFT,$aWindowsPos[0]+$aControlPos[0]+20, $aWindowsPos[1]+$aControlPos[1]+40, 1)
	  Sleep(1000)
	  Send("{DOWN}")
	  Send("{DOWN}")
	  Send("{DOWN}")
	  Send("{DOWN}")
	  Send("{DOWN}")
	  Send("{DOWN}")
	  Send("{DOWN}")
	  Send("{DOWN}")
	  Send("{ENTER}")
	  Sleep(1000)
	  $hWndE = WinWait("Export")
	  Sleep(1000)
	  Send($newFileName)
	  Send("{ENTER}")
	  $filesCount += 1

   Else
	  ConsoleWrite("Skipping existing file: " & $newFileName & @LF )
   EndIf
Next

ConsoleWrite("Converted files: " & $filesCount & @LF )