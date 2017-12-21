#cs ----------------------------------------------------------------------------

 Title: DIRF AutoBackup
 Version: 0.1-beta
 AutoIt Version: 3.3.14.2
 Author:         Eduardo Mozart de Oliveira

 Script Function:
	Backup autom�tico de declara��es da DIRF.

#ce ----------------------------------------------------------------------------

#include <MsgBoxConstants.au3>
#include "AutoIt-FSOClass\FSOClass.au3"

AutoItSetOption("MouseCoordMode", 2)
AutoItSetOption("TrayIconDebug", 1)
;OnAutoItExitRegister("OnAutoItExit")
HotKeySet("{ESCAPE}", "OnAutoItExit")

$DirfInstallDir = _FileSelectFolder()

$filelist = List($DirfInstallDir, "Dirf*.exe")
ConsoleWrite("$filelist.Count-1: " & $filelist.Count-1)
If $filelist.Count-1 < 0 Then
   MsgBox($MB_SYSTEMMODAL, "", "O execut�vel da DIRF n�o p�de ser encontrado na pasta selecionada.")
   Exit 1
EndIf

For $n = 0 To $filelist.Count-1
   $DirfExe = $filelist.Item($n)
Next


If ProcessExists($DirfExe) Then
   ProcessClose($DirfExe)
   ProcessWaitClose(DirfExe)
EndIf

; Excluir c�pia de seguran�a anteriores.
$filelist = List($DirfInstallDir & "\CopSeg", "*.dbk")
For $n = 0 To $filelist.Count-1
   ConsoleWrite( "(" & $n+1 & " de " & $filelist.Count & ") " & $filelist.Item($n) & @CRLF)
   FileDelete($filelist.Item($n))
Next

Run($DirfInstallDir & "\" & $DirfExe)

WinWaitActive("[CLASS:TfmAvisoTeste]")
Send("!f") ; Fechar anima��o

WinWaitActive("[CLASS:TFoInicial]") ; Dirf 2013
Send("!c") ; Cancelar


$Declaracao = 0
While 1
   Send("{ALTDOWN}")
   Send("f") ; Ferramentas
   Send("i") ; C�pia de Seguran�a
   Send("g") ; Gravar
   Send("{ALTUP}")

   WinWaitActive("Grava��o de c�pia de seguran�a")
   Send("{DOWN}") ; C�pia de Seguran�a de uma Declara��o
   Send("{ENTER}") ; Avan�ar


   Send("{TAB}") ; CPF/CNPJ
   Send("{TAB}") ; Ano
   For $n = 0 To $Declaracao-1
	  Send("{DOWN}")
   Next
   Send("{ENTER}") ; Avan�ar


   Send("{ENTER}") ; Avan�ar
   $Continuar = True
   While 1
	 Sleep(100)
	 ; Esta declara��o j� foi gravada. Deseja Sobrescrever?
     If WinExists("Confirma��o") Then
	    Send("!n") ; N�o
		$Continuar = False
		ExitLoop
	 EndIf

	 $pos = ControlGetPos("Grava��o de c�pia de seguran�a", "", "[CLASS:TPanel; INSTANCE:14]") ; Concluir
	 MouseClick("left", $pos[0], $pos[1])
     If Not WinExists("Grava��o de c�pia de seguran�a") Then
		ExitLoop
     EndIf
   WEnd

   If Not $Continuar Then ExitLoop

   $Declaracao = $Declaracao + 1
WEnd

MsgBox($MB_ICONINFORMATION, "", "Declara��es exportadas com sucesso.")

Func OnAutoItExit()
   Exit
EndFunc

Func _FileSelectFolder()
    ; Create a constant variable in Local scope of the message to display in FileSelectFolder.
    Local Const $sMessage = "Selecione uma pasta"

    ; Display an open dialog to select a file.
    Local $sFileSelectFolder = FileSelectFolder($sMessage, "")
    If @error Then
        ; Display the error message.
        MsgBox($MB_SYSTEMMODAL, "", "Nenhuma pasta foi selecionada.")
		Exit 1
    Else
        ; Display the selected folder.
        ; MsgBox($MB_SYSTEMMODAL, "", "Voc� escolheu a seguinte pasta:" & @CRLF & $sFileSelectFolder)
    EndIf

	Return $sFileSelectFolder
EndFunc   ;==>Example
