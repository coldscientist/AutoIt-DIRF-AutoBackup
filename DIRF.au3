#cs ----------------------------------------------------------------------------

 Title: DIRF AutoBackup
 Version: 0.1-beta
 AutoIt Version: 3.3.14.2
 Author:         Eduardo Mozart de Oliveira

 Script Function:
	Backup automático de declarações da DIRF.

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
   MsgBox($MB_SYSTEMMODAL, "", "O executável da DIRF não pôde ser encontrado na pasta selecionada.")
   Exit 1
EndIf

For $n = 0 To $filelist.Count-1
   $DirfExe = $filelist.Item($n)
Next


If ProcessExists($DirfExe) Then
   ProcessClose($DirfExe)
   ProcessWaitClose(DirfExe)
EndIf

; Excluir cópia de segurança anteriores.
$filelist = List($DirfInstallDir & "\CopSeg", "*.dbk")
For $n = 0 To $filelist.Count-1
   ConsoleWrite( "(" & $n+1 & " de " & $filelist.Count & ") " & $filelist.Item($n) & @CRLF)
   FileDelete($filelist.Item($n))
Next

Run($DirfInstallDir & "\" & $DirfExe)

WinWaitActive("[CLASS:TfmAvisoTeste]")
Send("!f") ; Fechar animação

WinWaitActive("[CLASS:TFoInicial]") ; Dirf 2013
Send("!c") ; Cancelar


$Declaracao = 0
While 1
   Send("{ALTDOWN}")
   Send("f") ; Ferramentas
   Send("i") ; Cópia de Segurança
   Send("g") ; Gravar
   Send("{ALTUP}")

   WinWaitActive("Gravação de cópia de segurança")
   Send("{DOWN}") ; Cópia de Segurança de uma Declaração
   Send("{ENTER}") ; Avançar


   Send("{TAB}") ; CPF/CNPJ
   Send("{TAB}") ; Ano
   For $n = 0 To $Declaracao-1
	  Send("{DOWN}")
   Next
   Send("{ENTER}") ; Avançar


   Send("{ENTER}") ; Avançar
   $Continuar = True
   While 1
	 Sleep(100)
	 ; Esta declaração já foi gravada. Deseja Sobrescrever?
     If WinExists("Confirmação") Then
	    Send("!n") ; Não
		$Continuar = False
		ExitLoop
	 EndIf

	 $pos = ControlGetPos("Gravação de cópia de segurança", "", "[CLASS:TPanel; INSTANCE:14]") ; Concluir
	 MouseClick("left", $pos[0], $pos[1])
     If Not WinExists("Gravação de cópia de segurança") Then
		ExitLoop
     EndIf
   WEnd

   If Not $Continuar Then ExitLoop

   $Declaracao = $Declaracao + 1
WEnd

MsgBox($MB_ICONINFORMATION, "", "Declarações exportadas com sucesso.")

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
        ; MsgBox($MB_SYSTEMMODAL, "", "Você escolheu a seguinte pasta:" & @CRLF & $sFileSelectFolder)
    EndIf

	Return $sFileSelectFolder
EndFunc   ;==>Example
