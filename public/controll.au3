#include-once
#include <StringConstants.au3>
#include <MsgBoxConstants.au3>
Func _PPT_Open($bVisible = Default)
	Local $oPPT
	If $bVisible = Default Then $bVisible = True
	$oPPT = ObjGet("", "PowerPoint.Application")
	If @error Then
		$oPPT = ObjCreate("PowerPoint.Application")
		If @error Then Return SetError(1, @error, 0)
	EndIf
	$oPPT.Visible = $bVisible
	Return $oPPT
 EndFunc   ;==>_PPT_Open

 Func _PPT_PresentationOpen($oPPT, $sFilePath, $bReadOnly = Default, $bVisible = Default)
	If Not IsObj($oPPT) Or ObjName($oPPT, 1) <> "_Application" Then Return SetError(1, @error, 0)
	If Not FileExists($sFilePath) Then Return SetError(2, 0, 0)
	If $bReadOnly = Default Then $bReadOnly = False
	If $bVisible = Default Then $bVisible = True
	Local $oPresentation = $oPPT.Presentations.Open($sFilePath, $bReadOnly, Default, $bVisible)
	If @error Then Return SetError(3, @error, 0)
	; If a read-write presentation was opened read-only then return an error
	If $bReadOnly = False And $oPresentation.Readonly = True Then Return SetError(0, 1, $oPresentation)
	Return $oPresentation
 EndFunc   ;==>_PPT_PresentationOpen
 Func _PPT_Close(ByRef $oPPT, $bSaveChanges = Default)
	If $bSaveChanges = Default Then $bSaveChanges = True
	If Not IsObj($oPPT) Or ObjName($oPPT, 1) <> "_Application" Then Return SetError(1, 0, 0)
	If $bSaveChanges Then
		For $oPresentation In $oPPT.Presentations
			If Not $oPresentation.Saved Then
				$oPresentation.Save()
				If @error Then Return SetError(3, @error, 0)
			EndIf
		Next
	EndIf
	$oPPT.Quit()
	If @error Then Return SetError(2, @error, 0)
	$oPPT = 0
	Return 1
 EndFunc   ;==>_PPT_Close




; Create application object and open an example presentation
$Dat = ''
If $CmdLine[0] > 0 Then
    If $CmdLine[1] <> @ScriptName Then
        $Dat = $CmdLine[1]
                Endif
Endif
Local $sPresentation = @ScriptDir & $Dat
Local $oPPT = _PPT_Open()

If @error Then Exit MsgBox($MB_SYSTEMMODAL, "PowerPoint UDF: _PPT_PresentationExport Example", "Error creating the PowerPoint application object." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
Local $oPresentation = _PPT_PresentationOpen($oPPT, $sPresentation)

If @error Then
	MsgBox($MB_SYSTEMMODAL, "PowerPoint UDF: _PPT_PresentationExport Example", "Error opening presentation '" & $sPresentation & "'." & @CRLF & "@error = " & @error & ", @extended = " & @extended)
	_PPT_Close($oPPT)
	Exit
EndIf

; *********************************************************************
; Export slides 2 and 3 of the presentation, format: PDF, output: Notes
; *********************************************************************
   $title = StringReplace($Dat, "\", "")
   $title = $title &' - PowerPoint'
    WinActive($title)
ControlClick( $title,"","[CLASS:mdiClass; INSTANCE:1]","left",1,159, 300)
ControlSend($title, "" , "" ,"!4")
Sleep(3000)

ControlClick ("Publish Presentation", "Publish", "[CLASS:Button; INSTANCE:112]" , "left" , 1 , 40, 12)
Sleep(30000)
Winclose("Presentation Preview")
ControlSend($title, "" , "" ,"^s")
Winclose($title)