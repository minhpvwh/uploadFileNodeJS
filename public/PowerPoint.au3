#include-once
#include <StringConstants.au3>
#include <PowerPointConstants.au3>

; #INDEX# =======================================================================================================================
; Title .........: Microsoft PowerPoint Function Library
; AutoIt Version : 3.3.12.0
; UDF Version ...: Alpha 5
; Language ......: English
; Description ...: A collection of functions for accessing and manipulating Microsoft PowerPoint files
; Author(s) .....: water
; Modified.......: 20170606 (YYYMMDD)
; Remarks .......: Based on the UDF written by toady (see Links)
; Link ..........: https://www.autoitscript.com/forum/topic/50254-powerpoint-wrapper
; Contributors ..:
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
Global $__iPPT_Debug = 0 ; Debug level. 0 = no debug information, 1 = Debug info to console, 2 = Debug info to MsgBox, 3 = Debug Info to File
Global $__sPPT_DebugFile = @ScriptDir & "\PowerPoint_Debug.txt" ; Debug file if $__iOL_Debug is set to 3
Global $__oPPT_Error ; COM Error handler
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
;_PPT_Open
;_PPT_Close
;_PPT_ErrorNotify
;_PPT_PresentationAttach
;_PPT_PresentationClose
;_PPT_PresentationExport
;_PPT_PresentationExportGraphic
;_PPT_PresentationList
;_PPT_PresentationNew
;_PPT_PresentationOpen
;_PPT_PresentationPrint
;_PPT_PresentationSave
;_PPT_PresentationSaveAs
;_PPT_SlideAdd
;_PPT_SlideCopyMove
;_PPT_SlideDelete
;_PPT_SlideShow
;_PPT_VersionInfo

;_PPT_SlideTextFrameSetText()
;_PPT_SlideTextFrameGetText()
;_PPT_SlideTextFrameSetFont()
;_PPT_SlideTextFrameSetFontSize()
;_PPT_SlideShapeCount()
;_PPT_SlideAddPicture()
;_PPT_SlideAddTable()
;_PPT_SlideAddTextBox()
;
; Functions - Slide Show Config Settings
; - - - - - - - - - - - - - - - - - - - - - -
;_PPT_SlideShowWithAnimation()
;_PPT_SlideShowWithNarration()
;_PPT_SlideShowAdvanceMode()
;_PPT_SlideShowAdvanceOnTime()
;_PPT_SlideShowAdvanceTime()
; ===============================================================================================================================

; #INTERNAL_USE_ONLY#============================================================================================================
;__PPT_ErrorHandler
;__PPT_SliderangeCreate
; ===============================================================================================================================

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_Open
; Description ...: Connects to an existing PowerPoint instance or creates a new one
; Syntax.........: _PPT_Open([$bVisible = True])
; Parameters ....: $bVisible - [optional] True specifies that the application will be visible (default = True).
; Return values .: Success - the PowerPoint application object.
;                  Failure - 0 and sets @error.
;                  |1 - Error returned by ObjCreate. @extended is set to the COM error code
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......: _PPT_Close
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
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

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_Close
; Description ...: Closes all presentations and the instance of the PowerPoint application
; Syntax.........: _PPT_Close($oPPT[, $iSaveChanges = True])
; Parameters ....: $oPPT         - PowerPoint application object as returned by _PPT_Open.
;                  $bSaveChanges - [optional] Specifies whether changed presentations should be saved before closing (default = True).
; Return values .: Success - 1.
;                  Failure - 0 and sets @error.
;                  |1 - $oPPT is not an object or not an application object
;                  |2 - Error returned by method Application.Quit. @extended is set to the COM error code
;                  |3 - Error returned by method Application.Save. @extended is set to the COM error code
; Author ........: water
; Modified ......:
; Remarks .......: _PPT_Close closes all presentations (even those opened manually by the user for this instance after _PPT_Open)
;                  and the specified PowerPoint instance.
; Related .......: _PPT_Open
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
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

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_ErrorNotify
; Description ...: Sets or queries the debug level
; Syntax.........: _PPT_ErrorNotify($iDebug[, $sDebugFile = @ScriptDir & "\PowerPoint_Debug.txt"])
; Parameters ....: $iDebug     - Debug level. Allowed values are:
;                  |-1 - Return the current settings
;                  |0  - Disable debugging
;                  |1  - Enable debugging. Output the debug info to the console
;                  |2  - Enable Debugging. Output the debug info to a MsgBox
;                  |3  - Enable Debugging. Output the debug info to a file defined by $sDebugFile
;                  |4  - Enable Debugging. The COM errors will be handled (the script no longer crashes) without any output
;                  $sDebugFile - Optional: File to write the debugging info to if $iDebug = 3 (Default = @ScriptDir & "\PowerPoint_Debug.txt")
; Return values .: Success (for $iDebug => 0) - 1, sets @extended to:
;                  |0 - The COM error handler for this UDF was already active
;                  |1 - A COM error handler has been initialized for this UDF
;                  Success (for $iDebug = -1) - one based one-dimensional array with the following elements:
;                  |1 - Debug level. Value from 0 to 3. Check parameter $iDebug for details
;                  |2 - Debug file. File to write the debugging info to as defined by parameter $sDebugFile
;                  |3 - True if the COM error handler has been defined for this UDF. False if debugging is set off or a COM error handler was already defined
;                  Failure - 0, sets @error to:
;                  |1 - $iDebug is not an integer or < -1 or > 4
;                  |2 - Installation of the custom error handler failed. @extended is set to the error code returned by ObjEvent
;                  |3 - COM error handler already set to another function
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_ErrorNotify($iDebug, $sDebugFile = Default)
	If $sDebugFile = Default Then $sDebugFile = ""
	If Not IsInt($iDebug) Or $iDebug < -1 Or $iDebug > 4 Then Return SetError(1, 0, 0)
	If $sDebugFile = "" Then $sDebugFile = @ScriptDir & "\PowerPoint_Debug.txt"
	Switch $iDebug
		Case -1
			Local $avDebug[4] = [3]
			$avDebug[1] = $__iPPT_Debug
			$avDebug[2] = $__sPPT_DebugFile
			$avDebug[3] = IsObj($__oPPT_Error)
			Return $avDebug
		Case 0
			$__iPPT_Debug = 0
			$__sPPT_DebugFile = ""
			$__oPPT_Error = 0
		Case Else
			$__iPPT_Debug = $iDebug
			$__sPPT_DebugFile = $sDebugFile
			; A COM error handler will be initialized only if one does not exist
			If ObjEvent("AutoIt.Error") = "" Then
				$__oPPT_Error = ObjEvent("AutoIt.Error", "__PPT_ErrorHandler") ; Creates a custom error handler
				If @error <> 0 Then Return SetError(2, @error, 0)
				Return SetError(0, 1, 1)
			ElseIf ObjEvent("AutoIt.Error") = "__PPT_ErrorHandler" Then
				Return SetError(0, 0, 1) ; COM error handler already set by a call to this function
			Else
				Return SetError(3, 0, 0) ; COM error handler already set to another function
			EndIf
	EndSwitch
	Return 1
EndFunc   ;==>_PPT_ErrorNotify

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_PresentationAttach
; Description ...: Attaches to the frist presentation where the search string matches based on the selected mode
; Syntax.........: _PPT_PresentationAttach($sString[, $sMode = "FilePath"[, $bPartialMatch = False]])
; Parameters ....: $sString       - String to search for.
;                  $sMode         - [optional] specifies search mode:
;                  |FileName      - Name of the open presentation
;                  |FilePath      - Full path to the open presentation (default)
;                  |Title         - Title of the PowerPoint window
;                  $bPartialMatch - [optional] When $sMode = Title then $sString must fully match when False (default) or partial if True
; Return values .: Success - the PowerPoint presentation object.
;                  Failure - 0 and sets @error.
;                  |1 - An error occurred. @extended is set to the COM error code
;                  |2 - $sMode is invalid
;                  |3 - $sString can't be found in any of the open presentations
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......: _PPT_PresentationClose, _PPT_PresentationNew, _PPT_PresentationOpen
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_PresentationAttach($sString, $sMode = Default, $bPartialMatch = Default)
	Local $oPresentation, $iCount = 0, $sCLSID_Presentation = "{91493444-5A91-11CF-8700-00AA0060263B}" ; Microsoft.Office.Interop.PowerPoint.PresentationClass
	If $sMode = Default Then $sMode = "FilePath"
	If $bPartialMatch = Default Then $bPartialMatch = False
	While True
		$oPresentation = ObjGet("", $sCLSID_Presentation, $iCount + 1)
		If @error Then Return SetError(1, @error, 0)
		Switch $sMode
			Case "filename"
				If $oPresentation.Name = $sString Then Return $oPresentation
			Case "filepath"
				If $oPresentation.FullName = $sString Then Return $oPresentation
			Case "title"
				If $bPartialMatch Then
					If StringInStr($oPresentation.Application.Caption, $sString) > 0 Then Return $oPresentation
				Else
					If $oPresentation.Application.Caption = $sString Then Return $oPresentation
				EndIf
			Case Else
				Return SetError(2, 0, 0)
		EndSwitch
		$iCount += 1
	WEnd
	Return SetError(3, 0, 0)
EndFunc   ;==>_PPT_PresentationAttach

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_PresentationClose
; Description ...: Closes the specified presentation
; Syntax.........: _PPT_PresentationClose($oPresentation[, $bSave = True])
; Parameters ....: $oPresentation - Presentation object.
;                  $bSave         - [optional] If True the presentation will be saved before closing (default = True).
; Return values .: Success - 1.
;                  Failure - 0 and sets @error.
;                  |1 - $oPresentation is not an object or not a presentation object
;                  |2 - Error occurred when saving the presentation. @extended is set to the COM error code returned by the Save method
;                  |3 - Error occurred when closing the presentation. @extended is set to the COM error code returned by the Close method
; Author ........: water
; Modified.......:
; Remarks .......: None
; Related .......: _PPT_PresentationAttach, _PPT_PresentationNew, _PPT_PresentationOpen
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_PresentationClose(ByRef $oPresentation, $bSave = Default)
	If Not IsObj($oPresentation) Or ObjName($oPresentation, 1) <> "_Presentation" Then Return SetError(1, 0, 0)
	If $bSave = Default Then $bSave = True
	If $bSave And Not $oPresentation.Saved Then
		$oPresentation.Save()
		If @error Then Return SetError(2, @error, 0)
	EndIf
	$oPresentation.Close()
	If @error Then Return SetError(3, @error, 0)
	$oPresentation = 0
	Return 1
EndFunc   ;==>_PPT_PresentationClose

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_PresentationExport
; Description ...: Export one/multiple/all slides of a presentation as PDF or XPS
; Syntax.........: _PPT_PresentationExport($oPresentation, $sPath[, $sSlide = Default[, $iFixedFormatType = $ppFixedFormatTypePDF[, $iOutputType = $ppFixedFormatTypePDF[, $bUseISO19005 = True)]]]])
; Parameters ....: $oPresentation    - Presentation object.
;                  $sPath            - Path/name of the exported file.
;                  $sSlide           - [optional] A string with the index number of the starting (and ending) slide to be exported (, separated by a hyphen) (default = Keyword Default = export all slides in the presentation).
;                  $iFixedFormatType - [optional] The format to which the slides should be exported. Can be any value of the PpFixedFormatType enumeration (default = $ppFixedFormatTypePDF).
;                  $iOutputType      - [optional] The type of output. Can be any value of the PpPrintOutputType enumeration (default = $ppPrintOutputSlides).
;                  $bUseISO19005     - [optional] Whether the resulting document is compliant with ISO 19005-1 (PDF/A) (default = True).
; Return values .: Success - 1.
;                  Failure - 0 and sets @error.
;                  |1 - $oPresentation is not an object or not a presentation object
;                  |2 - $sSlide is an object but not a sliderange object
;                  |3 - $sPath is empty
;                  |4 - Error occurred when exporting the presentation. @extended is set to the COM error code returned by the ExportAsFixedFormat method
;                  |5 - $sSlide is invalid. Has to be "StartingSlide-EndingSlide"
; Author ........: water
; Modified.......:
; Remarks .......: Method ExportAsFixedFormat only supports a single range of consecutive slides.
;                  To export a single slide simply pass the slide number as $sSlide.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_PresentationExport(ByRef $oPresentation, $sPath, $sSlide = Default, $iFixedFormatType = Default, $iOutputType = Default, $bUseISO19005 = Default)
	If Not IsObj($oPresentation) Or ObjName($oPresentation, 1) <> "_Presentation" Then Return SetError(1, 0, 0)
	If $sSlide = Default Then
		$oPrintRange = $oPresentation.PrintOptions.Ranges.Add(1, $oPresentation.Slides.Count)
	Else
		Local $aSlides = StringSplit($sSlide, "-", $STR_NOCOUNT)
		Local $iValues = UBound($aSlides)
		If $iValues < 1 Or $iValues > 2 Then
			Return SetError(5, 0, 0)
		ElseIf $iValues = 1 Then
			ReDim $aSlides[2]
			$aSlides[1] = $aSlides[0]
		EndIf
		If Number($aSlides[0]) = 0 Or Number($aSlides[1]) = 0 Then Return SetError(5, 0, 0)
		$oPrintRange = $oPresentation.PrintOptions.Ranges.Add($aSlides[0], $aSlides[1])
	EndIf
	If $sPath = "" Then Return SetError(3, 0, 0)
	If $iFixedFormatType = Default Then $iFixedFormatType = $ppFixedFormatTypePDF
	If $iOutputType = Default Then $iOutputType = $ppPrintOutputSlides
	If $bUseISO19005 = Default Then $bUseISO19005 = True
	$oPresentation.ExportAsFixedFormat($sPath, $iFixedFormatType, Default, Default, Default, $iOutputType, Default, $oPrintRange, $ppPrintSlideRange, Default, Default, Default, Default, Default, $bUseISO19005, Default)
	If @error Then Return SetError(4, @error, 0)
	Return 1
EndFunc   ;==>_PPT_PresentationExport

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_PresentationExportGraphic
; Description ...: Export one/multiple/all slides of a presentation in a graphic format
; Syntax.........: _PPT_PresentationExportGraphic($oPresentation[, $sPath = @ScriptDir[, $sSlide = Default[, $sFilter = "JPG"[, $bKeepRatio = True]]]])
; Parameters ....: $oPresentation - Presentation object.
;                  $sPath         - [optional] Directory where the graphics should be stored. The graphics are named Slide<n>. n is the slide number (default = keyword Default = @ScriptDir).
;                  $sSlide        - [optional] A string with the index number of the starting (and ending) slide to be exported (, separated by a hyphen) (default = Keyword Default = export all slides in the presentation).
;                  $sFilter       - [optional] The graphics format in which you want to export slides (default = JPG). See Remarks.
;                  $iScaleWidth   - [optional] The width in pixels of an exported slide (default = keyword Default = do not change the width).
;                  $iScaleHeight  - [optional] The height in pixels of an exported slide (default = keyword Default = do not change the heigth).
;                  $bKeepRatio    - [optional] If set To True the width:height ratio is preserved (default = True)
; Return values .: Success - 1.
;                  Failure - 0 and sets @error.
;                  |1 - $oPresentation is not an object or not a presentation object
;                  |2 - $sSlide is an object but not a sliderange object
;                  |3 - $sPath didn't exist and returned an error when the function tried to create it.
;                  |4 - Error occurred when exporting the presentation. @extended is set to the COM error code returned by the ExportAsFixedFormat method
;                  |5 - $sSlide is invalid. Has to be "StartingSlide-EndingSlide"
; Author ........: water
; Modified.......:
; Remarks .......: This function only supports a single range of consecutive slides.
;                  The specified graphics format must have an export filter registered in the Windows registry.
;                  You need to specify the registered extension (JPG, GIF etc.).
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_PresentationExportGraphic($oPresentation, $sPath = Default, $sSlide = Default, $sFilter = Default, $iScaleWidth = Default, $iScaleHeight = Default, $bKeepRatio = Default)
	Local $iStartingSlide, $iEndingSlide
	If Not IsObj($oPresentation) Or ObjName($oPresentation, 1) <> "_Presentation" Then Return SetError(1, 0, 0)
	If $sSlide = Default Then
		$iStartingSlide = 1
		$iEndingSlide = $oPresentation.Slides.Count
	Else
		Local $aSlides = StringSplit($sSlide, "-", $STR_NOCOUNT)
		Local $iValues = UBound($aSlides)
		If $iValues < 1 Or $iValues > 2 Then
			Return SetError(5, 0, 0)
		ElseIf $iValues = 1 Then
			ReDim $aSlides[2]
			$aSlides[1] = $aSlides[0]
		EndIf
		If Number($aSlides[0]) = 0 Or Number($aSlides[1]) = 0 Then Return SetError(5, 0, 0)
		$iStartingSlide = $aSlides[0]
		$iEndingSlide = $aSlides[1]
	EndIf
	If $sPath = Default Then $sPath = @ScriptDir
	If Not FileExists($sPath) Then
		If DirCreate($sPath) = 0 Then Return SetError(3, 0, 0)
	EndIf
	If $sFilter = Default Then $sFilter = "JPG"
	If $bKeepRatio = Default Then $bKeepRatio = True
	If $bKeepRatio = True And ($iScaleWidth <> Default Or $iScaleHeight <> Default) Then
		Local $iRatio = $oPresentation.Pagesetup.SlideWidth / $oPresentation.Pagesetup.SlideHeight
		If $iScaleWidth = Default Then
			$iScaleWidth = $oPresentation.Pagesetup.SlideWidth * ($iScaleHeight / $oPresentation.Pagesetup.SlideHeight)
		Else
			$iScaleHeight = $oPresentation.Pagesetup.SlideHeight * ($iScaleWidth / $oPresentation.Pagesetup.SlideWidth)
		EndIf
	EndIf
	For $i = $iStartingSlide To $iEndingSlide
		$oPresentation.Slides($i).Export($sPath & "\Slide" & $i & "." & $sFilter, $sFilter, $iScaleWidth, $iScaleHeight)
		If @error Then Return SetError(4, @error, 0)
	Next
	Return 1
EndFunc   ;==>_PPT_PresentationExportGraphic

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_PresentationList
; Description ...: Returns a list of currently open presentations
; Syntax.........: _PPT_PresentationList($oPPT)
; Parameters ....: $oPPT - PowerPoint application object to retrieve the list of presentations from
; Return values .: Success - a two-dimensional zero based array with the following information:
;                  |0 - Object of the presentation
;                  |1 - Name of the presentation/file
;                  |2 - Complete path to the presentation/file
;                  Failure - 0 and sets @error.
;                  |1 - $oPPT is not an object or not an application object
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......: None
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_PresentationList($oPPT)
	Local $aPresentations[1][3], $iIndex = 0
	If IsObj($oPPT) = 0 Or ObjName($oPPT, 1) <> "_Application" Then Return SetError(1, 0, 0)
	Local $iTemp = $oPPT.Presentations.Count
	ReDim $aPresentations[$iTemp][3]
	For $iIndex = 0 To $iTemp - 1
		$aPresentations[$iIndex][0] = $oPPT.Presentations($iIndex + 1)
		$aPresentations[$iIndex][1] = $oPPT.Presentations($iIndex + 1).Name
		$aPresentations[$iIndex][2] = $oPPT.Presentations($iIndex + 1).Path
	Next
	Return $aPresentations
EndFunc   ;==>_PPT_PresentationList

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_PresentationNew
; Description ...: Creates a new presentation
; Syntax.........: _PPT_PresentationNew($oPPT[, $sTemplate = "", [, $bVisible = True]])
; Parameters ....: $oPPT      - PowerPoint application object where you want to create the new presentation.
;                  $sTemplate - [optional] path to a template file (extension: pot, potx) or an existing presentation to be applied to the presentation (default = "" = no template).
;                  $bVisible  - [optional] True specifies that the presentation window will be visible (default = True).
; Return values .: Success - the presentation object.
;                  Failure - 0 and sets @error.
;                  |1 - $oPPT is not an object or not an application object
;                  |2 - Error returned by method Presentations.Add. @extended is set to the COM error code
;                  |3 - Specified template file does not exist
;                  |4 - Error returned by method ApplyTemplate. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......: _PPT_PresentationAttach, _PPT_PresentationClose, _PPT_PresentationOpen
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_PresentationNew($oPPT, $sTemplate = Default, $bVisible = Default)
	If Not IsObj($oPPT) Or ObjName($oPPT, 1) <> "_Application" Then Return SetError(1, 0, 0)
	If $sTemplate = Default Then $sTemplate = ""
	If $bVisible = Default Then $bVisible = True
	If $sTemplate <> "" And FileExists($sTemplate) = 0 Then Return SetError(3, 0, 0)
	Local $oPresentation = $oPPT.Presentations.Add($bVisible)
	If @error Then Return SetError(2, @error, 0)
	If $sTemplate <> "" Then
		$oPresentation.ApplyTemplate($sTemplate)
		If @error Then Return SetError(4, @error, 0)
	EndIf
	Return $oPresentation
EndFunc   ;==>_PPT_PresentationNew

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_PresentationOpen
; Description ...: Opens an existing presentation
; Syntax.........: _PPT_PresentationOpen($oPPT, $sFilePath[, $bReadOnly = False[, $bVisible = True]])
; Parameters ....: $oPPT      - PowerPoint application object where you want to open the presentation
;                  $sFilePath - Path and filename of the file to be opened.
;                  $bReadOnly - [optional] True opens the presentation as read-only (default = False).
;                  $bVisible  - [optional] True specifies that the presentation window will be visible (default = True).
; Return values .: Success - a presentation object. @extended is set to 1 if $bReadOnly = False but read-write access could not be granted. Please see the Remarks section for details.
;                  Failure - 0 and sets @error.
;                  |1 - $oPPT is not an object or not an application object
;                  |2 - Specified $sFilePath does not exist
;                  |3 - Unable to open $sFilePath. @extended is set to the COM error code returned by the Open method
; Author ........: water
; Modified.......:
; Remarks .......: When you set $bReadOnly = False but the presentation can't be opened read-write @extended is set to 1.
;                  The presentation was opened read-only because it has already been opened by another user/task or the file is set to read-only by the filesystem.
;                  If you modify the presentation you need to use _PPT_PresentationSaveAs() to save it to another location or with another name.
;                  +
;                  The PowerPoint object model does not allow to pass read/write passwords to open a protected presentation.
; Related .......: _PPT_PresentationAttach, _PPT_PresentationClose, _PPT_PresentationNew
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
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

#cs
	Set PR = ActivePresentation
	With PR.PrintOptions
	.FitToPage = msoCTrue
	.FrameSlides = msoCTrue
	.PrintColorType = ppPrintBlackAndWhite
	.PrintFontsAsGraphics = msoCTrue
	.OutputType = ppPrintOutputSlides
	.PrintHiddenSlides = msoCTrue
	.PrintInBackground = msoTrue
	End With
#ce
; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_PresentationPrint
; Description ...: Print one/multiple/all slides of a presentation
; Syntax.........: _PPT_PresentationPrint($oPresentation[, $iStartingSlide = Default[, $iEndingSlide = Default[, $iCopies = 1[, $sPrinter = Default]]]])
; Parameters ....: $oPresentation    - Presentation object.
;                  $iStartingSlide   - [optional] Number of the first slide to be printed (default = 1).
;                  $iEndingSlide     - [optional] Number of the last slide to be printed (default = last slide of the presentation).
;                  $iCopies          - [optional] Number of copies to be printer (default = 1).
;                  $sPrinter         - [optional] Sets the name of the printer (default = keyword Default = Active printer)
; Return values .: Success - 1.
;                  Failure - 0 and sets @error.
;                  |1 - $oPresentation is not an object or not a presentation object
;                  |2 - $sSlide is an object but not a sliderange object
;                  |3 - Error occurred when printing the presentation. @extended is set to the COM error code returned by the PrintOut method
;                  |4 - Error occurred when creating the WScript.Network object. @extended is set to the COM error code
;                  |5 - Error occurred when setting the default printer. @extended is set to the COM error code returned by the SetDefaultPrinter method
;                  |6 - Error occurred when re-setting the default printer. @extended is set to the COM error code returned by the SetDefaultPrinter method
;                  |7 - Error occurred when retrieving current default printer. @extended is set to the COM error code returned by the ActivePrinter property
;                  |8 - Error occurred when setting printer connection. @extended is set to the COM error code returned by the AddWindowsPrinterConnection method
; Author ........: water
; Modified.......:
; Remarks .......: Method PrintOut only supports a single range of consecutive slides.
;                  To print a single slide set parameters $iStartingSlide and $iEndingSlide to the same number.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_PresentationPrint($oPresentation, $iStartingSlide = Default, $iEndingSlide = Default, $iCopies = Default, $sPrinter = Default)
	Local $bPrinterChanged
	If Not IsObj($oPresentation) Or ObjName($oPresentation, 1) <> "_Presentation" Then Return SetError(1, 0, 0)
	If $iStartingSlide = Default Then $iStartingSlide = 1
	If $iEndingSlide = Default Then $iEndingSlide = $oPresentation.Slides.Count
	If $iCopies = Default Then $iCopies = 1
	If $sPrinter <> Default Then
		Local $sActivePrinter = $oPresentation.Parent.Parent.ActivePrinter
		If @error Then Return SetError(7, @error, 0)
		If $sActivePrinter <> $sPrinter Then
			Local $oWSHNetwork = ObjCreate("WScript.Network")
			If @error Then Return SetError(4, @error, 0)
			$oWSHNetwork.AddWindowsPrinterConnection($sPrinter)
			If @error Then Return SetError(8, @error, 0)
			$oWSHNetwork.SetDefaultPrinter($sPrinter)
			If @error Then Return SetError(5, @error, 0)
			$bPrinterChanged = True
		EndIf
	EndIf
	If Number($iStartingSlide) = 0 Or Number($iEndingSlide) = 0 Then Return SetError(2, 0, 0)
	$oPresentation.PrintOut($iStartingSlide, $iEndingSlide, Default, $iCopies)
	If @error Then Return SetError(3, @error, 0)
	If $bPrinterChanged Then
		$oWSHNetwork.SetDefaultPrinter($sActivePrinter)
		If @error Then Return SetError(6, @error, 0)
	EndIf
	Return 1
EndFunc   ;==>_PPT_PresentationPrint

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_PresentationSave
; Description ...: Saves a presentation
; Syntax.........: _PPT_PresentationSave($oPresentation)
; Parameters ....: $oPresentation - Object of the presentation to save.
; Return values .: Success - 1 and sets @extended.
;                  |0 - File has not been saved because it has not been changed since the last save or file open
;                  |1 - File has been saved because it has been changed since the last save or file open
;                  Failure - 0 and sets @error.
;                  |1 - $oPresentation is not an object or not a presentation object
;                  |2 - Error occurred when saving the presentation. @extended is set to the COM error code
; Author ........: water
; Modified.......:
; Remarks .......: A newly created presentation has to be saved using _PPT_PresentationSaveAs before
; Related .......: _PPT_PresentationSaveAs
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_PresentationSave($oPresentation)
	If Not IsObj($oPresentation) Or ObjName($oPresentation, 1) <> "_Presentation" Then Return SetError(1, 0, 0)
	If Not $oPresentation.Saved Then
		$oPresentation.Save()
		If @error Then Return SetError(2, @error, 0)
		Return SetError(0, 1, 1)
	EndIf
	Return 1
EndFunc   ;==>_PPT_PresentationSave

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_PresentationSaveAs
; Description ...: Saves a presentation with a new filename and/or type
; Syntax.........: _PPT_PresentationSaveAs($oPresentation, $sFilePath[, $iFormat = $ppSaveAsDefault[, $bOverWrite = False]])
; Parameters ....: $oPresentation - Presentation object to be saved.
;                  $sFilename     - Filename of the file to be written to (no extension).
;                  $iFormat       - [optional] PowerPoint writeable filetype. Can be any value of the PpSaveAsFileType enumeration.
;                                   (default = keyword Default = $ppSaveAsDefault (means: .ppt for PowerPoint < 2007, .pptx for PowerPoint 2007 and later).
;                  $bOverWrite    - [optional] True overwrites an already existing file (default = False).
; Return values .: Success - 1.
;                  Failure - 0 and sets @error.
;                  |1 - $oPresentation is not an object or not a presentation object
;                  |2 - $iFormat is not a number
;                  |3 - File exists, overwrite flag not set
;                  |4 - File exists but could not be deleted
;                  |5 - Error occurred when saving the presentation. @extended is set to the COM error code returned by the SaveAs method
; Author ........: water
; Modified.......:
; Remarks .......: $sFilename is the filename without extension. "test" is correct, "test.pps" is wrong.
; Related .......: _PPT_PresentationSave
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_PresentationSaveAs($oPresentation, $sFilePath, $iFormat = Default, $bOverWrite = Default)
	If Not IsObj($oPresentation) Or ObjName($oPresentation, 1) <> "_Presentation" Then Return SetError(1, 0, 0)
	If $iFormat = Default Then
		$iFormat = $ppSaveAsDefault
	Else
		If Not IsNumber($iFormat) Then Return SetError(2, 0, 0)
	EndIf
	If $bOverWrite = Default Then $bOverWrite = False
	If FileExists($sFilePath) Then
		If Not $bOverWrite Then Return SetError(3, 0, 0)
		Local $iResult = FileDelete($sFilePath)
		If $iResult = 0 Then Return SetError(4, 0, 0)
	EndIf
	$oPresentation.SaveAs($sFilePath, $iFormat)
	If @error Then Return SetError(5, @error, 0)
	Return 1
EndFunc   ;==>_PPT_PresentationSaveAs

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_SlideAdd
; Description ...: Add one or multiple slides to a presentation
; Syntax.........: _PPT_SlideAdd($oPresentation[, $vIndex = Default[, $iSlides = 1[, $sLayout = Default]]])
; Parameters ....: $oPresentation - Presentation object.
;                  $vIndex        - [optional] The name or index where the slide is to be added (default = keyword Default = add the slide at the end of the presentation).
;                  $iSlides       - [optional] Number of slides to be added (default = 1).
;                  $vLayout       - [optional] The layout of the slide (default = keyword Default = Layout of the preceding/next slide).
;                                   Can be any of the PpSlideLayout enumeration or a layout object.
; Return values .: Success - the slide object of the first slide added.
;                  Failure - 0 and sets @error.
;                  |1 - $oPresentation is not an object or not a presentation object
;                  |2 - $vIndex is a number and < 1 or > current number of slides + 1
;                  |3 - Error occurred when retrieving layout of a slide. @extended is set to the COM error code
;                  |4 - Error occurred when adding a slide. @extended is set to the COM error code returned by the Add/AddSlide method
;                  |5 - $iSlides is not a number or < 1
; Author ........: water
; Modified.......:
; Remarks .......: If $vLayout ist set to Default then the layout of the slide with index $vIndex is used.
;                  If there is no such slide then the layout of the preceding slide is being used.
;                  If there is no such slide then layout $ppLayoutBlank is being used.
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_SlideAdd($oPresentation, $vIndex = Default, $iSlides = Default, $vLayout = Default)
	Local $oSlide, $oSlideReturn, $iSlideCount, $oSlideLayout, $iDiff = 1
	If Not IsObj($oPresentation) Or ObjName($oPresentation, 1) <> "_Presentation" Then Return SetError(1, 0, 0)
	If $iSlides = Default Then $iSlides = 1
	If Not IsNumber($iSlides) Or $iSlides < 1 Then Return SetError(5, 0, 0)
	$iSlideCount = $oPresentation.Slides.Count()
	If $iSlideCount = 0 Then $iDiff = 2
	If $vIndex = Default Then
		$vIndex = $iSlideCount + 1
	Else
		If IsNumber($vIndex) And ($vIndex < 1 Or $vIndex > $iSlideCount + 1) Then Return SetError(2, 0, 0)
	EndIf
	$oSlideLayout = $vLayout
	For $i = 1 To $iSlides
		If $vLayout = Default Then
			$iSlideCount = $oPresentation.Slides.Count()
			Switch $iSlideCount
				Case 0
					; $vLayout = $ppLayoutBlank ; Default layout
					$oSlideLayout = $oPresentation.SlideMaster.CustomLayouts.Item(1) ; Titlelayout from SlideMaster
				Case 1
					$oSlideLayout = $oPresentation.SlideMaster.CustomLayouts.Item(2) ; Layout of first content slide
				Case Else
					$oSlideLayout = $oPresentation.Slides($vIndex - $iDiff + $i).CustomLayout ; Get the Layout of the slide where to insert the new slide
			EndSwitch
			If @error Then Return SetError(4, @error, 0)
		EndIf
		If IsObj($oSlideLayout) Then
			$oSlide = $oPresentation.Slides.AddSlide($vIndex - 1 + $i, $oSlideLayout)
		Else
			$oSlide = $oPresentation.Slides.Add($vIndex - 1 + $i, $oSlideLayout)
		EndIf
		If @error Then Return SetError(4, @error, 0)
		If $i = 1 Then $oSlideReturn = $oSlide
	Next
	Return $oSlideReturn
EndFunc   ;==>_PPT_SlideAdd

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_SlideCopyMove
; Description ...: Copies or moves the specified slide(s) before or after a specified slide in the same or a different presentation. Duplicates slides in the same presentation
; Syntax.........: _PPT_SlideCopyMove($oSourcePresentation, $vSourceSlide[, $oTargetPresentation = $oSourcePresentation[, $vTargetSlide = Default[, $iFunction = 1]]])
; Parameters ....: $oSourcePresentation - Presentation object of the source presentation.
;                  $vSourceSlide        - A SlideRange object or a string with index numbers or names of the slides to be processes, separated by comma.
;                  $oTargetPresentation - [optional] Presentation object of the target presentation where slides should be copied/moved to (default = keyword Default = $oSourcePresentation).
;                  $vTargetSlide        - [optional] Index or name of the slide that the slides on the Clipboard are to be pasted before (default = keyword Default = copy after the last slide in the presentation).
;                  $iFunction           - [optional] Specifies how to process the slides:
;                  |1: Copy the specified slide(s) to the target presentation (default)
;                  |2: Move the spcified slide(s) to the target presentation
;                  |3: Duplicate the specified slide(s) in the source presentation
;                  |4: Copy the specified slide(s) from the source presentation to the clipboard
;                  |5: Cut the specified slide(s) from the source presentation to the clipboard
; Return values .: Success - a sliderange object for $iFunction = 1, 2 to another presentation and 3.
;                            1 for $iFunction = 2 when moving within the same presentation, 4 and 5.
;                  Failure - 0 and sets @error.
;                  |1 - $oSourcePresentation is not an object or not a presentation object
;                  |2 - $oTargetPresentation is not an object or not a presentation object
;                  |3 - Error occurred when copying the slide(s). @extended is set to the COM error code returned by the Copy method
;                  |4 - Error occurred when cutting the slide(s). @extended is set to the COM error code returned by the Cut method
;                  |5 - Error occurred when moving the slide(s). @extended is set to the COM error code returned by the MoveTo method
;                  |6 - Error occurred when inserting slide(s). @extended is set to the COM error code returned by the Paste method
;                  |7 - Error occurred when duplicating slide(s). @extended is set to the COM error code returned by the Duplicate method
;                  |8 - $iFunction < 1 or $iFunction > 5
;                  |9 - You can't duplicate slides to another presentation. $oTargetPresentation = $oSourcePresentation is needed
;                  |10 - $vTargetSlide is a number and < 1 or > current number of slides + 1
;                  |11 - $vSourceSlide is an object but not a sliderange object
;                  |12 - Error occurred creating the SlideRange for the source presentation. @extended is set to the COM error code returned by the Duplicate method
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_SlideCopyMove($oSourcePresentation, $vSourceSlide, $oTargetPresentation = Default, $vTargetSlide = Default, $iFunction = Default)

	If Not IsObj($oSourcePresentation) Or ObjName($oSourcePresentation, 1) <> "_Presentation" Then Return SetError(1, 0, 0)
	If $oTargetPresentation = Default Then $oTargetPresentation = $oSourcePresentation
	If Not IsObj($oTargetPresentation) Or ObjName($oTargetPresentation, 1) <> "_Presentation" Then Return SetError(2, 0, 0)
	Local $iSlideCount = $oTargetPresentation.Slides.Count
	If $vTargetSlide = Default Then $vTargetSlide = $iSlideCount + 1
	If IsNumber($vTargetSlide) And ($vTargetSlide < 1 Or $vTargetSlide > $iSlideCount + 1) Then Return SetError(10, 0, 0)
	If $iFunction = Default Then $iFunction = 1
	If $iFunction < 1 Or $iFunction > 5 Then Return SetError(8, 0, 0)
	If $iFunction = 3 And $oSourcePresentation <> $oTargetPresentation Then Return SetError(9, 0, 0)
	If Not IsObj($vSourceSlide) Then
		Local $aSlides = StringSplit($vSourceSlide, ",", $STR_NOCOUNT)
		For $i = 0 To UBound($aSlides) - 1
			If Number($aSlides[$i]) > 0 Then $aSlides[$i] = Number($aSlides[$i]) ; Translate numeric string to a number
		Next
		$vSourceSlide = $oSourcePresentation.Slides.Range($aSlides)
		If @error Then Return SetError(12, @error, 0)
	Else
		If ObjName($vSourceSlide, 1) <> "SlideRange" Then Return SetError(11, 0, 0)
	EndIf
	Switch $iFunction
		Case 1 ; Copy to target presentation
			$vSourceSlide.Copy()
			If @error Then Return SetError(3, @error, 0)
			$oSlideRange = $oTargetPresentation.Slides.Paste($vTargetSlide)
			If @error Then Return SetError(6, @error, 0)
			Return $oSlideRange
		Case 2 ; Move to target presentation
			If $oSourcePresentation = $oTargetPresentation Then ; Move within the same presentation
				$vSourceSlide.MoveTo($vTargetSlide)
				If @error Then Return SetError(5, @error, 0)
				Return 1
			Else
				$vSourceSlide.Cut() ; Cut/paste to another presentation
				If @error Then Return SetError(4, @error, 0)
				$oSlideRange = $oTargetPresentation.Slides.Paste($vTargetSlide)
				If @error Then Return SetError(6, @error, 0)
				Return $oSlideRange
			EndIf
		Case 3 ; Duplicate
			$oSlideRange = $vSourceSlide.Duplicate()
			If @error Then Return SetError(7, @error, 0)
			Return $oSlideRange
		Case 4 ; Copy to the clipboard
			$vSourceSlide.Copy()
			If @error Then Return SetError(3, @error, 0)
			Return 1
		Case 5 ; Move to the clipboard
			$vSourceSlide.Cut() ; Cut/paste to another presentation
			If @error Then Return SetError(4, @error, 0)
			Return 1
	EndSwitch
EndFunc   ;==>_PPT_SlideCopyMove

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_SlideDelete
; Description ...: Delete one or multiple slides from a presentation
; Syntax.........: _PPT_SlideDelete($oPresentation, $vSlide)
; Parameters ....: $oPresentation - Presentation object.
;                  $vSlide        - A SlideRange object or a string with index numbers or names of the slides to be deleted, separated by comma.
; Return values .: Success - 1.
;                  Failure - 0 and sets @error.
;                  |1 - $oPresentation is not an object or not a presentation object
;                  |2 - $vSlide is an object but not a sliderange object
;                  |3 - Error deleting slides. @extended is set to the COM error code returned by the Delete method
;                  |4 - Error occurred creating the SlideRange from $vSlide. @extended is set to the COM error code returned by the Duplicate method
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_SlideDelete($oPresentation, $vSlide)
	If Not IsObj($oPresentation) Or ObjName($oPresentation, 1) <> "_Presentation" Then Return SetError(1, 0, 0)
	If Not IsObj($vSlide) Then
		Local $aSlides = StringSplit($vSlide, ",", $STR_NOCOUNT)
		For $i = 0 To UBound($aSlides) - 1
			If Number($aSlides[$i]) > 0 Then $aSlides[$i] = Number($aSlides[$i]) ; Translate numeric string to a number
		Next
		$vSlide = $oPresentation.Slides.Range($aSlides)
		If @error Then Return SetError(4, @error, 0)
	Else
		If ObjName($vSlide, 1) <> "SlideRange" Then Return SetError(2, 0, 0)
	EndIf
	$vSlide.Delete()
	If @error Then Return SetError(3, @error, 0)
	Return 1
EndFunc   ;==>_PPT_SlideDelete

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_SlideShow
; Description ...: Set properties for a slide show and run the show
; Syntax.........: _PPT_SlideShow($oPresentation[, $bRun = True[, $vStartingSlide = 1[, $vEndingSlide = Default[, $bLoop = True[, $iShowType = $ppShowTypeKiosk]]]]])
; Parameters ....: $oPresentation  - Presentation object.
;                  $bRun           - [optional] If True then the function starts the slide show (default = True).
;                  $vStartingSlide - [optional] Name or index of the first slide to be shown (default = 1).
;                  $vEndingSlide   - [optional] Name or index of the last slide to be shown (default = keyword Default = the last slide in the presentation).
;                  $bLoop          - [optional] If True then the slide show starts again when having reached the end (default = True).
;                  $iShowType      - [optional] Type of slide show defined by the PpSlideShowType enumeration (default = $ppShowTypeKiosk).
; Return values .: Success - the slidewindow object when $bRun = True, else 1.
;                  Failure - 0 and sets @error.
;                  |1 - $oPresentation is not an object or not a presentation object
;                  |2 - $vStartingSlide is a number and < 1 or > current number of slides
;                  |3 - $vEndingSlide is a number and < 1 or > current number of slides
;                  |4 - Error occurred when running the slide show. @extended is set to the COM error code returned by the Run method
; Author ........: water
; Modified.......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_SlideShow(ByRef $oPresentation, $bRun = Default, $vStartingSlide = Default, $vEndingSlide = Default, $bLoop = Default, $iShowType = Default)

	If Not IsObj($oPresentation) Or ObjName($oPresentation, 1) <> "_Presentation" Then Return SetError(1, 0, 0)
	Local $iSlideCount = $oPresentation.Slides.Count
	If $bRun = Default Then $bRun = True
	If $vStartingSlide = Default Then $vStartingSlide = 1
	If IsNumber($vStartingSlide) And ($vStartingSlide < 1 Or $vStartingSlide > $iSlideCount) Then Return SetError(2, 0, 0)
	If $vEndingSlide = Default Then $vEndingSlide = $iSlideCount
	If IsNumber($vEndingSlide) And ($vEndingSlide < 1 Or $vEndingSlide > $iSlideCount) Then Return SetError(3, 0, 0)
	If $bLoop = Default Then $bLoop = True
	If $iShowType = Default Then $iShowType = $ppShowTypeKiosk
	If $vStartingSlide <> Default Or $vEndingSlide <> Default Then ; User specified start or end slide => Set RangeType
		$oPresentation.SlideshowSettings.StartingSlide = $vStartingSlide
		$oPresentation.SlideshowSettings.EndingSlide = $vEndingSlide
		$oPresentation.SlideshowSettings.RangeType = $ppShowSlideRange
	Else
		$oPresentation.SlideshowSettings.RangeType = $ppShowAll ; Show all slides
	EndIf
	$oPresentation.SlideshowSettings.LoopUntilStopped = $bLoop
	$oPresentation.SlideshowSettings.ShowType = $iShowType
	If $bRun Then
		Local $oSlideShow = $oPresentation.SlideshowSettings.Run()
		If @error Then Return SetError(4, @error, 0)
		Return $oSlideShow
	EndIf
	Return 1
EndFunc   ;==>_PPT_SlideShow

; #FUNCTION# ====================================================================================================================
; Name...........: _PPT_VersionInfo
; Description ...: Returns an array of information about the PowerPoint UDF
; Syntax.........: _PPT_VersionInfo()
; Parameters ....: None
; Return values .: Success - one-dimensional one based array with the following information:
;                  |1 - Release Type (T=Test or V=Production)
;                  |2 - Major Version
;                  |3 - Minor Version
;                  |4 - Sub Version
;                  |5 - Release Date (YYYYMMDD)
;                  |6 - AutoIt version required
;                  |7 - List of authors separated by ","
;                  |8 - List of contributors separated by ","
; Author ........: water
; Modified.......:
; Remarks .......: Based on function _IE_VersionInfo written bei Dale Hohm
; Related .......:
; Link ..........:
; Example .......: Yes
; ===============================================================================================================================
Func _PPT_VersionInfo()

	Local $aVersionInfo[9] = [8, "T", 0, 0, 5.0, "20170606", "3.3.10.2", "water", ""]
	Return $aVersionInfo

EndFunc   ;==>_PPT_VersionInfo

; #INTERNAL_USE_ONLY#============================================================================================================
; Name ..........: __PPT_ErrorHandler
; Description ...: Called if an ObjEvent error occurs
; Syntax.........: __PPT_ErrorHandler()
; Parameters ....: None
; Return values .: @error is set to the COM error by AutoIt
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __PPT_ErrorHandler()

	Local $bHexNumber = Hex($__oPPT_Error.number, 8)
	Local $aVersionInfo = _PPT_VersionInfo()
	Local $sError = "COM Error Encountered in " & @ScriptName & @CRLF & _
			"PowerPoint UDF version = " & $aVersionInfo[2] & "." & $aVersionInfo[3] & "." & $aVersionInfo[4] & @CRLF & _
			"@AutoItVersion = " & @AutoItVersion & @CRLF & _
			"@AutoItX64 = " & @AutoItX64 & @CRLF & _
			"@Compiled = " & @Compiled & @CRLF & _
			"@OSArch = " & @OSArch & @CRLF & _
			"@OSVersion = " & @OSVersion & @CRLF & _
			"Scriptline = " & $__oPPT_Error.scriptline & @CRLF & _
			"NumberHex = " & $bHexNumber & @CRLF & _
			"Number = " & $__oPPT_Error.number & @CRLF & _
			"WinDescription = " & StringStripWS($__oPPT_Error.WinDescription, 2) & @CRLF & _
			"Description = " & StringStripWS($__oPPT_Error.Description, 2) & @CRLF & _
			"Source = " & $__oPPT_Error.Source & @CRLF & _
			"HelpFile = " & $__oPPT_Error.HelpFile & @CRLF & _
			"HelpContext = " & $__oPPT_Error.HelpContext & @CRLF & _
			"LastDllError = " & $__oPPT_Error.LastDllError
	If $__iPPT_Debug > 0 Then
		If $__iPPT_Debug = 1 Then ConsoleWrite($sError & @CRLF & "========================================================" & @CRLF)
		If $__iPPT_Debug = 2 Then MsgBox(64, "PowerPoint UDF - Debug Info", $sError)
		If $__iPPT_Debug = 3 Then FileWrite($__sPPT_DebugFile, @YEAR & "." & @MON & "." & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & " " & @CRLF & _
				"-------------------" & @CRLF & $sError & @CRLF & "========================================================" & @CRLF)
	EndIf

EndFunc   ;==>__PPT_ErrorHandler

; Soll Range analog zu _ArrayDisplay akzeptieren. "1,2" oder "1-2" oder "1:2"
; #INTERNAL_USE_ONLY#============================================================================================================
; Name...........: __PPT_SliderangeCreate
; Description ...: Creates a Sliderange object from a list of slides (index, name)
; Syntax.........: __PPT_SliderangeCreate($vSlide)
; Parameters ....: $oPresentation - Presentation object.
;                  $vSlide        - A SlideRange object or a string with index numbers or names of the slides to be deleted, separated by comma.
; Return values .: Success - Sliderange object.
;                  Failure - 0 and sets @error.
;                  |1 - $oPresentation is not an object or not a presentation object
;                  |2 - $vSlide is an object but not a sliderange object
;                  |3 - Error occurred creating the SlideRange from $vSlide. @extended is set to the COM error code returned by the Duplicate method
; Author ........: water
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __PPT_SliderangeCreate($oPresentation, $vSlide)
	If Not IsObj($oPresentation) Or ObjName($oPresentation, 1) <> "_Presentation" Then Return SetError(1, 0, 0)
	If Not IsObj($vSlide) Then
		Local $aSlides = StringSplit($vSlide, ",", $STR_NOCOUNT)
		For $i = 0 To UBound($aSlides) - 1
			If Number($aSlides[$i]) > 0 Then $aSlides[$i] = Number($aSlides[$i]) ; Translate numeric string to a number
		Next
		$vSlide = $oPresentation.Slides.Range($aSlides)
		If @error Then Return SetError(3, @error, 0)
	Else
		If ObjName($vSlide, 1) <> "SlideRange" Then Return SetError(1, 0, 0)
	EndIf
	Return $vSlide
EndFunc   ;==>__PPT_SliderangeCreate
