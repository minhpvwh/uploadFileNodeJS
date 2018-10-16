#include-once

; #INDEX# =======================================================================================================================
; Title .........: PowerPointConstants
; AutoIt Version : 3.3.12.0
; Language ......: English
; Description ...: Constants to be included in an AutoIt script when using the PowerPoint UDF.
; Author(s) .....: water
; Resources .....: PowerPoint 2010 Enumerations: https://msdn.microsoft.com/en-us/library/ff744042%28v=office.14%29.aspx
; ===============================================================================================================================

; #CONSTANTS# ===================================================================================================================
; PpFixedFormatType Enumeration. Specify the type of fixed-format file to export.
; See: https://msdn.microsoft.com/en-us/library/ff746754%28v=office.14%29.aspx
Global Const $ppFixedFormatTypePDF = 2 ; PDF format
Global Const $ppFixedFormatTypeXPS = 1 ; XPS format

; PpPrintOutputType Enumeration. Indicates which component (slides, handouts, notes pages, or an outline) of the presentation is to be printed.
; See: https://msdn.microsoft.com/en-us/library/ff744185%28v=office.14%29.aspx
Global Const $ppPrintOutputBuildSlides = 7 ; Build Slides
Global Const $ppPrintOutputFourSlideHandouts = 8 ; Four Slide Handouts
Global Const $ppPrintOutputNineSlideHandouts = 9 ; Nine Slide Handouts
Global Const $ppPrintOutputNotesPages = 5 ; Notes Pages
Global Const $ppPrintOutputOneSlideHandouts = 10 ; Single Slide Handouts
Global Const $ppPrintOutputOutline  = 6 ; Outline
Global Const $ppPrintOutputSixSlideHandouts = 4 ; Six Slide Handouts
Global Const $ppPrintOutputSlides = 1 ; Slides
Global Const $ppPrintOutputThreeSlideHandouts  = 3 ; Three Slide Handouts
Global Const $ppPrintOutputTwoSlideHandouts = 2 ; Two Slide Handouts

; PpPrintRangeType Enumeration. Specifies the type of print range for the presentation.
; See: https://msdn.microsoft.com/en-us/library/ff745585%28v=office.14%29.aspx
Global Const $ppPrintAll = 1 ; Print all slides in the presentation
Global Const $ppPrintCurrent = 3 ; Print the current slide from the presentation
Global Const $ppPrintNamedSlideShow = 5 ; Print a named slideshow
Global Const $ppPrintSelection = 2 ; Print a selection of slides
Global Const $ppPrintSlideRange = 4 ; Print a range of slides

; PpSaveAsFileType Enumeration. Specify type of file to save as.
; See: https://msdn.microsoft.com/en-us/library/ff746500%28v=office.14%29.aspx
Global Const $ppSaveAsAddIn = 8 ; Save as an AddIn
Global Const $ppSaveAsBMP = 19 ; Save as an BMP image
Global Const $ppSaveAsDefault = 11 ; Save in the default format
Global Const $ppSaveAsEMF = 23 ; Save in the Enhanced MetaFile (EMF) format
Global Const $ppSaveAsGIF = 16 ; Save as a GIF image
Global Const $ppSaveAsHTML = 12 ; Save as an HTML document
Global Const $ppSaveAsHTMLDual = 14 ; Save as HTML Dual version
Global Const $ppSaveAsHTMLv3 = 13 ; Save as HTMLv3
Global Const $ppSaveAsJPG = 17 ; Save as a JPG image
Global Const $ppSaveAsMetaFile = 15 ; Save as as a MetaFile
Global Const $ppSaveAsOpenXMLAddin = 30 ; Save as an open XML add-in
Global Const $ppSaveAsOpenXMLPresentation = 24 ; Save as an open XML presentation
Global Const $ppSaveAsOpenXMLPresentationMacroEnabled = 25 ; Save as a macro-enabled open XML presentation
Global Const $ppSaveAsOpenXMLShow = 28 ; Save as an open XML show
Global Const $ppSaveAsOpenXMLShowMacroEnabled = 29 ; Save as a macro-enabled open XML show
Global Const $ppSaveAsOpenXMLTemplate = 26 ; Save as an open XML template
Global Const $ppSaveAsOpenXMLTemplateMacroEnabled = 27 ; Save as a macro-enabled open XML template
Global Const $ppSaveAsOpenXMLTheme = 31 ; Save as an open XML theme
Global Const $ppSaveAsPDF = 32 ; Save as a PDF
Global Const $ppSaveAsPNG = 18 ; Save as a PNG image
Global Const $ppSaveAsPresentation = 1 ; Save as a presentation
Global Const $ppSaveAsRTF = 6 ; Save as an RTF
Global Const $ppSaveAsShow = 7 ; Save as a slideshow
Global Const $ppSaveAsTemplate = 5 ; Save as a template
Global Const $ppSaveAsTIF = 21 ; Save as a TIF file
Global Const $ppSaveAsWebArchive = 20 ; Save as a Web archive
Global Const $ppSaveAsXPS = 33 ; Save in the XML Paper Specification (XPS) format

; PpSlideLayout Enumeration. Specify the layout of the slide.
; See: https://msdn.microsoft.com/en-us/library/ff746831%28v=office.14%29.aspx
Global Const $ppLayoutBlank = 12 ; Blank
Global Const $ppLayoutChart = 8 ; Chart
Global Const $ppLayoutChartAndText = 6 ; Chart and text
Global Const $ppLayoutClipartAndText = 10 ; Clipart and text
Global Const $ppLayoutClipArtAndVerticalText = 26 ; ClipArt and vertical text
Global Const $ppLayoutCustom = 32 ; Custom
Global Const $ppLayoutFourObjects = 24 ; Four objects
Global Const $ppLayoutLargeObject = 15 ; Large object
Global Const $ppLayoutMediaClipAndText = 18 ; MediaClip and text
Global Const $ppLayoutMixed = -2 ; Mixed
Global Const $ppLayoutObject = 16 ; Object
Global Const $ppLayoutObjectAndText = 14 ; Object and text
Global Const $ppLayoutObjectAndTwoObjects = 30 ; Object and two objects
Global Const $ppLayoutObjectOverText = 19 ; Object over text
Global Const $ppLayoutOrgchart =  7 ; Organization chart
Global Const $ppLayoutTable = 4 ; Table
Global Const $ppLayoutText = 2 ; Text
Global Const $ppLayoutTextAndChart = 5 ; Text and chart
Global Const $ppLayoutTextAndClipart = 9 ; Text and clipart
Global Const $ppLayoutTextAndMediaClip = 17 ; Text and MediaClip
Global Const $ppLayoutTextAndObject = 13 ; Text and object
Global Const $ppLayoutTextAndTwoObjects = 21 ; Text and two objects
Global Const $ppLayoutTextOverObject = 20 ; Text over object
Global Const $ppLayoutTitle = 1 ; Title
Global Const $ppLayoutTitleOnly = 11 ; Title only
Global Const $ppLayoutTwoColumnText = 3 ; Two-column text
Global Const $ppLayoutTwoObjects = 29 ; Two objects
Global Const $ppLayoutTwoObjectsAndObject = 31 ; Two objects and object
Global Const $ppLayoutTwoObjectsAndText = 22 ; Two objects and text
Global Const $ppLayoutTwoObjectsOverText = 23 ; Two objects over text
Global Const $ppLayoutVerticalText = 25 ; Vertical text
Global Const $ppLayoutVerticalTitleAndText = 27 ; Vertical title and text
Global Const $ppLayoutVerticalTitleAndTextOverChart = 28 ; Vertical title and text over chart

; PpSlideShowRangeType Enumeration. Specify the type of the slideshow range.
; See: https://msdn.microsoft.com/en-us/library/ff744228%28v=office.14%29.aspx
Global Const $ppShowAll = 1 ; Show all
Global Const $ppShowNamedSlideShow = 3 ; Show named slideshow
Global Const $ppShowSlideRange = 2 ; Show slide range

; PpSlideShowType Enumeration. Specify the type of slide show.
; See: https://msdn.microsoft.com/en-us/library/ff745785%28v=office.14%29.aspx
Global Const $ppShowTypeKiosk = 3 ; Kiosk
Global Const $ppShowTypeSpeaker = 1 ; Speaker
Global Const $ppShowTypeWindow = 2 ; Window
; ===============================================================================================================================
