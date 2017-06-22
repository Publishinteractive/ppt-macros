Attribute VB_Name = "ppttodoc"
'---------------------------------------------------------------------------------
' The sample scripts are not supported under any Microsoft standard support
' program or service. The sample scripts are provided AS IS without warranty
' of any kind. Microsoft further disclaims all implied warranties including,
' without limitation, any implied warranties of merchantability or of fitness for
' a particular purpose. The entire risk arising out of the use or performance of
' the sample scripts and documentation remains with you. In no event shall
' Microsoft, its authors, or anyone else involved in the creation, production, or
' delivery of the scripts be liable for any damages whatsoever (including,
' without limitation, damages for loss of business profits, business interruption,
' loss of business information, or other pecuniary loss) arising out of the use
' of or inability to use the sample scripts or documentation, even if Microsoft
' has been advised of the possibility of such damages.
'---------------------------------------------------------------------------------
'
' Modified By Andrew Woods 01/06/2017
' Needs documentation!
'
'



Option Explicit

' ################################################
' The starting point of execution for this window.
' ################################################
Public Sub PPTtoDOC()
Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
StartTime = Timer

Call SendPowerPoint2Word(ActivePresentation.FullName, False)

'Determine how many minutes code took to run
SecondsElapsed = Round(Timer - StartTime, 2)
'MinutesElapsed = Format((Timer - StartTime) / 86400, "hh:mm:ss")
MsgBox "Conversion to Word complete. " & SecondsElapsed & " seconds elapsed", vbOKOnly, "Message"

End Sub

' #####################################
' Send PowerPoint presentation to Word.
' #####################################
Public Sub SendPowerPoint2Word(ByVal FullPath As String, ByVal IsOther As Boolean)
    Dim wdApp           As Object       ' Word application.
    Dim wdDoc           As Object       ' Word document.
    Dim sldEach         As Object       ' Each slide.
    Dim sldAll          As Object       ' All slides.
    Dim spNotes         As Object       ' Notes shapes.
    Dim spNotesPage     As Object       ' All shapes in notes page.
    Dim pptPresentation As Object       ' PowerPoint presentation.
    Dim sldHeight       As Single       ' Slide height.
    Dim sldWidth        As Single       ' Slide width.
    Dim strFilePath     As String       ' File path.
    Dim strFileName     As String       ' File name with no extension.
    Dim strNotesText    As String       ' Notes text.
    Dim intPageNumber   As Integer      ' Page number in Word.
    Dim strProgress     As String       ' The conversion progress.
    Dim shp             As Shape
    
    ' /* Constants declaration. */
    Const wdPaperCustom = 41
    Const wdStory = 6
    Const wdCharacter = 1
    Const wdExtend = 1
    Const wdGoToPage = 1
    Const wdGoToNext = 2
    
    On Error Resume Next
    
    ' To get the current progress.
    Dim iNow As Integer
    ' To count all slides.
    Dim iAll As Integer
    
    iNow = 0
    
    Set pptPresentation = GetObject(FullPath)
    
    ' Get the file path from the given path.
    strFilePath = CGPath(pptPresentation.Path)
    ' Get the file name from the given path.
    strFileName = GetFileNameFromFullPath(pptPresentation.FullName)
    
    Set wdApp = CreateObject("Word.Application")
    Set wdDoc = wdApp.Documents.Add
    
    'Set reference to all slides in PowerPoint presentation.
    Set sldAll = pptPresentation.Slides
    
    ' Get the number of all slides.
    iAll = sldAll.Count
    
    With wdApp.Selection
        ' /* Go through each slide object. */
        For Each sldEach In sldAll
            
            iNow = iNow + 1
            
            ' Display the progress.
            'ufrmP2W.Caption = "Converting slide " & CStr(iNow) & " of " & CStr(iAll)
            
            Set spNotesPage = sldEach.NotesPage.Shapes
            
            ' /* Read notes in the current slide. */
            For Each spNotes In spNotesPage
                If spNotes.HasTextFrame Then
                    If spNotes.PlaceholderFormat.Type = ppPlaceholderBody Then
                        strNotesText = spNotes.TextFrame.TextRange.Text
                        Exit For
                    End If
                End If
            Next
                      
            For Each shp In sldEach.Shapes
              If shp.TextFrame.HasText Then
              ' Make white text grey so it doesn't visually dissapear
                If shp.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255) Then
                  shp.TextFrame.TextRange.Font.Color.RGB = RGB(150, 150, 150)
                End If
                shp.TextFrame.TextRange.Copy
                .Paste
              ' Make the grey text white again.
                If shp.TextFrame.TextRange.Font.Color.RGB = RGB(150, 150, 150) Then
                  shp.TextFrame.TextRange.Font.Color.RGB = RGB(255, 255, 255)
                End If
              Else
                shp.Copy
                '.Paste
                .PasteAndFormat Type:=wdFormatOriginalFormatting
                .TypeParagraph
              End If
            Next
            
            ' To count the page number.
            intPageNumber = intPageNumber + 1
            
            .ShapeRange.Group
            .ShapeRange.Left = 0
            .ShapeRange.Ungroup
            
            ' Put the notes in as a paragraph
            If strNotesText <> "" Then
              .Text = strNotesText
            End If
            
            .EndKey wdStory
            .InsertNewPage
        Next
        
        ' /* To delete the last blank page in Word. */
        .TypeBackspace
        .TypeBackspace
        
        ' /* Copy the newline char to reduce a large data on clipboard. */
        .MoveRight wdCharacter, 1, wdExtend
        .Copy
    End With
    wdDoc.Content.Find.Execute FindTExt:="^l", ReplaceWith:="^p", Replace:=wdReplaceAll

    wdDoc.SaveAs strFilePath & strFileName
    wdDoc.Close
    wdApp.Quit
       
    ' /* Release memory. */
    Set wdApp = Nothing
    Set wdDoc = Nothing
    Set sldEach = Nothing
    Set sldAll = Nothing
    Set spNotes = Nothing
    Set spNotesPage = Nothing
    Set pptPresentation = Nothing
End Sub

' #########################################
' Get file name from a specified full path.
' #########################################
Private Function GetFileNameFromFullPath(ByVal FullPath As String) As String
    Dim lngPathSeparatorPosition    As Long     ' Path separator.
    Dim lngDotPosition              As Long     ' Dot position.
    Dim strFile                     As String   ' A full file name.
    
    GetFileNameFromFullPath = ""
    lngPathSeparatorPosition = InStrRev(FullPath, "\", , vbTextCompare)
    
    If lngPathSeparatorPosition <> 0 Then
        strFile = Right(FullPath, Len(FullPath) - lngPathSeparatorPosition)
        lngDotPosition = InStrRev(strFile, ".", , vbTextCompare)
        
        If lngDotPosition <> 0 Then GetFileNameFromFullPath = Left(strFile, lngDotPosition - 1)
    End If
End Function

' #####################
' Convert general path.
' #####################
Private Function CGPath(ByVal Path As String) As String
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    CGPath = Path
End Function

