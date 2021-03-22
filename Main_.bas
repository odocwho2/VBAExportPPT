Attribute VB_Name = "Main_"
Sub Main()
Application.EnableEvents = False
Application.EnableAnimations = False
Application.EnableEvents = False
OriginalWb = ThisWorkbook.Name
templatepath = Cells.Range("L1").Value
Dim PPApp As PowerPoint.Application
Dim PPPres As PowerPoint.Presentation
Set PowerPointApp = CreateObject("PowerPoint.Application")
DestinationPPT = templatepath
PowerPointApp.Presentations.Open (DestinationPPT)
Set PPApp = GetObject(class:="Powerpoint.Application")
Set PPPres = PPApp.ActivePresentation

i = 2
While Cells.Range("A" & i) <> ""
    CrtCheck = Range("A" & i)
    WbName = Range("B" & i)
    WsName = Range("C" & i)
    RangeOrChart = Range("D" & i)
    SlideNumber = Range("E" & i)
    LeftAlig = Range("F" & i)
    TopAlig = Range("G" & i)
    ShapeH = Range("H" & i)
    ShapeW = Range("I" & i)
    
    If CrtCheck = "Chart" Then
        Workbooks(WbName).Activate
        Workbooks(WbName).Worksheets(WsName).Shapes(RangeOrChart).Copy
        PPPres.Slides(SlideNumber).Select
        With PPPres.Slides(SlideNumber)
            .Shapes.PasteSpecial ppPasteEnhancedMetafile
            With .Shapes(.Shapes.Count)
                .LockAspectRatio = msoFalse
                .Left = 72 * LeftAlig
                .Top = 72 * TopAlig
                .Height = 72 * ShapeH
                .Width = 72 * ShapeW
            End With
        End With
    ElseIf CrtCheck = "Range" Then
        Workbooks(WbName).Activate
        Workbooks(WbName).Worksheets(WsName).Range(RangeOrChart).Copy
        PPPres.Slides(SlideNumber).Select
        With PPPres.Slides(SlideNumber)
            .Shapes.PasteSpecial ppPasteEnhancedMetafile
            With .Shapes(.Shapes.Count)
                .LockAspectRatio = msoFalse
                .Left = 72 * LeftAlig
                .Top = 72 * TopAlig
                .Height = 72 * ShapeH
                .Width = 72 * ShapeW
            End With
        End With
        
    ElseIf CrtCheck = "Workbook Open" Then
        Workbooks.Open (WbName)
    ElseIf CrtCheck = "Workbook Close" Then
        Workbooks(WbName).Close
    End If
    
Workbooks(OriginalWb).Activate
i = i + 1

Wend
Application.EnableAnimations = True
Application.EnableEvents = True
Application.EnableEvents = True
MsgBox "Done Exporting"
End Sub
