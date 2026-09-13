Attribute VB_Name = "modPrintReport"
Option Explicit
    Public Enum TextAlignment
        LeftAlign
        RightAlign
        CentreAlign
        JustifiedAlign
    End Enum

    Public Type PrintColumn
        StartX As Long
        EndX As Long
        ColWidth As Long
        Topic As String
        ColText() As String
        TextAlignmant As TextAlignment
    End Type

    Public Enum PrintOrientation
        Landscape
        Potraite
    End Enum

    Public Type PrintReport
        PrintColums() As PrintColumn
        TopMargin As Long
        LeftMargin As Long
        RightMargin As Long
        BottomMargin As Long
        ColSpace As Long
        ColFontName As String
        ColFontSize As Integer
        ColFontBold As Boolean
        ColFontItalic As Boolean
        ColFontUnderline As Boolean
        ColFontStrikeThrough As Boolean
        ColTopicFontName As String
        ColTopicFontSize As Integer
        ColTopicFontBold As Boolean
        ColTopicFontItalic As Boolean
        ColTopicFontUnderline As Boolean
        ColTopicFontStrikeThrough As Boolean
        ColTopicAlignment As TextAlignment
        TopFontName As String
        TopFontSize As Integer
        TopFontBold As Boolean
        TopFontItalic As Boolean
        TopFontUnderline As Boolean
        TopFontStrikeThrough As Boolean
        TopAlignment As TextAlignment
        ReportPrinter As String
        ReportPaper As String
        TopicFontName As String
        TopicFontSize As Integer
        TopicFontBold As Boolean
        TopicFontItalic As Boolean
        TopicFontUnderline As Boolean
        TopicFontStrikeThrough As Boolean
        TopicAlignment As TextAlignment
        SubTopicFontName As String
        SubTopicFontSize As Integer
        SubTopicFontBold As Boolean
        SubTopicFontItalic As Boolean
        SubTopicFontUnderline As Boolean
        SubTopicFontStrikeThrough As Boolean
        SubTopicAlignment As TextAlignment
        Topic As String
        Subtopic As String
        Header As String
        Footer As String
        HeaderFontName As String
        HeaderFontSize As Integer
        HeaderFontBold As Boolean
        HeaderFontItalic As Boolean
        HeaderFontUnderline As Boolean
        HeaderX As Long
        HeaderY As Long
        HeaderAlignment As TextAlignment
        FooterFontName As String
        FooterFontSize As Integer
        FooterFontBold As Boolean
        FooterFontItalic As Boolean
        FooterFontUnderline As Boolean
        FooterX As Long
        FooterY As Long
        FooterAlignment As TextAlignment
        PageNoX As Long
        PageNoY As Long
        PageNoFontName As String
        PageNoFontSize As Integer
        PageNoFontBold As Boolean
        PageNoFontItalic As Boolean
        PageNoFontUnderline As Boolean
        PageNoAlignment As TextAlignment
        ReportPrintOrientation As PrintOrientation
    End Type

Public Sub GetPrintDefaults(MyPrintDefaults As PrintReport)
    Dim MyPrinter As Printer
    For Each MyPrinter In VB.Printers
        If Printer.DeviceName = MyPrinter.DeviceName Then
            MyPrintDefaults.ReportPrinter = MyPrinter.DeviceName
        End If
    Next
    With MyPrintDefaults
        .ReportPrintOrientation = Potraite
        .TopMargin = 1440 / 2
        .LeftMargin = 1440 / 2
        .RightMargin = Printer.Width - (1440)
        .BottomMargin = Printer.height - (1440)
        .ColSpace = 350
        .ColFontName = "Verdana"
        .ColFontSize = 10
        .ColFontBold = False
        .ColFontItalic = False
        .ColFontUnderline = False
        .ColFontStrikeThrough = False
        .ColTopicAlignment = LeftAlign
        .ColTopicFontName = "Verdana"
        .ColTopicFontSize = 10
        .ColTopicFontBold = True
        .ColTopicFontItalic = False
        .ColTopicFontUnderline = True
        .ColTopicFontStrikeThrough = False
        .ColTopicAlignment = LeftAlign
        .TopFontName = "Verdana"
        .TopFontSize = 11
        .TopFontBold = True
        .TopFontItalic = False
        .TopFontUnderline = False
        .TopFontStrikeThrough = False
        .TopAlignment = CentreAlign
        .ReportPaper = "A4"
        .TopicFontName = "Verdana"
        .TopicFontSize = 13
        .TopicFontBold = True
        .TopicFontItalic = False
        .TopicFontUnderline = False
        .TopicFontStrikeThrough = False
        .TopicAlignment = CentreAlign
        .SubTopicFontName = "Verdana"
        .SubTopicFontSize = 12
        .SubTopicFontBold = True
        .SubTopicFontItalic = False
        .SubTopicFontUnderline = False
        .SubTopicFontStrikeThrough = False
        .SubTopicAlignment = CentreAlign
        .Topic = InstitutionName
        .Subtopic = InstitutionAddress
        .Header = ""
        .Footer = ""
        .HeaderFontName = "Verdana"
        .HeaderFontSize = 9
        .HeaderFontBold = False
        .HeaderFontItalic = False
        .HeaderFontUnderline = False
        .HeaderX = 1440
        .HeaderY = 1440 / 2
        .HeaderAlignment = LeftAlign
        .FooterFontName = "Verdana"
        .FooterFontSize = 9
        .FooterFontBold = False
        .FooterFontItalic = False
        .FooterFontUnderline = False
        .FooterX = 1440 * 0.5
        .FooterY = Printer.height - (1440 * 0.6)
        .FooterAlignment = RightAlign
        .PageNoX = Printer.Width - 1440 * 2
        .PageNoY = Printer.height - (1440 * 0.8)
        .PageNoFontName = "Verdana"
        .PageNoFontSize = 9
        .PageNoFontBold = False
        .PageNoFontItalic = False
        .PageNoFontUnderline = False
        .PageNoAlignment = CentreAlign
    End With
End Sub

Public Sub PrintMyReport(ReportToPrint As PrintReport, Optional AutoColumns As Boolean)
    Dim MyCol As Long
    Dim MyRow As Long
    Dim NowX As Long
    Dim NowY As Long
    Dim TotalCols As Integer
    Dim ThisPage As Integer
    
    ThisPage = 1
    On Error Resume Next
    If ReportToPrint.ReportPrintOrientation = Landscape Then
        Printer.Orientation = vbPRORLandscape '  PrinterOrientationConstants.cdlLandscape
    Else
        Printer.Orientation = PrinterOrientationConstants.cdlPortrait
    End If
    
    If AutoColumns = True Then
        NowX = ReportToPrint.LeftMargin
        For MyCol = 0 To UBound(ReportToPrint.PrintColums) - 1
            ReportToPrint.PrintColums(MyCol).StartX = NowX
            For MyRow = 0 To UBound(ReportToPrint.PrintColums(MyCol).ColText) - 1
                If Printer.TextWidth(ReportToPrint.PrintColums(MyCol).ColText(MyRow)) > ReportToPrint.PrintColums(MyCol).ColWidth Then
                    ReportToPrint.PrintColums(MyCol).ColWidth = Printer.TextWidth(ReportToPrint.PrintColums(MyCol).ColText(MyRow) & Space(2))
                End If
            Next
            If Printer.TextWidth(ReportToPrint.PrintColums(MyCol).Topic) > ReportToPrint.PrintColums(MyCol).ColWidth Then
            ReportToPrint.PrintColums(MyCol).ColWidth = Printer.TextWidth(ReportToPrint.PrintColums(MyCol).Topic & Space(2))
            End If
            
            ReportToPrint.PrintColums(MyCol).EndX = NowX + ReportToPrint.PrintColums(MyCol).ColWidth
            NowX = NowX + ReportToPrint.PrintColums(MyCol).ColWidth + ReportToPrint.ColSpace
        Next
    End If
    
    NowY = 0
    
    Printer.FontName = ReportToPrint.TopicFontName
    Printer.FontSize = ReportToPrint.TopicFontSize
    Printer.FontBold = ReportToPrint.TopicFontBold
    Printer.FontItalic = ReportToPrint.TopicFontItalic
    Printer.FontStrikethru = ReportToPrint.TopicFontStrikeThrough
    Printer.FontUnderline = ReportToPrint.TopicFontUnderline
    NowY = NowY + ReportToPrint.TopMargin
    Printer.CurrentY = NowY
    
    If ReportToPrint.TopicAlignment = LeftAlign Then
        Printer.CurrentX = ReportToPrint.LeftMargin
    ElseIf ReportToPrint.TopicAlignment = RightAlign Then
        Printer.CurrentX = ReportToPrint.RightMargin - Printer.TextWidth(ReportToPrint.Topic) + (Printer.TextWidth(ReportToPrint.Topic))
    ElseIf ReportToPrint.TopicAlignment = CentreAlign Then
        Printer.CurrentX = (((ReportToPrint.RightMargin + ReportToPrint.LeftMargin) / 2) - (Printer.TextWidth(ReportToPrint.Topic) / 2))
    End If
    
    Printer.Print ReportToPrint.Topic
    NowY = NowY + Printer.TextHeight(ReportToPrint.Topic)
    
    Printer.FontName = ReportToPrint.SubTopicFontName
    Printer.FontSize = ReportToPrint.SubTopicFontSize
    Printer.FontBold = ReportToPrint.SubTopicFontBold
    Printer.FontItalic = ReportToPrint.SubTopicFontItalic
    Printer.FontStrikethru = ReportToPrint.SubTopicFontStrikeThrough
    Printer.FontUnderline = ReportToPrint.SubTopicFontUnderline
    Printer.CurrentY = NowY
    If ReportToPrint.SubTopicAlignment = LeftAlign Then
        Printer.CurrentX = ReportToPrint.LeftMargin
    ElseIf ReportToPrint.SubTopicAlignment = RightAlign Then
        Printer.CurrentX = ReportToPrint.RightMargin - Printer.TextWidth(ReportToPrint.Subtopic)
    ElseIf ReportToPrint.SubTopicAlignment = CentreAlign Then
        Printer.CurrentX = (((ReportToPrint.RightMargin + ReportToPrint.LeftMargin) / 2) - (Printer.TextWidth(ReportToPrint.Subtopic) / 2))
    End If
    Printer.Print ReportToPrint.Subtopic
    NowY = NowY + Printer.TextHeight(ReportToPrint.Subtopic)
    
    Printer.FontName = ReportToPrint.ColTopicFontName
    Printer.FontSize = ReportToPrint.ColTopicFontSize
    Printer.FontBold = ReportToPrint.ColTopicFontBold
    Printer.FontItalic = ReportToPrint.ColTopicFontItalic
    Printer.FontStrikethru = ReportToPrint.ColTopicFontStrikeThrough
    Printer.FontUnderline = ReportToPrint.ColTopicFontUnderline
    For MyCol = 0 To UBound(ReportToPrint.PrintColums) - 1
        Printer.CurrentY = NowY
        If ReportToPrint.PrintColums(MyCol).TextAlignmant = LeftAlign Then
            Printer.CurrentX = ReportToPrint.PrintColums(MyCol).StartX
        ElseIf ReportToPrint.PrintColums(MyCol).TextAlignmant = RightAlign Then
            Printer.CurrentX = ReportToPrint.PrintColums(MyCol).EndX - Printer.TextWidth(ReportToPrint.PrintColums(MyCol).Topic)
        ElseIf ReportToPrint.PrintColums(MyCol).TextAlignmant = CentreAlign Then
            Printer.CurrentX = (((ReportToPrint.PrintColums(MyCol).EndX + ReportToPrint.PrintColums(MyCol).StartX) / 2) - (Printer.TextWidth(ReportToPrint.PrintColums(MyCol).Topic) / 2))
        End If
        Printer.Print ReportToPrint.PrintColums(MyCol).Topic
    Next
    NowY = NowY + Printer.TextHeight(ReportToPrint.PrintColums(MyCol).Topic)

    
    TotalCols = UBound(ReportToPrint.PrintColums(0).ColText)
    
    For MyRow = 0 To TotalCols - 1
        For MyCol = 0 To UBound(ReportToPrint.PrintColums) - 1
            Printer.CurrentY = NowY
            Printer.FontName = ReportToPrint.ColFontName
            Printer.FontSize = ReportToPrint.ColFontSize
            Printer.FontBold = ReportToPrint.ColFontBold
            Printer.FontItalic = ReportToPrint.ColFontItalic
            Printer.FontStrikethru = ReportToPrint.ColFontStrikeThrough
            Printer.FontUnderline = ReportToPrint.ColFontUnderline
            If ReportToPrint.PrintColums(MyCol).TextAlignmant = LeftAlign Then
                Printer.CurrentX = ReportToPrint.PrintColums(MyCol).StartX
            ElseIf ReportToPrint.PrintColums(MyCol).TextAlignmant = RightAlign Then
                Printer.CurrentX = ReportToPrint.PrintColums(MyCol).EndX - Printer.TextWidth(ReportToPrint.PrintColums(MyCol).ColText(MyRow))
            ElseIf ReportToPrint.PrintColums(MyCol).TextAlignmant = CentreAlign Then
                Printer.CurrentX = (((ReportToPrint.PrintColums(MyCol).EndX + ReportToPrint.PrintColums(MyCol).StartX) / 2) - (Printer.TextWidth(ReportToPrint.PrintColums(MyCol).ColText(MyRow)) / 2))
            End If
            Printer.Print ReportToPrint.PrintColums(MyCol).ColText(MyRow)
        Next
        
        NowY = NowY + Printer.TextHeight(ReportToPrint.PrintColums(0).ColText(0))
        
        If NowY > ReportToPrint.BottomMargin Then
            Printer.FontName = ReportToPrint.PageNoFontName
            Printer.FontSize = ReportToPrint.PageNoFontSize
            Printer.FontBold = ReportToPrint.PageNoFontBold
            Printer.FontItalic = ReportToPrint.PageNoFontItalic
            Printer.FontUnderline = ReportToPrint.PageNoFontUnderline
            If ReportToPrint.ReportPrintOrientation = Potraite Then
                Printer.CurrentX = ReportToPrint.PageNoX
                Printer.CurrentY = ReportToPrint.PageNoY
            Else
                Printer.CurrentX = ReportToPrint.PageNoX
                Printer.CurrentY = ReportToPrint.PageNoY
            End If
            Printer.Print "Page No. " & ThisPage
            
            Printer.FontName = ReportToPrint.HeaderFontName
            Printer.FontSize = ReportToPrint.HeaderFontSize
            Printer.FontBold = ReportToPrint.HeaderFontBold
            Printer.FontItalic = ReportToPrint.HeaderFontItalic
            Printer.FontUnderline = ReportToPrint.HeaderFontUnderline
            Printer.CurrentX = ReportToPrint.HeaderX
            Printer.CurrentY = ReportToPrint.HeaderY
            Printer.Print ReportToPrint.Header
            
        
            Printer.FontName = ReportToPrint.FooterFontName
            Printer.FontSize = ReportToPrint.FooterFontSize
            Printer.FontBold = ReportToPrint.FooterFontBold
            Printer.FontItalic = ReportToPrint.FooterFontItalic
            Printer.FontUnderline = ReportToPrint.FooterFontUnderline
            Printer.CurrentX = ReportToPrint.FooterX
            Printer.CurrentY = ReportToPrint.FooterY
            Printer.Print ReportToPrint.Footer
            
            Printer.NewPage
            NowY = ReportToPrint.TopMargin
            ThisPage = ThisPage + 1
            
            
        End If
    Next
    
    Printer.FontName = ReportToPrint.PageNoFontName
    Printer.FontSize = ReportToPrint.PageNoFontSize
    Printer.FontBold = ReportToPrint.PageNoFontBold
    Printer.FontItalic = ReportToPrint.PageNoFontItalic
    Printer.FontUnderline = ReportToPrint.PageNoFontUnderline
    Printer.CurrentX = ReportToPrint.PageNoX
    Printer.CurrentY = ReportToPrint.PageNoY
    Printer.Print "Page No. " & ThisPage
    
    Printer.FontName = ReportToPrint.HeaderFontName
    Printer.FontSize = ReportToPrint.HeaderFontSize
    Printer.FontBold = ReportToPrint.HeaderFontBold
    Printer.FontItalic = ReportToPrint.HeaderFontItalic
    Printer.FontUnderline = ReportToPrint.HeaderFontUnderline
    Printer.CurrentX = ReportToPrint.HeaderX
    Printer.CurrentY = ReportToPrint.HeaderY
    Printer.Print ReportToPrint.Header
    

    Printer.FontName = ReportToPrint.FooterFontName
    Printer.FontSize = ReportToPrint.FooterFontSize
    Printer.FontBold = ReportToPrint.FooterFontBold
    Printer.FontItalic = ReportToPrint.FooterFontItalic
    Printer.FontUnderline = ReportToPrint.FooterFontUnderline
    Printer.CurrentX = ReportToPrint.FooterX
    Printer.CurrentY = ReportToPrint.FooterY
    Printer.Print ReportToPrint.Footer
    
    Printer.EndDoc

End Sub

Public Sub GridPrint(PrintGrid As MSFlexGrid, MyPrintReport As PrintReport, Optional PrintTopic As String, Optional PrintSubTopic As String)
    If PrintGrid.Rows <= 1 Then
        MsgBox "Noting to Print"
        Exit Sub
    End If
    
    Dim GridReport As PrintReport
    Dim GridReportCols() As PrintColumn
    Dim ColData() As String
    Dim NumericCount As Long
    Dim NonNumericCount As Long
    
    Dim PrintableCols As Integer
    Dim i As Integer
    Dim n As Integer
    Dim MyCol As Long
    Dim MyRow As Long
    
    
    GridReport = MyPrintReport
    
    For i = 0 To PrintGrid.Cols - 1
        If PrintGrid.ColWidth(i) > 10 Then PrintableCols = PrintableCols + 1
    Next
    
    ReDim GridReportCols(PrintableCols)
    ReDim ColData(PrintGrid.Rows - 1)

    n = 0
    For MyCol = 0 To PrintGrid.Cols - 1
        PrintGrid.col = MyCol
        If PrintGrid.ColWidth(MyCol) > 10 Then
            GridReportCols(n).Topic = PrintGrid.TextMatrix(0, MyCol)
            NumericCount = 0
            NonNumericCount = 0
            For MyRow = 0 To PrintGrid.Rows - 2
                ColData(MyRow) = PrintGrid.TextMatrix(MyRow + 1, MyCol)
                If IsNumeric(PrintGrid.TextMatrix(MyRow + 1, MyCol)) = True Then
                    NumericCount = NumericCount + 1
                Else
                    NonNumericCount = NonNumericCount + 1
                End If
            Next MyRow
            PrintGrid.row = 1
            GridReportCols(n).ColText() = ColData()
            If NumericCount > NonNumericCount Then
                GridReportCols(n).TextAlignmant = RightAlign
            Else
                GridReportCols(n).TextAlignmant = LeftAlign
            End If
            n = n + 1
        End If
    Next MyCol
    

    
    GridReport.Topic = PrintTopic
    GridReport.Subtopic = PrintSubTopic
    GridReport.PrintColums() = GridReportCols()
    
    Call PrintMyReport(GridReport, True)


End Sub


