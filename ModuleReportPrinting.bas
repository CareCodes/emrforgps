Attribute VB_Name = "ModuleReportPrinting"
Option Explicit

Public Function PrintReportText(ByVal SuppliedText As String, X1 As Long, X2 As Long, Y1 As Long, Y2 As Long, PrintFontName As String, PrintFontSize As Long, PrintFontBold As Boolean, PrintFontItalic As Boolean, PrintFontUnderline As Boolean, PrintFontStrikethrough As Boolean, PrintAlignment As Integer, CheckOk As Boolean) As Boolean
    PrintReportText = SingleLinePrint(SuppliedText, X1, X2, Y1, Y2, PrintFontName, PrintFontSize, PrintFontBold, PrintFontItalic, PrintFontUnderline, PrintFontStrikethrough, PrintAlignment, CheckOk)
    'SuperScriptPrint SuppliedText, X1, X2, Y1, Y2, PrintFontName, PrintFontSize, PrintFontBold, PrintFontItalic, PrintFontUnderline, PrintFontStrikethrough, PrintAlignment
    
End Function
    

Private Function SingleLinePrint(ByVal SuppliedText As String, X1 As Long, X2 As Long, Y1 As Long, Y2 As Long, PrintFontName As String, PrintFontSize As Long, PrintFontBold As Boolean, PrintFontItalic As Boolean, PrintFontUnderline As Boolean, PrintFontStrikethrough As Boolean, PrintAlignment As Integer, CheckOk As Boolean) As Boolean
    Dim RemainingString As String
    Dim PrintString As String
    Dim RemainingLength As Long
    Dim LineBrackLength As Long
    Dim NewY1 As Long
    Dim ThisY1 As Long
            Printer.Font.Name = PrintFontName
            Printer.Font.Size = PrintFontSize
            Printer.Font.Bold = PrintFontBold
            Printer.Font.Underline = PrintFontUnderline
            Printer.Font.Strikethrough = PrintFontStrikethrough
    NewY1 = Y1
    RemainingString = SuppliedText
    While InStr(RemainingString, vbNewLine) > 0
        RemainingLength = Len(RemainingString)
        LineBrackLength = InStr(RemainingString, vbNewLine)
        PrintString = Left(RemainingString, LineBrackLength - 1)
        RemainingString = Right(RemainingString, RemainingLength - LineBrackLength - 1)
        SingleLinePrint = BreakLinePrint(PrintString, X1, X2, NewY1, Y2, PrintFontName, PrintFontSize, PrintFontBold, PrintFontItalic, PrintFontUnderline, PrintFontStrikethrough, PrintAlignment, CheckOk)
        ThisY1 = Printer.TextHeight(PrintString)
        NewY1 = NewY1 + ThisY1
    Wend
    SingleLinePrint = BreakLinePrint(RemainingString, X1, X2, NewY1, Y2, PrintFontName, PrintFontSize, PrintFontBold, PrintFontItalic, PrintFontUnderline, PrintFontStrikethrough, PrintAlignment, CheckOk)
End Function

Private Function BreakLinePrint(ByVal SuppliedText As String, X1 As Long, X2 As Long, Y1 As Long, Y2 As Long, PrintFontName As String, PrintFontSize As Long, PrintFontBold As Boolean, PrintFontItalic As Boolean, PrintFontUnderline As Boolean, PrintFontStrikethrough As Boolean, PrintAlignment As Integer, CheckOk As Boolean) As Boolean
    Dim PrintedText As String
    Dim ToPrint As String
    Dim RemainingText As String
    Dim LastDevide As Long
    Dim PrintCount As Long
    PrintCount = 0
    ToPrint = SuppliedText
    RemainingText = Empty
    
    While Trim(ToPrint) <> ""
        While Printer.TextWidth(ToPrint) > X2 - X1
            PrintCount = PrintCount + 1
            If InStr(ToPrint, " ") Then
                LastDevide = InStrRev(ToPrint, " ")
                RemainingText = Right(ToPrint, Len(ToPrint) - LastDevide + 1) & RemainingText
                ToPrint = Left(ToPrint, LastDevide - 1)
            Else
            
            End If
        Wend
        BreakLinePrint = SuperScriptPrint(Trim(ToPrint), X1, X2, Y1, Y2, PrintFontName, PrintFontSize, PrintFontBold, PrintFontItalic, PrintFontUnderline, PrintFontStrikethrough, PrintAlignment, CheckOk)
        If PrintCount > 0 Then
            Y1 = Y1 + Printer.TextHeight(ToPrint)
        End If
        ToPrint = RemainingText
        RemainingText = Empty
    Wend
    
End Function

Private Function SuperScriptPrint(ByVal SuppliedText As String, X1 As Long, X2 As Long, Y1 As Long, Y2 As Long, PrintFontName As String, PrintFontSize As Long, PrintFontBold As Boolean, PrintFontItalic As Boolean, PrintFontUnderline As Boolean, PrintFontStrikethrough As Boolean, PrintAlignment As Integer, CheckOk As Boolean) As Boolean
    SuperScriptPrint = True
    Dim SuperscriptPresent As Boolean
    Dim SuperscriptLocation As Long
    Dim BeforeSuperscriptText As String
    Dim AfterSuperscriptText As String
    Dim SuperScriptText As String
    Dim SpaceLocation As Long
    Dim SuperscriptElevation As Double
    
    Dim TotalTextWidth As Long
    Dim TextWidth1 As Long
    Dim TextWidth2 As Long
    Dim TextWidth3 As Long
    
    SuperscriptLocation = InStr(1, SuppliedText, "^")
    SuperscriptElevation = 0.2
    
    If SuperscriptLocation = 0 Then
        Printer.Font.Name = PrintFontName
        Printer.Font.Size = PrintFontSize
        Printer.Font.Bold = PrintFontBold
        Printer.Font.Underline = PrintFontUnderline
        Printer.Font.Strikethrough = PrintFontStrikethrough
        Printer.CurrentY = Y1
        If PrintAlignment = TextAlignment.LeftAlign Then
            Printer.CurrentX = X1
        ElseIf PrintAlignment = TextAlignment.RightAlign Then
            Printer.CurrentX = X2 - Printer.TextWidth(SuppliedText)
        ElseIf PrintAlignment = TextAlignment.CentreAlign Then
            Printer.CurrentX = ((X1 + X2) / 2) - (Printer.TextWidth(SuppliedText) / 2)
        Else
            Printer.CurrentX = ((X1 + X2) / 2) - (Printer.TextWidth(SuppliedText) / 2)
        End If
        Printer.Print SuppliedText
    Else
        If SuperscriptLocation > 0 Then BeforeSuperscriptText = Left(SuppliedText, SuperscriptLocation - 1)
        
        AfterSuperscriptText = Right(SuppliedText, (Len(SuppliedText) - SuperscriptLocation))
        
        If InStr(1, AfterSuperscriptText, " ") = 0 Then
            
            AfterSuperscriptText = Trim(AfterSuperscriptText)
            
            Printer.Font.Name = PrintFontName
            Printer.Font.Size = PrintFontSize
            Printer.Font.Bold = PrintFontBold
            Printer.Font.Underline = PrintFontUnderline
            Printer.Font.Strikethrough = PrintFontStrikethrough
            
            TextWidth1 = Printer.TextWidth(BeforeSuperscriptText)
            
            Printer.FontSize = PrintFontSize - 2
            
            TextWidth2 = Printer.TextWidth(AfterSuperscriptText)
            
            TextWidth3 = 0
            
            TotalTextWidth = TextWidth1 + TextWidth2 + TextWidth3
            
            Printer.Font.Name = PrintFontName
            Printer.Font.Size = PrintFontSize
            Printer.Font.Bold = PrintFontBold
            Printer.Font.Underline = PrintFontUnderline
            Printer.Font.Strikethrough = PrintFontStrikethrough
            
            If PrintAlignment = TextAlignment.LeftAlign Then
                Printer.CurrentX = X1
            ElseIf PrintAlignment = TextAlignment.RightAlign Then
                Printer.CurrentX = X2 - TotalTextWidth
            ElseIf PrintAlignment = TextAlignment.CentreAlign Then
                Printer.CurrentX = ((X1 + X2) / 2) - (TotalTextWidth / 2)
            Else
                Printer.CurrentX = ((X1 + X2) / 2) - (TotalTextWidth / 2)
            End If
            
            Printer.CurrentY = Y1
                
            Dim i As Integer
                
            If CheckOk = True Then
                If Printer.CurrentX < 0 Or Printer.CurrentX > Printer.Width Then
                    i = MsgBox("The Width of the Printed Items fall outside the Printeble are. Do you want to continue?", vbYesNo)
                    If i = vbNo Then
                        Printer.KillDoc
                        SuperScriptPrint = False
                        Exit Function
                    End If
                End If
                If Printer.CurrentY < 0 Or Printer.CurrentX > Printer.Height Then
                    i = MsgBox("The height of the Printed Items fall outside the Printeble are. Do you want to continue?", vbYesNo)
                    If i = vbNo Then
                        Printer.KillDoc
                        SuperScriptPrint = False
                        Exit Function
                    End If
                End If
            End If
            
            Printer.Print BeforeSuperscriptText
            
            
            
            If PrintAlignment = TextAlignment.LeftAlign Then
                Printer.CurrentX = X1 + TextWidth1
            ElseIf PrintAlignment = TextAlignment.RightAlign Then
                Printer.CurrentX = X2 - TotalTextWidth + TextWidth1
            ElseIf PrintAlignment = TextAlignment.CentreAlign Then
                Printer.CurrentX = ((X1 + X2) / 2) - (TotalTextWidth / 2) + TextWidth1
            Else
                Printer.CurrentX = ((X1 + X2) / 2) - (TotalTextWidth / 2) + TextWidth1
            End If
            Printer.CurrentY = Y1 - (Printer.TextHeight(BeforeSuperscriptText) * SuperscriptElevation)
            Printer.FontSize = PrintFontSize - 2
                        
            If CheckOk = True Then
                If Printer.CurrentX < 0 Or Printer.CurrentX > Printer.Width Then
                    i = MsgBox("The Width of the Printed Items fall outside the Printeble are. Do you want to continue?", vbYesNo)
                    If i = vbNo Then
                        Printer.KillDoc
                        SuperScriptPrint = False
                        Exit Function
                    End If
                End If
                If Printer.CurrentY < 0 Or Printer.CurrentX > Printer.Height Then
                    i = MsgBox("The height of the Printed Items fall outside the Printeble are. Do you want to continue?", vbYesNo)
                    If i = vbNo Then
                        Printer.KillDoc
                        SuperScriptPrint = False
                        Exit Function
                    End If
                End If
            End If
            
            Printer.Print AfterSuperscriptText
            
            Printer.FontSize = PrintFontSize
            
        Else
            SpaceLocation = InStr(1, AfterSuperscriptText, " ")
            SuperScriptText = Left(AfterSuperscriptText, SpaceLocation)
            AfterSuperscriptText = Right(AfterSuperscriptText, (Len(AfterSuperscriptText) - SpaceLocation))
            Printer.Font.Name = PrintFontName
            Printer.Font.Size = PrintFontSize
            Printer.Font.Bold = PrintFontBold
            Printer.Font.Underline = PrintFontUnderline
            Printer.Font.Strikethrough = PrintFontStrikethrough
            
            TextWidth1 = Printer.TextWidth(BeforeSuperscriptText)
            Printer.FontSize = PrintFontSize - 2
            TextWidth2 = Printer.TextWidth(SuperScriptText)
            Printer.FontSize = PrintFontSize
            TextWidth3 = Printer.TextWidth(AfterSuperscriptText)
            TotalTextWidth = TextWidth1 + TextWidth2 + TextWidth3
            
            If PrintAlignment = TextAlignment.LeftAlign Then
                Printer.CurrentX = X1
            ElseIf PrintAlignment = TextAlignment.RightAlign Then
                Printer.CurrentX = X2 - TotalTextWidth
            ElseIf PrintAlignment = TextAlignment.CentreAlign Then
                Printer.CurrentX = ((X1 + X2) / 2) - (TotalTextWidth / 2)
            Else
                Printer.CurrentX = ((X1 + X2) / 2) - (TotalTextWidth / 2)
            End If
            
            Printer.CurrentY = Y1
            
            If CheckOk = True Then
                
                If Printer.CurrentX < 0 Or Printer.CurrentX > Printer.Width Then
                    i = MsgBox("The Width of the Printed Items fall outside the Printeble are. Do you want to continue?", vbYesNo)
                    If i = vbNo Then
                        Printer.KillDoc
                        SuperScriptPrint = False
                        Exit Function
                    End If
                End If
                If Printer.CurrentY < 0 Or Printer.CurrentX > Printer.Height Then
                    i = MsgBox("The height of the Printed Items fall outside the Printeble are. Do you want to continue?", vbYesNo)
                    If i = vbNo Then
                        Printer.KillDoc
                        SuperScriptPrint = False
                        Exit Function
                    End If
                End If
                
            End If
            
            Printer.Print BeforeSuperscriptText
            
            
            If PrintAlignment = TextAlignment.LeftAlign Then
                Printer.CurrentX = X1 + TextWidth1
            ElseIf PrintAlignment = TextAlignment.RightAlign Then
                Printer.CurrentX = X2 - TotalTextWidth + TextWidth1
            ElseIf PrintAlignment = TextAlignment.CentreAlign Then
                Printer.CurrentX = ((X1 + X2) / 2) - (TotalTextWidth / 2) + TextWidth1
            Else
                Printer.CurrentX = ((X1 + X2) / 2) - (TotalTextWidth / 2) + TextWidth1
            End If
            
            
            Printer.CurrentY = Y1 - (Printer.TextHeight(BeforeSuperscriptText) * SuperscriptElevation)
            Printer.FontSize = PrintFontSize - 2
                
            If CheckOk = True Then
                If Printer.CurrentX < 0 Or Printer.CurrentX > Printer.Width Then
                    i = MsgBox("The Width of the Printed Items fall outside the Printeble are. Do you want to continue?", vbYesNo)
                    If i = vbNo Then
                        Printer.KillDoc
                        SuperScriptPrint = False
                        Exit Function
                    End If
                End If
                If Printer.CurrentY < 0 Or Printer.CurrentX > Printer.Height Then
                    i = MsgBox("The height of the Printed Items fall outside the Printeble are. Do you want to continue?", vbYesNo)
                    If i = vbNo Then
                        Printer.KillDoc
                        SuperScriptPrint = False
                        Exit Function
                    End If
                End If
            End If
            
            Printer.Print SuperScriptText
            
            
            
            Printer.FontSize = PrintFontSize
            Printer.CurrentY = Y1
            
            If PrintAlignment = TextAlignment.LeftAlign Then
                Printer.CurrentX = X1 + TextWidth1 + TextWidth2
            ElseIf PrintAlignment = TextAlignment.RightAlign Then
                Printer.CurrentX = X2 - TotalTextWidth + TextWidth1 + TextWidth2
            ElseIf PrintAlignment = TextAlignment.CentreAlign Then
                Printer.CurrentX = ((X1 + X2) / 2) - (TotalTextWidth / 2) + TextWidth1 + TextWidth2
            Else
                Printer.CurrentX = ((X1 + X2) / 2) - (TotalTextWidth / 2) + TextWidth1 + TextWidth2
            End If
            
            If CheckOk = True Then
                If Printer.CurrentX < 0 Or Printer.CurrentX > Printer.Width Then
                    i = MsgBox("The Width of the Printed Items fall outside the Printeble are. Do you want to continue?", vbYesNo)
                    If i = vbNo Then
                        Printer.KillDoc
                        SuperScriptPrint = False
                        Exit Function
                    End If
                End If
                If Printer.CurrentY < 0 Or Printer.CurrentX > Printer.Height Then
                    i = MsgBox("The height of the Printed Items fall outside the Printeble are. Do you want to continue?", vbYesNo)
                    If i = vbNo Then
                        Printer.KillDoc
                        SuperScriptPrint = False
                        Exit Function
                    End If
                End If
            End If
            
            Printer.Print AfterSuperscriptText
        End If
    End If
End Function

