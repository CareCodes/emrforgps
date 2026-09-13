Attribute VB_Name = "ModuleSecurity"


Public Function EncreptedWord(ByVal SuppliedWord As String) As String
    Dim WordLength
    Dim ProcessingString As String
    Dim ProcessingNumber As Long
    Dim A As Long
    Dim b As Double
    Dim TemProcessingString As String
    WordLength = Len(SuppliedWord)
    EncreptedWord = ""
    ProcessingString = ""
    For A = 1 To WordLength
    TemProcessingString = Mid(SuppliedWord, A, 1)
    ProcessingNumber = Asc(TemProcessingString)
    If ProcessingNumber < 80 Then
        b = Rnd() * 2
        If b < 1 Then
            ProcessingString = ProcessingString & "a"
            ProcessingNumber = ((ProcessingNumber + 3) * 3) + 100
            ProcessingString = ProcessingString & ProcessingNumber
        Else
            ProcessingString = ProcessingString & "c"
            ProcessingNumber = ((ProcessingNumber + 4) * 2) + 100
            ProcessingString = ProcessingString & ProcessingNumber
        End If
    Else
            b = Rnd() * 2
        If b < 1 Then
            ProcessingString = ProcessingString & "e"
            ProcessingNumber = ((ProcessingNumber - 3) * 3) + 100
            ProcessingString = ProcessingString & ProcessingNumber
        Else
            ProcessingString = ProcessingString & "f"
            ProcessingNumber = ((ProcessingNumber - 1) * 2) + 100
            ProcessingString = ProcessingString & ProcessingNumber
        End If
    End If
    Next
    For A = 1 To Len(ProcessingString)
    TemProcessingString = Mid(ProcessingString, A, 1)
    If TemProcessingString = "0" Then
        EncreptedWord = EncreptedWord & "g"
    ElseIf TemProcessingString = "1" Then EncreptedWord = EncreptedWord & "l"
    ElseIf TemProcessingString = "2" Then EncreptedWord = EncreptedWord & "k"
    ElseIf TemProcessingString = "3" Then EncreptedWord = EncreptedWord & "p"
    ElseIf TemProcessingString = "4" Then EncreptedWord = EncreptedWord & "h"
    ElseIf TemProcessingString = "5" Then EncreptedWord = EncreptedWord & "q"
    ElseIf TemProcessingString = "6" Then EncreptedWord = EncreptedWord & "n"
    ElseIf TemProcessingString = "7" Then EncreptedWord = EncreptedWord & "i"
    ElseIf TemProcessingString = "8" Then EncreptedWord = EncreptedWord & "r"
    ElseIf TemProcessingString = "9" Then EncreptedWord = EncreptedWord & "m"
    Else
    EncreptedWord = EncreptedWord & TemProcessingString
    End If
    Next A
End Function

Public Function DecreptedWord(ByVal SuppliedWord As String) As String
    Dim WordLength
    Dim ProcessingString As String
    Dim ProcessingNumber As Long
    Dim A As Long
    Dim b As Double
    Dim TemProcessingString As String
    Dim TemTemProcessingString As String
    
    ProcessingString = ""
    For A = 1 To Len(SuppliedWord)
    TemProcessingString = Mid(SuppliedWord, A, 1)
    If TemProcessingString = "g" Then
        ProcessingString = ProcessingString & "0"
    ElseIf TemProcessingString = "l" Then ProcessingString = ProcessingString & "1"
    ElseIf TemProcessingString = "k" Then ProcessingString = ProcessingString & "2"
    ElseIf TemProcessingString = "p" Then ProcessingString = ProcessingString & "3"
    ElseIf TemProcessingString = "h" Then ProcessingString = ProcessingString & "4"
    ElseIf TemProcessingString = "q" Then ProcessingString = ProcessingString & "5"
    ElseIf TemProcessingString = "n" Then ProcessingString = ProcessingString & "6"
    ElseIf TemProcessingString = "i" Then ProcessingString = ProcessingString & "7"
    ElseIf TemProcessingString = "r" Then ProcessingString = ProcessingString & "8"
    ElseIf TemProcessingString = "m" Then ProcessingString = ProcessingString & "9"
    Else
    ProcessingString = ProcessingString & TemProcessingString
    End If
    Next A
    DecreptedWord = ""
    For A = 1 To Len(ProcessingString) Step 4
    TemProcessingString = Mid(ProcessingString, A, 4)
        For b = 1 To 4
           TemTemProcessingString = Mid(TemProcessingString, b, 1)
           Select Case TemTemProcessingString
                Case "a": ProcessingNumber = (((Val(Mid(TemProcessingString, 2, 3))) - 100) / 3) - 3
                     DecreptedWord = DecreptedWord & Chr(ProcessingNumber)
                Case "c": ProcessingNumber = (((Val(Mid(TemProcessingString, 2, 3))) - 100) / 2) - 4
                       DecreptedWord = DecreptedWord & Chr(ProcessingNumber)
                Case "e": ProcessingNumber = (((Val(Mid(TemProcessingString, 2, 3))) - 100) / 3) + 3
                     DecreptedWord = DecreptedWord & Chr(ProcessingNumber)
                Case "f": ProcessingNumber = (((Val(Mid(TemProcessingString, 2, 3))) - 100) / 2) + 1
                     DecreptedWord = DecreptedWord & Chr(ProcessingNumber)
            End Select
        Next
    Next
End Function


