VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function NumberToWord(ByVal Number As Long) As String
    NumberToWord = WordNumber(Number, True)
End Function

Private Function ReplaceLastComma(SuppliedWord As String) As String
    Dim LastCommaPlace As Long
    Dim LeftString As String
    Dim RightString As String
    Dim WordLength As Long
    
    ReplaceLastComma = SuppliedWord
    If InStr(SuppliedWord, ",") <= 0 Then Exit Function
    WordLength = Len(SuppliedWord)
    LastCommaPlace = InStrRev(SuppliedWord, ",")
    LeftString = Left(SuppliedWord, LastCommaPlace - 1)
    RightString = Right(SuppliedWord, WordLength - LastCommaPlace)
    ReplaceLastComma = LeftString & " and" & RightString
End Function



Private Function WordNumber(ByVal Number As Long, Comma As Boolean) As String
    Dim temNumber As Long
    Dim TemString As String
    
    
    WordNumber = ""
    
    Do
        If Number >= 1000000000 Then
            temNumber = Number \ 1000000000
            TemString = WordNumber(temNumber, False)
            WordNumber = WordNumber & " " & TemString & " Billion"
            Number = Number - temNumber * 1000000000
            If Number = 0 Then
                Exit Do
            Else
                If Comma = True Then
                    WordNumber = WordNumber & ","
                End If
            End If
        ElseIf Number >= 1000000 Then
            temNumber = Number \ 1000000
            TemString = WordNumber(temNumber, False)
            WordNumber = WordNumber & " " & TemString & " Million"
            Number = Number - temNumber * 1000000
            If Number = 0 Then
                Exit Do
            Else
                If Comma = True Then
                    WordNumber = WordNumber & ","
                End If
            End If
        ElseIf Number >= 1000 Then
            temNumber = Number \ 1000
            TemString = WordNumber(temNumber, False)
            WordNumber = WordNumber & " " & TemString & " Thousand"
            Number = Number - temNumber * 1000
            If Number = 0 Then
                Exit Do
            Else
                If Comma = True Then
                    WordNumber = WordNumber & ","
                End If
            End If
        ElseIf Number >= 100 Then
            temNumber = Number \ 100
            TemString = WordNumber(temNumber, False)
            WordNumber = WordNumber & " " & TemString & " Hundred"
            Number = Number - temNumber * 100
            If Number = 0 Then
                Exit Do
            Else
                If Comma = True Then
                    WordNumber = WordNumber & ","
                End If
            End If
        ElseIf Number > 90 Then
            Number = Number - 90
            WordNumber = WordNumber & " " & "Nienty"
        ElseIf Number >= 80 Then
            Number = Number - 80
            WordNumber = WordNumber & " " & "Eighty"
        ElseIf Number >= 70 Then
            Number = Number - 70
            WordNumber = WordNumber & " " & "Seventy"
        ElseIf Number >= 60 Then
            Number = Number - 60
            WordNumber = WordNumber & " " & "Sixty"
        ElseIf Number >= 50 Then
            Number = Number - 50
            WordNumber = WordNumber & " " & "Fifty"
        ElseIf Number >= 40 Then
            Number = Number - 40
            WordNumber = WordNumber & " " & "Fourty"
        ElseIf Number >= 30 Then
            Number = Number - 30
            WordNumber = WordNumber & " " & "Thirty"
        ElseIf Number >= 20 Then
            Number = Number - 20
            WordNumber = WordNumber & " " & "Twenty"
        ElseIf Number = 19 Then
            Number = Number - 19
            WordNumber = WordNumber & " " & "Nineteen"
        ElseIf Number = 18 Then
            Number = Number - 18
            WordNumber = WordNumber & " " & "Eighteen"
        ElseIf Number = 17 Then
            Number = Number - 17
            WordNumber = WordNumber & " " & "Seventeen"
        ElseIf Number = 16 Then
            Number = Number - 16
            WordNumber = WordNumber & " " & "Sixteen"
        ElseIf Number = 15 Then
            Number = Number - 15
            WordNumber = WordNumber & " " & "Fiftween"
        ElseIf Number = 14 Then
            Number = Number - 14
            WordNumber = WordNumber & " " & "Fourteen"
        ElseIf Number = 13 Then
            Number = Number - 13
            WordNumber = WordNumber & " " & "Thirteen"
        ElseIf Number = 12 Then
            Number = Number - 12
            WordNumber = WordNumber & " " & "Twelve"
        ElseIf Number = 11 Then
            Number = Number - 11
            WordNumber = WordNumber & " " & "Eleven"
        ElseIf Number = 10 Then
            Number = Number - 10
            WordNumber = WordNumber & " " & "Ten"
        ElseIf Number = 9 Then
            Number = Number - 9
            WordNumber = WordNumber & " " & "Nine"
        ElseIf Number = 8 Then
            Number = Number - 8
            WordNumber = WordNumber & " " & "Eight"
        ElseIf Number = 7 Then
            Number = Number - 7
            WordNumber = WordNumber & " " & "Seven"
        ElseIf Number = 6 Then
            Number = Number - 6
            WordNumber = WordNumber & " " & "Six"
        ElseIf Number = 5 Then
            Number = Number - 5
            WordNumber = WordNumber & " " & "Five"
        ElseIf Number = 4 Then
            Number = Number - 4
            WordNumber = WordNumber & " " & "Four"
        ElseIf Number = 3 Then
            Number = Number - 3
            WordNumber = WordNumber & " " & "Three"
        ElseIf Number = 2 Then
            Number = Number - 2
            WordNumber = WordNumber & " " & "Two"
        ElseIf Number = 1 Then
            Number = Number - 1
            WordNumber = WordNumber & " " & "One"
        Else
            WordNumber = "Zero"
        End If
    Loop While Number > 0
    
    WordNumber = ReplaceLastComma(WordNumber)
    
    
    
End Function
