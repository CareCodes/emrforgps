Attribute VB_Name = "modSystem"
Option Explicit
    
    Dim sql As String
    Dim rsTem As New ADODB.Recordset
    
Public Function calculateBmi(height As Double, weight As Double)
    Dim ht As Double
    If height = 0 Or weight = 0 Then
        calculateBmi = 0
        Exit Function
    End If
    If height < 2 Then
        ht = height
    Else
        ht = height / 100
    End If
    calculateBmi = Format(weight / (ht * ht), "0.0")
End Function

