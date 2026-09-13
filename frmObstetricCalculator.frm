VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmObstetricCalculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Obstetric Calculater"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   7770
   Begin VB.TextBox txtPOA 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   3600
      Width           =   2175
   End
   Begin MSComCtl2.MonthView MonthView2 
      Height          =   2820
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   15794177
      CurrentDate     =   39892
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2820
      Left            =   4200
      TabIndex        =   2
      Top             =   600
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   15794177
      CurrentDate     =   39892
   End
   Begin MSComCtl2.MonthView MonthView3 
      Height          =   2820
      Left            =   2040
      TabIndex        =   3
      Top             =   4080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4974
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   15794177
      CurrentDate     =   39892
   End
   Begin VB.Label Label1 
      Caption         =   "LRMP"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "EDD"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "POA"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "frmObstetricCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetColours()
    Me.ForeColor = DefaultColourScheme.LabelForeColour
    Me.BackColor = DefaultColourScheme.LabelBackColour
    On Error Resume Next
    Dim MyControl As Control
    For Each MyControl In Controls
        If InStr(UCase(MyControl.Name), "BTN") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.ButtonForeColour
            MyControl.BackColor = DefaultColourScheme.ButtonBackColour
            MyControl.BorderColor = DefaultColourScheme.ButtonBorderColour
        ElseIf InStr(UCase(MyControl.Name), "LST") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
        ElseIf InStr(UCase(MyControl.Name), "TXTID") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
        ElseIf InStr(UCase(MyControl.Name), "CMB") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.ComboForeColour
            MyControl.BackColor = DefaultColourScheme.ComboBackColour
        ElseIf InStr(UCase(MyControl.Name), "TXT") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.TextForeColour
            MyControl.BackColor = DefaultColourScheme.TextBackColour
        ElseIf InStr(UCase(MyControl.Name), "DTP") > 0 Then
            MyControl.ForeColor = DefaultColourScheme.TextForeColour
            MyControl.BackColor = DefaultColourScheme.TextBackColour
        Else
            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
            MyControl.BackColor = DefaultColourScheme.LabelBackColour
            MyControl.BackStyle = 0
        End If
    Next
End Sub

Private Sub Form_Load()
    GetCommonSettings Me
    
    Call SetColours
    MonthView1.Value = Date
    MonthView3.Value = Date
    MonthView2.Value = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub

Private Sub MonthView2_DateClick(ByVal DateClicked As Date)
    MonthView1.Value = EDD(MonthView2.Value)
    txtPOA.Text = POA(MonthView2.Value, MonthView3.Value)
End Sub

Private Function EDD(LRMP As Date, Optional ExactDate As Boolean) As Date
'    EDD = LRMP + 280
    EDD = DateSerial(Year(LRMP), Month(LRMP) + 9, Day(LRMP) + 7)
End Function


Private Function POA(LRMP As Date, CalDate As Date) As String
    Dim POADates As Long
    POADates = DateDiff("d", LRMP, CalDate)
    POA = POADates \ 7 & " weeks"
    If POADates Mod 7 = 0 Then
    
    Else
        POA = POA & " and " & POADates Mod 7 & " days"
    End If
End Function

Private Sub MonthView3_DateClick(ByVal DateClicked As Date)
    txtPOA.Text = POA(MonthView2.Value, MonthView3.Value)
End Sub

