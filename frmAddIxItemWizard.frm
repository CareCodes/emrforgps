VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAddIxItemWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Report Item"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3540
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
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3540
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameSelect 
      Caption         =   "Select Item Type"
      Height          =   2775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
      Begin VB.OptionButton optCalc 
         Caption         =   "Calculation"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton optCircle 
         Caption         =   "Circle"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton optRectangle 
         Caption         =   "Rectangle"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton optLine 
         Caption         =   "Line"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton optValue 
         Caption         =   "Value"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton optLabel 
         Caption         =   "Label"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
   End
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAddIxItemWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    

Private Sub btnAdd_Click()
    Dim rsReportItem As New ADODB.Recordset
    Dim MyReportDefault As New clsReportDefault
    Dim OrderNo As Long
    
    With rsReportItem
        If .State = 1 Then .Close
        temSQL = "Select * from tblIxItem where Deleted = false AND IxID = " & frmIxFormat.FormIxID
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        OrderNo = .RecordCount
        If .State = 1 Then .Close
        temSQL = "Select * from tblIxItem"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
            !IxID = frmIxFormat.FormIxID
        If optCircle.Value = True Then
            !IxItem = "New Circle"
            !IsCircle = True
            !CircleRadius = 0.2
        ElseIf optLabel.Value = True Then
            !IxItem = "New Label"
            !IsLabel = True
            !ItemText = "New Label"
        ElseIf optLine.Value = True Then
            !IxItem = "New Line"
            !IsLine = True
        ElseIf optRectangle.Value = True Then
            !IxItem = "New Rectangle"
            !IsRectangle = True
        ElseIf optValue.Value = True Then
            !IxItem = "New Value"
            !IsValue = True
            !OrderNo = OrderNo
        ElseIf optCalc.Value = True Then
            !IsCalc = True
            !IxItem = "New Calculation"
        End If
        
        !X1 = 0.4
        !X2 = 0.6
        !Y1 = 0.4
        !Y2 = 0.43
        
        If optLabel.Value = True Then
            !FontName = MyReportDefault.LabelFontName
            !FontSize = MyReportDefault.LabelFontSize
            !FontBold = MyReportDefault.LabelFontBold
            !FontItalic = MyReportDefault.LabelFontItalic
            !FontStrikeThrough = MyReportDefault.LabelFontStrikeThrough
            !FontUnderline = MyReportDefault.LabelFontUnderline
        End If
        
         If optValue.Value = True Or optCalc.Value = True Then
            !FontName = MyReportDefault.TextFontName
            !FontSize = MyReportDefault.TextFontSize
            !FontBold = MyReportDefault.TextFontBold
            !FontItalic = MyReportDefault.TextFontItalic
            !FontStrikeThrough = MyReportDefault.TextFontStrikeThrough
            !FontUnderline = MyReportDefault.TextFontUnderline
        End If
       
       .Update
        Unload Me
    End With
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call SetColours
End Sub


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

