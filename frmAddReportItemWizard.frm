VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAddReportItemWizard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Report Item"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3510
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
   ScaleHeight     =   4155
   ScaleWidth      =   3510
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameSelect 
      Caption         =   "Select Item Type"
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
      Begin VB.OptionButton optCalculation 
         Caption         =   "Calculation"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   2175
      End
      Begin VB.OptionButton optValueField 
         Caption         =   "Value Field"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton optCircle 
         Caption         =   "Circle"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   2175
      End
      Begin VB.OptionButton optRectangle 
         Caption         =   "Rectangle"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   2175
      End
      Begin VB.OptionButton optLine 
         Caption         =   "Line"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   2175
      End
      Begin VB.OptionButton optText 
         Caption         =   "Text Field"
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
      Top             =   3600
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
      Top             =   3600
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
Attribute VB_Name = "frmAddReportItemWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    

Private Sub btnAdd_Click()
    Dim rsReportItem As New ADODB.Recordset
    Dim MyReportDefault As New clsReportDefault
    
    With rsReportItem
        
        If .State = 1 Then .Close
        temSql = "Select * from tblReportItem"
        .Open temSql, cnnLab, adOpenStatic, adLockOptimistic
        .AddNew
        !BackColour = 15724527
        !ForeColour = 986895
        If optCircle.Value = True Then
            !IsCircle = True
            !CircleRadius = 0.2
            !PrintItem = "New Circle"
        ElseIf optLabel.Value = True Then
            !IsLabel = True
            !PrintItem = "New Label"
            !LabelText = "New Label"
        ElseIf optValueField.Value = True Then
            !IsValue = True
            !PrintItem = "New Value"
        ElseIf optLine.Value = True Then
            !IsLine = True
            !PrintItem = "New Line"
        ElseIf optRectangle.Value = True Then
            !IsRectangle = True
            !PrintItem = "New Rectangle"
        ElseIf optText.Value = True Then
            !IsText = True
            !TextText = "Patient Name"
            !PrintItem = "New Text Field"
        ElseIf optCalculation.Value = True Then
            !IsCalculation = True
            !PrintItem = "New Calculation"
        End If
        !IsRelativeX1 = False
        !IsRelativeX2 = False
        !IsRelativeY1 = False
        !IsRelativeY2 = False
        
        !X1 = 0.4
        !X2 = 0.6
        !Y1 = 0.4
        !Y2 = 0.6
        
        If optLabel.Value = True Then
            !FontName = MyReportDefault.LabelFontName
            !FontSize = MyReportDefault.LabelFontSize
            !FontBold = MyReportDefault.LabelFontBold
            !FontItalic = MyReportDefault.LabelFontItalic
            !FONTSTRIKETHROUGH = MyReportDefault.LabelFontStrikeThrough
            !FontUnderline = MyReportDefault.LabelFontUnderline
        End If
        
         If optText.Value = True Then
            !FontName = MyReportDefault.TextFontName
            !FontSize = MyReportDefault.TextFontSize
            !FontBold = MyReportDefault.TextFontBold
            !FontItalic = MyReportDefault.TextFontItalic
            !FONTSTRIKETHROUGH = MyReportDefault.TextFontStrikeThrough
            !FontUnderline = MyReportDefault.TextFontUnderline
        End If
       
       .Update
        Unload Me
    End With
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
