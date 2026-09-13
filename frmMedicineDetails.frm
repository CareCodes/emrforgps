VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMedicineDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Daily Income"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12675
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
   ScaleHeight     =   8040
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   4680
      TabIndex        =   6
      Top             =   360
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Generic Name Details"
      TabPicture(0)   =   "frmMedicineDetails.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "rtbGenericDetails"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Trade Name Details"
      TabPicture(1)   =   "frmMedicineDetails.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "rtbTradeNameDetails"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin RichTextLib.RichTextBox rtbGenericDetails 
         Height          =   6375
         Left            =   -74880
         TabIndex        =   7
         Top             =   480
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   11245
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmMedicineDetails.frx":0038
      End
      Begin RichTextLib.RichTextBox rtbTradeNameDetails 
         Height          =   6375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   11245
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmMedicineDetails.frx":00BC
      End
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   11280
      TabIndex        =   3
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&lose"
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
   Begin MSFlexGridLib.MSFlexGrid GridTrade 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   11033
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbGeneric 
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   9960
      TabIndex        =   4
      Top             =   7440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Print"
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
   Begin VB.Label Label2 
      Caption         =   "Trade Name"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Generic Name"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmMedicineDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub GetSettings()
    GetCommonSettings Me
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
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
    Call SetColours
    Call FillCombos
    Call GetSettings
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub FillCombos()
    Dim Generic As New clsFillCombos
    Generic.FillBoolCombo cmbGeneric, "Item", "Item", "IsGenericName", True
End Sub

Private Sub FormatGrid()
    With GridTrade
        .Clear
        .Cols = 2
        .Rows = 1
        
        .Row = 0
        
        .Col = 0
        .Text = "Trade Name"
        
        .Col = 1
        .Text = "Category"
    End With
End Sub

Private Sub FillGrid()
    If IsNumeric(cmbGeneric.BoundText) = False Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            GridTrade.Rows = GridTrade.Rows + 1
            GridTrade.Row = GridTrade.Rows - 1

            GridTrade.Col = 0
            GridTrade.Text = !VisitID

            If GridTrade.Row Mod 2 = 0 Then
                GridTrade.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridTrade.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If

            GridTrade.Col = 1
            If IsNull(!VisitTime) = False Then
                GridTrade.Text = Format(!VisitTime, "HH:MM AMPM")
            Else
                GridTrade.Text = "Not Seen"
            End If
            If GridTrade.Row Mod 2 = 0 Then
                GridTrade.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridTrade.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If

            .MoveNext
        Wend
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub


