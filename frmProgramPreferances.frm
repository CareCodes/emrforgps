VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProgramPreferances 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Program Preferances"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
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
   ScaleHeight     =   5040
   ScaleWidth      =   6930
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameBackGroundPicture 
      Caption         =   "Background Picture"
      Height          =   1215
      Left            =   2160
      TabIndex        =   11
      Top             =   120
      Width           =   4695
      Begin btButtonEx.ButtonEx btnSelectBackgroundFile 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Select File"
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
      Begin VB.TextBox txtBackgroundPicture 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame FrameMultipleVisits 
      Caption         =   "Showing Visits"
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1935
      Begin VB.OptionButton optSingleVisit 
         Caption         =   "Single Visit"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optMultipleVisits 
         Caption         =   "Multiple Visits"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame frameColour 
      Caption         =   "Colour"
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
      Begin VB.OptionButton optSunny 
         Caption         =   "Sunny"
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton optEnergy 
         Caption         =   "Energy"
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optAqua 
         Caption         =   "Aqua"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame frmAutoSecession 
      Caption         =   "Auto Secession"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton optAutoSecessionNo 
         Caption         =   "No"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optAutoSecessionYes 
         Caption         =   "Yes"
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmProgramPreferances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
    Unload Me
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

Private Sub btnSelectBackgroundFile_Click()
    On Error Resume Next
    CommonDialog1.Flags = cdlOFNFileMustExist
    CommonDialog1.Flags = cdlOFNNoChangeDir
    CommonDialog1.DefaultExt = "jpg|jpe|jpeg|gif|bmp"
    CommonDialog1.Filter = "JPG|*.jpg|JPE|*.jpe|JPEG|*.jpeg|GIF|*.gif|BMF|*.bmf"
    On Error Resume Next
    CommonDialog1.ShowOpen
    If CommonDialog1.CancelError = False Then
        txtBackgroundPicture.Text = CommonDialog1.filename
        MDIImageFile = CommonDialog1.filename
        SaveSetting App.EXEName, "Options", "MDIImageFile", txtBackgroundPicture.Text
        Unload Me
    Else
        MsgBox "You have not selected valid database. The program may not function", vbCritical, "No database"
    End If
End Sub

Private Sub Form_Load()
    Call GetSettings
    Call SetColours
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveToMemory
    Call SaveSettings
    
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
    
    SaveSetting App.EXEName, "Options", "AutoSecession", optAutoSecessionYes.Value
    SaveSetting App.EXEName, "Options", "Aqua", optAqua.Value
    SaveSetting App.EXEName, "Options", "Energy", optEnergy.Value
    SaveSetting App.EXEName, "Options", "Sunny", optSunny.Value
    SaveSetting App.EXEName, "Options", "SingleVisit", optSingleVisit.Value
    SaveSetting App.EXEName, "Options", "MDIImageFile", txtBackgroundPicture.Text
End Sub

Private Sub GetSettings()
    GetCommonSettings Me
    
    If CBool(GetSetting(App.EXEName, "Options", "AutoSecession", True)) = True Then
        optAutoSecessionNo.Value = False
        optAutoSecessionYes.Value = True
    Else
        optAutoSecessionNo.Value = True
        optAutoSecessionYes.Value = False
    End If
    If CBool(GetSetting(App.EXEName, "Options", "Energy", True)) = True Then
        optEnergy.Value = True
    ElseIf CBool(GetSetting(App.EXEName, "Options", "Sunny", False)) = True Then
        optSunny.Value = True
    ElseIf CBool(GetSetting(App.EXEName, "Options", "Aqua", False)) = True Then
        optAqua.Value = True
    End If
    If GetSetting(App.EXEName, "Options", "SingleVisit", True) = True Then
        optSingleVisit.Value = True
    Else
        optMultipleVisits.Value = True
    End If
    txtBackgroundPicture.Text = GetSetting(App.EXEName, "Options", "MDIImageFile", True)
End Sub

Private Sub SaveToMemory()
    AutoSecession = optAutoSecessionYes.Value
    If optAqua.Value = True Then
        DefaultColourScheme = Aqua
    ElseIf optSunny.Value = True Then
        DefaultColourScheme = Sunny
    ElseIf optEnergy.Value = True Then
        DefaultColourScheme = Energy
    End If
    SingleVisit = optSingleVisit.Value
    MDIImageFile = txtBackgroundPicture.Text
End Sub

Private Sub optAqua_Click()
    Call SaveToMemory
    Call SetColours
End Sub

Private Sub optEnergy_Click()
    Call SaveToMemory
    Call SetColours
End Sub

Private Sub optSunny_Click()
    Call SaveToMemory
    Call SetColours
End Sub
