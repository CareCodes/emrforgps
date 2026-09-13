VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmSearchPatient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Patient"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9285
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
   ScaleHeight     =   8640
   ScaleWidth      =   9285
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   7920
      TabIndex        =   8
      Top             =   8040
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
   Begin VB.Frame Frame1 
      Caption         =   "Sex"
      Height          =   1575
      Left            =   4680
      TabIndex        =   3
      Top             =   0
      Width           =   1575
      Begin VB.OptionButton optAll 
         Caption         =   "&All"
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optFemale 
         Caption         =   "&Female"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optMale 
         Caption         =   "&Male"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid gridPatient 
      Height          =   6255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11033
      _Version        =   393216
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   2566
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Name"
      TabPicture(0)   =   "frmSearchPatient.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Address"
      TabPicture(1)   =   "frmSearchPatient.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtByAddress"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Birthday"
      TabPicture(2)   =   "frmSearchPatient.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "dtpDOB"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.TextBox txtByAddress 
         Height          =   360
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   4095
      End
      Begin VB.TextBox txtName 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtpDOB 
         Height          =   375
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMMM yyyy"
         Format          =   163774467
         CurrentDate     =   39984
      End
   End
End
Attribute VB_Name = "frmSearchPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim SearchCount As Long
    Public FormName As String
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub dtpDOB_Change()
    Call PatientSearch
    SearchCount = SearchCount + 1
End Sub

Private Sub Form_Activate()
    txtName.Text = Empty
End Sub

Private Sub Form_Load()
    Call SetColours
    Call FormatGrid
    Call GetSettings
End Sub

Private Sub GetSettings()
    SSTab1.Tab = Val(GetSetting(App.EXEName, Me.Name, SSTab1.Name, 0))
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


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, SSTab1.Name, SSTab1.Tab
End Sub

Private Sub gridPatient_DblClick()
    Dim temPatientID As Long
    temPatientID = Val(gridPatient.TextMatrix(gridPatient.row, 0))
    If frmPatientDetails.Name = FormName Then
        frmPatientDetails.Show
        frmPatientDetails.cmbPatient.BoundText = temPatientID
        frmPatientDetails.ZOrder 0
    ElseIf frmPatient.Name = FormName Then
        frmPatient.Show
        frmPatient.cmbPatient.BoundText = temPatientID
        frmPatient.ZOrder 0
    Else
        frmPatientDetails.Show
        frmPatientDetails.cmbPatient.BoundText = temPatientID
        frmPatientDetails.ZOrder 0
    End If
End Sub

Private Sub optAll_Click()
    Call PatientSearch
End Sub

Private Sub optFemale_Click()
    Call PatientSearch
End Sub

Private Sub optMale_Click()
    Call PatientSearch
End Sub

Private Sub txtByAddress_Change()
    Call PatientSearch
    SearchCount = SearchCount + 1
End Sub

Private Sub txtName_Change()
    SearchCount = SearchCount + 1
    Call PatientSearch
End Sub

Private Sub PatientSearch()
    Select Case SSTab1.Tab
        Case 0: NameSearch
        Case 1: AddressSearch
        Case 2: DOBSearch
    End Select
End Sub

Private Sub AddressSearch()
    On Error Resume Next
    Dim OldSearchCount As Long
    Call FormatGrid
    OldSearchCount = SearchCount
    If Trim(txtByAddress.Text) = "" Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    temSql = "Select * from tblPatient where "
    If optMale.Value = True Then
        temSql = temSql & " SexID = 1 AND "
    ElseIf optFemale.Value = True Then
        temSql = temSql & " SexID = 2 AND "
    Else
        
    End If
    temSql = temSql & " Address like '%" & Trim(txtByAddress.Text) & "%' order by Patient"
    With rsTem
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        gridPatient.Visible = False
        If .RecordCount > 0 Then
            While .EOF = False
                DoEvents
                If SearchCount <> OldSearchCount Then
                    gridPatient.Visible = True
                    Exit Sub
                End If
                gridPatient.Rows = gridPatient.Rows + 1
                gridPatient.row = gridPatient.Rows - 1
                gridPatient.col = 0
                gridPatient.Text = !PatientID
                gridPatient.col = 1
                gridPatient.Text = !Patient
                gridPatient.col = 2
                gridPatient.Text = CalculateAgeInWords(!DateOfBirth)
                .MoveNext
            Wend
        End If
        gridPatient.Visible = True
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub DOBSearch()
    On Error Resume Next

    Dim OldSearchCount As Long
    Call FormatGrid
    OldSearchCount = SearchCount
    Dim rsTem As New ADODB.Recordset
    temSql = "Select * from tblPatient where "
    If optMale.Value = True Then
        temSql = temSql & " SexID = 1 AND "
    ElseIf optFemale.Value = True Then
        temSql = temSql & " SexID = 2 AND "
    Else
        
    End If
    temSql = temSql & " DateOfBirth = #" & Format(dtpDOB.Value, "dd MMMM yyyy") & "# order by Patient"
    With rsTem
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        gridPatient.Visible = False
        If .RecordCount > 0 Then
            While .EOF = False
                DoEvents
                If SearchCount <> OldSearchCount Then
                    gridPatient.Visible = True
                    Exit Sub
                End If
                gridPatient.Rows = gridPatient.Rows + 1
                gridPatient.row = gridPatient.Rows - 1
                gridPatient.col = 0
                
                If gridPatient.row Mod 2 = 0 Then
                    gridPatient.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    gridPatient.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
                
                gridPatient.Text = !PatientID
                gridPatient.col = 1
                
                If gridPatient.row Mod 2 = 0 Then
                    gridPatient.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    gridPatient.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
                
                gridPatient.Text = !Patient
                gridPatient.col = 2
                
                If gridPatient.row Mod 2 = 0 Then
                    gridPatient.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    gridPatient.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
                
                gridPatient.Text = CalculateAgeInWords(!DateOfBirth)
                .MoveNext
            Wend
        End If
        gridPatient.Visible = True
        .Close
    End With
    Set rsTem = Nothing
End Sub


Private Sub NameSearch()
    On Error Resume Next

    Dim OldSearchCount As Long
    Call FormatGrid
    OldSearchCount = SearchCount
    If Trim(txtName.Text) = "" Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    temSql = "Select * from tblPatient where "
    If optMale.Value = True Then
        temSql = temSql & " SexID = 1 AND "
    ElseIf optFemale.Value = True Then
        temSql = temSql & " SexID = 2 AND "
    Else
        
    End If
    temSql = temSql & " Patient like '%" & Trim(txtName.Text) & "%' order by Patient"
    With rsTem
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        gridPatient.Visible = False
        If .RecordCount > 0 Then
            While .EOF = False
                DoEvents
                If SearchCount <> OldSearchCount Then
                    gridPatient.Visible = True
                    Exit Sub
                End If
                gridPatient.Rows = gridPatient.Rows + 1
                gridPatient.row = gridPatient.Rows - 1
                gridPatient.col = 0
                
                If gridPatient.row Mod 2 = 0 Then
                    gridPatient.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    gridPatient.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
                
                
                gridPatient.Text = !PatientID
                gridPatient.col = 1
                
                If gridPatient.row Mod 2 = 0 Then
                    gridPatient.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    gridPatient.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
                
                gridPatient.Text = !Patient
                gridPatient.col = 2
                
                If gridPatient.row Mod 2 = 0 Then
                    gridPatient.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    gridPatient.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
                
                gridPatient.Text = CalculateAgeInWords(!DateOfBirth)
                .MoveNext
            Wend
        End If
        gridPatient.Visible = True
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FormatGrid()
    With gridPatient
        .Clear
        .Rows = 1
        .Cols = 3
        .col = 0
        .Text = "Patient ID"
        .col = 1
        .Text = "Name"
        .col = 2
        .Text = "Age"
        .ColWidth(0) = 1200
        .ColWidth(2) = 1800
        .ColWidth(1) = .Width - .ColWidth(0) - .ColWidth(2) - 275
    End With
End Sub
