VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPatientMerge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient Merge"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10560
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatientMerge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10560
   Begin VB.TextBox txtPatient2 
      Height          =   360
      Left            =   5160
      TabIndex        =   24
      ToolTipText     =   "The name of the patient. Ex. Niluka Gunasekara"
      Top             =   3240
      Width           =   4935
   End
   Begin VB.TextBox txtAge2 
      Height          =   360
      Left            =   6120
      TabIndex        =   22
      ToolTipText     =   $"frmPatientMerge.frx":058A
      Top             =   4680
      Width           =   1815
   End
   Begin VB.TextBox txtID2 
      Height          =   360
      Left            =   8880
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtSearchId2 
      Height          =   360
      Left            =   3000
      TabIndex        =   17
      ToolTipText     =   "Type the patient ID to select a patient directly"
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtSearchID1 
      Height          =   360
      Left            =   3120
      TabIndex        =   16
      ToolTipText     =   "Type the patient ID to select a patient directly"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtID1 
      Height          =   360
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpDOB1 
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      ToolTipText     =   $"frmPatientMerge.frx":0664
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   93192195
      CurrentDate     =   39845
   End
   Begin VB.TextBox txtAge1 
      Height          =   360
      Left            =   6480
      TabIndex        =   9
      ToolTipText     =   $"frmPatientMerge.frx":073F
      Top             =   1920
      Width           =   1815
   End
   Begin MSDataListLib.DataCombo cmbTitle1 
      Height          =   360
      Left            =   6480
      TabIndex        =   5
      ToolTipText     =   "Select the Title of the patient, ex. Dr.(Mrs.)"
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtPatient1 
      Height          =   360
      Left            =   5520
      TabIndex        =   3
      ToolTipText     =   "The name of the patient. Ex. Niluka Gunasekara"
      Top             =   480
      Width           =   4935
   End
   Begin MSDataListLib.DataCombo cmbPatient1 
      Height          =   2340
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   $"frmPatientMerge.frx":0819
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4128
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSex1 
      Height          =   360
      Left            =   6480
      TabIndex        =   7
      ToolTipText     =   "Select the Sex of the patient. Ex. Female"
      Top             =   1440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   9240
      TabIndex        =   15
      ToolTipText     =   "This will close the Patient window. You will loose all the unsaved data."
      Top             =   6240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "This will save all the above entered details."
      Top             =   6120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Merge"
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
   Begin MSDataListLib.DataCombo cmbPatient2 
      Height          =   2340
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   $"frmPatientMerge.frx":08FE
      Top             =   3720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4128
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpDOB2 
      Height          =   375
      Left            =   6120
      TabIndex        =   21
      ToolTipText     =   $"frmPatientMerge.frx":09E3
      Top             =   5160
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   92864515
      CurrentDate     =   39845
   End
   Begin MSDataListLib.DataCombo cmbTitle2 
      Height          =   360
      Left            =   6120
      TabIndex        =   23
      ToolTipText     =   "Select the Title of the patient, ex. Dr.(Mrs.)"
      Top             =   3720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSex2 
      Height          =   360
      Left            =   6120
      TabIndex        =   25
      ToolTipText     =   "Select the Sex of the patient. Ex. Female"
      Top             =   4200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label13 
      Caption         =   "Name"
      Height          =   255
      Left            =   4320
      TabIndex        =   31
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "Title"
      Height          =   255
      Left            =   4320
      TabIndex        =   30
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Sex"
      Height          =   255
      Left            =   4320
      TabIndex        =   29
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Date of Birth"
      Height          =   255
      Left            =   4320
      TabIndex        =   28
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Age"
      Height          =   255
      Left            =   4320
      TabIndex        =   27
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "ID"
      Height          =   255
      Left            =   8160
      TabIndex        =   26
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Second Patient"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblID 
      Caption         =   "ID"
      Height          =   255
      Left            =   8520
      TabIndex        =   13
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Age"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Date of Birth"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Sex"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Title"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   4680
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "First Patient"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPatientMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'    Dim temSql As String
'    Dim rsPastVisits As New ADODB.Recordset
'
'Private Sub btnClose_Click()
'    Unload Me
'End Sub
'
'Private Sub cmbPatient1_Change()
'    Call ClearValues1
'    CurrentPatient1.ID = Val(cmbPatient1.BoundText)
'    If IsNumeric(cmbPatient1.BoundText) = True Then
'        Call DisplayDetails1
'        txtAge1.Text = CurrentPatient1.AgeInYears
'    End If
'End Sub
'
'Private Sub cmbPatient2_Change()
'    Call clearvalues2
'    CurrentPatient2.ID = Val(cmbPatient2.BoundText)
'    If IsNumeric(cmbPatient2.BoundText) = True Then
'        Call DisplayDetails2
'        txtAge2.Text = CurrentPatient2.AgeInYears
'    End If
'End Sub
'
'Private Sub Form_Load()
'    Call SetColours
'    Call SelectMode
'    Call ClearValues1
'    Call clearvalues2
'    Call FillCombos
'    Call FillOnceCombo
'    Call GetFormSettings
'End Sub
'
'
'Private Sub FillOnceCombo()
'    Dim Title1 As New clsFillCombos
'    Title1.FillAnyCombo cmbTitle1, "Title", True
'    Dim Sex1 As New clsFillCombos
'    Sex1.FillAnyCombo cmbSex1, "Sex", True
'    Dim Doc1 As New clsFillCombos
'    Doc1.FillBoolCombo cmbDoctor1, "staff", "staff", "IsADoctor", True
'    Dim Secession1 As New clsFillCombos
'    Secession1.FillBoolCombo cmbSecession1, "Secession", "Secession", "IsVisitSecession", True
'
'    Dim Title2 As New clsFillCombos
'    Title2.FillAnyCombo cmbTitle2, "Title", True
'    Dim Sex2 As New clsFillCombos
'    Sex2.FillAnyCombo cmbSex2, "Sex", True
'    Dim Doc2 As New clsFillCombos
'    Doc2.FillBoolCombo cmbDoctor2, "staff", "staff", "IsADoctor", True
'    Dim Secession2 As New clsFillCombos
'    Secession2.FillBoolCombo cmbSecession2, "Secession", "Secession", "IsVisitSecession", True
'End Sub
'
'Private Sub GetFormSettings()
'    GetCommonSettings Me
'
'
'
'End Sub
'
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Call SetFormSettings
'End Sub
'
'Private Sub SetFormSettings()
'    SaveSetting App.EXEName, Me.Name, "DoctorID", Val(cmbDoctor.BoundText)
'    SaveSetting App.EXEName, Me.Name, "SecessionID", Val(cmbSecession.BoundText)
'    SaveCommonSettings Me
'
'
'End Sub
'
'Private Sub SetColours()
'    Me.ForeColor = DefaultColourScheme.LabelForeColour
'    Me.BackColor = DefaultColourScheme.LabelBackColour
'    On Error Resume Next
'    Dim MyControl As Control
'    For Each MyControl In Controls
'        If InStr(UCase(MyControl.Name), "BTN") > 0 Then
'            MyControl.ForeColor = DefaultColourScheme.ButtonForeColour
'            MyControl.BackColor = DefaultColourScheme.ButtonBackColour
'            MyControl.BorderColor = DefaultColourScheme.ButtonBorderColour
'        ElseIf InStr(UCase(MyControl.Name), "LST") > 0 Then
'            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
'            MyControl.BackColor = DefaultColourScheme.LabelBackColour
'        ElseIf InStr(UCase(MyControl.Name), "TXTID") > 0 Then
'            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
'            MyControl.BackColor = DefaultColourScheme.LabelBackColour
'        ElseIf InStr(UCase(MyControl.Name), "CMB") > 0 Then
'            MyControl.ForeColor = DefaultColourScheme.ComboForeColour
'            MyControl.BackColor = DefaultColourScheme.ComboBackColour
'        ElseIf InStr(UCase(MyControl.Name), "TXT") > 0 Then
'            MyControl.ForeColor = DefaultColourScheme.TextForeColour
'            MyControl.BackColor = DefaultColourScheme.TextBackColour
'        ElseIf InStr(UCase(MyControl.Name), "DTP") > 0 Then
'            MyControl.ForeColor = DefaultColourScheme.TextForeColour
'            MyControl.BackColor = DefaultColourScheme.TextBackColour
'        Else
'            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
'            MyControl.BackColor = DefaultColourScheme.LabelBackColour
'            MyControl.BackStyle = 0
'        End If
'    Next
'End Sub
'
'Private Sub txtAge_LostFocus()
'    If Abs(DateDiff("yyyy", DateSerial(Year(Date) - Val(txtAge.Text), 1, 1), dtpDOB.Value)) > 2 Then
'        dtpDOB.Value = DateSerial(Year(Date) - Val(txtAge.Text), 1, 1)
'    End If
'End Sub
'
'Private Sub dtpDOB_LostFocus()
'    If Abs(DateDiff("y", DateSerial(Year(Date) - Val(txtAge.Text), 6, 31), dtpDOB.Value)) > 2 Then
'        txtAge.Text = Abs(DateDiff("yyyy", dtpDOB.Value, Date))
'    End If
'
'End Sub
'
'Private Sub EditMode()
'    btnAdd.Enabled = False
'    btnEdit.Enabled = False
'    cmbPatient.Enabled = False
'    txtSearchID.Enabled = False
'
'    cmbDoctor.Enabled = False
'    dtpDate.Enabled = False
'    cmbSecession.Enabled = False
'    txtWeight.Enabled = False
'    txtSBP.Enabled = False
'    txtDBP.Enabled = False
'    btnToQueue.Enabled = False
'
'    txtPatient.Enabled = True
'    txtAge.Enabled = True
'    txtID.Enabled = True
'    cmbTitle.Enabled = True
'    cmbSex.Enabled = True
'    dtpDOB.Enabled = True
'
'    FrameQueue.Enabled = False
'
'    btnSave.Enabled = True
'    btnCancel.Enabled = True
'    btnSaveQueue.Enabled = True
'
'End Sub
'
'Private Sub SelectMode()
'    btnAdd.Enabled = True
'    btnEdit.Enabled = True
'    cmbPatient.Enabled = True
'    txtSearchID.Enabled = True
'
'    cmbDoctor.Enabled = True
'    dtpDate.Enabled = True
'    cmbSecession.Enabled = True
'    txtWeight.Enabled = True
'    txtSBP.Enabled = True
'    txtDBP.Enabled = True
'    btnToQueue.Enabled = True
'
'    txtPatient.Enabled = False
'    txtAge.Enabled = False
'    txtID.Enabled = False
'    cmbTitle.Enabled = False
'    cmbSex.Enabled = False
'    dtpDOB.Enabled = False
'
'    FrameQueue.Enabled = True
'
'    btnSave.Enabled = False
'    btnCancel.Enabled = False
'    btnSaveQueue.Enabled = False
'
'End Sub
'
'Private Sub ClearValues()
'    txtPatient.Text = Empty
'    cmbTitle.Text = Empty
'    cmbSex.Text = Empty
'    lblID.Visible = False
'    txtID.Visible = False
'End Sub
'
'Private Function SaveNew() As Long
'    Dim rsTem As New ADODB.Recordset
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "Select * from tblPatient"
'        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'        .AddNew
'        !Patient = txtPatient.Text
'        !TitleID = Val(cmbTitle.BoundText)
'        !SexID = Val(cmbSex.BoundText)
'        !DateOfBirth = dtpDOB.Value
'        .Update
'        CurrentPatient.ID = !PatientID
'        SaveNew = !PatientID
'        .Close
'    End With
'    Set rsTem = Nothing
'End Function
'
'Private Sub SaveOld()
'    Dim rsTem As New ADODB.Recordset
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "Select * from tblPatient where PatientID = " & Val(cmbPatient.BoundText)
'        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
'        If .RecordCount > 0 Then
'            !Patient = txtPatient.Text
'            !TitleID = Val(cmbTitle.BoundText)
'            !SexID = Val(cmbSex.BoundText)
'            !DateOfBirth = dtpDOB.Value
'            .Update
'        End If
'        CurrentPatient.ID = !PatientID
'        .Close
'    End With
'    Set rsTem = Nothing
'End Sub
'
'Private Sub FillCombos()
'    Dim Pt As New clsFillCombos
'    Pt.FillAnyCombo cmbPatient, "Patient", True
'End Sub
'
'Private Sub DisplayDetails()
''    Dim rsTem As New ADODB.Recordset
''    With rsTem
''        If .State = 1 Then .Close
''        temSql = "Select * from tblPatient where PatientID = " & Val(cmbPatient.BoundText)
''        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
''        If .RecordCount > 0 Then
'            txtPatient.Text = CurrentPatient.Patient '  !Patient
'            cmbTitle.BoundText = CurrentPatient.TitleID '  !TitleID
'            cmbSex.BoundText = CurrentPatient.SexID ' !SexID
'            dtpDOB.Value = CurrentPatient.DateOfBirth '  !DateOfBirth
'            lblID.Visible = True
'            txtID.Visible = True
'            txtID.Text = CurrentPatient.ID '  !PatientID
''        End If
''        .Close
''    End With
''    Set rsTem = Nothing
'End Sub
'
'Private Sub txtSearchID_Change()
'    cmbPatient.BoundText = Val(txtSearchID.Text)
'End Sub
'
'
'
'
'
'
'
'
'Private Sub cmbPatient_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        If Shift = 1 Then
'            btnAdd.SetFocus
'        Else
'            btnEdit.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        cmbPatient.Text = Empty
'    End If
'End Sub
'
'Private Sub cmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        If Shift = 1 Then
'            cmbTitle.SetFocus
'        Else
'            txtAge.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        cmbSex.Text = Empty
'    End If
'End Sub
'
'Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        If Shift = 1 Then
'            btnCancel.SetFocus
'        Else
'            cmbTitle.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        txtPatient.Text = Empty
'    End If
'End Sub
'
'Private Sub cmbTitle_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        If Shift = 1 Then
'            txtPatient.SetFocus
'        Else
'            cmbSex.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        cmbTitle.Text = Empty
'    End If
'End Sub
'
'Private Sub txtAge_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        If Shift = 1 Then
'            cmbSex.SetFocus
'        Else
'            dtpDOB.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        txtAge.Text = Empty
'    End If
'End Sub
'
'Private Sub cmbDoctor_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        If Shift = 1 Then
'            txtPatient.SetFocus
'        Else
'            dtpDate.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        cmbDoctor.Text = Empty
'    End If
'End Sub
'
'Private Sub cmbSecession_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        If Shift = 1 Then
'            dtpDate.SetFocus
'        Else
'            txtWeight.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        cmbSecession.Text = Empty
'    End If
'End Sub
'
'Private Sub txtDBP_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        If Shift = 1 Then
'            txtSBP.SetFocus
'        Else
'            btnToQueue.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        txtDBP.Text = Empty
'    End If
'End Sub
'
'Private Sub txtSBP_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        If Shift = 1 Then
'            txtWeight.SetFocus
'        Else
'            txtDBP.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        txtSBP.Text = Empty
'    End If
'End Sub
'
'Private Sub txtWeight_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        If Shift = 1 Then
'            cmbSecession.SetFocus
'        Else
'            txtSBP.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        txtWeight.Text = Empty
'    End If
'End Sub
'
'Private Sub dtpDOB_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        If Shift = 1 Then
'            txtAge.SetFocus
'        Else
'            btnSave.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        dtpDOB.Value = Date
'    End If
'End Sub
'
'Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        KeyCode = Empty
'        If Shift = 1 Then
'            txtAge.SetFocus
'        Else
'            btnSave.SetFocus
'        End If
'    ElseIf KeyCode = vbKeyEscape Then
'        dtpDOB.Value = Date
'    End If
'End Sub
'
'Private Sub txtSBP_Change()
'    On Error Resume Next
'    If Len(txtSBP.Text) >= 3 Then txtDBP.SetFocus
'    If Len(txtSBP.Text) = 2 And Val(txtSBP.Text) > 30 Then txtDBP.SetFocus
'End Sub
'
'Private Sub txtDBP_Change()
'    On Error Resume Next
'    If Len(Trim(txtDBP.Text)) >= 3 Then btnToQueue.SetFocus
'    If Len(Trim(txtDBP.Text)) = 2 And Val(txtDBP.Text) > 20 Then btnToQueue.SetFocus
'End Sub
