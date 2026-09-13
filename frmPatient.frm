VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPatient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient"
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
   Icon            =   "frmPatient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10560
   Begin VB.Frame FrameQueue 
      Caption         =   "Queue details"
      Height          =   2295
      Left            =   5160
      TabIndex        =   19
      Top             =   3600
      Width           =   4815
      Begin VB.TextBox txtWeight 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   960
         TabIndex        =   26
         ToolTipText     =   "Enter the weight of the patient in kilogrames, not essential"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtSBP 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2760
         TabIndex        =   28
         ToolTipText     =   "Enter Systolic Blood Pressure hear"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtDBP 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   3360
         TabIndex        =   29
         ToolTipText     =   "Enter the diastolic blood pressure here"
         Top             =   1320
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   375
         Left            =   960
         TabIndex        =   23
         ToolTipText     =   "Select the date the consultation is done."
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd MMM yyyy"
         Format          =   100728835
         CurrentDate     =   39859
      End
      Begin MSDataListLib.DataCombo cmbSecession 
         Height          =   360
         Left            =   2760
         TabIndex        =   24
         ToolTipText     =   "Select the secession"
         Top             =   840
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbDoctor 
         Height          =   360
         Left            =   960
         TabIndex        =   21
         ToolTipText     =   "Select the doctor who will consult the patient."
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin btButtonEx.ButtonEx btnToQueue 
         Height          =   375
         Left            =   1320
         TabIndex        =   30
         ToolTipText     =   "Press this button to add the selected patient to the queue"
         Top             =   1800
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Appearance      =   3
         Caption         =   "Add To Queue"
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
      Begin VB.Label Label17 
         Caption         =   "Doctor"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label18 
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "kg"
         Height          =   255
         Left            =   2040
         TabIndex        =   33
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "BP                 mmHg"
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label21 
         Caption         =   "Weight"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.TextBox txtSearchID 
      Height          =   360
      Left            =   1560
      TabIndex        =   32
      ToolTipText     =   "Type the patient ID to select a patient directly"
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtID 
      Height          =   360
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtpDOB 
      Height          =   375
      Left            =   6480
      TabIndex        =   13
      ToolTipText     =   $"frmPatient.frx":058A
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMM yyyy"
      Format          =   100728835
      CurrentDate     =   39845
   End
   Begin VB.TextBox txtAge 
      Height          =   360
      Left            =   6480
      TabIndex        =   11
      ToolTipText     =   $"frmPatient.frx":0665
      Top             =   1920
      Width           =   1815
   End
   Begin MSDataListLib.DataCombo cmbTitle 
      Height          =   360
      Left            =   6480
      TabIndex        =   7
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
   Begin VB.TextBox txtPatient 
      Height          =   360
      Left            =   5520
      TabIndex        =   5
      ToolTipText     =   "The name of the patient. Ex. Niluka Gunasekara"
      Top             =   480
      Width           =   4935
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Click this button if you are adding a new patient, whose details has never entered before."
      Top             =   5880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Add"
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
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      ToolTipText     =   $"frmPatient.frx":073F
      Top             =   5880
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Edit"
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
   Begin MSDataListLib.DataCombo cmbPatient 
      Height          =   5220
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   $"frmPatient.frx":07E0
      Top             =   600
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   9208
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSex 
      Height          =   360
      Left            =   6480
      TabIndex        =   9
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
      TabIndex        =   31
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
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   375
      Left            =   8760
      TabIndex        =   16
      ToolTipText     =   "The details you entered above will all be lost if you press this button."
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Cancel"
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
      Left            =   5160
      TabIndex        =   14
      ToolTipText     =   "This will save all the above entered details."
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Save"
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
   Begin btButtonEx.ButtonEx btnSaveQueue 
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      ToolTipText     =   "Above entered details will be saved and the patient will be added to the queue."
      Top             =   3000
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Save && Queue"
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
   Begin btButtonEx.ButtonEx btnSearch 
      Height          =   375
      Left            =   2760
      TabIndex        =   34
      ToolTipText     =   "Click here to search the patients from there Name, Address, Date of Birth, etc."
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Sea&rch"
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
   Begin VB.Label lblID 
      Caption         =   "ID"
      Height          =   255
      Left            =   8520
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Age"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Date of Birth"
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Sex"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Title"
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "&Patient"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsPastVisits As New ADODB.Recordset
    
Private Sub btnAdd_Click()
    Dim temText As String
    If IsNumeric(cmbPatient.BoundText) = False Then
        temText = cmbPatient.Text
    Else
        temText = Empty
    End If
    cmbPatient.Text = Empty
    Call EditMode
    txtPatient.Text = temText
    txtPatient.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbPatient.Text = Empty
    cmbPatient.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbPatient.BoundText) = False Then Exit Sub
    Call EditMode
    txtPatient.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    Dim temPatientID As Long
    If Trim(txtPatient.Text) = Empty Then
        MsgBox "You have not entered a name"
        txtPatient.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPatient.BoundText) = True Then
        temPatientID = cmbPatient.BoundText
        Call SaveOld
    Else
        temPatientID = SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbPatient.Text = Empty
    cmbPatient.SetFocus
    cmbPatient.BoundText = temPatientID
End Sub

Private Sub btnSaveQueue_Click()
    Dim temPatientID As Long
    If Trim(txtPatient.Text) = Empty Then
        MsgBox "You have not entered a name"
        txtPatient.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPatient.BoundText) = True Then
        Call SaveOld
        temPatientID = cmbPatient.BoundText
    Else
        temPatientID = SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbPatient.Text = Empty
    cmbPatient.BoundText = temPatientID
    Dim rsQueue As New ADODB.Recordset
    If IsNumeric(cmbDoctor.BoundText) = False Then
        MsgBox "Please select a doctor"
        cmbDoctor.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbSecession.BoundText) = False Then
        MsgBox "Please select a secession"
        cmbSecession.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPatient.BoundText) = False Then
        MsgBox "Please select a patient"
        cmbPatient.SetFocus
        Exit Sub
    End If
    
    Dim i As Integer
    i = MsgBox("Patient" & vbTab & ":" & vbTab & CurrentPatient.NameWithTitle & vbNewLine & "Age   " & vbTab & ":" & vbTab & CurrentPatient.AgeInWords & vbNewLine & "Date    " & vbTab & ":" & vbTab & Format(dtpDate.Value, "dd MMMM yyyy") & vbNewLine & "Secession" & vbTab & ":" & vbTab & cmbSecession.Text & vbNewLine & "Number" & vbTab & ":" & vbTab & GetLastQueueNo(Val(cmbDoctor.BoundText), dtpDate.Value, Val(cmbSecession.BoundText)) & vbNewLine & vbNewLine & "Add to Queue?", vbYesNo)
    If i = vbYes Then
        With rsQueue
            temSql = "Select * from tblQueue"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !PatientID = Val(cmbPatient.BoundText)
            !SecessionID = Val(cmbSecession.BoundText)
            !QueueDate = dtpDate.Value
            !DoctorID = Val(cmbDoctor.BoundText)
            !QueueNo = GetLastQueueNo(Val(cmbDoctor.BoundText), dtpDate.Value, Val(cmbSecession.BoundText))
            !BookedUserID = UserID
            !BookedAt = Now
            !Weight = Val(txtWeight.Text)
            !SBP = Val(txtSBP.Text)
            !DBP = Val(txtDBP.Text)
            !Age = CurrentPatient.AgeInWords
            .Update
            MsgBox "Added to Queue. Number is " & !QueueNo
            .Close
        End With
    End If
    
    
End Sub

Private Sub btnSearch_Click()
    frmSearchPatient.Show
    frmSearchPatient.FormName = Me.Name
    frmSearchPatient.ZOrder 0
End Sub

Private Sub btnToQueue_Click()
    Dim rsQueue As New ADODB.Recordset
    If IsNumeric(cmbDoctor.BoundText) = False Then
        MsgBox "Please select a doctor"
        cmbDoctor.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbSecession.BoundText) = False Then
        MsgBox "Please select a secession"
        cmbSecession.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPatient.BoundText) = False Then
        MsgBox "Please select a patient"
        cmbPatient.SetFocus
        Exit Sub
    End If
    Dim i As Integer
    i = MsgBox("Patient" & vbTab & ":" & vbTab & CurrentPatient.NameWithTitle & vbNewLine & "Age   " & vbTab & ":" & vbTab & CurrentPatient.AgeInWords & vbNewLine & vbNewLine & "Doctor  " & vbTab & ":" & vbTab & cmbDoctor.Text & vbNewLine & "Date    " & vbTab & ":" & vbTab & Format(dtpDate.Value, "dd MMMM yyyy") & vbNewLine & "Secession" & vbTab & ":" & vbTab & cmbSecession.Text & vbNewLine & "Number" & vbTab & ":" & vbTab & GetLastQueueNo(Val(cmbDoctor.BoundText), dtpDate.Value, Val(cmbSecession.BoundText)) & vbNewLine & vbNewLine & "Add to Queue?", vbYesNo)
    If i = vbNo Then Exit Sub
    
    
    With rsQueue
        temSql = "Select * from tblQueue"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !PatientID = Val(cmbPatient.BoundText)
        !SecessionID = Val(cmbSecession.BoundText)
        !QueueDate = dtpDate.Value
        !DoctorID = Val(cmbDoctor.BoundText)
        !QueueNo = GetLastQueueNo(Val(cmbDoctor.BoundText), dtpDate.Value, Val(cmbSecession.BoundText))
        !BookedUserID = UserID
        !BookedAt = Now
        !Weight = Val(txtWeight.Text)
        !SBP = Val(txtSBP.Text)
        !DBP = Val(txtDBP.Text)
        !Age = CurrentPatient.AgeInWords
        .Update
'        MsgBox "Added to Queue. Number is " & !QueueNo
        .Close
    End With
    
    txtWeight.Text = Empty
    txtSBP.Text = Empty
    txtDBP.Text = Empty
    
End Sub


Private Sub cmbPatient_Change()
    Call ClearValues
    CurrentPatient.ID = Val(cmbPatient.BoundText)
    If IsNumeric(cmbPatient.BoundText) = True Then
        Call DisplayDetails
        txtAge.Text = CurrentPatient.AgeInYears
    End If
End Sub

Private Sub Form_Load()
    Call SetColours
    Call SelectMode
    Call ClearValues
    Call FillCombos
    Call FillOnceCombo
    Call GetFormSettings
    dtpDate.Value = Date
     If AutoSecession = True Then Call SetAutoSecession
End Sub

Private Sub SetAutoSecession()
    
    If Time > #2:00:00 PM# Then
        cmbSecession.Text = "Evening"
    Else
        cmbSecession.Text = "Morning"
    End If
End Sub


Private Sub FillOnceCombo()
    Dim Title As New clsFillCombos
    Title.FillAnyCombo cmbTitle, "Title", True
    Dim Sex As New clsFillCombos
    Sex.FillAnyCombo cmbSex, "Sex", True
    Dim Doc As New clsFillCombos
    Doc.FillBoolCombo cmbDoctor, "staff", "staff", "IsADoctor", True
    Dim Secession As New clsFillCombos
    Secession.FillBoolCombo cmbSecession, "Secession", "Secession", "IsVisitSecession", True
End Sub

Private Sub GetFormSettings()
    GetCommonSettings Me
    
    
    cmbDoctor.BoundText = Val(GetSetting(App.EXEName, Me.Name, "DoctorID", 0))
    cmbSecession.BoundText = Val(GetSetting(App.EXEName, Me.Name, "SecessionID", 0))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetFormSettings
End Sub

Private Sub SetFormSettings()
    SaveSetting App.EXEName, Me.Name, "DoctorID", Val(cmbDoctor.BoundText)
    SaveSetting App.EXEName, Me.Name, "SecessionID", Val(cmbSecession.BoundText)
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

Private Sub txtAge_LostFocus()
    If Abs(DateDiff("yyyy", DateSerial(Year(Date) - Val(txtAge.Text), 1, 1), dtpDOB.Value)) > 2 Then
        dtpDOB.Value = DateSerial(Year(Date) - Val(txtAge.Text), 1, 1)
    End If
End Sub

Private Sub dtpDOB_LostFocus()
    If Abs(DateDiff("y", DateSerial(Year(Date) - Val(txtAge.Text), 6, 31), dtpDOB.Value)) > 2 Then
        txtAge.Text = Abs(DateDiff("yyyy", dtpDOB.Value, Date))
    End If

End Sub

Private Sub EditMode()
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    cmbPatient.Enabled = False
    txtSearchID.Enabled = False
    
    cmbDoctor.Enabled = False
    dtpDate.Enabled = False
    cmbSecession.Enabled = False
    txtWeight.Enabled = False
    txtSBP.Enabled = False
    txtDBP.Enabled = False
    btnToQueue.Enabled = False
    
    txtPatient.Enabled = True
    txtAge.Enabled = True
    txtID.Enabled = True
    cmbTitle.Enabled = True
    cmbSex.Enabled = True
    dtpDOB.Enabled = True
    
    FrameQueue.Enabled = False
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    btnSaveQueue.Enabled = True
    
End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    cmbPatient.Enabled = True
    txtSearchID.Enabled = True
    
    cmbDoctor.Enabled = True
    dtpDate.Enabled = True
    cmbSecession.Enabled = True
    txtWeight.Enabled = True
    txtSBP.Enabled = True
    txtDBP.Enabled = True
    btnToQueue.Enabled = True
    
    txtPatient.Enabled = False
    txtAge.Enabled = False
    txtID.Enabled = False
    cmbTitle.Enabled = False
    cmbSex.Enabled = False
    dtpDOB.Enabled = False
    
    FrameQueue.Enabled = True
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
    btnSaveQueue.Enabled = False
    
End Sub

Private Sub ClearValues()
    txtPatient.Text = Empty
    cmbTitle.Text = Empty
    cmbSex.Text = Empty
    lblID.Visible = False
    txtID.Visible = False
End Sub

Private Function SaveNew() As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPatient"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Patient = txtPatient.Text
        !TitleID = Val(cmbTitle.BoundText)
        !SexID = Val(cmbSex.BoundText)
        !DateOfBirth = dtpDOB.Value
        .Update
        CurrentPatient.ID = !PatientID
        SaveNew = !PatientID
        .Close
    End With
    Set rsTem = Nothing
End Function

Private Sub SaveOld()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPatient where PatientID = " & Val(cmbPatient.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Patient = txtPatient.Text
            !TitleID = Val(cmbTitle.BoundText)
            !SexID = Val(cmbSex.BoundText)
            !DateOfBirth = dtpDOB.Value
            .Update
        End If
        CurrentPatient.ID = !PatientID
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim Pt As New clsFillCombos
    Pt.FillAnyCombo cmbPatient, "Patient", True
End Sub

Private Sub DisplayDetails()
'    Dim rsTem As New ADODB.Recordset
'    With rsTem
'        If .State = 1 Then .Close
'        temSql = "Select * from tblPatient where PatientID = " & Val(cmbPatient.BoundText)
'        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
'        If .RecordCount > 0 Then
            txtPatient.Text = CurrentPatient.Patient '  !Patient
            cmbTitle.BoundText = CurrentPatient.TitleID '  !TitleID
            cmbSex.BoundText = CurrentPatient.SexID ' !SexID
            dtpDOB.Value = CurrentPatient.DateOfBirth '  !DateOfBirth
            lblID.Visible = True
            txtID.Visible = True
            txtID.Text = CurrentPatient.ID '  !PatientID
'        End If
'        .Close
'    End With
'    Set rsTem = Nothing
End Sub

Private Sub txtSearchID_Change()
    cmbPatient.BoundText = Val(txtSearchID.Text)
End Sub








Private Sub cmbPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            btnAdd.SetFocus
        Else
            btnEdit.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbPatient.Text = Empty
    End If
End Sub

Private Sub cmbSex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            cmbTitle.SetFocus
        Else
            txtAge.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbSex.Text = Empty
    End If
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            btnCancel.SetFocus
        Else
            cmbTitle.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtPatient.Text = Empty
    End If
End Sub

Private Sub cmbTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtPatient.SetFocus
        Else
            cmbSex.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbTitle.Text = Empty
    End If
End Sub

Private Sub txtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            cmbSex.SetFocus
        Else
            dtpDOB.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtAge.Text = Empty
    End If
End Sub

Private Sub cmbDoctor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtPatient.SetFocus
        Else
            dtpDate.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbDoctor.Text = Empty
    End If
End Sub

Private Sub cmbSecession_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            dtpDate.SetFocus
        Else
            txtWeight.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbSecession.Text = Empty
    End If
End Sub

Private Sub txtDBP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtSBP.SetFocus
        Else
            btnToQueue.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtDBP.Text = Empty
    End If
End Sub

Private Sub txtSBP_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtWeight.SetFocus
        Else
            txtDBP.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtSBP.Text = Empty
    End If
End Sub

Private Sub txtWeight_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            cmbSecession.SetFocus
        Else
            txtSBP.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        txtWeight.Text = Empty
    End If
End Sub

Private Sub dtpDOB_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtAge.SetFocus
        Else
            btnSave.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        dtpDOB.Value = Date
    End If
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If Shift = 1 Then
            txtAge.SetFocus
        Else
            btnSave.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        dtpDOB.Value = Date
    End If
End Sub

Private Sub txtSBP_Change()
    On Error Resume Next
    If Len(txtSBP.Text) >= 3 Then txtDBP.SetFocus
    If Len(txtSBP.Text) = 2 And Val(txtSBP.Text) > 30 Then txtDBP.SetFocus
End Sub

Private Sub txtDBP_Change()
    On Error Resume Next
    If Len(Trim(txtDBP.Text)) >= 3 Then btnToQueue.SetFocus
    If Len(Trim(txtDBP.Text)) = 2 And Val(txtDBP.Text) > 20 Then btnToQueue.SetFocus
End Sub
