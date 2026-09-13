VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmQueue 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Queue"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQueue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   7455
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   6240
      Top             =   6000
   End
   Begin VB.OptionButton optToComplete 
      BackColor       =   &H80000001&
      Caption         =   "To Complete"
      Height          =   240
      Left            =   240
      TabIndex        =   18
      Top             =   1080
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.OptionButton optCompleted 
      BackColor       =   &H80000001&
      Caption         =   "Completed"
      Height          =   240
      Left            =   2640
      Picture         =   "frmQueue.frx":08CA
      TabIndex        =   17
      Top             =   1080
      Width           =   2415
   End
   Begin VB.OptionButton optRemoved 
      BackColor       =   &H80000001&
      Caption         =   "Removed"
      Height          =   240
      Left            =   5040
      TabIndex        =   16
      Top             =   1080
      Width           =   2175
   End
   Begin VB.OptionButton optToIssue 
      BackColor       =   &H80000001&
      Caption         =   "To Issue Medicines"
      Height          =   240
      Left            =   240
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.OptionButton optIssued 
      BackColor       =   &H80000001&
      Caption         =   "Medicines Issued"
      Height          =   240
      Left            =   2640
      TabIndex        =   14
      Top             =   1440
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.OptionButton optAll 
      BackColor       =   &H80000001&
      Caption         =   "All"
      Height          =   240
      Left            =   5040
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6240
      Top             =   5520
   End
   Begin btButtonEx.ButtonEx btnVisit 
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16777215
      Caption         =   "Visit"
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
   Begin MSDataListLib.DataList lstQueue 
      Height          =   6060
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   10689
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   183173123
      CurrentDate     =   39859
   End
   Begin MSDataListLib.DataCombo cmbSecession 
      Height          =   360
      Left            =   4800
      TabIndex        =   3
      Top             =   600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbDoctor 
      Height          =   360
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnPatient 
      Height          =   375
      Left            =   5880
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Patient"
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
   Begin btButtonEx.ButtonEx btnRemove 
      Height          =   375
      Left            =   5880
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Remove"
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   7800
      Width           =   1335
      _ExtentX        =   2355
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
   Begin btButtonEx.ButtonEx btnIncome 
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   4080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Summery"
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
   Begin btButtonEx.ButtonEx btnAddAgain 
      Height          =   375
      Left            =   5880
      TabIndex        =   12
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Add Again"
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Session"
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor"
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim ColourChange As Boolean
    Dim r As Long
    Dim G As Long
    Dim b As Long

Private Sub ChangeColour()
    
End Sub

Private Sub btnAddAgain_Click()
    Dim i As Integer
    If IsNumeric(lstQueue.BoundText) = False Then
        MsgBox "Please select a patient"
        lstQueue.SetFocus
        Exit Sub
    End If
    
    Dim temLong As Integer
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSql = "Select * from tblQueue where QueueID = " & Val(lstQueue.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If !Removed = False Then
                MsgBox "This visit is not a removed visit"
                .Close
                Exit Sub
            End If
            If !Completed = True Then
                MsgBox "This visit is completed. Therefore you can't remove it or add it back"
                .Close
                Exit Sub
            End If
            temLong = MsgBox("Are you sure you want to Add this patient back to the queue?", vbYesNo)
            If temLong = vbYes Then
                !Removed = False
                !RemovedTime = Null
                !RemovedUserID = Null
                .Update
            End If
        End If
        .Close
    End With
    Call FillQueue

End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnIncome_Click()
    Unload frmQueueIncome
    frmQueueIncome.Show
    frmQueueIncome.ZOrder 0
    frmQueueIncome.dtpDate.Value = dtpDate.Value
    frmQueueIncome.cmbDoctor.BoundText = Val(cmbDoctor.BoundText)
    frmQueueIncome.cmbSecession.BoundText = Val(cmbSecession.BoundText)
End Sub

Private Sub btnPatient_Click()
    Dim i As Integer
    If IsNumeric(lstQueue.BoundText) = False Then
        MsgBox "Please select a patient"
        lstQueue.SetFocus
        Exit Sub
    End If
    
    If SingleVisit = True Then
        Unload frmNewVisit
        Unload frmPatientDetails
        Unload frmPastHistory
    End If
    
    Dim ThisPatientID As Long
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSql = "Select * from tblQueue where QueueID = " & Val(lstQueue.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            ThisPatientID = !PatientID
        End If
        .Close
    End With
    
    frmPatientDetails.Show
    frmPatientDetails.ZOrder 0
    frmPatientDetails.cmbPatient.BoundText = ThisPatientID
End Sub

Private Sub btnRemove_Click()
    Dim i As Integer
    If IsNumeric(lstQueue.BoundText) = False Then
        MsgBox "Please select a patient"
        lstQueue.SetFocus
        Exit Sub
    End If
    
    Dim temLong As Integer
    temLong = MsgBox("Are you sure you want to remove this patient?", vbYesNo)
    If temLong = vbNo Then Exit Sub
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSql = "Select * from tblQueue where QueueID = " & Val(lstQueue.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If !Completed = True Then
                MsgBox "This visit is cimpleted. Therefore you can't remove it"
                .Close
                Exit Sub
            End If
            If !Removed = True Then
                MsgBox "This visit is already removed"
                .Close
                Exit Sub
            End If
            !Removed = True
            !RemovedTime = Now
            !RemovedUserID = UserID
            .Update
        End If
        .Close
    End With
    Call FillQueue
End Sub

Private Sub btnVisit_Click()
    Screen.MousePointer = vbHourglass
    DoEvents
    DoEvents
    Dim i As Integer
    If IsNumeric(lstQueue.BoundText) = False Then
        MsgBox "Please select a patient"
        lstQueue.SetFocus
        Exit Sub
    End If
    If SingleVisit = True Then
        Unload frmNewVisit
        Unload frmPatientDetails
        Unload frmPastHistory
    End If
    
    currentDoctorId = Val(cmbDoctor.BoundText)
    
    Dim NewVisitID As Long
    Dim ThisPatientID As Long
    Dim ThisQueueID As Long
    Dim VisitWt As Double
    Dim SBP As Long
    Dim DBP As Long
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSql = "Select * from tblQueue where QueueID = " & Val(lstQueue.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            ThisPatientID = !PatientID
            VisitWt = !weight
            SBP = !SBP
            DBP = !DBP
            !Completed = True
            !CompletedTime = Now
            !CompletedUserID = UserID
            .Update
            ThisQueueID = !QueueID
        End If
    End With
    
    NewVisitID = CurrentVisit.NewVisitID(ThisPatientID, Date, Val(cmbDoctor.BoundText), Val(cmbSecession.BoundText), ThisQueueID)
    CurrentPatient.ID = ThisPatientID
    
    If SingleVisit = True Then
        Unload frmNewVisit
        frmNewVisit.Show
        frmNewVisit.ZOrder 0
        frmNewVisit.txtWeight = VisitWt
        frmNewVisit.txtSBP = SBP
        frmNewVisit.txtDBP = DBP
    Else
        Dim frmQueueNewVisit As New frmNewVisit
        frmQueueNewVisit.Show
        frmQueueNewVisit.ZOrder 0
        frmQueueNewVisit.txtWeight = VisitWt
        frmQueueNewVisit.txtSBP = SBP
        frmQueueNewVisit.txtDBP = DBP
    End If
    FillQueue
    Screen.MousePointer = vbDefault
    DoEvents
End Sub

Private Sub cmbDoctor_Change()
    Call FillQueue
End Sub

Private Sub cmbSecession_Change()
    Call FillQueue
End Sub

Private Sub dtpDate_Change()
    Call FillQueue
End Sub

Private Sub Form_Activate()
    
    ColourChange = True
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

Private Sub Form_Deactivate()
    ColourChange = False
End Sub

Private Sub Form_Load()
    Call SetColours
    Call FillCombos
    Call GetFormSettings
    dtpDate.Value = Date
    If AutoSecession = True Then Call SetAutoSecession
    Call FillQueue

End Sub

Private Sub SetAutoSecession()
'     If AutoSecession = True Then Call SetAutoSecession
    
    If Time > #2:00:00 PM# Then
        cmbSecession.Text = "Evening"
    Else
        cmbSecession.Text = "Morning"
    End If
End Sub

Private Sub SaveQueueDetails()

End Sub

Private Sub FillCombos()
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

Private Sub FillQueue()
    Dim temQID As Long
    temQID = Val(lstQueue.BoundText)
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSql = "SELECT tblQueue.QueueNo & ' ' & tblPatient.Patient & ' (' & tblPatient.PatientID & ') ' & ' - ' & tblQueue.Age  as QueueDisplay, tblQueue.QueueID FROM tblQueue LEFT JOIN tblPatient ON tblQueue.PatientID = tblPatient.PatientID where DoctorID = " & Val(cmbDoctor.BoundText) & " AND QueueDate = #" & Format(dtpDate.Value, "dd MMMM yyyy") & "# AND SEcessionID = " & Val(cmbSecession.BoundText) & " AND tblQueue.Deleted = False "
        
        If optCompleted.Value = True Then
            temSql = temSql & " and Completed = true "
        ElseIf optToComplete.Value = True Then
            temSql = temSql & " and  Completed = false AND Removed = False "
        ElseIf optRemoved.Value = True Then
            temSql = temSql & " AND Removed = true "
        ElseIf optToIssue.Value = True Then
            temSql = temSql & " AND (Issued = False AND Completed = True) "
        ElseIf optIssued.Value = True Then
            temSql = temSql & " AND Issued = true "
        ElseIf optAll.Value = True Then
        
        End If
        
        temSql = temSql & " Order by QueueNo"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With lstQueue
        Set .RowSource = rsQueue
        .ListField = "QueueDisplay"
        .BoundColumn = "QueueID"
        .BoundText = temQID
    End With
End Sub


Private Sub lstQueue_DblClick()
    btnVisit_Click
End Sub

Private Sub optAll_Click()
    Call FillQueue
End Sub

Private Sub optCompleted_Click()
    Call FillQueue
End Sub

Private Sub optIssued_Click()
    Call FillQueue
End Sub

Private Sub optRemoved_Click()
    Call FillQueue
End Sub

Private Sub optToComplete_Click()
    Call FillQueue
End Sub

Private Sub optToIssue_Click()
    Call FillQueue
End Sub

Private Sub Timer1_Timer()
    Call FillQueue
End Sub
