VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPatientIx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Investigations"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
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
   ScaleHeight     =   7770
   ScaleWidth      =   6840
   Begin VB.TextBox txtIxID 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2880
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtIxEditID 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   3360
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtIxDeleteID 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2400
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSDataListLib.DataCombo cmbIx 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid gridIx 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   11245
      _Version        =   393216
      WordWrap        =   -1  'True
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnIxAdd 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   7200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   65280
      Caption         =   "&Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnIxDelete 
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   7200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   255
      Caption         =   "&Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnIxResults 
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   7200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   65535
      Caption         =   "&Results"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Investigation"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmPatientIx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim FormPatient As New clsPatient
    Dim FormVisit As New clsVisit
    
Private Sub FillIx()
    Dim rsIx As New ADODB.Recordset
    With rsIx
        If .State = 1 Then .Close
        temSql = "SELECT tblPatientIX.PatientIXID, tblIx.Ix, tblIx.IxID FROM tblIx RIGHT JOIN tblPatientIX ON tblIx.IxID = tblPatientIX.IxID where tblPatientIX.Deleted = False AND tblPatientIX.PatientID = " & FormPatient.ID & " Order by PatientIXID Desc"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        
        gridIx.Clear
        gridIx.Rows = 1
        gridIx.Cols = 6
        gridIx.ColWidth(0) = 0
        gridIx.ColWidth(1) = gridIx.Width - 175
        gridIx.ColWidth(2) = 0
        gridIx.ColWidth(3) = 0
        gridIx.ColWidth(4) = 2500
        gridIx.ColWidth(5) = 2500
        gridIx.row = 0
        gridIx.col = 1
        gridIx.Text = "Ix"
        
        While .EOF = False
            gridIx.Rows = gridIx.Rows + 1
            gridIx.row = gridIx.Rows - 1
            gridIx.col = 0
            gridIx.Text = !PatientIxID
            gridIx.col = 1
            gridIx.Text = !Ix
            gridIx.col = 3
            gridIx.Text = !IxID
            .MoveNext
        Wend
    End With
    Set rsIx = Nothing
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

Private Sub Form_Load()
    GetCommonSettings Me
    FormPatient.ID = CurrentPatient.ID
    FormVisit.VisitID = CurrentVisit.VisitID
    Call SetColours
    Call FillCombos
    Call FillIx
End Sub


Private Sub FillCombos()
    Dim Ix As New clsFillCombos
    Ix.FillAnyCombo cmbIx, "Ix", True
End Sub


Private Sub btnIxAdd_Click()
    If IsNumeric(cmbIx.BoundText) = False Then
        MsgBox "Please select a Ix"
        cmbIx.SetFocus
        Exit Sub
    End If
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If Val(txtIxEditID.Text) <> 0 Then
            temSql = "Select * from tblPatientIX where PatientIXID = " & Val(txtIxEditID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !IxID = Val(cmbIx.BoundText)
                .Update
            Else
                .AddNew
                !PatientID = FormPatient.ID
                !IxID = Val(cmbIx.BoundText)
                !Requested = True
                !RequestedDate = Date
                !RequestedTime = Time
                !RequestedDateTime = Now
                !RequestedUserID = UserID
                .Update
            End If
        Else
            temSql = "Select * from tblPatientIX"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !PatientID = FormPatient.ID
            !IxID = Val(cmbIx.BoundText)
            !Requested = True
            !RequestedDate = Date
            !RequestedTime = Time
            !RequestedDateTime = Now
            !RequestedUserID = UserID
            
            .Update
        End If
        .Close
    End With
    Call FillIx
    cmbIx.Text = Empty
    txtIxEditID.Text = Empty
    txtIxID.Text = Empty
    cmbIx.SetFocus
    On Error Resume Next: SendKeys "{Escape}"
End Sub

Private Sub btnIxDelete_Click()
    If IsNumeric(txtIxDeleteID.Text) = False Then
        MsgBox "Please select a Ix to delete"
        Exit Sub
    End If
    Dim rsDx As New ADODB.Recordset
    With rsDx
        If Val(txtIxDeleteID.Text) <> 0 Then
            temSql = "Select * from tblPatientIX where PatientIXID = " & Val(txtIxDeleteID.Text)
            If .State = 1 Then .Close
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Deleted = True
                !DeletedTime = Now
                !DeletedUserID = UserID
                .Update
            End If
        End If
        .Close
    End With
    Call FillIx
    cmbIx.Text = Empty
    txtIxEditID.Text = Empty
    txtIxID.Text = Empty
    cmbIx.SetFocus
    On Error Resume Next: SendKeys "{Escape}"
End Sub

Private Sub btnIxResults_Click()
    Dim temGridRow As Long
    Dim temPatientIXID As Long
    Dim rsTemPatientIx As New ADODB.Recordset
    With rsTemPatientIx
        If .State = 1 Then .Close
        temSql = "Select * from tblPatientIx where PatientIXID = " & Val(txtIxDeleteID.Text) & " AND IxID = " & Val(txtIxID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            temPatientIXID = !PatientIxID
        Else
            .AddNew
            !PatientID = FormPatient.ID
            !IxID = Val(txtIxID.Text)
            !DoctorID = UserID
            !PatientIxID = Val(txtIxDeleteID.Text)
            !Requested = True
            !RequestedDate = Date
            !RequestedTime = Time
            !RequestedDateTime = Now
            !RequestedUserID = UserID
            
            .Update
            temPatientIXID = !PatientIxID
        End If
        .Close
    End With
    Unload frmIxReport
    frmIxReport.myPATIENTiXid = temPatientIXID
    frmIxReport.InvestigationDetails
    frmIxReport.Show
    frmIxReport.ZOrder 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub

Private Sub gridIx_Click()
    Dim temGridRow As Long
    With gridIx
        temGridRow = .row
        .col = 3
        txtIxID.Text = .Text
        If IsNumeric(.Text) = False Then
            MsgBox "Please select an Investigation"
            gridIx.SetFocus
            Exit Sub
        End If
        .col = 0
        txtIxDeleteID.Text = .Text
        If IsNumeric(.Text) = False Then
            MsgBox "Please select an Investigation"
            gridIx.SetFocus
            Exit Sub
        End If
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub gridIx_DblClick()
    Dim temGridRow As Long
    Dim temPatientIXID As Long
    With gridIx
        temGridRow = .row
        .col = 0
        txtIxEditID.Text = .Text
        .ColSel = .Cols - 1
        .col = 3
        cmbIx.BoundText = Val(.Text)
    End With
End Sub

