VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSystemFindings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Findings"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSystemFindings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10770
   Begin VB.Frame frameAge 
      Caption         =   "&Age Range"
      Height          =   1575
      Left            =   5040
      TabIndex        =   17
      Top             =   4920
      Width           =   5535
      Begin VB.TextBox txtTo 
         Height          =   360
         Left            =   3000
         TabIndex        =   23
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtFrom 
         Height          =   360
         Left            =   1080
         TabIndex        =   21
         Top             =   1080
         Width           =   855
      End
      Begin VB.OptionButton optLimitedAge 
         Caption         =   "&Limited to Age Range"
         Height          =   240
         Left            =   360
         TabIndex        =   19
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton optAllAges 
         Caption         =   "A&ll ages"
         Height          =   240
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblTo 
         Caption         =   "&To"
         Height          =   255
         Left            =   2280
         TabIndex        =   22
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblFrom 
         Caption         =   "&From"
         Height          =   255
         Left            =   360
         TabIndex        =   20
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.Frame frameSex 
      Caption         =   "Se&x"
      Height          =   1455
      Left            =   5040
      TabIndex        =   13
      Top             =   3360
      Width           =   5535
      Begin VB.OptionButton optBoth 
         Caption         =   "&Both"
         Height          =   240
         Left            =   960
         TabIndex        =   16
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton optFemale 
         Caption         =   "&Female"
         Height          =   240
         Left            =   960
         TabIndex        =   15
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optMale 
         Caption         =   "&Male"
         Height          =   240
         Left            =   960
         TabIndex        =   14
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame frameBelongTo 
      Caption         =   "&Belongs to"
      Height          =   1455
      Left            =   5040
      TabIndex        =   9
      Top             =   1800
      Width           =   5535
      Begin VB.OptionButton optHx 
         Caption         =   "&History"
         Height          =   240
         Left            =   960
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton optEx 
         Caption         =   "&Examination"
         Height          =   240
         Left            =   960
         TabIndex        =   11
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton optOther 
         Caption         =   "&Other"
         Height          =   240
         Left            =   960
         TabIndex        =   12
         Top             =   1080
         Width           =   1695
      End
   End
   Begin VB.TextBox txtFinding 
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   1320
      Width           =   4575
   End
   Begin MSDataListLib.DataCombo cmbFinding 
      Height          =   5700
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   10054
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSystem 
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   7200
      Width           =   1455
      _ExtentX        =   2566
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
      Left            =   1680
      TabIndex        =   5
      Top             =   7200
      Width           =   1335
      _ExtentX        =   2355
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   7200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Delete"
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
      Left            =   9360
      TabIndex        =   26
      Top             =   7200
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
      Left            =   8160
      TabIndex        =   25
      Top             =   6720
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
      Left            =   6840
      TabIndex        =   24
      Top             =   6720
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
   Begin VB.Label lblFindingText 
      Caption         =   "&Finding"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblFindingCmb 
      BackStyle       =   0  'Transparent
      Caption         =   "&Finding"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblSystem 
      Caption         =   "&System"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSystemFindings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim Finding As New clsFillCombos
    Dim rsSystemFinding As New ADODB.Recordset
    
Private Sub btnAdd_Click()
    If IsNumeric(cmbSystem.BoundText) = False Then
        MsgBox "Please select a system"
        cmbSystem.SetFocus
        Exit Sub
    End If
    cmbFinding.Text = Empty
    Call ClearValues
    Call EditMode
    txtFinding.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbFinding.Text = Empty
    cmbFinding.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    If IsNumeric(cmbFinding.BoundText) = False Then
        MsgBox "Please select a finding to delete"
        cmbFinding.SetFocus
        Exit Sub
    End If
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete?", vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblSystemFinding where SystemFindingID = " & Val(cmbFinding.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedTime = Now
            !DeletedUserID = UserID
            .Update
        End If
        .Close
    End With
    Call ClearValues
    Call FillCombos
    Call SelectMode
    cmbFinding.Text = Empty
    cmbFinding.SetFocus
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbSystem.BoundText) = False Then
        MsgBox "Please select a system"
        cmbSystem.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbFinding.BoundText) = False Then
        MsgBox "Please select a finding to edit"
        cmbFinding.SetFocus
        Exit Sub
    End If
    Call EditMode
    txtFinding.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtFinding.Text) = Empty Then
        MsgBox "Please enter a finding to save"
        txtFinding.SetFocus
        Exit Sub
    End If
    If optEx.Value = False And optHx.Value = False And optOther.Value = False Then
        MsgBox "Please select a value category"
        optHx.SetFocus
        Exit Sub
    End If
    If optMale.Value = False And optFemale.Value = False And optBoth.Value = False Then
        MsgBox "Please select the sex"
        optMale.SetFocus
        Exit Sub
    End If
    If optLimitedAge.Value = True And (IsNumeric(txtFrom.Text) = False Or IsNumeric(txtTo.Text) = False) Then
        MsgBox "If you have selected an age range, please enter two valid ages in years)"
        If IsNumeric(txtFrom.Text) = False Then
            txtFrom.SetFocus
        Else
            txtTo.SetFocus
        End If
        On Error Resume Next: SendKeys "{home}+{end}"
        Exit Sub
    End If
    If optLimitedAge.Value = True Then
        If Val(txtFrom.Text) > Val(txtTo.Text) Then
            MsgBox "The age range you entered is not valid. Please recheck"
            txtFrom.SetFocus
            Exit Sub
        End If
    ElseIf optAllAges.Value = True Then
        txtFrom.Text = 0
        txtTo.Text = 200
    End If
    If IsNumeric(cmbFinding.BoundText) = False Then
        Call SaveNew
    Else
        Call SaveOld
    End If
    Call ClearValues
    Call SelectMode
    Call FillCombos
    cmbFinding.Text = Empty
    cmbFinding.SetFocus
End Sub

Private Sub cmbFinding_Change()
    Call ClearValues
    If IsNumeric(cmbFinding.BoundText) = True Then Call DisplayDetails
End Sub

Private Sub cmbSystem_Change()
    Call FillCombos
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
    Call SetColours
    Call FillSysCOmbo
    Call SelectMode
End Sub

Private Sub FillSysCOmbo()
    Dim Systems As New clsFillCombos
    Systems.FillAnyCombo cmbSystem, "System", True
    cmbFinding.Text = Empty
End Sub

Private Sub FillCombos()
'    Finding.FillLongCombo cmbFinding, "SystemFinding", "SystemFinding", "SystemID", Val(cmbSystem.BoundText), True
'    cmbFinding.Text = Empty
    With rsSystemFinding
        If .State = 1 Then .Close
        temSql = "Select * from tblSystemFinding where Deleted = False and SystemID = " & Val(cmbSystem.BoundText) & " Order by IsHistory, IsExamination, IsOther , SystemFInding"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbFinding
        Set .RowSource = rsSystemFinding
        .ListField = "SystemFinding"
        .BoundColumn = "SystemFindingID"
    End With
End Sub

Private Sub ClearValues()
    txtFinding.Text = Empty
    optEx.Value = False
    optHx.Value = False
    optOther.Value = False
    optMale.Value = False
    optFemale.Value = False
    optBoth.Value = False
    optAllAges.Value = False
    optLimitedAge.Value = False
    txtFrom.Text = Empty
    txtTo.Text = Empty
End Sub

Private Sub DisplayDetails()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblSystemFinding where SystemFindingID = " & Val(cmbFinding.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtFinding.Text = !SystemFinding
            optEx.Value = !IsExamination
            optHx.Value = !IsHistory
            optOther.Value = !IsOther
            If !ForMale = True And !ForFemale = True Then
                optBoth.Value = True
            ElseIf !ForMale = True Then
                optMale.Value = True
            ElseIf !ForFemale = True Then
                optFemale.Value = True
            Else
                
            End If
            If !AllAges = True Then
                optAllAges.Value = True
            ElseIf !LimitedAge = True Then
                optLimitedAge.Value = True
                txtFrom.Text = !FromAge
                txtTo.Text = !ToAge
            End If
            
        End If
        .Close
    End With
End Sub

Private Sub SaveOld()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblSystemFinding where SystemFindingID = " & Val(cmbFinding.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !SystemFinding = txtFinding.Text
            !SystemID = Val(cmbSystem.BoundText)
            !IsExamination = optEx.Value
            !IsHistory = optHx.Value
            !IsOther = optOther.Value
            If optBoth.Value = True Then
                !ForMale = True
                !ForFemale = True
            ElseIf optMale.Value = True Then
                !ForMale = True
                !ForFemale = False
            ElseIf optFemale.Value = True Then
                !ForMale = False
                !ForFemale = True
            Else
                !ForMale = False
                !ForFemale = False
            End If
            !LimitedAge = optLimitedAge.Value
            !AllAges = optAllAges.Value
            !FromAge = Val(txtFrom.Text)
            !ToAge = Val(txtTo.Text)
            .Update
        End If
        .Close
    End With
End Sub

Private Sub SaveNew()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblSystemFinding"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !SystemFinding = txtFinding.Text
        !SystemID = Val(cmbSystem.BoundText)
        !IsExamination = optEx.Value
        !IsHistory = optHx.Value
        !IsOther = optOther.Value
        If optBoth.Value = True Then
            !ForMale = True
            !ForFemale = True
        ElseIf optMale.Value = True Then
            !ForMale = True
            !ForFemale = False
        ElseIf optFemale.Value = True Then
            !ForMale = False
            !ForFemale = True
        Else
            !ForMale = False
            !ForFemale = False
        End If
        !LimitedAge = optLimitedAge.Value
        !AllAges = optAllAges.Value
        !FromAge = Val(txtFrom.Text)
        !ToAge = Val(txtTo.Text)
        .Update
        .Close
    End With
End Sub

Private Sub SelectMode()
    cmbSystem.Enabled = True
    cmbFinding.Enabled = True
    btnEdit.Enabled = True
    btnAdd.Enabled = True
    btnDelete.Enabled = True

    lblSystem.Enabled = True
    lblFindingCmb.Enabled = True
    lblFindingText.Enabled = False
    frameAge.Enabled = False
    frameBelongTo.Enabled = False
    frameSex.Enabled = False
    lblFrom.Enabled = False
    lblTo.Enabled = False

    txtFinding.Enabled = False
    optEx.Enabled = False
    optHx.Enabled = False
    optOther.Enabled = False
    
    optBoth.Enabled = False
    optMale.Enabled = False
    optFemale.Enabled = False
    optAllAges.Enabled = False
    optLimitedAge.Enabled = False
    txtFrom.Enabled = False
    txtTo.Enabled = False

    btnSave.Enabled = False
    btnCancel.Enabled = False

End Sub

Private Sub EditMode()
    cmbSystem.Enabled = False
    cmbFinding.Enabled = False
    btnEdit.Enabled = False
    btnAdd.Enabled = False
    btnDelete.Enabled = False

    lblSystem.Enabled = False
    lblFindingCmb.Enabled = False
    lblFindingText.Enabled = True
    frameAge.Enabled = True
    frameBelongTo.Enabled = True
    frameSex.Enabled = True
    lblFrom.Enabled = True
    lblTo.Enabled = True

    txtFinding.Enabled = True
    optEx.Enabled = True
    optHx.Enabled = True
    optOther.Enabled = True

    optBoth.Enabled = True
    optMale.Enabled = True
    optFemale.Enabled = True
    optAllAges.Enabled = True
    optLimitedAge.Enabled = True
    txtFrom.Enabled = True
    txtTo.Enabled = True
    btnSave.Enabled = True
    btnCancel.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub

Private Sub optAllAges_Click()
    Call SetAges
End Sub

Private Sub SetAges()
    If optAllAges.Value = True Then
        lblFrom.Visible = False
        lblTo.Visible = False
        txtFrom.Visible = False
        txtTo.Visible = False
    ElseIf optLimitedAge.Value = True Then
        lblFrom.Visible = True
        lblTo.Visible = True
        txtFrom.Visible = True
        txtTo.Visible = True
    End If
End Sub

Private Sub optLimitedAge_Click()
    Call SetAges
End Sub
