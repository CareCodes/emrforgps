VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDoctors 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctors Details"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12120
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
   ScaleHeight     =   8490
   ScaleWidth      =   12120
   Begin VB.TextBox txtHowToAddress 
      Height          =   360
      Left            =   6360
      TabIndex        =   15
      Top             =   3840
      Width           =   5535
   End
   Begin VB.TextBox txtDesignation 
      Height          =   720
      Left            =   6360
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   960
      Width           =   5535
   End
   Begin VB.TextBox txtDoctor 
      Height          =   360
      Left            =   8040
      TabIndex        =   7
      Top             =   480
      Width           =   3855
   End
   Begin VB.TextBox txtComments 
      Height          =   960
      Left            =   6360
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   6240
      Width           =   5535
   End
   Begin VB.TextBox txtAddress 
      Height          =   1440
      Left            =   6360
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1800
      Width           =   5535
   End
   Begin VB.TextBox txtCode 
      Height          =   360
      Left            =   6360
      TabIndex        =   17
      Top             =   4320
      Width           =   5535
   End
   Begin VB.TextBox txtTelephone 
      Height          =   360
      Left            =   6360
      TabIndex        =   19
      Top             =   4800
      Width           =   5535
   End
   Begin VB.TextBox txteMail 
      Height          =   360
      Left            =   6360
      TabIndex        =   21
      Top             =   5280
      Width           =   5535
   End
   Begin VB.TextBox txtWebsite 
      Height          =   345
      Left            =   6360
      TabIndex        =   23
      Top             =   5760
      Width           =   5535
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   6480
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
      TabIndex        =   3
      Top             =   6480
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
      TabIndex        =   4
      Top             =   6480
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
   Begin MSDataListLib.DataCombo cmbDoctor 
      Height          =   5940
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   10478
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   10680
      TabIndex        =   28
      Top             =   7920
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
      Left            =   9240
      TabIndex        =   27
      Top             =   7320
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
      Left            =   7920
      TabIndex        =   26
      Top             =   7320
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
   Begin MSDataListLib.DataCombo cmbSpeciality 
      Height          =   360
      Left            =   6360
      TabIndex        =   13
      Top             =   3360
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbTitle 
      Height          =   360
      Left            =   6360
      TabIndex        =   6
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label10 
      Caption         =   "Speciality"
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "How to Address"
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Designation"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Doctors"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Doctor"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Comments"
      Height          =   255
      Left            =   4680
      TabIndex        =   24
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Code"
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Telephone"
      Height          =   255
      Left            =   4680
      TabIndex        =   18
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "eMail"
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Website"
      Height          =   255
      Left            =   4680
      TabIndex        =   22
      Top             =   5760
      Width           =   1215
   End
End
Attribute VB_Name = "frmDoctors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub btnAdd_Click()
    Dim temText As String
    If IsNumeric(cmbDoctor.BoundText) = False Then
        temText = cmbDoctor.Text
    Else
        temText = Empty
    End If
    cmbDoctor.Text = Empty
    Call EditMode
    txtDoctor.Text = temText
    cmbTitle.SetFocus
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbDoctor.Text = Empty
    cmbDoctor.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete " & cmbDoctor.Text, vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblStaff where StaffID = " & Val(cmbDoctor.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedTime = Now
            !DeletedUserID = UserID
            .Update
            MsgBox "Deleted"
        Else
            MsgBox "Nothing to Delete"
        End If
        .Close
    End With
    Set rsTem = Nothing
    Call FillCombos
    cmbDoctor.SetFocus
    cmbDoctor.Text = Empty
    
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbDoctor.BoundText) = False Then Exit Sub
    Call EditMode
    txtDoctor.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtDoctor.Text) = Empty Then
        MsgBox "You have not entered a Doctor"
        txtDoctor.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbDoctor.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbDoctor.Text = Empty
    cmbDoctor.SetFocus
End Sub

Private Sub cmbDoctor_Change()
    Call ClearValues
    If IsNumeric(cmbDoctor.BoundText) = True Then Call DisplayDetails
End Sub


Private Sub cmbSpeciality_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbSpeciality.Text = Empty
    End If
End Sub

Private Sub cmbTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbTitle.Text = Empty
    End If
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
    Call SetColours
    GetCommonSettings Me
    
    
    Call SelectMode
    Call FillCombos
End Sub

Private Sub EditMode()
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    cmbDoctor.Enabled = False
    
    txtDoctor.Enabled = True
    txtComments.Enabled = True
    txtAddress.Enabled = True
    txtCode.Enabled = True
    txteMail.Enabled = True
    txtWebsite.Enabled = True
    txtTelephone.Enabled = True
    txtDesignation.Enabled = True
    cmbSpeciality.Enabled = True
    txtHowToAddress.Enabled = True
    cmbTitle.Enabled = True
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    
End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    cmbDoctor.Enabled = True
    
    txtDoctor.Enabled = False
    txtComments.Enabled = False
    txtAddress.Enabled = False
    txtCode.Enabled = False
    txteMail.Enabled = False
    txtWebsite.Enabled = False
    txtDesignation.Enabled = False
    cmbSpeciality.Enabled = False
    txtHowToAddress.Enabled = False
    txtTelephone.Enabled = False
    cmbTitle.Enabled = False
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
End Sub

Private Sub ClearValues()
    txtDoctor.Text = Empty
    txtComments.Text = Empty
    txtAddress.Text = Empty
    txtCode.Text = Empty
    txteMail.Text = Empty
    txtWebsite.Text = Empty
    txtTelephone.Text = Empty
    txtDesignation.Text = Empty
    cmbSpeciality.Text = Empty
    txtHowToAddress.Text = Empty
    cmbTitle.Text = Empty
End Sub

Private Sub SaveNew()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblStaff"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Staff = txtDoctor.Text
        !TitleID = Val(cmbTitle.BoundText)
        !Comments = txtComments.Text
        !Address = txtAddress.Text
        !Designation = txtDesignation.Text
        !Code = txtCode.Text
        !Telephone = txtTelephone.Text
        !eMail = txteMail.Text
        !Website = txtWebsite.Text
        !IsADoctor = True
        !HowToAddress = txtHowToAddress.Text
        !CategoryID = Val(cmbSpeciality.BoundText)
        .Update
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub SaveOld()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblStaff where StaffID = " & Val(cmbDoctor.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        !Staff = txtDoctor.Text
        !TitleID = Val(cmbTitle.BoundText)
        !Comments = txtComments.Text
        !Address = txtAddress.Text
        !Designation = txtDesignation.Text
        !Code = txtCode.Text
        !Telephone = txtTelephone.Text
        !eMail = txteMail.Text
        !Website = txtWebsite.Text
        !IsADoctor = True
        !HowToAddress = txtHowToAddress.Text
        !CategoryID = Val(cmbSpeciality.BoundText)
        .Update
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim It As New clsFillCombos
    It.FillBoolCombo cmbDoctor, "Staff", "Staff", "IsADoctor", True
    Dim Cat As New clsFillCombos
    Cat.FillBoolCombo cmbSpeciality, "Category", "Category", "IsDoctorSpeciality", True
    Dim Title As New clsFillCombos
    Title.FillAnyCombo cmbTitle, "Title", True
End Sub

Private Sub DisplayDetails()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblStaff where StaffID = " & Val(cmbDoctor.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtDoctor.Text = !Staff
            If IsNull(!TitleID) = False Then
                cmbTitle.BoundText = !TitleID
            End If
            txtComments.Text = Format(!Comments, "")
            txtAddress.Text = Format(!Address, "")
            txtDesignation.Text = Format(!Designation, "")
            txtCode.Text = Format(!Code, "")
            txtTelephone.Text = Format(!Telephone, "")
            txteMail.Text = Format(!eMail, "")
            txtWebsite.Text = Format(!Website, "")
            txtHowToAddress.Text = Format(!HowToAddress, "")
            cmbSpeciality.BoundText = !CategoryID
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub
