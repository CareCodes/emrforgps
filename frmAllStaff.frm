VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAllStaff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "All Staff Details"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11985
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
   ScaleHeight     =   7785
   ScaleWidth      =   11985
   Begin VB.CheckBox chkUser 
      Caption         =   "User"
      Height          =   255
      Left            =   9120
      TabIndex        =   28
      Top             =   6600
      Width           =   2415
   End
   Begin VB.CheckBox chkDoctor 
      Caption         =   "Doctor"
      Height          =   255
      Left            =   6360
      TabIndex        =   27
      Top             =   6600
      Width           =   2415
   End
   Begin VB.TextBox txtUserName 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   6360
      TabIndex        =   26
      Top             =   3000
      Width           =   5535
   End
   Begin VB.TextBox txtPassword 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   6360
      PasswordChar    =   "*"
      TabIndex        =   14
      Top             =   3480
      Width           =   5535
   End
   Begin VB.TextBox txtDesignation 
      Height          =   720
      Left            =   6360
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   5535
   End
   Begin VB.TextBox txtStaff 
      Height          =   360
      Left            =   8040
      TabIndex        =   7
      Top             =   120
      Width           =   3855
   End
   Begin VB.TextBox txtComments 
      Height          =   960
      Left            =   6360
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   5400
      Width           =   5535
   End
   Begin VB.TextBox txtAddress 
      Height          =   1440
      Left            =   6360
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   1440
      Width           =   5535
   End
   Begin VB.TextBox txtTelephone 
      Height          =   360
      Left            =   6360
      TabIndex        =   16
      Top             =   3960
      Width           =   5535
   End
   Begin VB.TextBox txteMail 
      Height          =   360
      Left            =   6360
      TabIndex        =   18
      Top             =   4440
      Width           =   5535
   End
   Begin VB.TextBox txtWebsite 
      Height          =   345
      Left            =   6360
      TabIndex        =   20
      Top             =   4920
      Width           =   5535
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   2
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
      TabIndex        =   3
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
      TabIndex        =   4
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
   Begin MSDataListLib.DataCombo cmbStaff 
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
      Left            =   9720
      TabIndex        =   25
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
      Left            =   8280
      TabIndex        =   24
      Top             =   7200
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
      Left            =   6960
      TabIndex        =   23
      Top             =   7200
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
   Begin MSDataListLib.DataCombo cmbTitle 
      Height          =   360
      Left            =   6360
      TabIndex        =   6
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label10 
      Caption         =   "User Name"
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label8 
      Caption         =   "Password"
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label Label9 
      Caption         =   "Designation"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Staff"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Staff"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Comments"
      Height          =   255
      Left            =   4680
      TabIndex        =   21
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Telephone"
      Height          =   255
      Left            =   4680
      TabIndex        =   15
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "eMail"
      Height          =   255
      Left            =   4680
      TabIndex        =   17
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Website"
      Height          =   255
      Left            =   4680
      TabIndex        =   19
      Top             =   4920
      Width           =   1215
   End
End
Attribute VB_Name = "frmAllStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub btnAdd_Click()
    Dim temText As String
    If IsNumeric(cmbStaff.BoundText) = False Then
        temText = cmbStaff.Text
    Else
        temText = Empty
    End If
    cmbStaff.Text = Empty
    Call EditMode
    txtStaff.Text = temText
    cmbTitle.SetFocus
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbStaff.Text = Empty
    cmbStaff.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete " & cmbStaff.Text, vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblStaff where StaffID = " & Val(cmbStaff.BoundText)
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
    cmbStaff.SetFocus
    cmbStaff.Text = Empty
    
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbStaff.BoundText) = False Then Exit Sub
    Call EditMode
    txtStaff.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtStaff.Text) = Empty Then
        MsgBox "You have not entered a Staff"
        txtStaff.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbStaff.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbStaff.Text = Empty
    cmbStaff.SetFocus
End Sub

Private Sub cmbStaff_Change()
    Call ClearValues
    If IsNumeric(cmbStaff.BoundText) = True Then Call DisplayDetails
End Sub

Private Sub cmbTitle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbTitle.Text = Empty
    End If
End Sub

Private Sub Form_Load()
    Call SelectMode
    Call FillCombos
    GetCommonSettings Me
    
    Call SetColours
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

Private Sub EditMode()
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    cmbStaff.Enabled = False
    
    txtStaff.Enabled = True
    txtComments.Enabled = True
    txtAddress.Enabled = True
    txtUserName.Enabled = True
    txteMail.Enabled = True
    txtWebsite.Enabled = True
    txtTelephone.Enabled = True
    txtDesignation.Enabled = True
    txtPassword.Enabled = True
    cmbTitle.Enabled = True
    chkDoctor.Enabled = True
    chkUser.Enabled = True
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    
End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    cmbStaff.Enabled = True
    
    txtStaff.Enabled = False
    txtComments.Enabled = False
    txtAddress.Enabled = False
    txtUserName.Enabled = False
    txteMail.Enabled = False
    txtWebsite.Enabled = False
    txtDesignation.Enabled = False
    txtPassword.Enabled = False
    txtTelephone.Enabled = False
    cmbTitle.Enabled = False
    chkDoctor.Enabled = False
    chkUser.Enabled = False
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
End Sub

Private Sub ClearValues()
    txtStaff.Text = Empty
    txtComments.Text = Empty
    txtAddress.Text = Empty
    txtUserName.Text = Empty
    txteMail.Text = Empty
    txtWebsite.Text = Empty
    txtTelephone.Text = Empty
    txtDesignation.Text = Empty
    txtPassword.Text = Empty
    cmbTitle.Text = Empty
    chkUser.Value = 0
    chkDoctor.Value = 0
End Sub

Private Sub SaveNew()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblStaff"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Staff = txtStaff.Text
        !TitleID = Val(cmbTitle.BoundText)
        !Comments = txtComments.Text
        !Address = txtAddress.Text
        !Designation = txtDesignation.Text
        !Password = EncreptedWord(txtPassword.Text)
        !Telephone = txtTelephone.Text
        !eMail = txteMail.Text
        !Website = txtWebsite.Text
        !UserName = EncreptedWord(txtUserName.Text)
        
        If chkDoctor.Value = 1 Then
            !IsADoctor = True
        Else
            !IsADoctor = False
        End If
        
        If chkUser.Value = 1 Then
            !IsAUser = True
        Else
            !IsAUser = False
        End If
        
        .Update
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub SaveOld()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblStaff where StaffID = " & Val(cmbStaff.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        !Staff = txtStaff.Text
        !TitleID = Val(cmbTitle.BoundText)
        !Comments = txtComments.Text
        !Address = txtAddress.Text
        !Designation = txtDesignation.Text
        !Password = EncreptedWord(txtPassword.Text)
        !Telephone = txtTelephone.Text
        !eMail = txteMail.Text
        !Website = txtWebsite.Text
        !UserName = EncreptedWord(txtUserName.Text)

        If chkDoctor.Value = 1 Then
            !IsADoctor = True
        Else
            !IsADoctor = False
        End If
        
        If chkUser.Value = 1 Then
            !IsAUser = True
        Else
            !IsAUser = False
        End If
        
        .Update
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim It As New clsFillCombos
    It.FillAnyCombo cmbStaff, "Staff", True
    Dim Title As New clsFillCombos
    Title.FillAnyCombo cmbTitle, "Title", True
End Sub

Private Sub DisplayDetails()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblStaff where StaffID = " & Val(cmbStaff.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtStaff.Text = !Staff
            If IsNull(!TitleID) = False Then
                cmbTitle.BoundText = !TitleID
            End If
            txtComments.Text = Format(!Comments, "")
            txtAddress.Text = Format(!Address, "")
            txtDesignation.Text = Format(!Designation, "")
            txtPassword.Text = DecreptedWord(Format(!Password, ""))
            txtTelephone.Text = Format(!Telephone, "")
            txteMail.Text = Format(!eMail, "")
            txtWebsite.Text = Format(!Website, "")
            txtUserName.Text = DecreptedWord(Format(!UserName, ""))
            
            If !IsADoctor = True Then
                chkDoctor.Value = 1
            Else
                chkDoctor.Value = 0
            End If
            
            If !IsAUser = True Then
                chkUser.Value = 1
            Else
                chkUser.Value = 0
            End If
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub
