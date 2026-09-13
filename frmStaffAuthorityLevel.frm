VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmStaffAuthorityLevel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff Details"
   ClientHeight    =   7800
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
   ScaleHeight     =   7800
   ScaleWidth      =   12120
   Begin MSDataListLib.DataCombo cmbStaff 
      Height          =   5940
      Left            =   120
      TabIndex        =   0
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
      TabIndex        =   5
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
      Left            =   9240
      TabIndex        =   4
      Top             =   6840
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
      TabIndex        =   3
      Top             =   6840
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
   Begin MSDataListLib.DataCombo cmbAuthority 
      Height          =   360
      Left            =   5280
      TabIndex        =   2
      Top             =   840
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "Authority"
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmStaffAuthorityLevel"
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
    SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtStaff.Text) = Empty Then
        MsgBox "You have not entered a Staff"
        txtStaff.SetFocus
        Exit Sub
    End If
    If NoSuchUserName = False Then
        MsgBox "The user name you entered is already taken, please try with something else"
        txtUserName.SetFocus
        SendKeys "{home}+{end}"
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
    Me.Top = GetSetting(App.EXEName, Me.Name, "Top", Me.Top)
    Me.Left = GetSetting(App.EXEName, Me.Name, "Left", Me.Left)
    Call SetColours
    Call SelectMode
    Call FillCombos
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
        !Email = txteMail.Text
        !Website = txtWebsite.Text
        !UserName = EncreptedWord(txtUserName.Text)
        !IsAUser = True
        .Update
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Function NoSuchUserName() As Boolean
    NoSuchUserName = True
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblStaff where Deleted = false AND StaffID <> " & Val(cmbStaff.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If IsNull(!UserName) = False Then
                txtOtherUserNames.Text = DecreptedWord(!UserName)
                If UCase(txtUserName.Text) = UCase(txtOtherUserNames.Text) Then
                    NoSuchUserName = False
                    Exit Function
                End If
            End If
            .MoveNext
        Wend
        .Close
    End With
End Function


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
        !Email = txteMail.Text
        !Website = txtWebsite.Text
        !UserName = EncreptedWord(txtUserName.Text)
        !IsAUser = True
        .Update
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim It As New clsFillCombos
    It.FillBoolCombo cmbStaff, "Staff", "staff", "IsAUser", True
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
            txteMail.Text = Format(!Email, "")
            txtWebsite.Text = Format(!Website, "")
            txtUserName.Text = DecreptedWord(Format(!UserName, ""))
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting App.EXEName, Me.Name, "Top", Me.Top
    SaveSetting App.EXEName, Me.Name, "Left", Me.Left

End Sub


