VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmInstitutions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Institutions"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10905
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInstitutions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   10905
   Begin VB.CheckBox chkInsurer 
      Caption         =   "Insurer"
      Height          =   255
      Left            =   8880
      TabIndex        =   12
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CheckBox chkHospital 
      Caption         =   "Hospital"
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CheckBox chkCustomer 
      Caption         =   "Customer"
      Height          =   255
      Left            =   8880
      TabIndex        =   10
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CheckBox chkSupplier 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CheckBox chkImporter 
      Caption         =   "Importer"
      Height          =   255
      Left            =   8880
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.CheckBox chkManufacturer 
      Caption         =   "Manufacturer"
      Height          =   255
      Left            =   6000
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.CheckBox chkBlackListed 
      Caption         =   "Blacklisted"
      Height          =   255
      Left            =   6000
      TabIndex        =   25
      Top             =   6000
      Width           =   3135
   End
   Begin VB.TextBox txtBalance 
      Height          =   345
      Left            =   6000
      TabIndex        =   24
      Top             =   5520
      Width           =   4575
   End
   Begin VB.TextBox txtWebsite 
      Height          =   345
      Left            =   6000
      TabIndex        =   22
      Top             =   5040
      Width           =   4575
   End
   Begin VB.TextBox txteMail 
      Height          =   345
      Left            =   6000
      TabIndex        =   20
      Top             =   4560
      Width           =   4575
   End
   Begin VB.TextBox txtTelephone 
      Height          =   345
      Left            =   6000
      TabIndex        =   18
      Top             =   4080
      Width           =   4575
   End
   Begin VB.TextBox txtCode 
      Height          =   345
      Left            =   6000
      TabIndex        =   16
      Top             =   3600
      Width           =   4575
   End
   Begin VB.TextBox txtAddress 
      Height          =   1425
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   2040
      Width           =   4575
   End
   Begin VB.TextBox txtComments 
      Height          =   960
      Left            =   6000
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   6480
      Width           =   4575
   End
   Begin VB.TextBox txtSupplier 
      Height          =   360
      Left            =   6000
      TabIndex        =   6
      Top             =   480
      Width           =   4575
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7560
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
      Top             =   7560
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
      Top             =   7560
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
   Begin MSDataListLib.DataCombo cmbSupplier 
      Height          =   6900
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   12171
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   9360
      TabIndex        =   30
      Top             =   8040
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
      TabIndex        =   29
      Top             =   7560
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
      TabIndex        =   28
      Top             =   7560
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
   Begin VB.Label Label8 
      Caption         =   "Balance"
      Height          =   255
      Left            =   4800
      TabIndex        =   23
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Website"
      Height          =   255
      Left            =   4800
      TabIndex        =   21
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "eMail"
      Height          =   255
      Left            =   4800
      TabIndex        =   19
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Telephone"
      Height          =   255
      Left            =   4800
      TabIndex        =   17
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Code"
      Height          =   255
      Left            =   4800
      TabIndex        =   15
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   255
      Left            =   4680
      TabIndex        =   13
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Comments"
      Height          =   255
      Left            =   4680
      TabIndex        =   26
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Institution"
      Height          =   255
      Left            =   4680
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Institutions"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmInstitutions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    
Private Sub btnAdd_Click()
    Dim temText As String
    If IsNumeric(cmbSupplier.BoundText) = False Then
        temText = cmbSupplier.Text
    Else
        temText = Empty
    End If
    cmbSupplier.Text = Empty
    Call EditMode
    txtSupplier.Text = temText
    txtSupplier.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbSupplier.Text = Empty
    cmbSupplier.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete " & cmbSupplier.Text, vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblInstitution where InstitutionID = " & Val(cmbSupplier.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
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
    cmbSupplier.SetFocus
    cmbSupplier.Text = Empty
    
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbSupplier.BoundText) = False Then Exit Sub
    Call EditMode
    txtSupplier.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtSupplier.Text) = Empty Then
        MsgBox "You have not entered a Supplier"
        txtSupplier.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbSupplier.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbSupplier.Text = Empty
    cmbSupplier.SetFocus
End Sub

Private Sub cmbSupplier_Change()
    Call ClearValues
    If IsNumeric(cmbSupplier.BoundText) = True Then Call DisplayDetails
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
    cmbSupplier.Enabled = False
    
    txtSupplier.Enabled = True
    txtComments.Enabled = True
    txtAddress.Enabled = True
    txtCode.Enabled = True
    txteMail.Enabled = True
    txtWebsite.Enabled = True
    txtBalance.Enabled = True
    txtTelephone.Enabled = True
    chkBlackListed.Enabled = True
    chkSupplier.Enabled = True
    chkImporter.Enabled = True
    chkManufacturer.Enabled = True
    chkHospital.Enabled = True
    chkCustomer.Enabled = True
    chkInsurer.Enabled = True
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    
End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    cmbSupplier.Enabled = True
    
    txtSupplier.Enabled = False
    txtComments.Enabled = False
    txtAddress.Enabled = False
    txtCode.Enabled = False
    txteMail.Enabled = False
    txtWebsite.Enabled = False
    txtBalance.Enabled = False
    txtTelephone.Enabled = False
    chkBlackListed.Enabled = False
    chkSupplier.Enabled = True
    chkImporter.Enabled = True
    chkManufacturer.Enabled = True
    chkHospital.Enabled = True
    chkCustomer.Enabled = True
    chkInsurer.Enabled = True
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
End Sub

Private Sub ClearValues()
    txtSupplier.Text = Empty
    txtComments.Text = Empty
    txtAddress.Text = Empty
    txtCode.Text = Empty
    txteMail.Text = Empty
    txtWebsite.Text = Empty
    txtBalance.Text = Empty
    txtTelephone.Text = Empty
    chkBlackListed.Value = 0
    chkSupplier.Value = 0
    chkImporter.Value = 0
    chkManufacturer.Value = 0
    chkHospital.Value = 0
    chkCustomer.Value = 0
    chkInsurer.Value = 0
End Sub

Private Sub SaveNew()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblInstitution"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Institution = txtSupplier.Text
        !Comments = txtComments.Text
        !Address = txtAddress.Text
        !InstitutionCode = txtCode.Text
        !Telephone = txtTelephone.Text
        !eMail = txteMail.Text
        !Website = txtWebsite.Text
        If chkBlackListed.Value = 1 Then
            !BlackListed = True
        Else
            !BlackListed = False
        End If
        If chkImporter.Value = 1 Then
            !IsImporter = True
        Else
            !IsImporter = False
        End If
        If chkManufacturer.Value = 1 Then
            !IsManufacturer = True
        Else
            !IsManufacturer = False
        End If
        If chkSupplier.Value = 1 Then
            !IsSupplier = True
        Else
            !IsSupplier = False
        End If
        If chkInsurer.Value = 1 Then
            !IsInsurer = True
        Else
            !IsInsurer = False
        End If
        If chkHospital.Value = 1 Then
            !IsHospital = True
        Else
            !IsHospital = False
        End If
        If chkCustomer.Value = 1 Then
            !IsCUstomer = True
        Else
            !IsCUstomer = False
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
        temSQL = "Select * from tblInstitution where InstitutionID = " & Val(cmbSupplier.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
        !Institution = txtSupplier.Text
        !Comments = txtComments.Text
        !Address = txtAddress.Text
        !InstitutionCode = txtCode.Text
        !Telephone = txtTelephone.Text
        !eMail = txteMail.Text
        !Website = txtWebsite.Text
        If chkBlackListed.Value = 1 Then
            !BlackListed = True
        Else
            !BlackListed = False
        End If
        If chkImporter.Value = 1 Then
            !IsImporter = True
        Else
            !IsImporter = False
        End If
        If chkManufacturer.Value = 1 Then
            !IsManufacturer = True
        Else
            !IsManufacturer = False
        End If
        If chkSupplier.Value = 1 Then
            !IsSupplier = True
        Else
            !IsSupplier = False
        End If
        If chkInsurer.Value = 1 Then
            !IsInsurer = True
        Else
            !IsInsurer = False
        End If
        If chkHospital.Value = 1 Then
            !IsHospital = True
        Else
            !IsHospital = False
        End If
        If chkCustomer.Value = 1 Then
            !IsCUstomer = True
        Else
            !IsCUstomer = False
        End If
        .Update
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim It As New clsFillCombos
    It.FillAnyCombo cmbSupplier, "Institution", True
End Sub

Private Sub DisplayDetails()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "Select * from tblInstitution where InstitutionID = " & Val(cmbSupplier.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtSupplier.Text = Format(!Institution, "")
            txtComments.Text = Format(!Comments, "")
            txtAddress.Text = Format(!Address, "")
            txtCode.Text = Format(!InstitutionCode, "")
            txtTelephone.Text = Format(!Telephone, "")
            txteMail.Text = Format(!eMail, "")
            txtWebsite.Text = Format(!Website, "")
            txtBalance.Text = Format(!Credit, "0.00")
            If !BlackListed = True Then
                chkBlackListed.Value = 1
            Else
                chkBlackListed.Value = 0
            End If
            If !IsImporter = True Then chkImporter.Value = 1
            If !IsCUstomer = True Then chkCustomer.Value = 1
            If !IsSupplier = True Then chkSupplier.Value = 1
            If !IsManufacturer = True Then chkManufacturer.Value = 1
'            If !IsHospital = True Then chkHospital.Value = 1
'            If !IsInsurer = True Then chkInsurer.Value = 1
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
End Sub
