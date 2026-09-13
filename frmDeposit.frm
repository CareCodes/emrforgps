VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmDeposit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deposit"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6195
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
   ScaleHeight     =   4845
   ScaleWidth      =   6195
   Begin btButtonEx.ButtonEx btnUpdate 
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Update"
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
   Begin VB.TextBox txtComments 
      Height          =   1575
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   4455
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   1920
      Width           =   4455
   End
   Begin MSDataListLib.DataCombo cmbDrawer 
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Top             =   480
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Top             =   960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   4680
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Close"
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
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   1440
      TabIndex        =   5
      Top             =   1440
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label5 
      Caption         =   "Category"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Method"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Value"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Drawer"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmDeposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub ClearValues()
    txtComments.Text = Empty
    cmbCategory.Text = Empty
    txtValue.Text = Empty
End Sub

Private Sub btnUpdate_Click()
    If IsNumeric(cmbDrawer.BoundText) = False Then
        MsgBox "Please select a drawer"
        cmbDrawer.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPaymentMethod.BoundText) = False Then
        MsgBox "Please select a Payment Method"
        cmbPaymentMethod.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbCategory.BoundText) = False Then
        MsgBox "Please select a category"
        cmbCategory.SetFocus
        Exit Sub
    End If
    If ChangeDrawerBalances(Val(cmbDrawer.BoundText), Val(cmbPaymentMethod.BoundText), Val(txtValue.Text), True, False, False) = True Then
        Dim rsTem As New ADODB.Recordset
        With rsTem
            If .State = 1 Then .Close
            temSql = "Select * from tblDeposit"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !ToDepartmentID = UserID
            !DepositDate = Date
            !DepositTime = Time
            !DepositDateTime = Now
            !DepositValue = Val(txtValue.Text)
            !Comments = txtComments.Text
            !DepositUserID = UserID
            !DepositCategoryID = Val(cmbCategory.BoundText)
            !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
            !ToDrawerID = Val(cmbDrawer.BoundText)
            
            .Update
        End With
        MsgBox "Deposited"
        Call ClearValues
        cmbCategory.SetFocus
    Else
        MsgBox "NOT Deposited"
        cmbCategory.SetFocus
    End If
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtValue.SetFocus
        On Error Resume Next: SendKeys "{home}+{end}"
    ElseIf KeyCode = vbKeyEscape Then
        cmbCategory.Text = Empty
    End If
End Sub

Private Sub cmbDrawer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbPaymentMethod.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbDrawer.Text = Empty
    End If
End Sub

Private Sub cmbPaymentMethod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbCategory.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbPaymentMethod.Text = Empty
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call GetSettings
    Call SetColours
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub txtComments_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnUpdate_Click
    ElseIf KeyCode = vbKeyEscape Then
        txtComments.Text = Empty
    End If
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtComments.SetFocus
        On Error Resume Next: SendKeys "{home}+{end}"
    ElseIf KeyCode = vbKeyEscape Then
        txtValue.Text = Empty
    End If
End Sub

Private Sub txtValue_LostFocus()
    txtValue.Text = Format(Val(txtValue.Text), "0.00")
End Sub

Private Sub FillCombos()
    Dim PeyMethod As New clsFillCombos
    PeyMethod.FillAnyCombo cmbPaymentMethod, "PaymentMEthod", True
    
    Dim Cat As New clsFillCombos
    Cat.FillBoolCombo cmbCategory, "Category", "Category", "IsDepositCategory", True
    
    Dim Drawer As New clsFillCombos
    Drawer.FillAnyCombo cmbDrawer, "Drawer", True
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


Private Sub SaveSettings()
    SaveCommonSettings Me
    
    SaveSetting App.EXEName, Me.Name, cmbDrawer.Name, cmbDrawer.BoundText
    SaveSetting App.EXEName, Me.Name, cmbPaymentMethod.Name, cmbPaymentMethod.BoundText
End Sub

Private Sub GetSettings()
    GetCommonSettings Me
    
    cmbDrawer.BoundText = GetSetting(App.EXEName, Me.Name, cmbDrawer.Name, 1)
    cmbPaymentMethod.BoundText = GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1)
End Sub
