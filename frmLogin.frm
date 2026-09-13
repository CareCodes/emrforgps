VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSDataListLib.DataCombo cmbDepartment 
      Height          =   360
      Left            =   1800
      TabIndex        =   7
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
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
   Begin VB.TextBox txtTemUsername 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1560
      Width           =   3765
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   3735
   End
   Begin btButtonEx.ButtonEx cmdCancel 
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16711680
      Caption         =   "&Cancel"
      ForeColor       =   16777215
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
   Begin btButtonEx.ButtonEx cmdOK 
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      BackColor       =   16711680
      Caption         =   "&Login"
      ForeColor       =   16777215
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Department"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FSys As New Scripting.FileSystemObject
    Dim SuppliedWord As String
    
    Dim rsStaff As New ADODB.Recordset
    Dim rsHospital As New ADODB.Recordset
    Dim temSQL As String
    Dim constr As String
    Dim TemUserPassward As String
    
Private Sub cmbDepartment_Change()
On Error Resume Next
    Dim rsTemDept As New ADODB.Recordset
    With rsTemDept
        If .State = 1 Then .Close
        temSQL = "Select * from tblDepartment where DepartmentID  = " & Val(cmbDepartment.BoundText)
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            IndoorDepartmentID = !IndoorDepartmentID
            OutdoorDepartmentID = !OutdoorDepartmentID
        End If
        .Close
    End With
End Sub

Private Sub cmbDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyReturn Then
        txtUserName.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    End
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    Dim TemResponce As Byte
    Dim UserNameFound As Boolean
    UserNameFound = False
    If IsNumeric(cmbDepartment.BoundText) = False Then
        MsgBox "Please select a department"
        cmbDepartment.SetFocus
        Exit Sub
    End If
    TemUserPassward = txtPassword.Text
    If Trim(txtUserName.Text) = "" Then
        TemResponce = MsgBox("You have not entered a username", vbCritical, "Username")
        txtUserName.SetFocus
        Exit Sub
    End If
    If Trim(txtPassword.Text) = "" Then
        TemResponce = MsgBox("You have not entered a password", vbCritical, "Password")
        txtPassword.SetFocus
        Exit Sub
    End If
    With rsStaff
        If .State = 1 Then .Close
        temSQL = "Select tblstaff.* from tblstaff where Deleted = False AND IsAUser=True AND UserName Is not NUll and Password  is Not  Null"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount < 1 Then Exit Sub
        .MoveFirst
        Do While Not .EOF
            txtTemUsername.Text = DecreptedWord(!UserName)
            If UCase(txtUserName.Text) = UCase(txtTemUsername.Text) Then
                UserNameFound = True
                If txtPassword.Text = DecreptedWord(!Password) Then
                    UserName = UCase(txtUserName.Text)
                    UserID = !StaffID
                    UserDepartment = cmbDepartment.Text
                    UserDepartmentID = Val(cmbDepartment.BoundText)

                    If Not IsNull(!AuthorityID) Then
                        UserAuthority = !AuthorityID
                    Else
                        UserAuthority = 0
                    End If
                    Select Case UserAuthority
                        Case 1: UserAuthorityLevel = Viewer
                        Case 2: UserAuthorityLevel = Analyzer
                        Case 3: UserAuthorityLevel = OrdinaryUser
                        Case 4: UserAuthorityLevel = PowerUser
                        Case 5: UserAuthorityLevel = SuperUser
                        Case 6: UserAuthorityLevel = Administrator
                        Case Else: UserAuthorityLevel = NotIdentified
                    End Select
                    Exit Do
                Else
                    TemResponce = MsgBox("The username and password you entered are not matching. Please try again", vbCritical, "Wrong Username and Password")
                    txtUserName.SetFocus
                    On Error Resume Next: SendKeys "{home}+{end}"
                    Exit Sub
                End If
            Else
            End If
            .MoveNext
        Loop
        .Close
        If UserNameFound = False Then
            TemResponce = MsgBox("There is no such  a username, Please try again", vbCritical, "Username")
            txtUserName.SetFocus
            On Error Resume Next: SendKeys "{home}+{end}"
            Exit Sub
        End If
        End With
        
'        If DemoCopy = True Then
'            If Date > ExpiaryDate Or DemoCount > 5000 Then
'                MsgBox "Demo Copy EXPIRED"
'                End
'            Else
'                MsgBox "Demo Copy Only. Will Expire Soon"
'                SaveSetting App.EXEName, "Options", "DemoCount", DemoCount + 1
'            End If
'        End If
        DoEvents
        Screen.MousePointer = vbHourglass
        DoEvents
        MDIMain.Show
        Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call SetColours
    Dim Dept As New clsFillCombos
    Dept.FillAnyCombo cmbDepartment, "Department", True
    cmbDepartment.BoundText = Val(GetSetting(App.EXEName, Me.Name, "DepartmentID", 0))
    Call alterDatabase
End Sub


Private Sub alterDatabase()
    Call createIndoorIssue
    Call updateSaleItem
End Sub

Private Sub updateSaleItem()
    On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "ALTER TABLE tblSaleItem ADD COLUMN IndoorBillID INTEGER;"
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        .Update
        If .State = 1 Then .Close
    End With

End Sub

Private Sub createIndoorIssue()
    On Error Resume Next
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSQL = "CREATE TABLE tblIndoorBill ([IndoorBillID] AUTOINCREMENT, [BillDateTime]   DATETIME, [BillDate]    DATETIME, [BillTime]    DATETIME, [BillUserID]  INTEGER, [DepartmentID]     INTEGER, [QueueID] INTEGER, [SaleBillCompleted]     YesNo, [SaleCategoryID] INTEGER, [PaymentMethodID]  INTEGER,  [GrossTotal] DOUBLE, [Discount]  DOUBLE, [DiscountPercent]   DOUBLE, [NetTotal]  DOUBLE, [Cancelled]  YesNo, [CancelledDateTime]  DATETIME,  [CancelledDate]  DATETIME,  [CancelledTime]  DATETIME, [CancelledUserID]     INTEGER, [PurchaseValue]   DOUBLE, " & _
                    "[Completed]     YesNo, [CompletedDate]  DATETIME,  [CompletedTime]  DATETIME, [CompletedDateTime]   DATETIME, [CompletedUserID]     INTEGER, [Deleted]  YesNo, [DeletedDateTime]    DATETIME, [DeletedDate]     DATETIME, [DeletedTime]     DATETIME, [DeletedUserID]   INTEGER, [TotalReturnValue]    DOUBLE, [TotalSaleReturnValue]  DOUBLE, [SaleValueMonthly]  DOUBLE, [PurchaseValueMonthly]  DOUBLE, [SaleValueAnnually] DOUBLE, [PurchaseValueAnnually] DOUBLE, [ForOutDoor]         YesNo, " & _
                    "[ForStaff]      YesNo,  [ForBHT]        YesNo, [ForUnit]        YesNo, [ForInstitution]         YesNo, [CustomerID]     INTEGER,  [DrawerID]    INTEGER,  [Comments] LONGTEXT, " & _
                    " CONSTRAINT [PK_tblIndoorBill] PRIMARY KEY ([IndoorBillID]))"
                    
        .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
        If .State = 1 Then .Update
        If .State = 1 Then .Close
    End With
End Sub

Private Sub SetColours()
On Error Resume Next
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


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    SaveSetting App.EXEName, Me.Name, "DepartmentID", cmbDepartment.BoundText
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 And txtUserName.Text <> "" Then cmdOK_Click: Exit Sub
    If KeyAscii = 13 And txtUserName.Text = "" Then txtUserName.SetFocus: Exit Sub
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then txtPassword.SetFocus
End Sub

