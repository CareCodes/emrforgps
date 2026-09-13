VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmTodaysIncome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Income"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
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
   ScaleHeight     =   8595
   ScaleWidth      =   10200
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   6600
      Width           =   2055
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   4695
   End
   Begin MSFlexGridLib.MSFlexGrid gridIncome 
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   7011
      _Version        =   393216
      ScrollTrack     =   -1  'True
      AllowUserResizing=   2
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   495
      Left            =   8760
      TabIndex        =   10
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Delete"
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
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   495
      Left            =   8760
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Add"
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
      Height          =   855
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   1560
      Width           =   4695
   End
   Begin MSDataListLib.DataCombo cmbDepartment 
      Height          =   360
      Left            =   2160
      TabIndex        =   16
      Top             =   8160
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   131858435
      CurrentDate     =   39776
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   8760
      TabIndex        =   18
      Top             =   8040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "C&lose"
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
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   7320
      TabIndex        =   17
      Top             =   8040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Print"
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
   Begin MSDataListLib.DataCombo cmbDrawer 
      Height          =   360
      Left            =   2160
      TabIndex        =   14
      Top             =   7680
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   2160
      TabIndex        =   12
      Top             =   7200
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label8 
      Caption         =   "Total"
      Height          =   375
      Left            =   5880
      TabIndex        =   19
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Payment Method"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   7200
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Drawer"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   7680
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Value"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Comments"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Category"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Department"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   8160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTodaysIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String

Private Sub btnAdd_Click()
    Dim i As Integer
    Dim temBillID As Long
    If IsNumeric(txtValue.Text) = False Then
        MsgBox "Please enter a value"
        txtValue.SetFocus
        On Error Resume Next: SendKeys "{home}+{end}"
        Exit Sub
    End If
    If IsNumeric(cmbDepartment.BoundText) = False Then
        MsgBox "Please select a Unit"
        cmbDepartment.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbCategory.BoundText) = False Then
        MsgBox "Please select a category"
        cmbCategory.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbPaymentMethod.BoundText) = False Then
        MsgBox "Please select a Payment Method"
        cmbPaymentMethod.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbDrawer.BoundText) = False Then
        MsgBox "Please select a Drawer"
        cmbDrawer.SetFocus
        Exit Sub
    End If
    Dim rsIncome As New ADODB.Recordset
    With rsIncome
        If .State = 1 Then .Close
        temSql = "Select * from tblIncome"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Value = Val(txtValue.Text)
        !IncomeCategoryID = Val(cmbCategory.BoundText)
        !DepartmentID = Val(cmbDepartment.BoundText)
        !Comments = txtComments.Text
        !IncomeDate = Format(dtpDate.Value, "dd MMMM yyyy")
        !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
        !IncomeMethodID = IncomeMethod.OtherIncome
        !IncomeUserID = UserID
        !IncomeDateTime = Now
        !IncomeDate = Date
        !IncomeTime = Time
        !AddedDate = Date
        !AddedUserID = UserID
        !AddedTime = Time
        !AddedDateTime = Now
        !DrawerID = Val(cmbDrawer.BoundText)
        .Update
        .Close
    End With
    Dim temBool As Boolean
    temBool = ChangeDrawerBalances(Val(cmbDrawer.BoundText), Val(cmbPaymentMethod.BoundText), Val(txtValue.Text), True, False, False)
    Call ClearAddValues
    Call FormatGrid
    Call FillGrid
    cmbCategory.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim temBool As Boolean
    With GridIncome
        Dim temRow As Integer
        temRow = .row
        If .Rows <= 1 Then Exit Sub
        If .row < 1 Then Exit Sub
        If IsNumeric(.TextMatrix(temRow, 4)) = False Then Exit Sub
    End With
    Dim rsIncome As New ADODB.Recordset
    With rsIncome
        If .State = 1 Then .Close
        temSql = "Select * from tblIncome where IncomeID = " & Val(GridIncome.TextMatrix(temRow, 4))
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedDate = Date
            !DeletedTime = Time
            !DeletedDateTime = Now
            !DeletedUserID = UserID
            .Update
            temBool = ChangeDrawerBalances(!DrawerID, !PaymentMethodID, !Value, False, True, False)
        End If
        .Close
    End With
    Call FormatGrid
    Call FillGrid
    cmbCategory.SetFocus
End Sub

Private Sub btnPrint_Click()
    Dim tabReport As Long
    Dim tab1 As Long
    Dim tab2 As Long

    tabReport = 70
    tab1 = 5
    tab2 = 40
    
    Printer.Print
    Printer.Font.Bold = True
    Printer.Print Tab(tabReport); "Income Report"
    Printer.Font.Bold = False
    Printer.Print
    Printer.Print Tab(tab1); "Department :";
    Printer.Print Tab(tab2); cmbDepartment.Text
    Printer.Print Tab(tab1); "Date  :";
    Printer.Print Tab(tab2); dtpDate.Value;
    Printer.Print

    Dim i As Integer
    Dim tabNo As Long
    Dim tabCategory As Long
    Dim tabComments As Long
    Dim tabValue As Long
    
    tabNo = 10
    tabCategory = 20
    tabComments = 60
    tabValue = 115

    With GridIncome
        For i = 0 To .Rows - 1
            Printer.Print
            Printer.Print Tab(tabNo - Len(.TextMatrix(i, 0))); .TextMatrix(i, 0);
            Printer.Print Tab(tabCategory); .TextMatrix(i, 1);
            Printer.Print Tab(tabComments); .TextMatrix(i, 2);
            Printer.Print Tab(tabValue - Len(.TextMatrix(i, 3))); .TextMatrix(i, 3);
            Printer.Print
        Next
    End With

   Printer.EndDoc
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

Private Sub cmbDepartment_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpDate_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbCategory.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        dtpDate.Value = Date
    End If
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FormatGrid
    Call GetSettings
    Call SetColours
End Sub

Private Sub FillCombos()
    Dim PeyMethod As New clsFillCombos
    PeyMethod.FillAnyCombo cmbPaymentMethod, "PaymentMEthod", True
    Dim Cat As New clsFillCombos
    Cat.FillBoolCombo cmbCategory, "Category", "Category", "IsIncomeCategory", True
    Dim Drawer As New clsFillCombos
    Drawer.FillAnyCombo cmbDrawer, "Drawer", True
    Dim Dept As New clsFillCombos
    Dept.FillAnyCombo cmbDepartment, "Department", True
End Sub

Private Sub FormatGrid()
    With GridIncome
        .Clear
        .Cols = 5
        .Rows = 1
        .row = 0
        .col = 0
        .CellAlignment = 4
        .Text = "No."
        .col = 1
        .CellAlignment = 4
        .Text = "Category"
        .col = 2
        .CellAlignment = 4
        .Text = "Comments"
        .col = 3
        .CellAlignment = 4
        .Text = "Value"
        .col = 4
        .Text = "ID"
        .ColWidth(0) = 600
        .ColWidth(1) = 2500
        .ColWidth(2) = 3500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1
    End With
End Sub

Private Sub FillGrid()
    If IsNumeric(cmbDepartment.BoundText) = False Then Exit Sub
    Dim MyTotal As Double
    Dim rsIncome As New ADODB.Recordset
    With rsIncome
        If .State = 1 Then .Close
        temSql = "SELECT tblIncome.*, tblCategory.Category FROM tblIncome LEFT JOIN tblCategory ON tblIncome.IncomeCategoryID = tblCategory.CategoryID where tblIncome.IncomeDate = #" & Format(dtpDate.Value, "dd MMMM yyyy") & "# And tblIncome.DepartmentID = " & Val(cmbDepartment.BoundText) & " And tblIncome.Deleted = false  AND IncomeMethodID = " & IncomeMethod.OtherIncome
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            .MoveLast
            .MoveFirst
            Dim i As Integer
            GridIncome.Rows = .RecordCount + 1
            For i = 1 To .RecordCount
                GridIncome.TextMatrix(i, 0) = i
                GridIncome.TextMatrix(i, 1) = !Category
                GridIncome.TextMatrix(i, 2) = Format(!Comments, "")
                GridIncome.TextMatrix(i, 3) = Format(!Value, "#,##0.00")
                MyTotal = MyTotal + !Value
                GridIncome.TextMatrix(i, 4) = !IncomeID
                GridIncome.row = i
                If GridIncome.row Mod 2 = 0 Then
                    GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
                Else
                    GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
                End If
                .MoveNext
            Next
        End If
    End With
    txtTotal.Text = Format(MyTotal, "0.00")
End Sub

Private Sub ClearAddValues()
    txtComments.Text = Empty
    txtValue.Text = Empty
    cmbCategory.Text = Empty
End Sub

Private Sub ClearUpdateValues()

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
    SaveSetting App.EXEName, Me.Name, cmbDepartment.Name, cmbDepartment.BoundText
End Sub

Private Sub GetSettings()
    dtpDate.Value = Date
    GetCommonSettings Me
    
    cmbDrawer.BoundText = GetSetting(App.EXEName, Me.Name, cmbDrawer.Name, 1)
    cmbPaymentMethod.BoundText = GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1)
    cmbDepartment.BoundText = GetSetting(App.EXEName, Me.Name, cmbDepartment.Name, 1)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub txtComments_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnAdd_Click
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
