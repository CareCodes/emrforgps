VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPatientIxReception 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Investigation Reception"
   ClientHeight    =   8385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatientIxReception.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtVisitIxID 
      Height          =   375
      Left            =   1080
      TabIndex        =   36
      Top             =   4200
      Width           =   855
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   7680
      Width           =   3735
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   7200
      Width           =   3735
   End
   Begin VB.TextBox txtPatientIxID 
      Height          =   375
      Left            =   7440
      TabIndex        =   27
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtVisitID 
      Height          =   375
      Left            =   1080
      TabIndex        =   26
      Top             =   4680
      Width           =   855
   End
   Begin VB.TextBox txtQueueID 
      Height          =   375
      Left            =   1080
      TabIndex        =   25
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox txtIxCost 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtIxValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CheckBox chkNewPatient 
      Height          =   255
      Left            =   9240
      TabIndex        =   22
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox txtCost 
      Height          =   375
      Left            =   1080
      TabIndex        =   17
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox txtNTotal 
      Height          =   375
      Left            =   8760
      TabIndex        =   16
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox txtDiscount 
      Height          =   360
      Left            =   8760
      TabIndex        =   15
      Top             =   6840
      Width           =   1695
   End
   Begin VB.TextBox txtGTotal 
      Height          =   375
      Left            =   8760
      TabIndex        =   14
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "&Print"
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   7800
      Value           =   1  'Checked
      Width           =   855
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   9240
      TabIndex        =   0
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
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
   Begin MSDataListLib.DataCombo cmbPatientName 
      Height          =   360
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   5295
      _ExtentX        =   9340
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
   Begin MSDataListLib.DataCombo cmbDoc 
      Height          =   360
      Left            =   2040
      TabIndex        =   2
      Top             =   600
      Width           =   5295
      _ExtentX        =   9340
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
   Begin MSDataListLib.DataCombo cmbIns 
      Height          =   360
      Left            =   2040
      TabIndex        =   3
      Top             =   1080
      Width           =   5295
      _ExtentX        =   9340
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
   Begin btButtonEx.ButtonEx btnSettle 
      Height          =   375
      Left            =   7920
      TabIndex        =   7
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Settle"
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   9240
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
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
      Height          =   375
      Left            =   9240
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
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
   Begin MSDataListLib.DataCombo cmbIx 
      Height          =   360
      Left            =   2040
      TabIndex        =   11
      Top             =   1560
      Width           =   5295
      _ExtentX        =   9340
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
   Begin MSFlexGridLib.MSFlexGrid GridIx 
      Height          =   4095
      Left            =   2040
      TabIndex        =   13
      Top             =   2040
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7223
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin btButtonEx.ButtonEx btnNewPatient 
      Height          =   375
      Left            =   9240
      TabIndex        =   21
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "New Patient"
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
      Left            =   2040
      TabIndex        =   28
      Top             =   6720
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   2040
      TabIndex        =   30
      Top             =   6240
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label11 
      Caption         =   "Paper"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Label Label10 
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Label Label9 
      Caption         =   "&Payment Method"
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "&Drawer"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "N&et Total"
      Height          =   375
      Left            =   6720
      TabIndex        =   20
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Dis&count"
      Height          =   375
      Left            =   6720
      TabIndex        =   19
      Top             =   6840
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "&Total"
      Height          =   375
      Left            =   6720
      TabIndex        =   18
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "&Investigation"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Patient &Name"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Referring Doctor"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Referring Institution"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
End
Attribute VB_Name = "frmPatientIxReception"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim MyIx As New clsIx
    Dim i As Integer
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub GetSettings()
    GetCommonSettings Me
    cmbDrawer.BoundText = GetSetting(App.EXEName, Me.Name, cmbDrawer.Name, 1)
    cmbPaymentMethod.BoundText = GetSetting(App.EXEName, Me.Name, cmbPaymentMethod.Name, 1)
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
    SaveSetting App.EXEName, Me.Name, cmbDrawer.Name, Val(cmbDrawer.BoundText)
    SaveSetting App.EXEName, Me.Name, cmbPaymentMethod.Name, Val(cmbPaymentMethod.BoundText)
End Sub

Private Sub SetColours()
    Me.ForeColor = DefaultColourScheme.LabelForeColour
    Me.BackColor = DefaultColourScheme.LabelBackColour
    cmbDrawer.BoundText = GetSetting(App.EXEName, Me.Name, cmbDrawer.Name, 1)
    
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

Private Sub btnAdd_Click()
    If IsNumeric(cmbIx.BoundText) = False Then
        MsgBox "Please select an Investigation"
        cmbIx.SetFocus
        Exit Sub
    End If
    GridIx.Rows = GridIx.Rows + 1
    GridIx.row = GridIx.Rows - 1
    
    GridIx.col = 0
    GridIx.Text = GridIx.row
            
    GridIx.col = 1
    GridIx.Text = cmbIx.Text
            
    GridIx.col = 2
    GridIx.Text = Val(txtIxValue.Text)
            
    GridIx.col = 3
    GridIx.Text = Val(cmbIx.BoundText)
            
    GridIx.col = 4
    GridIx.Text = Val(txtIxCost.Text)
                
    GridIx.col = 5
    GridIx.Text = Val(txtPatientIxID.Text)
                
    Call ColourGrid(GridIx)
    Call CalculateTotals
    Call CalculateNetTotal
    Call ClearAddValues
    cmbIx.SetFocus
End Sub

Private Sub CalculateTotals()
    Dim TotalValue As Double
    Dim TotalCost As Double
    For i = 1 To GridIx.Rows - 1
        TotalValue = TotalValue + Val(GridIx.TextMatrix(i, 2))
        TotalCost = TotalCost + Val(GridIx.TextMatrix(i, 4))
    Next
    txtGTotal.Text = Format(TotalValue, "0.00")
    txtCost.Text = Format(TotalCost, "0.00")
End Sub

Private Sub CalculateNetTotal()
    txtNTotal.Text = Format(Val(txtGTotal.Text) - Val(txtDiscount.Text), "0.00")
End Sub

Private Sub btnDelete_Click()
    With GridIx
        .RemoveItem (.row)
    End With
    Call CalculateTotals
    Call CalculateNetTotal
    Call ClearAddValues
    Call ColourGrid(GridIx)
    cmbIx.SetFocus
End Sub

Private Sub btnNewPatient_Click()
    frmPatient.Show
    frmPatient.ZOrder 0
End Sub

Private Sub btnSettle_Click()
    Dim temPatientID As Long
    If IsNumeric(cmbPatientName.BoundText) = False Then
        MsgBox "Please select a patient?"
        cmbPatientName.SetFocus
        Exit Sub
    End If
    
    Call WritePatientIx(WritePatientIxBill)
    
    ChangeDrawerBalances Val(cmbDrawer.BoundText), Val(cmbPaymentMethod.BoundText), Val(txtNTotal.Text), True, False, False
    
    If chkPrint.Value = 1 Then Call eLabPrint(temPatientID)
    
    Call FormatGrid
    Call ClearAddValues
    Call ClearSettleValues
    Call ColourGrid
    
    cmbPatientName.SetFocus

End Sub

Private Function WritePatientIxBill() As Long
    Dim rsTem As New ADODB.Recordset
    For i = 1 To GridIx.Rows - 1
        With rsTem
            If .State = 1 Then .Close
            temSql = "SELECT * FROM tblPatientIxBill"
            .Open temSql, cnnLab, adOpenStatic, adLockOptimistic
            .AddNew
            !Cost = Val(txtCost.Text)
            !BillDateTime = Now
            !BillDate = Date
            !BillTime = Time
            !BillUserID = UserID
            !PatientID = Val(cmbPatientName.BoundText)
            !VisitID = Val(txtVisitID.Text)
            !PatientIxID = Val(txtPatientIxID.Text)
            !QueueID = Val(txtQueueID.Text)
            !DepartmentID = UserDepartmentID
            !IxBillCompleted = True
            !IxCategoryID = 1
            !PaymentMethodID = Val(cmbpaymentmethodid.BoundText)
            !GrossTotal = Val(txtGTotal.Text)
            !Discount = Val(txtDiscount.Text)
            !DiscountPercent = (Val(txtDiscount.Text) / Val(txtGTotal.Text)) * 100
            !NetTotal = Val(txtNTotal.Text)
            !Completed = True
            !CompletedDate = Date
            !CompletedTime = Time
            !CompletedDateTime = Now
            !CompletedUserID = UserID
            !DrawerID = Val(cmbDrawer.BoundText)
            .Update
            WritePatientIxBill = !IxBillID
            .Close
        End With
    Next
End Function

Private Sub WritePatientIx(BillID As Long)
    Dim i As Integer
    For i = 1 To .Rows - 1
        With rsPatientIx
        If .State = 1 Then .Close
        temSql = "SELECT * from tblPatientIX where PatientIxID = " & Val(GridIx.TextMatrix(i, 5))
        .Open temSql, cnnLab, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
            
            Else
                .AddNew
                !Requested = True
                !RequestedDate = Date
                !RequestedTime = Time
                !RequestedDateTime = Now
                !RequestedUserID = UserID
            End If
            !PatientIxBillID = BillID
            !IxID = Val(GridIx.TextMatrix(i, 3))
            !PatientID = Val(cmbPatientName.BoundText)
            !VisitIxID = Val(txtVisitIxID.Text)
            !VisitID = Val(txtVisitID.Text)
            !DoctorID = Val(cmbDoc.BoundText)
            !InstitutionID = Val(cmbIns.BoundText)
            !Value = Val(GridIx.TextMatrix(i, 2))
            !Cost = Val(GridIx.TextMatrix(i, 4))

            !Billed = True
            !BilledDate = Date
            !BilledTime = Time
            !BilledDateTime = Now
            !BilledUserID = UserID
            !PaymentMethodID = Val(cmbPaymentMethod.BoundText)
            !DrawerID = Val(cmbDrawer.BoundText)
            !DepartmentID = UserDepartmentID
            !PatientIxBillID = temPatientIxBillID
            
            !Printed = False
            !Issued = False
            !Deleted = False
        .Update
        End With
    Next
End Sub


Private Sub cmbDoc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbIns.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbDoc.Text = Empty
    End If
End Sub

Private Sub cmbIns_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbIx.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbIns.Text = Empty
    End If
End Sub

Private Sub cmbIx_Change()
    MyIx.ID = Val(cmbIx.BoundText)
    txtIxValue.Text = Format(MyIx.Value, "0.00")
    txtIxCost.Text = Format(MyIx.Cost, "0.00")
End Sub

Private Sub ClearAddValues()
    cmbIx.Text = Empty
    txtIxValue.Text = Empty
    txtCost.Text = Empty
End Sub

Private Sub ClearBillValues()
    txtGTotal.Text = Empty
    txtDiscount.Text = Empty
    txtNTotal.Text = Empty
    cmbPatientName.Text = Empty
    cmbIns.Text = Empty
    cmbDoc.Text = Empty
End Sub

Private Sub cmbIx_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnAdd_Click
    ElseIf KeyCode = vbKeyEscape Then
        cmbIx.Text = Empty
    End If
End Sub

Private Sub cmbpatientName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbDoc.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbPatientName.Text = Empty
    ElseIf KeyCode = vbKeyF5 Then
        cmbPatientName.Text = Empty
        Call FillCombos
    End If
End Sub

Private Sub cmbPrinter_Change()
    FillPrinterPapers cmbPrinter, cmbPaper
End Sub

Private Sub cmbPrinter_Click()
    FillPrinterPapers cmbPrinter, cmbPaper
End Sub

Private Sub Form_Activate()
    If chkNewPatient.Value = 1 Then
        Call FillCombos
        cmbPatientName.BoundText = CurrentPatient.ID
        chkNewPatient.Value = 0
    End If
End Sub

Private Sub Form_Load()
    FillPrinters cmbPrinter
    Call SetColours
    Call FillCombos
    Call GetSettings
    Call FormatGrid
    Call ColourGrid(GridIx)
End Sub

Private Sub FillCombos()
    Dim Doc As New clsFillCombos
    Doc.FillBoolCombo cmbDoc, "staff", "staff", "IsADoctor", True
    Dim RefIns As New clsFillCombos
    RefIns.FillBoolCombo cmbIns, "Institution", "Institution", "IsHOspital", True
    Dim Pt As New clsFillCombos
    Pt.FillAnyCombo cmbPatientName, "Patient", True
    Dim Ix As New clsFillCombos
    Ix.FillAnyCombo cmbIx, "Ix", True
    Dim Drawer As New clsFillCombos
    Drawer.FillAnyCombo cmbDrawer, "Drawer", True
    Dim Pay As New clsFillCombos
    Pay.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "ToReceive", True
End Sub

Private Sub FormatGrid()
    With GridIx
        .Clear
        .Cols = 6
        .Rows = 1
        
        .row = 0
        
        .col = 0
        .Text = "No."
        
        .col = 1
        .Text = "Ix"
        
        .col = 2
        .Text = "Value"
        
        .col = 3
        .Text = "IxID"
        
        .col = 4
        .Text = "Cost"
        
        .col = 5
        .Text = "PatientIxID"
        
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub GridIx_Click()
    With GridIx
        .col = .Cols - 1
        .ColSel = 0
    End With
End Sub

Private Sub txtDiscount_Change()
    Call CalculateNetTotal
End Sub
