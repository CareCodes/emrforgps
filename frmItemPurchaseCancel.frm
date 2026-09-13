VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmItemPurchaseCancel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Cancellation"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15240
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
   MDIChild        =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   15240
   Begin VB.TextBox txtRePaymentComments 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   11640
      MultiLine       =   -1  'True
      TabIndex        =   34
      Top             =   8025
      Width           =   3495
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   7560
      Width           =   3255
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   8040
      Width           =   3255
   End
   Begin VB.TextBox txtPaymentComments 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   360
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   645
      Width           =   3255
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   360
      Left            =   12000
      TabIndex        =   19
      Top             =   765
      Width           =   2535
   End
   Begin VB.TextBox txtInvoice 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   360
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   3015
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   62324739
      CurrentDate     =   39691
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   9600
      TabIndex        =   20
      Top             =   9000
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.TextBox txtBillID 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   23
      Top             =   1560
      Width           =   3015
   End
   Begin btButtonEx.ButtonEx btnPurchase 
      Height          =   375
      Left            =   11640
      TabIndex        =   21
      Top             =   9000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "Cancel &Purchase"
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
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   375
      Left            =   13560
      TabIndex        =   22
      Top             =   9000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
   Begin MSDataListLib.DataCombo cmbChecked 
      Height          =   360
      Left            =   6480
      TabIndex        =   7
      Top             =   1560
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbReceived 
      Height          =   360
      Left            =   6480
      TabIndex        =   9
      Top             =   2040
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSupplier 
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbPayment 
      Height          =   360
      Left            =   6480
      TabIndex        =   24
      Top             =   180
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbDrawer 
      Height          =   360
      Left            =   6480
      TabIndex        =   14
      Top             =   1110
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid GridItem 
      Height          =   4935
      Left            =   120
      TabIndex        =   32
      Top             =   2520
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   8705
      _Version        =   393216
      WordWrap        =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
   End
   Begin MSDataListLib.DataCombo cmbRePayment 
      Height          =   360
      Left            =   11640
      TabIndex        =   35
      Top             =   7560
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbReDrawer 
      Height          =   360
      Left            =   11640
      TabIndex        =   36
      Top             =   8490
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label11 
      Caption         =   "Re-Payment &Method"
      Height          =   255
      Left            =   9600
      TabIndex        =   39
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Label Label10 
      Caption         =   "Re-Paymen&t Details"
      Height          =   255
      Left            =   9600
      TabIndex        =   38
      Top             =   7980
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Repaid from Drawer"
      Height          =   255
      Left            =   9600
      TabIndex        =   37
      Top             =   8460
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Bill ID"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "Drawer"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   13
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Total Sale Value"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10080
      TabIndex        =   31
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label lblTotalSaleValue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Enabled         =   0   'False
      Height          =   375
      Left            =   12000
      TabIndex        =   30
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "&Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "PPaper"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   8040
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Paymen&t Details"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   11
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Invoice &Date"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "&Discount"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10080
      TabIndex        =   29
      Top             =   765
      Width           =   1215
   End
   Begin VB.Label lblNetTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Enabled         =   0   'False
      Height          =   375
      Left            =   12000
      TabIndex        =   26
      Top             =   1275
      Width           =   2535
   End
   Begin VB.Label lblGrossTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Enabled         =   0   'False
      Height          =   375
      Left            =   12000
      TabIndex        =   25
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label23 
      Caption         =   "Payment &Method"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   10
      Top             =   180
      Width           =   1935
   End
   Begin VB.Label Label22 
      Caption         =   "Invoice &No."
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label21 
      Caption         =   "C&hecked by"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "&Received by"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label24 
      Caption         =   "&Supplier"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Net Total"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10080
      TabIndex        =   28
      Top             =   1275
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Gross Total"
      Enabled         =   0   'False
      Height          =   255
      Left            =   10080
      TabIndex        =   27
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmItemPurchaseCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim NumForms As Long
    Dim FI1 As FORM_INFO_1
    Dim aFI1() As FORM_INFO_1
    Dim Temp() As Byte
    Dim BytesNeeded As Long
    Dim PrinterName As String
    Dim PrinterHandle As Long
    Dim FormItem As String
    Dim RetVal As Long
    Dim FormSize As SIZEL
    Dim SetPrinter As Boolean
    Dim SuppliedWord As String
    
    Dim FSys As New Scripting.FileSystemObject
    
    Private CsetPrinter As New cSetDfltPrinter
    
    Dim temSql As String
    Dim i As Integer
    
    Dim DisplayChange As Boolean
    
    Dim MyBatch As ItemBatch

    Dim FormPurchaseBill As New clsPurchaseBill
    Dim FormPurchaseItem As New clsItem
    
Private Sub FillCombos()
    Dim Sup As New clsFillCombos
    Sup.FillBoolCombo cmbSupplier, "Institution", "Institution", "IsSupplier", True
    Dim CheckedUser As New clsFillCombos
    CheckedUser.FillAnyCombo cmbChecked, "Staff", True
    Dim ReceivedUser As New clsFillCombos
    ReceivedUser.FillAnyCombo cmbReceived, "Staff", True
    Dim Drawer As New clsFillCombos
    Drawer.FillAnyCombo cmbDrawer, "Drawer", True
    Dim Pay As New clsFillCombos
    Pay.FillBoolCombo cmbPayment, "PaymentMethod", "PaymentMethod", "ToPay", True
    
    Dim ReDrawer As New clsFillCombos
    ReDrawer.FillAnyCombo cmbReDrawer, "Drawer", True
    Dim Repay As New clsFillCombos
    Repay.FillBoolCombo cmbRePayment, "PaymentMethod", "PaymentMethod", "ToReceive", True
End Sub

Private Sub RemoveItems()
    Dim temBool As Boolean
    Dim rsItem As New ADODB.Recordset
    With rsItem
        If .State = 1 Then .Close
        temSql = "Select * from tblPurchaseItem where Deleted = False AND PurchaseBillID = " & Val(txtBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            While .EOF = False
                temBool = ChangeStocks(!ItemID, !ItemBatchID, !PurchaseQuentity + !PurchaseFreeQuentity, !DepartmentID, False, True, False)
                !Deleted = True
                !DeletedDate = Date
                !DeletedDateTime = Now
                !DeletedTime = Time
                !DeletedUserID = UserID
                !ReturnValue = !PurchaseValue
                !SaleReturnValue = !SaleValue
                !ReturnRate = !PurchaseRate
                !ReturnQuentity = !PurchaseQuentity + !PurchaseFreeQuentity
                .Update
                .MoveNext
            Wend
        End If
        .Close
    End With
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnPurchase_Click()
    If Val(txtBillID.Text) = 0 Then
        MsgBox "Noting to cancel the purchase"
        Exit Sub
    End If
    If GridItem.Rows < 2 Then
        MsgBox "Noting to cancel the purchase"
        Exit Sub
    End If
    If Val(cmbReDrawer.BoundText) = 0 Then
        MsgBox "Please select a drawer"
        cmbReDrawer.SetFocus
        Exit Sub
    End If
    If Val(cmbRePayment.BoundText) = 0 Then
        MsgBox "Please select a re-payment method"
        cmbRePayment.SetFocus
        Exit Sub
    End If
    
    Call RemoveItems
    Call ChangeDrawerBalances(Val(cmbReDrawer.BoundText), Val(cmbRePayment.BoundText), Val(lblNetTotal.Caption), True, False, False)
    Call ChangeDrawerBalances(Val(cmbDrawer.BoundText), Val(cmbPayment.BoundText), Val(lblNetTotal.Caption), False, True, False)
    
    Dim rsExpence As New ADODB.Recordset
    With rsExpence
        If .State = 1 Then .Close
        temSql = "Select * from tblIncome"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !DepartmentID = UserDepartmentID
        !IncomeMethodID = 6
        !FromInstitutionID = Val(cmbSupplier.BoundText)
        !BillID = Val(txtBillID.Text)
        !PaymentMethodID = Val(cmbRePayment.BoundText)
        !PaymentComments = txtRePaymentComments.Text
        !Value = lblGrossTotal.Caption
        !AddedUserID = UserID
        !AddedDateTime = Now
        !AddedDate = Date
        !AddedTime = Time
        !IncomeDate = Date
        !IncomeTime = Time
        !IncomeDateTime = Now
        !IncomeUserID = UserID
        .Update
        .Close
    End With
    
    Dim rsPurchaseBill As New ADODB.Recordset
    With rsPurchaseBill
        If .State = 1 Then .Close
        temSql = "Select * from tblPurchaseBill where PurchaseBillID = " & Val(txtBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !cancelled = True
            !CancelledUserID = UserID
            !CancelledDateTime = Now
            !CancelledDate = Date
            !CancelledTime = Time
            !RePaymentMethodID = Val(cmbPayment.BoundText)
            !RePaymentComments = txtRePaymentComments.Text
            !ReDrawerID = Val(cmbDrawer.BoundText)
            !TotalReturnValue = Val(lblNetTotal.Caption)
            !TotalSaleReturnValue = Val(lblTotalSaleValue.Caption)
            .Update
            MsgBox "Successfully Cancelled"
        Else
            .Close
            MsgBox "Not Saved"
            Exit Sub
        End If
    End With
    
    If Val(cmbPayment.BoundText) = 2 Then
        Dim rsSup As New ADODB.Recordset
        With rsSup
            If .State = 1 Then .Close
            temSql = "Select * from tblInstitution where InstitutionID = " & Val(cmbSupplier.BoundText)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            If .RecordCount > 0 Then
                !Credit = !Credit - Val(lblNetTotal.Caption)
                .Update
            End If
            .Close
        End With
    End If
    
    ChangeDrawerBalances Val(cmbDrawer.BoundText), Val(cmbPayment.BoundText), Val(lblNetTotal.Caption), True, False, False
    
    Unload Me
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
    Call FillCombos
    Call FillPrinters
    Call FormatGrid
    Call GetSettings
End Sub

Private Sub GetSettings()
    Dim temCtrlName As String
    Dim MyControl As Control
    GetCommonSettings Me
    cmbRePayment.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbRePayment.Name, 1))
    chkPrint.Value = Val(GetSetting(App.EXEName, Me.Name, chkPrint.Name, 1))
    cmbReDrawer.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbReDrawer.Name, 1))
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, "Printer", "")
    cmbPrinter_Click
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, "Paper", "")
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
    SaveSetting App.EXEName, Me.Name, cmbRePayment.Name, cmbRePayment.BoundText
    SaveSetting App.EXEName, Me.Name, cmbReDrawer.Name, cmbReDrawer.BoundText
    SaveSetting App.EXEName, Me.Name, chkPrint.Name, chkPrint.Value
    SaveSetting App.EXEName, Me.Name, "Printer", cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, "Paper", cmbPaper.Text
End Sub

Private Sub cmbPrinter_Change()
    cmbPrinter_Click
End Sub

Private Sub cmbPrinter_Click()
    'cmbPaper.Clear
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        With FormSize
            .cx = PrescreptionPaperHeight
            .cy = PrescreptionPaperWidth
        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                cmbPaper.AddItem PtrCtoVbString(.pName)
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub FillPrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub FormatGrid()
    With GridItem
        .Cols = 11
        .Rows = 2
        .RowHeight(0) = .RowHeight(1) * 2
        .Rows = 1
        .Row = 0
        


        .Col = 0
        For i = 0 To .Cols - 1
            .Col = i
            .CellAlignment = 4
            Select Case i
                Case 1:     .Text = "No"
                            .ColWidth(i) = 400
                Case 2:     .Text = "Item"
                            .ColWidth(i) = 3600
                Case 3:     .Text = "Purchase Rate"
                            .ColWidth(i) = 1000
                Case 4:     .Text = "Purchase Qty"
                            .ColWidth(i) = 1000
                Case 5:     .Text = "Free Qty"
                            .ColWidth(i) = 1000
                Case 6:     .Text = "Purchase Value"
                            .ColWidth(i) = 1000
                Case 7:     .Text = "Batch"
                            .ColWidth(i) = 1400
                Case 8:     .Text = "Expiary"
                            .ColWidth(i) = 1600
                Case 9:     .Text = "Slaes Rate"
                            .ColWidth(i) = 1000
                Case 10:     .Text = "Sale Value"
                            .ColWidth(i) = 1000
                Case Else:  .ColWidth(i) = 0
            End Select
        Next i
    
    End With
    '   0   PurchaseItemID
    '   1   No
    '   2   Item
    '   3   Purchase Rate
    '   4   Purchase Qty
    '   5   Free Qty
    '   6   Purchase Value
    '   7   Batch
    '   8   Expiary
    '   9   Sale Rate
    '   10  Sale Value

End Sub

Private Sub FillGrid()
    Dim GrossTotal As Double
    Dim SaleValueTotal As Double
    Dim rsPItem As New ADODB.Recordset
    With rsPItem
        temSql = "SELECT tblItemBatch.ItemBatch, tblItemBatch.DOE, tblPurchaseItem.* FROM tblPurchaseItem LEFT JOIN tblItemBatch ON tblPurchaseItem.ItemBatchID = tblItemBatch.ItemBatchID WHERE tblPurchaseItem.PurchaseBillID = " & Val(txtBillID.Text) & " AND tblPurchaseItem.Deleted = false"
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            While .EOF = False
                FormPurchaseItem.ID = !ItemID
                GridItem.Rows = GridItem.Rows + 1
                GridItem.Row = GridItem.Rows - 1
                GridItem.Col = 0
                GridItem.Text = !PurchaseItemID
                GridItem.Col = 1
                GridItem.Text = GridItem.Row
                GridItem.Col = 2
                If Trim(FormPurchaseItem.Code) <> "" Then
                    GridItem.Text = FormPurchaseItem.Item & " (" & FormPurchaseItem.Code & ")"
                Else
                    GridItem.Text = FormPurchaseItem.Item
                End If
                GridItem.Col = 6
                GrossTotal = GrossTotal + !PurchaseValue
                SaleValueTotal = SaleValueTotal + !SaleValue
                GridItem.Text = Format(!PurchaseValue, "0.00")
                GridItem.Col = 7
                GridItem.Text = !ItemBatch
                GridItem.Col = 8
                GridItem.Text = Format(!DOE, "dd MMM yyyy")
'                If SSTab1.Tab = 0 Then
                    GridItem.Col = 3
                    GridItem.Text = Format(!PurchaseRate, "0.00")
                    GridItem.Col = 4
                    GridItem.Text = !PurchaseQuentity
                    GridItem.Col = 5
                    GridItem.Text = !PurchaseFreeQuentity
                    GridItem.Col = 9
                    GridItem.Text = Format(!SaleRate, "0.00")
                    GridItem.Col = 10
                    GridItem.Text = Format(!SaleValue, "0.00")
'                Else
'                    GridItem.Col = 3
'                    GridItem.Text = Format(!PurchaseRate * FormPurchaseItem.IssueUnitsPerPack, "0.00")
'                    GridItem.Col = 4
'                    GridItem.Text = Format(!PurchaseQuentity / FormPurchaseItem.IssueUnitsPerPack, "0.00")
'                    GridItem.Col = 5
'                    GridItem.Text = Format(!PurchaseFreeQuentity / FormPurchaseItem.IssueUnitsPerPack, "0.00")
'                    GridItem.Col = 9
'                    GridItem.Text = Format(!SaleRate * FormPurchaseItem.IssueUnitsPerPack, "0.00")
'                    GridItem.Col = 10
'                    GridItem.Text = Format(!SaleValue, "0.00")
'                End If
                .MoveNext
            Wend
        End If
        .Close
    End With
'    lblGrossTotal.Caption = Format(GrossTotal, "0.00")
'    lblTotalSaleValue.Caption = Format(SaleValueTotal, "0.00")
End Sub

Private Sub DisplayBillDetails(BillID As Long)
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblPurchaseBill where PurchaseBillID = " & Val(txtBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            If !cancelled = True Then
                MsgBox "This bill is already cancelled. You can't cancel it again"
                Unload Me
            End If
            cmbPayment.BoundText = Val(Format(!PaymentMethodID, "0"))
            cmbSupplier.BoundText = Val(Format(!SupplierID, "0"))
            cmbChecked.BoundText = Val(Format(!CheckedUserID, "0"))
            cmbReceived.BoundText = Val(Format(!ReceivedUserID, "0"))
            cmbDrawer.BoundText = Val(Format(!DrawerID, "0"))
            txtPaymentComments.Text = Format(!PaymentComments, "")
            dtpDate.Value = CDate(Format(!InvoiceDate, "dd MMMM yyyy"))
            txtInvoice.Text = Format(!InvoiceNumber, "")
            lblGrossTotal.Caption = Format(!Total, "0.00")
            txtDiscount.Text = Format(!Discount, "0.00")
            lblNetTotal.Caption = Format(!NetTotal, "0.00")
            lblTotalSaleValue.Caption = Format(!SaleValue, "0.00")
        Else
            MsgBox "No such purchase Bill Is Found"
            Unload Me
        End If
        .Close
    End With
End Sub

Private Sub txtBillID_Change()
    Call DisplayBillDetails(Val(txtBillID.Text))
    Call FormatGrid
    Call FillGrid
End Sub
