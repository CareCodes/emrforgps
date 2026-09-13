VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSearchToCancelPurchaseBills 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Purchase Bills to Cancel"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12495
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
   ScaleHeight     =   9165
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   8640
      Width           =   3735
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   7920
      Width           =   3735
   End
   Begin VB.TextBox txtBillID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2400
      TabIndex        =   18
      Top             =   7080
      Width           =   1815
   End
   Begin VB.TextBox txtNetTotal 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10440
      TabIndex        =   15
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txtRepayments 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10440
      TabIndex        =   13
      Top             =   7560
      Width           =   1935
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   11160
      TabIndex        =   7
      Top             =   8520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
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
   Begin VB.TextBox txtGrossTotal 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10440
      TabIndex        =   6
      Top             =   7080
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid GridIncome 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   9551
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   90701827
      CurrentDate     =   39859
   End
   Begin MSDataListLib.DataCombo cmbSupplier 
      Height          =   360
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnPrint 
      Height          =   495
      Left            =   9840
      TabIndex        =   8
      Top             =   8520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Appearance      =   3
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
   Begin MSDataListLib.DataCombo cmbPaymentMethod 
      Height          =   360
      Left            =   1200
      TabIndex        =   9
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   90701827
      CurrentDate     =   39859
   End
   Begin btButtonEx.ButtonEx btnView 
      Height          =   375
      Left            =   4320
      TabIndex        =   17
      Top             =   7080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Go to Cancel"
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
   Begin VB.Label Label10 
      Caption         =   "Paper"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   8400
      Width           =   2655
   End
   Begin VB.Label Label9 
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Label Label8 
      Caption         =   "Purchase Bill ID"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label Label7 
      Caption         =   "Net Purchase Value"
      Height          =   255
      Left            =   8280
      TabIndex        =   16
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "Total Repayments"
      Height          =   255
      Left            =   8280
      TabIndex        =   14
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "To"
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Payment"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Total Purchase Value"
      Height          =   255
      Left            =   8280
      TabIndex        =   5
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Supplier"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "From"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmSearchToCancelPurchaseBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
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
    Dim i As Integer

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub GetSettings()
    GetCommonSettings Me
    dtpFrom.Value = Date
    dtpTo.Value = Date
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
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

Private Sub btnPrint_Click()
    Dim ThisReportFormat As PrintReport
    Dim temTopic As String
    Dim temSubTopic As String
    
    temTopic = InstitutionName & " - Purchase Bills from "
    
    If cmbSupplier.Text <> "" Then
        temTopic = cmbSupplier.Text
    Else
        temTopic = "All Suppliers "
    End If
    
    If dtpFrom.Value = dtpTo.Value Then
        temSubTopic = "on " & Format(dtpFrom.Value, "dd MMMM yyyy")
    Else
        temSubTopic = "From " & Format(dtpFrom.Value, "dd MMMM yyyy") & " to " & Format(dtpTo.Value, "dd MMMM yyyy")
    End If
    
    If cmbPaymentMethod.Text <> "" Then
        temTopic = cmbPaymentMethod.Text & " - " & temSubTopic
    End If
    
    GetPrintDefaults ThisReportFormat
    ThisReportFormat.ReportPrintOrientation = Landscape
    
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    PrinterName = Printer.DeviceName
    
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    
    Dim MyPrinter As Printer
    
    For Each MyPrinter In Printers
        If MyPrinter.DeviceName = cmbPrinter.Text Then
            Set Printer = MyPrinter
        End If
    Next
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        GridPrint GridIncome, ThisReportFormat, temTopic, temSubTopic
        Printer.EndDoc
    End If

End Sub

Private Sub btnView_Click()
    If IsNumeric(txtBillID.Text) = True Then
        frmItemPurchaseCancel.Show
        frmItemPurchaseCancel.ZOrder 0
        frmItemPurchaseCancel.txtBillID.Text = Val(txtBillID.Text)
    Else
        MsgBox "Please select a bill"
    End If
End Sub

Private Sub cmbSupplier_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbSupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbPaymentMethod.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbSupplier.Text = Empty
    End If
End Sub

Private Sub cmbPaymentMethod_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub cmbPaymentMethod_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbPaymentMethod.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbPaymentMethod.Text = Empty
    End If
End Sub


Private Sub dtpFrom_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub dtpTo_Change()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_Load()
    Call SetColours
    Call FillCombos
    Call FormatGrid
    Call GetSettings
    Call FillGrid
End Sub

Private Sub FillCombos()
    Dim PM As New clsFillCombos
    PM.FillBoolCombo cmbPaymentMethod, "PaymentMethod", "PaymentMethod", "ToReceive", True
    Dim Supp As New clsFillCombos
    Supp.FillBoolCombo cmbSupplier, "Institution", "Institution", "IsSupplier", True
End Sub

Private Sub FormatGrid()
    With GridIncome
        .Clear
        .Cols = 6
        .Rows = 1
        
        .Row = 0
        
        .Col = 0
        .Text = "Bill ID"
        
        .Col = 1
        .Text = "Date"
        
        .Col = 2
        .Text = "Supplier"
        
        .Col = 3
        .Text = "Payment Method"
        
        .Col = 4
        .Text = "Value"
        
        .Col = 5
        .Text = "Comments"
        
    End With
End Sub

Private Sub FillGrid()
    Dim rsTem As New ADODB.Recordset
    Dim TotalPayments As Double
    Dim TotalReturn As Double
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblPurchaseBill.*, tblPaymentMethod.PaymentMethod, tblInstitution.Institution "
        temSql = temSql & "FROM (tblPaymentMethod RIGHT JOIN tblPurchaseBill ON tblPaymentMethod.PaymentMethodID = tblPurchaseBill.PaymentMethodID) LEFT JOIN tblInstitution ON tblPurchaseBill.SupplierID = tblInstitution.InstitutionID "
        temSql = temSql & "WHERE tblPurchaseBill.CompletedDate Between #" & Format(dtpFrom.Value, "dd MMMM yyyy") & "# And #" & Format(dtpTo.Value, "dd MMMM yyyy") & "# "
        If IsNumeric(cmbPaymentMethod.BoundText) = True Then
            temSql = temSql & "AND tblPurchaseBill.PaymentMethodID = " & Val(cmbPaymentMethod.BoundText) & " "
        End If
        If IsNumeric(cmbSupplier.BoundText) = True Then
            temSql = temSql & "And tblPurchaseBill.SupplierID = " & Val(cmbSupplier.BoundText) & " "
        End If
        temSql = temSql & "ORDER BY tblPurchaseBill.CompletedDateTime"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If IsNull(!NetTotal) = False Then
                TotalPayments = TotalPayments + !NetTotal
            End If
            
            GridIncome.Rows = GridIncome.Rows + 1
            GridIncome.Row = GridIncome.Rows - 1
            
            
            GridIncome.Col = 0
            GridIncome.Text = !PurchaseBillID
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            GridIncome.Col = 1
            If IsNull(!CompletedDate) = False Then
                GridIncome.Text = Format(!CompletedDate, LongDateFormat)
            End If
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            GridIncome.Col = 2
            If IsNull(!Institution) = False Then
                GridIncome.Text = !Institution
            End If
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            GridIncome.Col = 3
            If IsNull(!PaymentMethod) = False Then
                GridIncome.Text = !PaymentMethod
            End If
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            
            GridIncome.Col = 4
            GridIncome.Text = Format(!NetTotal, "0.00")
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            
            GridIncome.Col = 5
            If !cancelled Then
                GridIncome.Text = "Cancelled on " & Format(!CancelledDate, LongDateFormat) & " at " & !CancelledTime
                TotalReturn = TotalReturn + Format(!TotalReturnValue, "0.00")
            End If
'            If !Cancelled Then
'                GridIncome.Text = "Cancelled on " & Format(!CancelledDate, LongDateFormat) & " at " & !CancelledTime
'                TotalReturn = TotalReturn + Format(!TotalReturnValue, "0.00")
'            End If
            
            If GridIncome.Row Mod 2 = 0 Then
                GridIncome.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridIncome.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            .MoveNext
        Wend
    End With

    txtGrossTotal.Text = Format(TotalPayments, "0.00")
    txtRepayments.Text = Format(TotalReturn, "0.00")
    txtNetTotal.Text = Format(TotalPayments - TotalReturn, "0.00")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub GridIncome_Click()
    txtBillID.Text = GridIncome.TextMatrix(GridIncome.Row, 0)
    GridIncome.Col = GridIncome.Cols - 1
    GridIncome.ColSel = 0
End Sub

Private Sub GridIncome_DblClick()
    Call btnView_Click
End Sub

