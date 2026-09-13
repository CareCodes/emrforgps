VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPharmacySaleView 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "View Pharmacy Sale"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
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
   ScaleHeight     =   7305
   ScaleWidth      =   14760
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid GridItem 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10610
      _Version        =   393216
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
   Begin VB.TextBox txtPaid 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   11760
      TabIndex        =   15
      Top             =   5760
      Width           =   2775
   End
   Begin VB.TextBox txtBalance 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   11760
      TabIndex        =   17
      Top             =   6240
      Width           =   2775
   End
   Begin VB.TextBox txtRunningSaleM 
      Height          =   405
      Left            =   10320
      TabIndex        =   38
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtRunningPurchaseM 
      Height          =   405
      Left            =   11040
      TabIndex        =   37
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtComments 
      Enabled         =   0   'False
      Height          =   2295
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
   Begin VB.TextBox txtGrossTotal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   11760
      TabIndex        =   9
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox txtNetTotal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   11760
      TabIndex        =   13
      Top             =   5280
      Width           =   2775
   End
   Begin VB.TextBox txtRunningPurchaseY 
      Height          =   405
      Left            =   13200
      TabIndex        =   36
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtRunningSaleY 
      Height          =   405
      Left            =   12480
      TabIndex        =   35
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtTotalCost 
      Height          =   405
      Left            =   11760
      TabIndex        =   34
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCostRate 
      Height          =   405
      Left            =   11760
      TabIndex        =   33
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCostValue 
      Height          =   405
      Left            =   12480
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtEditID 
      Height          =   405
      Left            =   10320
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDeleteID 
      Height          =   405
      Left            =   11040
      TabIndex        =   30
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtBillID 
      Height          =   405
      Left            =   10800
      TabIndex        =   29
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtPaymentComments 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   11760
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3000
      Width           =   2775
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
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
      Left            =   9840
      TabIndex        =   18
      Top             =   6840
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   11760
      TabIndex        =   11
      Top             =   4800
      Width           =   2775
   End
   Begin VB.ComboBox cmbPaper 
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   6720
      Width           =   3255
   End
   Begin VB.ComboBox cmbPrinter 
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
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   6240
      Width           =   3255
   End
   Begin btButtonEx.ButtonEx btnPurchase 
      Height          =   375
      Left            =   11400
      TabIndex        =   19
      Top             =   6720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   375
      Left            =   13080
      TabIndex        =   20
      Top             =   6720
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
      Height          =   405
      Left            =   1440
      TabIndex        =   22
      Top             =   6240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   714
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbReceived 
      Height          =   405
      Left            =   1440
      TabIndex        =   24
      Top             =   6720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   714
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSaleCategory 
      Height          =   405
      Left            =   11760
      TabIndex        =   3
      Top             =   2520
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   714
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbCustomer 
      Height          =   405
      Left            =   11760
      TabIndex        =   7
      Top             =   3840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   714
      _Version        =   393216
      Enabled         =   0   'False
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtQueueID 
      Height          =   405
      Left            =   13920
      TabIndex        =   39
      Top             =   6240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "&Cash"
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
      Left            =   9720
      TabIndex        =   14
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Balance"
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
      Left            =   9720
      TabIndex        =   16
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label11 
      Caption         =   "Customer"
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
      Left            =   9720
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label23 
      Caption         =   "Sale Type"
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
      Left            =   9720
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Paymen&t Details"
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
      Left            =   9720
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Gross Total"
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
      Left            =   9720
      TabIndex        =   8
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Net Total"
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
      Left            =   9720
      TabIndex        =   12
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "&Received by"
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
      Left            =   120
      TabIndex        =   23
      Top             =   6720
      Width           =   1455
   End
   Begin VB.Label Label21 
      Caption         =   "C&hecked by"
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
      Left            =   120
      TabIndex        =   21
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "&Discount"
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
      Left            =   9720
      TabIndex        =   10
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Paper"
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
      Left            =   4920
      TabIndex        =   27
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "&Printer"
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
      Left            =   4920
      TabIndex        =   25
      Top             =   6240
      Width           =   1575
   End
End
Attribute VB_Name = "frmPharmacySaleView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim CsetPrinter As New cSetDfltPrinter
    Dim SaleItem As New clsItem
    Dim ThisSale As New clsSaleCategory
    Dim SaleBill As New clsSaleBill
    Dim SaleCategory As New clsSaleCategory
    Dim tr As Long
    
    Dim rsBatch As New ADODB.Recordset
    Dim Customer As New clsFillCombos
    
    Dim FocusControl As Control
    
    Dim NumForms As Long, i As Long
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
    
    
Private Sub FillCombos()
    Dim Checked As New clsFillCombos
    Checked.FillBoolCombo cmbChecked, "Staff", "Staff", "IsAUser", True
    Dim Received As New clsFillCombos
    Received.FillBoolCombo cmbReceived, "Staff", "Staff", "IsAUser", True
    Dim SCat As New clsFillCombos
    SCat.FillAnyCombo cmbSaleCategory, "SaleCategory", True
End Sub


Private Sub GetSettings()
    On Error Resume Next
    GetCommonSettings Me
    
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, "Printer", "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, "Paper", "")
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, "Printer", cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, "Paper", cmbPaper.Text
    SaveCommonSettings Me
    
End Sub

Private Sub FormatGrid()
    With GridItem
        .Clear
        
        .Rows = 1
        .Cols = 6
        
        .Row = 0
        
        .Col = 0
        .Text = "No."
        
        .Col = 1
        .Text = "Item"
        
        .Col = 2
        .Text = "Quentity"
        
        .Col = 3
        .Text = "Rate"
        
        .Col = 4
        .Text = "Value"
        
        .Col = 5
        .Text = "ID"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 2800
        .ColWidth(2) = 1800
        .ColWidth(3) = 1800
        .ColWidth(4) = 1800
        .ColWidth(5) = 0
        
    End With
End Sub

Private Sub FillGrid()
    Dim rsTem As New ADODB.Recordset
    Dim TotalValue As Double
    Dim TotalCost As Double
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblItem.Item, tblItem.ItemCode, tblSaleItem.PurchaseValue, tblSaleItem.SaleQuentity & ' ' &  tblItemUnit.ItemUnit as IssueQuentity, tblSaleItem.SaleRate, tblSaleItem.SaleValue, tblSaleItem.SaleItemID FROM (tblSaleItem LEFT JOIN tblItem ON tblSaleItem.ItemID = tblItem.ItemID) LEFT JOIN tblItemUnit ON tblItem.IssueUnitID = tblItemUnit.ItemUnitID WHERE (((tblSaleItem.SaleBillID)=" & Val(txtBillID.Text) & ") AND ((tblSaleItem.Deleted) = False))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            GridItem.Rows = GridItem.Rows + 1
            GridItem.Row = GridItem.Rows - 1
            GridItem.Col = 0
            GridItem.CellAlignment = 1
            GridItem.Text = GridItem.Row
            GridItem.Col = 1
            GridItem.CellAlignment = 1
            If !ItemCode <> "" Then
                GridItem.Text = !Item & "(" & !ItemCode & ")"
            Else
                GridItem.Text = !Item
            End If
            GridItem.Col = 2
            GridItem.CellAlignment = 7
            GridItem.Text = !IssueQuentity
            GridItem.Col = 3
            GridItem.CellAlignment = 7
            GridItem.Text = Format(!SaleRate, "0.00")
            
            GridItem.Col = 4
            
           
            GridItem.CellAlignment = 7
            GridItem.Text = Format(!SaleValue, "0.00")
            
            TotalValue = TotalValue + !SaleValue
            TotalCost = TotalCost + !PurchaseValue
            
            GridItem.Col = 5
'            GridItem.CellAlignment = 4
            GridItem.Text = !SaleItemID
            
            .MoveNext
        Wend
        .Close
    End With
    txtGrossTotal.Text = Format(TotalValue, "0.00")
    txtTotalCost.Text = TotalCost
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


Private Sub btnPurchase_Click()
    Call PrintBill
End Sub

Private Sub PrintBill()
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    Dim MyFOnt As ReportFont

    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    
    Dim CenterX As Long
    Dim FieldX As Long
    Dim SubFieldX As Long
    Dim ValueX As Long
    Dim SubValueX As Long
    Dim AllLines() As String
    Dim i As Integer
    Dim temY As Long
    
    
    
    With MyFOnt
        .Name = "Arial"
        .Bold = False
        .Italic = False
        .Size = 12
        .Italic = False
        .Underline = False
        .Strikethrough = False
    End With
    
    If SelectForm(cmbPaper.Text, Me.hwnd) = 1 Then
        CenterX = Printer.Width / 2
        FieldX = (1440) * 0.3
        SubFieldX = (1440) * 2
        ValueX = Printer.Width - (1440) * 0.5
        SubValueX = Printer.Width - 1440 * 1.2
        
        MyFOnt.Bold = False
        MyFOnt.Size = 11
        PrintingText 0, 1440 * 1, Printer.Width, 0, InstitutionName, CentreAlign, MyFOnt
        
        MyFOnt.Bold = False
        MyFOnt.Size = 10
        PrintingText 0, 0, Printer.Width, 0, InstitutionAddress, CentreAlign, MyFOnt
        PrintingText 0, 0, Printer.Width, 0, "Pharmacy BILL - " & cmbCustomer.Text, CentreAlign, MyFOnt
        
        
        MyFOnt.Size = 8
        
        Printer.Print
        Printer.Print

        For i = 0 To GridItem.Rows - 1
            temY = Printer.CurrentY
            PrintingText FieldX, temY, ValueX, 0, GridItem.TextMatrix(i, 1), LeftAlign, MyFOnt
            PrintingText SubFieldX, temY, ValueX, 0, GridItem.TextMatrix(i, 2), LeftAlign, MyFOnt
            PrintingText SubValueX, temY, SubValueX, 0, GridItem.TextMatrix(i, 3), RightAlign, MyFOnt
            PrintingText ValueX, temY, ValueX, 0, GridItem.TextMatrix(i, 4), RightAlign, MyFOnt
        Next
        
        Printer.Print
        
        temY = Printer.CurrentY
        PrintingText FieldX, temY, FieldX, 0, "Total", LeftAlign, MyFOnt
        PrintingText FieldX, temY, ValueX, 0, txtGrossTotal.Text, RightAlign, MyFOnt
        
        If Val(txtDiscount.Text) > 0 Then
            
            temY = Printer.CurrentY
            PrintingText FieldX, temY, FieldX, 0, "Discount", LeftAlign, MyFOnt
            PrintingText FieldX, temY, ValueX, 0, txtDiscount.Text, RightAlign, MyFOnt
            
            temY = Printer.CurrentY
            PrintingText FieldX, temY, FieldX, 0, "Net Total", LeftAlign, MyFOnt
            PrintingText FieldX, temY, ValueX, 0, txtNetTotal.Text, RightAlign, MyFOnt
            
            
        End If
        
        
        Printer.EndDoc
        
    End If

End Sub


Private Sub PrintingText(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, PrintText As String, PrintAlignment As TextAlignment, ReportPrintFont As ReportFont)
    
    If PrintAlignment = LeftAlign Then
        Printer.CurrentX = X1
    ElseIf PrintAlignment = RightAlign Then
        Printer.CurrentX = X2 - Printer.TextWidth(PrintText)
    ElseIf PrintAlignment = CentreAlign Then
        Printer.CurrentX = (X1 + X2 / 2) - (Printer.TextWidth(PrintText) / 2)
    Else
        Printer.CurrentX = X1
    End If
    If Y1 <> 0 Then Printer.CurrentY = Y1
    Printer.Font.Name = ReportPrintFont.Name
    Printer.Font.Size = ReportPrintFont.Size
    Printer.Font.Italic = ReportPrintFont.Italic
    Printer.Font.Bold = ReportPrintFont.Bold
    Printer.Font.Underline = ReportPrintFont.Underline
    
    Printer.Print PrintText
End Sub


Private Sub DisplaySaleBill()
    Dim rsTemBill As New ADODB.Recordset
    With rsTemBill
        If .State = 1 Then .Close
        temSql = "Select * from tblSaleBill where SaleBillID = " & Val(txtBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbSaleCategory.BoundText = !SaleCategoryID
            txtGrossTotal.Text = !GrossTotal
            txtDiscount.Text = !Discount
            txtNetTotal.Text = !NetTotal
            cmbCustomer.BoundText = !CustomerID
            txtComments.Text = Format(!Comments, "")
        End If
        .Close
    End With
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
    Call PopulatePrinters
    Call SetColours
    Call FillCombos
    Call FormatGrid
    Call GetSettings
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub


Private Sub PopulatePrinters()
    Dim MyPrinter As Printer
    For Each MyPrinter In Printers
        cmbPrinter.AddItem MyPrinter.DeviceName
    Next
End Sub

Private Sub PopulatePapers()
    cmbPaper.Clear
    SetPrinter = False
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
'        With FormSize
'            .cx = PrescreptionPaperHeight
'            .cy = PrescreptionPaperWidth
'        End With
        ReDim aFI1(1)
        RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
        ReDim Temp(BytesNeeded)
        ReDim aFI1(BytesNeeded / Len(FI1))
        RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, NumForms)
        Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
        For i = 0 To NumForms - 1
            With aFI1(i)
                'FormItem = PtrCtoVbString(.pName) & " - " & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm   (" & i + 1 & ")"
                'ComboBillPrinterPapers.AddItem FormItem
                cmbPaper.AddItem PtrCtoVbString(.pName)
'                ListBillPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
            End With
        Next i
        ClosePrinter (PrinterHandle)
    End If
End Sub

Private Sub cmbPrinter_Change()
    Call PopulatePapers
End Sub

Private Sub cmbPrinter_Click()
    Call PopulatePapers
End Sub

Private Sub txtBillID_Change()
    Call FormatGrid
    Call FillGrid
    Call DisplaySaleBill
End Sub
