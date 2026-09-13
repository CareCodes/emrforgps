VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmItemStock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Stock"
   ClientHeight    =   9090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11400
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
   ScaleHeight     =   9090
   ScaleWidth      =   11400
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   840
      TabIndex        =   14
      Top             =   8640
      Width           =   4575
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   840
      TabIndex        =   13
      Top             =   8160
      Width           =   4575
   End
   Begin btButtonEx.ButtonEx btnFill 
      Height          =   375
      Left            =   6600
      TabIndex        =   4
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Fill"
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
   Begin VB.Frame Frame1 
      Caption         =   "Order By"
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   7440
      Width           =   4215
      Begin VB.OptionButton optQuentity 
         Caption         =   "Quentity"
         Height          =   240
         Left            =   2040
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   9960
      TabIndex        =   12
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
   Begin btButtonEx.ButtonEx bttnPrint 
      Height          =   495
      Left            =   8640
      TabIndex        =   11
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
   Begin MSFlexGridLib.MSFlexGrid GridStock 
      Height          =   6255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11033
      _Version        =   393216
      AllowUserResizing=   1
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbDepartment 
      Height          =   360
      Left            =   1560
      TabIndex        =   3
      Top             =   720
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label29 
      Caption         =   "Paper"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label30 
      Caption         =   "Printer"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   8160
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "&Department"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "&Category"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   375
      Left            =   6720
      TabIndex        =   10
      Top             =   7560
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Total Value"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   7560
      Width           =   1575
   End
End
Attribute VB_Name = "frmItemStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim temSelect As String
    Dim temFrom As String
    Dim temWhere As String
    Dim temGroupBy As String
    Dim temOrderBY As String
    Dim temHaving As String
    Dim rsBatchStock As New ADODB.Recordset
    Dim temTopic As String
    Dim temSubTopic As String
    Dim TotalValue As Double
    Dim rsViewCategory As New ADODB.Recordset
    
    Dim CsetPrinter As New cSetDfltPrinter
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
    
    
    
Private Sub FormatGrid()
    With GridStock
        .Clear
        .Rows = 1
        .Cols = 3
        .Row = 0
        .Col = 0
        .Text = "Item"
        .CellAlignment = 4
        .Col = 1
        .Text = "Quentity"
        .CellAlignment = 4
        .Col = 2
        .Text = "Value"
        .CellAlignment = 4
        .ColWidth(0) = 5500
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
    End With
End Sub

Private Sub FillGrid()
    Screen.MousePointer = vbHourglass
    DoEvents
    With rsBatchStock
        If .State = 1 Then .Close
        
        temSelect = "SELECT tblItem.Item, Sum(tblItemBatchStock.Stock) AS SumOfStock, tblItem.PurchasePrice"
        
        temFrom = "FROM (tblItem LEFT JOIN tblItemBatch ON tblItem.ItemID = tblItemBatch.ItemID) LEFT JOIN tblItemBatchStock ON tblItemBatch.ItemBatchID = tblItemBatchStock.ItemBatchID"
        
        If IsNumeric(cmbCategory.BoundText) = False Then
            temWhere = " "
        Else
            temWhere = "WHERE (((tblItem.ItemCategoryID) = " & Val(cmbCategory.BoundText) & "))"
        End If
        
        
        If IsNumeric(cmbDepartment.BoundText) = False Then
            temGroupBy = "GROUP BY tblItem.Item, tblItem.PurchasePrice, tblItem.IsItemName"
            temHaving = "HAVING (((Sum(tblItemBatchStock.Stock))>0) AND ((tblItem.IsItemName)=True))"
        Else
            temGroupBy = "GROUP BY tblItem.Item, tblItem.PurchasePrice, tblItem.IsItemName, tblItemBatchStock.DepartmentID "
            temHaving = "HAVING (((Sum(tblItemBatchStock.Stock))>0) AND ((tblItem.IsItemName)=True) AND ((tblItemBatchStock.DepartmentID)=" & Val(cmbDepartment.BoundText) & "))"
        End If
        If optItem.Value = True Then
            temOrderBY = "ORDER BY tblItem.IsItemName"
        ElseIf optQuentity.Value = True Then
            temOrderBY = "ORDER BY Sum(tblItemBatchStock.Stock)"
        End If
        
        temSql = temSelect & " " & temFrom & " " & temWhere & " " & temGroupBy & " " & temHaving & " " & temOrderBY
        
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        i = 0
        TotalValue = 0
        GridStock.Visible = False
        While .EOF = False
            i = i + 1
            GridStock.Rows = GridStock.Rows + 1
            GridStock.Row = GridStock.Rows - 1
            
            GridStock.Col = 0
            GridStock.CellAlignment = 1
            If Not IsNull(!Item) Then GridStock.Text = !Item
            If GridStock.Row Mod 2 = 0 Then
                GridStock.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridStock.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            GridStock.Col = 1
            GridStock.CellAlignment = 7
            If Not IsNull(!SumOfStock) Then GridStock.Text = !SumOfStock
            If GridStock.Row Mod 2 = 0 Then
                GridStock.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridStock.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            GridStock.Col = 2
            GridStock.CellAlignment = 7
            If Not IsNull(!SumOfStock) And Not IsNull(!PurchasePrice) Then
                GridStock.Text = Format(!SumOfStock * !PurchasePrice, "#,##0.00")
                TotalValue = TotalValue + !SumOfStock * !PurchasePrice
            End If
            If GridStock.Row Mod 2 = 0 Then
                GridStock.CellBackColor = DefaultColourScheme.GridLightBackColour
            Else
                GridStock.CellBackColor = DefaultColourScheme.GridDarkBackColour
            End If
            
            .MoveNext
        Wend
        GridStock.Visible = True
    End With
    lblValue.Caption = Format(TotalValue, "#,##0.00")
    Screen.MousePointer = vbDefault
    DoEvents
End Sub

Private Sub btnFill_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnPrint_Click()
    Dim ThisReportFormat As PrintReport
    Dim temTopic As String
    Dim temSubTopic As String
    
    temTopic = InstitutionName & " - Item Stock"
    
    If cmbDepartment.Text <> "" And cmbCategory.Text <> "" Then
        temTopic = InstitutionName & " - " & cmbDepartment.Text
        temSubTopic = "Item Stock - " & cmbCategory.Text
    ElseIf cmbDepartment.Text <> "" And cmbCategory.Text = "" Then
        temTopic = InstitutionName & " - " & cmbDepartment.Text
        temSubTopic = "Item Stock"
    ElseIf cmbDepartment.Text = "" And cmbCategory.Text <> "" Then
        temTopic = InstitutionName
        temSubTopic = "Item Stock - " & cmbCategory.Text
    Else
        temTopic = InstitutionName
        temSubTopic = "Item Stock"
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
        GridPrint GridStock, ThisReportFormat, temTopic, temSubTopic
        Printer.EndDoc
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

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbCategory.Text = Empty
    End If
End Sub

Private Sub cmbDepartment_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = Empty
        cmbDepartment.Text = Empty
    End If
End Sub

Private Sub Form_Load()
    Call SetColours
    Call FillCombos
    Call FormatGrid
    Call FillGrid
    Call PopulatePrinters
    cmbPrinter_Click
    Call GetSettings
End Sub

Private Sub SaveSettings()
    SaveCommonSettings Me
    SaveSetting App.EXEName, Me.Name, cmbPrinter.Name, cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, cmbPaper.Name, cmbPaper.Text
End Sub

Private Sub GetSettings()
    GetCommonSettings Me
    On Error Resume Next
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, cmbPrinter.Name, "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, cmbPaper.Name, "A4")
End Sub

Private Sub FillCombos()
    With rsViewCategory
        If .State = 1 Then .Close
        temSql = "Select * from tblCategory Where Deleted = False and IsItemCategory = True order by Category"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCategory
        Set .RowSource = rsViewCategory
        .ListField = "Category"
        .BoundColumn = "CategoryID"
    End With
    Dim Dept As New clsFillCombos
    Dept.FillAnyCombo cmbDepartment, "Department", True
End Sub

Private Sub optAscending_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optDescending_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optExpiary_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveSettings
End Sub

Private Sub optItem_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optQuentity_Click()
    Call FormatGrid
    Call FillGrid
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
                'ComboPrescreptionPrinterPapers.AddItem FormItem
                cmbPaper.AddItem PtrCtoVbString(.pName)
'                ListPrescreptionPrinterPapers.AddItem PtrCtoVbString(.pName) & vbTab & .Size.cx / 1000 & " mm X " & .Size.cy / 1000 & " mm"
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

