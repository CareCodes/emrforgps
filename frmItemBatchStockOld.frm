VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmItemBatchStockOld 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Batch-vice Stock"
   ClientHeight    =   8880
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
   ScaleHeight     =   8880
   ScaleWidth      =   11400
   Begin VB.Frame Frame2 
      Caption         =   "Order By"
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   8160
      Width           =   4215
      Begin VB.OptionButton optAscending 
         Caption         =   "Ascending"
         Height          =   240
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optDescending 
         Caption         =   "Descinding"
         Height          =   240
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order By"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   7440
      Width           =   4215
      Begin VB.OptionButton optExpiary 
         Caption         =   "Expiary"
         Height          =   240
         Left            =   3000
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optValue 
         Caption         =   "Value"
         Height          =   240
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optQuentity 
         Caption         =   "Quentity"
         Height          =   240
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optItem 
         Caption         =   "Item"
         Height          =   240
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin btButtonEx.ButtonEx bttnClose 
      Height          =   495
      Left            =   9960
      TabIndex        =   2
      Top             =   8160
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
      TabIndex        =   1
      Top             =   8160
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
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11880
      _Version        =   393216
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   1560
      TabIndex        =   13
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label2 
      Caption         =   "Category"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblValue 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   7560
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "Total Value"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   7560
      Width           =   1575
   End
End
Attribute VB_Name = "frmItemBatchStockOld"
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
    Dim rsBatchStock As New ADODB.Recordset
    Dim temTopic As String
    Dim temSubTopic As String
    Dim i As Integer
    Dim TotalValue As Double
    Dim rsViewCategory As New ADODB.Recordset
    
Private Sub FormatGrid()
    With GridStock
        .Clear
        .Rows = 1
        .Cols = 5
        .Row = 0
        .Col = 0
        .Text = "Item"
        .CellAlignment = 4
        .Col = 1
        .Text = "Batch"
        .CellAlignment = 4
        .Col = 2
        .Text = "Expiary"
        .CellAlignment = 4
        .Col = 3
        .Text = "Quentity"
        .CellAlignment = 4
        .Col = 4
        .Text = "Value"
        .CellAlignment = 4
        .ColWidth(0) = 5500
        .ColWidth(1) = 1100
        .ColWidth(2) = 1400
        .ColWidth(3) = 1100
        .ColWidth(4) = 1400
    
    End With
End Sub

Private Sub FillGrid()
    Screen.MousePointer = vbHourglass
    DoEvents
    With rsBatchStock
        If .State = 1 Then .Close
        temSelect = "SELECT tblItem.Item, tblItemBatch.ItemBatch, tblItemBatchStock.Stock, [tblItem].[PurchasePrice]*[tblItemBatchStock].[Stock] AS StockValue, tblItemBatch.DOE"
        temFrom = "FROM (tblItem LEFT JOIN tblItemBatch ON tblItem.ItemID = tblItemBatch.ItemID) LEFT JOIN tblItemBatchStock ON tblItemBatch.ItemBatchID = tblItemBatchStock.ItemBatchID"
        If IsNumeric(cmbCategory.BoundText) = False Then
            temWhere = "WHERE (((tblItemBatchStock.Stock)>0) AND ((tblItem.IsItemName)=True))"
        Else
            temWhere = "WHERE (((tblItemBatchStock.Stock)>0) AND ((tblItem.IsItemName)=True) AND ((tblItem.ItemCategoryID)=" & Val(cmbCategory.BoundText) & "))"
        End If
        If optItem.Value = True Then
            temOrderBY = "ORDER BY tblItem.Item"
        ElseIf optValue.Value = True Then
            temOrderBY = "ORDER BY [tblItem].[PurchasePrice]*[tblItemBatchStock].[Stock]"
        ElseIf optQuentity.Value = True Then
            temOrderBY = "ORDER BY [tblItemBatchStock].[Stock]"
        ElseIf optExpiary.Value = True Then
            temOrderBY = "ORDER BY tblItemBatch.DOE"
        End If
        If optDescending.Value = True Then temOrderBY = temOrderBY & " DESC"
        temSql = temSelect & " " & temFrom & " " & temWhere & " " & temOrderBY
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        i = 0
        TotalValue = 0
        If .RecordCount > 0 Then
            .MoveLast
            GridStock.Rows = .RecordCount + 1
            .MoveFirst
            While .EOF = False
                i = i + 1
                If Not IsNull(!Item) Then GridStock.TextMatrix(i, 0) = !Item
                If Not IsNull(!ItemBatch) Then GridStock.TextMatrix(i, 1) = !ItemBatch
                If Not IsNull(!DOE) Then GridStock.TextMatrix(i, 2) = Format(!DOE, "MMMM yyyy")
                If Not IsNull(!Stock) Then GridStock.TextMatrix(i, 3) = !Stock
                If Not IsNull(!StockValue) Then GridStock.TextMatrix(i, 4) = Format(!StockValue, "#,##0.00")
                If Not IsNull(!StockValue) Then TotalValue = TotalValue + !StockValue
                .MoveNext
            Wend
        End If
    End With
    lblValue.Caption = Format(TotalValue, "#,##0.00")
    Screen.MousePointer = vbDefault
    DoEvents
End Sub

Private Sub bttnClose_Click()
    Unload Me
End Sub

Private Sub bttnPrint_Click()
    Dim RetVal As Integer
    Dim TemResponce As Integer
    Dim CsetPrinter As New cSetDfltPrinter
    CsetPrinter.SetPrinterAsDefault (ReportPrinterName)
    If SelectForm(ReportPaperName, Me.hwnd) = 1 Then


    End If
End Sub

Private Sub cmbCategory_Change()
    Call FormatGrid
    Call FillGrid
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
    GetCommonSettings Me
    
    Call SetColours
    Call FillCombos
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub FillCombos()
    With rsViewCategory
        If .State = 1 Then .Close
        temSql = "Select * from tblCategory Where Deleted = False AND IsItemCategory = True order by Category"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCategory
        Set .RowSource = rsViewCategory
        .ListField = "Category"
        .BoundColumn = "CategoryID"
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

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

Private Sub optItem_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optQuentity_Click()
    Call FormatGrid
    Call FillGrid
End Sub

Private Sub optValue_Click()
    Call FormatGrid
    Call FillGrid
End Sub
