VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmIndoorIssue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indoor Issue"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14760
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
   ScaleHeight     =   8490
   ScaleWidth      =   14760
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtQueueID 
      Height          =   375
      Left            =   10320
      TabIndex        =   41
      Top             =   5400
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   10920
      Top             =   6960
   End
   Begin VB.TextBox txtRunningSaleM 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11160
      TabIndex        =   40
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtRunningPurchaseM 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11880
      TabIndex        =   39
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtComments 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   9600
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   1200
      Width           =   4935
   End
   Begin VB.TextBox txtRunningPurchaseY 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   13920
      TabIndex        =   38
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtRunningSaleY 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   13320
      TabIndex        =   37
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtTotalCost 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12600
      TabIndex        =   36
      Top             =   5400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCostRate 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11760
      TabIndex        =   35
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtCostValue 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12480
      TabIndex        =   34
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtEditID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   10320
      TabIndex        =   33
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtDeleteID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   11040
      TabIndex        =   32
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtBillID 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9600
      TabIndex        =   31
      Top             =   2520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CheckBox chkPrint 
      Caption         =   "Print"
      Height          =   255
      Left            =   9960
      TabIndex        =   19
      Top             =   7920
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cmbPaper 
      Height          =   360
      Left            =   11280
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   4920
      Width           =   3255
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   360
      Left            =   11280
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   4440
      Width           =   3255
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   7920
      TabIndex        =   12
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSDataListLib.DataList lstBatch 
      Height          =   915
      Left            =   9600
      TabIndex        =   15
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   1614
      _Version        =   393216
      IntegralHeight  =   0   'False
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
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtRate 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
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
   Begin MSDataListLib.DataCombo cmbItem 
      Height          =   360
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   4335
      _ExtentX        =   7646
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
   Begin MSDataListLib.DataCombo cmbCode 
      Height          =   360
      Left            =   1320
      TabIndex        =   5
      Top             =   1080
      Width           =   4335
      _ExtentX        =   7646
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
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   375
      Left            =   7920
      TabIndex        =   14
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
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
   Begin MSFlexGridLib.MSFlexGrid GridItem 
      Height          =   6735
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11880
      _Version        =   393216
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin btButtonEx.ButtonEx btnPurchase 
      Height          =   375
      Left            =   11400
      TabIndex        =   20
      Top             =   7920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Appearance      =   3
      BorderColor     =   16711680
      Caption         =   "&Issue"
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
      TabIndex        =   21
      Top             =   7920
      Width           =   1455
      _ExtentX        =   2566
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
      Left            =   11280
      TabIndex        =   23
      Top             =   3480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
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
   Begin MSDataListLib.DataCombo cmbReceived 
      Height          =   360
      Left            =   11280
      TabIndex        =   25
      Top             =   3960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
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
   Begin MSDataListLib.DataCombo cmbCustomer 
      Height          =   360
      Left            =   11280
      TabIndex        =   18
      Top             =   3000
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   635
      _Version        =   393216
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
   Begin EmrForGps.UCPopup ucpPrescreption 
      Left            =   9960
      Top             =   7320
      _ExtentX        =   318
      _ExtentY        =   370
   End
   Begin VB.Label Label11 
      Caption         =   "Customer"
      Height          =   255
      Left            =   9600
      TabIndex        =   17
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "&Received by"
      Height          =   255
      Left            =   9600
      TabIndex        =   24
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label21 
      Caption         =   "C&hecked by"
      Height          =   255
      Left            =   9600
      TabIndex        =   22
      Top             =   3600
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Paper"
      Height          =   255
      Left            =   9600
      TabIndex        =   28
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "&Printer"
      Height          =   255
      Left            =   9600
      TabIndex        =   26
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label lblIssueUnit 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   30
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label4 
      Caption         =   "Price"
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Rate"
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "&Quantity"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label28 
      Caption         =   "Cato&gery"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label34 
      Caption         =   "&Item"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label35 
      Caption         =   "C&ode"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
End
Attribute VB_Name = "frmIndoorIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim CsetPrinter As New cSetDfltPrinter
    Dim SaleItem As New clsItem
    Dim ThisSale As New clsSaleCategory
    Dim IndoorBill As New clsIndoorBill
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
    
    Dim ToRight As Boolean
    
    Dim AlternateItem() As String
    Dim AlternateItemID() As Long
    Dim AlternateItemCount As Long
    Dim ItemRow As Long
    Dim TopRow As Long
    
    
Private Sub FillCombos()
    Dim ICat As New clsFillCombos
    ICat.FillBoolCombo cmbCategory, "Category", "Category", "IsItemCategory", True
    Dim Checked As New clsFillCombos
    Checked.FillBoolCombo cmbChecked, "Staff", "Staff", "IsAUser", True
    Dim Received As New clsFillCombos
    Received.FillBoolCombo cmbReceived, "Staff", "Staff", "IsAUser", True
    Dim ItemName As New clsFillCombos
    Dim ItemCode As New clsFillCombos
    ItemName.FillBoolCombo cmbItem, "Item", "Item", "IsItemName", True
    ItemCode.FillSpecificFieldBoolCombo cmbCode, "Item", "ItemCode", "ItemCode", "IsItemName", True
    Dim Pt As New clsFillCombos
    Pt.FillAnyCombo cmbCustomer, "Patient", True
End Sub

Private Sub ClearAddValues()
    cmbCategory.Text = Empty
    cmbItem.Text = Empty
    cmbCode.Text = Empty
    txtQty.Text = Empty
    txtRate.Text = Empty
    txtPrice.Text = Empty
    txtCostRate.Text = Empty
    txtCostValue.Text = Empty
    txtDeleteID.Text = Empty
    txtEditID.Text = Empty
    Call ClearBatches
End Sub

Private Sub ClearBatches()
    Dim temID As Long
    With rsBatch
        If .State = 1 Then .Close
    End With
    With lstBatch
        Set .RowSource = Nothing
        .BoundColumn = Empty
        .ListField = Empty
        .BoundText = Empty
    End With
End Sub

Private Sub ClearBillValues()
    txtBillID.Text = Empty
    txtRunningPurchaseY.Text = Empty
    txtRunningPurchaseM.Text = Empty
    txtRunningSaleY.Text = Empty
    txtRunningSaleM.Text = Empty
    txtTotalCost.Text = Empty
    txtComments.Text = Empty
    txtDeleteID.Text = Empty
    txtEditID.Text = Empty
    txtQueueID.Text = Empty
    cmbCustomer.Text = Empty
End Sub

Private Sub GetSettings()
    ToRight = True
    
    On Error Resume Next
    
    cmbChecked.BoundText = UserID
    cmbReceived.BoundText = UserID
    
    GetCommonSettings Me
    
    cmbChecked.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbChecked.Name, UserID))
    cmbReceived.BoundText = Val(GetSetting(App.EXEName, Me.Name, cmbReceived.Name, UserID))
    chkPrint.Value = Val(GetSetting(App.EXEName, Me.Name, chkPrint.Name, 1))
    cmbPrinter.Text = GetSetting(App.EXEName, Me.Name, "Printer", "")
    cmbPaper.Text = GetSetting(App.EXEName, Me.Name, "Paper", "")
    Dim temText As String
    temText = GetSetting(App.EXEName, Me.Name, "FocusControl", cmbCategory.Name)
    Dim MyCtrl As Control
    For Each MyCtrl In Controls
        If MyCtrl.Name = temText Then
            MyCtrl.TabIndex = 1
            Set FocusControl = MyCtrl
        End If
    Next
End Sub

Private Sub SaveSettings()
    SaveSetting App.EXEName, Me.Name, "Printer", cmbPrinter.Text
    SaveSetting App.EXEName, Me.Name, "Paper", cmbPaper.Text
    SaveSetting App.EXEName, Me.Name, cmbReceived.Name, cmbReceived.BoundText
    SaveSetting App.EXEName, Me.Name, cmbChecked.Name, cmbChecked.BoundText
    SaveSetting App.EXEName, Me.Name, chkPrint.Name, chkPrint.Value
    SaveCommonSettings Me
    
    SaveSetting App.EXEName, Me.Name, "FocusControl", FocusControl.Name
End Sub

Private Sub ListBatches(ItemID As Long)
    Dim rsBatch As New ADODB.Recordset
    Dim temBatchID As Long
    With rsBatch
        If .State = 1 Then .Close
        temSql = "SELECT tblItemBatch.ItemBatch & ' ' &  tblItemBatch.DOE & ' (' &  tblItemBatchStock.Stock & ')' as BatchDetails , tblItemBatch.ItemBatchID " & _
                    "FROM tblItemBatch LEFT JOIN tblItemBatchStock ON tblItemBatch.ItemBatchID = tblItemBatchStock.ItemBatchID " & _
                    "Where (((tblItemBatchStock.DepartmentID)=" & UserDepartmentID & ") AND ((tblItemBatch.ItemID) = " & ItemID & ")) " & _
                    "ORDER BY tblItemBatch.DOE"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temBatchID = !ItemBatchID
        Else
            temBatchID = 0
        End If
    End With
    With lstBatch
        Set .RowSource = rsBatch
        .ListField = "BatchDetails"
        .BoundColumn = "ItemBatchID"
        .BoundText = temBatchID
    End With
End Sub

Private Sub DisplayRate()
    txtRate.Text = SaleItem.SalePrice
    txtCostRate.Text = SaleItem.PurchasePrice
End Sub

Private Sub CalculateValues()
    txtPrice.Text = Format(Val(txtRate.Text) * Val(txtQty.Text), "0.00")
    txtCostValue.Text = Val(txtCostRate.Text) * Val(txtQty.Text)
End Sub

Private Sub FormatGrid()
    With GridItem
        .Clear
        
        .Rows = 1
        .Cols = 6
        
        .row = 0
        
        .col = 0
        .Text = "No."
        
        .col = 1
        .Text = "Item"
        
        .col = 2
        .Text = "Quentity"
        
        .col = 3
        .Text = "Rate"
        
        .col = 4
        .Text = "Value"
        
        .col = 5
        .Text = "ID"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 3800
        .ColWidth(2) = 1600
        .ColWidth(3) = 1600
        .ColWidth(4) = 1600
        .ColWidth(5) = 0
        
    End With
End Sub

Private Sub FillGrid()
    Dim rsTem As New ADODB.Recordset
    Dim TotalValue As Double
    Dim TotalCost As Double
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblItem.Item, tblItem.ItemCode, tblSaleItem.PurchaseValue, tblSaleItem.SaleQuentity & ' ' &  tblItemUnit.ItemUnit as IssueQuentity, tblSaleItem.SaleRate, tblSaleItem.SaleValue, tblSaleItem.SaleItemID FROM (tblSaleItem LEFT JOIN tblItem ON tblSaleItem.ItemID = tblItem.ItemID) LEFT JOIN tblItemUnit ON tblItem.IssueUnitID = tblItemUnit.ItemUnitID WHERE (((tblSaleItem.IndoorBillID)=" & Val(txtBillID.Text) & ") AND ((tblSaleItem.Deleted) = False))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            GridItem.Rows = GridItem.Rows + 1
            GridItem.row = GridItem.Rows - 1
            GridItem.col = 0
            GridItem.CellAlignment = 1
            GridItem.Text = GridItem.row
            GridItem.col = 1
            GridItem.CellAlignment = 1
            If !ItemCode <> "" Then
                GridItem.Text = !Item & "   (" & !ItemCode & ")"
            Else
                GridItem.Text = !Item
            End If
            GridItem.col = 2
            GridItem.CellAlignment = 7
            GridItem.Text = !IssueQuentity
            GridItem.col = 3
            GridItem.CellAlignment = 7
            GridItem.Text = Format(!SaleRate, "0.00")
            
            GridItem.col = 4
            
           
            GridItem.CellAlignment = 7
            GridItem.Text = Format(!SaleValue, "0.00")
            
            TotalValue = TotalValue + !SaleValue
            TotalCost = TotalCost + !PurchaseValue
            
            GridItem.col = 5
'            GridItem.CellAlignment = 4
            GridItem.Text = !SaleItemID
            
            .MoveNext
        Wend
        .Close
    End With
    
 End Sub

Public Sub btnAdd_Click()
    Dim AddStock As Boolean
    Dim MyTemID As Long
    If IsNumeric(cmbItem.BoundText) = False Then
        tr = MsgBox("You have not entered the item to add", vbCritical, "Item?")
        FocusControl.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtQty.Text) = False Or Val(txtQty.Text) = 0 Then
        tr = MsgBox("You have not entered the quentity", vbCritical, "Quentity?")
        txtQty.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtRate.Text) = False Or Val(txtRate.Text) = 0 Then
        tr = MsgBox("You have not entered the rate", vbCritical, "Rate")
        txtRate.SetFocus
        On Error Resume Next: SendKeys "{home}+{end}"
        Exit Sub
    End If
    If IsNumeric(lstBatch.BoundText) = False Then
        tr = MsgBox("Please select a batch", vbCritical, "Batch?")
        lstBatch.SetFocus
        Exit Sub
    End If
    
    While CalculateStock(cmbItem.BoundText, lstBatch.BoundText, UserDepartmentID).Stock < Val(txtQty.Text)
        If CalculateStock(cmbItem.BoundText, , UserDepartmentID).Stock > Val(txtQty.Text) Then
            Call FillBatchesByStock
            lstBatch.SetFocus
            On Error Resume Next
            SendKeys "{home}"
            btnAdd_Click
            Exit Sub
        Else
            MsgBox "No Adequate stocks"
            FocusControl.SetFocus
            Exit Sub
        End If
    Wend

    If IsNumeric(txtBillID.Text) = False Then
        txtBillID.Text = IndoorBill.NewIndoorBillID
    End If

    Dim rsSaleItem As New ADODB.Recordset
    With rsSaleItem
        If .State = 1 Then .Close
        temSql = "Select * from tblSaleItem where SaleItemID = " & Val(txtEditID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If !ItemID = cmbItem.BoundText And !ItemBatchID = lstBatch.BoundText Then
                If Val(txtQty.Text) > !SaleQuentity Then
                    If CalculateStock(cmbItem.BoundText, lstBatch.BoundText, UserDepartmentID).Stock < Val(txtQty.Text) - !SaleQuentity Then
                        tr = MsgBox("There are no adequate stocks for that batch", vbCritical, "No Stocks")
                        lstBatch.SetFocus
                        .Close
                        Exit Sub
                    End If
                End If
                AddStock = ChangeStocks(cmbItem.BoundText, lstBatch.BoundText, txtQty.Text - !SaleQuentity, UserDepartmentID, False, True, True)
            Else
                If CalculateStock(cmbItem.BoundText, lstBatch.BoundText, UserDepartmentID).Stock < Val(txtQty.Text) Then
                    tr = MsgBox("There are no adequate stocks for that batch", vbCritical, "No Stocks")
                    lstBatch.SetFocus
                    .Close
                    Exit Sub
                End If
                AddStock = ChangeStocks(cmbItem.BoundText, lstBatch.BoundText, txtQty.Text, UserDepartmentID, False, True, True)
                AddStock = ChangeStocks(!ItemID, !ItemBatchID, !SaleQuentity, UserDepartmentID, True, False, True)
            End If
        
        Else
            If CalculateStock(cmbItem.BoundText, lstBatch.BoundText, UserDepartmentID).Stock < Val(txtQty.Text) Then
                tr = MsgBox("There are no adequate stocks for that batch", vbCritical, "No Stocks")
                lstBatch.SetFocus
                .Close
                Exit Sub
            End If
            .AddNew
            AddStock = ChangeStocks(cmbItem.BoundText, lstBatch.BoundText, txtQty.Text, UserDepartmentID, False, True, True)
        End If
        !IndoorBillID = Val(txtBillID.Text)
        !ItemID = Val(cmbItem.BoundText)
        !ItemBatchID = Val(lstBatch.BoundText)
        !SaleRate = Val(txtRate.Text)
        !SaleQuentity = Val(txtQty.Text)
        !SaleFreeQuentity = 0
        !SaleValue = Val(txtPrice.Text)
        !FreeSaleValue = 0
        !TotalSaleValue = Val(txtPrice.Text)
        !PurchaseRate = Val(txtCostRate.Text)
        !PurchaseValue = Val(txtCostValue.Text)
        !PurchaseFreeValue = 0
        !TotalPurchaseValue = Val(txtCostValue.Text)
        .Update
    End With
        Call ClearAddValues
        Call FormatGrid
        Call FillGrid
        FocusControl.SetFocus

End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim temBool As Boolean
    Dim rsTem As New ADODB.Recordset
    With rsTem
        temSql = "Select * from tblSaleItem Where SaleItemID = " & Val(txtDeleteID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedTime = Now
            !DeletedUserID = UserID
            temBool = ChangeStocks(!ItemID, !ItemBatchID, !SaleQuentity, UserDepartmentID, True, False, False)
            .Update
        End If
    End With
    Call FormatGrid
    Call FillGrid
    Call ClearAddValues
    FocusControl.SetFocus
End Sub

Private Sub btnPurchase_Click()
    
    If CanSettle = False Then Exit Sub
    
    Call WriteIndoorBill
    If chkPrint.Value = 1 Then Call PrintBill
    
    Call ClearAddValues
    Call ClearBillValues
    Call FormatGrid
    
    txtBillID.Text = IndoorBill.NewIndoorBillID

    FocusControl.SetFocus
    

End Sub

Private Sub PrintBill()
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    Dim MyFOnt As ReportFont

    PrinterName = Printer.DeviceName
    If OpenPrinter(PrinterName, PrinterHandle, 0&) Then
        ClosePrinter (PrinterHandle)
    End If
    CsetPrinter.SetPrinterAsDefault (cmbPrinter.Text)
    

End Sub

Private Sub WriteIndoorBill()
    Dim rsTemBill As New ADODB.Recordset
    With rsTemBill
        If .State = 1 Then .Close
        temSql = "Select * from tblIndoorBill where IndoorBillID = " & Val(txtBillID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount <= 0 Then
            .AddNew
        End If
        !SaleBillCompleted = True
        !SaleCategoryID = SaleCategory.SaleCategoryID
        !PaymentMethodID = SaleCategory.PaymentMethodID
        !PurchaseValue = Val(txtTotalCost.Text)
        !Completed = True
        !CompletedDateTime = Now
        !CompletedDate = Date
        !CompletedTime = Time
        !CompletedUserID = UserID
        !SaleValueMonthly = Val(txtRunningSaleM.Text)
        !PurchaseValueMonthly = Val(txtRunningPurchaseM.Text)
        !SaleValueAnnually = Val(txtRunningSaleY.Text)
        !PurchaseValueAnnually = Val(txtRunningPurchaseY.Text)
        !CustomerID = Val(cmbCustomer.BoundText)
        !ForOutDoor = SaleCategory.ForOutDoor
        !ForBHT = SaleCategory.ForBHT
        !ForInstitution = SaleCategory.ForInstitution
        !ForUnit = SaleCategory.ForUnit
        !QueueID = Val(txtQueueID.Text)
        !Comments = txtComments.Text
        .Update
    
        txtBillID.Text = !IndoorBillID
        .Close
    End With
End Sub


Private Function CanSettle() As Boolean
    Dim tr As Integer
    CanSettle = False
    
    If GridItem.Rows <= 1 Then
        tr = MsgBox("There are no items to sell", vbCritical, "No Items")
        FocusControl.SetFocus
        Exit Function
    End If
    
    
    
    If SaleCategory.PaymentMethod = "CASH" Then
    
    ElseIf SaleCategory.PaymentMethod = "CREDIT" Then
        If IsNumeric(cmbCustomer.BoundText) = False Then
            tr = MsgBox("Please select the customer", vbCritical, "Customer?")
            cmbCustomer.SetFocus
            Exit Function
        End If
    ElseIf SaleCategory.PaymentMethod = "CHEQUE" Then
    
    ElseIf SaleCategory.PaymentMethod = "CREDIT CARD" Then
    
    End If
    
    If SaleCategory.ForBHT = True Then
        If IsNumeric(cmbCustomer.BoundText) = False Then
            tr = MsgBox("You have not selected the BHT number", vbCritical, "BHT?")
            cmbCustomer.SetFocus
            Exit Function
        End If
    ElseIf SaleCategory.ForOutDoor = True Then
    
    ElseIf SaleCategory.ForStaff = True Then
        If IsNumeric(cmbCustomer.BoundText) = False Then
            tr = MsgBox("You have not selected the staff member to whom the items are issued", vbCritical, "Staff member?")
            cmbCustomer.SetFocus
            Exit Function
        End If
    ElseIf SaleCategory.ForUnit = True Then
        If IsNumeric(cmbCustomer.BoundText) = False Then
            tr = MsgBox("You have not selected the unit")
            cmbCustomer.SetFocus
            Exit Function
        End If
    ElseIf SaleCategory.ForInstitution = True Then
        If IsNumeric(cmbCustomer.BoundText) = False Then
            tr = MsgBox("You have not selected the institution")
            cmbCustomer.SetFocus
            Exit Function
        End If
    End If
    
    CanSettle = True
End Function

Private Sub cmbCategory_Change()
    Dim ItemName As New clsFillCombos
    Dim ItemCode As New clsFillCombos
    cmbItem.Text = Empty
    cmbCode.Text = Empty
    If IsNumeric(cmbCategory.BoundText) = True Then
        ItemName.FillLongBoolCombo cmbItem, "Item", "Item", "ItemCategoryID", Val(cmbCategory.BoundText), "IsItemName", True
        ItemCode.FillLongBoolCombo cmbCode, "Item", "ItemCode", "ItemCategoryID", Val(cmbCategory.BoundText), "IsItemName", True
    Else
        ItemName.FillBoolCombo cmbItem, "Item", "Item", "IsItemName", True
        ItemCode.FillSpecificFieldBoolCombo cmbCode, "Item", "ItemCode", "ItemCode", "IsItemName", True
    End If
End Sub

Private Sub cmbCategory_Click(Area As Integer)
    Set FocusControl = cmbCategory
End Sub

Private Sub cmbCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        If cmbItem.Name = FocusControl.Name Then
            cmbItem.SetFocus
        Else
            cmbCode.SetFocus
        End If
    ElseIf KeyCode = vbKeyEscape Then
        cmbCategory.Text = Empty
    End If
End Sub

Private Sub cmbCode_Change()
    If IsNumeric(cmbCode.BoundText) = True Then cmbItem.BoundText = cmbCode.BoundText
End Sub

Private Sub cmbCode_Click(Area As Integer)
    Set FocusControl = cmbCode
End Sub

Private Sub cmbCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        cmbCode = Empty
    ElseIf KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtQty.SetFocus
    End If
End Sub

Private Sub cmbItem_Change()
    If IsNumeric(cmbItem.BoundText) = False Then Exit Sub
    cmbCode.BoundText = cmbItem.BoundText
    SaleItem.ID = cmbItem.BoundText
    lblIssueUnit.Caption = SaleItem.IssueUnit
    Call DisplayRate
    Call FillBatches
End Sub

Private Sub FillBatches()
    Dim temID As Long
    With rsBatch
        If .State = 1 Then .Close
        temSql = "SELECT tblItemBatch.ItemBatchID, tblItemBatch.ItemBatch & ' - ' & tblItemBatchStock.Stock & ' ' & tblItemUnit.ItemUnit & ' ' & tblItemBatch.DOE  AS Display " & _
                    "FROM ((tblItem LEFT JOIN tblItemBatch ON tblItem.ItemID = tblItemBatch.ItemID) LEFT JOIN tblItemBatchStock ON tblItemBatch.ItemBatchID = tblItemBatchStock.ItemBatchID) LEFT JOIN tblItemUnit ON tblItem.IssueUnitID = tblItemUnit.ItemUnitID " & _
                    "WHERE tblItemBatchStock.DepartmentID = " & UserDepartmentID & " AND tblItem.ItemID = " & Val(cmbItem.BoundText) & "  AND ( tblItemBatchStock.Stock > 0  OR (SELECT tblSaleItem.ItemBatchID FROM tblSaleItem WHERE tblSaleItem.SaleItemID = " & Val(txtEditID.Text) & "  AND tblSaleItem.Deleted = False ) = tblItemBatch.ItemBatchID  ) " & _
                    "Order by tblItemBatch.DOE DESC, tblItemBatchStock.Stock DESC  "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temID = !ItemBatchID
        Else
            temID = 0
        End If
    End With
    With lstBatch
        Set .RowSource = rsBatch
        .BoundColumn = "ItemBatchID"
        .ListField = "Display"
        .BoundText = temID
    End With
End Sub


Private Sub FillBatchesByStock()
    Dim temID As Long
    With rsBatch
        If .State = 1 Then .Close
        temSql = "SELECT tblItemBatch.ItemBatchID, tblItemBatch.ItemBatch & ' - ' & tblItemBatchStock.Stock & ' ' & tblItemUnit.ItemUnit & ' ' & tblItemBatch.DOE  AS Display " & _
                    "FROM ((tblItem LEFT JOIN tblItemBatch ON tblItem.ItemID = tblItemBatch.ItemID) LEFT JOIN tblItemBatchStock ON tblItemBatch.ItemBatchID = tblItemBatchStock.ItemBatchID) LEFT JOIN tblItemUnit ON tblItem.IssueUnitID = tblItemUnit.ItemUnitID " & _
                    "WHERE tblItemBatchStock.DepartmentID = " & UserDepartmentID & " AND tblItem.ItemID = " & Val(cmbItem.BoundText) & "  AND ( tblItemBatchStock.Stock > 0  OR (SELECT tblSaleItem.ItemBatchID FROM tblSaleItem WHERE tblSaleItem.SaleItemID = " & Val(txtEditID.Text) & "  AND tblSaleItem.Deleted = False ) = tblItemBatch.ItemBatchID  ) " & _
                    "Order by tblItemBatchStock.Stock DESC, tblItemBatch.DOE DESC  "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temID = !ItemBatchID
        Else
            temID = 0
        End If
    End With
    With lstBatch
        Set .RowSource = rsBatch
        .BoundColumn = "ItemBatchID"
        .ListField = "Display"
        .BoundText = temID
    End With

End Sub

Private Sub cmbItem_Click(Area As Integer)
    Set FocusControl = cmbItem
End Sub

Private Sub cmbItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        txtQty.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbItem.Text = Empty
    End If
End Sub

Private Sub cmbSaleCategory_Change()
    Call SetSaleCategory
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
'    txtBillID.Text = IndoorBill.NewIndoorBillID
    Call FillCombos
    Call FormatGrid
    Call GetSettings

    Call SetSaleCategory
    Call FillGrid
    
    If Date <> TodayStartedDate Then End
    
End Sub

Private Sub SetSaleCategory()
    If SaleCategory.ForOutDoor = True Then
        Customer.FillAnyCombo cmbCustomer, "Patient", True
    ElseIf SaleCategory.ForBHT = True Then
        Customer.FillAnyCombo cmbCustomer, "BHT", True
    ElseIf SaleCategory.ForStaff = True Then
        Customer.FillAnyCombo cmbCustomer, "Staff", True
    ElseIf SaleCategory.ForUnit = True Then
        Customer.FillAnyCombo cmbCustomer, "Department", True
    ElseIf SaleCategory.ForInstitution = True Then
        Customer.FillAnyCombo cmbCustomer, "Institution", True
    End If
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call ClearGrid
    Call SaveSettings
End Sub

Private Sub ClearGrid()
    With GridItem
        While .Rows >= 2
            .row = .Rows - 1
            txtDeleteID.Text = Val(.TextMatrix(.row, 5))
            btnDelete_Click
        Wend
    End With
End Sub

Private Sub GridItem_Click()
    With GridItem
        txtDeleteID.Text = Val(.TextMatrix(.row, 5))
    End With
End Sub

Private Sub GridItem_DblClick()
    With GridItem
        txtEditID.Text = Val(.TextMatrix(.row, 5))
    End With
    Dim rsItemEdit As New ADODB.Recordset
    With rsItemEdit
        If .State = 1 Then .Close
        temSql = "SELECT tblItem.ItemCategoryID, tblItemBatch.ItemBatchID, tblItemBatch.DOE, tblItemBatch.DOM, tblSaleItem.* FROM (tblSaleItem LEFT JOIN tblItem ON tblSaleItem.ItemID = tblItem.ItemID) LEFT JOIN tblItemBatch ON tblSaleItem.ItemBatchID = tblItemBatch.ItemBatchID WHERE (((tblSaleItem.SaleItemID)=" & Val(txtEditID.Text) & ")) "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbCategory.BoundText = !ItemCategoryID
            cmbItem.BoundText = !ItemID
            cmbCode.BoundText = !ItemID
            txtQty.Text = !SaleQuentity
            txtRate.Text = Format(!SaleRate, "0.00")
            txtPrice.Text = Format(!SaleValue, "0.00")
            Call ListBatches(Val(!ItemID))
            lstBatch.BoundText = ![tblItemBatch.ItemBatchID]
        End If
        .Close
    End With
    
End Sub


Private Sub GridItem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim temSelect As String
    
    ReDim AlternateItem(7)
    ReDim AlternateItemID(7)
    
    If Button = vbRightButton Then
        ItemRow = Y \ GridItem.RowHeight(0)
        If ItemRow < 1 Then Exit Sub
        ItemRow = ItemRow + GridItem.TopRow - 1
        TopRow = GridItem.TopRow
        If ItemRow > GridItem.Rows - 1 Then Exit Sub
        
        
        If GridItem.RowSel <> ItemRow Then GridItem.RowSel = ItemRow
        
        
        AlternateItem(0) = "X 1/2"
        AlternateItemID(0) = 0
        AlternateItem(1) = "X 1/4"
        AlternateItemID(1) = 1
        AlternateItem(2) = "-"
        AlternateItemID(2) = 2
        AlternateItem(3) = "X 2"
        AlternateItemID(3) = 3
        AlternateItem(4) = "X 4"
        AlternateItemID(4) = 4
        AlternateItem(5) = "-"
        AlternateItemID(5) = 5
        AlternateItem(6) = "Delete"
        AlternateItemID(6) = 6
        AlternateItem(7) = "Edit"
        AlternateItemID(7) = 7
        
        ucpPrescreption.Popup AlternateItem()
        
        RxPopAction (ucpPrescreption.ReturnIndex)
    End If
End Sub


Private Sub RxPopAction(RxPopUpReturnIndex As Long)
    Dim myTopRow As Long
    Dim myBotRow As Long
    
    With GridItem
        If .RowSel > .row Then
            myTopRow = .row
            myBotRow = .RowSel
        Else
            myTopRow = .RowSel
            myBotRow = .row
        End If
        
        
        Dim i As Integer
        
        If RxPopUpReturnIndex < 0 Then Exit Sub
        
        For i = myTopRow To myBotRow
            If AlternateItemID(RxPopUpReturnIndex) = 0 Then
                .row = i
                GridItem_DblClick
                txtQty.Text = txtQty.Text \ 2
                btnAdd_Click
            ElseIf AlternateItemID(RxPopUpReturnIndex) = 1 Then
                .row = i
                GridItem_DblClick
                txtQty.Text = txtQty.Text \ 4
                btnAdd_Click
            ElseIf AlternateItemID(RxPopUpReturnIndex) = 2 Then
            
            ElseIf AlternateItemID(RxPopUpReturnIndex) = 3 Then
                .row = i
                GridItem_DblClick
                txtQty.Text = txtQty.Text * 2
                btnAdd_Click
            ElseIf AlternateItemID(RxPopUpReturnIndex) = 4 Then
                .row = i
                GridItem_DblClick
                txtQty.Text = txtQty.Text * 4
                btnAdd_Click
            ElseIf AlternateItemID(RxPopUpReturnIndex) = 5 Then
                    
            ElseIf AlternateItemID(RxPopUpReturnIndex) = 6 Then
                .row = i
                GridItem_Click
                btnDelete_Click
            ElseIf AlternateItemID(RxPopUpReturnIndex) = 7 Then
                .row = i
                GridItem_DblClick
                Exit Sub
            End If
        
        Next
    End With
End Sub


Private Sub txtCostRate_Change()
    Call CalculateValues
End Sub


Private Sub txtQty_Change()
    Call CalculateValues
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnAdd_Click
    ElseIf KeyCode = vbKeyEscape Then
        txtQty.Text = Empty
    End If
End Sub

Private Sub txtRate_Change()
    Call CalculateValues
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

