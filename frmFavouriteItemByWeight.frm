VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmFavouriteItemByWeight 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Favourite Item By Weight"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14970
   BeginProperty Font 
      Name            =   "Verdana"
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
   ScaleHeight     =   7320
   ScaleWidth      =   14970
   Begin VB.TextBox txtDeletedID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1080
      TabIndex        =   26
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAddID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtTo 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtFrom 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtIssueQty 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   12240
      TabIndex        =   19
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtDuration 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   9120
      TabIndex        =   14
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtDose 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   1080
      Width           =   855
   End
   Begin VB.CheckBox chkIndoor 
      Height          =   255
      Left            =   11760
      TabIndex        =   17
      Top             =   1080
      Width           =   255
   End
   Begin MSDataListLib.DataCombo cmbItem 
      Height          =   360
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbDoseUnit 
      Height          =   360
      Left            =   5640
      TabIndex        =   10
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbFrequencyUnit 
      Height          =   360
      Left            =   7440
      TabIndex        =   12
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbDurationUnit 
      Height          =   360
      Left            =   10080
      TabIndex        =   15
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbIssueUnit 
      Height          =   360
      Left            =   13080
      TabIndex        =   20
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid gridItem 
      Height          =   4935
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   8705
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   13680
      TabIndex        =   21
      Top             =   1560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      Left            =   13680
      TabIndex        =   23
      Top             =   2040
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   13800
      TabIndex        =   24
      Top             =   6840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "C&lose"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      Left            =   2280
      TabIndex        =   7
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "Category"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Indoor"
      Height          =   255
      Left            =   11520
      TabIndex        =   16
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Wt. To"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Wt. From"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Issue"
      Height          =   255
      Left            =   12240
      TabIndex        =   18
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Duration"
      Height          =   255
      Left            =   9120
      TabIndex        =   13
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Frequency"
      Height          =   255
      Left            =   7440
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Dose"
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Item"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmFavouriteItemByWeight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsTem As New ADODB.Recordset
    

Private Sub btnAdd_Click()
    If IsNumeric(cmbItem.BoundText) = False Then
        MsgBox "Please select an Item"
        cmbItem.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtFrom.Text) = False Then
        MsgBox "Please select start weight"
        txtFrom.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtTo.Text) = False Then
        MsgBox "Please select end weight"
        txtTo.SetFocus
        Exit Sub
    End If
'    If IsNumeric(txtDose.Text) = False Then
'        MsgBox "Please select a Dose"
'        txtDose.SetFocus
'        Exit Sub
'    End If
'    If IsNumeric(cmbDoseUnit.BoundText) = False Then
'        MsgBox "Please select a Dose Unit"
'        cmbDoseUnit.SetFocus
'        Exit Sub
'    End If
    If IsNumeric(cmbFrequencyUnit.BoundText) = False Then
        MsgBox "Please select a frequency"
        cmbFrequencyUnit.SetFocus
        Exit Sub
    End If
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(txtAddID.Text) = True Then
            temSql = "Select * from tblFavouriteItem where FavouriteItemID = " & Val(txtAddID.Text)
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        Else
            temSql = "Select * from tblFavouriteItem"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !ItemID = Val(cmbItem.BoundText)
            !ByWeight = True
        End If
        !FromWeight = txtFrom.Text
        !ToWeight = txtTo.Text
        If IsNumeric(txtDose.Text) = True Then
            !Dose = txtDose.Text
        Else
            !Dose = Null
        End If
        !DoseUnitID = Val(cmbDoseUnit.BoundText)
        !FrequencyUnitID = Val(cmbFrequencyUnit.BoundText)
        !Duration = Val(txtDuration.Text)
        !DurationUnitID = Val(cmbDurationUnit.BoundText)
        !IssueQuentity = Val(txtIssueQty.Text)
        !IssueUnitID = Val(cmbIssueUnit.BoundText)
        !ItemCategoryID = Val(cmbCategory.BoundText)
        If chkIndoor.Value = 1 Then
            !Indoor = True
            !IndoorOutdoor = "Indoor"
        Else
            !Indoor = False
            !IndoorOutdoor = "Outdoor"
        End If
        .Update
    End With
    Call FormatGrid
    Call FillGrid
    Call ClearValues
    txtFrom.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    If IsNumeric(txtDeletedID.Text) = False Then
        MsgBox "Please select one to delete"
        GridItem.SetFocus
        Exit Sub
    End If
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblFavouriteItem where FavouriteItemID = " & Val(txtDeletedID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedTime = Now
            !DeletedUserID = UserID
            .Update
        End If
        .Close
    End With
    Call FormatGrid
    Call FillGrid
    Call ClearValues
    txtFrom.SetFocus
End Sub

Private Sub cmbItem_Change()
    Call FormatGrid
    Call ClearValues
    Call FillGrid
    
    Dim temItem As New clsItem
    temItem.ID = Val(cmbItem.BoundText)
    If temItem.IsItem = True Then
        cmbCategory.Enabled = False
        cmbCategory.BoundText = temItem.CategoryID
    Else
        cmbCategory.Enabled = True
        cmbCategory.Text = Empty
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
        End If
    Next
End Sub

Private Sub Form_Load()
    Call SetColours
    Call FillCombos
    GetCommonSettings Me
    
    
End Sub

Private Sub FillCombos()
    Dim Item As New clsFillCombos
    Item.FillAnyCombo cmbItem, "Item", True
    
    Dim IUnit As New clsFillCombos
    IUnit.FillBoolCombo cmbIssueUnit, "ItemUnit", "ItemUnit", "IsIssueUnit", True
    
    '   Change1
    Dim DUnit As New clsFillCombos
    DUnit.FillBoolCombo cmbDoseUnit, "ItemUnit", "ItemUnit", "IsDoseUnit", True
    '   Change1
    
    
    Dim FUnit As New clsFillCombos
    FUnit.FillBoolCombo cmbFrequencyUnit, "ItemUnit", "ItemUnit", "IsFrequencyUnit", True
    
    Dim Dura As New clsFillCombos
    Dura.FillBoolCombo cmbDurationUnit, "ItemUnit", "ItemUnit", "IsDurationUnit", True
    
    Dim ICat As New clsFillCombos
    ICat.FillBoolCombo cmbCategory, "Category", "Category", "IsItemCategory", True
End Sub

Private Sub FillGrid()
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblFavouriteItem.FavouriteItemID, 'From ' & [FromWeight] & ' To ' & [ToWeight] & ' kg' AS WeightRange, [tblItem.Item] & ' (' & [tblCategory.Category] & ') ' & [tblFavouriteItem].[Dose] & ' ' & [tblDoseUnit].[ItemUnit] & ' ' & [tblFrequencyUnit].[ItemUnit] & ' ' & ' for ' & [tblFavouriteItem].[Duration] & ' ' & [tblDurationUnit].[ItemUnit] & ' (' & [tblFavouriteItem].[IndoorOutdoor] & ')' AS Prescreption, 'Issue ' & tblFavouriteItem.IssueQuentity & ' ' & tblIssueUnit.ItemUnit AS ToIssue " & _
                    "FROM ((((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblFavouriteItem ON tblIssueUnit.ItemUnitID = tblFavouriteItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblFavouriteItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblFavouriteItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblFavouriteItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblCategory ON tblFavouriteItem.ItemCategoryID = tblCategory.CategoryID) LEFT JOIN tblItem ON tblFavouriteItem.ItemID = tblItem.ItemID " & _
                    "WHERE (((tblFavouriteItem.ItemID)=" & Val(cmbItem.BoundText) & ") AND ((tblFavouriteItem.ByWeight)=True) AND ((tblFavouriteItem.Deleted)=False))" & _
                    "ORDER BY tblFavouriteItem.FromWeight, tblFavouriteItem.ToWeight"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            GridItem.Rows = GridItem.Rows + 1
            GridItem.Row = GridItem.Rows - 1
            GridItem.Col = 0
            GridItem.Text = ![FavouriteItemID]
            GridItem.Col = 1
            GridItem.Text = !WeightRange
            GridItem.Col = 2
            GridItem.Text = !Prescreption
            GridItem.Col = 3
            GridItem.Text = !ToIssue
            .MoveNext
        Wend
    End With
End Sub

Private Sub FormatGrid()
    With GridItem
        .Rows = 1
        .Cols = 4
        .Clear
        .Row = 0
        .Col = 0
        .Text = "ID"
        .Col = 1
        .Text = "Weight Range"
        .Col = 2
        .Text = "How to prescribe"
        .Col = 3
        .Text = "Indoor Issue Quentity"
        .ColWidth(0) = 0
        .ColWidth(1) = 2500
        .ColWidth(3) = 2900
        .ColWidth(2) = .Width - 2650 - 2900
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub

Private Sub GridItem_Click()
    Dim temRow As Integer
    With GridItem
        temRow = .Row
        txtDeletedID.Text = .TextMatrix(temRow, 0)
    End With
End Sub

Private Sub ClearValues()
    cmbDoseUnit.Text = Empty
    cmbDurationUnit.Text = Empty
    cmbFrequencyUnit.Text = Empty
    cmbIssueUnit.Text = Empty
    txtAddID.Text = Empty
    txtDeletedID.Text = Empty
    txtDose.Text = Empty
    txtDuration.Text = Empty
    txtTo.Text = Empty
    txtFrom.Text = Empty
    txtIssueQty.Text = Empty
    cmbCategory.Text = Empty
    chkIndoor.Value = 0
End Sub

Private Sub GridItem_DblClick()
    Dim temRow As Integer
    Dim EditID As Long
    With GridItem
        temRow = .Row
        EditID = .TextMatrix(temRow, 0)
    End With
    Call ClearValues
    txtAddID.Text = EditID
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblFavouriteItem where FavouriteItemID = " & EditID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbDoseUnit.BoundText = !DoseUnitID
            cmbIssueUnit.BoundText = !IssueUnitID
            cmbFrequencyUnit.BoundText = !FrequencyUnitID
            cmbDurationUnit.BoundText = !DurationUnitID
            txtDose.Text = !Dose
            txtDuration.Text = !Duration
            txtIssueQty.Text = !IssueQuentity
            If !Indoor = True Then
                chkIndoor.Value = 1
            Else
                chkIndoor.Value = 0
            End If
            txtFrom.Text = !FromWeight
            txtTo.Text = !ToWeight
            If IsNull(!ItemCategoryID) = False Then
                cmbCategory.BoundText = !ItemCategoryID
            End If
        End If
        .Close
    End With
End Sub
