VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmFavouriteItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Favourite Item"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12600
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
   ScaleWidth      =   12600
   Begin VB.TextBox txtDeletedID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1080
      TabIndex        =   26
      Top             =   6600
      Width           =   855
   End
   Begin VB.TextBox txtAddID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      TabIndex        =   25
      Top             =   6600
      Width           =   855
   End
   Begin VB.TextBox txtTo 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1200
      TabIndex        =   21
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtFrom 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtIssueQty 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   9720
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtDuration 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtDose 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.CheckBox chkIndoor 
      Height          =   255
      Left            =   9240
      TabIndex        =   0
      Top             =   1080
      Width           =   255
   End
   Begin MSDataListLib.DataCombo cmbItem 
      Height          =   360
      Left            =   1320
      TabIndex        =   4
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
      Left            =   3120
      TabIndex        =   5
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
      Left            =   4920
      TabIndex        =   6
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
      Left            =   7560
      TabIndex        =   7
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
      Left            =   10560
      TabIndex        =   8
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   9960
      TabIndex        =   9
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid gridItem 
      Height          =   4935
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8705
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   11400
      TabIndex        =   11
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
      Left            =   11400
      TabIndex        =   12
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
      Left            =   11400
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Indoor"
      Height          =   255
      Left            =   9000
      TabIndex        =   23
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Age To"
      Height          =   255
      Left            =   1200
      TabIndex        =   22
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Age From"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Category"
      Height          =   255
      Left            =   8880
      TabIndex        =   18
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Issue"
      Height          =   255
      Left            =   9720
      TabIndex        =   17
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Duration"
      Height          =   255
      Left            =   6600
      TabIndex        =   16
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Frequency"
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Dose"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Item"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmFavouriteItem"
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
        MsgBox "Please select start age"
        txtFrom.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtTo.Text) = False Then
        MsgBox "Please select end age"
        txtTo.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtDose.Text) = False Then
        MsgBox "Please select a Dose"
        txtDose.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbDoseUnit.BoundText) = False Then
        MsgBox "Please select a Dose Unit"
        cmbDoseUnit.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbFrequencyUnit.BoundText) = False Then
        MsgBox "Please select a frequency"
        cmbFrequencyUnit.SetFocus
        Exit Sub
    End If
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(txtAddID.Text) = True Then
            temSql = "Select * from tblFavouriteItem where ItemID = " & Val(txtAddID.Text) ' & " And ByAge = true AND Deleted = False"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        Else
            temSql = "Select * from tblFavouriteItem"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !ItemID = Val(cmbItem.BoundText)
            !ByAge = True
        End If
        !FromAge = txtFrom.Text * 12
        !ToAge = txtTo.Text * 12
        !Dose = txtDose.Text
        !DoseUnitID = Val(cmbDoseUnit.BoundText)
        !FrequencyUnitID = Val(cmbFrequencyUnit.BoundText)
        !Duration = Val(txtDuration.Text)
        !DurationUnitID = Val(cmbDurationUnit.BoundText)
        !IssueQuentity = Val(txtIssueQty.Text)
        !IssueUnitID = Val(cmbIssueUnit.BoundText)
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

Private Sub btnDelete_Click()
    If IsNumeric(txtDeletedID.Text) = False Then
        MsgBox "Please select one to delete"
        gridItem.SetFocus
        Exit Sub
    End If
    With rsTem
        temSql = "Select * from tblFavouriteItem where ItemID = " & Val(txtDeletedID.Text)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !deleted = True
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

Private Sub Form_Load()
    Call FillCombos
    
End Sub

Private Sub FillCombos()
    Dim Cat As New clsFillCombos
    Cat.FillAnyCombo cmbCategory, "ItemCategory", True

    Dim Item As New clsFillCombos
    Item.FillAnyCombo cmbItem, "Item", True
    
    Dim IUnit As New clsFillCombos
    IUnit.FillBoolCombo cmbIssueUnit, "ItemUnit", "ItemUnit", "IsIssueUnit", True
    
    Dim IUnit1 As New clsFillCombos
    IUnit1.FillBoolCombo cmbDoseUnit, "ItemUnit", "ItemUnit", "IsIssueUnit", True
    
    
    Dim FUnit As New clsFillCombos
    FUnit.FillAnyCombo cmbFrequencyUnit, "FrequencyUnit", True
    
    Dim Dura As New clsFillCombos
    Dura.FillAnyCombo cmbDurationUnit, "DurationUnit", True
    
End Sub

Private Sub FillGrid()
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblFavouriteItem.FavouriteItemID, 'From ' & Round([FromAge]/12,2) & ' To ' & Round([ToAge]/12,2) & ' Years' AS AgeRange, tblFavouriteItem.Dose & ' ' & tblDoseUnit.ItemUnit & ' ' & tblFrequencyUnit.FrequencyUnit & ' ' & ' for ' & tblFavouriteItem.Duration & ' ' & tblDurationUnit.DurationUnit & ' (' & tblFavouriteItem.IndoorOutdoor & ')' AS Prescreption, 'Issue ' & tblFavouriteItem.IssueQuentity & ' ' & tblIssueUnit.ItemUnitID AS ToIssue " & _
                    "FROM tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN ((tblFavouriteItem LEFT JOIN tblFrequencyUnit ON tblFavouriteItem.FrequencyUnitID = tblFrequencyUnit.FrequencyUnitID) LEFT JOIN tblDurationUnit ON tblFavouriteItem.DurationUnitID = tblDurationUnit.DurationUnitID) ON tblIssueUnit.ItemUnitID = tblFavouriteItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblFavouriteItem.DoseUnitID " & _
                    "WHERE (((tblFavouriteItem.ItemID)=" & Val(cmbItem.BoundText) & ") AND ((tblFavouriteItem.ByAge)=True) AND ((tblFavouriteItem.Deleted)=False))" & _
                    "ORDER BY tblFavouriteItem.FromAge, tblFavouriteItem.ToAge"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            gridItem.Rows = gridItem.Rows + 1
            gridItem.Row = gridItem.Rows - 1
            gridItem.Col = 0
            gridItem.Text = !FavouriteItemID
            gridItem.Col = 1
            gridItem.Text = !AgeRange
            gridItem.Col = 2
            gridItem.Text = !Prescreption
            gridItem.Col = 3
            gridItem.Text = !ToIssue
            .MoveNext
        Wend
    End With
End Sub

Private Sub FormatGrd()
    With gridItem
        .Rows = 1
        .Cols = 4
        .Clear
        .Row = 0
        .Col = 0
        .Text = "ID"
        .Col = 1
        .Text = "Age Range"
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

Private Sub gridItem_Click()
    Dim temRow As Integer
    With gridItem
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
    chkIndoor.Value = 0
End Sub

Private Sub gridItem_DblClick()
    Dim temRow As Integer
    Dim EditID As Long
    With gridItem
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
            cmbDurationUnit.BoundText = !DurationUnit
            txtDose.Text = !Dose
            txtDuration.Text = !Duration
            txtIssueQty.Text = !IssueQuentity
            If !Indoor = True Then
                chkIndoor.Value = 1
            Else
                chkIndoor.Value = 0
            End If
            txtFrom.Text = Round(!FromAge / 12, 2)
            txtTo.Text = Round(!ToAge / 12, 2)
        End If
        .Close
    End With
End Sub
