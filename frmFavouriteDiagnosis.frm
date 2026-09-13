VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmFavouriteDiagnosis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Favourite Diagnosis By Age"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10005
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
   ScaleWidth      =   10005
   Begin VB.TextBox txtDeletedID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   6600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtAddID 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   120
      TabIndex        =   12
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
   Begin MSDataListLib.DataCombo cmbDiagnosis 
      Height          =   360
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSFlexGridLib.MSFlexGrid gridItem 
      Height          =   4935
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8705
      _Version        =   393216
      SelectionMode   =   1
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   8640
      TabIndex        =   8
      Top             =   1080
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
      Left            =   8640
      TabIndex        =   10
      Top             =   1560
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
      Left            =   8640
      TabIndex        =   11
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
   Begin MSDataListLib.DataCombo cmbItem 
      Height          =   360
      Left            =   2400
      TabIndex        =   7
      Top             =   1080
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label3 
      Caption         =   "&Item"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Age &To"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Age &From"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Dia&gnosis"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmFavouriteDiagnosis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsTem As New ADODB.Recordset
    

Private Sub btnAdd_Click()
    If IsNumeric(cmbDiagnosis.BoundText) = False Then
        MsgBox "Please select an Diagnosis"
        cmbDiagnosis.SetFocus
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
    If IsNumeric(cmbItem.BoundText) = False Then
        MsgBox "Please select an Item"
        cmbItem.SetFocus
        Exit Sub
    End If
    With rsTem
        If .State = 1 Then .Close
        If IsNumeric(txtAddID.Text) = True Then
            temSql = "Select * from tblFavouriteDiagnosis where FavouriteDiagnosisID = " & Val(txtAddID.Text) ' & " And Deleted = False"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        Else
            temSql = "Select * from tblFavouriteDiagnosis"
            .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
            .AddNew
            !DiagnosisID = Val(cmbDiagnosis.BoundText)
        End If
        !FromAge = Val(txtFrom.Text) * 12
        !ToAge = Val(txtTo.Text) * 12
        !ItemID = Val(cmbItem.BoundText)
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
        temSql = "Select * from tblFavouriteDiagnosis where FavouriteDiagnosisID = " & Val(txtDeletedID.Text)
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

Private Sub cmbDiagnosis_Change()
    Call FormatGrid
    Call ClearValues
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
    Dim Diag As New clsFillCombos
    Diag.FillAnyCombo cmbDiagnosis, "Diagnosis", True
End Sub

Private Sub FillGrid()
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblFavouriteDiagnosis.FavouriteDiagnosisID, 'From ' & [FromAge]/12 & ' To ' & [ToAge]/12 & ' Years' AS AgeRange, tblItem.Item " & _
                    "FROM tblFavouriteDiagnosis LEFT JOIN tblItem ON tblFavouriteDiagnosis.ItemID = tblItem.ItemID " & _
                    "WHERE (((tblFavouriteDiagnosis.Deleted)=False) AND ((tblFavouriteDiagnosis.DiagnosisID)=" & Val(cmbDiagnosis.BoundText) & "))" & _
                    "ORDER BY tblFavouriteDiagnosis.FromAge, tblFavouriteDiagnosis.ToAge, tblItem.Item"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            GridItem.Rows = GridItem.Rows + 1
            GridItem.Row = GridItem.Rows - 1
            GridItem.Col = 0
            GridItem.Text = ![FavouriteDiagnosisID]
            GridItem.Col = 1
            GridItem.Text = !AgeRange
            GridItem.Col = 2
            GridItem.Text = !Item
            .MoveNext
        Wend
    End With
End Sub

Private Sub FormatGrid()
    With GridItem
        .Rows = 1
        .Cols = 3
        .Clear
        .Row = 0
        .Col = 0
        .Text = "ID"
        .Col = 1
        .Text = "Age Range"
        .Col = 2
        .Text = "Item prescribe"
        .ColWidth(0) = 0
        .ColWidth(1) = 2900
        .ColWidth(2) = .Width - 100 - 2900
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
    cmbItem.Text = Empty
    txtAddID.Text = Empty
    txtDeletedID.Text = Empty
    txtTo.Text = Empty
    txtFrom.Text = Empty
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
        temSql = "Select * from tblFavouriteDiagnosis where FavouriteDiagnosisID = " & EditID
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbItem.BoundText = !ItemID
            txtFrom.Text = !FromAge / 12
            txtTo.Text = !ToAge / 12
        End If
        .Close
    End With
End Sub
