VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmItemUnitRepationship 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Unit Relationship"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
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
   ScaleHeight     =   4350
   ScaleWidth      =   7800
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Appearance      =   3
      Caption         =   "&Save"
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
   Begin VB.TextBox txtRatio 
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   1200
      Width           =   1695
   End
   Begin MSDataListLib.DataCombo cmbIUnit2 
      Height          =   360
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbIUnit1 
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   635
      _Version        =   393216
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnDelete 
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   3720
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label Label1 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   2
      Top             =   1200
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   4560
      Y1              =   1440
      Y2              =   1440
   End
End
Attribute VB_Name = "frmItemUnitRepationship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsUnit1 As New ADODB.Recordset
    Dim rsUnit2 As New ADODB.Recordset
    
Private Sub FillCombos()
    With rsUnit1
        If .State = 1 Then .Close
        temSql = "Select * from tblItemUnit where (IsIssueUnit = true or IsStrengthUnit = true or IsDoseUnit = true or IsPackUnit = true) and Deleted = False order by ItemUnit"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbIUnit1
        Set .RowSource = rsUnit1
        .ListField = "ItemUnit"
        .BoundColumn = "itemUnitID"
    End With
    With rsUnit2
        If .State = 1 Then .Close
        temSql = "Select * from tblItemUnit where (IsIssueUnit = true or IsStrengthUnit = true or IsDoseUnit = true or IsPackUnit = true) and Deleted = False order by ItemUnit"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbIUnit2
        Set .RowSource = rsUnit2
        .ListField = "ItemUnit"
        .BoundColumn = "itemUnitID"
    End With
End Sub

Private Sub btnDelete_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblItemUnitRatio where ItemUnitID1 = " & Val(cmbIUnit1.BoundText) & " AND ItemUnitID2 = " & Val(cmbIUnit2.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Ratio = Empty
            !Deleted = True
            !DeletedTime = Now
            !DeletedUserID = UserID
            .Update
        End If
        .Close
    End With

End Sub

Private Sub btnSave_Click()
    If IsNumeric(cmbIUnit1.BoundText) = False Then
        MsgBox "Please select a Unit"
        cmbIUnit1.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbIUnit2.BoundText) = False Then
        MsgBox "Please select a Unit"
        cmbIUnit2.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtRatio.Text) = False Then
        MsgBox "Please enter a ratio"
        txtRatio.SetFocus
        Exit Sub
    End If
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblItemUnitRatio where ItemUnitID1 = " & Val(cmbIUnit1.BoundText) & " AND ItemUnitID2 = " & Val(cmbIUnit2.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Ratio = txtRatio.Text
        Else
            .AddNew
            !ItemUnitID1 = cmbIUnit1.BoundText
            !ItemUnitID2 = cmbIUnit2.BoundText
            !Ratio = txtRatio.Text
        End If
        .Update
        .Close
    End With
    cmbIUnit1.Text = Empty
    cmbIUnit2.Text = Empty
    txtRatio.Text = Empty
End Sub

Private Sub cmbIUnit1_Click(Area As Integer)
    Call FindRatio
End Sub

Private Sub FindRatio()
    txtRatio.Text = Empty
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblItemUnitRatio where ItemUnitID1 = " & Val(cmbIUnit1.BoundText) & " AND ItemUnitID2 = " & Val(cmbIUnit2.BoundText) & " AND Deleted = False"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtRatio.Text = !Ratio
        End If
        .Close
    End With
End Sub

Private Sub cmbIUnit2_Click(Area As Integer)
    Call FindRatio
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
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub
