VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAssignIxValuesToFields 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Investigation Item Values"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
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
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   6915
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   2160
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
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   4200
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
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
   Begin MSDataListLib.DataCombo cmbIx 
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbIxItem 
      Height          =   360
      Left            =   1680
      TabIndex        =   3
      Top             =   600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbIxItemValueCategory 
      Height          =   360
      Left            =   1680
      TabIndex        =   5
      Top             =   1080
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbDefaultValue 
      Height          =   360
      Left            =   1680
      TabIndex        =   7
      Top             =   1560
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label4 
      Caption         =   "&Default Value"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "&Value Category"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Value &Field"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "&Investigation"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAssignIxValuesToFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsViewIxItem As New ADODB.Recordset
    Dim rsViewIxItemValueCategory As New ADODB.Recordset
    Dim rsViewValues As New ADODB.Recordset
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    If IsNumeric(cmbIx.BoundText) = False Then
        MsgBox ("Please select investigation")
        cmbIx.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbIxItem.BoundText) = False Then
        MsgBox "Please selct a value field"
        cmbIxItem.SetFocus
        Exit Sub
    End If
'    If IsNumeric(cmbIxItemValueCategory.BoundText) = False Then
'        MsgBox "Please select a value category"
'        cmbIxItemValueCategory.SetFocus
'        Exit Sub
'    End If
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIxItem where IxItemID = " & Val(cmbIxItem.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
             !IxValueCategoryID = Val(cmbIxItemValueCategory.BoundText)
             !IxValueDefaultID = Val(cmbDefaultValue.BoundText)
            .Update
        End If
        .Close
    End With
    cmbDefaultValue.Text = Empty
    cmbIxItemValueCategory.Text = Empty
    cmbIxItem.SetFocus
End Sub

Private Sub cmbDefaultValue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        btnSave_Click
    ElseIf KeyCode = vbKeyEscape Then
        cmbDefaultValue.Text = Empty
    End If
End Sub

Private Sub cmbIx_Change()
    With rsViewIxItem
        If .State = 1 Then .Close
        temSql = "Select * from tblIxItem where IxID = " & Val(cmbIx.BoundText) & " AND Deleted = False AND IsValue = true Order by IxItem"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbIxItem
        Set .RowSource = rsViewIxItem
        .ListField = "IxItem"
        .BoundColumn = "IxItemID"
    End With
    cmbIxItem.Text = Empty
    cmbDefaultValue.Text = Empty
    cmbIxItemValueCategory.Text = Empty
End Sub


Private Sub cmbIx_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbIxItem.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbIx.Text = Empty
    End If
End Sub

Private Sub cmbIxItem_Change()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIxItem where IxItemID = " & Val(cmbIxItem.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            cmbIxItemValueCategory.BoundText = !IxValueCategoryID
            If IsNull(!IxValueDefaultID) = False Then
                cmbDefaultValue.BoundText = !IxValueDefaultID
            Else
                cmbDefaultValue.Text = Empty
            End If
        Else
            cmbIxItemValueCategory.Text = Empty
            cmbDefaultValue.BoundText = Empty
        End If
    End With
End Sub

Private Sub cmbIxItem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbIxItemValueCategory.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbIxItem.Text = Empty
    End If
End Sub

Private Sub cmbIxItemValueCategory_Change()
    With rsViewValues
        If .State = 1 Then .Close
        temSql = "SELECT tblIxValue.IxValueID, tblIxValue.IxValue FROM (tblIxValuesForCategories RIGHT JOIN tblCategory ON tblIxValuesForCategories.CategoryID = tblCategory.CategoryID) LEFT JOIN tblIxValue ON  tblIxValuesForCategories.ValueID = tblIxValue.IxValueID WHERE (((tblCategory.CategoryID)=" & Val(cmbIxItemValueCategory.BoundText) & ")) ORDER BY tblIxValue.IxValue "
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbDefaultValue
        Set .RowSource = rsViewValues
        .ListField = "IxValue"
        .BoundColumn = "IxValueID"
        .BoundText = IxValueDefaultID(Val(cmbIxItem.BoundText))
    End With
     
End Sub

Private Function IxValueDefaultID(IxItemID As Long) As Long
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblIxItem.IxValueDefaultID FROM tblIxItem WHERE (((tblIxItem.IxItemID)=" & IxItemID & "))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 1 Then
            IxValueDefaultID = !IxValueDefaultID
        Else
            IxValueDefaultID = 0
        End If
        .Close
    End With
End Function

Private Sub cmbIxItemValueCategory_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = Empty
        cmbDefaultValue.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        cmbIxItemValueCategory.Text = Empty
    End If
End Sub

Private Sub Form_Load()
    GetCommonSettings Me
    
    Call SetColours
    Dim Ix As New clsFillCombos
    Ix.FillAnyCombo cmbIx, "Ix", True
    Dim IxItemValueCategory As New clsFillCombos
    IxItemValueCategory.FillBoolCombo cmbIxItemValueCategory, "Category", "Category", "IsIxValueCategory"
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub
