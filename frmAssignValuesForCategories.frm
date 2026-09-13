VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmAssignValuesForCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assign Investigation Values to Value Categories"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10050
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
   ScaleHeight     =   7125
   ScaleWidth      =   10050
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   8400
      TabIndex        =   7
      Top             =   6480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Close"
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
   Begin VB.ListBox lstSelectedIDs 
      Height          =   5580
      Left            =   9480
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox lstSelected 
      Height          =   5460
      Left            =   5160
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   720
      Width           =   4575
   End
   Begin VB.ListBox lstAll 
      Height          =   6000
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   240
      Width           =   3855
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   5160
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   2880
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin btButtonEx.ButtonEx btnRemove 
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   3360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Webdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstAllIDs 
      Height          =   6060
      Left            =   4080
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frmAssignValuesForCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    

Private Sub btnAdd_Click()
    If IsNumeric(cmbCategory.BoundText) = False Then
        MsgBox "Please select a category"
        cmbCategory.SetFocus
        Exit Sub
    End If
    
    Dim i As Integer
    Dim rsValue As New ADODB.Recordset
    For i = 0 To lstAll.ListCount - 1
        If lstAll.Selected(i) = True Then
            With rsValue
                If .State = 1 Then .Close
                temSql = "Select * from tblIxValuesForCategories where ValueID = " & Val(lstAllIDs.List(i)) & "  AND CategoryID = " & Val(cmbCategory.BoundText)
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount <= 0 Then
                    .AddNew
                    !ValueID = Val(lstAllIDs.List(i))
                    !CategoryID = Val(cmbCategory.BoundText)
                    .Update
                End If
                If .State = 1 Then .Close
            End With
        End If
    Next
    For i = 0 To lstAll.ListCount - 1
        lstAll.Selected(i) = False
        lstAllIDs.Selected(i) = False
    Next
    Call FillSelected
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnRemove_Click()
    Dim i As Integer
    Dim rsValue As New ADODB.Recordset
    For i = 0 To lstSelected.ListCount - 1
        If lstSelected.Selected(i) = True Then
            With rsValue
                If .State = 1 Then .Close
                temSql = "Delete * from tblIxValuesForCategories where IxValuesForCategoriesID = " & Val(lstSelectedIDs.List(i))
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .State = 1 Then .Close
            End With
        End If
    Next
    Call FillSelected
End Sub

Private Sub cmbCategory_Change()
    Call FillSelected
End Sub

Private Sub FillSelected()
    lstSelected.Clear
    lstSelectedIDs.Clear
    Dim rsValue As New ADODB.Recordset
    With rsValue
        If .State = 1 Then .Close
        temSql = "SELECT tblIxValuesForCategories.IxValuesForCategoriesID, tblIxValue.IxValue FROM tblIxValuesForCategories INNER JOIN tblIxValue ON tblIxValuesForCategories.ValueID = tblIxValue.IxValueID Where CategoryID = " & Val(cmbCategory.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            lstSelected.AddItem !IxValue
            lstSelectedIDs.AddItem !IxValuesForCategoriesID
            .MoveNext
        Wend
        .Close
    End With
End Sub

Private Sub Form_Load()
    Call FillCombos
    Call FillLists
    GetCommonSettings Me
    
    Call SetColours
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

Private Sub FillCombos()
    Dim Cat As New clsFillCombos
    Cat.FillBoolCombo cmbCategory, "Category", "Category", "IsIxValueCategory", True
End Sub

Private Sub FillLists()
    lstAll.Clear
    lstAllIDs.Clear
    Dim rsValue As New ADODB.Recordset
    With rsValue
        If .State = 1 Then .Close
        temSql = "Select * from tblIxValue where Deleted = False order by IxValue"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            lstAll.AddItem !IxValue
            lstAllIDs.AddItem !IxValueID
            .MoveNext
        Wend
        .Close
    End With
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub
