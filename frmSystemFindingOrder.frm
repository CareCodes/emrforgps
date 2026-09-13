VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmSystemFindingOrder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Order"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4875
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
   ScaleHeight     =   6600
   ScaleWidth      =   4875
   Begin VB.OptionButton optOther 
      Caption         =   "Other"
      Height          =   240
      Left            =   3360
      TabIndex        =   8
      Top             =   960
      Width           =   855
   End
   Begin VB.OptionButton optEx 
      Caption         =   "Examination"
      Height          =   240
      Left            =   1620
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.OptionButton optHx 
      Caption         =   "History"
      Height          =   240
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Value           =   -1  'True
      Width           =   975
   End
   Begin MSDataListLib.DataList lstSystemFinding 
      Height          =   4620
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   8149
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnUp 
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Top             =   2640
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "5"
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
   Begin btButtonEx.ButtonEx btnDown 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   3120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "6"
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
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSDataListLib.DataCombo cmbSystem 
      Height          =   360
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label1 
      Caption         =   "System"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSystemFindingOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsSys As New ADODB.Recordset
    
Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDown_Click()
    If IsNumeric(lstSystemFinding.BoundText) = False Then
        MsgBox "Please select a System to move"
        lstSystemFinding.SetFocus
        Exit Sub
    End If
    Dim temID As Long
    temID = lstSystemFinding.BoundText
    
    Dim rsTem As New ADODB.Recordset
    Dim ChangedNo As Long
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblSystemFinding where SystemFindingID = " & lstSystemFinding.BoundText
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            ChangedNo = !OrderNo
            !OrderNo = !OrderNo + 1
            .Update
        End If
        temSql = "Select * from tblSystemFinding where  Deleted = False AND SystemID = " & Val(cmbSystem.BoundText) & " AND  SystemFindingID <> " & lstSystemFinding.BoundText & " AND orderNo <= " & ChangedNo + 1
        If optEx.Value = True Then
            temSql = temSql & " AND IsExamination = True"
        ElseIf optHx.Value = True Then
            temSql = temSql & " AND IsHistory = True"
        ElseIf optOther.Value = True Then
            temSql = temSql & " AND IsOther = True"
        End If
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            !OrderNo = !OrderNo - 1
            .Update
            .MoveNext
        Wend
        .Close
    End With
    Call FillList
    lstSystemFinding.BoundText = temID
End Sub

Private Sub btnUp_Click()
    If IsNumeric(lstSystemFinding.BoundText) = False Then
        MsgBox "Please select a System to move"
        lstSystemFinding.SetFocus
        Exit Sub
    End If
    Dim temID As Long
    temID = lstSystemFinding.BoundText
    
    Dim rsTem As New ADODB.Recordset
    Dim ChangedNo As Long
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblSystemFinding where SystemFindingID = " & lstSystemFinding.BoundText
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            ChangedNo = !OrderNo
            If !OrderNo <= 1 Then
                .Close
                Exit Sub
            End If
            !OrderNo = !OrderNo - 1
            .Update
        End If
        temSql = "Select * from tblSystemFinding where Deleted = False AND SystemID = " & Val(cmbSystem.BoundText) & " AND SystemFindingID <> " & lstSystemFinding.BoundText & " AND orderNo >= " & ChangedNo - 1
        If optEx.Value = True Then
            temSql = temSql & " AND IsExamination = True"
        ElseIf optHx.Value = True Then
            temSql = temSql & " AND IsHistory = True"
        ElseIf optOther.Value = True Then
            temSql = temSql & " AND IsOther = True"
        End If
        If .State = 1 Then .Close
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            !OrderNo = !OrderNo + 1
            .Update
            .MoveNext
        Wend
        .Close
    End With
    Call FillList
    lstSystemFinding.BoundText = temID
End Sub

Private Sub cmbSystem_Click(Area As Integer)
    Call FillList
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
    Call FillSysCOmbo
End Sub

Private Sub FillSysCOmbo()
    Dim Systems As New clsFillCombos
    Systems.FillAnyCombo cmbSystem, "System", True
End Sub


Private Sub FillList()
    Dim temNo As Long
    temNo = 1
    With rsSys
        If .State = 1 Then .Close
        temSql = "Select * from tblSystemFinding Where Deleted = False AND SystemID = " & Val(cmbSystem.BoundText)
        If optEx.Value = True Then
            temSql = temSql & " And IsExamination = True"
        ElseIf optHx.Value = True Then
            temSql = temSql & " And IsHistory = True"
        ElseIf optOther.Value = True Then
            temSql = temSql & " And IsOther = True"
        End If
        temSql = temSql & " order by OrderNo, systemfindingid"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        While .EOF = False
            !OrderNo = temNo
            .Update
            temNo = temNo + 1
            .MoveNext
        Wend
    End With
    With lstSystemFinding
        Set .RowSource = rsSys
        .ListField = "SystemFinding"
        .BoundColumn = "SystemFindingID"
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub

Private Sub optEx_Click()
    Call FillList
End Sub

Private Sub optHx_Click()
    Call FillList
End Sub

Private Sub optOther_Click()
    Call FillList
End Sub
