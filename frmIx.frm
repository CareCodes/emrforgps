VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmIx 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Investigations"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIx.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   10410
   Begin VB.TextBox txtcomments 
      Height          =   1095
      Left            =   5640
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3000
      Width           =   4455
   End
   Begin VB.TextBox txtIx 
      Height          =   360
      Left            =   5640
      TabIndex        =   6
      Top             =   120
      Width           =   4455
   End
   Begin VB.TextBox txtVolume 
      Height          =   360
      Left            =   5640
      TabIndex        =   4
      Top             =   2520
      Width           =   4455
   End
   Begin MSDataListLib.DataCombo cmbIx 
      Height          =   3780
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   6668
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
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
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Edit"
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
      Left            =   2520
      TabIndex        =   3
      Top             =   4080
      Width           =   975
      _ExtentX        =   1720
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
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   5640
      TabIndex        =   5
      Top             =   600
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   4200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Cancel"
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
   Begin MSDataListLib.DataCombo cmbDept 
      Height          =   360
      Left            =   5640
      TabIndex        =   10
      Top             =   1080
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbTube 
      Height          =   360
      Left            =   5640
      TabIndex        =   11
      Top             =   1560
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbSpeciman 
      Height          =   360
      Left            =   5640
      TabIndex        =   18
      Top             =   2040
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      ListField       =   ""
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   9120
      TabIndex        =   20
      Top             =   4800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Cancel"
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
   Begin VB.Label Label7 
      Caption         =   "S&peciman"
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "C&omments"
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "&Investigation"
      Height          =   255
      Left            =   3840
      TabIndex        =   16
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Categ&ory"
      Height          =   255
      Left            =   3840
      TabIndex        =   15
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "&Department"
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Volume"
      Height          =   255
      Left            =   3840
      TabIndex        =   13
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label6 
      Caption         =   "&Tube"
      Height          =   255
      Left            =   3840
      TabIndex        =   12
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "frmIx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    
Private Sub btnAdd_Click()
    Dim temText As String
    If IsNumeric(cmbIx.BoundText) = False Then
        temText = cmbIx.Text
    Else
        temText = Empty
    End If
    cmbIx.Text = Empty
    Call EditMode
    txtIx.Text = temText
    txtIx.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call SelectMode
    cmbIx.Text = Empty
    cmbIx.SetFocus
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete " & cmbIx.Text, vbYesNo)
    If i = vbNo Then Exit Sub
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIx where IxID = " & Val(cmbIx.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedTime = Now
            !DeletedUserID = UserID
            .Update
            MsgBox "Deleted"
        Else
            MsgBox "Nothing to Delete"
        End If
        .Close
    End With
    Set rsTem = Nothing
    Call FillCombos
    cmbIx.SetFocus
    cmbIx.Text = Empty
    
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbIx.BoundText) = False Then Exit Sub
    Call EditMode
    txtIx.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub btnSave_Click()
    If Trim(txtIx.Text) = Empty Then
        MsgBox "You have not entered an Investigation"
        txtIx.SetFocus
        Exit Sub
    End If
    If IsNumeric(cmbIx.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call SelectMode
    Call ClearValues
    Call FillCombos
    cmbIx.Text = Empty
    cmbIx.SetFocus
End Sub

Private Sub cmbIx_Change()
    Call ClearValues
    If IsNumeric(cmbIx.BoundText) = True Then Call DisplayDetails
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
    Call SelectMode
    Call FillCombos
End Sub

Private Sub EditMode()
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    cmbIx.Enabled = False
    
    txtIx.Enabled = True
    txtComments.Enabled = True
    cmbSpeciman.Enabled = True
    cmbTube.Enabled = True
    cmbDept.Enabled = True
    txtVolume.Enabled = True
    cmbCategory.Enabled = True
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    
End Sub

Private Sub SelectMode()
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    cmbIx.Enabled = True
    
    txtIx.Enabled = False
    txtComments.Enabled = False
    cmbSpeciman.Enabled = False
    cmbTube.Enabled = False
    cmbDept.Enabled = False
    txtVolume.Enabled = False
    cmbCategory.Enabled = False
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
End Sub

Private Sub ClearValues()
    txtIx.Text = Empty
    txtComments.Text = Empty
    cmbSpeciman.Text = Empty
    cmbTube.Text = Empty
    cmbDept.Text = Empty
    txtVolume.Text = Empty
    cmbCategory.Text = Empty
End Sub

Private Sub SaveNew()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIx"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Ix = txtIx.Text
        !IxCategoryID = Val(cmbCategory.BoundText)
        !DepartmentID = Val(cmbDept.BoundText)
        !Comments = txtComments.Text
        !SpecimanID = Val(cmbSpeciman.BoundText)
        !TubeVolume = Val(txtVolume.Text)
        !TubeID = Val(cmbTube.BoundText)
        .Update
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub SaveOld()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIx where IxID = " & Val(cmbIx.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Ix = txtIx.Text
            !Comments = txtComments.Text
            !DepartmentID = Val(cmbDept.BoundText)
            !IxCategoryID = Val(cmbCategory.BoundText)
            !SpecimanID = Val(cmbSpeciman.BoundText)
            !TubeID = Val(cmbTube.BoundText)
            !TubeVolume = Val(txtVolume.Text)
        .Update
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub FillCombos()
    Dim Ix As New clsFillCombos
    Ix.FillAnyCombo cmbIx, "Ix", True
    Dim Speciman As New clsFillCombos
    Speciman.FillBoolCombo cmbSpeciman, "Category", "Category", "IsSpecimanCategory", True
    Dim Tube As New clsFillCombos
    Tube.FillBoolCombo cmbTube, "Category", "Category", "IsTubeCategory", True
    Dim Dept As New clsFillCombos
    Dept.FillAnyCombo cmbDept, "Department", True
    Dim IxCat As New clsFillCombos
    IxCat.FillBoolCombo cmbCategory, "Category", "Category", "IsIxCategory", True
End Sub

Private Sub DisplayDetails()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblIx where IxID = " & Val(cmbIx.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            txtIx.Text = !Ix
            txtComments.Text = Format(!Comments, "")
            cmbSpeciman.BoundText = !SpecimanID
            cmbTube.BoundText = !TubeID
            cmbDept.BoundText = !DepartmentID
            cmbCategory.BoundText = !IxCategoryID
            txtVolume.Text = !TubeVolume
        End If
        .Close
    End With
    Set rsTem = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub
