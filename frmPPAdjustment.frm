VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPPAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Price Adjustment"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7545
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
   ScaleHeight     =   5805
   ScaleWidth      =   7545
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   4800
      TabIndex        =   27
      Top             =   5160
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
   Begin VB.TextBox txtComments 
      Alignment       =   1  'Right Justify
      Height          =   1335
      Left            =   3240
      TabIndex        =   26
      Top             =   3720
      Width           =   4095
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   495
      Left            =   6120
      TabIndex        =   28
      Top             =   5160
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   2055
      Left            =   2160
      TabIndex        =   6
      Top             =   1560
      Width           =   5205
      _ExtentX        =   9181
      _ExtentY        =   3625
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Issue Units"
      TabPicture(0)   =   "frmPPAdjustment.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblIUnit(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblIUnit(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblIUnit(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtICurrentStock"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtINewStock"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtIAdjustment"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Pack Units"
      TabPicture(1)   =   "frmPPAdjustment.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtPAdjustment"
      Tab(1).Control(1)=   "txtPNewStock"
      Tab(1).Control(2)=   "txtPCurrentStock"
      Tab(1).Control(3)=   "lblPUnit(2)"
      Tab(1).Control(4)=   "lblPUnit(1)"
      Tab(1).Control(5)=   "lblPUnit(0)"
      Tab(1).Control(6)=   "Label9"
      Tab(1).Control(7)=   "Label8"
      Tab(1).Control(8)=   "Label7"
      Tab(1).ControlCount=   9
      Begin VB.TextBox txtPAdjustment 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   -73320
         TabIndex        =   23
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtPNewStock 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   -73320
         TabIndex        =   20
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtPCurrentStock 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73320
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtIAdjustment 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtINewStock 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox txtICurrentStock 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblIUnit 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   3240
         TabIndex        =   15
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblIUnit 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   12
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblPUnit 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   2
         Left            =   -71760
         TabIndex        =   24
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblPUnit 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   1
         Left            =   -71760
         TabIndex        =   21
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblPUnit 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   -71760
         TabIndex        =   18
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblIUnit 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Index           =   0
         Left            =   3240
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "&Adjustment"
         Height          =   255
         Left            =   -74880
         TabIndex        =   22
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label8 
         Caption         =   "&New Price"
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "C&urrent Price"
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label6 
         Caption         =   "&Adjustment"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "&New Price"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "C&urrent Price"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   2295
      End
   End
   Begin MSDataListLib.DataCombo cmbCategory 
      Height          =   360
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbItem 
      Height          =   360
      Left            =   2160
      TabIndex        =   3
      Top             =   600
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbCode 
      Height          =   360
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label11 
      Caption         =   "C&ode"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label10 
      Caption         =   "Co&mments"
      Height          =   255
      Left            =   2160
      TabIndex        =   25
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "&Item"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "&Category"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmPPAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String
    Dim rsItem As New ADODB.Recordset
    Dim rsCode As New ADODB.Recordset
    Dim MyItem As New clsItem
    
Private Sub GetFormSettings()
    GetCommonSettings Me
    
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    Dim rsTem As New ADODB.Recordset
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblUserAction"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !UserID = UserID
        !DepartmentID = UserDepartmentID
        !ActionID = UserAction.PurchasePriceAdjustment
        !ActionDate = Date
        !ActionTime = Time
        !ActionDateTime = Now
        !RefID = Val(cmbItem.BoundText)
        !RefValue = Val(txtIAdjustment.Text)
        !RefComments = txtComments.Text
        !PreviousValue = Val(txtICurrentStock.Text)
        !NewValue = Val(txtINewStock.Text)
        .Update
    End With
    SetPPrice Val(cmbItem.BoundText), Val(txtINewStock.Text)
    SetLastPPrice Val(cmbItem.BoundText), Val(txtICurrentStock.Text)
    Call ClearValues
    cmbCategory.SetFocus
End Sub

Private Sub ClearValues()
    cmbCategory.Text = Empty
    cmbItem.Text = Empty
    cmbCode.Text = Empty
    txtIAdjustment.Text = Empty
    txtICurrentStock.Text = Empty
    txtINewStock.Text = Empty
    txtPAdjustment.Text = Empty
    txtPCurrentStock.Text = Empty
    txtPNewStock.Text = Empty
    txtComments.Text = Empty
End Sub

Private Sub DisplayBatch()
    txtICurrentStock.Text = Empty
    txtIAdjustment.Text = Empty
    txtINewStock.Text = Empty
    txtPCurrentStock.Text = Empty
    txtPAdjustment.Text = Empty
    txtPNewStock.Text = Empty
    If IsNumeric(cmbItem.BoundText) = False Then Exit Sub
    txtICurrentStock.Text = GetPPrice(cmbItem.BoundText)
End Sub

Private Sub cmbCategory_Change()
    With rsItem
        If .State = 1 Then .Close
        If IsNumeric(cmbCategory.BoundText) = True Then
            temSql = "Select * from tblItem where ItemcATEGORYid = " & Val(cmbCategory.BoundText) & " AND Deleted = false Order by Item"
        Else
            temSql = "Select * from tblItem where Deleted = false order by Item"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbItem
        Set .RowSource = rsItem
        .ListField = "Item"
        .BoundColumn = "ItemID"
        .Text = Empty
    End With
    With rsCode
        If .State = 1 Then .Close
        If IsNumeric(cmbCategory.BoundText) = True Then
            temSql = "Select * from tblItem where ItemcATEGORYid = " & Val(cmbCategory.BoundText) & " AND Deleted = false Order by Item"
        Else
            temSql = "Select * from tblItem where Deleted = false order by Item"
        End If
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbCode
        Set .RowSource = rsCode
        .ListField = "ItemCode"
        .BoundColumn = "ItemID"
        .Text = Empty
    End With
    
End Sub

Private Sub cmbCode_Change()
    cmbItem.BoundText = Val(cmbCode.BoundText)
End Sub

Private Sub cmbItem_Change()
    MyItem.ID = Val(cmbItem.BoundText)
    cmbCode.BoundText = Val(cmbItem.BoundText)
    Dim i As Integer
    For i = 0 To lblIUnit.UBound
        lblIUnit(i).Caption = MyItem.IssueUnit
    Next
    For i = 0 To lblPUnit.UBound
        lblPUnit(i).Caption = MyItem.PackUnit
    Next
    Call DisplayBatch
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetFormSettings
End Sub

Private Sub SetFormSettings()
    SaveCommonSettings Me
    
End Sub
    
Private Sub Form_Load()
    Call SetColours
    Call FillCombos
    Call GetFormSettings
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
    Dim ICat As New clsFillCombos
    ICat.FillBoolCombo cmbCategory, "Category", "Category", "IsItemCategory", True
End Sub

Private Sub txtIAdjustment_Lostfocus()
    If MyItem.IssueUnitsPerPack = 0 Then Exit Sub
    txtPAdjustment.Text = Round(Val(txtIAdjustment.Text) * MyItem.IssueUnitsPerPack, 2)
End Sub

Private Sub txtICurrentStock_Change()
    If MyItem.IssueUnitsPerPack = 0 Then Exit Sub
    txtPCurrentStock.Text = Round(Val(txtICurrentStock.Text) * MyItem.IssueUnitsPerPack, 2)
    txtINewStock.Text = Val(txtICurrentStock.Text) + Val(txtIAdjustment.Text)
End Sub

Private Sub txtINewStock_Lostfocus()
    If MyItem.IssueUnitsPerPack = 0 Then Exit Sub
    txtPNewStock.Text = Round(Val(txtINewStock.Text) * MyItem.IssueUnitsPerPack, 2)
    txtIAdjustment.Text = Val(txtINewStock.Text) - Val(txtICurrentStock.Text)
End Sub

Private Sub txtPAdjustment_Change()
    If SSTab1.Tab = 1 Then
        If MyItem.IssueUnitsPerPack = 0 Then Exit Sub
        txtIAdjustment.Text = Val(txtPAdjustment.Text) / MyItem.IssueUnitsPerPack
        txtIAdjustment_Lostfocus
    End If
End Sub

Private Sub txtPNewStock_Change()
    If SSTab1.Tab = 1 Then
        If MyItem.IssueUnitsPerPack = 0 Then Exit Sub
        txtINewStock.Text = Val(txtPNewStock.Text) / MyItem.IssueUnitsPerPack
        txtINewStock_Lostfocus
    End If
End Sub
