VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11655
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
   ScaleHeight     =   8835
   ScaleWidth      =   11655
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   4560
      TabIndex        =   5
      Top             =   360
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   12938
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Details"
      TabPicture(0)   =   "frmItem.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(2)=   "lblIssueUnitsToPack"
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(4)=   "lblStrengthToIssueUnit"
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(8)=   "Label12"
      Tab(0).Control(9)=   "Label13"
      Tab(0).Control(10)=   "Label23"
      Tab(0).Control(11)=   "Label19"
      Tab(0).Control(12)=   "lblIssueUnit"
      Tab(0).Control(13)=   "lblIssueUnit1"
      Tab(0).Control(14)=   "Label22"
      Tab(0).Control(15)=   "lblIUnit2"
      Tab(0).Control(16)=   "Label25"
      Tab(0).Control(17)=   "cmbGenericName"
      Tab(0).Control(18)=   "cmbCategory"
      Tab(0).Control(19)=   "cmbIUnit"
      Tab(0).Control(20)=   "cmbPUnit"
      Tab(0).Control(21)=   "cmbSUnit"
      Tab(0).Control(22)=   "txtIssueUnitsPerPack"
      Tab(0).Control(23)=   "txtStrengthPerIssueUnit"
      Tab(0).Control(24)=   "txtItem"
      Tab(0).Control(25)=   "txtItemCode"
      Tab(0).Control(26)=   "txtMinIssueQty"
      Tab(0).Control(27)=   "txtMaxIQty"
      Tab(0).Control(28)=   "txtDivQty"
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Ordering"
      TabPicture(1)   =   "frmItem.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label20"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label18"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label16"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label15"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label17"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmbManufacturer"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmbSupplier"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmbMinQtyi"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmbMinQtyp"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmbImporter"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "cmbROQi"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmbROLi"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "cmbROQp"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "cmbROLp"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtMinQtyp"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtMinQtyi"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtROQi"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtROLi"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtROLp"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtROQp"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Other"
      TabPicture(2)   =   "frmItem.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label11"
      Tab(2).Control(1)=   "txtComments"
      Tab(2).ControlCount=   2
      Begin VB.TextBox txtDivQty 
         Height          =   375
         Left            =   -72600
         TabIndex        =   33
         Top             =   5760
         Width           =   1095
      End
      Begin VB.TextBox txtMaxIQty 
         Height          =   375
         Left            =   -72600
         TabIndex        =   30
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox txtComments 
         Height          =   6240
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   57
         Top             =   840
         Width           =   6615
      End
      Begin VB.TextBox txtMinIssueQty 
         Height          =   375
         Left            =   -72600
         TabIndex        =   27
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox txtROQp 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtROLp 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtROLi 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtROQi 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtMinQtyi 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtMinQtyp 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   3120
         Width           =   1335
      End
      Begin VB.TextBox txtItemCode 
         Height          =   360
         Left            =   -73440
         TabIndex        =   9
         Top             =   960
         Width           =   5175
      End
      Begin VB.TextBox txtItem 
         Height          =   360
         Left            =   -73440
         TabIndex        =   7
         Top             =   480
         Width           =   5175
      End
      Begin VB.TextBox txtStrengthPerIssueUnit 
         Height          =   375
         Left            =   -72600
         TabIndex        =   21
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox txtIssueUnitsPerPack 
         Height          =   375
         Left            =   -72600
         TabIndex        =   24
         Top             =   4320
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo cmbSUnit 
         Height          =   360
         Left            =   -72600
         TabIndex        =   15
         Top             =   2400
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbPUnit 
         Height          =   360
         Left            =   -72600
         TabIndex        =   19
         Top             =   3360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbIUnit 
         Height          =   360
         Left            =   -72600
         TabIndex        =   17
         Top             =   2880
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbCategory 
         Height          =   360
         Left            =   -73440
         TabIndex        =   11
         Top             =   1440
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbROLp 
         Height          =   360
         Left            =   3360
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   960
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbROQp 
         Height          =   360
         Left            =   3360
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   2040
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbROLi 
         Height          =   360
         Left            =   3360
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbROQi 
         Height          =   360
         Left            =   3360
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1560
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbImporter 
         Height          =   360
         Left            =   1920
         TabIndex        =   53
         Top             =   4560
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbMinQtyp 
         Height          =   360
         Left            =   3360
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   3120
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbMinQtyi 
         Height          =   360
         Left            =   3360
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2640
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Locked          =   -1  'True
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbSupplier 
         Height          =   360
         Left            =   1920
         TabIndex        =   55
         Top             =   5040
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbManufacturer 
         Height          =   360
         Left            =   1920
         TabIndex        =   51
         Top             =   4080
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbGenericName 
         Height          =   360
         Left            =   -73440
         TabIndex        =   13
         Top             =   1920
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   635
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label25 
         Caption         =   "Must be multiplies of"
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   5760
         Width           =   2775
      End
      Begin VB.Label lblIUnit2 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -71400
         TabIndex        =   34
         Top             =   5760
         Width           =   3135
      End
      Begin VB.Label Label22 
         Caption         =   "Maximum Issue Qty."
         Height          =   255
         Left            =   -74880
         TabIndex        =   29
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label lblIssueUnit1 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -71400
         TabIndex        =   31
         Top             =   5280
         Width           =   3135
      End
      Begin VB.Label Label11 
         Caption         =   "Comments"
         Height          =   255
         Left            =   -74760
         TabIndex        =   56
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblIssueUnit 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -71400
         TabIndex        =   28
         Top             =   4800
         Width           =   3135
      End
      Begin VB.Label Label19 
         Caption         =   "Minimum Issue Qty."
         Height          =   255
         Left            =   -74880
         TabIndex        =   26
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label Label23 
         Caption         =   "Pack Unit:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   18
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Generic Name"
         Height          =   255
         Left            =   -74880
         TabIndex        =   12
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label17 
         Caption         =   "Manufacturer"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "Reorder &Quentity"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label15 
         Caption         =   "&Reorder Leval"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label16 
         Caption         =   "&Supplier"
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label Label18 
         Caption         =   "&Importer"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   4560
         Width           =   2055
      End
      Begin VB.Label Label20 
         Caption         =   "&Minimum Order"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label12 
         Caption         =   "Item Code"
         Height          =   255
         Left            =   -74880
         TabIndex        =   8
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Item"
         Height          =   255
         Left            =   -74880
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Issue units per Pack"
         Height          =   255
         Left            =   -74880
         TabIndex        =   23
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label6 
         Caption         =   "Strength Unit"
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label lblStrengthToIssueUnit 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -71400
         TabIndex        =   22
         Top             =   3840
         Width           =   3135
      End
      Begin VB.Label Label8 
         Caption         =   "Strength of an Issue Unit"
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label lblIssueUnitsToPack 
         BorderStyle     =   1  'Fixed Single
         Height          =   375
         Left            =   -71400
         TabIndex        =   25
         Top             =   4320
         Width           =   3135
      End
      Begin VB.Label Label4 
         Caption         =   "Catogery"
         Height          =   255
         Left            =   -74880
         TabIndex        =   10
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Issue Unit"
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         Top             =   2880
         Width           =   1695
      End
   End
   Begin btButtonEx.ButtonEx btnAdd 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   8280
      Width           =   1455
      _ExtentX        =   2566
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
   Begin btButtonEx.ButtonEx btnEdit 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Edit"
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
      Left            =   3120
      TabIndex        =   4
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSDataListLib.DataCombo cmbItem 
      Height          =   7860
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   13864
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   1
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   10320
      TabIndex        =   60
      Top             =   8400
      Width           =   1215
      _ExtentX        =   2143
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
   Begin btButtonEx.ButtonEx btnCancel 
      Height          =   375
      Left            =   8280
      TabIndex        =   59
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Cancel"
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
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   375
      Left            =   6960
      TabIndex        =   58
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Save"
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
   Begin VB.Label Label1 
      Caption         =   "Diagnosis"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsTem As New ADODB.Recordset
    Dim temSql As String
    Dim MyItem As New clsItem
    
Private Sub btnAdd_Click()
    Dim temText As String
    If IsNumeric(cmbItem.BoundText) = True Then
        temText = Empty
    Else
        temText = cmbItem.Text
    End If
    cmbItem.Text = Empty
    txtItem.Text = temText
    Call EditMode
    SSTab1.Tab = 0
    txtItem.SetFocus
    On Error Resume Next: SendKeys "{home}+{end}"
End Sub

Private Sub btnCancel_Click()
    Call ClearValues
    Call FillCombos
    Call SelectMode
    cmbItem.SetFocus
    cmbItem.Text = Empty
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    If IsNumeric(cmbItem.BoundText) = False Then
        MsgBox "Please select an Item to delete"
        cmbItem.SetFocus
        Exit Sub
    End If
    
    If CalculateStock(Val(cmbItem.BoundText)).Stock > 0 Then
        MsgBox "Please remove stock first."
        cmbItem.SetFocus
        Exit Sub
    End If
    
    
    Dim i As Integer
    i = MsgBox("Are you sure you want to delete " & cmbItem.Text & "?", vbYesNo)
    If i = vbNo Then Exit Sub
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblItem where ItemID = " & Val(cmbItem.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Deleted = True
            !DeletedUserID = UserID
            !DeletedTime = Now
            .Update
            MsgBox "Successfully Deleted"
        End If
        .Close
    End With
    Call ClearValues
    Call FillCombos
    cmbItem.Text = Empty
    cmbItem.SetFocus
End Sub

Private Sub btnEdit_Click()
    If IsNumeric(cmbItem.BoundText) = True Then
        Call EditMode
        txtItem.SetFocus
        On Error Resume Next: SendKeys "{home}+{end}"
    End If
End Sub

Private Sub btnSave_Click()
    If IsNumeric(cmbItem.BoundText) = True Then
        Call SaveOld
    Else
        Call SaveNew
    End If
    Call ClearValues
    Call FillCombos
    Call SelectMode
    cmbItem.SetFocus
    cmbItem.Text = Empty
End Sub

Private Sub SaveOld()
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblItem where ItemID = " & Val(cmbItem.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !Item = Trim(txtItem.Text)
            !ItemCode = Trim(txtItemCode.Text)
            !IsItemName = True
'            !GroupID = Val(cmbGroup.BoundText)
'            !SubGroupID = Val(cmbSubGroup.BoundText)
'            !TradeNameID = Val(cmbTradeName.BoundText)
            !GenericNameID = Val(cmbGenericName.BoundText)
            !ItemCategoryID = Val(cmbCategory.BoundText)
            !StrengthUnitID = Val(cmbSUnit.BoundText)
            !IssueUnitID = Val(cmbIUnit.BoundText)
            !PackUnitID = Val(cmbPUnit.BoundText)
            !StrengthUnitID = Val(cmbSUnit.BoundText)
            !MinIQty = Val(txtMinIssueQty.Text)
            !MaxIQty = Val(txtMaxIQty.Text)
            !DivQty = Val(txtDivQty.Text)
            If IsNumeric(txtStrengthPerIssueUnit.Text) = True Then
                !StrengthUnitsPerIssueUnit = txtStrengthPerIssueUnit.Text
            Else
                !StrengthUnitsPerIssueUnit = 1
            End If
            If IsNumeric(txtIssueUnitsPerPack.Text) = True Then
                !IssueUnitsPerPack = txtIssueUnitsPerPack.Text
            Else
                !IssueUnitsPerPack = 1
            End If
            !ROL = Val(txtROLi.Text)
            !ROQ = Val(txtROQi.Text)
            !MinROQ = Val(txtMinQtyi.Text)
            !ManufacturerID = Val(cmbManufacturer.BoundText)
            !ImporterID = Val(cmbImporter.BoundText)
            !SupplierID = Val(cmbSupplier.BoundText)
            !Comments = txtComments.Text
            .Update
        End If
        .Close
    End With
End Sub

Private Sub SaveNew()
    With rsTem
        If .State = 1 Then .Close
        temSql = "Select * from tblItem"
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        .AddNew
        !Item = Trim(txtItem.Text)
        !ItemCode = Trim(txtItemCode.Text)
        !IsItemName = True
'        !GroupID = Val(cmbGroup.BoundText)
'        !SubGroupID = Val(cmbSubGroup.BoundText)
'        !TradeNameID = Val(cmbTradeName.BoundText)
        !GenericNameID = Val(cmbGenericName.BoundText)
        !ItemCategoryID = Val(cmbCategory.BoundText)
        !StrengthUnitID = Val(cmbSUnit.BoundText)
        !IssueUnitID = Val(cmbIUnit.BoundText)
        !PackUnitID = Val(cmbPUnit.BoundText)
        !StrengthUnitID = Val(cmbSUnit.BoundText)
        !MinIQty = Val(txtMinIssueQty.Text)
        !MaxIQty = Val(txtMaxIQty.Text)
        !DivQty = Val(txtDivQty.Text)
        !AddedDate = Date
        !AddedTime = Time
        !AddedUserID = UserID
        If IsNumeric(txtStrengthPerIssueUnit.Text) = True Then
            !StrengthUnitsPerIssueUnit = txtStrengthPerIssueUnit.Text
        Else
            !StrengthUnitsPerIssueUnit = 1
        End If
        If IsNumeric(txtIssueUnitsPerPack.Text) = True Then
            !IssueUnitsPerPack = txtIssueUnitsPerPack.Text
        Else
            !IssueUnitsPerPack = 1
        End If
        !ROL = Val(txtROLi.Text)
        !ROQ = Val(txtROQi.Text)
        !MinROQ = Val(txtMinQtyi.Text)
        !ManufacturerID = Val(cmbManufacturer.BoundText)
        !ImporterID = Val(cmbImporter.BoundText)
        !SupplierID = Val(cmbSupplier.BoundText)
        !Comments = txtComments.Text
        .Update
        .Close
    End With
End Sub

Private Sub cmbGenericName_Change()
    Dim TemItem As New clsItem
    TemItem.ID = Val(cmbGenericName.BoundText)
'    cmbGroup.BoundText = TemItem.GroupID
'    cmbSubGroup.BoundText = TemItem.SubGroupID
End Sub

Private Sub cmbItem_Change()
    Call ClearValues
    If IsNumeric(cmbItem.BoundText) = True Then Call DisplayDetails
End Sub

Private Sub cmbIUnit_Change()
    Call SetUnits
End Sub

Private Sub cmbPUnit_Change()
    Call SetUnits
End Sub

Private Sub cmbSUnit_Change()
    Call SetUnits
End Sub

'Private Sub cmbTradeName_Change()
'    Dim TemItem As New clsItem
'    TemItem.ID = Val(cmbTradeName.BoundText)
'    cmbGenericName.BoundText = TemItem.GenericNameID
'End Sub

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
    Call SetColours
    Call FillCombos
    Call SelectMode
    GetCommonSettings Me
    
    
    
End Sub

Private Sub FillCombos()
'    Dim Gp As New clsFillCombos
'    Gp.FillBoolCombo cmbGroup, "Item", "Item", "IsGroup", True
'    Dim SGp As New clsFillCombos
'    SGp.FillBoolCombo cmbSubGroup, "Item", "Item", "IsSubgroup", True
'    Dim TN As New clsFillCombos
'    TN.FillBoolCombo cmbTradeName, "Item", "Item", "IsTradeName", True
    Dim It As New clsFillCombos
    It.FillBoolCombo cmbItem, "Item", "Item", "IsItemName", True
    Dim Generic As New clsFillCombos
    Generic.FillBoolCombo cmbGenericName, "Item", "Item", "IsGenericName", True
    
    Dim IUnit1 As New clsFillCombos
    IUnit1.FillBoolCombo cmbIUnit, "ItemUnit", "ItemUnit", "IsIssueUnit", True
    Dim IUnit2 As New clsFillCombos
    IUnit2.FillBoolCombo cmbMinQtyi, "ItemUnit", "ItemUnit", "IsIssueUnit", True
    Dim IUnit3 As New clsFillCombos
    IUnit3.FillBoolCombo cmbROLi, "ItemUnit", "ItemUnit", "IsIssueUnit", True
    Dim IUnit4 As New clsFillCombos
    IUnit4.FillBoolCombo cmbROQi, "ItemUnit", "ItemUnit", "IsIssueUnit", True
    
    Dim PUnit1 As New clsFillCombos
    PUnit1.FillBoolCombo cmbPUnit, "ItemUnit", "ItemUnit", "IsPackUnit", True
    Dim PUnit2 As New clsFillCombos
    PUnit2.FillBoolCombo cmbMinQtyp, "ItemUnit", "ItemUnit", "IsPackUnit", True
    Dim PUnit3 As New clsFillCombos
    PUnit3.FillBoolCombo cmbROLp, "ItemUnit", "ItemUnit", "IsPackUnit", True
    Dim PUnit4 As New clsFillCombos
    PUnit4.FillBoolCombo cmbROQp, "ItemUnit", "ItemUnit", "IsPackUnit", True
    
    Dim SUnit1 As New clsFillCombos
    SUnit1.FillBoolCombo cmbSUnit, "ItemUnit", "ItemUnit", "IsStrengthUnit", True
    
    Dim Cat As New clsFillCombos
    Cat.FillBoolCombo cmbCategory, "CaTEGORY", "Category", "IsItemCategory", True
    
    Dim Manuf As New clsFillCombos
    Manuf.FillBoolCombo cmbManufacturer, "Institution", "Institution", "IsManufacturer", True
    
    Dim Sup As New clsFillCombos
    Sup.FillBoolCombo cmbSupplier, "Institution", "Institution", "IsSupplier", True
    
    Dim Impo As New clsFillCombos
    Impo.FillBoolCombo cmbImporter, "Institution", "Institution", "IsImporter", True
    
End Sub

Private Sub EditMode()
    
    cmbItem.Enabled = False
    btnAdd.Enabled = False
    btnEdit.Enabled = False
    btnDelete.Enabled = False
    
    btnSave.Enabled = True
    btnCancel.Enabled = True
    
    txtComments.Enabled = True
    txtIssueUnitsPerPack.Enabled = True
    txtItem.Enabled = True
    txtItemCode.Enabled = True
    txtMinQtyi.Enabled = True
    txtMinQtyp.Enabled = True
    txtROLi.Enabled = True
    txtROLp.Enabled = True
    txtROQi.Enabled = True
    txtROQp.Enabled = True
    txtStrengthPerIssueUnit.Enabled = True
    txtMinIssueQty.Enabled = True
    txtMaxIQty.Enabled = True
    txtDivQty.Enabled = True
    cmbCategory.Enabled = True
    cmbSupplier.Enabled = True
    cmbManufacturer.Enabled = True
'    cmbGroup.Enabled = True
    cmbImporter.Enabled = True
    cmbIUnit.Enabled = True
    cmbManufacturer.Enabled = True
    cmbMinQtyi.Enabled = True
    cmbMinQtyp.Enabled = True
    cmbPUnit.Enabled = True
    cmbROLi.Enabled = True
    cmbROLp.Enabled = True
    cmbROQi.Enabled = True
    cmbROQp.Enabled = True
'    cmbSubGroup.Enabled = True
    cmbSUnit.Enabled = True
'    cmbTradeName.Enabled = True
    cmbGenericName.Enabled = True
    
End Sub

Private Sub SelectMode()
    
    cmbItem.Enabled = True
    btnAdd.Enabled = True
    btnEdit.Enabled = True
    btnDelete.Enabled = True
    
    btnSave.Enabled = False
    btnCancel.Enabled = False
    
    txtComments.Enabled = False
    txtIssueUnitsPerPack.Enabled = False
    txtItem.Enabled = False
    txtItemCode.Enabled = False
    txtMinQtyi.Enabled = False
    txtMinQtyp.Enabled = False
    txtROLi.Enabled = False
    txtROLp.Enabled = False
    txtROQi.Enabled = False
    txtROQp.Enabled = False
    txtStrengthPerIssueUnit.Enabled = False
    txtMinIssueQty.Enabled = False
    txtMaxIQty.Enabled = False
    txtDivQty.Enabled = False
    
    cmbSupplier.Enabled = False
    cmbCategory.Enabled = False
    cmbManufacturer.Enabled = False
'    cmbGroup.Enabled = False
    cmbImporter.Enabled = False
    cmbIUnit.Enabled = False
    cmbManufacturer.Enabled = False
    cmbMinQtyi.Enabled = False
    cmbMinQtyp.Enabled = False
    cmbPUnit.Enabled = False
    cmbROLi.Enabled = False
    cmbROLp.Enabled = False
    cmbROQi.Enabled = False
    cmbROQp.Enabled = False
'    cmbSubGroup.Enabled = False
    cmbSUnit.Enabled = False
    cmbGenericName.Enabled = False
'    cmbTradeName.Enabled = False
End Sub

Private Sub ClearValues()
    txtComments.Text = Empty
    txtIssueUnitsPerPack.Text = Empty
    txtItem.Text = Empty
    txtItemCode.Text = Empty
    txtMinQtyi.Text = Empty
    txtMinQtyp.Text = Empty
    txtROLi.Text = Empty
    txtROLp.Text = Empty
    txtROQi.Text = Empty
    txtROQp.Text = Empty
    txtStrengthPerIssueUnit.Text = Empty
    txtMinIssueQty.Text = Empty
    txtDivQty.Text = Empty
    txtMaxIQty.Text = Empty
    
    cmbCategory.Text = Empty
    cmbManufacturer.Text = Empty
'    cmbGroup.Text = Empty
    cmbImporter.Text = Empty
    cmbIUnit.Text = Empty
    cmbManufacturer.Text = Empty
    cmbMinQtyi.Text = Empty
    cmbMinQtyp.Text = Empty
    cmbPUnit.Text = Empty
    cmbROLi.Text = Empty
    cmbROLp.Text = Empty
    cmbROQi.Text = Empty
    cmbROQp.Text = Empty
'    cmbSubGroup.Text = Empty
    cmbSUnit.Text = Empty
    cmbSupplier.Text = Empty
'    cmbTradeName.Text = Empty
    cmbGenericName.Text = Empty
    
End Sub

Private Sub DisplayDetails()
    If IsNumeric(cmbItem.BoundText) = True Then
        MyItem.ID = Val(cmbItem.BoundText)
    End If
    cmbCategory.BoundText = MyItem.CategoryID
    cmbSupplier.BoundText = MyItem.SupplierID
'    cmbGroup.BoundText = MyItem.GroupID
    cmbImporter.BoundText = MyItem.ImporterID
    cmbIUnit.BoundText = MyItem.IssueUnitID
    cmbManufacturer.BoundText = MyItem.ManufacturerID
    cmbMinQtyi.BoundText = MyItem.IssueUnitID
    cmbMinQtyp.BoundText = MyItem.PackUnitID
    cmbPUnit.BoundText = MyItem.PackUnitID
    cmbROLi.BoundText = MyItem.IssueUnitID
    cmbROLp.BoundText = MyItem.PackUnitID
    cmbROQi.BoundText = MyItem.IssueUnitID
    cmbROQp.BoundText = MyItem.PackUnitID
'    cmbSubGroup.BoundText = MyItem.SubGroupID
    cmbSUnit.BoundText = MyItem.StrengthUnitID
'    cmbTradeName.BoundText = MyItem.TradeNameID
    cmbGenericName.BoundText = MyItem.GenericNameID
    
    txtComments.Text = MyItem.Comments
    txtIssueUnitsPerPack.Text = MyItem.IssueUnitsPerPack
    txtItem.Text = MyItem.Item
    txtItemCode.Text = MyItem.Code
    txtMinQtyi.Text = MyItem.MinQty
    txtROLi.Text = MyItem.ROL
    txtROQi.Text = MyItem.ROQ
    txtStrengthPerIssueUnit.Text = MyItem.StrengthUnitsPerIssueUnit
    txtMinIssueQty.Text = MyItem.MinIQty
    
    txtMaxIQty.Text = MyItem.MaxIQty
    txtDivQty.Text = MyItem.DivQty
    
End Sub

Private Sub SetUnits()
    lblStrengthToIssueUnit.Caption = cmbSUnit.Text & " / " & cmbIUnit.Text
    lblIssueUnitsToPack.Caption = cmbIUnit.Text & " / " & cmbPUnit.Text
    lblIssueUnit.Caption = cmbIUnit.Text
    lblIssueUnit1.Caption = cmbIUnit.Text
    lblIUnit2.Caption = cmbIUnit.Text
    cmbROLi.Text = cmbIUnit.Text
    cmbROLp.Text = cmbPUnit.Text
    cmbROQi.Text = cmbIUnit.Text
    cmbROQp.Text = cmbPUnit.Text
    cmbMinQtyi.Text = cmbIUnit.Text
    cmbMinQtyp.Text = cmbPUnit.Text
End Sub

Private Sub SetValues()
    If Val(txtIssueUnitsPerPack.Text) <> 0 Then
        txtROLp.Text = Val(txtROLi.Text) / Val(txtIssueUnitsPerPack.Text)
        txtROLi.Text = Val(txtROLp.Text) * Val(txtIssueUnitsPerPack.Text)
        txtROQp.Text = Val(txtROQi.Text) / Val(txtIssueUnitsPerPack.Text)
        txtROQi.Text = Val(txtROQp.Text) * Val(txtIssueUnitsPerPack.Text)
        txtROLp.Text = Val(txtROLi.Text) / Val(txtIssueUnitsPerPack.Text)
        txtROLi.Text = Val(txtROLp.Text) * Val(txtIssueUnitsPerPack.Text)
        txtMinQtyp.Text = Val(txtMinQtyi.Text) / Val(txtIssueUnitsPerPack.Text)
        txtMinQtyi.Text = Val(txtMinQtyp.Text) * Val(txtIssueUnitsPerPack.Text)
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveCommonSettings Me
    

End Sub

Private Sub txtIssueUnitsPerPack_Change()
    Call SetValues
End Sub

Private Sub txtMinQtyi_Change()
    Call SetValues
End Sub

'Private Sub txtMinQtyp_Change()
'    Call SetValues
'End Sub

Private Sub txtROLi_Change()
    Call SetValues
End Sub

'Private Sub txtROLp_Change()
'    Call SetValues
'End Sub

Private Sub txtROQi_Change()
    Call SetValues
End Sub

'Private Sub txtROQp_Change()
'    Call SetValues
'End Sub

Private Sub txtStrengthPerIssueUnit_Change()
    Call SetValues
End Sub
