VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAuthorityPrevilagesMenuVisible 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Authority Previlages - Visibility of Menus"
   ClientHeight    =   7695
   ClientLeft      =   1275
   ClientTop       =   -2445
   ClientWidth     =   13815
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
   ScaleHeight     =   7695
   ScaleWidth      =   13815
   Begin MSDataListLib.DataCombo cmbAuthority 
      Height          =   360
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   635
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      Text            =   ""
   End
   Begin btButtonEx.ButtonEx btnAll 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Select All"
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
      Left            =   12240
      TabIndex        =   3
      Top             =   7080
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
   Begin TabDlg.SSTab sstMenus 
      Height          =   6375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   13485
      _ExtentX        =   23786
      _ExtentY        =   11245
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmAuthorityPrevilagesMenuVisible.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
   End
   Begin btButtonEx.ButtonEx btnSave 
      Height          =   495
      Left            =   10800
      TabIndex        =   5
      Top             =   7080
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
   Begin btButtonEx.ButtonEx btnNone 
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   7080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Select None"
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
      Caption         =   "Authority"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAuthorityPrevilagesMenuVisible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim FilledOnce As Boolean
    Dim temSql As String
    Dim rsAuthority As New ADODB.Recordset

Private Sub btnAll_Click()
    Dim MyCtrl As Control
    For Each MyCtrl In Controls
        If TypeOf MyCtrl Is CheckBox Then
            MyCtrl.Value = 1
        End If
    Next
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnNone_Click()
    Dim MyCtrl As Control
    For Each MyCtrl In Controls
        If TypeOf MyCtrl Is CheckBox Then
            MyCtrl.Value = 0
        End If
    Next
End Sub

Private Sub btnSave_Click()
    If IsNumeric(cmbAuthority.BoundText) = False Then
        MsgBox "Please select an Authority"
        cmbAuthority.SetFocus
        Exit Sub
    End If
    Dim rsTem As New ADODB.Recordset
    Dim MyCtrl As Control
    
    For Each MyCtrl In Controls
        If TypeOf MyCtrl Is CheckBox Then
            With rsTem
                If .State = 1 Then .Close
                temSql = "Select * from tblUserAuthorityControl where AuthorityID = " & Val(cmbAuthority.BoundText) & " AND ControlID = " & Val(MyCtrl.Tag)
                .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    If MyCtrl.Value = 1 Then
                        !Visible = True
                    Else
                        !Visible = False
                    End If
                    .Update
                Else
                    .AddNew
                    !AuthorityID = Val(cmbAuthority.BoundText)
                    !ControlID = Val(MyCtrl.Tag)
                    If MyCtrl.Value = 1 Then
                        !Visible = True
                    Else
                        !Visible = False
                    End If
                    .Update
                End If
            End With
        End If
    Next
    MsgBox "Saved"
End Sub

Private Sub cmbAuthority_Change()
    Dim rsTem As New ADODB.Recordset
    Dim MyCtrl As Control
    For Each MyCtrl In Controls
        If TypeOf MyCtrl Is CheckBox Then
            MyCtrl.Value = 0
        End If
    Next
    If IsNumeric(cmbAuthority.BoundText) = False Then Exit Sub
    For Each MyCtrl In Controls
        If TypeOf MyCtrl Is CheckBox Then
            MyCtrl.Value = 0
        End If
    Next
    For Each MyCtrl In Controls
        If TypeOf MyCtrl Is CheckBox Then
            With rsTem
                If .State = 1 Then .Close
                temSql = "Select * from tblUserAuthorityControl where AuthorityID = " & Val(cmbAuthority.BoundText) & " AND ControlID = " & Val(MyCtrl.Tag)
                .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                If .RecordCount > 0 Then
                    If !Visible = True Then
                        MyCtrl.Value = 1
                    Else
                        MyCtrl.Value = 0
                    End If
                End If
                .Close
            End With
        End If
    Next

End Sub

Private Sub Form_Activate()
    If FilledOnce = False Then
        FilledOnce = True
        Call FillControls
        'Call SetColours
    End If
End Sub

Private Sub FillCombos()
    With rsAuthority
        If .State = 1 Then .Close
        temSql = "SELECT tblAuthority.Authority, tblAuthority.AuthorityID FROM tblAuthority ORDER BY tblAuthority.Authority"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With cmbAuthority
        Set .RowSource = rsAuthority
        .BoundColumn = "AuthorityID"
        .ListField = "Authority"
    End With
End Sub

Private Sub Form_Load()
    FilledOnce = False
    Call FillCombos
End Sub

Private Sub FillControls()

    Screen.MousePointer = vbHourglass
    Dim rsTem As New ADODB.Recordset
    Dim rsFormegory As New ADODB.Recordset
    Dim rsMenu As New ADODB.Recordset
    
    Dim i As Integer
    
    Dim MyFrame As Frame
    Dim MyChkBox As CheckBox
    
    Dim FormName() As String
    Dim FormID() As Long
    Dim MenuName() As String
    Dim MenuNameID() As Long
    
    Dim LabelX As Long
    Dim LabelWidth As Long
    Dim LabelHeight As Long
    
    Dim LabelComboY As Long
    Dim LabelComboWidth As Long
    Dim LabelComboHeight As Long
    
    With rsFormegory
        If .State = 1 Then .Close
        temSql = "SELECT * from tblControl where Deleted = False AND MainMenu = True"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        ReDim FormID(.RecordCount)
        ReDim FormName(.RecordCount)
        i = 0
        
        If .RecordCount > 0 Then
            sstMenus.Visible = True
            .MoveLast
            sstMenus.Tabs = .RecordCount
            sstMenus.TabsPerRow = .RecordCount
            .MoveFirst
            
            
            While .EOF = False
                sstMenus.Tab = i
                sstMenus.TabCaption(i) = !ControlText
                FormID(i) = !ControlID
                FormName(i) = !Control
                
                If rsMenu.State = 1 Then rsMenu.Close
                temSql = "Select * from tblControl where MainMenuID = " & !ControlID & " AND ControlType =  " & ControlType.MenuItem & " AND MainMenu = False "
                rsMenu.Open temSql, cnnStores, adOpenStatic, adLockReadOnly
                
                LabelX = 100
                LabelWidth = 2500
                LabelHeight = 360
                
                LabelComboY = 380
                LabelComboWidth = 2600
                LabelComboHeight = 380
                
                While rsMenu.EOF = False
                    
                    If LabelComboY + LabelComboHeight > sstMenus.Height Then
                        LabelComboY = 380
                        LabelX = LabelX + LabelComboWidth
                    End If
                    If rsMenu!ControlType = ControlType.SSTab Then
                        Set MyChkBox = Me.Controls.Add("VB.checkbox", "chk" & rsMenu!Control & rsMenu!ControlIndex, Me)
                    Else
                        Set MyChkBox = Me.Controls.Add("VB.checkbox", "chk" & rsMenu!Control, Me)
                    End If
                    Set MyChkBox.Container = sstMenus
                    MyChkBox.Height = LabelHeight
                    MyChkBox.Width = LabelWidth
                    MyChkBox.Top = LabelComboY
                    MyChkBox.Left = LabelX
                    MyChkBox.Caption = rsMenu!ControlText
                    MyChkBox.Tag = rsMenu!ControlID
                    MyChkBox.Visible = True
                    LabelComboY = LabelComboY + LabelComboHeight
                    rsMenu.MoveNext
                Wend
                
                .MoveNext
                i = i + 1
            Wend
        Else
            sstMenus.Visible = False
        End If
    End With
    Screen.MousePointer = vbDefault
End Sub

Private Sub Setcolours()
'    Me.ForeColor = DefaultColourScheme.LabelForeColour
'    Me.BackColor = DefaultColourScheme.LabelBackColour
'    On Error Resume Next
'    Dim MyControl As Control
'    For Each MyControl In Controls
'        If InStr(UCase(MyControl.Name), "BTN") > 0 Then
'            MyControl.ForeColor = DefaultColourScheme.ButtonForeColour
'            MyControl.BackColor = DefaultColourScheme.ButtonBackColour
'            MyControl.BorderColor = DefaultColourScheme.ButtonBorderColour
'        ElseIf InStr(UCase(MyControl.Name), "LST") > 0 Then
'            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
'            MyControl.BackColor = DefaultColourScheme.LabelBackColour
'        ElseIf InStr(UCase(MyControl.Name), "CMB") > 0 Then
'            MyControl.ForeColor = DefaultColourScheme.ComboForeColour
'            MyControl.BackColor = DefaultColourScheme.ComboBackColour
'        ElseIf InStr(UCase(MyControl.Name), "TXT") > 0 Then
'            MyControl.ForeColor = DefaultColourScheme.TextForeColour
'            MyControl.BackColor = DefaultColourScheme.TextBackColour
'        ElseIf InStr(UCase(MyControl.Name), "DTP") > 0 Then
'            MyControl.ForeColor = DefaultColourScheme.TextForeColour
'            MyControl.BackColor = DefaultColourScheme.TextBackColour
'        Else
'            MyControl.ForeColor = DefaultColourScheme.LabelForeColour
'            MyControl.BackColor = DefaultColourScheme.LabelBackColour
'            MyControl.BackStyle = 0
'        End If
'    Next
End Sub

