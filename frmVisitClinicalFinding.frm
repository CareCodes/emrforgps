VERSION 5.00
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmVisitClinicalFinding 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clinical Findings"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12480
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
   ScaleHeight     =   8415
   ScaleWidth      =   12480
   Begin btButtonEx.ButtonEx btnClose 
      Height          =   375
      Left            =   11160
      TabIndex        =   4
      Top             =   7920
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
   Begin TabDlg.SSTab sstCategory 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   13573
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "History"
      TabPicture(0)   =   "frmVisitClinicalFinding.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "sstHx"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Examination"
      TabPicture(1)   =   "frmVisitClinicalFinding.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "sstEx"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Other"
      TabPicture(2)   =   "frmVisitClinicalFinding.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "sstOx"
      Tab(2).ControlCount=   1
      Begin TabDlg.SSTab sstEx 
         Height          =   7215
         Left            =   -74880
         TabIndex        =   2
         Top             =   360
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   12726
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "frmVisitClinicalFinding.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Tab 1"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Tab 2"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
      End
      Begin TabDlg.SSTab sstOx 
         Height          =   7215
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   12726
         _Version        =   393216
         TabHeight       =   520
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "frmVisitClinicalFinding.frx":0070
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Tab 1"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Tab 2"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
      End
      Begin TabDlg.SSTab sstHx 
         Height          =   7215
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   12726
         _Version        =   393216
         TabHeight       =   794
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "frmVisitClinicalFinding.frx":008C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
         TabCaption(1)   =   "Tab 1"
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Tab 2"
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
      End
   End
   Begin btButtonEx.ButtonEx bbtnAdd 
      Height          =   375
      Left            =   8760
      TabIndex        =   5
      Top             =   7920
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "&Add To Comments"
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
End
Attribute VB_Name = "frmVisitClinicalFinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSQL As String
    Dim SysHx() As Long
    Dim SysEx() As Long
    Dim SysOx() As Long
    Dim SystemFindingID() As Long
    Dim SystemFindingNo As Long
    Dim FilledOnce As Boolean
    Dim FormPatient As New clsPatient
    Dim FormVisit As New clsVisit
    
Private Sub bbtnAdd_Click()
    Call SaveDetails
    Dim temText As String
    Dim LastCommaPos As Long
    
    Dim rsTem As New ADODB.Recordset
    Dim rsSF As New ADODB.Recordset
    Dim rsValue As New ADODB.Recordset
    Dim rsVV As New ADODB.Recordset
    
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
        
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT tblSystem.System, tblSystem.SystemID " & _
                    "FROM tblSystemFinding LEFT JOIN tblSystem ON tblSystemFinding.SystemID = tblSystem.SystemID " & _
                    "WHERE (((tblSystem.Deleted)=False) AND ((tblSystemFinding.Deleted)=False)) " & _
                    "GROUP BY tblSystem.System, tblSystem.SystemID, tblSystem.OrderNo, tblSystemFinding.IsHistory " & _
                    "Having (((tblSystemFinding.IsHistory) = True)) " & _
                    "ORDER BY tblSystem.OrderNo "
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temText = "History" & vbNewLine
            While .EOF = False
                temText = temText & vbTab & !System & vbNewLine
                If rsSF.State = 1 Then rsSF.Close
                temSQL = "Select * from tblSystemFinding where SystemID = " & !SystemID & " AND IsHistory = true and Deleted = False order by OrderNO"
                rsSF.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                While rsSF.EOF = False
                    If rsValue.State = 1 Then rsValue.Close
                    temSQL = "Select * from tblSystemFindingValue where SystemFindingID = " & rsSF!SystemFindingID & " AND Deleted = False order by SystemFindingValue"
                    If rsVV.State = 1 Then rsVV.Close
                    temSQL = "Select * from tblVisitSystemFinding where SystemFindingID = " & rsSF!SystemFindingID & " AND Deleted = False And VisitID = " & FormVisit.VisitID
                    rsVV.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                    If rsVV.RecordCount > 0 Then
                        While rsVV.EOF = False
                            If rsVV.RecordCount > 0 Then
                                If Trim(rsVV!VisitSystemFinding) <> "" Then
                                    temText = temText & rsVV!VisitSystemFinding & ", "
                                End If
                            End If
                            rsVV.MoveNext
                        Wend
                    Else
                    
                    End If
                    rsVV.Close
                    rsSF.MoveNext
                Wend
                .MoveNext
            Wend
                    
            temText = temText & vbNewLine
        End If
    End With

    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT tblSystem.System, tblSystem.SystemID " & _
                    "FROM tblSystemFinding LEFT JOIN tblSystem ON tblSystemFinding.SystemID = tblSystem.SystemID " & _
                    "WHERE (((tblSystem.Deleted)=False) AND ((tblSystemFinding.Deleted)=False)) " & _
                    "GROUP BY tblSystem.System, tblSystem.SystemID, tblSystem.OrderNo, tblSystemFinding.IsExamination " & _
                    "Having (((tblSystemFinding.IsExamination) = True)) " & _
                    "ORDER BY tblSystem.OrderNo"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temText = temText & "Examination" & vbNewLine
            While .EOF = False
                temText = temText & vbTab & !System & vbNewLine
                If rsSF.State = 1 Then rsSF.Close
                temSQL = "Select * from tblSystemFinding where SystemID = " & !SystemID & " AND IsExamination = true and Deleted = False order by OrderNO"
                rsSF.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                While rsSF.EOF = False
                    If rsVV.State = 1 Then rsVV.Close
                    temSQL = "Select * from tblVisitSystemFinding where SystemFindingID = " & rsSF!SystemFindingID & " AND Deleted = False And VisitID = " & FormVisit.VisitID
                    rsVV.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                    If rsVV.RecordCount > 0 Then
                        While rsVV.EOF = False
                            If rsVV.RecordCount > 0 Then
                                If Trim(rsVV!VisitSystemFinding) <> "" Then
                                    temText = temText & rsVV!VisitSystemFinding & ","
                                End If
                            End If
                            rsVV.MoveNext
                        Wend
                    End If
                    rsVV.Close
                    rsSF.MoveNext
                Wend
                .MoveNext
            Wend
            temText = temText & vbNewLine
        Else
        
        End If
    End With


    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT tblSystem.System, tblSystem.SystemID " & _
                    "FROM tblSystemFinding LEFT JOIN tblSystem ON tblSystemFinding.SystemID = tblSystem.SystemID " & _
                    "WHERE (((tblSystem.Deleted)=False) AND ((tblSystemFinding.Deleted)=False)) " & _
                    "GROUP BY tblSystem.System, tblSystem.SystemID, tblSystem.OrderNo, tblSystemFinding.IsOther " & _
                    "Having (((tblSystemFinding.IsOther) = True)) " & _
                    "ORDER BY tblSystem.OrderNo"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            temText = temText & "Other Findings" & vbNewLine
            While .EOF = False
                temText = temText & vbTab & !System & vbNewLine
                If rsSF.State = 1 Then rsSF.Close
                temSQL = "Select * from tblSystemFinding where SystemID = " & !SystemID & " AND IsOther = true and Deleted = False order by OrderNO"
                rsSF.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                While rsSF.EOF = False
                    If rsVV.State = 1 Then rsVV.Close
                    temSQL = "Select * from tblVisitSystemFinding where SystemFindingID = " & rsSF!SystemFindingID & " AND Deleted = False And VisitID = " & FormVisit.VisitID
                    rsVV.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                    If rsVV.RecordCount > 0 Then
                        While rsVV.EOF = False
                            If rsVV.RecordCount > 0 Then
                                If Trim(rsVV!VisitSystemFinding) <> "" Then
                                    temText = temText & rsVV!VisitSystemFinding & ", "
                                End If
                            End If
                            rsVV.MoveNext
                        Wend
                    End If
                    rsVV.Close
                    rsSF.MoveNext
                Wend
                .MoveNext
            Wend
            temText = temText & vbNewLine
        Else
        End If
    End With
    
    If Trim(frmNewVisit.txtComments.Text) = Empty Then
        frmNewVisit.txtComments.Text = temText
    Else
        frmNewVisit.txtComments.Text = frmNewVisit.txtComments.Text & vbNewLine & temText
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If FilledOnce = True Then
        Exit Sub
    Else
        FilledOnce = True
        Call SetSSTs
        Call SetColours
    End If
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
    FormPatient.ID = CurrentPatient.ID
    FormVisit.VisitID = CurrentVisit.VisitID
    
    FilledOnce = False
    Me.Top = GetSetting(App.EXEName, Me.Name, "Top", Me.Top)
    Me.Left = GetSetting(App.EXEName, Me.Name, "Left", Me.Left)
    
End Sub

Private Sub SetSSTs()
    
    Screen.MousePointer = vbHourglass
    


    Dim rsTem As New ADODB.Recordset
    Dim rsSF As New ADODB.Recordset
    Dim rsValue As New ADODB.Recordset
    Dim rsVV As New ADODB.Recordset
    
    Dim i As Integer
    
    Dim MyFrame As Frame
    Dim MyLabel As Label
    Dim MyCombo As ComboBox
    
    Dim LabelX As Long
    Dim LabelWidth As Long
    Dim LabelHeight As Long
    
    Dim ComboX As Long
    Dim ComboWidth As Long
    Dim ComboHeight As Long
    
    Dim LabelComboY As Long
    Dim LabelComboWidth As Long
    Dim LabelComboHeight As Long
    
    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT tblSystem.System, tblSystem.SystemID " & _
                    "FROM tblSystemFinding LEFT JOIN tblSystem ON tblSystemFinding.SystemID = tblSystem.SystemID " & _
                    "WHERE (((tblSystem.Deleted)=False) AND ((tblSystemFinding.Deleted)=False)) " & _
                    "GROUP BY tblSystem.System, tblSystem.SystemID, tblSystem.OrderNo, tblSystemFinding.IsHistory " & _
                    "Having (((tblSystemFinding.IsHistory) = True)) " & _
                    "ORDER BY tblSystem.OrderNo "
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        ReDim SysHx(.RecordCount)
        i = 0
        If .RecordCount > 0 Then
            sstHx.Visible = True
            .MoveLast
            sstHx.Tabs = .RecordCount
            sstHx.TabsPerRow = .RecordCount
            .MoveFirst
            
            
            While .EOF = False
                sstHx.Tab = i
                sstHx.TabCaption(i) = !System
                SysHx(i) = !SystemID
                
                If rsSF.State = 1 Then rsSF.Close
                temSQL = "Select * from tblSystemFinding where SystemID = " & !SystemID & " AND IsHistory = true and Deleted = False order by OrderNO"
                rsSF.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                
                
                LabelX = 100
                LabelWidth = 1500
                LabelHeight = 500
                
                ComboX = 1600
                ComboWidth = 4000
                ComboHeight = 360
                
                LabelComboY = 600
                LabelComboWidth = 5600
                LabelComboHeight = 500
                
                
                
                While rsSF.EOF = False
                    
                    If LabelComboY + LabelComboHeight > sstHx.Height Then
                        LabelComboY = 600
                        LabelX = LabelX + LabelComboWidth
                        ComboX = ComboX + LabelComboWidth
                    End If
                    
                    Set MyLabel = Me.Controls.Add("VB.label", "lbl" & rsSF!SystemFindingID, Me)
                    Set MyLabel.Container = sstHx
                    MyLabel.Height = LabelHeight
                    MyLabel.Width = LabelWidth
                    MyLabel.Top = LabelComboY
                    MyLabel.Left = LabelX
                    MyLabel.BackStyle = 0
                    MyLabel.Caption = rsSF!SystemFinding
                    MyLabel.Visible = True
                    
                    Set MyCombo = Me.Controls.Add("VB.ComboBox", "cmb" & rsSF!SystemFindingID, Me)
                    Set MyCombo.Container = sstHx
                    MyCombo.Width = ComboWidth
                    MyCombo.Top = LabelComboY
                    MyCombo.Left = ComboX
                    MyCombo.Visible = True
                    
                    If rsValue.State = 1 Then rsValue.Close
                    temSQL = "Select * from tblSystemFindingValue where SystemFindingID = " & rsSF!SystemFindingID & " AND Deleted = False order by SystemFindingValue"
                    rsValue.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                    While rsValue.EOF = False
                        MyCombo.AddItem rsValue!SystemFindingValue
                        If rsValue!IsSystemFindingDefault = True Then MyCombo.Text = rsValue!SystemFindingValue
                        rsValue.MoveNext
                    Wend
                    rsValue.Close

                    If rsVV.State = 1 Then rsVV.Close
                    temSQL = "Select * from tblVisitSystemFinding where SystemFindingID = " & rsSF!SystemFindingID & " AND Deleted = False And VisitID = " & FormVisit.VisitID
                    rsVV.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                    While rsVV.EOF = False
                        If rsVV.RecordCount > 0 Then
                            MyCombo.Text = rsVV!VisitSystemFinding
                        End If
                        rsVV.MoveNext
                    Wend
                    rsVV.Close
                    
                    LabelComboY = LabelComboY + LabelComboHeight
                    
                    rsSF.MoveNext
                Wend
                
                .MoveNext
                i = i + 1
            Wend
        Else
            sstHx.Visible = False
        End If
    End With

    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT tblSystem.System, tblSystem.SystemID " & _
                    "FROM tblSystemFinding LEFT JOIN tblSystem ON tblSystemFinding.SystemID = tblSystem.SystemID " & _
                    "WHERE (((tblSystem.Deleted)=False) AND ((tblSystemFinding.Deleted)=False)) " & _
                    "GROUP BY tblSystem.System, tblSystem.SystemID, tblSystem.OrderNo, tblSystemFinding.IsExamination " & _
                    "Having (((tblSystemFinding.IsExamination) = True)) " & _
                    "ORDER BY tblSystem.OrderNo"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        ReDim SysEx(.RecordCount)
        i = 0
        If .RecordCount > 0 Then
            sstEx.Visible = True
            .MoveLast
            sstEx.Tabs = .RecordCount
            sstEx.TabsPerRow = .RecordCount
            .MoveFirst
            While .EOF = False
                sstEx.TabCaption(i) = !System
                SysEx(i) = !SystemID
                sstEx.Tab = i
                
                
                If rsSF.State = 1 Then rsSF.Close
                temSQL = "Select * from tblSystemFinding where SystemID = " & !SystemID & " AND IsExamination = true and Deleted = False order by OrderNO"
                rsSF.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                
                
                LabelX = 100
                LabelWidth = 1500
                LabelHeight = 500
                
                ComboX = 1600
                ComboWidth = 4000
                ComboHeight = 360
                
                LabelComboY = 600
                LabelComboWidth = 5600
                LabelComboHeight = 500
                
                
                
                While rsSF.EOF = False
                    
                    If LabelComboY + LabelComboHeight > sstEx.Height Then
                        LabelComboY = 600
                        LabelX = LabelX + LabelComboWidth
                        ComboX = ComboX + LabelComboWidth
                    End If
                    
                    Set MyLabel = Me.Controls.Add("VB.label", "lbl" & rsSF!SystemFindingID, Me)
                    Set MyLabel.Container = sstEx
                    MyLabel.Height = LabelHeight
                    MyLabel.Width = LabelWidth
                    MyLabel.Top = LabelComboY
                    MyLabel.BackStyle = 0
                    MyLabel.Left = LabelX
                    MyLabel.Caption = rsSF!SystemFinding
                    MyLabel.Visible = True
                    
                    Set MyCombo = Me.Controls.Add("VB.ComboBox", "cmb" & rsSF!SystemFindingID, Me)
                    Set MyCombo.Container = sstEx
                    MyCombo.Width = ComboWidth
                    MyCombo.Top = LabelComboY
                    MyCombo.Left = ComboX
                    MyCombo.Visible = True
                    
                    If rsValue.State = 1 Then rsValue.Close
                    temSQL = "Select * from tblSystemFindingValue where SystemFindingID = " & rsSF!SystemFindingID & " AND Deleted = False order by SystemFindingValue"
                    rsValue.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                    While rsValue.EOF = False
                        MyCombo.AddItem rsValue!SystemFindingValue
                        If rsValue!IsSystemFindingDefault = True Then MyCombo.Text = rsValue!SystemFindingValue
                        rsValue.MoveNext
                    Wend
                    rsValue.Close

                    If rsVV.State = 1 Then rsVV.Close
                    temSQL = "Select * from tblVisitSystemFinding where SystemFindingID = " & rsSF!SystemFindingID & " AND Deleted = False And VisitID = " & FormVisit.VisitID
                    rsVV.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                    While rsVV.EOF = False
                        If rsVV.RecordCount > 0 Then
                            MyCombo.Text = rsVV!VisitSystemFinding
                        End If
                        rsVV.MoveNext
                    Wend
                    rsVV.Close
                    
                    LabelComboY = LabelComboY + LabelComboHeight
                    
                    rsSF.MoveNext
                Wend
                
                
                .MoveNext
                i = i + 1
            Wend
        Else
            sstEx.Visible = False
        End If
    End With


    With rsTem
        If .State = 1 Then .Close
        temSQL = "SELECT tblSystem.System, tblSystem.SystemID " & _
                    "FROM tblSystemFinding LEFT JOIN tblSystem ON tblSystemFinding.SystemID = tblSystem.SystemID " & _
                    "WHERE (((tblSystem.Deleted)=False) AND ((tblSystemFinding.Deleted)=False)) " & _
                    "GROUP BY tblSystem.System, tblSystem.SystemID, tblSystem.OrderNo, tblSystemFinding.IsOther " & _
                    "Having (((tblSystemFinding.IsOther) = True)) " & _
                    "ORDER BY tblSystem.OrderNo"
        .Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
        ReDim SysOx(.RecordCount)
        i = 0
        If .RecordCount > 0 Then
            sstOx.Visible = True
            .MoveLast
            sstOx.Tabs = .RecordCount
            sstOx.TabsPerRow = .RecordCount
            .MoveFirst
            While .EOF = False
                sstOx.TabCaption(i) = !System
                SysOx(i) = !SystemID
                sstOx.Tab = i
                
                
                If rsSF.State = 1 Then rsSF.Close
                temSQL = "Select * from tblSystemFinding where SystemID = " & !SystemID & " AND IsOther = true and Deleted = False order by OrderNO"
                rsSF.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                
                
                LabelX = 100
                LabelWidth = 1500
                LabelHeight = 500
                
                ComboX = 1600
                ComboWidth = 4000
                ComboHeight = 360
                
                LabelComboY = 600
                LabelComboWidth = 5600
                LabelComboHeight = 500
                
                
                
                While rsSF.EOF = False
                    
                    If LabelComboY + LabelComboHeight > sstOx.Height Then
                        LabelComboY = 600
                        LabelX = LabelX + LabelComboWidth
                        ComboX = ComboX + LabelComboWidth
                    End If
                    
                    Set MyLabel = Me.Controls.Add("VB.label", "lbl" & rsSF!SystemFindingID, Me)
                    Set MyLabel.Container = sstOx
                    MyLabel.Height = LabelHeight
                    MyLabel.Width = LabelWidth
                    MyLabel.Top = LabelComboY
                    MyLabel.Left = LabelX
                    MyLabel.BackStyle = 0
                    MyLabel.Caption = rsSF!SystemFinding
                    MyLabel.Visible = True
                    
                    Set MyCombo = Me.Controls.Add("VB.ComboBox", "cmb" & rsSF!SystemFindingID, Me)
                    Set MyCombo.Container = sstOx
                    MyCombo.Width = ComboWidth
                    MyCombo.Top = LabelComboY
                    MyCombo.Left = ComboX
                    MyCombo.Visible = True
                    
                    If rsValue.State = 1 Then rsValue.Close
                    temSQL = "Select * from tblSystemFindingValue where SystemFindingID = " & rsSF!SystemFindingID & " AND Deleted = False order by SystemFindingValue"
                    rsValue.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                    While rsValue.EOF = False
                        MyCombo.AddItem rsValue!SystemFindingValue
                        If rsValue!IsSystemFindingDefault = True Then MyCombo.Text = rsValue!SystemFindingValue
                        rsValue.MoveNext
                    Wend
                    rsValue.Close

                    If rsVV.State = 1 Then rsVV.Close
                    temSQL = "Select * from tblVisitSystemFinding where SystemFindingID = " & rsSF!SystemFindingID & " AND Deleted = False And VisitID = " & FormVisit.VisitID
                    rsVV.Open temSQL, cnnStores, adOpenStatic, adLockReadOnly
                    While rsVV.EOF = False
                        If rsVV.RecordCount > 0 Then
                            MyCombo.Text = rsVV!VisitSystemFinding
                        End If
                        rsVV.MoveNext
                    Wend
                    rsVV.Close
                    
                    LabelComboY = LabelComboY + LabelComboHeight
                    
                    rsSF.MoveNext
                Wend
                
                
                
                .MoveNext
                i = i + 1
            Wend
        Else
            sstOx.Visible = False
        End If
    End With
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SaveDetails
End Sub

Private Sub SaveDetails()
    
    Screen.MousePointer = vbHourglass

    Dim MyCombo As Control
    Dim rsTem As New ADODB.Recordset
    Dim temID As Long
    SaveSetting App.EXEName, Me.Name, "Top", Me.Top
    SaveSetting App.EXEName, Me.Name, "Left", Me.Left
    For Each MyCombo In Controls
        If InStr(MyCombo.Name, "cmb") > 0 Then
            temID = Val(Right(MyCombo.Name, (Len(MyCombo.Name) - 3)))
            With rsTem
                If .State = 1 Then .Close
                temSQL = "Select * from tblVisitSystemFinding where VisitID = " & FormVisit.VisitID & " AND SystemFindingID = " & temID
                .Open temSQL, cnnStores, adOpenStatic, adLockOptimistic
                If .RecordCount > 0 Then
                    !VisitSystemFinding = MyCombo.Text
                    .Update
                Else
                    .AddNew
                    !VisitID = FormVisit.VisitID
                    !PatientID = FormPatient.ID
                    !SystemFindingID = temID
                    !VisitSystemFinding = MyCombo.Text
                    .Update
                End If
                .Close
            End With
        End If
    Next
    
    Screen.MousePointer = vbDefault
    
End Sub
