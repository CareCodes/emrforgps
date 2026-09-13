VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{575E4548-F564-4A9D-8667-6EE848F77EB8}#1.0#0"; "ButtonEx.ocx"
Begin VB.Form frmPharmacyList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pharmacy List"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7440
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
   ScaleHeight     =   8055
   ScaleWidth      =   7440
   Begin VB.ListBox lstNotAvailable 
      Height          =   1740
      Left            =   3000
      TabIndex        =   15
      Top             =   5880
      Width           =   2655
   End
   Begin VB.ListBox lstAvailable 
      Height          =   1740
      Left            =   120
      TabIndex        =   14
      Top             =   5880
      Width           =   2655
   End
   Begin VB.OptionButton optAll 
      Caption         =   "All"
      Height          =   480
      Left            =   3480
      TabIndex        =   12
      Top             =   1440
      Width           =   3255
   End
   Begin VB.OptionButton optOutSideIssue 
      Caption         =   "Outside Medicines from outside"
      Height          =   480
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   3255
   End
   Begin VB.OptionButton optOutdoorIssued 
      Caption         =   "Outdoor Medicines Issued"
      Height          =   600
      Left            =   3480
      TabIndex        =   10
      Top             =   960
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   6240
      Top             =   4920
   End
   Begin VB.OptionButton optIndoorIssued 
      Caption         =   "Indoor Medicines Issued"
      Height          =   600
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin VB.OptionButton optToOutdoorIssue 
      Caption         =   "To Issue Outdoor Medicines"
      Height          =   480
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Value           =   -1  'True
      Width           =   3375
   End
   Begin VB.OptionButton optToIndoorIssue 
      Caption         =   "To Issue Indoor Medicines"
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3015
   End
   Begin btButtonEx.ButtonEx btnSale 
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Sale"
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
   Begin MSDataListLib.DataList lstQueue 
      Height          =   3420
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6033
      _Version        =   393216
   End
   Begin btButtonEx.ButtonEx btnOutSide 
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   3000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Outside"
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
      Left            =   6000
      TabIndex        =   6
      Top             =   7560
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
   Begin btButtonEx.ButtonEx btnIncome 
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Income"
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
   Begin btButtonEx.ButtonEx btnIndoor 
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Appearance      =   3
      Caption         =   "Indoor"
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd MMMM yyyy"
      Format          =   57671683
      CurrentDate     =   39859
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Available"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Not Available"
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   5640
      Width           =   2775
   End
End
Attribute VB_Name = "frmPharmacyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim temSql As String

Private Sub chkCompleted_Click()
    Call FillQueue
End Sub

Private Sub chkRemoved_Click()
    Call FillQueue
End Sub

Private Sub chkToComplete_Click()
    Call FillQueue
End Sub

Private Sub btnAddAgain_Click()
    Dim temLong As Integer
    temLong = MsgBox("Are you sure you want to Add this patient?", vbYesNo)
    If temLong = vbNo Then Exit Sub
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSql = "Select * from tblQueue where QueueID = " & Val(lstQueue.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If !Completed = True Then
                MsgBox "This visit is cimpleted. Therefore you can't remove it"
                .Close
                Exit Sub
            End If
            If !Removed = False Then
                MsgBox "This visit is not a removed visit"
                .Close
                Exit Sub
            End If
            !Removed = False
            !RemovedTime = Null
            !RemovedUserID = Null
            .Update
        End If
        .Close
    End With
    Call FillQueue

End Sub

Private Sub btnRemove_Click()
    Dim temLong As Integer
    temLong = MsgBox("Are you sure you want to remove this patient?", vbYesNo)
    If temLong = vbNo Then Exit Sub
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSql = "Select * from tblQueue where QueueID = " & Val(lstQueue.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            If !Completed = True Then
                MsgBox "This visit is cimpleted. Therefore you can't remove it"
                .Close
                Exit Sub
            End If
            If !Removed = True Then
                MsgBox "This visit is already removed"
                .Close
                Exit Sub
            End If
            !Removed = True
            !RemovedTime = Now
            !RemovedUserID = UserID
            .Update
        End If
        .Close
    End With
    Call FillQueue
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnIncome_Click()
    frmPharmacyQueueIncome.Show
    frmPharmacyQueueIncome.ZOrder 0
End Sub

Private Sub btnOutSide_Click()
    Dim temLong As Integer
    temLong = MsgBox("Are you sure you want let patient to take outside medicines from any other pharmacy?", vbYesNo)
    If temLong = vbNo Then Exit Sub
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSql = "Select * from tblQueue where QueueID = " & Val(lstQueue.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !OutsideIssued = True
            !OutsideIssuedTime = Now
            !OutsideIssuedUserID = Null
            .Update
        End If
        .Close
    End With
    Call FillQueue
End Sub

Private Sub btnSale_Click()
    If IsNumeric(lstQueue.BoundText) = False Then
        Beep
        Exit Sub
    End If
    
    Dim MyVisitID As Long
    
    frmPharmacySale.Show
    frmPharmacySale.Caption = "Pharmacy Sale - " & lstQueue.Text
    
    
    Dim rsTemItem As New ADODB.Recordset
    With rsTemItem
        If .State = 1 Then .Close
        temSql = "SELECT tblVisit.VisitID, tblVisitItem.PItemID, tblVisitItem.PIssueQuentity " & _
                    "FROM tblVisit LEFT JOIN tblVisitItem ON tblVisit.VisitID = tblVisitItem.VisitID " & _
                    "WHERE tblVisit.VisitID = " & Val(lstQueue.BoundText) & " AND tblVisitItem.Indoor = False AND tblVisitItem.Deleted = False"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        If .RecordCount > 0 Then
            MyVisitID = !VisitID
            While .EOF = False
                If CalculateStock(!PItemID, 0, UserDepartmentID).Stock >= !PIssueQuentity Then
                    frmPharmacySale.cmbItem.BoundText = !PItemID
                    frmPharmacySale.txtQty.Text = !PIssueQuentity
                    frmPharmacySale.btnAdd_Click
                End If
                .MoveNext
            Wend
        End If
        .Close
    End With
    
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSql = "Select * from tblQueue where QueueID = " & Val(lstQueue.BoundText)
        .Open temSql, cnnStores, adOpenStatic, adLockOptimistic
        If .RecordCount > 0 Then
            !OutdoorIssued = True
            !OutdoorIssuedTime = Now
            !OutdoorIssuedUserID = Null
            .Update
        End If
        .Close
    End With
    Call FillQueue
    
    
    Dim temText As String
    Dim rsTem As New ADODB.Recordset
    
    With rsTem
        If .State = 1 Then .Close
        temSql = "SELECT tblVisitItem.VisitItemID, tblVisitItem.ItemID, tblVisitItem.Indoor,  [tblItem.Item] & ' '   &  tblCategory.Category   & ' '   & [tblVisitItem].[Dose] & ' ' & [tblDoseUnit].[ItemUnit] & ' ' & [tblFrequencyUnit].[ItemUnit] & ' ' & ' for ' & [tblVisitItem].[Duration] & ' ' & [tblDurationUnit].[ItemUnit] AS Prescreption, 'Issue ' & tblVisitItem.IssueQuentity & ' ' & tblIssueUnit.ItemUnit AS ToIssue " & _
                    "FROM ((((tblItemUnit AS tblDoseUnit RIGHT JOIN (tblItemUnit AS tblIssueUnit RIGHT JOIN tblVisitItem ON tblIssueUnit.ItemUnitID = tblVisitItem.IssueUnitID) ON tblDoseUnit.ItemUnitID = tblVisitItem.DoseUnitID) LEFT JOIN tblItemUnit AS tblFrequencyUnit ON tblVisitItem.FrequencyUnitID = tblFrequencyUnit.ItemUnitID) LEFT JOIN tblItemUnit AS tblDurationUnit ON tblVisitItem.DurationUnitID = tblDurationUnit.ItemUnitID) LEFT JOIN tblItem ON tblVisitItem.ItemID = tblItem.ItemID) LEFT JOIN tblCategory ON tblVisitItem.ItemCategoryID = tblCategory.CategoryID " & _
                    "WHERE (((tblVisitItem.VisitID)=" & MyVisitID & ") AND ((tblVisitItem.Deleted)=False))" & _
                    "Order by tblVisitItem.VisitItemID"
                    
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If !Indoor = False Then
                temText = temText & !Prescreption & vbNewLine
            End If
            .MoveNext
        Wend
    End With
    
    Dim MyVisit As New clsVisit
    
    MyVisit.VisitID = MyVisitID
    
    frmPharmacySale.cmbCustomer.BoundText = MyVisit.PatientID
    frmPharmacySale.txtComments.Text = temText
    frmPharmacySale.txtQueueID.Text = MyVisit.QueueID
    frmPharmacySale.ZOrder 0
    
End Sub

Private Sub cmbDoctor_Change()
    Call FillQueue
End Sub

Private Sub cmbSecession_Change()
    Call FillQueue
End Sub

Private Sub dtpDate_Change()
    Call FillQueue
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
    Call SetColours
    Call FillCombos
    Call GetFormSettings
    dtpDate.Value = Date
    Call FillQueue
     If AutoSecession = True Then Call SetAutoSecession
End Sub

Private Sub SetAutoSecession()

End Sub


Private Sub FillCombos()

End Sub

Private Sub GetFormSettings()
    GetCommonSettings Me
    
    
    On Error Resume Next
    Dim MyControl  As Control
    For Each MyControl In Controls
        If Left(MyControl.Name, 3) = "opt" Then
            MyControl.Value = GetSetting(App.EXEName, Me.Name, MyControl.Name, "False")
        End If
    Next
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call SetFormSettings
End Sub

Private Sub SetFormSettings()
    SaveCommonSettings Me
    
    
    Dim MyControl  As Control
    For Each MyControl In Controls
        If Left(MyControl.Name, 3) = "opt" Then
            SaveSetting App.EXEName, Me.Name, MyControl.Name, MyControl.Value
        End If
        If Left(MyControl.Name, 3) = "cmb" Then
            SaveSetting App.EXEName, Me.Name, MyControl.Name, MyControl.Text
        End If
    Next


End Sub

Private Sub FillQueue()
    Dim temQID As Long
    temQID = Val(lstQueue.BoundText)
    Dim rsQueue As New ADODB.Recordset
    With rsQueue
        If .State = 1 Then .Close
        temSql = "SELECT tblVisit.VisitID & ' ' & tblPatient.Patient & ' (' & tblPatient.PatientID & ') ' as QueueDisplay, tblVisit.VisitID FROM tblVisit LEFT JOIN tblPatient ON tblVisit.PatientID = tblPatient.PatientID where tblVisit.Deleted = False AND tblVisit.Deleted = False AND tblVisit.VisitDate = #" & Format(dtpDate.Value, "dd MMMM yyyy") & "# "
        
        If optToIndoorIssue.Value = True Then
            temSql = temSql & " and tblVisit.Completed = true And tblVisit.IndoorIssued = False  AND tblVisit.HasIndoorDrugs = True "
        ElseIf optToOutdoorIssue.Value = True Then
            temSql = temSql & " and  tblVisit.Completed = true AND tblVisit.OutdoorIssued = false AND tblVisit.HasOutdoorDrugs = True And tblVisit.OutSideIssued = False "
        ElseIf optIndoorIssued.Value = True Then
            temSql = temSql & " and  tblVisit.Completed = true  AND tblVisit.IndoorIssued = true  AND tblVisit.HasIndoorDrugs = True "
        ElseIf optOutdoorIssued.Value = True Then
            temSql = temSql & " and  tblVisit.Completed = true  AND tblVisit.OutdoorIssued = True And tblVisit.HasOutdoorDrugs = True "
        ElseIf optOutdoorIssued.Value = True Then
            temSql = temSql & " and  tblVisit.Completed = true  AND tblVisit.OutsideIssued = true AND tblVisit.HasOutdoorDrugs = True "
        ElseIf optOutSideIssue.Value = True Then
            temSql = temSql & " And tblVisit.Completed = true And tblVisit.OutSideIssued = true"
        ElseIf optAll.Value = True Then
            
        End If
        
        temSql = temSql & " Order by VisitID"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
    End With
    With lstQueue
        Set .RowSource = rsQueue
        .ListField = "QueueDisplay"
        .BoundColumn = "VIsitID"
        .BoundText = temQID
    End With
End Sub

Private Sub lstQueue_Click()
    lstAvailable.Clear
    lstNotAvailable.Clear
    Dim rsTemItem As New ADODB.Recordset
    With rsTemItem
        If .State = 1 Then .Close
        temSql = "SELECT tblItem.Item, tblVisitItem.PItemID, tblVisitItem.PIssueQuentity FROM (tblVisit LEFT JOIN tblVisitItem ON tblVisit.VisitID = tblVisitItem.VisitID) LEFT JOIN tblItem ON tblVisitItem.PItemID = tblItem.ItemID " & _
                    "WHERE (((tblVisitItem.PItemID)<>0) AND ((tblVisit.VisitID)=" & Val(lstQueue.BoundText) & ") AND ((tblVisitItem.Indoor)=False) AND ((tblVisitItem.Deleted)=False))"
        .Open temSql, cnnStores, adOpenStatic, adLockReadOnly
        While .EOF = False
            If CalculateStock(!PItemID, 0, UserDepartmentID).Stock < !PIssueQuentity Then
                lstNotAvailable.AddItem !Item
            Else
                lstAvailable.AddItem !Item
            End If
            .MoveNext
        Wend
        .Close
    End With

End Sub

Private Sub optAll_Click()
    Call FillQueue

End Sub

Private Sub optIndoorIssued_Click()
    Call FillQueue

End Sub

Private Sub optOutdoorIssued_Click()
    Call FillQueue

End Sub

Private Sub optOutSideIssue_Click()
    Call FillQueue

End Sub

Private Sub optToIndoorIssue_Click()
    Call FillQueue
End Sub

Private Sub optToOutdoorIssue_Click()
    Call FillQueue

End Sub

Private Sub Timer1_Timer()
    Call FillQueue
End Sub

